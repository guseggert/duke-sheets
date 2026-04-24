//! XLS RC4 CryptoAPI decryption (MS-OFFCRYPTO §2.3.5 + MS-XLS §2.2.10).
//!
//! This covers `FilePass` records with `wEncryptionType=1`, `vMajor` in
//! {2, 3, 4}, `vMinor=2`. The legacy MD5-KDF variant (`vMajor=1`,
//! `vMinor=1`) is handled separately by `rc4_legacy`.
//!
//! # Algorithm outline
//!
//! 1. **KDF**: `h0 = sha1(salt || password_utf16le)`, then
//!    `h_final(block) = sha1(h0 || u32le(block))`. For `key_size_bits=40`
//!    the key is `h_final[..5] || zeros(11)` (128-bit RC4 key with the
//!    last 11 bytes zeroed); for other sizes it's the first
//!    `key_size_bits/8` bytes of `h_final`.
//!
//! 2. **Password verify**: decrypt `encrypted_verifier` and
//!    `encrypted_verifier_hash` with a *single* RC4 instance keyed at
//!    `block=0`. The keystream continues across them. `sha1(verifier)`
//!    must match the first 20 bytes of the decrypted verifier hash.
//!
//! 3. **Stream decrypt**: feed the Workbook stream through RC4 in
//!    1024-byte blocks, re-keying each block from scratch with
//!    `makekey(..., block=stream_offset/1024)`. Plaintext regions
//!    (record headers, the 7 excluded records, the BoundSheet8
//!    `lbPlyPos` field, and FilePass itself) are co-indexed with the
//!    ciphertext: the RC4 engine runs over zero-filled placeholders,
//!    and the known plaintext is overlaid at the end. This preserves
//!    absolute stream offsets — critical because `BoundSheet8.lbPlyPos`
//!    points to absolute Workbook-stream positions of sheet BOFs.

use rc4::{KeyInit, Rc4, StreamCipher};
use sha1::{Digest, Sha1};
use subtle::ConstantTimeEq;

use crate::error::{CryptoError, CryptoResult};
use crate::password::utf16le_bytes;
use crate::xls::record_walk;

/// Inputs to the RC4 CryptoAPI KDF, parsed from the FilePass record.
#[derive(Debug, Clone)]
pub struct Rc4CryptoApiParams {
    /// 16-byte random salt.
    pub salt: [u8; 16],
    /// Key length in bits (40 or 128 most common; can be up to 512).
    pub key_size_bits: u32,
    /// 16-byte ciphertext of the verifier.
    pub encrypted_verifier: [u8; 16],
    /// 32-byte ciphertext of the verifier hash. The first 20 bytes after
    /// decryption are compared against `sha1(verifier)`.
    pub encrypted_verifier_hash: [u8; 32],
}

/// Derive the RC4 key for a given 1024-byte block index.
///
/// Implements MS-OFFCRYPTO §2.3.5.2:
/// `h0 = sha1(salt || password_utf16le); key = sha1(h0 || u32le(block))`
/// truncated to `key_size_bits`. The 40-bit case returns a 16-byte
/// buffer whose last 11 bytes are zero (the RC4 key is always 128-bit
/// wide on the wire; the KDF constrains its effective strength).
pub fn make_key(password: &str, salt: &[u8; 16], key_size_bits: u32, block: u32) -> Vec<u8> {
    let pw_utf16 = utf16le_bytes(password);
    let mut h = Sha1::new();
    h.update(salt);
    h.update(&pw_utf16);
    let h0 = h.finalize();

    let mut h = Sha1::new();
    h.update(h0);
    h.update(block.to_le_bytes());
    let h_final = h.finalize();

    if key_size_bits == 40 {
        let mut out = vec![0u8; 16];
        out[..5].copy_from_slice(&h_final[..5]);
        out
    } else {
        let key_bytes = (key_size_bits / 8) as usize;
        h_final[..key_bytes].to_vec()
    }
}

/// Check that `password` decrypts the FilePass verifier correctly.
///
/// Uses a constant-time comparison on the final SHA-1 digest. The slow
/// defense against brute force is the KDF itself — this just prevents
/// trivial timing oracles on the verifier check.
pub fn verify_password(params: &Rc4CryptoApiParams, password: &str) -> bool {
    let key = make_key(password, &params.salt, params.key_size_bits, 0);
    let mut rc4 = Rc4::new_from_slice(&key[..16]).expect("16-byte key is valid for RC4");

    let mut verifier = params.encrypted_verifier;
    rc4.apply_keystream(&mut verifier);

    let mut verifier_hash = params.encrypted_verifier_hash;
    rc4.apply_keystream(&mut verifier_hash);

    let mut computed = Sha1::new();
    computed.update(verifier);
    let computed = computed.finalize();

    // `computed` is 20 bytes; `verifier_hash` is 32 (zero-padded to RC4
    // block alignment on the write side). Compare the leading 20 bytes.
    computed.ct_eq(&verifier_hash[..20]).into()
}

/// Apply RC4 CryptoAPI decryption to a byte buffer, re-keying every
/// `block_size` bytes (1024 for XLS, 512 for OOXML Binary RC4).
///
/// Every block gets a fresh RC4 instance (no state carries across
/// blocks); only the derived key differs per block. This function does
/// not know about XLS record structure — callers must zero-pad
/// plaintext positions in the input and overlay them in the output.
pub fn decrypt_raw(
    password: &str,
    salt: &[u8; 16],
    key_size_bits: u32,
    block_size: usize,
    ciphertext: &[u8],
) -> Vec<u8> {
    let mut out = Vec::with_capacity(ciphertext.len());

    for (block_idx, chunk) in ciphertext.chunks(block_size).enumerate() {
        let key = make_key(password, salt, key_size_bits, block_idx as u32);
        let mut rc4 = Rc4::new_from_slice(&key[..16]).expect("16-byte key is valid for RC4");
        let mut buf = chunk.to_vec();
        rc4.apply_keystream(&mut buf);
        out.extend_from_slice(&buf);
    }

    out
}

/// Decrypt an XLS Workbook stream that was encrypted with RC4 CryptoAPI.
///
/// `workbook_stream` is the full stream including the FilePass record.
/// The output is the same length; the FilePass record has its body
/// zeroed (but its 4-byte header is preserved) so `BoundSheet8.lbPlyPos`
/// offsets remain valid.
pub fn decrypt_workbook_stream(
    workbook_stream: &[u8],
    password: &str,
    params: &Rc4CryptoApiParams,
) -> CryptoResult<Vec<u8>> {
    if !verify_password(params, password) {
        return Err(CryptoError::BadPassword);
    }

    let classified = record_walk::classify(workbook_stream)?;
    const BLOCK_SIZE: usize = 1024;
    let decrypted = decrypt_raw(
        password,
        &params.salt,
        params.key_size_bits,
        BLOCK_SIZE,
        &classified.ciphertext,
    );
    Ok(record_walk::apply_overlay(decrypted, &classified.overlay))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn make_key_for_known_vectors_is_deterministic() {
        // msoffcrypto reference: for password "test", salt [0x11, 0x22, ..., 0xFF][..16],
        // key_size=128, block=0, we should get a stable 16-byte key.
        let salt: [u8; 16] = [
            0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, //
            0x99, 0xAA, 0xBB, 0xCC, 0xDD, 0xEE, 0xFF, 0x00,
        ];
        let k1 = make_key("test", &salt, 128, 0);
        let k2 = make_key("test", &salt, 128, 0);
        assert_eq!(k1, k2);
        assert_eq!(k1.len(), 16);
    }

    #[test]
    fn make_key_differs_per_block() {
        let salt = [0u8; 16];
        let k0 = make_key("pw", &salt, 128, 0);
        let k1 = make_key("pw", &salt, 128, 1);
        assert_ne!(k0, k1);
    }

    #[test]
    fn make_key_40bit_pads_to_16_bytes_with_zeros() {
        // MS-OFFCRYPTO: the 40-bit variant is wire-compatible with the
        // 128-bit RC4 cipher but zero-pads the tail.
        let salt = [0u8; 16];
        let k = make_key("pw", &salt, 40, 0);
        assert_eq!(k.len(), 16);
        assert_eq!(&k[5..], &[0u8; 11]);
    }

    #[test]
    fn make_key_empty_password_is_stable() {
        let salt = [0u8; 16];
        let k = make_key("", &salt, 128, 0);
        assert_eq!(k.len(), 16);
    }

    /// Verifier-check on a constructed RC4 CryptoAPI params block. This
    /// test locks in the exact sequence of calls (one RC4 instance
    /// spanning both encrypted_verifier and encrypted_verifier_hash)
    /// against the known-vector path by construction: we encrypt a
    /// known verifier with RC4, take its SHA-1, encrypt the hash, and
    /// verify that decrypt-and-compare passes.
    #[test]
    fn verify_password_accepts_correct_password() {
        let password = "duke-test-pw";
        let salt = [
            0xAA, 0xBB, 0xCC, 0xDD, 0xEE, 0xFF, 0x00, 0x11, //
            0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99,
        ];
        let key_size_bits = 128u32;

        // Build a verifier and its hash, then encrypt both with the same
        // RC4 instance — mirroring how a writer would lay them down.
        let verifier = [0x13u8; 16];
        let mut hasher = Sha1::new();
        hasher.update(verifier);
        let verifier_hash = hasher.finalize();
        let mut verifier_hash_padded = [0u8; 32];
        verifier_hash_padded[..20].copy_from_slice(&verifier_hash);

        let key = make_key(password, &salt, key_size_bits, 0);
        let mut rc4 = Rc4::new_from_slice(&key[..16]).unwrap();

        let mut enc_verifier = verifier;
        rc4.apply_keystream(&mut enc_verifier);
        let mut enc_verifier_hash = verifier_hash_padded;
        rc4.apply_keystream(&mut enc_verifier_hash);

        let params = Rc4CryptoApiParams {
            salt,
            key_size_bits,
            encrypted_verifier: enc_verifier,
            encrypted_verifier_hash: enc_verifier_hash,
        };
        assert!(verify_password(&params, password));
        assert!(!verify_password(&params, "wrong-password"));
    }

    #[test]
    fn decrypt_raw_rekeys_every_block() {
        let salt = [0u8; 16];
        // Two blocks' worth of identical plaintext → two different
        // ciphertexts because the key changes per block.
        let plaintext = vec![0x41u8; 2048];
        let mut encrypted = Vec::with_capacity(2048);
        for (block_idx, chunk) in plaintext.chunks(1024).enumerate() {
            let key = make_key("pw", &salt, 128, block_idx as u32);
            let mut rc4 = Rc4::new_from_slice(&key[..16]).unwrap();
            let mut c = chunk.to_vec();
            rc4.apply_keystream(&mut c);
            encrypted.extend_from_slice(&c);
        }
        // First 1024 bytes ciphertext MUST differ from second 1024
        // because of per-block re-keying.
        assert_ne!(&encrypted[..1024], &encrypted[1024..]);

        // Round-trip: same function decrypts.
        let decrypted = decrypt_raw("pw", &salt, 128, 1024, &encrypted);
        assert_eq!(decrypted, plaintext);
    }
}
