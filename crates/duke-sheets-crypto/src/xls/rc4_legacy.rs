//! XLS Legacy RC4 decryption (MS-OFFCRYPTO §2.3.6 + MS-XLS §2.4.117).
//!
//! Applies to FilePass records with `wEncryptionType=1`, `vMajor=1`,
//! `vMinor=1` — the Excel 97/2000-era encryption before CryptoAPI. Same
//! RC4 cipher and 1024-byte re-keying as `rc4_cryptoapi`, but the key
//! derivation uses MD5 and has a different salt-expansion step.
//!
//! # KDF
//! ```text
//! h0 = md5(password_utf16le)             # 16 bytes
//! intermediate = (h0[..5] || salt) * 16  # 336 bytes (16 repetitions)
//! h1 = md5(intermediate)                 # 16 bytes
//! for each 1024-byte block:
//!     h_final = md5(h1[..5] || u32le(block))
//!     key = h_final                      # 128-bit
//! ```
//!
//! # FilePass layout
//! After `wEncryptionType || vMajor || vMinor` (6 bytes, already
//! consumed by the dispatcher), the body is a flat 48-byte blob:
//! ```text
//! u8[16] Salt
//! u8[16] EncryptedVerifier
//! u8[16] EncryptedVerifierHash
//! ```
//! Unlike the CryptoAPI variant there's no `EncryptionHeader` /
//! `EncryptionVerifier` framing; just the three fixed-size fields.
//!
//! # Password verification
//! Decrypt `encrypted_verifier` and `encrypted_verifier_hash` with a
//! single RC4 instance keyed at block 0. The keystream continues across
//! the two reads. Compare `md5(decrypted_verifier)` against
//! `decrypted_verifier_hash` (both 16 bytes).

use md5::{Digest as _, Md5};
use rc4::{KeyInit, Rc4, StreamCipher};
use subtle::ConstantTimeEq;

use crate::error::{CryptoError, CryptoResult};
use crate::password::utf16le_bytes;
use crate::xls::record_walk;

/// Parameters parsed from the legacy RC4 FilePass body.
#[derive(Debug, Clone)]
pub struct Rc4LegacyParams {
    pub salt: [u8; 16],
    pub encrypted_verifier: [u8; 16],
    pub encrypted_verifier_hash: [u8; 16],
}

/// Parse the 48-byte body that follows `wEncryptionType/vMajor/vMinor`
/// in a legacy RC4 FilePass record.
pub(crate) fn parse_filepass_body(body_after_header: &[u8]) -> CryptoResult<Rc4LegacyParams> {
    if body_after_header.len() < 48 {
        return Err(CryptoError::InvalidFormat(format!(
            "legacy RC4 FilePass body too short: {} (need 48)",
            body_after_header.len()
        )));
    }
    let mut salt = [0u8; 16];
    salt.copy_from_slice(&body_after_header[0..16]);
    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&body_after_header[16..32]);
    let mut encrypted_verifier_hash = [0u8; 16];
    encrypted_verifier_hash.copy_from_slice(&body_after_header[32..48]);
    Ok(Rc4LegacyParams {
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

/// Derive the RC4 key for a given 1024-byte block index.
pub fn make_key(password: &str, salt: &[u8; 16], block: u32) -> [u8; 16] {
    let pw_utf16 = utf16le_bytes(password);

    // h0 = MD5(password_utf16le)
    let mut md = Md5::new();
    md.update(&pw_utf16);
    let h0 = md.finalize();

    // intermediate = (h0[..5] || salt) * 16  → 336 bytes
    let mut intermediate = Vec::with_capacity(16 * (5 + 16));
    for _ in 0..16 {
        intermediate.extend_from_slice(&h0[..5]);
        intermediate.extend_from_slice(salt);
    }

    // h1 = MD5(intermediate)
    let mut md = Md5::new();
    md.update(&intermediate);
    let h1 = md.finalize();

    // h_final = MD5(h1[..5] || u32le(block))
    let mut md = Md5::new();
    md.update(&h1[..5]);
    md.update(block.to_le_bytes());
    let h_final = md.finalize();

    let mut key = [0u8; 16];
    key.copy_from_slice(&h_final);
    key
}

/// Check that `password` decrypts the FilePass verifier correctly.
pub fn verify_password(params: &Rc4LegacyParams, password: &str) -> bool {
    let key = make_key(password, &params.salt, 0);
    let mut rc4 = Rc4::new_from_slice(&key).expect("16-byte key is valid for RC4");

    let mut verifier = params.encrypted_verifier;
    rc4.apply_keystream(&mut verifier);

    let mut verifier_hash = params.encrypted_verifier_hash;
    rc4.apply_keystream(&mut verifier_hash);

    let mut md = Md5::new();
    md.update(verifier);
    let computed = md.finalize();

    computed.ct_eq(&verifier_hash).into()
}

/// Apply the legacy RC4 keystream to a byte buffer, re-keying every
/// 1024 bytes per MS-OFFCRYPTO §2.3.6.
///
/// RC4 is its own inverse, so this function is symmetric: feeding it
/// plaintext yields ciphertext and vice versa. Every block gets a fresh
/// RC4 instance keyed off `make_key(password, salt, block_idx)`.
pub fn apply_keystream(password: &str, params: &Rc4LegacyParams, input: &[u8]) -> Vec<u8> {
    const BLOCK_SIZE: usize = 1024;
    let mut out = Vec::with_capacity(input.len());
    for (block_idx, chunk) in input.chunks(BLOCK_SIZE).enumerate() {
        let key = make_key(password, &params.salt, block_idx as u32);
        let mut rc4 = Rc4::new_from_slice(&key).expect("16-byte key is valid for RC4");
        let mut buf = chunk.to_vec();
        rc4.apply_keystream(&mut buf);
        out.extend_from_slice(&buf);
    }
    out
}

/// Decrypt an XLS Workbook stream encrypted with legacy RC4 (MD5 KDF).
///
/// Output length equals input length; the FilePass record body is
/// zeroed in the output so subsequent parsers treat the stream as
/// plaintext.
pub fn decrypt_workbook_stream(
    stream: &[u8],
    password: &str,
    params: &Rc4LegacyParams,
) -> CryptoResult<Vec<u8>> {
    if !verify_password(params, password) {
        return Err(CryptoError::BadPassword);
    }

    let classified = record_walk::classify(stream)?;
    let decrypted = apply_keystream(password, params, &classified.ciphertext);
    Ok(record_walk::apply_overlay(decrypted, &classified.overlay))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn make_key_is_deterministic() {
        let salt = [0x11u8; 16];
        let k1 = make_key("test", &salt, 0);
        let k2 = make_key("test", &salt, 0);
        assert_eq!(k1, k2);
        assert_eq!(k1.len(), 16);
    }

    #[test]
    fn make_key_varies_per_block() {
        let salt = [0u8; 16];
        let k0 = make_key("pw", &salt, 0);
        let k1 = make_key("pw", &salt, 1);
        assert_ne!(k0, k1);
    }

    #[test]
    fn make_key_differs_from_cryptoapi_variant() {
        // Safety net: if the MD5 path and the SHA-1 path ever accidentally
        // return the same key for the same inputs, something is very wrong.
        let salt = [0u8; 16];
        let legacy = make_key("pw", &salt, 0);
        let cryptoapi = crate::xls::rc4_cryptoapi::make_key("pw", &salt, 128, 0);
        assert_ne!(&legacy[..], &cryptoapi[..]);
    }

    #[test]
    fn verify_password_round_trips() {
        let password = "duke-test-pw";
        let salt = [
            0xAA, 0xBB, 0xCC, 0xDD, 0xEE, 0xFF, 0x00, 0x11, //
            0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99,
        ];

        let verifier = [0x13u8; 16];
        let mut md = Md5::new();
        md.update(verifier);
        let verifier_hash: [u8; 16] = md.finalize().into();

        let key = make_key(password, &salt, 0);
        let mut rc4 = Rc4::new_from_slice(&key).unwrap();
        let mut enc_verifier = verifier;
        rc4.apply_keystream(&mut enc_verifier);
        let mut enc_verifier_hash = verifier_hash;
        rc4.apply_keystream(&mut enc_verifier_hash);

        let params = Rc4LegacyParams {
            salt,
            encrypted_verifier: enc_verifier,
            encrypted_verifier_hash: enc_verifier_hash,
        };
        assert!(verify_password(&params, password));
        assert!(!verify_password(&params, "wrong"));
    }

    #[test]
    fn parse_filepass_body_rejects_short_input() {
        let short = [0u8; 47];
        assert!(parse_filepass_body(&short).is_err());
    }

    #[test]
    fn parse_filepass_body_accepts_48_bytes() {
        let body = [0u8; 48];
        let params = parse_filepass_body(&body).expect("parse ok");
        assert_eq!(params.salt, [0u8; 16]);
    }

    #[test]
    fn apply_keystream_rekeys_every_block() {
        let params = Rc4LegacyParams {
            salt: [0u8; 16],
            encrypted_verifier: [0u8; 16],
            encrypted_verifier_hash: [0u8; 16],
        };
        let plaintext = vec![0x41u8; 2048];
        let encrypted = apply_keystream("pw", &params, &plaintext);
        assert_ne!(
            &encrypted[..1024],
            &encrypted[1024..],
            "per-block re-keying should yield different ciphertext for identical plaintext"
        );
        let decrypted = apply_keystream("pw", &params, &encrypted);
        assert_eq!(
            decrypted, plaintext,
            "RC4 is symmetric: round-trip recovers input"
        );
    }
}
