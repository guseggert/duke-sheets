//! OOXML Binary Document RC4 CryptoAPI (MS-OFFCRYPTO §2.3.5).
//!
//! Same binary `EncryptionHeader`/`EncryptionVerifier` layout as Standard
//! Encryption (§2.3.4.5) but with `AlgID=0x6801` (RC4) and
//! `ProviderType=0x0001`. Same SHA-1 KDF as our XLS RC4 CryptoAPI
//! variant — we reuse `xls::rc4_cryptoapi::make_key` and
//! `apply_keystream` since the cipher math is identical; only the
//! re-key block size and verifier hash size differ:
//!
//! | Variant | Re-key boundary | Verifier hash size |
//! |---|---|---|
//! | XLS RC4 CryptoAPI | 1024 bytes | 32 (zero-padded) |
//! | OOXML Binary RC4 | **512 bytes** | **20** (no padding, since RC4 has no block) |
//!
//! In practice this variant is rare; modern Office uses Agile instead.
//! Some legacy corporate tools still produce it.

use rc4::{KeyInit, Rc4, StreamCipher};
use sha1::{Digest as _, Sha1};
use subtle::ConstantTimeEq;

use crate::error::{CryptoError, CryptoResult};
use crate::xls::rc4_cryptoapi::{apply_keystream, make_key};

/// Parsed Binary RC4 EncryptionHeader + EncryptionVerifier.
#[derive(Debug, Clone)]
pub struct BinaryRc4Descriptor {
    pub key_size_bits: u32,
    pub salt: [u8; 16],
    pub encrypted_verifier: [u8; 16],
    /// 20 bytes on the wire — one SHA-1 digest, unpadded.
    pub encrypted_verifier_hash: [u8; 20],
}

pub fn decrypt(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> CryptoResult<Vec<u8>> {
    let descriptor = parse_descriptor(encryption_info)?;
    if !verify_password(password, &descriptor) {
        return Err(CryptoError::BadPassword);
    }
    decrypt_package(password, &descriptor, encrypted_package)
}

fn parse_descriptor(stream: &[u8]) -> CryptoResult<BinaryRc4Descriptor> {
    if stream.len() < 4 + 8 {
        return Err(CryptoError::InvalidFormat(
            "Binary RC4 EncryptionInfo too short".into(),
        ));
    }
    // Skip vMajor/vMinor (4 bytes), then EncryptionHeaderFlags (4 bytes),
    // then read the EncryptionHeaderSize (4 bytes).
    let header_size = u32::from_le_bytes([stream[8], stream[9], stream[10], stream[11]]) as usize;
    let header_start = 12usize;
    let header_end = header_start
        .checked_add(header_size)
        .ok_or_else(|| CryptoError::InvalidFormat("Binary RC4 header size overflow".into()))?;
    if header_end > stream.len() {
        return Err(CryptoError::InvalidFormat(format!(
            "Binary RC4 EncryptionHeader overruns: claims {header_size} bytes"
        )));
    }
    let hdr = &stream[header_start..header_end];
    if hdr.len() < 32 {
        return Err(CryptoError::InvalidFormat(format!(
            "Binary RC4 EncryptionHeader too short: {}",
            hdr.len()
        )));
    }
    let alg_id = u32::from_le_bytes([hdr[8], hdr[9], hdr[10], hdr[11]]);
    let raw_key_size = u32::from_le_bytes([hdr[16], hdr[17], hdr[18], hdr[19]]);
    if alg_id != 0x0000_6801 && alg_id != 0 {
        return Err(CryptoError::UnsupportedVariant(format!(
            "Binary RC4 algId={alg_id:#010x} (expected 0x6801 or 0)"
        )));
    }
    let key_size_bits = if raw_key_size == 0 { 40 } else { raw_key_size };

    let verifier = &stream[header_end..];
    if verifier.len() < 4 + 16 + 16 + 4 + 20 {
        return Err(CryptoError::InvalidFormat(format!(
            "Binary RC4 EncryptionVerifier too short: {}",
            verifier.len()
        )));
    }
    let salt_size = u32::from_le_bytes([verifier[0], verifier[1], verifier[2], verifier[3]]);
    if salt_size != 16 {
        return Err(CryptoError::InvalidFormat(format!(
            "Binary RC4 saltSize={salt_size} (expected 16)"
        )));
    }
    let mut salt = [0u8; 16];
    salt.copy_from_slice(&verifier[4..20]);
    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&verifier[20..36]);
    let verifier_hash_size =
        u32::from_le_bytes([verifier[36], verifier[37], verifier[38], verifier[39]]);
    if verifier_hash_size != 20 {
        return Err(CryptoError::InvalidFormat(format!(
            "Binary RC4 verifierHashSize={verifier_hash_size} (expected 20)"
        )));
    }
    let mut encrypted_verifier_hash = [0u8; 20];
    encrypted_verifier_hash.copy_from_slice(&verifier[40..60]);

    Ok(BinaryRc4Descriptor {
        key_size_bits,
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

fn verify_password(password: &str, d: &BinaryRc4Descriptor) -> bool {
    let key = make_key(password, &d.salt, d.key_size_bits, 0);
    let mut rc4 = Rc4::new_from_slice(&key[..16]).expect("16-byte key");
    let mut verifier = d.encrypted_verifier;
    rc4.apply_keystream(&mut verifier);
    let mut verifier_hash = d.encrypted_verifier_hash;
    rc4.apply_keystream(&mut verifier_hash);
    let computed = Sha1::digest(verifier);
    computed.ct_eq(&verifier_hash).into()
}

fn decrypt_package(
    password: &str,
    d: &BinaryRc4Descriptor,
    encrypted_package: &[u8],
) -> CryptoResult<Vec<u8>> {
    if encrypted_package.len() < 8 {
        return Err(CryptoError::InvalidFormat(
            "Binary RC4 EncryptedPackage shorter than 8-byte size prefix".into(),
        ));
    }
    let total_size = u64::from_le_bytes([
        encrypted_package[0],
        encrypted_package[1],
        encrypted_package[2],
        encrypted_package[3],
        encrypted_package[4],
        encrypted_package[5],
        encrypted_package[6],
        encrypted_package[7],
    ]) as usize;
    let ciphertext = &encrypted_package[8..];

    // OOXML Binary RC4 re-keys every 512 bytes (vs XLS's 1024).
    const BLOCK_SIZE: usize = 512;
    let mut decrypted = apply_keystream(password, &d.salt, d.key_size_bits, BLOCK_SIZE, ciphertext);
    decrypted.truncate(total_size);
    Ok(decrypted)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Verifier round-trip: build a (verifier, verifierHash) pair using
    /// the same KDF the descriptor parser will look up, then check that
    /// `verify_password` accepts the right password and rejects the wrong
    /// one.
    #[test]
    fn verify_password_round_trips() {
        let password = "duke-test-pw";
        let salt = [0x11u8; 16];
        let key_size_bits = 128u32;

        let key = make_key(password, &salt, key_size_bits, 0);
        let mut rc4 = Rc4::new_from_slice(&key[..16]).unwrap();
        let mut verifier = [0x42u8; 16];
        rc4.apply_keystream(&mut verifier);
        let mut verifier_hash = Sha1::digest([0x42u8; 16]).to_vec();
        rc4.apply_keystream(&mut verifier_hash);

        let mut ev = [0u8; 16];
        ev.copy_from_slice(&verifier);
        let mut evh = [0u8; 20];
        evh.copy_from_slice(&verifier_hash);

        let d = BinaryRc4Descriptor {
            key_size_bits,
            salt,
            encrypted_verifier: ev,
            encrypted_verifier_hash: evh,
        };
        assert!(verify_password(password, &d));
        assert!(!verify_password("wrong", &d));
    }
}
