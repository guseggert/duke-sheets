//! ECMA-376 Standard Encryption (MS-OFFCRYPTO §2.3.4.5–§2.3.4.9).
//!
//! Office 2007-era encryption. LibreOffice still defaults to this for
//! password-protected `.xlsx` (via `CryptoType=StrongEncryptionDataSpace`).
//! Distinct from Agile (Office 2010+):
//!
//! - Binary EncryptionHeader (vs Agile's XML descriptor)
//! - **AES-ECB**, not CBC. Yes, ECB on a whole package.
//! - SHA-1 KDF (Agile defaults to SHA-512)
//! - 50,000 iterations (Agile spec ceiling 10M)
//! - No HMAC integrity check
//! - AES-128 / AES-192 / AES-256 selectable by EncryptionHeader.KeySize
//!
//! # KDF (with the post-iteration XOR-pad step)
//! ```text
//! H_0     = SHA1(salt || password_utf16le)
//! H_{i+1} = SHA1(u32le(i) || H_i)        for i in [0, 50000)
//! H       = SHA1(H_50000 || u32le(0))    // single extra hash, block=0
//!
//! buf1    = H xored against 0x36 repeated to 64 bytes
//! x1      = SHA1(buf1)
//! buf2    = H xored against 0x5C repeated to 64 bytes
//! x2      = SHA1(buf2)
//!
//! derived_key = (x1 || x2)[..keyBits/8]
//! ```
//! The 0x36/0x5C step is *only* in Standard, not Agile. Forgetting it
//! makes every Standard file fail decryption even though every other
//! piece of the KDF is correct — the most subtle Standard-vs-Agile
//! difference.

use aes::cipher::generic_array::GenericArray;
use aes::cipher::{BlockDecrypt, BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use sha1::{Digest as _, Sha1};
use subtle::ConstantTimeEq;

use crate::error::{CryptoError, CryptoResult};
use crate::password::utf16le_bytes;
use crate::random::random_bytes;

const ITER_COUNT: u32 = 50_000;

/// Parsed Standard EncryptionHeader + EncryptionVerifier.
#[derive(Debug, Clone)]
pub struct StandardDescriptor {
    pub key_size_bits: u32,
    pub salt: [u8; 16],
    pub encrypted_verifier: [u8; 16],
    /// 32 bytes on the wire (AES block-aligned); first 20 are the SHA-1
    /// of the verifier after AES-ECB decryption.
    pub encrypted_verifier_hash: [u8; 32],
}

/// Top-level Standard decrypt entry point.
///
/// `encryption_info` is the full `/EncryptionInfo` stream including the
/// 4-byte version header; `encrypted_package` is the full
/// `/EncryptedPackage` stream including the 8-byte total-size prefix.
pub fn decrypt(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> CryptoResult<Vec<u8>> {
    let descriptor = parse_descriptor(encryption_info)?;
    let key = derive_key(password, &descriptor.salt, descriptor.key_size_bits);
    if !verify_password(&key, &descriptor)? {
        return Err(CryptoError::BadPassword);
    }
    decrypt_package(&key, encrypted_package)
}

/// Parse the binary EncryptionInfo stream for a Standard-encrypted
/// OOXML file. Stream layout, after the 4-byte vMajor/vMinor:
///
/// ```text
///   u32 EncryptionHeaderFlags
///   u32 EncryptionHeaderSize        // length of EncryptionHeader
///   EncryptionHeader (size bytes):
///     u32 Flags
///     u32 SizeExtra                  // must be 0
///     u32 AlgID                      // 0x660E=AES-128, 0x660F=AES-192, 0x6610=AES-256
///     u32 AlgIDHash                  // 0x8004=SHA-1
///     u32 KeySize                    // bits
///     u32 ProviderType               // 0x18 = AES
///     u32 Reserved1
///     u32 Reserved2
///     UTF-16LE CSPName               // null-terminated
///   EncryptionVerifier:
///     u32 SaltSize                   // = 16
///     u8[16] Salt
///     u8[16] EncryptedVerifier
///     u32 VerifierHashSize           // = 20 (SHA-1)
///     u8[32] EncryptedVerifierHash   // 32 to align to AES block
/// ```
fn parse_descriptor(stream: &[u8]) -> CryptoResult<StandardDescriptor> {
    if stream.len() < 4 + 8 {
        return Err(CryptoError::InvalidFormat(
            "Standard EncryptionInfo too short for header".into(),
        ));
    }
    let header_size_off = 4 + 4;
    let header_size = u32::from_le_bytes([
        stream[header_size_off],
        stream[header_size_off + 1],
        stream[header_size_off + 2],
        stream[header_size_off + 3],
    ]) as usize;
    let header_start = header_size_off + 4;
    let header_end = header_start
        .checked_add(header_size)
        .ok_or_else(|| CryptoError::InvalidFormat("Standard header size overflow".into()))?;
    if header_end > stream.len() {
        return Err(CryptoError::InvalidFormat(format!(
            "Standard EncryptionHeader overruns stream: claims {header_size} bytes"
        )));
    }
    let hdr = &stream[header_start..header_end];
    if hdr.len() < 32 {
        return Err(CryptoError::InvalidFormat(format!(
            "Standard EncryptionHeader too short: {}",
            hdr.len()
        )));
    }
    let alg_id = u32::from_le_bytes([hdr[8], hdr[9], hdr[10], hdr[11]]);
    let key_size_bits = u32::from_le_bytes([hdr[16], hdr[17], hdr[18], hdr[19]]);

    let key_size_bits = match (alg_id, key_size_bits) {
        (0x0000_660E, _) | (_, 128) => 128,
        (0x0000_660F, _) | (_, 192) => 192,
        (0x0000_6610, _) | (_, 256) => 256,
        (other_alg, other_keysize) => {
            return Err(CryptoError::UnsupportedVariant(format!(
                "Standard AES algId={other_alg:#010x} keySize={other_keysize}"
            )));
        }
    };

    let verifier = &stream[header_end..];
    if verifier.len() < 4 + 16 + 16 + 4 + 32 {
        return Err(CryptoError::InvalidFormat(format!(
            "Standard EncryptionVerifier too short: {}",
            verifier.len()
        )));
    }
    let salt_size = u32::from_le_bytes([verifier[0], verifier[1], verifier[2], verifier[3]]);
    if salt_size != 16 {
        return Err(CryptoError::InvalidFormat(format!(
            "Standard saltSize={salt_size} (expected 16)"
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
            "Standard verifierHashSize={verifier_hash_size} (expected 20)"
        )));
    }
    let mut encrypted_verifier_hash = [0u8; 32];
    encrypted_verifier_hash.copy_from_slice(&verifier[40..72]);

    Ok(StandardDescriptor {
        key_size_bits,
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

/// Derive the AES key from the password. See the module docstring for
/// the formula; the 0x36/0x5C step is the easy-to-miss part.
fn derive_key(password: &str, salt: &[u8; 16], key_size_bits: u32) -> Vec<u8> {
    let pw_utf16 = utf16le_bytes(password);
    let mut h = Sha1::digest([salt.as_slice(), &pw_utf16].concat()).to_vec();
    for i in 0..ITER_COUNT {
        let mut buf = Vec::with_capacity(4 + h.len());
        buf.extend_from_slice(&i.to_le_bytes());
        buf.extend_from_slice(&h);
        h = Sha1::digest(&buf).to_vec();
    }
    // Single extra hash with block counter 0.
    let mut tail = Vec::with_capacity(h.len() + 4);
    tail.extend_from_slice(&h);
    tail.extend_from_slice(&0u32.to_le_bytes());
    let h = Sha1::digest(&tail);

    // The 0x36/0x5C XOR-pad step. SHA-1 digest is 20 bytes; we extend
    // (with zero-padding) to a 64-byte block and XOR against 0x36 / 0x5C.
    let mut buf1 = [0x36u8; 64];
    let mut buf2 = [0x5Cu8; 64];
    for i in 0..h.len() {
        buf1[i] ^= h[i];
        buf2[i] ^= h[i];
    }
    let x1 = Sha1::digest(buf1);
    let x2 = Sha1::digest(buf2);

    let key_bytes = (key_size_bits / 8) as usize;
    let mut combined = Vec::with_capacity(40);
    combined.extend_from_slice(&x1);
    combined.extend_from_slice(&x2);
    combined[..key_bytes].to_vec()
}

fn verify_password(key: &[u8], d: &StandardDescriptor) -> CryptoResult<bool> {
    let verifier = aes_ecb_decrypt(key, &d.encrypted_verifier)?;
    let stored_hash = aes_ecb_decrypt(key, &d.encrypted_verifier_hash)?;
    let computed = Sha1::digest(&verifier);
    Ok(computed.ct_eq(&stored_hash[..20]).into())
}

fn aes_ecb_decrypt(key: &[u8], ciphertext: &[u8]) -> CryptoResult<Vec<u8>> {
    if ciphertext.len() % 16 != 0 || ciphertext.is_empty() {
        return Err(CryptoError::InvalidFormat(format!(
            "AES-ECB ciphertext length {} not a positive multiple of 16",
            ciphertext.len()
        )));
    }
    let mut out = ciphertext.to_vec();
    match key.len() {
        16 => {
            let cipher = Aes128::new(GenericArray::from_slice(key));
            for chunk in out.chunks_exact_mut(16) {
                let block = GenericArray::from_mut_slice(chunk);
                cipher.decrypt_block(block);
            }
        }
        24 => {
            let cipher = Aes192::new(GenericArray::from_slice(key));
            for chunk in out.chunks_exact_mut(16) {
                let block = GenericArray::from_mut_slice(chunk);
                cipher.decrypt_block(block);
            }
        }
        32 => {
            let cipher = Aes256::new(GenericArray::from_slice(key));
            for chunk in out.chunks_exact_mut(16) {
                let block = GenericArray::from_mut_slice(chunk);
                cipher.decrypt_block(block);
            }
        }
        n => {
            return Err(CryptoError::UnsupportedVariant(format!(
                "AES key length {n} not in {{16, 24, 32}}"
            )));
        }
    }
    Ok(out)
}

fn aes_ecb_encrypt(key: &[u8], plaintext: &[u8]) -> CryptoResult<Vec<u8>> {
    if plaintext.len() % 16 != 0 || plaintext.is_empty() {
        return Err(CryptoError::InvalidFormat(format!(
            "AES-ECB plaintext length {} not a positive multiple of 16",
            plaintext.len()
        )));
    }
    let mut out = plaintext.to_vec();
    match key.len() {
        16 => {
            let cipher = Aes128::new(GenericArray::from_slice(key));
            for chunk in out.chunks_exact_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(chunk));
            }
        }
        24 => {
            let cipher = Aes192::new(GenericArray::from_slice(key));
            for chunk in out.chunks_exact_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(chunk));
            }
        }
        32 => {
            let cipher = Aes256::new(GenericArray::from_slice(key));
            for chunk in out.chunks_exact_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(chunk));
            }
        }
        n => {
            return Err(CryptoError::UnsupportedVariant(format!(
                "AES key length {n} not in {{16, 24, 32}}"
            )));
        }
    }
    Ok(out)
}

/// Caller-tunable parameters for [`encrypt`].
#[derive(Debug, Clone)]
pub struct StandardWriteOptions {
    /// AES key size in bits. Standard supports 128, 192, 256.
    pub key_bits: u32,
}

impl Default for StandardWriteOptions {
    fn default() -> Self {
        Self { key_bits: 256 }
    }
}

/// The two top-level streams the caller must wrap in a CFB envelope.
/// Standard encryption does NOT require the `\x06DataSpaces` tree —
/// LibreOffice produces Standard files with just these two streams,
/// and our reader (and Office's) accepts them.
#[derive(Debug)]
pub struct StandardEnvelopeParts {
    pub encryption_info: Vec<u8>,
    pub encrypted_package: Vec<u8>,
}

/// Encrypt with ECMA-376 Standard encryption (AES-ECB + SHA-1 KDF).
///
/// Returns the parts of a CFB envelope; the caller is responsible for
/// CFB packaging. See the module docstring for KDF details and
/// [`StandardWriteOptions`] for parameters.
pub fn encrypt(
    plaintext: &[u8],
    password: &str,
    opts: &StandardWriteOptions,
) -> CryptoResult<StandardEnvelopeParts> {
    if !matches!(opts.key_bits, 128 | 192 | 256) {
        return Err(CryptoError::UnsupportedVariant(format!(
            "Standard keyBits={} (only 128/192/256 supported)",
            opts.key_bits
        )));
    }

    let salt_vec = random_bytes(16)?;
    let mut salt = [0u8; 16];
    salt.copy_from_slice(&salt_vec);

    let key = derive_key(password, &salt, opts.key_bits);

    let verifier_input_vec = random_bytes(16)?;
    let mut verifier_input = [0u8; 16];
    verifier_input.copy_from_slice(&verifier_input_vec);

    let verifier_hash = Sha1::digest(verifier_input);
    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..20].copy_from_slice(&verifier_hash);

    let encrypted_verifier = aes_ecb_encrypt(&key, &verifier_input)?;
    let encrypted_verifier_hash = aes_ecb_encrypt(&key, &verifier_hash_padded)?;

    let total_size = plaintext.len();
    let padded = pad_to_block(plaintext, 16);
    let encrypted_payload = if padded.is_empty() {
        Vec::new()
    } else {
        aes_ecb_encrypt(&key, &padded)?
    };
    let mut encrypted_package = Vec::with_capacity(8 + encrypted_payload.len());
    encrypted_package.extend_from_slice(&(total_size as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&encrypted_payload);

    let encryption_info = build_encryption_info_stream(
        opts.key_bits,
        &salt,
        &encrypted_verifier,
        &encrypted_verifier_hash,
    );

    Ok(StandardEnvelopeParts {
        encryption_info,
        encrypted_package,
    })
}

fn pad_to_block(data: &[u8], block: usize) -> Vec<u8> {
    let rem = data.len() % block;
    if rem == 0 {
        data.to_vec()
    } else {
        let mut v = Vec::with_capacity(data.len() + block - rem);
        v.extend_from_slice(data);
        v.resize(data.len() + (block - rem), 0);
        v
    }
}

/// Build the full `/EncryptionInfo` stream for Standard encryption:
/// 4-byte version header (vMajor=4, vMinor=2) + EncryptionHeaderFlags +
/// EncryptionHeaderSize + EncryptionHeader + EncryptionVerifier.
fn build_encryption_info_stream(
    key_bits: u32,
    salt: &[u8; 16],
    encrypted_verifier: &[u8],
    encrypted_verifier_hash: &[u8],
) -> Vec<u8> {
    // [MS-OFFCRYPTO] §2.3.4.4 EncryptionHeader.Flags:
    //   bit 2 (0x04) fCryptoAPI
    //   bit 5 (0x20) fAES
    const FLAGS: u32 = 0x24;

    let alg_id = match key_bits {
        128 => 0x0000_660Eu32,
        192 => 0x0000_660Fu32,
        256 => 0x0000_6610u32,
        _ => unreachable!("guarded in encrypt()"),
    };

    let csp_name_utf16: Vec<u16> = "Microsoft Enhanced RSA and AES Cryptographic Provider"
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();
    let csp_name_bytes: Vec<u8> = csp_name_utf16
        .iter()
        .flat_map(|u| u.to_le_bytes())
        .collect();

    let mut header = Vec::new();
    header.extend_from_slice(&FLAGS.to_le_bytes());
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&alg_id.to_le_bytes());
    header.extend_from_slice(&0x0000_8004u32.to_le_bytes()); // AlgIDHash = SHA-1
    header.extend_from_slice(&key_bits.to_le_bytes());
    header.extend_from_slice(&0x0000_0018u32.to_le_bytes()); // ProviderType = AES
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
    header.extend_from_slice(&csp_name_bytes);

    let mut out = Vec::new();
    out.extend_from_slice(&4u16.to_le_bytes()); // vMajor
    out.extend_from_slice(&2u16.to_le_bytes()); // vMinor
    out.extend_from_slice(&FLAGS.to_le_bytes()); // EncryptionHeaderFlags
    out.extend_from_slice(&(header.len() as u32).to_le_bytes()); // EncryptionHeaderSize
    out.extend_from_slice(&header);

    out.extend_from_slice(&16u32.to_le_bytes()); // SaltSize
    out.extend_from_slice(salt);
    out.extend_from_slice(encrypted_verifier);
    out.extend_from_slice(&20u32.to_le_bytes()); // VerifierHashSize (SHA-1)
    out.extend_from_slice(encrypted_verifier_hash);

    out
}

fn decrypt_package(key: &[u8], encrypted_package: &[u8]) -> CryptoResult<Vec<u8>> {
    if encrypted_package.len() < 8 {
        return Err(CryptoError::InvalidFormat(
            "Standard EncryptedPackage shorter than 8-byte size prefix".into(),
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
    if ciphertext.is_empty() {
        return Ok(Vec::new());
    }
    let mut decrypted = aes_ecb_decrypt(key, ciphertext)?;
    decrypted.truncate(total_size);
    Ok(decrypted)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn iter_count_is_50000() {
        assert_eq!(ITER_COUNT, 50_000);
    }

    #[test]
    fn derive_key_is_deterministic() {
        let salt = [0x42u8; 16];
        let k1 = derive_key("test", &salt, 128);
        let k2 = derive_key("test", &salt, 128);
        assert_eq!(k1, k2);
        assert_eq!(k1.len(), 16);
    }

    #[test]
    fn derive_key_varies_with_password() {
        let salt = [0u8; 16];
        let k1 = derive_key("pw1", &salt, 128);
        let k2 = derive_key("pw2", &salt, 128);
        assert_ne!(k1, k2);
    }

    #[test]
    fn derive_key_supports_aes_128_192_256() {
        let salt = [0u8; 16];
        assert_eq!(derive_key("pw", &salt, 128).len(), 16);
        assert_eq!(derive_key("pw", &salt, 192).len(), 24);
        assert_eq!(derive_key("pw", &salt, 256).len(), 32);
    }

    /// AES-256 derivation must use both x1 and x2 (combined exceeds the
    /// 20-byte SHA-1 digest size); this test guards against the
    /// implementation accidentally truncating to x1 only.
    #[test]
    fn aes_256_key_uses_combined_x1_x2() {
        let salt = [0u8; 16];
        let k128 = derive_key("pw", &salt, 128);
        let k256 = derive_key("pw", &salt, 256);
        assert_eq!(k256[..16], k128[..]);
        assert_ne!(&k256[16..], &[0u8; 16]);
    }

    #[test]
    fn verify_password_round_trips() {
        let password = "duke-test-pw";
        let salt = [
            0xCAu8, 0xFE, 0xBA, 0xBE, 0xDE, 0xAD, 0xBE, 0xEF, 0x12, 0x34, 0x56, 0x78, 0x9A, 0xBC,
            0xDE, 0xF0,
        ];

        let key = derive_key(password, &salt, 128);
        let verifier_plain = [0x77u8; 16];
        let verifier_hash = Sha1::digest(verifier_plain);
        let mut verifier_hash_padded = [0u8; 32];
        verifier_hash_padded[..20].copy_from_slice(&verifier_hash);

        // Encrypt both with AES-ECB so verify_password can decrypt them.
        let encrypted_verifier = encrypt_ecb_for_test(&key, &verifier_plain);
        let encrypted_verifier_hash = encrypt_ecb_for_test(&key, &verifier_hash_padded);
        let mut ev = [0u8; 16];
        ev.copy_from_slice(&encrypted_verifier);
        let mut evh = [0u8; 32];
        evh.copy_from_slice(&encrypted_verifier_hash);

        let d = StandardDescriptor {
            key_size_bits: 128,
            salt,
            encrypted_verifier: ev,
            encrypted_verifier_hash: evh,
        };
        let key2 = derive_key(password, &d.salt, d.key_size_bits);
        assert!(verify_password(&key2, &d).unwrap());
        let wrong_key = derive_key("wrong", &d.salt, d.key_size_bits);
        assert!(!verify_password(&wrong_key, &d).unwrap());
    }

    /// AES-ECB encrypt for test setup. Production code only decrypts.
    fn encrypt_ecb_for_test(key: &[u8], plaintext: &[u8]) -> Vec<u8> {
        use aes::cipher::BlockEncrypt;
        assert_eq!(plaintext.len() % 16, 0);
        let mut out = plaintext.to_vec();
        match key.len() {
            16 => {
                let cipher = Aes128::new(GenericArray::from_slice(key));
                for chunk in out.chunks_exact_mut(16) {
                    cipher.encrypt_block(GenericArray::from_mut_slice(chunk));
                }
            }
            32 => {
                let cipher = Aes256::new(GenericArray::from_slice(key));
                for chunk in out.chunks_exact_mut(16) {
                    cipher.encrypt_block(GenericArray::from_mut_slice(chunk));
                }
            }
            _ => panic!("unsupported test key length"),
        }
        out
    }
}
