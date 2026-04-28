//! ECMA-376 Agile encryption (MS-OFFCRYPTO §2.3.4.10–§2.3.4.15).
//!
//! Used by Office 2010+ for password-protected `.xlsx`. The
//! `/EncryptionInfo` stream contains a 4-byte version header, a 4-byte
//! flags word (always `0x40` for Agile), then an XML descriptor.
//! `/EncryptedPackage` contains an 8-byte little-endian total-plaintext-
//! size prefix followed by AES-CBC ciphertext of the inner ZIP, split
//! into 4096-byte segments each with a derived IV.
//!
//! # KDF
//! Office's iterated hash KDF; SHA-512 by default for modern files:
//! ```text
//! H_0     = H(keyEncryptorSalt || password_utf16le)
//! H_{i+1} = H(u32le(i) || H_i)         for i in [0, spinCount)
//! H_final = H(H_spinCount || blockKey)
//! intermediate_key = H_final[..keyBits/8]
//! ```
//!
//! Block keys (MS-OFFCRYPTO §2.3.4.13):
//!
//! | Constant | Used for |
//! |---|---|
//! | `fea7d2763b4b9e79` | password verifier (input) |
//! | `d7aa0f6d3061344e` | password verifier (hash) |
//! | `146e0be7abacd0d6` | wrapping the data secret key |
//! | `5fb2ad010cb9e1f6` | wrapping the HMAC key |
//! | `a0677f02b22c8433` | wrapping the HMAC value |
//!
//! # Two independent salts
//! - `keyEncryptor.saltValue` (a.k.a. keyEncryptorSalt) — feeds the
//!   password KDF and is the IV for unwrapping the verifier and key.
//! - `keyData.saltValue` (a.k.a. keyDataSalt) — used for per-segment
//!   data IVs and for HMAC key/value unwrap IVs.
//!
//! Treating them as one is the most common implementation bug.

use aes::{Aes128, Aes192, Aes256};
use cbc::cipher::{block_padding::NoPadding, BlockDecryptMut, BlockEncryptMut, KeyIvInit};
use hmac::{Hmac, Mac};
use sha1::Sha1;
use sha2::{Digest as _, Sha256, Sha384, Sha512};
use subtle::ConstantTimeEq;

use crate::error::{CryptoError, CryptoResult};
use crate::password::utf16le_bytes;
use crate::random::random_bytes;

const BLK_VERIFIER_HASH_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const BLK_VERIFIER_HASH_VALUE: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const BLK_ENCRYPTED_KEY_VALUE: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
const BLK_DATA_INTEGRITY_KEY: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
const BLK_DATA_INTEGRITY_VALUE: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

/// Maximum spinCount we accept on read. Spec ceiling per
/// MS-OFFCRYPTO §2.3.4.10. Above this we reject the file as malformed
/// rather than honor a DoS-grade KDF.
const MAX_SPIN_COUNT: u32 = 10_000_000;

/// Hash algorithm used by the Agile KDF. Reading the XML descriptor's
/// `hashAlgorithm` attribute against this enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HashAlgo {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlgo {
    fn parse(s: &str) -> CryptoResult<Self> {
        match s {
            "SHA1" => Ok(Self::Sha1),
            "SHA256" => Ok(Self::Sha256),
            "SHA384" => Ok(Self::Sha384),
            "SHA512" => Ok(Self::Sha512),
            _ => Err(CryptoError::UnsupportedVariant(format!(
                "Agile hashAlgorithm={s}"
            ))),
        }
    }

    pub fn digest_size(self) -> usize {
        match self {
            HashAlgo::Sha1 => 20,
            HashAlgo::Sha256 => 32,
            HashAlgo::Sha384 => 48,
            HashAlgo::Sha512 => 64,
        }
    }

    fn name(self) -> &'static str {
        match self {
            HashAlgo::Sha1 => "SHA1",
            HashAlgo::Sha256 => "SHA256",
            HashAlgo::Sha384 => "SHA384",
            HashAlgo::Sha512 => "SHA512",
        }
    }

    fn hmac(self, key: &[u8], data: &[u8]) -> Vec<u8> {
        match self {
            HashAlgo::Sha1 => {
                let mut m = <Hmac<Sha1> as Mac>::new_from_slice(key).expect("HMAC key any size");
                m.update(data);
                m.finalize().into_bytes().to_vec()
            }
            HashAlgo::Sha256 => {
                let mut m = <Hmac<Sha256> as Mac>::new_from_slice(key).expect("HMAC key any size");
                m.update(data);
                m.finalize().into_bytes().to_vec()
            }
            HashAlgo::Sha384 => {
                let mut m = <Hmac<Sha384> as Mac>::new_from_slice(key).expect("HMAC key any size");
                m.update(data);
                m.finalize().into_bytes().to_vec()
            }
            HashAlgo::Sha512 => {
                let mut m = <Hmac<Sha512> as Mac>::new_from_slice(key).expect("HMAC key any size");
                m.update(data);
                m.finalize().into_bytes().to_vec()
            }
        }
    }

    fn hash(self, input: &[u8]) -> Vec<u8> {
        match self {
            HashAlgo::Sha1 => Sha1::digest(input).to_vec(),
            HashAlgo::Sha256 => Sha256::digest(input).to_vec(),
            HashAlgo::Sha384 => Sha384::digest(input).to_vec(),
            HashAlgo::Sha512 => Sha512::digest(input).to_vec(),
        }
    }
}

/// One `<keyData>` or `<keyEncryptor>` parameter set.
#[derive(Debug, Clone)]
struct CryptoParams {
    salt_value: Vec<u8>,
    block_size: usize,
    key_bits: u32,
    hash_size: usize,
    cipher_algorithm: String,
    cipher_chaining: String,
    hash_algorithm: HashAlgo,
}

/// Parsed Agile `<encryption>` descriptor.
#[derive(Debug, Clone)]
pub struct AgileDescriptor {
    key_data: CryptoParams,
    encrypted_hmac_key: Vec<u8>,
    encrypted_hmac_value: Vec<u8>,
    key_encryptor: CryptoParams,
    spin_count: u32,
    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
    encrypted_key_value: Vec<u8>,
}

/// Top-level Agile decrypt entry point.
///
/// `encryption_info` is the full `/EncryptionInfo` stream including the
/// 8-byte version/flags header. `encrypted_package` is the full
/// `/EncryptedPackage` stream including the 8-byte total-size prefix.
///
/// Verifies the data-integrity HMAC over the encrypted package after
/// successful password unwrap; returns
/// [`CryptoError::IntegrityCheckFailed`] on mismatch. Use
/// [`decrypt_with_options`] to opt out of the check.
pub fn decrypt(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> CryptoResult<Vec<u8>> {
    decrypt_with_options(
        encryption_info,
        encrypted_package,
        password,
        &AgileReadOptions::default(),
    )
}

/// Caller-tunable parameters for Agile decryption.
#[derive(Debug, Clone, Default)]
pub struct AgileReadOptions {
    /// Skip the HMAC integrity check after decryption. Off by default
    /// (matching Office behaviour). Turn on for forensic / recovery
    /// scenarios where you want to read damaged or tampered files.
    pub skip_integrity_check: bool,
}

/// Agile decrypt with explicit options. See [`decrypt`] and
/// [`AgileReadOptions`].
pub fn decrypt_with_options(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
    opts: &AgileReadOptions,
) -> CryptoResult<Vec<u8>> {
    if encryption_info.len() < 8 {
        return Err(CryptoError::InvalidFormat(
            "Agile EncryptionInfo missing 8-byte header".into(),
        ));
    }
    let descriptor = parse_descriptor(&encryption_info[8..])?;
    let secret_key = recover_secret_key(&descriptor, password)?;
    if !opts.skip_integrity_check {
        verify_hmac(&descriptor, &secret_key, encrypted_package)?;
    }
    decrypt_package(&descriptor.key_data, &secret_key, encrypted_package)
}

fn verify_hmac(
    d: &AgileDescriptor,
    secret_key: &[u8],
    encrypted_package: &[u8],
) -> CryptoResult<()> {
    let block_size = d.key_data.block_size as usize;
    let iv_hmac_key = derive_iv_from_salt(
        &d.key_data.salt_value,
        &BLK_DATA_INTEGRITY_KEY,
        d.key_data.hash_algorithm,
        block_size,
    );
    let iv_hmac_value = derive_iv_from_salt(
        &d.key_data.salt_value,
        &BLK_DATA_INTEGRITY_VALUE,
        d.key_data.hash_algorithm,
        block_size,
    );
    let unwrapped_key = aes_cbc_decrypt(secret_key, &iv_hmac_key, &d.encrypted_hmac_key)?;
    let unwrapped_value = aes_cbc_decrypt(secret_key, &iv_hmac_value, &d.encrypted_hmac_value)?;
    let hash_size = d.key_data.hash_algorithm.digest_size();
    if unwrapped_key.len() < hash_size || unwrapped_value.len() < hash_size {
        return Err(CryptoError::InvalidFormat(format!(
            "HMAC key/value too short after unwrap: key={} value={} hash_size={}",
            unwrapped_key.len(),
            unwrapped_value.len(),
            hash_size
        )));
    }
    let hmac_key = &unwrapped_key[..hash_size];
    let stored_value = &unwrapped_value[..hash_size];
    let computed = d.key_data.hash_algorithm.hmac(hmac_key, encrypted_package);
    let ok: bool = computed.ct_eq(stored_value).into();
    if !ok {
        return Err(CryptoError::IntegrityCheckFailed);
    }
    Ok(())
}

/// Parse the XML descriptor that follows the 8-byte EncryptionInfo
/// header.
fn parse_descriptor(xml: &[u8]) -> CryptoResult<AgileDescriptor> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);

    let mut key_data: Option<CryptoParams> = None;
    let mut encrypted_hmac_key: Option<Vec<u8>> = None;
    let mut encrypted_hmac_value: Option<Vec<u8>> = None;
    let mut key_encryptor: Option<CryptoParams> = None;
    let mut spin_count: Option<u32> = None;
    let mut encrypted_verifier_hash_input: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_value: Option<Vec<u8>> = None;
    let mut encrypted_key_value: Option<Vec<u8>> = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"keyData" => key_data = Some(parse_crypto_params(&e)?),
                    b"dataIntegrity" => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"encryptedHmacKey" => {
                                    encrypted_hmac_key = Some(decode_b64(&attr.value)?);
                                }
                                b"encryptedHmacValue" => {
                                    encrypted_hmac_value = Some(decode_b64(&attr.value)?);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"encryptedKey" => {
                        let params = parse_crypto_params(&e)?;
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"spinCount" => {
                                    let s = std::str::from_utf8(&attr.value).map_err(|_| {
                                        CryptoError::InvalidFormat("spinCount not utf8".into())
                                    })?;
                                    spin_count = Some(s.parse().map_err(|_| {
                                        CryptoError::InvalidFormat(format!(
                                            "spinCount not a number: {s}"
                                        ))
                                    })?);
                                }
                                b"encryptedVerifierHashInput" => {
                                    encrypted_verifier_hash_input = Some(decode_b64(&attr.value)?);
                                }
                                b"encryptedVerifierHashValue" => {
                                    encrypted_verifier_hash_value = Some(decode_b64(&attr.value)?);
                                }
                                b"encryptedKeyValue" => {
                                    encrypted_key_value = Some(decode_b64(&attr.value)?);
                                }
                                _ => {}
                            }
                        }
                        key_encryptor = Some(params);
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(CryptoError::InvalidFormat(format!(
                    "Agile EncryptionInfo XML parse error: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    let key_data = key_data.ok_or_else(|| {
        CryptoError::InvalidFormat("Agile EncryptionInfo missing <keyData>".into())
    })?;
    let key_encryptor = key_encryptor.ok_or_else(|| {
        CryptoError::InvalidFormat(
            "Agile EncryptionInfo missing <keyEncryptors>/<keyEncryptor>/<encryptedKey>".into(),
        )
    })?;
    let spin_count = spin_count
        .ok_or_else(|| CryptoError::InvalidFormat("Agile encryptedKey missing spinCount".into()))?;
    if spin_count > MAX_SPIN_COUNT {
        return Err(CryptoError::InvalidFormat(format!(
            "Agile spinCount={spin_count} exceeds MS-OFFCRYPTO §2.3.4.10 limit of {MAX_SPIN_COUNT}"
        )));
    }

    Ok(AgileDescriptor {
        key_data,
        encrypted_hmac_key: encrypted_hmac_key.unwrap_or_default(),
        encrypted_hmac_value: encrypted_hmac_value.unwrap_or_default(),
        key_encryptor,
        spin_count,
        encrypted_verifier_hash_input: encrypted_verifier_hash_input.ok_or_else(|| {
            CryptoError::InvalidFormat(
                "Agile encryptedKey missing encryptedVerifierHashInput".into(),
            )
        })?,
        encrypted_verifier_hash_value: encrypted_verifier_hash_value.ok_or_else(|| {
            CryptoError::InvalidFormat(
                "Agile encryptedKey missing encryptedVerifierHashValue".into(),
            )
        })?,
        encrypted_key_value: encrypted_key_value.ok_or_else(|| {
            CryptoError::InvalidFormat("Agile encryptedKey missing encryptedKeyValue".into())
        })?,
    })
}

fn parse_crypto_params(e: &quick_xml::events::BytesStart<'_>) -> CryptoResult<CryptoParams> {
    let mut salt_value = Vec::new();
    let mut block_size = 0usize;
    let mut key_bits = 0u32;
    let mut hash_size = 0usize;
    let mut cipher_algorithm = String::new();
    let mut cipher_chaining = String::new();
    let mut hash_algorithm = HashAlgo::Sha512;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"saltValue" => salt_value = decode_b64(&attr.value)?,
            b"blockSize" => {
                block_size = parse_u32(&attr.value)? as usize;
            }
            b"keyBits" => {
                key_bits = parse_u32(&attr.value)?;
            }
            b"hashSize" => {
                hash_size = parse_u32(&attr.value)? as usize;
            }
            b"cipherAlgorithm" => {
                cipher_algorithm = std::str::from_utf8(&attr.value)
                    .map_err(|_| CryptoError::InvalidFormat("cipherAlgorithm not utf8".into()))?
                    .to_string();
            }
            b"cipherChaining" => {
                cipher_chaining = std::str::from_utf8(&attr.value)
                    .map_err(|_| CryptoError::InvalidFormat("cipherChaining not utf8".into()))?
                    .to_string();
            }
            b"hashAlgorithm" => {
                let s = std::str::from_utf8(&attr.value)
                    .map_err(|_| CryptoError::InvalidFormat("hashAlgorithm not utf8".into()))?;
                hash_algorithm = HashAlgo::parse(s)?;
            }
            _ => {}
        }
    }

    Ok(CryptoParams {
        salt_value,
        block_size,
        key_bits,
        hash_size,
        cipher_algorithm,
        cipher_chaining,
        hash_algorithm,
    })
}

fn parse_u32(bytes: &[u8]) -> CryptoResult<u32> {
    let s = std::str::from_utf8(bytes)
        .map_err(|_| CryptoError::InvalidFormat("attribute not utf8".into()))?;
    s.parse()
        .map_err(|_| CryptoError::InvalidFormat(format!("attribute not a number: {s}")))
}

fn decode_b64(bytes: &[u8]) -> CryptoResult<Vec<u8>> {
    let s = std::str::from_utf8(bytes)
        .map_err(|_| CryptoError::InvalidFormat("attribute not utf8".into()))?;
    decode_base64(s)
}

/// Tiny base64 decoder. quick-xml does not include base64 support and
/// the standard `base64` crate isn't already in the workspace; rolling
/// our own keeps the dependency footprint of the crypto crate minimal.
/// Accepts only standard base64 (`+/`), whitespace-skipping, with `=`
/// padding. Rejects other characters with `InvalidFormat`.
fn decode_base64(s: &str) -> CryptoResult<Vec<u8>> {
    fn lookup(c: u8) -> Option<u8> {
        match c {
            b'A'..=b'Z' => Some(c - b'A'),
            b'a'..=b'z' => Some(c - b'a' + 26),
            b'0'..=b'9' => Some(c - b'0' + 52),
            b'+' => Some(62),
            b'/' => Some(63),
            _ => None,
        }
    }

    let cleaned: Vec<u8> = s
        .bytes()
        .filter(|c| !c.is_ascii_whitespace() && *c != b'=')
        .collect();
    let mut out = Vec::with_capacity(cleaned.len() * 3 / 4 + 2);
    let mut chunk = [0u8; 4];
    let mut chunk_len = 0usize;
    for c in cleaned {
        let v = lookup(c)
            .ok_or_else(|| CryptoError::InvalidFormat(format!("invalid base64 char {c:#x}")))?;
        chunk[chunk_len] = v;
        chunk_len += 1;
        if chunk_len == 4 {
            out.push((chunk[0] << 2) | (chunk[1] >> 4));
            out.push((chunk[1] << 4) | (chunk[2] >> 2));
            out.push((chunk[2] << 6) | chunk[3]);
            chunk_len = 0;
        }
    }
    match chunk_len {
        0 => {}
        2 => {
            out.push((chunk[0] << 2) | (chunk[1] >> 4));
        }
        3 => {
            out.push((chunk[0] << 2) | (chunk[1] >> 4));
            out.push((chunk[1] << 4) | (chunk[2] >> 2));
        }
        _ => {
            return Err(CryptoError::InvalidFormat("truncated base64 input".into()));
        }
    }
    Ok(out)
}

/// Office iterated-hash KDF (MS-OFFCRYPTO §2.3.4.11 + §2.3.4.13).
///
/// `salt` is `keyEncryptorSalt` for verifier/key-wrap operations and
/// `keyDataSalt` for HMAC-wrap operations. Returns
/// `intermediate_key[..key_bytes]` after the block-key finalization.
fn derive_intermediate_key(
    password: &str,
    salt: &[u8],
    spin_count: u32,
    block_key: &[u8],
    hash: HashAlgo,
    key_bytes: usize,
) -> Vec<u8> {
    let pw_utf16 = utf16le_bytes(password);

    let mut h = hash.hash(&[salt, &pw_utf16].concat());
    for i in 0..spin_count {
        let mut buf = Vec::with_capacity(4 + h.len());
        buf.extend_from_slice(&i.to_le_bytes());
        buf.extend_from_slice(&h);
        h = hash.hash(&buf);
    }

    let mut buf = Vec::with_capacity(h.len() + block_key.len());
    buf.extend_from_slice(&h);
    buf.extend_from_slice(block_key);
    let h_final = hash.hash(&buf);

    h_final[..key_bytes].to_vec()
}

/// Verify the supplied password and return the secret_key used for the
/// data stream. Matches MS-OFFCRYPTO §2.3.4.13 step-by-step.
fn recover_secret_key(d: &AgileDescriptor, password: &str) -> CryptoResult<Vec<u8>> {
    let key_bytes = (d.key_encryptor.key_bits / 8) as usize;

    let key_vhi = derive_intermediate_key(
        password,
        &d.key_encryptor.salt_value,
        d.spin_count,
        &BLK_VERIFIER_HASH_INPUT,
        d.key_encryptor.hash_algorithm,
        key_bytes,
    );
    let key_vhv = derive_intermediate_key(
        password,
        &d.key_encryptor.salt_value,
        d.spin_count,
        &BLK_VERIFIER_HASH_VALUE,
        d.key_encryptor.hash_algorithm,
        key_bytes,
    );
    let key_ekv = derive_intermediate_key(
        password,
        &d.key_encryptor.salt_value,
        d.spin_count,
        &BLK_ENCRYPTED_KEY_VALUE,
        d.key_encryptor.hash_algorithm,
        key_bytes,
    );

    let verifier = aes_cbc_decrypt(
        &key_vhi,
        &d.key_encryptor.salt_value,
        &d.encrypted_verifier_hash_input,
    )?;
    let expected_hash = aes_cbc_decrypt(
        &key_vhv,
        &d.key_encryptor.salt_value,
        &d.encrypted_verifier_hash_value,
    )?;

    let computed = d.key_encryptor.hash_algorithm.hash(&verifier);
    let computed_len = computed.len();
    let cmp_len = computed_len.min(expected_hash.len());
    let ok: bool = computed[..cmp_len].ct_eq(&expected_hash[..cmp_len]).into();
    if !ok {
        return Err(CryptoError::BadPassword);
    }

    let secret_key = aes_cbc_decrypt(
        &key_ekv,
        &d.key_encryptor.salt_value,
        &d.encrypted_key_value,
    )?;
    let key_data_bytes = (d.key_data.key_bits / 8) as usize;
    if secret_key.len() < key_data_bytes {
        return Err(CryptoError::InvalidFormat(format!(
            "decrypted secret key is {} bytes, need {}",
            secret_key.len(),
            key_data_bytes
        )));
    }
    Ok(secret_key[..key_data_bytes].to_vec())
}

fn aes_cbc_decrypt(key: &[u8], iv: &[u8], ciphertext: &[u8]) -> CryptoResult<Vec<u8>> {
    if ciphertext.len() % 16 != 0 || ciphertext.is_empty() {
        return Err(CryptoError::InvalidFormat(format!(
            "AES-CBC ciphertext length {} not a positive multiple of 16",
            ciphertext.len()
        )));
    }
    let mut iv_buf = [0u8; 16];
    iv_buf.copy_from_slice(&iv[..iv.len().min(16)]);
    let mut buf = ciphertext.to_vec();

    match key.len() {
        16 => {
            type Aes128CbcDec = cbc::Decryptor<Aes128>;
            let cipher = Aes128CbcDec::new_from_slices(key, &iv_buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-128-CBC init: {e}")))?;
            cipher
                .decrypt_padded_mut::<NoPadding>(&mut buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-128-CBC decrypt: {e}")))?;
        }
        24 => {
            type Aes192CbcDec = cbc::Decryptor<Aes192>;
            let cipher = Aes192CbcDec::new_from_slices(key, &iv_buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-192-CBC init: {e}")))?;
            cipher
                .decrypt_padded_mut::<NoPadding>(&mut buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-192-CBC decrypt: {e}")))?;
        }
        32 => {
            type Aes256CbcDec = cbc::Decryptor<Aes256>;
            let cipher = Aes256CbcDec::new_from_slices(key, &iv_buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-256-CBC init: {e}")))?;
            cipher
                .decrypt_padded_mut::<NoPadding>(&mut buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-256-CBC decrypt: {e}")))?;
        }
        n => {
            return Err(CryptoError::UnsupportedVariant(format!(
                "Agile AES key length {n} not in {{16, 24, 32}}"
            )));
        }
    }
    Ok(buf)
}

/// Decrypt the EncryptedPackage stream using the recovered secret key.
fn decrypt_package(
    key_data: &CryptoParams,
    secret_key: &[u8],
    encrypted_package: &[u8],
) -> CryptoResult<Vec<u8>> {
    if encrypted_package.len() < 8 {
        return Err(CryptoError::InvalidFormat(
            "EncryptedPackage shorter than 8-byte size prefix".into(),
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

    if !key_data.cipher_algorithm.eq_ignore_ascii_case("AES")
        || !key_data
            .cipher_chaining
            .eq_ignore_ascii_case("ChainingModeCBC")
    {
        return Err(CryptoError::UnsupportedVariant(format!(
            "Agile keyData cipher={}/{} (only AES/CBC supported)",
            key_data.cipher_algorithm, key_data.cipher_chaining
        )));
    }

    const SEGMENT: usize = 4096;
    let mut out = Vec::with_capacity(ciphertext.len());
    for (i, segment) in ciphertext.chunks(SEGMENT).enumerate() {
        let mut salt_with_block_key = Vec::with_capacity(key_data.salt_value.len() + 4);
        salt_with_block_key.extend_from_slice(&key_data.salt_value);
        salt_with_block_key.extend_from_slice(&(i as u32).to_le_bytes());
        let mut iv = key_data.hash_algorithm.hash(&salt_with_block_key);
        iv.truncate(key_data.block_size);
        if iv.len() < key_data.block_size {
            iv.resize(key_data.block_size, 0);
        }

        let mut decrypted = aes_cbc_decrypt(secret_key, &iv, segment)?;
        out.append(&mut decrypted);
    }

    out.truncate(total_size);
    let _ = key_data.hash_size; // silence unused-field lint until HMAC verify is wired
    Ok(out)
}

/// Caller-tunable parameters for [`encrypt`].
///
/// `Default` produces Office 2010+ defaults: AES-256, SHA-512, 100 000
/// spinCount. Pin explicit values for reproducible output.
#[derive(Debug, Clone)]
pub struct AgileWriteOptions {
    /// AES key size in bits. Must be 128, 192, or 256.
    pub key_bits: u32,
    /// Spin count for the password KDF. Office's default is 100 000;
    /// MS-OFFCRYPTO §2.3.4.10 caps at 10 000 000.
    pub spin_count: u32,
    /// Hash algorithm used everywhere in the descriptor. SHA-512 is the
    /// modern default; SHA-256/384 also work; SHA-1 only for legacy
    /// reader compatibility (and is rejected by some readers).
    pub hash: HashAlgo,
}

impl Default for AgileWriteOptions {
    fn default() -> Self {
        Self {
            key_bits: 256,
            spin_count: 100_000,
            hash: HashAlgo::Sha512,
        }
    }
}

/// Encrypted OOXML envelope parts ready for CFB packaging.
///
/// `encryption_info` and `encrypted_package` are the two top-level
/// streams (`/EncryptionInfo` and `/EncryptedPackage`).
/// `dataspaces` carries the four `\x06DataSpaces`-tree streams Office
/// requires for type detection (MS-OFFCRYPTO §2.4); each entry is a
/// `(path, bytes)` pair where the path uses leading-slash form and
/// includes the `\x06` prefix bytes on the storages and `Primary`
/// stream. The matching storages (containers) the caller must also
/// create are listed in [`AGILE_REQUIRED_STORAGES`].
///
/// The caller is responsible for assembling these parts into a CFB v3
/// compound file. `duke-sheets-xls`'s `CompoundFileBuilder` is a
/// suitable backend.
#[derive(Debug)]
pub struct AgileEnvelopeParts {
    pub encryption_info: Vec<u8>,
    pub encrypted_package: Vec<u8>,
    pub dataspaces: Vec<(String, Vec<u8>)>,
}

/// Storage paths the caller must create (in this order) before adding
/// the streams from [`AgileEnvelopeParts::dataspaces`]. Without these
/// storages, Office rejects the file with "type detection failed".
pub const AGILE_REQUIRED_STORAGES: &[&str] = &[
    "/\u{0006}DataSpaces",
    "/\u{0006}DataSpaces/DataSpaceInfo",
    "/\u{0006}DataSpaces/TransformInfo",
    "/\u{0006}DataSpaces/TransformInfo/StrongEncryptionTransform",
];

/// Encrypt an OOXML inner-ZIP byte buffer with Agile encryption,
/// producing the parts that make up a complete CFB envelope.
///
/// The returned [`AgileEnvelopeParts`]:
/// 1. Round-trips through [`decrypt`] (or [`super::decrypt`]) with the
///    same password once the parts are reassembled.
/// 2. Carries an HMAC over the full encrypted package for integrity
///    verification by external readers.
///
/// Random salts, IVs, and keys are drawn from the OS RNG via
/// `getrandom`; no caller-provided seed.
///
/// CFB packaging is the caller's responsibility; the crypto crate stays
/// out of file-format territory so it can run on no-std-ish targets
/// like WASM without pulling in CFB-write code that isn't otherwise
/// reachable from a decrypt-only build.
pub fn encrypt(
    plaintext: &[u8],
    password: &str,
    opts: &AgileWriteOptions,
) -> CryptoResult<AgileEnvelopeParts> {
    if !matches!(opts.key_bits, 128 | 192 | 256) {
        return Err(CryptoError::UnsupportedVariant(format!(
            "Agile keyBits={} (only 128/192/256 supported)",
            opts.key_bits
        )));
    }
    if opts.spin_count > MAX_SPIN_COUNT {
        return Err(CryptoError::UnsupportedVariant(format!(
            "Agile spinCount={} exceeds MS-OFFCRYPTO limit of {MAX_SPIN_COUNT}",
            opts.spin_count
        )));
    }

    let key_bytes = (opts.key_bits / 8) as usize;
    let block_size = 16usize;
    let hash_size = opts.hash.digest_size();

    let key_data_salt = random_bytes(block_size)?;
    let key_encryptor_salt = random_bytes(block_size)?;
    let secret_key = random_bytes(key_bytes)?;
    let hmac_key = random_bytes(hash_size)?;

    let encrypted_package = encrypt_package(plaintext, &secret_key, &key_data_salt, opts.hash)?;

    let hmac_value = opts.hash.hmac(&hmac_key, &encrypted_package);

    let iv_hmac_key = derive_iv_from_salt(
        &key_data_salt,
        &BLK_DATA_INTEGRITY_KEY,
        opts.hash,
        block_size,
    );
    let iv_hmac_value = derive_iv_from_salt(
        &key_data_salt,
        &BLK_DATA_INTEGRITY_VALUE,
        opts.hash,
        block_size,
    );
    let encrypted_hmac_key = aes_cbc_encrypt(
        &secret_key,
        &iv_hmac_key,
        &pad_to_block(&hmac_key, block_size),
    )?;
    let encrypted_hmac_value = aes_cbc_encrypt(
        &secret_key,
        &iv_hmac_value,
        &pad_to_block(&hmac_value, block_size),
    )?;

    let key_vhi = derive_intermediate_key(
        password,
        &key_encryptor_salt,
        opts.spin_count,
        &BLK_VERIFIER_HASH_INPUT,
        opts.hash,
        key_bytes,
    );
    let key_vhv = derive_intermediate_key(
        password,
        &key_encryptor_salt,
        opts.spin_count,
        &BLK_VERIFIER_HASH_VALUE,
        opts.hash,
        key_bytes,
    );
    let key_ekv = derive_intermediate_key(
        password,
        &key_encryptor_salt,
        opts.spin_count,
        &BLK_ENCRYPTED_KEY_VALUE,
        opts.hash,
        key_bytes,
    );

    let verifier_input = random_bytes(block_size)?;
    let verifier_hash = opts.hash.hash(&verifier_input);

    let encrypted_verifier_hash_input =
        aes_cbc_encrypt(&key_vhi, &key_encryptor_salt, &verifier_input)?;
    let encrypted_verifier_hash_value = aes_cbc_encrypt(
        &key_vhv,
        &key_encryptor_salt,
        &pad_to_block(&verifier_hash, block_size),
    )?;
    let encrypted_key_value = aes_cbc_encrypt(
        &key_ekv,
        &key_encryptor_salt,
        &pad_to_block(&secret_key, block_size),
    )?;

    let xml = build_descriptor_xml(&DescriptorParams {
        key_bits: opts.key_bits,
        block_size: block_size as u32,
        hash_size: hash_size as u32,
        hash: opts.hash,
        spin_count: opts.spin_count,
        key_data_salt: &key_data_salt,
        key_encryptor_salt: &key_encryptor_salt,
        encrypted_hmac_key: &encrypted_hmac_key,
        encrypted_hmac_value: &encrypted_hmac_value,
        encrypted_verifier_hash_input: &encrypted_verifier_hash_input,
        encrypted_verifier_hash_value: &encrypted_verifier_hash_value,
        encrypted_key_value: &encrypted_key_value,
    });

    let encryption_info = build_encryption_info_stream(&xml);
    let dataspaces_root = "/\u{0006}DataSpaces";
    let dataspaces = vec![
        (
            format!("{dataspaces_root}/Version"),
            build_dataspaces_version(),
        ),
        (
            format!("{dataspaces_root}/DataSpaceMap"),
            build_dataspace_map(),
        ),
        (
            format!("{dataspaces_root}/DataSpaceInfo/StrongEncryptionDataSpace"),
            build_strong_encryption_dataspace(),
        ),
        (
            format!("{dataspaces_root}/TransformInfo/StrongEncryptionTransform/\u{0006}Primary"),
            build_primary_transform(),
        ),
    ];

    Ok(AgileEnvelopeParts {
        encryption_info,
        encrypted_package,
        dataspaces,
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

fn derive_iv_from_salt(
    salt: &[u8],
    block_key: &[u8; 8],
    hash: HashAlgo,
    block_size: usize,
) -> Vec<u8> {
    let mut buf = Vec::with_capacity(salt.len() + block_key.len());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(block_key);
    let mut iv = hash.hash(&buf);
    if iv.len() < block_size {
        iv.resize(block_size, 0);
    } else {
        iv.truncate(block_size);
    }
    iv
}

fn aes_cbc_encrypt(key: &[u8], iv: &[u8], plaintext: &[u8]) -> CryptoResult<Vec<u8>> {
    if plaintext.len() % 16 != 0 {
        return Err(CryptoError::InvalidFormat(format!(
            "AES-CBC plaintext length {} not a multiple of 16",
            plaintext.len()
        )));
    }
    let mut iv_buf = [0u8; 16];
    iv_buf.copy_from_slice(&iv[..iv.len().min(16)]);
    let mut buf = vec![0u8; plaintext.len()];
    buf[..plaintext.len()].copy_from_slice(plaintext);

    match key.len() {
        16 => {
            type Aes128CbcEnc = cbc::Encryptor<Aes128>;
            let cipher = Aes128CbcEnc::new_from_slices(key, &iv_buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-128-CBC init: {e}")))?;
            cipher
                .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-128-CBC encrypt: {e}")))?;
        }
        24 => {
            type Aes192CbcEnc = cbc::Encryptor<Aes192>;
            let cipher = Aes192CbcEnc::new_from_slices(key, &iv_buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-192-CBC init: {e}")))?;
            cipher
                .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-192-CBC encrypt: {e}")))?;
        }
        32 => {
            type Aes256CbcEnc = cbc::Encryptor<Aes256>;
            let cipher = Aes256CbcEnc::new_from_slices(key, &iv_buf)
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-256-CBC init: {e}")))?;
            cipher
                .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .map_err(|e| CryptoError::InvalidFormat(format!("AES-256-CBC encrypt: {e}")))?;
        }
        n => {
            return Err(CryptoError::UnsupportedVariant(format!(
                "Agile AES key length {n} not in {{16, 24, 32}}"
            )));
        }
    }
    Ok(buf)
}

fn encrypt_package(
    plaintext: &[u8],
    secret_key: &[u8],
    key_data_salt: &[u8],
    hash: HashAlgo,
) -> CryptoResult<Vec<u8>> {
    const SEGMENT: usize = 4096;
    let block_size = 16usize;

    let total_size = plaintext.len();
    let padded = pad_to_block(plaintext, block_size);

    let mut out = Vec::with_capacity(8 + padded.len());
    out.extend_from_slice(&(total_size as u64).to_le_bytes());

    for (i, segment) in padded.chunks(SEGMENT).enumerate() {
        let mut salt_with_block_key = Vec::with_capacity(key_data_salt.len() + 4);
        salt_with_block_key.extend_from_slice(key_data_salt);
        salt_with_block_key.extend_from_slice(&(i as u32).to_le_bytes());
        let mut iv = hash.hash(&salt_with_block_key);
        if iv.len() < block_size {
            iv.resize(block_size, 0);
        } else {
            iv.truncate(block_size);
        }
        let enc = aes_cbc_encrypt(secret_key, &iv, segment)?;
        out.extend_from_slice(&enc);
    }
    Ok(out)
}

struct DescriptorParams<'a> {
    key_bits: u32,
    block_size: u32,
    hash_size: u32,
    hash: HashAlgo,
    spin_count: u32,
    key_data_salt: &'a [u8],
    key_encryptor_salt: &'a [u8],
    encrypted_hmac_key: &'a [u8],
    encrypted_hmac_value: &'a [u8],
    encrypted_verifier_hash_input: &'a [u8],
    encrypted_verifier_hash_value: &'a [u8],
    encrypted_key_value: &'a [u8],
}

fn build_descriptor_xml(p: &DescriptorParams<'_>) -> Vec<u8> {
    let salt_size = 16u32;
    let hash_name = p.hash.name();
    let kds_b64 = encode_base64(p.key_data_salt);
    let kes_b64 = encode_base64(p.key_encryptor_salt);
    let ehk_b64 = encode_base64(p.encrypted_hmac_key);
    let ehv_b64 = encode_base64(p.encrypted_hmac_value);
    let evhi_b64 = encode_base64(p.encrypted_verifier_hash_input);
    let evhv_b64 = encode_base64(p.encrypted_verifier_hash_value);
    let ekv_b64 = encode_base64(p.encrypted_key_value);

    let mut s = String::new();
    s.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    s.push_str("<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\" xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">");
    s.push_str(&format!(
        "<keyData saltSize=\"{salt_size}\" blockSize=\"{block_size}\" keyBits=\"{key_bits}\" hashSize=\"{hash_size}\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"{hash_name}\" saltValue=\"{kds_b64}\"/>",
        block_size = p.block_size,
        key_bits = p.key_bits,
        hash_size = p.hash_size,
    ));
    s.push_str(&format!(
        "<dataIntegrity encryptedHmacKey=\"{ehk_b64}\" encryptedHmacValue=\"{ehv_b64}\"/>"
    ));
    s.push_str("<keyEncryptors>");
    s.push_str(
        "<keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">",
    );
    s.push_str(&format!(
        "<p:encryptedKey spinCount=\"{spin_count}\" saltSize=\"{salt_size}\" blockSize=\"{block_size}\" keyBits=\"{key_bits}\" hashSize=\"{hash_size}\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"{hash_name}\" saltValue=\"{kes_b64}\" encryptedVerifierHashInput=\"{evhi_b64}\" encryptedVerifierHashValue=\"{evhv_b64}\" encryptedKeyValue=\"{ekv_b64}\"/>",
        spin_count = p.spin_count,
        block_size = p.block_size,
        key_bits = p.key_bits,
        hash_size = p.hash_size,
    ));
    s.push_str("</keyEncryptor>");
    s.push_str("</keyEncryptors>");
    s.push_str("</encryption>");
    s.into_bytes()
}

/// Build the full `/EncryptionInfo` stream: 8-byte version/flags header
/// (vMajor=4 vMinor=4 flags=0x40 = fAgile) followed by the XML
/// descriptor body.
fn build_encryption_info_stream(xml: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(8 + xml.len());
    out.extend_from_slice(&4u16.to_le_bytes());
    out.extend_from_slice(&4u16.to_le_bytes());
    out.extend_from_slice(&0x0000_0040u32.to_le_bytes());
    out.extend_from_slice(xml);
    out
}

/// Encode `s` as length-prefixed UTF-16LE: `u32 byte_length || UTF-16LE
/// bytes || zero pad to next 4-byte boundary`. Used throughout the
/// DataSpaces structures (MS-OFFCRYPTO §2.5.1.1).
fn encode_prefixed_utf16le(s: &str) -> Vec<u8> {
    let utf16: Vec<u16> = s.encode_utf16().collect();
    let byte_len = utf16.len() * 2;
    let padded = byte_len.div_ceil(4) * 4;
    let mut out = Vec::with_capacity(4 + padded);
    out.extend_from_slice(&(byte_len as u32).to_le_bytes());
    for u in &utf16 {
        out.extend_from_slice(&u.to_le_bytes());
    }
    while out.len() < 4 + padded {
        out.push(0);
    }
    out
}

/// Build the `/\x06DataSpaces/Version` stream (MS-OFFCRYPTO §2.5.4).
/// Identifies the FeatureIdentifier and the reader/updater/writer
/// version triplet (all 1.0 for current Office).
fn build_dataspaces_version() -> Vec<u8> {
    let mut out = encode_prefixed_utf16le("Microsoft.Container.DataSpaces");
    for v in [1u16, 0u16, 1, 0, 1, 0] {
        out.extend_from_slice(&v.to_le_bytes());
    }
    out
}

/// Build the `/\x06DataSpaces/DataSpaceMap` stream (MS-OFFCRYPTO §2.5.5).
///
/// Maps `EncryptedPackage` (a stream) to the `StrongEncryptionDataSpace`
/// — the link Office follows when opening the file.
fn build_dataspace_map() -> Vec<u8> {
    let component_name = encode_prefixed_utf16le("EncryptedPackage");
    let dataspace_name = encode_prefixed_utf16le("StrongEncryptionDataSpace");
    let entry_inner = {
        let mut v = Vec::new();
        v.extend_from_slice(&1u32.to_le_bytes()); // ReferenceComponentCount
        v.extend_from_slice(&0u32.to_le_bytes()); // ReferenceComponentType (stream)
        v.extend_from_slice(&component_name);
        v.extend_from_slice(&dataspace_name);
        v
    };
    let entry_total = 4 + entry_inner.len();
    let mut out = Vec::with_capacity(8 + entry_total);
    out.extend_from_slice(&8u32.to_le_bytes()); // HeaderLength
    out.extend_from_slice(&1u32.to_le_bytes()); // EntryCount
    out.extend_from_slice(&(entry_total as u32).to_le_bytes());
    out.extend_from_slice(&entry_inner);
    out
}

/// Build the `/\x06DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace`
/// stream (MS-OFFCRYPTO §2.5.6). Lists the transforms applied to the
/// EncryptedPackage; for password-encrypted files, just one
/// `StrongEncryptionTransform`.
fn build_strong_encryption_dataspace() -> Vec<u8> {
    let transform_name = encode_prefixed_utf16le("StrongEncryptionTransform");
    let mut out = Vec::with_capacity(8 + transform_name.len());
    out.extend_from_slice(&8u32.to_le_bytes()); // HeaderLength
    out.extend_from_slice(&1u32.to_le_bytes()); // EntryCount
    out.extend_from_slice(&transform_name);
    out
}

/// Build the
/// `/\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary`
/// stream (MS-OFFCRYPTO §2.5.7). Carries the well-known Office
/// encryption-transform GUID and an empty EncryptionTransformInfo (the
/// real key data lives in `/EncryptionInfo`).
fn build_primary_transform() -> Vec<u8> {
    const TRANSFORM_ID: &str = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}";
    const TRANSFORM_NAME: &str = "Microsoft.Container.EncryptionTransform";

    let id = encode_prefixed_utf16le(TRANSFORM_ID);
    let name = encode_prefixed_utf16le(TRANSFORM_NAME);

    let transform_length = 4 + 4 + id.len();

    let mut out = Vec::with_capacity(transform_length + name.len() + 12 + 16);
    out.extend_from_slice(&(transform_length as u32).to_le_bytes());
    out.extend_from_slice(&1u32.to_le_bytes()); // TransformType: Protection
    out.extend_from_slice(&id);
    out.extend_from_slice(&name);
    for v in [1u16, 0u16, 1, 0, 1, 0] {
        out.extend_from_slice(&v.to_le_bytes());
    }
    out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionName length = 0
    out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionBlockSize = 0
    out.extend_from_slice(&0u32.to_le_bytes()); // CipherMode = 0
    out.extend_from_slice(&4u32.to_le_bytes()); // Reserved (MUST be 4)
    out
}

/// Standard base64 encode (RFC 4648, `+/` alphabet, `=` padding). Mirror
/// of [`decode_base64`] above; rolled by hand to avoid a dependency on
/// the `base64` crate for a few thousand bytes per file.
fn encode_base64(bytes: &[u8]) -> String {
    const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity(((bytes.len() + 2) / 3) * 4);
    let mut i = 0;
    while i + 3 <= bytes.len() {
        let n =
            (u32::from(bytes[i]) << 16) | (u32::from(bytes[i + 1]) << 8) | u32::from(bytes[i + 2]);
        out.push(ALPHABET[((n >> 18) & 0x3F) as usize] as char);
        out.push(ALPHABET[((n >> 12) & 0x3F) as usize] as char);
        out.push(ALPHABET[((n >> 6) & 0x3F) as usize] as char);
        out.push(ALPHABET[(n & 0x3F) as usize] as char);
        i += 3;
    }
    let rem = bytes.len() - i;
    if rem == 1 {
        let n = u32::from(bytes[i]) << 16;
        out.push(ALPHABET[((n >> 18) & 0x3F) as usize] as char);
        out.push(ALPHABET[((n >> 12) & 0x3F) as usize] as char);
        out.push('=');
        out.push('=');
    } else if rem == 2 {
        let n = (u32::from(bytes[i]) << 16) | (u32::from(bytes[i + 1]) << 8);
        out.push(ALPHABET[((n >> 18) & 0x3F) as usize] as char);
        out.push(ALPHABET[((n >> 12) & 0x3F) as usize] as char);
        out.push(ALPHABET[((n >> 6) & 0x3F) as usize] as char);
        out.push('=');
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn block_keys_match_spec() {
        // Spec values from MS-OFFCRYPTO §2.3.4.13.
        assert_eq!(
            BLK_VERIFIER_HASH_INPUT,
            [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79]
        );
        assert_eq!(
            BLK_VERIFIER_HASH_VALUE,
            [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E]
        );
        assert_eq!(
            BLK_ENCRYPTED_KEY_VALUE,
            [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6]
        );
        assert_eq!(
            BLK_DATA_INTEGRITY_KEY,
            [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6]
        );
        assert_eq!(
            BLK_DATA_INTEGRITY_VALUE,
            [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33]
        );
    }

    #[test]
    fn hashalgo_parses_known_values() {
        assert_eq!(HashAlgo::parse("SHA1").unwrap(), HashAlgo::Sha1);
        assert_eq!(HashAlgo::parse("SHA256").unwrap(), HashAlgo::Sha256);
        assert_eq!(HashAlgo::parse("SHA384").unwrap(), HashAlgo::Sha384);
        assert_eq!(HashAlgo::parse("SHA512").unwrap(), HashAlgo::Sha512);
        assert!(HashAlgo::parse("MD5").is_err());
    }

    #[test]
    fn base64_round_trips_known_strings() {
        assert_eq!(decode_base64("").unwrap(), Vec::<u8>::new());
        assert_eq!(decode_base64("aGVsbG8=").unwrap(), b"hello");
        assert_eq!(decode_base64("YWJjZGVmZw==").unwrap(), b"abcdefg");
        assert_eq!(decode_base64("aGVsbG8gd29ybGQ=").unwrap(), b"hello world");
    }

    #[test]
    fn base64_rejects_non_base64_chars() {
        assert!(decode_base64("not~base64").is_err());
    }

    #[test]
    fn parse_descriptor_extracts_known_fields() {
        let xml = br##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
           hashAlgorithm="SHA512" saltValue="aGVsbG8gc2FsdAAAAAA="/>
  <dataIntegrity encryptedHmacKey="aGVsbG8=" encryptedHmacValue="d29ybGQ="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
        spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64"
        cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"
        saltValue="c2FsdF9rZXlfZW5jcnlwdG9y" encryptedVerifierHashInput="dmhpAAAAAAAAAAAAAAAAAA=="
        encryptedVerifierHashValue="dmh2AAAAAAAAAAAAAAAAAA=="
        encryptedKeyValue="ZWt2AAAAAAAAAAAAAAAAAA=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"##;
        let d = parse_descriptor(xml).expect("parse ok");
        assert_eq!(d.spin_count, 100_000);
        assert_eq!(d.key_data.key_bits, 256);
        assert_eq!(d.key_data.hash_algorithm, HashAlgo::Sha512);
        assert_eq!(d.key_data.salt_value, b"hello salt\x00\x00\x00\x00");
        assert_eq!(d.key_encryptor.salt_value, b"salt_key_encryptor");
        assert!(!d.encrypted_hmac_key.is_empty());
        assert!(!d.encrypted_hmac_value.is_empty());
    }

    /// AES-CBC round-trip for each key size we claim to support.
    /// Validates that the dispatch in `aes_cbc_decrypt` plumbs through
    /// the right cipher type for 128- and 192-bit keys (the path that
    /// the corpus surfaced an unsupported error on).
    #[test]
    fn aes_cbc_decrypt_supports_128_192_256() {
        use cbc::cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};
        let plaintext = [0x77u8; 64];
        let iv = [0x13u8; 16];

        for key_len in [16usize, 24, 32] {
            let key = vec![0x42u8; key_len];
            let mut buf = plaintext.to_vec();
            match key_len {
                16 => {
                    type Aes128CbcEnc = cbc::Encryptor<Aes128>;
                    let cipher = Aes128CbcEnc::new_from_slices(&key, &iv).unwrap();
                    cipher
                        .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                        .unwrap();
                }
                24 => {
                    type Aes192CbcEnc = cbc::Encryptor<Aes192>;
                    let cipher = Aes192CbcEnc::new_from_slices(&key, &iv).unwrap();
                    cipher
                        .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                        .unwrap();
                }
                32 => {
                    type Aes256CbcEnc = cbc::Encryptor<Aes256>;
                    let cipher = Aes256CbcEnc::new_from_slices(&key, &iv).unwrap();
                    cipher
                        .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                        .unwrap();
                }
                _ => unreachable!(),
            }
            let decrypted = aes_cbc_decrypt(&key, &iv, &buf).expect("decrypt ok");
            assert_eq!(
                decrypted,
                plaintext,
                "round-trip for {}-bit key",
                key_len * 8
            );
        }
    }

    #[test]
    fn aes_cbc_decrypt_rejects_unsupported_key_lengths() {
        let iv = [0u8; 16];
        let ct = [0u8; 16];
        let err = aes_cbc_decrypt(&[0u8; 17], &iv, &ct).unwrap_err();
        assert!(matches!(err, CryptoError::UnsupportedVariant(_)));
    }

    #[test]
    fn parse_descriptor_rejects_excessive_spincount() {
        let xml = br##"<?xml version="1.0"?>
<encryption xmlns="x">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
           hashAlgorithm="SHA512" saltValue=""/>
  <keyEncryptors>
    <keyEncryptor>
      <p:encryptedKey xmlns:p="x" spinCount="999999999"
        saltSize="16" blockSize="16" keyBits="256" hashSize="64"
        cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"
        saltValue="" encryptedVerifierHashInput="" encryptedVerifierHashValue=""
        encryptedKeyValue=""/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"##;
        let err = parse_descriptor(xml).expect_err("must reject huge spinCount");
        assert!(matches!(err, CryptoError::InvalidFormat(_)));
    }
}
