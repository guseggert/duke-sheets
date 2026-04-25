//! OOXML (ECMA-376) encryption variants: Agile, Standard, and Binary RC4.
//!
//! Encrypted OOXML files are CFB containers (not ZIPs) with two streams:
//!
//! - `/EncryptionInfo` — header describing the encryption scheme
//! - `/EncryptedPackage` — ciphertext of the inner ZIP archive
//!
//! The first 4 bytes of EncryptionInfo identify the variant via
//! `vMajor`/`vMinor` (MS-OFFCRYPTO §2.3.4):
//!
//! | vMajor.vMinor | Variant |
//! |---|---|
//! | 4.4 | Agile (XML descriptor, AES-CBC + HMAC) |
//! | 3.2 / 4.2 | Standard (binary header, AES-ECB) |
//! | 2.2 / 3.2 / 4.2 | Binary RC4 CryptoAPI (when fAES flag is unset) |

pub mod agile;
pub mod standard;

use crate::error::{CryptoError, CryptoResult};

/// Detected OOXML encryption variant from the `/EncryptionInfo` stream
/// header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OoxmlVariant {
    Agile,
    Standard,
    BinaryRc4,
}

/// Inspect the first bytes of an `/EncryptionInfo` stream and identify
/// which encryption variant the file uses. Returns
/// [`CryptoError::InvalidFormat`] if the header doesn't match any
/// known OOXML variant.
pub fn detect_variant(encryption_info: &[u8]) -> CryptoResult<OoxmlVariant> {
    if encryption_info.len() < 8 {
        return Err(CryptoError::InvalidFormat(
            "EncryptionInfo too short for header".into(),
        ));
    }
    let v_major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let v_minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);
    let flags = u32::from_le_bytes([
        encryption_info[4],
        encryption_info[5],
        encryption_info[6],
        encryption_info[7],
    ]);

    match (v_major, v_minor) {
        (4, 4) => Ok(OoxmlVariant::Agile),
        (2 | 3 | 4, 2) => {
            // MS-OFFCRYPTO §2.3.4.4 EncryptionHeaderFlags:
            //   bit 2 (0x04) = fCryptoAPI - must be set in any 2007+ variant
            //   bit 5 (0x20) = fAES       - set => Standard, clear => Binary RC4
            const F_CRYPTOAPI: u32 = 0x04;
            const F_AES: u32 = 0x20;
            if flags & F_CRYPTOAPI == 0 {
                return Err(CryptoError::InvalidFormat(format!(
                    "EncryptionInfo flags {flags:#010x} lack fCryptoAPI bit"
                )));
            }
            if flags & F_AES != 0 {
                Ok(OoxmlVariant::Standard)
            } else {
                Ok(OoxmlVariant::BinaryRc4)
            }
        }
        _ => Err(CryptoError::InvalidFormat(format!(
            "unknown EncryptionInfo version vMajor={v_major} vMinor={v_minor}"
        ))),
    }
}

/// Decrypt an `/EncryptedPackage` stream given a password and the raw
/// `/EncryptionInfo` stream bytes. Dispatches to the variant
/// implementation in submodules.
pub fn decrypt(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> CryptoResult<Vec<u8>> {
    match detect_variant(encryption_info)? {
        OoxmlVariant::Agile => agile::decrypt(encryption_info, encrypted_package, password),
        OoxmlVariant::Standard => standard::decrypt(encryption_info, encrypted_package, password),
        OoxmlVariant::BinaryRc4 => Err(CryptoError::UnsupportedVariant(
            "OOXML Binary RC4 CryptoAPI (planned for a later phase)".into(),
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn header(v_major: u16, v_minor: u16, flags: u32) -> [u8; 8] {
        let mut h = [0u8; 8];
        h[0..2].copy_from_slice(&v_major.to_le_bytes());
        h[2..4].copy_from_slice(&v_minor.to_le_bytes());
        h[4..8].copy_from_slice(&flags.to_le_bytes());
        h
    }

    #[test]
    fn detects_agile() {
        let h = header(4, 4, 0x00000040);
        assert_eq!(detect_variant(&h).unwrap(), OoxmlVariant::Agile);
    }

    #[test]
    fn detects_standard() {
        // fCryptoAPI (bit 2 = 0x04) | fAES (bit 5 = 0x20) = 0x24
        let h = header(4, 2, 0x00000024);
        assert_eq!(detect_variant(&h).unwrap(), OoxmlVariant::Standard);
    }

    #[test]
    fn detects_binary_rc4() {
        let h = header(3, 2, 0x00000004); // fCryptoAPI only
        assert_eq!(detect_variant(&h).unwrap(), OoxmlVariant::BinaryRc4);
    }

    #[test]
    fn rejects_unknown_version() {
        let h = header(7, 99, 0);
        assert!(matches!(
            detect_variant(&h),
            Err(CryptoError::InvalidFormat(_))
        ));
    }
}
