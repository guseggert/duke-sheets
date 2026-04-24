//! XLS (BIFF8) FilePass-based encryption: RC4 CryptoAPI, legacy RC4 (MD5),
//! and XOR obfuscation.
//!
//! See `docs/PASSWORD_SUPPORT.md` and MS-XLS §2.2.10 for how FilePass
//! interacts with the Workbook stream.

pub mod rc4_cryptoapi;
pub mod rc4_legacy;
pub(crate) mod record_walk;

use crate::error::{CryptoError, CryptoResult};

/// The BIFF8 `FilePass` record type.
pub const FILEPASS_RECORD_TYPE: u16 = 0x002F;

/// Decrypt an XLS Workbook stream in place.
///
/// Finds the `FilePass` record (which must appear directly after the
/// globals BOF), dispatches to the appropriate variant decryptor, and
/// returns a new buffer with record bodies decrypted. The output is the
/// same length as the input; the `FilePass` record has its body zeroed
/// so that subsequent passes treat the stream as unencrypted but all
/// absolute offsets (notably `BoundSheet8.lbPlyPos`) remain valid.
///
/// Returns [`CryptoError::InvalidFormat`] if the stream is not an
/// encrypted Workbook (no `FilePass` at the expected position). Returns
/// [`CryptoError::UnsupportedVariant`] for variants we don't yet
/// implement (legacy RC4 MD5, XOR).
pub fn decrypt_workbook_stream(stream: &[u8], password: &str) -> CryptoResult<Vec<u8>> {
    let filepass = find_filepass(stream)?;
    match parse_filepass(filepass.body)? {
        FilePassVariant::Rc4CryptoApi(params) => {
            rc4_cryptoapi::decrypt_workbook_stream(stream, password, &params)
        }
        FilePassVariant::Rc4Legacy(params) => {
            rc4_legacy::decrypt_workbook_stream(stream, password, &params)
        }
        FilePassVariant::Xor => Err(CryptoError::UnsupportedVariant(
            "XLS XOR Obfuscation".into(),
        )),
    }
}

/// The FilePass record's position and body within the Workbook stream.
#[derive(Debug)]
struct FilePass<'a> {
    #[allow(dead_code)]
    /// Offset of the FilePass record header in `stream`.
    header_offset: usize,
    /// FilePass body bytes (excluding the 4-byte record header).
    body: &'a [u8],
}

fn find_filepass(stream: &[u8]) -> CryptoResult<FilePass<'_>> {
    let mut cursor = 0usize;
    while cursor + 4 <= stream.len() {
        let record_type = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let size = u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]) as usize;
        let body_start = cursor + 4;
        let body_end = body_start + size;
        if body_end > stream.len() {
            return Err(CryptoError::InvalidFormat(format!(
                "record at {cursor:#x} extends past stream end"
            )));
        }
        if record_type == FILEPASS_RECORD_TYPE {
            return Ok(FilePass {
                header_offset: cursor,
                body: &stream[body_start..body_end],
            });
        }
        // Per MS-XLS §2.1.6, FilePass (if present) must appear directly
        // after the globals BOF. Stop scanning once we've passed the
        // first BOF's body; anything later isn't the FilePass we want.
        if record_type == 0x0809 && cursor > 0 {
            break;
        }
        cursor = body_end;
    }
    Err(CryptoError::InvalidFormat(
        "no FilePass record found in Workbook stream".into(),
    ))
}

/// Parsed FilePass variant.
enum FilePassVariant {
    Rc4CryptoApi(rc4_cryptoapi::Rc4CryptoApiParams),
    Rc4Legacy(rc4_legacy::Rc4LegacyParams),
    Xor,
}

fn parse_filepass(body: &[u8]) -> CryptoResult<FilePassVariant> {
    if body.len() < 2 {
        return Err(CryptoError::InvalidFormat(
            "FilePass body too short for wEncryptionType".into(),
        ));
    }
    let w_encryption_type = u16::from_le_bytes([body[0], body[1]]);

    match w_encryption_type {
        0 => Ok(FilePassVariant::Xor),
        1 => {
            if body.len() < 6 {
                return Err(CryptoError::InvalidFormat(
                    "RC4 FilePass missing vMajor/vMinor".into(),
                ));
            }
            let v_major = u16::from_le_bytes([body[2], body[3]]);
            let v_minor = u16::from_le_bytes([body[4], body[5]]);
            match (v_major, v_minor) {
                (1, 1) => {
                    let params = rc4_legacy::parse_filepass_body(&body[6..])?;
                    Ok(FilePassVariant::Rc4Legacy(params))
                }
                (2, 2) | (3, 2) | (4, 2) => {
                    let params = parse_rc4_cryptoapi_header(&body[6..])?;
                    Ok(FilePassVariant::Rc4CryptoApi(params))
                }
                _ => Err(CryptoError::UnsupportedVariant(format!(
                    "RC4 FilePass vMajor={v_major:#x} vMinor={v_minor:#x}"
                ))),
            }
        }
        _ => Err(CryptoError::UnsupportedVariant(format!(
            "FilePass wEncryptionType={w_encryption_type:#x}"
        ))),
    }
}

/// Parse the MS-OFFCRYPTO EncryptionHeader + EncryptionVerifier that
/// follow the `wEncryptionType / vMajor / vMinor` triple of a FilePass
/// RC4 CryptoAPI record.
///
/// Layout (MS-OFFCRYPTO §2.3.2 + §2.3.3):
///   EncryptionHeader:
///     u32 Flags
///     u32 SizeExtra       // must be 0
///     u32 AlgID           // 0x6801 = RC4
///     u32 AlgIDHash       // 0x8004 = SHA-1
///     u32 KeySize         // in bits
///     u32 ProviderType    // 0x0001 = RC4
///     u32 Reserved1
///     u32 Reserved2
///     WCHAR[] CSPName     // UTF-16LE null-terminated
///   EncryptionVerifier:
///     u32 SaltSize        // must be 16
///     u8[16] Salt
///     u8[16] EncryptedVerifier
///     u32 VerifierHashSize // must be 20
///     u8[32] EncryptedVerifierHash
fn parse_rc4_cryptoapi_header(data: &[u8]) -> CryptoResult<rc4_cryptoapi::Rc4CryptoApiParams> {
    // EncryptionHeader starts with Flags (u32). MS-OFFCRYPTO actually
    // prefixes it with a `u32 Size` telling how many bytes the
    // EncryptionHeader spans (including its Size field). Parse that
    // first.
    if data.len() < 4 {
        return Err(CryptoError::InvalidFormat(
            "RC4 CryptoAPI header missing size prefix".into(),
        ));
    }
    let header_size = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
    if 4 + header_size > data.len() {
        return Err(CryptoError::InvalidFormat(format!(
            "RC4 CryptoAPI EncryptionHeader claims {header_size} bytes but only {} remain",
            data.len() - 4
        )));
    }
    // Parse fixed-size fields of the EncryptionHeader (32 bytes before
    // the variable-length CSPName).
    let hdr = &data[4..4 + header_size];
    if hdr.len() < 32 {
        return Err(CryptoError::InvalidFormat(format!(
            "RC4 CryptoAPI EncryptionHeader too short: {}",
            hdr.len()
        )));
    }
    let alg_id = u32::from_le_bytes([hdr[8], hdr[9], hdr[10], hdr[11]]);
    let key_size_bits = u32::from_le_bytes([hdr[16], hdr[17], hdr[18], hdr[19]]);
    let provider_type = u32::from_le_bytes([hdr[20], hdr[21], hdr[22], hdr[23]]);

    if alg_id != 0x0000_6801 && alg_id != 0 {
        // algId=0 means "default" per spec (RC4 CryptoAPI); 0x6801
        // means RC4 explicitly. Anything else is a different cipher.
        return Err(CryptoError::UnsupportedVariant(format!(
            "RC4 CryptoAPI header algId={alg_id:#010x} (expected 0x6801 or 0)"
        )));
    }
    if provider_type != 0x0000_0001 && provider_type != 0 {
        return Err(CryptoError::UnsupportedVariant(format!(
            "RC4 CryptoAPI header providerType={provider_type:#010x}"
        )));
    }

    // EncryptionVerifier follows the EncryptionHeader.
    let verifier_start = 4 + header_size;
    let verifier = &data[verifier_start..];
    if verifier.len() < 4 + 16 + 16 + 4 + 32 {
        return Err(CryptoError::InvalidFormat(format!(
            "RC4 CryptoAPI EncryptionVerifier too short: {}",
            verifier.len()
        )));
    }

    let salt_size = u32::from_le_bytes([verifier[0], verifier[1], verifier[2], verifier[3]]);
    if salt_size != 16 {
        return Err(CryptoError::InvalidFormat(format!(
            "RC4 CryptoAPI saltSize={salt_size} (expected 16)"
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
            "RC4 CryptoAPI verifierHashSize={verifier_hash_size} (expected 20)"
        )));
    }

    let mut encrypted_verifier_hash = [0u8; 32];
    encrypted_verifier_hash.copy_from_slice(&verifier[40..72]);

    Ok(rc4_cryptoapi::Rc4CryptoApiParams {
        salt,
        key_size_bits,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn find_filepass_not_found_when_absent() {
        let mut stream = Vec::new();
        // Just a BOF (globals) with no FilePass follow-up.
        stream.extend_from_slice(&0x0809u16.to_le_bytes());
        stream.extend_from_slice(&8u16.to_le_bytes());
        stream.extend_from_slice(&[0u8; 8]);
        stream.extend_from_slice(&0x000Au16.to_le_bytes()); // EOF
        stream.extend_from_slice(&0u16.to_le_bytes());

        let err = find_filepass(&stream).unwrap_err();
        assert!(matches!(err, CryptoError::InvalidFormat(_)));
    }

    #[test]
    fn find_filepass_locates_record_after_bof() {
        let mut stream = Vec::new();
        // BOF (globals)
        stream.extend_from_slice(&0x0809u16.to_le_bytes());
        stream.extend_from_slice(&8u16.to_le_bytes());
        stream.extend_from_slice(&[0u8; 8]);
        // FilePass (XOR obfuscation, wEncryptionType=0, no body)
        stream.extend_from_slice(&0x002Fu16.to_le_bytes());
        stream.extend_from_slice(&2u16.to_le_bytes());
        stream.extend_from_slice(&[0, 0]);

        let fp = find_filepass(&stream).unwrap();
        assert_eq!(fp.header_offset, 12);
        assert_eq!(fp.body, &[0, 0]);
    }

    #[test]
    fn parse_filepass_xor() {
        let body = [0, 0];
        assert!(matches!(parse_filepass(&body), Ok(FilePassVariant::Xor)));
    }

    #[test]
    fn parse_filepass_rc4_legacy_is_recognized() {
        // wEncryptionType=1, vMajor=1, vMinor=1 + 48 zero bytes (minimum viable body)
        let mut body = vec![1, 0, 1, 0, 1, 0];
        body.extend_from_slice(&[0u8; 48]);
        assert!(matches!(
            parse_filepass(&body),
            Ok(FilePassVariant::Rc4Legacy(_))
        ));
    }

    #[test]
    fn parse_filepass_unknown_encryption_type_is_unsupported() {
        let body = [99, 0];
        assert!(matches!(
            parse_filepass(&body),
            Err(CryptoError::UnsupportedVariant(_))
        ));
    }
}
