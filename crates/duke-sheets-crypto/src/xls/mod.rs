//! XLS (BIFF8) FilePass-based encryption: RC4 CryptoAPI, legacy RC4 (MD5),
//! and XOR obfuscation.
//!
//! See MS-XLS §2.2.10 for how FilePass interacts with the Workbook stream.

pub mod rc4_cryptoapi;
pub mod rc4_legacy;
pub(crate) mod record_walk;
pub mod xor_obfuscation;

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
/// implement.
pub fn decrypt_workbook_stream(stream: &[u8], password: &str) -> CryptoResult<Vec<u8>> {
    let filepass = find_filepass(stream)?;
    match parse_filepass(filepass.body)? {
        FilePassVariant::Rc4CryptoApi(params) => {
            rc4_cryptoapi::decrypt_workbook_stream(stream, password, &params)
        }
        FilePassVariant::Rc4Legacy(params) => {
            rc4_legacy::decrypt_workbook_stream(stream, password, &params)
        }
        FilePassVariant::Xor(params) => {
            xor_obfuscation::decrypt_workbook_stream(stream, password, &params)
        }
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
    Xor(xor_obfuscation::XorParams),
}

fn parse_filepass(body: &[u8]) -> CryptoResult<FilePassVariant> {
    if body.len() < 2 {
        return Err(CryptoError::InvalidFormat(
            "FilePass body too short for wEncryptionType".into(),
        ));
    }
    let w_encryption_type = u16::from_le_bytes([body[0], body[1]]);

    match w_encryption_type {
        0 => {
            let params = xor_obfuscation::parse_filepass_body(&body[2..])?;
            Ok(FilePassVariant::Xor(params))
        }
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
/// XLS FilePass body layout for RC4 CryptoAPI (vMinor=2), per
/// MS-OFFCRYPTO §2.3.4 + msoffcrypto-tool's `_parse_header_RC4CryptoAPI`:
///
/// ```text
///   u32 EncryptionHeaderFlags     // duplicates the per-header Flags below; skip
///   u32 EncryptionHeaderSize      // length of the EncryptionHeader that follows
///   EncryptionHeader (size bytes):
///     u32 Flags
///     u32 SizeExtra                // must be 0
///     u32 AlgID                    // 0x6801 = RC4 (or 0 = default)
///     u32 AlgIDHash                // 0x8004 = SHA-1
///     u32 KeySize                  // in bits (40 / 128 / etc.; 0 = default 40)
///     u32 ProviderType             // 0x0001 = RC4 (or 0 = default)
///     u32 Reserved1
///     u32 Reserved2
///     UTF-16LE CSPName             // null-terminated, variable length
///   EncryptionVerifier:
///     u32 SaltSize                 // = 16
///     u8[16] Salt
///     u8[16] EncryptedVerifier
///     u32 VerifierHashSize         // = 20 (SHA-1 digest size)
///     u8[20] EncryptedVerifierHash // 20 for RC4; AES variants pad to 32
/// ```
fn parse_rc4_cryptoapi_header(data: &[u8]) -> CryptoResult<rc4_cryptoapi::Rc4CryptoApiParams> {
    if data.len() < 8 {
        return Err(CryptoError::InvalidFormat(
            "RC4 CryptoAPI header missing flags+size prefix".into(),
        ));
    }
    let header_size = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;
    let header_start = 8usize;
    let header_end = header_start
        .checked_add(header_size)
        .ok_or_else(|| CryptoError::InvalidFormat("RC4 CryptoAPI header size overflow".into()))?;
    if header_end > data.len() {
        return Err(CryptoError::InvalidFormat(format!(
            "RC4 CryptoAPI EncryptionHeader claims {header_size} bytes but only {} remain",
            data.len() - header_start
        )));
    }
    let hdr = &data[header_start..header_end];
    if hdr.len() < 32 {
        return Err(CryptoError::InvalidFormat(format!(
            "RC4 CryptoAPI EncryptionHeader too short: {}",
            hdr.len()
        )));
    }
    let alg_id = u32::from_le_bytes([hdr[8], hdr[9], hdr[10], hdr[11]]);
    let raw_key_size = u32::from_le_bytes([hdr[16], hdr[17], hdr[18], hdr[19]]);
    let provider_type = u32::from_le_bytes([hdr[20], hdr[21], hdr[22], hdr[23]]);

    if alg_id != 0x0000_6801 && alg_id != 0 {
        return Err(CryptoError::UnsupportedVariant(format!(
            "RC4 CryptoAPI header algId={alg_id:#010x} (expected 0x6801 or 0)"
        )));
    }
    if provider_type != 0x0000_0001 && provider_type != 0 {
        return Err(CryptoError::UnsupportedVariant(format!(
            "RC4 CryptoAPI header providerType={provider_type:#010x}"
        )));
    }

    // MS-OFFCRYPTO §2.3.4.5: KeySize=0 means "format default". For
    // RC4 CryptoAPI the default is 40 bits.
    let key_size_bits = if raw_key_size == 0 { 40 } else { raw_key_size };

    let verifier = &data[header_end..];
    if verifier.len() < 4 + 16 + 16 + 4 + 20 {
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

    // RC4 CryptoAPI's encrypted_verifier_hash is 20 bytes on the wire
    // (one SHA-1 digest). Our params struct holds 32 bytes for AES
    // alignment compatibility; the unused tail stays zero.
    let mut encrypted_verifier_hash = [0u8; 32];
    encrypted_verifier_hash[..20].copy_from_slice(&verifier[40..60]);

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
        // wEncryptionType=0 + 4 bytes XOR params (key + verification_bytes)
        let body = [0u8, 0, 0xAB, 0xCD, 0x0A, 0x9A];
        assert!(matches!(parse_filepass(&body), Ok(FilePassVariant::Xor(_))));
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
