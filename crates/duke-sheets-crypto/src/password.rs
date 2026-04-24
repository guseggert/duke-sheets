//! Password encoding helpers shared by every Office crypto variant.
//!
//! Every KDF in MS-OFFCRYPTO takes the password as UTF-16LE without BOM.
//! Getting this wrong silently produces the wrong derived key and every
//! decrypt fails verifier check. Keep this boring and centralized.

/// Encode a UTF-8 password as UTF-16LE (no BOM). Empty strings round-trip
/// to an empty byte buffer.
pub(crate) fn utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for codepoint in password.encode_utf16() {
        out.extend_from_slice(&codepoint.to_le_bytes());
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn utf16le_empty_password_is_empty() {
        assert_eq!(utf16le_bytes(""), Vec::<u8>::new());
    }

    #[test]
    fn utf16le_ascii_password_is_padded() {
        // MS-OFFCRYPTO KDFs treat "a" as two bytes: 0x61 0x00, not one byte.
        // This is the #1 password implementation bug.
        assert_eq!(utf16le_bytes("a"), vec![0x61, 0x00]);
        assert_eq!(
            utf16le_bytes("hello"),
            vec![0x68, 0x00, 0x65, 0x00, 0x6C, 0x00, 0x6C, 0x00, 0x6F, 0x00]
        );
    }

    #[test]
    fn utf16le_multibyte_bmp_codepoint() {
        // Latin Small Letter A With Acute, U+00E1 → 0xE1 0x00.
        assert_eq!(utf16le_bytes("\u{00E1}"), vec![0xE1, 0x00]);
    }

    #[test]
    fn utf16le_astral_is_surrogate_pair() {
        // U+1F600 (grinning face) is outside the BMP; UTF-16 encodes it
        // as a surrogate pair. Office actually rejects non-BMP passwords
        // (old Windows API constraint) but we still want to produce
        // well-formed UTF-16LE so the caller can detect the problem
        // cleanly via a failing verifier, not an encoding panic.
        let bytes = utf16le_bytes("\u{1F600}");
        assert_eq!(bytes.len(), 4);
        // D83D DE00 in little-endian byte order.
        assert_eq!(bytes, vec![0x3D, 0xD8, 0x00, 0xDE]);
    }
}
