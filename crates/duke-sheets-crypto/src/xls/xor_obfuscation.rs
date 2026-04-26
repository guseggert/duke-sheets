//! XLS XOR Obfuscation (MS-OFFCRYPTO §2.3.6, MS-XLS §2.4.117 with
//! `wEncryptionType=0`).
//!
//! Excel 95-era encryption — weak by design and not actual crypto.
//! Three lookup tables drive a 16-byte rolling XOR mask. Per byte:
//! XOR with the current mask entry, then rotate-right by 5 bits. The
//! initial mask index for each encrypted run is `(stream_offset + run_len) % 16`,
//! with a `+4` adjustment for the encrypted tail of `BoundSheet8` to
//! account for the 4-byte unencrypted `lbPlyPos` prefix.
//!
//! Passwords are capped at 15 characters by the algorithm itself
//! (`INITIAL_CODE` has 15 entries).

use crate::error::{CryptoError, CryptoResult};
use crate::xls::record_walk;

const PAD_ARRAY: [u8; 15] = [
    0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00,
];

const INITIAL_CODE: [u16; 15] = [
    0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9,
    0x280C, 0xA96A, 0x4EC3,
];

/// 15 rows × 7 entries, row-major. The reference implementation walks
/// this flat array backwards from index 104; no explicit row indexing.
const XOR_MATRIX: [u16; 105] = [
    0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, // row  0
    0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF, // row  1
    0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, // row  2
    0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40, // row  3
    0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5, // row  4
    0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, // row  5
    0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9, // row  6
    0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, // row  7
    0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, // row  8
    0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, // row  9
    0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, // row 10
    0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C, // row 11
    0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD, // row 12
    0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC, // row 13
    0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4, // row 14
];

/// XOR Obfuscation magic (MS-OFFCRYPTO §2.3.7.4).
const VERIFIER_MAGIC: u16 = 0xCE4B;

/// Maximum password length supported by XOR Obfuscation.
pub const MAX_PASSWORD_LEN: usize = 15;

/// Parameters parsed from an XOR-Obfuscation FilePass record body.
#[derive(Debug, Clone, Copy)]
pub struct XorParams {
    /// 16-bit Key field as stored in FilePass. Spec calls this "Key";
    /// the runtime XorKey is regenerated from the password rather than
    /// read from this field.
    pub stored_key: u16,
    /// 16-bit verification value used by `verify_password`.
    pub verification_bytes: u16,
}

/// Parse the 4 bytes of FilePass body that follow `wEncryptionType=0`.
pub(crate) fn parse_filepass_body(body_after_wet: &[u8]) -> CryptoResult<XorParams> {
    if body_after_wet.len() < 4 {
        return Err(CryptoError::InvalidFormat(format!(
            "XOR FilePass body too short: {} (need 4)",
            body_after_wet.len()
        )));
    }
    Ok(XorParams {
        stored_key: u16::from_le_bytes([body_after_wet[0], body_after_wet[1]]),
        verification_bytes: u16::from_le_bytes([body_after_wet[2], body_after_wet[3]]),
    })
}

/// Compute the verification value for a given password and compare it
/// against the FilePass-stored verification bytes.
///
/// The algorithm processes password bytes in **reverse** order, starting
/// with `length` and ending with the first character (length goes through
/// last in the reversed sequence, so first read in forward order).
pub fn verify_password(params: &XorParams, password: &[u8]) -> bool {
    if password.is_empty() || password.len() > MAX_PASSWORD_LEN {
        return false;
    }
    let mut seq = Vec::with_capacity(password.len() + 1);
    seq.push(password.len() as u8);
    seq.extend_from_slice(password);
    seq.reverse();

    let mut verifier: u16 = 0;
    for b in seq {
        let intermediate1 = if verifier & 0x4000 != 0 { 1 } else { 0 };
        let intermediate2 = (verifier << 1) & 0x7FFF;
        let intermediate3 = intermediate1 ^ intermediate2;
        verifier = intermediate3 ^ b as u16;
    }
    (verifier ^ VERIFIER_MAGIC) == params.verification_bytes
}

/// Derive the runtime 16-bit `XorKey` from the password
/// (`CreateXorKey_Method1`, MS-OFFCRYPTO §2.3.7.2).
fn create_xor_key(password: &[u8]) -> u16 {
    debug_assert!(!password.is_empty() && password.len() <= MAX_PASSWORD_LEN);
    let mut key = INITIAL_CODE[password.len() - 1];
    let mut current = (XOR_MATRIX.len() - 1) as i32;

    for &orig in password.iter().rev() {
        let mut ch = orig;
        for _ in 0..7 {
            if ch & 0x40 != 0 {
                key ^= XOR_MATRIX[current as usize];
            }
            ch = ch.wrapping_shl(1);
            current -= 1;
        }
    }
    key
}

/// Rotate-right by 1 bit (8-bit width) of `b1 ^ b2`, used as the
/// "stamp" function during XorArray construction.
fn xor_ror1(b1: u8, b2: u8) -> u8 {
    let v = b1 ^ b2;
    (v >> 1) | ((v & 1) << 7)
}

/// Rotate-right by 5 bits (8-bit width), used during data byte unmasking.
fn ror5(v: u8) -> u8 {
    (v >> 5) | ((v & 0x1F) << 3)
}

/// Build the 16-byte `XorArray` rolling mask
/// (`CreateXorArray_Method1`, MS-OFFCRYPTO §2.3.7.3).
fn create_xor_array(password: &[u8]) -> [u8; 16] {
    debug_assert!(!password.is_empty() && password.len() <= MAX_PASSWORD_LEN);
    let n = password.len();
    let mut obf = [0u8; 16];
    let xor_key = create_xor_key(password);
    let hi = ((xor_key & 0xFF00) >> 8) as u8;
    let lo = (xor_key & 0x00FF) as u8;

    let mut index = n;

    if n % 2 == 1 {
        obf[index] = xor_ror1(PAD_ARRAY[0], hi);
        index -= 1;
        obf[index] = xor_ror1(password[n - 1], lo);
        if index > 0 {
            index -= 1;
        } else {
            index = 0;
        }
    }

    while index > 0 {
        index -= 1;
        obf[index] = xor_ror1(password[index], hi);
        if index == 0 {
            break;
        }
        index -= 1;
        obf[index] = xor_ror1(password[index], lo);
    }

    let mut idx: i32 = 15;
    let mut pad_idx: i32 = 15 - n as i32;
    while pad_idx > 0 {
        obf[idx as usize] = xor_ror1(PAD_ARRAY[pad_idx as usize], hi);
        idx -= 1;
        pad_idx -= 1;
        if pad_idx <= 0 {
            break;
        }
        obf[idx as usize] = xor_ror1(PAD_ARRAY[pad_idx as usize], lo);
        idx -= 1;
        pad_idx -= 1;
    }

    obf
}

/// Decrypt an XLS Workbook stream encrypted with XOR Obfuscation.
pub fn decrypt_workbook_stream(
    stream: &[u8],
    password: &str,
    params: &XorParams,
) -> CryptoResult<Vec<u8>> {
    let pw_bytes = password.as_bytes();
    if pw_bytes.is_empty() {
        return Err(CryptoError::InvalidFormat(
            "XOR Obfuscation does not accept empty passwords".into(),
        ));
    }
    if pw_bytes.len() > MAX_PASSWORD_LEN {
        return Err(CryptoError::InvalidFormat(format!(
            "XOR Obfuscation password must be \u{2264} {} characters (got {})",
            MAX_PASSWORD_LEN,
            pw_bytes.len()
        )));
    }
    if !verify_password(params, pw_bytes) {
        return Err(CryptoError::BadPassword);
    }

    let xor_array = create_xor_array(pw_bytes);
    let classified = record_walk::classify(stream)?;

    // Walk the classified stream byte by byte. For each byte where the
    // overlay is None (encrypted), apply XOR + ROR-5 with the rolling
    // mask. The mask index for each run is set to (stream_offset_of_run_end + adjustment) % 16.
    let mut out = Vec::with_capacity(stream.len());
    let mut i = 0usize;
    while i < classified.overlay.len() {
        match classified.overlay[i] {
            Some(b) => {
                out.push(b);
                i += 1;
            }
            None => {
                // Determine the length of this encrypted run and whether
                // it follows a BoundSheet8 lbPlyPos prefix (4 plaintext
                // bytes immediately before this run, within the same
                // record).
                let run_start = i;
                let mut run_end = i;
                while run_end < classified.overlay.len() && classified.overlay[run_end].is_none() {
                    run_end += 1;
                }
                let count = run_end - run_start;
                let is_bound_sheet_tail = run_start >= 4
                    && classified.overlay[run_start - 4..run_start]
                        .iter()
                        .all(Option::is_some)
                    && classified.overlay[run_start - 8..run_start - 4]
                        .iter()
                        .all(Option::is_some)
                    && (run_start >= 8
                        && classified.ciphertext[run_start - 8] == 0x85
                        && classified.ciphertext[run_start - 7] == 0x00);

                let adjustment = if is_bound_sheet_tail { 4 } else { 0 };
                let mut idx = (run_start + count + adjustment) % 16;
                for k in 0..count {
                    let cb = classified.ciphertext[run_start + k];
                    let unmasked = ror5(cb ^ xor_array[idx]);
                    out.push(unmasked);
                    idx = (idx + 1) % 16;
                }
                i = run_end;
            }
        }
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// MS-OFFCRYPTO test vector: password "VelvetSweatshop" must yield
    /// verifier 0x9A0A (the value Excel writes into FilePass after the
    /// 0xCE4B XOR).
    #[test]
    fn velvet_sweatshop_verifier_is_known_value() {
        let params = XorParams {
            stored_key: 0,
            verification_bytes: 0x9A0A,
        };
        assert!(verify_password(&params, b"VelvetSweatshop"));
    }

    #[test]
    fn verify_rejects_wrong_password() {
        let params = XorParams {
            stored_key: 0,
            verification_bytes: 0x9A0A,
        };
        assert!(!verify_password(&params, b"hunter2"));
    }

    #[test]
    fn verify_rejects_empty_or_overlong_passwords() {
        let params = XorParams {
            stored_key: 0,
            verification_bytes: 0x9A0A,
        };
        assert!(!verify_password(&params, b""));
        assert!(!verify_password(&params, b"this_is_sixteen!"));
    }

    #[test]
    fn create_xor_array_is_deterministic() {
        let a1 = create_xor_array(b"VelvetSweatshop");
        let a2 = create_xor_array(b"VelvetSweatshop");
        assert_eq!(a1, a2);
    }

    #[test]
    fn create_xor_array_varies_with_password() {
        let a = create_xor_array(b"hunter2");
        let b = create_xor_array(b"correct horse");
        assert_ne!(a, b);
    }

    /// Round-trip XOR (since the algorithm is its own inverse if you
    /// reverse the rotate direction): encrypt a record body using the
    /// inverse operations, then decrypt and confirm we recover the
    /// plaintext.
    #[test]
    fn round_trip_one_record() {
        // Construct a stream: BOF (excluded, plaintext body) + a 0x0203
        // (NUMBER) record we'll fully encrypt.
        let mut stream = Vec::new();
        stream.extend_from_slice(&0x0809u16.to_le_bytes()); // BOF type
        stream.extend_from_slice(&8u16.to_le_bytes()); // size
        stream.extend_from_slice(&[0u8; 8]); // BOF body
        let plain_body = [
            0x11u8, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xAA, 0xBB, 0xCC, 0xDD, 0xEE,
        ];
        stream.extend_from_slice(&0x0203u16.to_le_bytes()); // NUMBER type
        stream.extend_from_slice(&(plain_body.len() as u16).to_le_bytes()); // size
        let body_offset = stream.len();
        // Encrypt body with the XOR algorithm in reverse: ROL-5 then XOR
        // (so decrypt = XOR then ROR-5 inverts it).
        let pw = b"hunter2";
        let xor_array = create_xor_array(pw);
        let count = plain_body.len();
        let mut idx = (body_offset + count) % 16;
        for &b in &plain_body {
            let rolled = (b << 5) | (b >> 3);
            let masked = rolled ^ xor_array[idx];
            stream.push(masked);
            idx = (idx + 1) % 16;
        }

        // Decrypt and verify.
        let params = XorParams {
            stored_key: 0,
            verification_bytes: {
                // Compute matching VerifyPassword output for "hunter2".
                let mut seq = vec![pw.len() as u8];
                seq.extend_from_slice(pw);
                seq.reverse();
                let mut v: u16 = 0;
                for b in seq {
                    let i1 = if v & 0x4000 != 0 { 1 } else { 0 };
                    let i2 = (v << 1) & 0x7FFF;
                    v = (i1 ^ i2) ^ b as u16;
                }
                v ^ VERIFIER_MAGIC
            },
        };
        let decrypted = decrypt_workbook_stream(&stream, "hunter2", &params)
            .expect("decrypt round-trip should succeed");

        // Decrypted body should match plaintext.
        assert_eq!(&decrypted[body_offset..body_offset + count], &plain_body);
        // Header preserved.
        assert_eq!(decrypted[12], 0x03);
        assert_eq!(decrypted[13], 0x02);
    }
}
