//! Record-walking helper shared by all XLS FilePass-based decryption
//! variants (legacy RC4, RC4 CryptoAPI, XOR).
//!
//! The "two-buffer trick" from msoffcrypto-tool: build a ciphertext
//! buffer that's co-indexed with the raw Workbook stream but has zeros
//! in plaintext regions, and a parallel overlay buffer that has the
//! known plaintext in those positions and `None` where decryption
//! applies. The cipher runs over the whole ciphertext buffer,
//! maintaining keystream alignment across plaintext regions, and the
//! overlay is applied on top of the decrypted output.
//!
//! This design preserves absolute stream offsets — critical because
//! `BoundSheet8.lbPlyPos` points to absolute byte positions of sheet
//! BOFs.

use crate::error::{CryptoError, CryptoResult};

/// BIFF8 record types whose bodies MUST remain plaintext in an encrypted
/// Workbook stream (MS-XLS §2.2.10).
const EXCLUDED_RECORD_TYPES: &[u16] = &[
    0x0809, // BOF
    0x002F, // FilePass (body zeroed in output; see decrypt_workbook_stream)
    0x0194, // UsrExcl
    0x0195, // FileLock
    0x00E1, // InterfaceHdr
    0x0196, // RRDInfo
    0x0138, // RRDHead
];

/// BoundSheet8: first 4 bytes of body (`lbPlyPos`) plaintext, rest encrypted.
const BOUND_SHEET_8: u16 = 0x0085;

/// Direction in which a Workbook stream is being classified.
///
/// The two directions agree on every record except `FilePass`. On
/// decrypt, the FilePass record is neutered: type bytes overwritten
/// with zeros and body zeroed, so downstream BIFF parsers ignore it.
/// On encrypt, the FilePass record's real bytes pass through verbatim
/// — the caller has already built and inserted it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Direction {
    Decrypt,
    Encrypt,
}

/// Classified workbook stream: ciphertext with zero-filled plaintext
/// holes, plus a co-indexed overlay of known plaintext bytes.
pub(crate) struct ClassifiedStream {
    pub ciphertext: Vec<u8>,
    pub overlay: Vec<Option<u8>>,
}

/// Walk the records in a workbook stream and classify each byte as
/// either "plaintext" (provide overlay byte) or "encrypted" (overlay
/// is `None`, ciphertext buffer holds the real bytes).
pub(crate) fn classify(stream: &[u8], direction: Direction) -> CryptoResult<ClassifiedStream> {
    let mut ciphertext = Vec::with_capacity(stream.len());
    let mut overlay: Vec<Option<u8>> = Vec::with_capacity(stream.len());

    let mut cursor = 0usize;
    while cursor + 4 <= stream.len() {
        let record_type = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let size = u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]) as usize;
        let body_start = cursor + 4;
        let body_end = body_start + size;

        if body_end > stream.len() {
            return Err(CryptoError::InvalidFormat(format!(
                "record at offset {cursor:#x} extends past stream end: type={record_type:#06x} size={size}"
            )));
        }

        if record_type == 0x002F && direction == Direction::Decrypt {
            // FilePass on decrypt: neuter the record so downstream
            // parsers treat the stream as plaintext. Zero the 2-byte
            // type (`0x002F` becomes `0x0000` - an unknown record that
            // BIFF readers skip) but keep the 2-byte size so every
            // absolute offset (notably `BoundSheet8.lbPlyPos`) remains
            // valid. On encrypt, the FilePass record falls through to
            // the EXCLUDED_RECORD_TYPES branch below: its real bytes
            // are preserved verbatim because the caller built them.
            ciphertext.extend_from_slice(&[0, 0, 0, 0]);
            overlay.extend_from_slice(&[
                Some(0),
                Some(0),
                Some(stream[cursor + 2]),
                Some(stream[cursor + 3]),
            ]);
            ciphertext.resize(ciphertext.len() + size, 0);
            overlay.extend(std::iter::repeat(Some(0u8)).take(size));
            cursor = body_end;
            continue;
        }

        // Record header is always plaintext.
        ciphertext.extend_from_slice(&[0, 0, 0, 0]);
        overlay.extend_from_slice(&[
            Some(stream[cursor]),
            Some(stream[cursor + 1]),
            Some(stream[cursor + 2]),
            Some(stream[cursor + 3]),
        ]);

        if EXCLUDED_RECORD_TYPES.contains(&record_type) {
            // Plaintext body: cipher runs over zero placeholders.
            ciphertext.resize(ciphertext.len() + size, 0);
            overlay.extend(stream[body_start..body_end].iter().map(|b| Some(*b)));
        } else if record_type == BOUND_SHEET_8 && size >= 4 {
            // First 4 bytes (lbPlyPos) plaintext, rest encrypted.
            ciphertext.extend_from_slice(&[0, 0, 0, 0]);
            overlay.extend(stream[body_start..body_start + 4].iter().map(|b| Some(*b)));
            ciphertext.extend_from_slice(&stream[body_start + 4..body_end]);
            overlay.extend(std::iter::repeat(None).take(size - 4));
        } else {
            // Fully encrypted record body.
            ciphertext.extend_from_slice(&stream[body_start..body_end]);
            overlay.extend(std::iter::repeat(None).take(size));
        }

        cursor = body_end;
    }

    // Trailing bytes that don't form a complete record header are
    // treated as encrypted. In practice the Workbook stream is always
    // record-aligned, so this is a defensive path.
    if cursor < stream.len() {
        ciphertext.extend_from_slice(&stream[cursor..]);
        overlay.extend(std::iter::repeat(None).take(stream.len() - cursor));
    }

    debug_assert_eq!(ciphertext.len(), stream.len());
    debug_assert_eq!(overlay.len(), stream.len());

    Ok(ClassifiedStream {
        ciphertext,
        overlay,
    })
}

/// Apply the plaintext overlay onto a decrypted buffer, restoring known
/// plaintext bytes.
pub(crate) fn apply_overlay(mut decrypted: Vec<u8>, overlay: &[Option<u8>]) -> Vec<u8> {
    for (i, o) in overlay.iter().enumerate() {
        if let Some(b) = o {
            decrypted[i] = *b;
        }
    }
    decrypted
}
