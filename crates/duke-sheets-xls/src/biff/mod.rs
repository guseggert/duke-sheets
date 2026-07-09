//! BIFF8 (Binary Interchange File Format) handling.
//!
//! This module provides the record-level abstraction for reading BIFF8 streams.
//! A BIFF8 stream is a sequence of records, each with a 4-byte header
//! (2 bytes record type + 2 bytes body length) followed by the body.
//!
//! CONTINUE records (type 0x003C) extend the body of the preceding record
//! beyond the 8224-byte per-record limit.

pub mod escher;
pub mod formula;
pub mod obj;
pub mod parser;
pub mod records;
pub mod strings;

use crate::error::{XlsError, XlsResult};
use std::io::{Read, Seek};

/// A single BIFF8 record (with CONTINUE bodies already merged).
#[derive(Debug)]
pub struct BiffRecord {
    /// Record type ID (e.g. `records::SST`, `records::NUMBER`).
    pub record_type: u16,
    /// Record body bytes (CONTINUE records have been concatenated).
    pub data: Vec<u8>,
    /// Byte offset of this record's header in the stream (for debugging).
    pub stream_offset: u64,
    /// Byte offsets within `data` where each CONTINUE record body begins.
    /// Empty if the record had no CONTINUE extensions.
    /// These offsets are needed by the SST parser to handle encoding changes
    /// at CONTINUE boundaries.
    pub continue_offsets: Vec<usize>,
}

/// Reads all BIFF8 records from a byte stream, merging CONTINUE records
/// into their parent.
///
/// Returns the records in order. Each record's `data` field contains the
/// full body (including any CONTINUE extensions).
pub fn read_all_records<R: Read + Seek>(stream: &mut R) -> XlsResult<Vec<BiffRecord>> {
    let mut records: Vec<BiffRecord> = Vec::new();
    let mut header_buf = [0u8; 4];

    loop {
        let stream_offset = stream.stream_position().map_err(|e| XlsError::Io(e))?;

        // Read 4-byte record header
        match stream.read_exact(&mut header_buf) {
            Ok(()) => {}
            Err(e) if e.kind() == std::io::ErrorKind::UnexpectedEof => break,
            Err(e) => return Err(XlsError::Io(e)),
        }

        let record_type = u16::from_le_bytes([header_buf[0], header_buf[1]]);
        let body_len = u16::from_le_bytes([header_buf[2], header_buf[3]]) as usize;

        // Read body
        let mut body = vec![0u8; body_len];
        if body_len > 0 {
            stream.read_exact(&mut body).map_err(|e| XlsError::Io(e))?;
        }

        if record_type == records::CONTINUE {
            // Append to the previous record's data, tracking the boundary offset
            if let Some(prev) = records.last_mut() {
                prev.continue_offsets.push(prev.data.len());
                prev.data.extend_from_slice(&body);
            }
            // If there's no previous record, we just drop the orphaned CONTINUE
        } else {
            records.push(BiffRecord {
                record_type,
                data: body,
                stream_offset,
                continue_offsets: Vec::new(),
            });
        }
    }

    Ok(records)
}

/// Extract the BOF record fields from a record body.
///
/// Returns `(version, substream_type)`.
/// - `version` should be `0x0600` for BIFF8
/// - `substream_type`: 0x0005 = workbook globals, 0x0010 = worksheet, etc.
pub fn parse_bof(data: &[u8]) -> XlsResult<(u16, u16)> {
    if data.len() < 4 {
        return Err(XlsError::InvalidFormat("BOF record too short".into()));
    }
    let version = u16::from_le_bytes([data[0], data[1]]);
    let dt = u16::from_le_bytes([data[2], data[3]]);
    Ok((version, dt))
}

/// Check the workbook globals block for a FILEPASS record and return a clean
/// `Encrypted` error if one is present. Per [MS-XLS] §2.1.7.4 a FILEPASS record
/// directly after the globals BOF marks the entire stream (except the header
/// records) as ciphertext. Without decryption we cannot make sense of anything
/// that follows, so fail fast rather than emitting garbage parse errors.
pub fn check_not_encrypted(records: &[BiffRecord]) -> XlsResult<()> {
    for rec in records {
        if rec.record_type == records::EOF {
            // Only inspect the globals block; subsequent EOF-terminated substreams
            // cannot introduce new encryption.
            break;
        }
        if rec.record_type == records::FILEPASS {
            let enc_kind = if rec.data.len() >= 2 {
                match u16::from_le_bytes([rec.data[0], rec.data[1]]) {
                    0 => "XOR obfuscation",
                    1 => "RC4",
                    other => {
                        return Err(XlsError::Encrypted(format!(
                            "workbook is password-protected (unknown encryption type {other:#x})"
                        )));
                    }
                }
            } else {
                "unknown"
            };
            return Err(XlsError::Encrypted(format!(
                "workbook is password-protected ({enc_kind})"
            )));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(ty: u16, data: Vec<u8>) -> BiffRecord {
        BiffRecord {
            record_type: ty,
            data,
            stream_offset: 0,
            continue_offsets: Vec::new(),
        }
    }

    #[test]
    fn check_not_encrypted_accepts_clean_stream() {
        let recs = vec![
            record(records::BOF, vec![0x00, 0x06, 0x05, 0x00, 0, 0, 0, 0]),
            record(records::FONT, vec![0; 26]),
            record(records::EOF, vec![]),
        ];
        check_not_encrypted(&recs).expect("clean stream must parse");
    }

    #[test]
    fn check_not_encrypted_rejects_rc4() {
        // Encryption type 1 = RC4.
        let recs = vec![
            record(records::BOF, vec![0x00, 0x06, 0x05, 0x00, 0, 0, 0, 0]),
            record(records::FILEPASS, vec![0x01, 0x00]),
            record(records::EOF, vec![]),
        ];
        let err = check_not_encrypted(&recs).unwrap_err();
        match err {
            XlsError::Encrypted(msg) => assert!(msg.contains("RC4"), "msg={msg}"),
            other => panic!("expected Encrypted, got {other:?}"),
        }
    }

    #[test]
    fn check_not_encrypted_rejects_xor_obfuscation() {
        // Encryption type 0 = XOR obfuscation.
        let recs = vec![
            record(records::BOF, vec![0x00, 0x06, 0x05, 0x00, 0, 0, 0, 0]),
            record(records::FILEPASS, vec![0x00, 0x00]),
        ];
        let err = check_not_encrypted(&recs).unwrap_err();
        match err {
            XlsError::Encrypted(msg) => assert!(msg.contains("XOR"), "msg={msg}"),
            other => panic!("expected Encrypted, got {other:?}"),
        }
    }

    #[test]
    fn check_not_encrypted_ignores_filepass_after_globals_eof() {
        // FILEPASS must appear in the globals block (before first EOF).
        // A stray FILEPASS-like record later in the stream is not a valid
        // encryption marker and must not trigger a false positive.
        let recs = vec![
            record(records::BOF, vec![0x00, 0x06, 0x05, 0x00, 0, 0, 0, 0]),
            record(records::EOF, vec![]),
            record(records::FILEPASS, vec![0x01, 0x00]),
        ];
        check_not_encrypted(&recs).expect("FILEPASS after globals EOF must not trigger");
    }
}
