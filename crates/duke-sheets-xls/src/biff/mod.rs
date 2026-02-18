//! BIFF8 (Binary Interchange File Format) handling.
//!
//! This module provides the record-level abstraction for reading BIFF8 streams.
//! A BIFF8 stream is a sequence of records, each with a 4-byte header
//! (2 bytes record type + 2 bytes body length) followed by the body.
//!
//! CONTINUE records (type 0x003C) extend the body of the preceding record
//! beyond the 8224-byte per-record limit.

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
            // Append to the previous record's data
            if let Some(prev) = records.last_mut() {
                prev.data.extend_from_slice(&body);
            }
            // If there's no previous record, we just drop the orphaned CONTINUE
        } else {
            records.push(BiffRecord {
                record_type,
                data: body,
                stream_offset,
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
