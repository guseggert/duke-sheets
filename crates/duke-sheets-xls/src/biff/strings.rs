//! BIFF8 Unicode string decoding.
//!
//! BIFF8 strings have a complex encoding:
//! - Header: char_count (2 bytes) + flags (1 byte)
//! - Flags bit 0 (`fHighByte`): 0 = compressed Latin-1, 1 = uncompressed UTF-16LE
//! - Flags bit 2 (`fExtSt`): extended string data follows (Asian phonetic)
//! - Flags bit 3 (`fRichSt`): rich text run array follows
//! - If fRichSt: 2-byte run count follows the flags
//! - If fExtSt: 4-byte extended data size follows
//! - Then the character data
//! - Then the rich text runs (4 bytes each) if fRichSt
//! - Then the extended data if fExtSt
//!
//! In SST records, strings can span CONTINUE records. The CONTINUE record
//! can change the encoding (compressed ↔ uncompressed) mid-string via a
//! new flags byte at the start of the continuation.

use super::parser::{read_u16, read_u32, read_u8};
use crate::error::{XlsError, XlsResult};

/// Decoded result of reading a BIFF8 Unicode string.
#[derive(Debug, Clone)]
pub struct BiffString {
    pub text: String,
    /// Total bytes consumed from the buffer (including header, runs, ext data).
    pub bytes_consumed: usize,
}

/// Read a BIFF8 "short" string (1-byte length prefix, used in BOUNDSHEET etc.).
pub fn read_short_string(data: &[u8], offset: &mut usize) -> XlsResult<String> {
    let char_count = read_u8(data, offset)? as u16;
    let flags = read_u8(data, offset)?;
    read_character_data(data, offset, char_count, flags)
}

/// Read a BIFF8 Unicode string with a 2-byte length prefix (used in SST, LABEL, etc.).
///
/// Returns the decoded string. This does NOT handle CONTINUE boundaries —
/// use `read_sst_string` for SST records that may span continuations.
pub fn read_unicode_string(data: &[u8], offset: &mut usize) -> XlsResult<String> {
    let char_count = read_u16(data, offset)?;
    let flags = read_u8(data, offset)?;

    let is_rich = (flags & 0x08) != 0;
    let has_ext = (flags & 0x04) != 0;

    let run_count = if is_rich { read_u16(data, offset)? } else { 0 };
    let ext_size = if has_ext { read_u32(data, offset)? } else { 0 };

    let text = read_character_data(data, offset, char_count, flags)?;

    // Skip rich text runs (4 bytes each: char_pos u16 + font_idx u16)
    if is_rich {
        *offset += run_count as usize * 4;
    }
    // Skip extended string data
    if has_ext {
        *offset += ext_size as usize;
    }

    Ok(text)
}

/// Read character data (no header) given char_count and flags byte.
fn read_character_data(
    data: &[u8],
    offset: &mut usize,
    char_count: u16,
    flags: u8,
) -> XlsResult<String> {
    let is_wide = (flags & 0x01) != 0;
    let count = char_count as usize;

    if is_wide {
        // UTF-16LE: 2 bytes per character
        let byte_len = count * 2;
        if *offset + byte_len > data.len() {
            return Err(XlsError::Parse(format!(
                "string data too short: need {} bytes at offset {}, have {}",
                byte_len,
                *offset,
                data.len() - *offset
            )));
        }
        let mut chars = Vec::with_capacity(count);
        for i in 0..count {
            let lo = data[*offset + i * 2];
            let hi = data[*offset + i * 2 + 1];
            chars.push(u16::from_le_bytes([lo, hi]));
        }
        *offset += byte_len;
        String::from_utf16(&chars)
            .map_err(|e| XlsError::Parse(format!("invalid UTF-16 string: {e}")))
    } else {
        // Compressed Latin-1: 1 byte per character
        if *offset + count > data.len() {
            return Err(XlsError::Parse(format!(
                "string data too short: need {} bytes at offset {}, have {}",
                count,
                *offset,
                data.len() - *offset
            )));
        }
        let s: String = data[*offset..*offset + count]
            .iter()
            .map(|&b| b as char)
            .collect();
        *offset += count;
        Ok(s)
    }
}

/// Parse the entire SST (Shared String Table) from a concatenated buffer
/// (SST body + all CONTINUE bodies already joined).
///
/// The SST body starts with:
/// - `total_strings` (4 bytes, u32) — total string refs in workbook
/// - `unique_strings` (4 bytes, u32) — number of unique strings in this table
/// - Then `unique_strings` Unicode string entries
pub fn parse_sst(data: &[u8]) -> XlsResult<Vec<String>> {
    let mut offset = 0;

    let _total_strings = read_u32(data, &mut offset)?;
    let unique_count = read_u32(data, &mut offset)? as usize;

    let mut strings = Vec::with_capacity(unique_count);

    for i in 0..unique_count {
        match read_unicode_string(data, &mut offset) {
            Ok(s) => strings.push(s),
            Err(e) => {
                // If we hit a parse error near the end, log and stop.
                // Some XLS files have SST padding or truncation issues.
                log::warn!("SST parse error at string {i}/{unique_count}: {e}");
                break;
            }
        }
    }

    Ok(strings)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_read_compressed_string() {
        // 3-char compressed string "ABC"
        // char_count = 3 (u16 LE), flags = 0x00, data = "ABC"
        let data = [0x03, 0x00, 0x00, b'A', b'B', b'C'];
        let mut offset = 0;
        let s = read_unicode_string(&data, &mut offset).unwrap();
        assert_eq!(s, "ABC");
        assert_eq!(offset, 6);
    }

    #[test]
    fn test_read_wide_string() {
        // 2-char UTF-16 string "Hi"
        // char_count = 2 (u16 LE), flags = 0x01, data = H\0i\0
        let data = [0x02, 0x00, 0x01, b'H', 0x00, b'i', 0x00];
        let mut offset = 0;
        let s = read_unicode_string(&data, &mut offset).unwrap();
        assert_eq!(s, "Hi");
        assert_eq!(offset, 7);
    }

    #[test]
    fn test_read_short_string() {
        // 1-byte length prefix: 2 chars, compressed
        let data = [0x02, 0x00, b'O', b'K'];
        let mut offset = 0;
        let s = read_short_string(&data, &mut offset).unwrap();
        assert_eq!(s, "OK");
    }

    #[test]
    fn test_parse_sst() {
        // SST with 2 total refs, 2 unique strings: "A" and "BC"
        let mut buf = Vec::new();
        buf.extend_from_slice(&2u32.to_le_bytes()); // total
        buf.extend_from_slice(&2u32.to_le_bytes()); // unique
                                                    // String "A": char_count=1, flags=0, data='A'
        buf.extend_from_slice(&[0x01, 0x00, 0x00, b'A']);
        // String "BC": char_count=2, flags=0, data='BC'
        buf.extend_from_slice(&[0x02, 0x00, 0x00, b'B', b'C']);

        let strings = parse_sst(&buf).unwrap();
        assert_eq!(strings, vec!["A", "BC"]);
    }
}
