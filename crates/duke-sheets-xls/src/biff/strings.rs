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

/// A formatting run within an SST string: start position + font index.
#[derive(Debug, Clone, PartialEq)]
pub struct FormattingRun {
    /// 0-based character position where this run begins.
    pub char_pos: u16,
    /// Font index into the workbook's FONT record table.
    pub font_index: u16,
}

/// An entry in the Shared String Table.
#[derive(Debug, Clone)]
pub enum SstEntry {
    /// Plain text string with no formatting runs.
    Plain(String),
    /// Rich text: the full text plus formatting run markers.
    Rich {
        text: String,
        runs: Vec<FormattingRun>,
    },
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
///
/// This is the low-level decoder: given a character count and a flags byte
/// (bit 0 = wide/compressed), reads the appropriate number of bytes and
/// returns the decoded string.
pub(crate) fn read_character_data(
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
/// (SST body + all CONTINUE bodies already joined), without CONTINUE awareness.
///
/// This is the legacy parser that treats the buffer as flat. Use
/// [`parse_sst_continued`] for correct handling of SST records with
/// CONTINUE extensions that may change string encoding at boundaries.
pub fn parse_sst(data: &[u8]) -> XlsResult<Vec<String>> {
    parse_sst_continued(data, &[])
}

/// Parse the SST with full CONTINUE boundary awareness.
///
/// When a string's character data spans a CONTINUE record boundary, the
/// first byte at the boundary is a flags byte that may change the encoding
/// (Latin-1 ↔ UTF-16LE). This parser correctly handles those transitions.
///
/// The SST body starts with:
/// - `total_strings` (4 bytes, u32) — total string refs in workbook
/// - `unique_strings` (4 bytes, u32) — number of unique strings in this table
/// - Then `unique_strings` Unicode string entries
pub fn parse_sst_continued(data: &[u8], continue_offsets: &[usize]) -> XlsResult<Vec<String>> {
    let mut reader = ContinueReader::new(data, continue_offsets);

    let _total_strings = reader.read_u32()?;
    let unique_count = reader.read_u32()? as usize;

    let mut strings = Vec::with_capacity(unique_count);

    for i in 0..unique_count {
        match reader.read_sst_string() {
            Ok(s) => strings.push(s),
            Err(e) => {
                log::warn!("SST parse error at string {i}/{unique_count}: {e}");
                break;
            }
        }
    }

    Ok(strings)
}

/// Parse the SST with full CONTINUE boundary awareness, preserving formatting runs.
///
/// Like `parse_sst_continued`, but returns `SstEntry` variants that carry
/// rich text formatting run data (character position + font index pairs)
/// instead of stripping them.
pub fn parse_sst_entries(data: &[u8], continue_offsets: &[usize]) -> XlsResult<Vec<SstEntry>> {
    let mut reader = ContinueReader::new(data, continue_offsets);

    let _total_strings = reader.read_u32()?;
    let unique_count = reader.read_u32()? as usize;

    let mut entries = Vec::with_capacity(unique_count);

    for i in 0..unique_count {
        match reader.read_sst_entry() {
            Ok(entry) => entries.push(entry),
            Err(e) => {
                log::warn!("SST parse error at string {i}/{unique_count}: {e}");
                break;
            }
        }
    }

    Ok(entries)
}

// ─── ContinueReader ─────────────────────────────────────────────────────

/// A reader over a concatenated BIFF8 record buffer that tracks CONTINUE
/// record boundaries.
///
/// For character data, CONTINUE boundaries contain a 1-byte flags field
/// that may change the string encoding. For all other data (headers, rich
/// text runs, extended data), boundaries are transparent.
struct ContinueReader<'a> {
    data: &'a [u8],
    /// Sorted byte offsets within `data` where each CONTINUE body begins.
    boundaries: &'a [usize],
    /// Current read position.
    pos: usize,
}

impl<'a> ContinueReader<'a> {
    fn new(data: &'a [u8], boundaries: &'a [usize]) -> Self {
        Self {
            data,
            boundaries,
            pos: 0,
        }
    }

    /// Bytes from current position to the next CONTINUE boundary, or to end of
    /// data if no more boundaries ahead.
    fn bytes_to_next_boundary(&self) -> usize {
        for &b in self.boundaries {
            if b > self.pos {
                return b - self.pos;
            }
        }
        self.data.len() - self.pos
    }

    // ── Raw reads (boundaries are transparent) ──────────────────────

    fn read_u8(&mut self) -> XlsResult<u8> {
        if self.pos >= self.data.len() {
            return Err(XlsError::Parse("unexpected end of data reading u8".into()));
        }
        let v = self.data[self.pos];
        self.pos += 1;
        Ok(v)
    }

    fn read_u16(&mut self) -> XlsResult<u16> {
        if self.pos + 2 > self.data.len() {
            return Err(XlsError::Parse("unexpected end of data reading u16".into()));
        }
        let v = u16::from_le_bytes([self.data[self.pos], self.data[self.pos + 1]]);
        self.pos += 2;
        Ok(v)
    }

    fn read_u32(&mut self) -> XlsResult<u32> {
        if self.pos + 4 > self.data.len() {
            return Err(XlsError::Parse("unexpected end of data reading u32".into()));
        }
        let v = u32::from_le_bytes([
            self.data[self.pos],
            self.data[self.pos + 1],
            self.data[self.pos + 2],
            self.data[self.pos + 3],
        ]);
        self.pos += 4;
        Ok(v)
    }

    /// Advance position by `n` bytes (for run data / ext data — no flags
    /// byte at boundaries).
    fn skip(&mut self, n: usize) {
        self.pos += n;
    }

    // ── CONTINUE-aware character data reading ───────────────────────

    /// Read a full SST Unicode string, handling CONTINUE boundaries in
    /// the character data portion.
    fn read_sst_string(&mut self) -> XlsResult<String> {
        let char_count = self.read_u16()? as usize;
        let flags = self.read_u8()?;

        let is_rich = (flags & 0x08) != 0;
        let has_ext = (flags & 0x04) != 0;
        let is_wide = (flags & 0x01) != 0;

        let run_count = if is_rich {
            self.read_u16()? as usize
        } else {
            0
        };
        let ext_size = if has_ext {
            self.read_u32()? as usize
        } else {
            0
        };

        // Character data — this is where CONTINUE encoding changes happen.
        let text = self.read_chars(char_count, is_wide)?;

        // Rich text runs (4 bytes each) — raw data, no flags at boundaries.
        if run_count > 0 {
            self.skip(run_count * 4);
        }
        // Extended string data — raw data, no flags at boundaries.
        if ext_size > 0 {
            self.skip(ext_size);
        }

        Ok(text)
    }

    /// Read a full SST Unicode string entry, preserving formatting runs.
    fn read_sst_entry(&mut self) -> XlsResult<SstEntry> {
        let char_count = self.read_u16()? as usize;
        let flags = self.read_u8()?;

        let is_rich = (flags & 0x08) != 0;
        let has_ext = (flags & 0x04) != 0;
        let is_wide = (flags & 0x01) != 0;

        let run_count = if is_rich {
            self.read_u16()? as usize
        } else {
            0
        };
        let ext_size = if has_ext {
            self.read_u32()? as usize
        } else {
            0
        };

        // Character data — this is where CONTINUE encoding changes happen.
        let text = self.read_chars(char_count, is_wide)?;

        // Rich text runs (4 bytes each: char_pos u16 + font_idx u16).
        let mut formatting_runs = Vec::new();
        if run_count > 0 {
            for _ in 0..run_count {
                if self.pos + 4 <= self.data.len() {
                    let char_pos =
                        u16::from_le_bytes([self.data[self.pos], self.data[self.pos + 1]]);
                    let font_index =
                        u16::from_le_bytes([self.data[self.pos + 2], self.data[self.pos + 3]]);
                    formatting_runs.push(FormattingRun {
                        char_pos,
                        font_index,
                    });
                    self.pos += 4;
                } else {
                    // Truncated run data — skip what we can
                    self.skip(run_count.saturating_sub(formatting_runs.len()) * 4);
                    break;
                }
            }
        }
        // Extended string data — raw data, no flags at boundaries.
        if ext_size > 0 {
            self.skip(ext_size);
        }

        if formatting_runs.is_empty() {
            Ok(SstEntry::Plain(text))
        } else {
            Ok(SstEntry::Rich {
                text,
                runs: formatting_runs,
            })
        }
    }

    /// Read `char_count` characters from the buffer, handling encoding
    /// changes at CONTINUE boundaries.
    ///
    /// At each CONTINUE boundary crossed during character reading, a 1-byte
    /// flags field is consumed. Bit 0 of that byte indicates the encoding
    /// for the remaining characters (0 = Latin-1, 1 = UTF-16LE).
    fn read_chars(&mut self, char_count: usize, initial_wide: bool) -> XlsResult<String> {
        if char_count == 0 {
            return Ok(String::new());
        }

        let mut code_units: Vec<u16> = Vec::with_capacity(char_count);
        let mut remaining = char_count;
        let mut wide = initial_wide;

        while remaining > 0 {
            if self.pos >= self.data.len() {
                return Err(XlsError::Parse(format!(
                    "SST string data truncated: {} chars remaining at offset {}",
                    remaining, self.pos
                )));
            }

            let avail = self.bytes_to_next_boundary();
            let char_width: usize = if wide { 2 } else { 1 };
            let chars_can_fit = avail / char_width;
            let n = remaining.min(chars_can_fit);

            // Read n characters in the current encoding.
            for _ in 0..n {
                if wide {
                    let lo = self.data[self.pos];
                    let hi = self.data[self.pos + 1];
                    code_units.push(u16::from_le_bytes([lo, hi]));
                    self.pos += 2;
                } else {
                    code_units.push(self.data[self.pos] as u16);
                    self.pos += 1;
                }
            }
            remaining -= n;

            if remaining > 0 {
                // We need to cross a CONTINUE boundary.
                // Skip any leftover padding bytes (e.g. 1 byte when a wide
                // char can't fit before the boundary).
                let leftover = avail - n * char_width;
                self.pos += leftover;

                // Consume the flags byte at the boundary.
                if self.pos >= self.data.len() {
                    return Err(XlsError::Parse(
                        "SST string truncated at CONTINUE boundary".into(),
                    ));
                }
                let new_flags = self.data[self.pos];
                self.pos += 1;
                wide = (new_flags & 0x01) != 0;
            }
        }

        String::from_utf16(&code_units)
            .map_err(|e| XlsError::Parse(format!("invalid UTF-16 in SST string: {e}")))
    }
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

    // ── CONTINUE boundary tests ─────────────────────────────────────

    /// Helper: build an SST buffer from segments (simulating record + CONTINUE bodies).
    /// Returns (concatenated_data, continue_offsets).
    fn build_sst_segments(segments: &[&[u8]]) -> (Vec<u8>, Vec<usize>) {
        let mut data = Vec::new();
        let mut offsets = Vec::new();
        for (i, seg) in segments.iter().enumerate() {
            if i > 0 {
                offsets.push(data.len());
            }
            data.extend_from_slice(seg);
        }
        (data, offsets)
    }

    #[test]
    fn test_sst_continue_same_encoding() {
        // SST with 1 string "ABCDE" (5 chars, compressed).
        // Split: first record has "AB", CONTINUE has "CDE" (same encoding).
        //
        // Record body: total(4) + unique(4) + header(3: char_count=5, flags=0) + "AB"
        // CONTINUE body: flags_byte(0x00) + "CDE"
        let mut seg0 = Vec::new();
        seg0.extend_from_slice(&1u32.to_le_bytes()); // total_strings
        seg0.extend_from_slice(&1u32.to_le_bytes()); // unique_strings
        seg0.extend_from_slice(&5u16.to_le_bytes()); // char_count = 5
        seg0.push(0x00); // flags: compressed
        seg0.extend_from_slice(b"AB"); // first 2 chars

        let seg1: &[u8] = &[
            0x00, // flags byte at CONTINUE boundary: still compressed
            b'C', b'D', b'E', // remaining 3 chars
        ];

        let (data, offsets) = build_sst_segments(&[&seg0, seg1]);
        let strings = parse_sst_continued(&data, &offsets).unwrap();
        assert_eq!(strings, vec!["ABCDE"]);
    }

    #[test]
    fn test_sst_continue_latin1_to_utf16() {
        // SST with 1 string "AB\u{00E9}" (3 chars).
        // Starts compressed (Latin-1) with "AB", then CONTINUE switches to
        // UTF-16LE for the accented char "é" (U+00E9).
        //
        // Record body: header + "AB"
        // CONTINUE body: flags(0x01=wide) + é as UTF-16LE (0xE9, 0x00)
        let mut seg0 = Vec::new();
        seg0.extend_from_slice(&1u32.to_le_bytes());
        seg0.extend_from_slice(&1u32.to_le_bytes());
        seg0.extend_from_slice(&3u16.to_le_bytes()); // 3 chars total
        seg0.push(0x00); // flags: compressed
        seg0.extend_from_slice(b"AB");

        let seg1: &[u8] = &[
            0x01, // flags: wide (UTF-16LE)
            0xE9, 0x00, // 'é' in UTF-16LE
        ];

        let (data, offsets) = build_sst_segments(&[&seg0, seg1]);
        let strings = parse_sst_continued(&data, &offsets).unwrap();
        assert_eq!(strings, vec!["AB\u{00E9}"]);
    }

    #[test]
    fn test_sst_continue_utf16_to_latin1() {
        // SST with 1 string: starts UTF-16LE, CONTINUE switches to Latin-1.
        // "H" as wide + "ello" as compressed.
        let mut seg0 = Vec::new();
        seg0.extend_from_slice(&1u32.to_le_bytes());
        seg0.extend_from_slice(&1u32.to_le_bytes());
        seg0.extend_from_slice(&5u16.to_le_bytes()); // 5 chars total
        seg0.push(0x01); // flags: wide
        seg0.extend_from_slice(&[b'H', 0x00]); // 'H' as UTF-16LE

        let seg1: &[u8] = &[
            0x00, // flags: compressed (Latin-1)
            b'e', b'l', b'l', b'o',
        ];

        let (data, offsets) = build_sst_segments(&[&seg0, seg1]);
        let strings = parse_sst_continued(&data, &offsets).unwrap();
        assert_eq!(strings, vec!["Hello"]);
    }

    #[test]
    fn test_sst_continue_multiple_strings_across_boundary() {
        // SST with 3 strings: "Hi", "World", "!" — "World" spans the boundary.
        // Record body: header + "Hi" + partial "World" header + "Wo"
        // CONTINUE body: flags(0x00) + "rld" + "!" header + "!"
        let mut seg0 = Vec::new();
        seg0.extend_from_slice(&3u32.to_le_bytes()); // total
        seg0.extend_from_slice(&3u32.to_le_bytes()); // unique = 3

        // String "Hi"
        seg0.extend_from_slice(&2u16.to_le_bytes());
        seg0.push(0x00);
        seg0.extend_from_slice(b"Hi");

        // String "World" (5 chars) — header + first 2 chars in this segment
        seg0.extend_from_slice(&5u16.to_le_bytes());
        seg0.push(0x00);
        seg0.extend_from_slice(b"Wo");

        // CONTINUE: flags byte + remaining 3 chars of "World" + string "!"
        let mut seg1 = Vec::new();
        seg1.push(0x00); // flags: still compressed
        seg1.extend_from_slice(b"rld");

        // String "!"
        seg1.extend_from_slice(&1u16.to_le_bytes());
        seg1.push(0x00);
        seg1.push(b'!');

        let (data, offsets) = build_sst_segments(&[&seg0, &seg1]);
        let strings = parse_sst_continued(&data, &offsets).unwrap();
        assert_eq!(strings, vec!["Hi", "World", "!"]);
    }

    #[test]
    fn test_sst_continue_boundary_between_strings() {
        // SST with 2 strings: "AB" and "CD".
        // "AB" ends exactly at the boundary. "CD" starts in the CONTINUE.
        // No flags byte needed because no character data is split.
        let mut seg0 = Vec::new();
        seg0.extend_from_slice(&2u32.to_le_bytes());
        seg0.extend_from_slice(&2u32.to_le_bytes());

        // String "AB" — fits entirely in segment 0
        seg0.extend_from_slice(&2u16.to_le_bytes());
        seg0.push(0x00);
        seg0.extend_from_slice(b"AB");

        // CONTINUE starts with string "CD" header (no flags byte)
        let mut seg1 = Vec::new();
        seg1.extend_from_slice(&2u16.to_le_bytes());
        seg1.push(0x00);
        seg1.extend_from_slice(b"CD");

        let (data, offsets) = build_sst_segments(&[&seg0, &seg1]);
        let strings = parse_sst_continued(&data, &offsets).unwrap();
        assert_eq!(strings, vec!["AB", "CD"]);
    }

    #[test]
    fn test_sst_continue_multiple_boundaries() {
        // SST with 1 long string split across 3 segments (2 CONTINUE records).
        // "ABCDEFGH" — 8 chars compressed.
        // Seg0: header + "ABC"
        // Seg1: flags(0x00) + "DE"
        // Seg2: flags(0x00) + "FGH"
        let mut seg0 = Vec::new();
        seg0.extend_from_slice(&1u32.to_le_bytes());
        seg0.extend_from_slice(&1u32.to_le_bytes());
        seg0.extend_from_slice(&8u16.to_le_bytes());
        seg0.push(0x00);
        seg0.extend_from_slice(b"ABC");

        let seg1: &[u8] = &[0x00, b'D', b'E'];
        let seg2: &[u8] = &[0x00, b'F', b'G', b'H'];

        let (data, offsets) = build_sst_segments(&[&seg0, seg1, seg2]);
        let strings = parse_sst_continued(&data, &offsets).unwrap();
        assert_eq!(strings, vec!["ABCDEFGH"]);
    }

    #[test]
    fn test_sst_continue_rich_text_spans_boundary() {
        // SST with 1 rich-text string "AB" with 1 formatting run.
        // The run data spans the CONTINUE boundary (no flags byte for runs).
        //
        // String header: char_count=2, flags=0x08 (rich), run_count=1
        // Char data: "AB" (2 bytes)
        // Run data: 4 bytes (char_pos u16 + font_idx u16) — split across boundary
        let mut seg0 = Vec::new();
        seg0.extend_from_slice(&1u32.to_le_bytes());
        seg0.extend_from_slice(&1u32.to_le_bytes());
        seg0.extend_from_slice(&2u16.to_le_bytes()); // 2 chars
        seg0.push(0x08); // flags: rich text
        seg0.extend_from_slice(&1u16.to_le_bytes()); // 1 run
        seg0.extend_from_slice(b"AB"); // char data
        seg0.extend_from_slice(&[0x00, 0x00]); // first 2 bytes of run (char_pos=0)

        // CONTINUE: remaining 2 bytes of run data (font_idx=1)
        let seg1: &[u8] = &[0x01, 0x00];

        let (data, offsets) = build_sst_segments(&[&seg0, seg1]);
        let strings = parse_sst_continued(&data, &offsets).unwrap();
        assert_eq!(strings, vec!["AB"]);
    }

    #[test]
    fn test_sst_no_continue_offsets() {
        // Verify parse_sst_continued works identically to parse_sst when
        // there are no CONTINUE boundaries (empty offsets).
        let mut buf = Vec::new();
        buf.extend_from_slice(&2u32.to_le_bytes());
        buf.extend_from_slice(&2u32.to_le_bytes());
        buf.extend_from_slice(&[0x01, 0x00, 0x00, b'A']);
        buf.extend_from_slice(&[0x02, 0x00, 0x00, b'B', b'C']);

        let strings = parse_sst_continued(&buf, &[]).unwrap();
        assert_eq!(strings, vec!["A", "BC"]);
    }

    #[test]
    fn test_parse_sst_entries_plain() {
        // SST with 2 plain strings: "A" and "BC"
        let mut buf = Vec::new();
        buf.extend_from_slice(&2u32.to_le_bytes()); // total
        buf.extend_from_slice(&2u32.to_le_bytes()); // unique
        buf.extend_from_slice(&[0x01, 0x00, 0x00, b'A']);
        buf.extend_from_slice(&[0x02, 0x00, 0x00, b'B', b'C']);

        let entries = parse_sst_entries(&buf, &[]).unwrap();
        assert_eq!(entries.len(), 2);
        assert!(matches!(&entries[0], SstEntry::Plain(s) if s == "A"));
        assert!(matches!(&entries[1], SstEntry::Plain(s) if s == "BC"));
    }

    #[test]
    fn test_parse_sst_entries_rich_text() {
        // SST with 1 rich-text string "AB" with 1 formatting run at char 1, font 5.
        // char_count=2, flags=0x08 (rich), run_count=1, data="AB",
        // run: char_pos=1(u16) + font_idx=5(u16)
        let mut buf = Vec::new();
        buf.extend_from_slice(&1u32.to_le_bytes()); // total
        buf.extend_from_slice(&1u32.to_le_bytes()); // unique
        buf.extend_from_slice(&2u16.to_le_bytes()); // char_count = 2
        buf.push(0x08); // flags: rich text
        buf.extend_from_slice(&1u16.to_le_bytes()); // run_count = 1
        buf.extend_from_slice(b"AB"); // char data
        buf.extend_from_slice(&1u16.to_le_bytes()); // run char_pos = 1
        buf.extend_from_slice(&5u16.to_le_bytes()); // run font_idx = 5

        let entries = parse_sst_entries(&buf, &[]).unwrap();
        assert_eq!(entries.len(), 1);
        match &entries[0] {
            SstEntry::Rich { text, runs } => {
                assert_eq!(text, "AB");
                assert_eq!(runs.len(), 1);
                assert_eq!(runs[0].char_pos, 1);
                assert_eq!(runs[0].font_index, 5);
            }
            other => panic!("Expected Rich, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_sst_entries_rich_text_multiple_runs() {
        // SST with 1 string "Hello World" with 2 runs:
        // Run 0: char_pos=0, font=1 (bold)
        // Run 1: char_pos=6, font=2 (italic)
        let text = "Hello World";
        let mut buf = Vec::new();
        buf.extend_from_slice(&1u32.to_le_bytes()); // total
        buf.extend_from_slice(&1u32.to_le_bytes()); // unique
        buf.extend_from_slice(&(text.len() as u16).to_le_bytes()); // char_count
        buf.push(0x08); // flags: rich text
        buf.extend_from_slice(&2u16.to_le_bytes()); // run_count = 2
        buf.extend_from_slice(text.as_bytes()); // char data
                                                // Run 1: char_pos=0, font_idx=1
        buf.extend_from_slice(&0u16.to_le_bytes());
        buf.extend_from_slice(&1u16.to_le_bytes());
        // Run 2: char_pos=6, font_idx=2
        buf.extend_from_slice(&6u16.to_le_bytes());
        buf.extend_from_slice(&2u16.to_le_bytes());

        let entries = parse_sst_entries(&buf, &[]).unwrap();
        assert_eq!(entries.len(), 1);
        match &entries[0] {
            SstEntry::Rich { text, runs } => {
                assert_eq!(text, "Hello World");
                assert_eq!(runs.len(), 2);
                assert_eq!(
                    runs[0],
                    FormattingRun {
                        char_pos: 0,
                        font_index: 1
                    }
                );
                assert_eq!(
                    runs[1],
                    FormattingRun {
                        char_pos: 6,
                        font_index: 2
                    }
                );
            }
            other => panic!("Expected Rich, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_sst_entries_mixed_plain_and_rich() {
        // SST with 2 strings: "Plain" (no runs) and "Bold" (1 run)
        let mut buf = Vec::new();
        buf.extend_from_slice(&2u32.to_le_bytes()); // total
        buf.extend_from_slice(&2u32.to_le_bytes()); // unique

        // String 1: plain "Plain"
        buf.extend_from_slice(&5u16.to_le_bytes());
        buf.push(0x00); // flags: no rich, no ext
        buf.extend_from_slice(b"Plain");

        // String 2: rich "Bold" with 1 run at pos 0, font 3
        buf.extend_from_slice(&4u16.to_le_bytes());
        buf.push(0x08); // flags: rich
        buf.extend_from_slice(&1u16.to_le_bytes()); // 1 run
        buf.extend_from_slice(b"Bold");
        buf.extend_from_slice(&0u16.to_le_bytes()); // char_pos = 0
        buf.extend_from_slice(&3u16.to_le_bytes()); // font_idx = 3

        let entries = parse_sst_entries(&buf, &[]).unwrap();
        assert_eq!(entries.len(), 2);
        assert!(matches!(&entries[0], SstEntry::Plain(s) if s == "Plain"));
        match &entries[1] {
            SstEntry::Rich { text, runs } => {
                assert_eq!(text, "Bold");
                assert_eq!(runs.len(), 1);
                assert_eq!(
                    runs[0],
                    FormattingRun {
                        char_pos: 0,
                        font_index: 3
                    }
                );
            }
            other => panic!("Expected Rich, got {:?}", other),
        }
    }
}
