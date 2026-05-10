pub(crate) mod compiler;
pub mod parser;
pub mod ptg;
pub mod records;
pub mod token_parser;

use std::io::{BufReader, Read};

use crate::error::{XlsbError, XlsbResult};

pub struct RecordIter<R: Read> {
    reader: BufReader<R>,
    scratch: [u8; 1],
}

impl<R: Read> RecordIter<R> {
    pub fn new(reader: R) -> Self {
        Self {
            reader: BufReader::new(reader),
            scratch: [0u8; 1],
        }
    }

    fn read_u8(&mut self) -> std::io::Result<u8> {
        self.reader.read_exact(&mut self.scratch)?;
        Ok(self.scratch[0])
    }

    /// Read a variable-length record type (1 or 2 bytes).
    ///
    /// Bit 7 of the first byte is a continuation flag:
    /// - Clear: type = byte (7-bit, 0..127)
    /// - Set: type = (b1 & 0x7F) | ((b2 & 0x7F) << 7) (14-bit)
    pub fn read_type(&mut self) -> std::io::Result<u16> {
        let b = self.read_u8()?;
        if (b & 0x80) == 0 {
            Ok(b as u16)
        } else {
            let b2 = self.read_u8()?;
            Ok((b & 0x7F) as u16 | (((b2 & 0x7F) as u16) << 7))
        }
    }

    /// Read a variable-length record size (1..4 bytes) and fill `buf` with the payload.
    ///
    /// Returns the payload length.
    pub fn fill_buffer(&mut self, buf: &mut Vec<u8>) -> std::io::Result<usize> {
        let mut b = self.read_u8()?;
        let mut len = (b & 0x7F) as usize;
        for i in 1..4 {
            if (b & 0x80) == 0 {
                break;
            }
            b = self.read_u8()?;
            len |= ((b & 0x7F) as usize) << (7 * i);
        }
        if buf.len() < len {
            buf.resize(len, 0);
        }
        if len > 0 {
            self.reader.read_exact(&mut buf[..len])?;
        }
        Ok(len)
    }

    pub fn next_record(&mut self, buf: &mut Vec<u8>) -> XlsbResult<(u16, usize)> {
        let typ = self.read_type().map_err(|e| {
            if e.kind() == std::io::ErrorKind::UnexpectedEof {
                XlsbError::Parse("unexpected end of BIFF12 stream".into())
            } else {
                XlsbError::Io(e)
            }
        })?;
        let len = self.fill_buffer(buf).map_err(XlsbError::Io)?;
        Ok((typ, len))
    }

    pub fn skip_to(
        &mut self,
        target_type: u16,
        skip_blocks: &[(u16, Option<u16>)],
        buf: &mut Vec<u8>,
    ) -> XlsbResult<usize> {
        loop {
            let (typ, len) = self.next_record(buf)?;
            if typ == target_type {
                return Ok(len);
            }
            if let Some(&(_, Some(end_type))) = skip_blocks.iter().find(|(start, _)| *start == typ)
            {
                loop {
                    let (inner_typ, _) = self.next_record(buf)?;
                    if inner_typ == end_type {
                        break;
                    }
                }
            }
        }
    }
}

pub(crate) fn encode_type(typ: u16, out: &mut Vec<u8>) {
    if typ < 128 {
        out.push(typ as u8);
    } else {
        out.push((typ & 0x7F) as u8 | 0x80);
        out.push(((typ >> 7) & 0x7F) as u8);
    }
}

pub(crate) fn encode_len(mut len: usize, out: &mut Vec<u8>) {
    loop {
        let mut b = (len & 0x7F) as u8;
        len >>= 7;
        if len > 0 {
            b |= 0x80;
        }
        out.push(b);
        if len == 0 {
            break;
        }
    }
}

#[cfg(test)]
pub(crate) fn build_record(typ: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    encode_type(typ, &mut out);
    encode_len(payload.len(), &mut out);
    out.extend_from_slice(payload);
    out
}

/// Encode a string as XLWideString: u32 char_count + UTF-16LE bytes.
pub(crate) fn encode_wide_str(s: &str) -> Vec<u8> {
    let utf16: Vec<u16> = s.encode_utf16().collect();
    let mut out = Vec::with_capacity(4 + utf16.len() * 2);
    out.extend_from_slice(&(utf16.len() as u32).to_le_bytes());
    for code_unit in &utf16 {
        out.extend_from_slice(&code_unit.to_le_bytes());
    }
    out
}

/// Encode an XLNullableWideString: cchCharacters=0xFFFFFFFF for NULL,
/// otherwise cchCharacters + UTF-16LE bytes. Empty `Some("")` is
/// treated as NULL because Excel emits the null marker for absent
/// optional strings.
pub(crate) fn encode_nullable_wide_str(s: Option<&str>) -> Vec<u8> {
    match s {
        None => 0xFFFFFFFFu32.to_le_bytes().to_vec(),
        Some(text) if text.is_empty() => 0xFFFFFFFFu32.to_le_bytes().to_vec(),
        Some(text) => encode_wide_str(text),
    }
}

pub(crate) struct RecordWriter<W: std::io::Write> {
    writer: W,
}

impl<W: std::io::Write> RecordWriter<W> {
    pub fn new(writer: W) -> Self {
        Self { writer }
    }

    pub fn write_record(&mut self, record_type: u16, payload: &[u8]) -> std::io::Result<()> {
        let mut header = Vec::with_capacity(6);
        encode_type(record_type, &mut header);
        encode_len(payload.len(), &mut header);
        self.writer.write_all(&header)?;
        self.writer.write_all(payload)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_single_byte_type() {
        // Type 5, 3 bytes payload [0xAA, 0xBB, 0xCC]
        let data = build_record(5, &[0xAA, 0xBB, 0xCC]);
        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();
        let (typ, len) = iter.next_record(&mut buf).unwrap();
        assert_eq!(typ, 5);
        assert_eq!(len, 3);
        assert_eq!(&buf[..len], &[0xAA, 0xBB, 0xCC]);
    }

    #[test]
    fn test_two_byte_type() {
        // Type 0x009C (BRT_BUNDLE_SH = 156), needs 2 bytes
        let data = build_record(0x009C, &[0x01, 0x02]);
        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();
        let (typ, len) = iter.next_record(&mut buf).unwrap();
        assert_eq!(typ, 0x009C);
        assert_eq!(len, 2);
        assert_eq!(&buf[..len], &[0x01, 0x02]);
    }

    #[test]
    fn test_empty_payload() {
        let data = build_record(0x0081, &[]);
        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();
        let (typ, len) = iter.next_record(&mut buf).unwrap();
        assert_eq!(typ, 0x0081);
        assert_eq!(len, 0);
    }

    #[test]
    fn test_multi_byte_length() {
        // 200 bytes payload - length needs 2 bytes (200 > 127)
        let payload: Vec<u8> = (0..200).map(|i| (i & 0xFF) as u8).collect();
        let data = build_record(10, &payload);
        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();
        let (typ, len) = iter.next_record(&mut buf).unwrap();
        assert_eq!(typ, 10);
        assert_eq!(len, 200);
        assert_eq!(&buf[..len], &payload[..]);
    }

    #[test]
    fn test_multiple_records() {
        let mut data = Vec::new();
        data.extend_from_slice(&build_record(1, &[0x10]));
        data.extend_from_slice(&build_record(2, &[0x20, 0x21]));
        data.extend_from_slice(&build_record(3, &[0x30, 0x31, 0x32]));

        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();

        let (typ, len) = iter.next_record(&mut buf).unwrap();
        assert_eq!(typ, 1);
        assert_eq!(&buf[..len], &[0x10]);

        let (typ, len) = iter.next_record(&mut buf).unwrap();
        assert_eq!(typ, 2);
        assert_eq!(&buf[..len], &[0x20, 0x21]);

        let (typ, len) = iter.next_record(&mut buf).unwrap();
        assert_eq!(typ, 3);
        assert_eq!(&buf[..len], &[0x30, 0x31, 0x32]);
    }

    #[test]
    fn test_skip_to_simple() {
        let mut data = Vec::new();
        data.extend_from_slice(&build_record(1, &[0xAA]));
        data.extend_from_slice(&build_record(2, &[0xBB]));
        data.extend_from_slice(&build_record(3, &[0xCC])); // target

        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();
        let len = iter.skip_to(3, &[], &mut buf).unwrap();
        assert_eq!(len, 1);
        assert_eq!(buf[0], 0xCC);
    }

    #[test]
    fn test_skip_to_with_block_skip() {
        // Records: [begin_block(10), inner(99), end_block(11), target(5)]
        let mut data = Vec::new();
        data.extend_from_slice(&build_record(10, &[0x01])); // begin block
        data.extend_from_slice(&build_record(99, &[0x02])); // inner (should be skipped)
        data.extend_from_slice(&build_record(11, &[])); // end block
        data.extend_from_slice(&build_record(5, &[0xFF])); // target

        let skip_blocks = &[(10u16, Some(11u16))];
        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();
        let len = iter.skip_to(5, skip_blocks, &mut buf).unwrap();
        assert_eq!(len, 1);
        assert_eq!(buf[0], 0xFF);
    }

    #[test]
    fn test_skip_to_not_found() {
        let data = build_record(1, &[0xAA]);
        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();
        // Looking for type 99 which doesn't exist - should hit EOF
        let result = iter.skip_to(99, &[], &mut buf);
        assert!(result.is_err());
    }

    #[test]
    fn test_eof_error() {
        let data: Vec<u8> = vec![];
        let mut iter = RecordIter::new(Cursor::new(data));
        let mut buf = Vec::new();
        let result = iter.next_record(&mut buf);
        assert!(result.is_err());
    }

    #[test]
    fn test_type_encoding_roundtrip() {
        for typ in [0u16, 1, 50, 127, 128, 255, 0x009C, 0x01EE, 0x0269, 0x3FFF] {
            let mut encoded = Vec::new();
            encode_type(typ, &mut encoded);
            encode_len(0, &mut encoded); // zero-length payload
            let mut iter = RecordIter::new(Cursor::new(encoded));
            let mut buf = Vec::new();
            let (decoded_typ, len) = iter.next_record(&mut buf).unwrap();
            assert_eq!(decoded_typ, typ, "roundtrip failed for type {typ:#06X}");
            assert_eq!(len, 0);
        }
    }

    #[test]
    fn test_length_encoding_roundtrip() {
        for payload_len in [0usize, 1, 127, 128, 255, 16383, 16384, 100_000] {
            let payload: Vec<u8> = (0..payload_len).map(|i| (i & 0xFF) as u8).collect();
            let data = build_record(1, &payload);
            let mut iter = RecordIter::new(Cursor::new(data));
            let mut buf = Vec::new();
            let (_, len) = iter.next_record(&mut buf).unwrap();
            assert_eq!(
                len, payload_len,
                "roundtrip failed for length {payload_len}"
            );
            assert_eq!(&buf[..len], &payload[..]);
        }
    }
}
