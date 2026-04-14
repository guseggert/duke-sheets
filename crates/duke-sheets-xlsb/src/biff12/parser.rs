use crate::error::{XlsbError, XlsbResult};

#[inline]
pub fn read_u16(buf: &[u8], off: usize) -> u16 {
    u16::from_le_bytes([buf[off], buf[off + 1]])
}

#[inline]
pub fn read_u32(buf: &[u8], off: usize) -> u32 {
    u32::from_le_bytes([buf[off], buf[off + 1], buf[off + 2], buf[off + 3]])
}

#[inline]
pub fn read_i32(buf: &[u8], off: usize) -> i32 {
    i32::from_le_bytes([buf[off], buf[off + 1], buf[off + 2], buf[off + 3]])
}

#[inline]
pub fn read_f64(buf: &[u8], off: usize) -> f64 {
    f64::from_le_bytes(buf[off..off + 8].try_into().unwrap())
}

/// Decode an RK-encoded number (4 bytes).
///
/// Bit 0: divide by 100. Bit 1: integer (i30) vs float (upper 30 bits of f64).
#[inline]
pub fn decode_rk(buf: &[u8], off: usize) -> f64 {
    let rk = read_u32(buf, off);
    let div100 = (rk & 0x01) != 0;
    let is_int = (rk & 0x02) != 0;

    let value = if is_int {
        ((rk as i32) >> 2) as f64
    } else {
        let upper = (rk & 0xFFFF_FFFC) as u64;
        f64::from_bits(upper << 32)
    };

    if div100 {
        value / 100.0
    } else {
        value
    }
}

/// XLWideString: u32 char count followed by UTF-16LE data.
/// Returns (decoded_string, total_bytes_consumed).
pub fn wide_str(buf: &[u8], off: usize) -> XlsbResult<(String, usize)> {
    if off + 4 > buf.len() {
        return Err(XlsbError::Parse(
            "wide_str: buffer too short for length".into(),
        ));
    }
    let char_count = read_u32(buf, off) as usize;
    let byte_len = char_count * 2;
    let str_start = off + 4;
    let str_end = str_start + byte_len;
    if str_end > buf.len() {
        return Err(XlsbError::Parse(format!(
            "wide_str: need {} bytes but only {} available",
            byte_len,
            buf.len() - str_start
        )));
    }
    let (cow, _, had_errors) = encoding_rs::UTF_16LE.decode(&buf[str_start..str_end]);
    if had_errors {
        log::warn!("wide_str: UTF-16LE decoding had errors");
    }
    Ok((cow.into_owned(), 4 + byte_len))
}

/// 24-bit iStyleRef from cell record bytes 4..7.
#[inline]
pub fn cell_style_ref(buf: &[u8]) -> u32 {
    u32::from_le_bytes([buf[4], buf[5], buf[6], 0])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_read_u16() {
        assert_eq!(read_u16(&[0x01, 0x00], 0), 1);
        assert_eq!(read_u16(&[0xFF, 0xFF], 0), 0xFFFF);
        assert_eq!(read_u16(&[0x00, 0x00, 0x34, 0x12], 2), 0x1234);
    }

    #[test]
    fn test_read_u32() {
        assert_eq!(read_u32(&[0x01, 0x00, 0x00, 0x00], 0), 1);
        assert_eq!(read_u32(&[0xFF, 0xFF, 0xFF, 0xFF], 0), 0xFFFF_FFFF);
        assert_eq!(read_u32(&[0x00, 0x78, 0x56, 0x34, 0x12], 1), 0x12345678);
    }

    #[test]
    fn test_read_i32() {
        assert_eq!(read_i32(&[0x01, 0x00, 0x00, 0x00], 0), 1);
        assert_eq!(read_i32(&[0xFF, 0xFF, 0xFF, 0xFF], 0), -1);
        let bytes = (-42i32).to_le_bytes();
        assert_eq!(read_i32(&bytes, 0), -42);
    }

    #[test]
    fn test_read_f64() {
        let bytes = 3.14f64.to_le_bytes();
        assert_eq!(read_f64(&bytes, 0), 3.14);
        assert_eq!(read_f64(&0.0f64.to_le_bytes(), 0), 0.0);
        assert_eq!(read_f64(&1.0f64.to_le_bytes(), 0), 1.0);
        // with offset
        let mut buf = vec![0xAA, 0xBB];
        buf.extend_from_slice(&42.5f64.to_le_bytes());
        assert_eq!(read_f64(&buf, 2), 42.5);
    }

    #[test]
    fn test_decode_rk_integer() {
        // Integer 100: value = 100, bits = (100 << 2) | 0x02
        let rk = ((100i32 << 2) as u32) | 0x02;
        let buf = rk.to_le_bytes();
        assert_eq!(decode_rk(&buf, 0), 100.0);
    }

    #[test]
    fn test_decode_rk_negative_integer() {
        // Integer -7: value = -7, bits = (-7 << 2) | 0x02
        let rk = ((-7i32 << 2) as u32) | 0x02;
        let buf = rk.to_le_bytes();
        assert_eq!(decode_rk(&buf, 0), -7.0);
    }

    #[test]
    fn test_decode_rk_integer_div100() {
        // Integer 1234 / 100 = 12.34
        let rk = ((1234i32 << 2) as u32) | 0x03; // is_int + div100
        let buf = rk.to_le_bytes();
        assert!((decode_rk(&buf, 0) - 12.34).abs() < 1e-10);
    }

    #[test]
    fn test_decode_rk_float() {
        // Float: upper 30 bits of f64 1.5, with lower 2 bits = 0
        let f_bits = 1.5f64.to_bits();
        let upper = (f_bits >> 32) as u32;
        let rk = upper & 0xFFFF_FFFC; // clear bottom 2 bits, no flags
        let buf = rk.to_le_bytes();
        let decoded = decode_rk(&buf, 0);
        // The bottom 32 bits of the f64 are zeroed, so result may differ slightly
        let expected = f64::from_bits((upper as u64 & 0xFFFF_FFFC) << 32);
        assert_eq!(decoded, expected);
    }

    #[test]
    fn test_decode_rk_float_div100() {
        // Float 150.0 / 100 = 1.5
        let f_bits = 150.0f64.to_bits();
        let upper = (f_bits >> 32) as u32;
        let rk = (upper & 0xFFFF_FFFC) | 0x01; // div100, not int
        let buf = rk.to_le_bytes();
        let decoded = decode_rk(&buf, 0);
        let expected = f64::from_bits(((upper & 0xFFFF_FFFC) as u64) << 32) / 100.0;
        assert!((decoded - expected).abs() < 1e-10);
    }

    #[test]
    fn test_wide_str_ascii() {
        // "hello" = 5 chars, UTF-16LE
        let mut buf = Vec::new();
        buf.extend_from_slice(&5u32.to_le_bytes()); // char count
        for &ch in b"hello" {
            buf.push(ch);
            buf.push(0);
        }
        let (s, consumed) = wide_str(&buf, 0).unwrap();
        assert_eq!(s, "hello");
        assert_eq!(consumed, 4 + 10); // 4 bytes length + 5*2 bytes data
    }

    #[test]
    fn test_wide_str_empty() {
        let buf = 0u32.to_le_bytes();
        let (s, consumed) = wide_str(&buf, 0).unwrap();
        assert_eq!(s, "");
        assert_eq!(consumed, 4);
    }

    #[test]
    fn test_wide_str_with_offset() {
        let mut buf = vec![0xAA, 0xBB]; // padding
        buf.extend_from_slice(&2u32.to_le_bytes());
        // "AB" in UTF-16LE
        buf.extend_from_slice(&[0x41, 0x00, 0x42, 0x00]);
        let (s, consumed) = wide_str(&buf, 2).unwrap();
        assert_eq!(s, "AB");
        assert_eq!(consumed, 4 + 4);
    }

    #[test]
    fn test_wide_str_accented() {
        // "café" = 4 chars: c(0x63) a(0x61) f(0x66) é(0xE9)
        let mut buf = Vec::new();
        buf.extend_from_slice(&4u32.to_le_bytes());
        for &ch in &[0x0063u16, 0x0061, 0x0066, 0x00E9] {
            buf.extend_from_slice(&ch.to_le_bytes());
        }
        let (s, consumed) = wide_str(&buf, 0).unwrap();
        assert_eq!(s, "café");
        assert_eq!(consumed, 4 + 8);
    }

    #[test]
    fn test_wide_str_emoji() {
        // "😀" = U+1F600, encoded as UTF-16 surrogate pair: D83D DE00
        let mut buf = Vec::new();
        buf.extend_from_slice(&2u32.to_le_bytes()); // 2 UTF-16 code units
        buf.extend_from_slice(&0xD83Du16.to_le_bytes());
        buf.extend_from_slice(&0xDE00u16.to_le_bytes());
        let (s, consumed) = wide_str(&buf, 0).unwrap();
        assert_eq!(s, "😀");
        assert_eq!(consumed, 4 + 4);
    }

    #[test]
    fn test_wide_str_buffer_too_short_for_length() {
        let buf = [0x00, 0x00]; // only 2 bytes, need 4 for length
        assert!(wide_str(&buf, 0).is_err());
    }

    #[test]
    fn test_wide_str_buffer_too_short_for_data() {
        let mut buf = Vec::new();
        buf.extend_from_slice(&10u32.to_le_bytes()); // claims 10 chars = 20 bytes
        buf.extend_from_slice(&[0x41, 0x00]); // only 2 bytes of data
        assert!(wide_str(&buf, 0).is_err());
    }

    #[test]
    fn test_cell_style_ref() {
        let buf = [0x00, 0x00, 0x00, 0x00, 0x42, 0x01, 0x00, 0x00];
        assert_eq!(cell_style_ref(&buf), 0x00000142);
    }

    #[test]
    fn test_cell_style_ref_max_24bit() {
        let buf = [0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0x00];
        assert_eq!(cell_style_ref(&buf), 0x00FFFFFF);
    }
}
