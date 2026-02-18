//! Low-level binary parsing helpers for BIFF8 records.
//!
//! All multi-byte integers in BIFF8 are little-endian.

use crate::error::{XlsError, XlsResult};

/// Read a `u8` from a byte slice at `offset`, advancing `offset`.
#[inline]
pub fn read_u8(data: &[u8], offset: &mut usize) -> XlsResult<u8> {
    if *offset >= data.len() {
        return Err(XlsError::Parse(format!(
            "unexpected end of data at offset {}, need 1 byte",
            *offset
        )));
    }
    let v = data[*offset];
    *offset += 1;
    Ok(v)
}

/// Read a `u16` (little-endian) from a byte slice at `offset`, advancing `offset`.
#[inline]
pub fn read_u16(data: &[u8], offset: &mut usize) -> XlsResult<u16> {
    if *offset + 2 > data.len() {
        return Err(XlsError::Parse(format!(
            "unexpected end of data at offset {}, need 2 bytes",
            *offset
        )));
    }
    let v = u16::from_le_bytes([data[*offset], data[*offset + 1]]);
    *offset += 2;
    Ok(v)
}

/// Read a `u32` (little-endian) from a byte slice at `offset`, advancing `offset`.
#[inline]
pub fn read_u32(data: &[u8], offset: &mut usize) -> XlsResult<u32> {
    if *offset + 4 > data.len() {
        return Err(XlsError::Parse(format!(
            "unexpected end of data at offset {}, need 4 bytes",
            *offset
        )));
    }
    let v = u32::from_le_bytes([
        data[*offset],
        data[*offset + 1],
        data[*offset + 2],
        data[*offset + 3],
    ]);
    *offset += 4;
    Ok(v)
}

/// Read an `i16` (little-endian).
#[inline]
pub fn read_i16(data: &[u8], offset: &mut usize) -> XlsResult<i16> {
    read_u16(data, offset).map(|v| v as i16)
}

/// Read an `i32` (little-endian).
#[inline]
pub fn read_i32(data: &[u8], offset: &mut usize) -> XlsResult<i32> {
    read_u32(data, offset).map(|v| v as i32)
}

/// Read an `f64` (IEEE 754 double, little-endian) from a byte slice.
#[inline]
pub fn read_f64(data: &[u8], offset: &mut usize) -> XlsResult<f64> {
    if *offset + 8 > data.len() {
        return Err(XlsError::Parse(format!(
            "unexpected end of data at offset {}, need 8 bytes",
            *offset
        )));
    }
    let bytes: [u8; 8] = data[*offset..*offset + 8].try_into().unwrap();
    *offset += 8;
    Ok(f64::from_le_bytes(bytes))
}

/// Decode an RK-encoded number.
///
/// RK encoding (4 bytes):
/// - Bit 0: if 1, the decoded number should be divided by 100
/// - Bit 1: if 1, value is an integer (bits 2..31 as signed 30-bit int)
///           if 0, value is an IEEE 754 double (bits 2..31 are the upper 30 bits,
///           lower 34 bits of the double are zero)
#[inline]
pub fn decode_rk(rk: u32) -> f64 {
    let div100 = (rk & 0x01) != 0;
    let is_integer = (rk & 0x02) != 0;

    let value = if is_integer {
        // Signed 30-bit integer in bits 2..31
        ((rk as i32) >> 2) as f64
    } else {
        // IEEE 754 double with upper 30 bits from rk (bits 2..31)
        // and lower 34 bits set to zero.
        let upper = (rk & 0xFFFF_FFFC) as u64;
        let bits = upper << 32;
        f64::from_bits(bits)
    };

    if div100 {
        value / 100.0
    } else {
        value
    }
}

/// Read an RK value from 4 bytes at `offset`.
#[inline]
pub fn read_rk(data: &[u8], offset: &mut usize) -> XlsResult<f64> {
    let raw = read_u32(data, offset)?;
    Ok(decode_rk(raw))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_decode_rk_integer() {
        // Integer 42: bits 2..31 = 42, bit 1 = 1 (integer), bit 0 = 0 (no /100)
        // rk = (42 << 2) | 2 = 0xAA
        let rk = (42u32 << 2) | 0x02;
        assert_eq!(decode_rk(rk), 42.0);
    }

    #[test]
    fn test_decode_rk_integer_negative() {
        // Integer -5: shift left 2, set bit 1 for integer
        let rk = ((-5i32 << 2) as u32) | 0x02;
        assert_eq!(decode_rk(rk), -5.0);
    }

    #[test]
    fn test_decode_rk_integer_div100() {
        // Integer 4200 / 100 = 42.0
        // bit 1 = 1 (integer), bit 0 = 1 (/100)
        let rk = (4200u32 << 2) | 0x03;
        assert_eq!(decode_rk(rk), 42.0);
    }

    #[test]
    fn test_decode_rk_float() {
        // Float: encode 42.0 as RK
        // Upper 30 bits of the double go into bits 2..31, bits 0-1 = 0
        let bits = 42.0_f64.to_bits();
        let upper = ((bits >> 32) as u32) & 0xFFFF_FFFC;
        let rk = upper; // bit 0 = 0 (no /100), bit 1 = 0 (float)
        assert_eq!(decode_rk(rk), 42.0);
    }

    #[test]
    fn test_decode_rk_real_values() {
        // Values observed from LibreOffice MULRK output:
        // 42.0  -> 0x000000AA (integer, no div100)
        assert_eq!(decode_rk(0x000000AA), 42.0);
        // 3.14  -> 0x000004EB (integer 314, div100)
        assert!((decode_rk(0x000004EB) - 3.14).abs() < f64::EPSILON);
        // -100  -> 0xFFFFFE72 (integer -100, no div100)
        assert_eq!(decode_rk(0xFFFFFE72), -100.0);
        // 0     -> 0x00000002 (integer 0, no div100)
        assert_eq!(decode_rk(0x00000002), 0.0);
    }

    #[test]
    fn test_read_u16() {
        let data = [0x34, 0x12];
        let mut off = 0;
        assert_eq!(read_u16(&data, &mut off).unwrap(), 0x1234);
        assert_eq!(off, 2);
    }

    #[test]
    fn test_read_f64() {
        let val = 3.14_f64;
        let bytes = val.to_le_bytes();
        let mut off = 0;
        let result = read_f64(&bytes, &mut off).unwrap();
        assert!((result - val).abs() < f64::EPSILON);
    }
}
