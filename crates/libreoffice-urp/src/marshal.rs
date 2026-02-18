//! Binary serialization/deserialization of UNO types for the URP wire format.
//!
//! All multi-byte integers are big-endian. Strings are UTF-8 with compressed
//! length prefix. The "compressed number" format uses 1 byte for values < 0xFF,
//! or 5 bytes (0xFF prefix + u32 BE) for larger values.

use bytes::{Buf, BufMut, Bytes, BytesMut};

use crate::error::{Result, UrpError};
use crate::protocol::LruCache;
use crate::types::{Any, Type, TypeClass, UnoException, UnoValue};

// ============================================================================
// Compressed number encoding
// ============================================================================

/// Write a compressed u32. Values < 255 use 1 byte; >= 255 use 5 bytes.
pub fn write_compressed(buf: &mut BytesMut, value: u32) {
    if value < 0xFF {
        buf.put_u8(value as u8);
    } else {
        buf.put_u8(0xFF);
        buf.put_u32(value);
    }
}

/// Read a compressed u32.
pub fn read_compressed(buf: &mut Bytes) -> Result<u32> {
    if buf.remaining() < 1 {
        return Err(UrpError::Marshal(
            "unexpected end of data reading compressed number".into(),
        ));
    }
    let first = buf.get_u8();
    if first < 0xFF {
        Ok(first as u32)
    } else {
        if buf.remaining() < 4 {
            return Err(UrpError::Marshal(
                "unexpected end of data reading compressed number (extended)".into(),
            ));
        }
        Ok(buf.get_u32())
    }
}

// ============================================================================
// String encoding
// ============================================================================

/// Write a UTF-8 string with compressed length prefix.
pub fn write_string(buf: &mut BytesMut, s: &str) {
    let bytes = s.as_bytes();
    write_compressed(buf, bytes.len() as u32);
    buf.put_slice(bytes);
}

/// Read a UTF-8 string with compressed length prefix.
pub fn read_string(buf: &mut Bytes) -> Result<String> {
    let len = read_compressed(buf)? as usize;
    if buf.remaining() < len {
        return Err(UrpError::Marshal(format!(
            "unexpected end of data reading string of length {len}, only {} bytes remaining",
            buf.remaining()
        )));
    }
    let bytes = buf.copy_to_bytes(len);
    String::from_utf8(bytes.to_vec())
        .map_err(|e| UrpError::Marshal(format!("invalid UTF-8 in string: {e}")))
}

// ============================================================================
// Type encoding
// ============================================================================

/// Write a UNO Type to the buffer.
///
/// Simple types (Void..Any) are a single byte with cache_flag=0.
/// Complex types (Enum, Struct, Exception, Sequence, Interface) include
/// a cache index and optionally the type name.
pub fn write_type(buf: &mut BytesMut, ty: &Type, cache_index: u16, new: bool) {
    let tc = ty.class as u8;
    if ty.class.is_simple() {
        // Simple types: just the type class byte, no cache
        buf.put_u8(tc);
    } else if new {
        // Complex type, providing new value: set cache flag (bit 7)
        buf.put_u8(tc | 0x80);
        buf.put_u16(cache_index);
        write_string(buf, &ty.name);
    } else {
        // Complex type, reading from cache: clear cache flag
        buf.put_u8(tc);
        buf.put_u16(cache_index);
    }
}

/// Read a UNO Type from the buffer.
///
/// Returns `(Type, cache_index, is_new)`. For simple types, cache_index is 0xFFFF
/// and is_new is false.
pub fn read_type(buf: &mut Bytes) -> Result<(Type, u16, bool)> {
    if buf.remaining() < 1 {
        return Err(UrpError::Marshal(
            "unexpected end of data reading type".into(),
        ));
    }
    let byte = buf.get_u8();
    let cache_flag = byte & 0x80 != 0;
    let tc_byte = byte & 0x7F;

    let tc = TypeClass::from_byte(tc_byte).ok_or_else(|| UrpError::UnknownTypeClass(tc_byte))?;

    if tc.is_simple() {
        Ok((
            Type {
                class: tc,
                name: String::new(),
            },
            0xFFFF,
            false,
        ))
    } else {
        if buf.remaining() < 2 {
            return Err(UrpError::Marshal(
                "unexpected end of data reading type cache index".into(),
            ));
        }
        let cache_index = buf.get_u16();
        if cache_flag {
            // New type name provided
            let name = read_string(buf)?;
            Ok((Type { class: tc, name }, cache_index, true))
        } else {
            // Must read from cache — caller is responsible
            Ok((
                Type {
                    class: tc,
                    name: String::new(),
                },
                cache_index,
                false,
            ))
        }
    }
}

// ============================================================================
// Value encoding
// ============================================================================

/// Write a UNO value, given its expected type.
/// Uses no OID caching for interface references (always sends full OID + 0xFFFF).
pub fn write_value(buf: &mut BytesMut, value: &UnoValue, ty: &Type) {
    write_value_cached(buf, value, ty, None);
}

/// Write a UNO value with an optional OID cache for interface references.
/// The cache is shared with the header-level OID cache (matching LibreOffice's behavior).
pub fn write_value_cached(
    buf: &mut BytesMut,
    value: &UnoValue,
    ty: &Type,
    oid_cache: Option<&mut LruCache<String>>,
) {
    match value {
        UnoValue::Void => {}
        UnoValue::Bool(b) => buf.put_u8(if *b { 1 } else { 0 }),
        UnoValue::Byte(b) => buf.put_u8(*b),
        UnoValue::Short(n) => buf.put_i16(*n),
        UnoValue::UnsignedShort(n) => buf.put_u16(*n),
        UnoValue::Long(n) => buf.put_i32(*n),
        UnoValue::UnsignedLong(n) => buf.put_u32(*n),
        UnoValue::Hyper(n) => buf.put_i64(*n),
        UnoValue::UnsignedHyper(n) => buf.put_u64(*n),
        UnoValue::Float(f) => buf.put_f32(*f),
        UnoValue::Double(d) => buf.put_f64(*d),
        UnoValue::Char(c) => buf.put_u16(*c),
        UnoValue::String(s) => write_string(buf, s),
        UnoValue::Type(t) => {
            // When writing a Type as a value, use cache index 0xFFFF (no caching for now)
            write_type(buf, t, 0xFFFF, true);
        }
        UnoValue::Any(a) => {
            write_type(buf, &a.type_desc, 0xFFFF, true);
            if a.type_desc.class != TypeClass::Void {
                write_value_cached(buf, &a.value, &a.type_desc, oid_cache);
            }
        }
        UnoValue::Enum(n) => buf.put_i32(*n),
        UnoValue::Struct(members) => {
            // Struct fields are written with their concrete types (no type tags).
            // We infer the wire type from each member's UnoValue variant.
            // Higher-level code constructs structs with correctly-typed members
            // (e.g., CellAddress → [Short, Long, Long]).
            for member in members {
                let member_type = member.infer_type();
                write_value_cached(buf, member, &member_type, None);
            }
        }
        UnoValue::Exception(exc) => {
            // Exception = Message (string) + Context (interface, null)
            write_string(buf, &exc.message);
            // Null XInterface for Context — always uncached
            write_string(buf, "");
            buf.put_u16(0xFFFF);
        }
        UnoValue::Sequence(items) => {
            write_compressed(buf, items.len() as u32);
            // Determine element type from the Type's name
            // For []byte, items are Byte values
            if ty.name == "[]byte" {
                for item in items {
                    if let UnoValue::Byte(b) = item {
                        buf.put_u8(*b);
                    }
                }
            } else {
                let elem_type_name = ty.name.strip_prefix("[]").unwrap_or("");
                let elem_type = Type {
                    class: guess_type_class(elem_type_name),
                    name: elem_type_name.to_string(),
                };
                for item in items {
                    write_value(buf, item, &elem_type);
                }
            }
        }
        UnoValue::Interface(oid) => {
            // Interface reference: OID string + cache index
            // Uses the shared OID cache (same as message header) when available.
            if let Some(cache) = oid_cache {
                if oid.is_empty() {
                    // Null interface — no caching
                    write_string(buf, "");
                    buf.put_u16(0xFFFF);
                } else {
                    let (cache_index, is_new) = cache.insert_or_get(oid.clone());
                    if is_new {
                        // New OID: write full string + cache index
                        write_string(buf, oid);
                        buf.put_u16(cache_index);
                    } else {
                        // Cached: write empty string + cache index
                        write_string(buf, "");
                        buf.put_u16(cache_index);
                    }
                }
            } else {
                // No cache available — always write full OID
                write_string(buf, oid);
                buf.put_u16(0xFFFF);
            }
        }
    }
}

/// Read a UNO value given its expected type.
pub fn read_value(buf: &mut Bytes, ty: &Type) -> Result<UnoValue> {
    read_value_cached(buf, ty, &mut [const { None }; 256])
}

/// Read a UNO value with an OID cache for resolving interface references.
/// The cache is shared between header-level OID and body-level interface references
/// (matching LibreOffice's behavior).
pub fn read_value_cached(
    buf: &mut Bytes,
    ty: &Type,
    oid_cache: &mut [Option<String>; 256],
) -> Result<UnoValue> {
    match ty.class {
        TypeClass::Void => Ok(UnoValue::Void),
        TypeClass::Boolean => {
            ensure_remaining(buf, 1, "boolean")?;
            Ok(UnoValue::Bool(buf.get_u8() != 0))
        }
        TypeClass::Byte => {
            ensure_remaining(buf, 1, "byte")?;
            Ok(UnoValue::Byte(buf.get_u8()))
        }
        TypeClass::Short => {
            ensure_remaining(buf, 2, "short")?;
            Ok(UnoValue::Short(buf.get_i16()))
        }
        TypeClass::UnsignedShort => {
            ensure_remaining(buf, 2, "unsigned short")?;
            Ok(UnoValue::UnsignedShort(buf.get_u16()))
        }
        TypeClass::Long => {
            ensure_remaining(buf, 4, "long")?;
            Ok(UnoValue::Long(buf.get_i32()))
        }
        TypeClass::UnsignedLong => {
            ensure_remaining(buf, 4, "unsigned long")?;
            Ok(UnoValue::UnsignedLong(buf.get_u32()))
        }
        TypeClass::Hyper => {
            ensure_remaining(buf, 8, "hyper")?;
            Ok(UnoValue::Hyper(buf.get_i64()))
        }
        TypeClass::UnsignedHyper => {
            ensure_remaining(buf, 8, "unsigned hyper")?;
            Ok(UnoValue::UnsignedHyper(buf.get_u64()))
        }
        TypeClass::Float => {
            ensure_remaining(buf, 4, "float")?;
            Ok(UnoValue::Float(buf.get_f32()))
        }
        TypeClass::Double => {
            ensure_remaining(buf, 8, "double")?;
            Ok(UnoValue::Double(buf.get_f64()))
        }
        TypeClass::Char => {
            ensure_remaining(buf, 2, "char")?;
            Ok(UnoValue::Char(buf.get_u16()))
        }
        TypeClass::String => {
            let s = read_string(buf)?;
            Ok(UnoValue::String(s))
        }
        TypeClass::Type => {
            let (t, _cache_index, _is_new) = read_type(buf)?;
            Ok(UnoValue::Type(t))
        }
        TypeClass::Any => {
            let (inner_type, _cache_index, _is_new) = read_type(buf)?;
            if inner_type.class == TypeClass::Void {
                Ok(UnoValue::Any(Box::new(Any {
                    type_desc: inner_type,
                    value: UnoValue::Void,
                })))
            } else {
                let value = read_value_cached(buf, &inner_type, oid_cache)?;
                Ok(UnoValue::Any(Box::new(Any {
                    type_desc: inner_type,
                    value,
                })))
            }
        }
        TypeClass::Enum => {
            ensure_remaining(buf, 4, "enum")?;
            Ok(UnoValue::Enum(buf.get_i32()))
        }
        TypeClass::Struct => {
            // For structs, we need to know the member types.
            // Handle known structs here; unknown structs are an error.
            read_known_struct(buf, &ty.name, oid_cache)
        }
        TypeClass::Exception => {
            // Exception base: Message (string) + Context (XInterface)
            let message = read_string(buf)?;
            // Read the Context interface reference (usually null) using OID cache
            let context_oid_str = read_string(buf)?;
            if buf.remaining() >= 2 {
                let cache_index = buf.get_u16();
                // Resolve via OID cache (same logic as Interface arm)
                if !context_oid_str.is_empty() && cache_index != 0xFFFF {
                    oid_cache[cache_index as usize] = Some(context_oid_str);
                }
            }
            Ok(UnoValue::Exception(UnoException {
                type_name: ty.name.clone(),
                message,
            }))
        }
        TypeClass::Sequence => {
            let count = read_compressed(buf)? as usize;
            let elem_type_name = ty.name.strip_prefix("[]").unwrap_or("");
            let elem_type = Type {
                class: guess_type_class(elem_type_name),
                name: elem_type_name.to_string(),
            };

            if elem_type_name == "byte" {
                // Optimized: raw bytes
                ensure_remaining(buf, count, "byte sequence")?;
                let items: Vec<UnoValue> = buf
                    .copy_to_bytes(count)
                    .iter()
                    .map(|b| UnoValue::Byte(*b))
                    .collect();
                Ok(UnoValue::Sequence(items))
            } else {
                let mut items = Vec::with_capacity(count);
                for _ in 0..count {
                    items.push(read_value_cached(buf, &elem_type, oid_cache)?);
                }
                Ok(UnoValue::Sequence(items))
            }
        }
        TypeClass::Interface => {
            // Interface reference: OID string + cache index
            // Uses the same OID cache as the message header (shared ReaderState.oid_cache)
            let oid_str = read_string(buf)?;
            ensure_remaining(buf, 2, "interface cache index")?;
            let cache_index = buf.get_u16();

            let resolved = if oid_str.is_empty() && cache_index != 0xFFFF {
                // Empty string means read from cache
                oid_cache[cache_index as usize].clone().unwrap_or_default()
            } else {
                // New OID — populate cache if index is valid
                if cache_index != 0xFFFF && !oid_str.is_empty() {
                    oid_cache[cache_index as usize] = Some(oid_str.clone());
                }
                oid_str
            };
            Ok(UnoValue::Interface(resolved))
        }
    }
}

// ============================================================================
// Helpers
// ============================================================================

fn ensure_remaining(buf: &Bytes, needed: usize, what: &str) -> Result<()> {
    if buf.remaining() < needed {
        Err(UrpError::Marshal(format!(
            "unexpected end of data reading {what}: need {needed} bytes, have {}",
            buf.remaining()
        )))
    } else {
        Ok(())
    }
}

/// Guess the TypeClass from a type name string.
pub fn guess_type_class(name: &str) -> TypeClass {
    match name {
        "" | "void" => TypeClass::Void,
        "boolean" => TypeClass::Boolean,
        "byte" => TypeClass::Byte,
        "short" => TypeClass::Short,
        "unsigned short" => TypeClass::UnsignedShort,
        "long" => TypeClass::Long,
        "unsigned long" => TypeClass::UnsignedLong,
        "hyper" => TypeClass::Hyper,
        "unsigned hyper" => TypeClass::UnsignedHyper,
        "float" => TypeClass::Float,
        "double" => TypeClass::Double,
        "string" => TypeClass::String,
        "type" => TypeClass::Type,
        "any" => TypeClass::Any,
        n if n.starts_with("[]") => TypeClass::Sequence,
        // Heuristic: interface names typically contain ".X" (e.g., ...XInterface)
        n if n.contains(".X") => TypeClass::Interface,
        // Structs/exceptions: assume struct if has dots but no ".X"
        n if n.contains('.') => TypeClass::Struct,
        _ => TypeClass::Interface, // Default guess
    }
}

/// Read a known struct type by name.
fn read_known_struct(
    buf: &mut Bytes,
    name: &str,
    oid_cache: &mut [Option<String>; 256],
) -> Result<UnoValue> {
    match name {
        "com.sun.star.beans.PropertyValue" => {
            // PropertyValue: Name (string), Handle (long), Value (any), State (enum)
            let prop_name = read_string(buf)?;
            let _handle = if buf.remaining() >= 4 {
                buf.get_i32()
            } else {
                0
            };
            let value = read_value_cached(buf, &Type::any(), oid_cache)?;
            let _state = if buf.remaining() >= 4 {
                buf.get_i32()
            } else {
                0
            };
            // Return as a struct with the relevant fields
            Ok(UnoValue::Struct(vec![
                UnoValue::String(prop_name),
                UnoValue::Long(_handle),
                value,
                UnoValue::Enum(_state),
            ]))
        }
        "com.sun.star.bridge.ProtocolProperty" => {
            // ProtocolProperty: Name (string), Value (any)
            let prop_name = read_string(buf)?;
            let value = read_value_cached(buf, &Type::any(), oid_cache)?;
            Ok(UnoValue::Struct(vec![UnoValue::String(prop_name), value]))
        }
        _ => Err(UrpError::Marshal(format!("unknown struct type: {name}"))),
    }
}

/// Write a PropertyValue struct.
pub fn write_property_value(buf: &mut BytesMut, name: &str, value: &UnoValue, value_type: &Type) {
    // Name
    write_string(buf, name);
    // Handle
    buf.put_i32(0);
    // Value as Any
    write_type(buf, value_type, 0xFFFF, true);
    write_value(buf, value, value_type);
    // State = DIRECT_VALUE (0)
    buf.put_i32(0);
}

/// Write a ProtocolProperty struct.
pub fn write_protocol_property(
    buf: &mut BytesMut,
    name: &str,
    value: &UnoValue,
    value_type: &Type,
) {
    // Name
    write_string(buf, name);
    // Value as Any
    write_type(buf, value_type, 0xFFFF, true);
    write_value(buf, value, value_type);
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compressed_small() {
        let mut buf = BytesMut::new();
        write_compressed(&mut buf, 42);
        assert_eq!(buf.as_ref(), &[42]);

        let mut bytes = buf.freeze();
        assert_eq!(read_compressed(&mut bytes).unwrap(), 42);
    }

    #[test]
    fn test_compressed_zero() {
        let mut buf = BytesMut::new();
        write_compressed(&mut buf, 0);
        assert_eq!(buf.as_ref(), &[0]);

        let mut bytes = buf.freeze();
        assert_eq!(read_compressed(&mut bytes).unwrap(), 0);
    }

    #[test]
    fn test_compressed_254() {
        let mut buf = BytesMut::new();
        write_compressed(&mut buf, 254);
        assert_eq!(buf.as_ref(), &[254]);

        let mut bytes = buf.freeze();
        assert_eq!(read_compressed(&mut bytes).unwrap(), 254);
    }

    #[test]
    fn test_compressed_255() {
        let mut buf = BytesMut::new();
        write_compressed(&mut buf, 255);
        assert_eq!(buf.as_ref(), &[0xFF, 0, 0, 0, 255]);

        let mut bytes = buf.freeze();
        assert_eq!(read_compressed(&mut bytes).unwrap(), 255);
    }

    #[test]
    fn test_compressed_large() {
        let mut buf = BytesMut::new();
        write_compressed(&mut buf, 100_000);
        assert_eq!(buf.len(), 5);

        let mut bytes = buf.freeze();
        assert_eq!(read_compressed(&mut bytes).unwrap(), 100_000);
    }

    #[test]
    fn test_string_roundtrip() {
        let mut buf = BytesMut::new();
        write_string(&mut buf, "hello world");

        let mut bytes = buf.freeze();
        assert_eq!(read_string(&mut bytes).unwrap(), "hello world");
    }

    #[test]
    fn test_string_empty() {
        let mut buf = BytesMut::new();
        write_string(&mut buf, "");

        let mut bytes = buf.freeze();
        assert_eq!(read_string(&mut bytes).unwrap(), "");
    }

    #[test]
    fn test_string_utf8() {
        let mut buf = BytesMut::new();
        let s = "Hej v\u{00e4}rlden \u{1f600}"; // Swedish + emoji
        write_string(&mut buf, s);

        let mut bytes = buf.freeze();
        assert_eq!(read_string(&mut bytes).unwrap(), s);
    }

    #[test]
    fn test_type_simple() {
        let mut buf = BytesMut::new();
        write_type(&mut buf, &Type::string(), 0xFFFF, false);
        assert_eq!(buf.as_ref(), &[TypeClass::String as u8]);

        let mut bytes = buf.freeze();
        let (ty, cache_idx, is_new) = read_type(&mut bytes).unwrap();
        assert_eq!(ty.class, TypeClass::String);
        assert_eq!(cache_idx, 0xFFFF);
        assert!(!is_new);
    }

    #[test]
    fn test_type_interface_new() {
        let iface = Type::interface("com.sun.star.uno.XInterface");
        let mut buf = BytesMut::new();
        write_type(&mut buf, &iface, 0x0005, true);

        // Should be: 0x96 (22 | 0x80), 0x0005, then the string
        assert_eq!(buf[0], TypeClass::Interface as u8 | 0x80);
        assert_eq!(buf[1], 0x00);
        assert_eq!(buf[2], 0x05);

        let mut bytes = buf.freeze();
        let (ty, cache_idx, is_new) = read_type(&mut bytes).unwrap();
        assert_eq!(ty.class, TypeClass::Interface);
        assert_eq!(ty.name, "com.sun.star.uno.XInterface");
        assert_eq!(cache_idx, 5);
        assert!(is_new);
    }

    #[test]
    fn test_type_interface_cached() {
        let iface = Type::interface("");
        let mut buf = BytesMut::new();
        write_type(&mut buf, &iface, 0x0005, false);

        assert_eq!(buf[0], TypeClass::Interface as u8); // No cache flag
        assert_eq!(buf[1], 0x00);
        assert_eq!(buf[2], 0x05);

        let mut bytes = buf.freeze();
        let (ty, cache_idx, is_new) = read_type(&mut bytes).unwrap();
        assert_eq!(ty.class, TypeClass::Interface);
        assert!(ty.name.is_empty());
        assert_eq!(cache_idx, 5);
        assert!(!is_new);
    }

    #[test]
    fn test_value_bool() {
        let mut buf = BytesMut::new();
        write_value(&mut buf, &UnoValue::Bool(true), &Type::boolean());
        write_value(&mut buf, &UnoValue::Bool(false), &Type::boolean());

        let mut bytes = buf.freeze();
        assert_eq!(
            read_value(&mut bytes, &Type::boolean()).unwrap(),
            UnoValue::Bool(true)
        );
        assert_eq!(
            read_value(&mut bytes, &Type::boolean()).unwrap(),
            UnoValue::Bool(false)
        );
    }

    #[test]
    fn test_value_long() {
        let mut buf = BytesMut::new();
        write_value(&mut buf, &UnoValue::Long(12345), &Type::long());

        let mut bytes = buf.freeze();
        assert_eq!(
            read_value(&mut bytes, &Type::long()).unwrap(),
            UnoValue::Long(12345)
        );
    }

    #[test]
    fn test_value_double() {
        let mut buf = BytesMut::new();
        write_value(&mut buf, &UnoValue::Double(3.14), &Type::double());

        let mut bytes = buf.freeze();
        assert_eq!(
            read_value(&mut bytes, &Type::double()).unwrap(),
            UnoValue::Double(3.14)
        );
    }

    #[test]
    fn test_value_string() {
        let mut buf = BytesMut::new();
        write_value(&mut buf, &UnoValue::String("test".into()), &Type::string());

        let mut bytes = buf.freeze();
        assert_eq!(
            read_value(&mut bytes, &Type::string()).unwrap(),
            UnoValue::String("test".into())
        );
    }

    #[test]
    fn test_value_any_void() {
        let mut buf = BytesMut::new();
        let value = UnoValue::Any(Box::new(Any {
            type_desc: Type::void(),
            value: UnoValue::Void,
        }));
        write_value(&mut buf, &value, &Type::any());

        let mut bytes = buf.freeze();
        let result = read_value(&mut bytes, &Type::any()).unwrap();
        assert_eq!(result, value);
    }

    #[test]
    fn test_value_any_string() {
        let mut buf = BytesMut::new();
        let value = UnoValue::Any(Box::new(Any {
            type_desc: Type::string(),
            value: UnoValue::String("hello".into()),
        }));
        write_value(&mut buf, &value, &Type::any());

        let mut bytes = buf.freeze();
        let result = read_value(&mut bytes, &Type::any()).unwrap();
        assert_eq!(result, value);
    }

    #[test]
    fn test_value_sequence_of_bytes() {
        let mut buf = BytesMut::new();
        let items = vec![UnoValue::Byte(1), UnoValue::Byte(2), UnoValue::Byte(3)];
        let value = UnoValue::Sequence(items.clone());
        let ty = Type::sequence_of_bytes();
        write_value(&mut buf, &value, &ty);

        let mut bytes = buf.freeze();
        let result = read_value(&mut bytes, &ty).unwrap();
        assert_eq!(result, UnoValue::Sequence(items));
    }
}
