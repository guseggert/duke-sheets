//! URP message protocol: request/reply encoding, caching, and protocol negotiation.
//!
//! This module handles the message-level protocol:
//! - Request messages (method calls on remote objects)
//! - Reply messages (return values or exceptions)
//! - OID/Type/TID caching (256-entry tables)
//! - Protocol properties negotiation

use bytes::{Buf, BufMut, Bytes, BytesMut};

use crate::error::{Result, UrpError};
use crate::marshal;
use crate::types::Type;

// ============================================================================
// Header flag constants
// ============================================================================

const FLAG_LONGHEADER: u8 = 0x80;
const FLAG_REQUEST: u8 = 0x40;
const FLAG_NEWTYPE: u8 = 0x20;
const FLAG_NEWOID: u8 = 0x10;
const FLAG_NEWTID: u8 = 0x08;
const FLAG_FUNCTIONID16: u8 = 0x04;
const FLAG_MOREFLAGS: u8 = 0x01;

// Reply-specific flags
const FLAG_EXCEPTION: u8 = 0x20;

// More-flags byte
const FLAG_MUSTREPLY: u8 = 0x80;
const FLAG_SYNCHRONOUS: u8 = 0x40;

/// Special OID for protocol properties negotiation.
pub const OID_PROTOCOL_PROPERTIES: &str = "UrpProtocolProperties";

/// Special TID for protocol properties negotiation.
pub const TID_PROTOCOL_PROPERTIES: &[u8] = b".UrpProtocolPropertiesTid";

/// Function ID for queryInterface (on XInterface).
pub const FN_QUERY_INTERFACE: u16 = 0;
/// Function ID for release (on XInterface).
pub const FN_RELEASE: u16 = 2;
/// Function ID for requestChange (on XProtocolProperties).
pub const FN_REQUEST_CHANGE: u16 = 4;
/// Function ID for commitChange (on XProtocolProperties).
pub const FN_COMMIT_CHANGE: u16 = 5;

// ============================================================================
// Parsed message types
// ============================================================================

/// A parsed URP request message.
#[derive(Debug, Clone)]
pub struct UrpRequest {
    pub function_id: u16,
    pub type_name: Type,
    pub oid: String,
    pub tid: Vec<u8>,
    pub must_reply: bool,
    pub synchronous: bool,
    /// Serialized in-parameters (raw bytes, to be deserialized per the method signature).
    pub body: Bytes,
}

/// A parsed URP reply message.
#[derive(Debug, Clone)]
pub struct UrpReply {
    pub tid: Vec<u8>,
    pub is_exception: bool,
    /// If is_exception: serialized exception (Any).
    /// Otherwise: serialized return value + out parameters.
    pub body: Bytes,
}

/// Any message received from the wire.
#[derive(Debug, Clone)]
pub enum UrpMessage {
    Request(UrpRequest),
    Reply(UrpReply),
}

// ============================================================================
// Reader-side cache state
// ============================================================================

/// Reader-side cache state for decoding incoming messages.
pub struct ReaderState {
    pub type_cache: [Option<Type>; 256],
    pub oid_cache: [Option<String>; 256],
    pub tid_cache: [Option<Vec<u8>>; 256],
    /// First-level cache (last-used values)
    pub last_type: Option<Type>,
    pub last_oid: Option<String>,
    pub last_tid: Option<Vec<u8>>,
}

impl ReaderState {
    pub fn new() -> Self {
        Self {
            type_cache: std::array::from_fn(|_| None),
            oid_cache: std::array::from_fn(|_| None),
            tid_cache: std::array::from_fn(|_| None),
            last_type: None,
            last_oid: None,
            last_tid: None,
        }
    }

    /// Decode a message from raw bytes.
    pub fn decode_message(&mut self, data: Bytes) -> Result<UrpMessage> {
        if data.is_empty() {
            return Err(UrpError::Protocol("empty message".into()));
        }

        let first_byte = data[0];

        if first_byte & FLAG_LONGHEADER == 0 {
            // Short request header
            self.decode_short_request(data)
        } else if first_byte & FLAG_REQUEST != 0 {
            // Long request header
            self.decode_long_request(data)
        } else {
            // Reply header
            self.decode_reply(data)
        }
    }

    fn decode_short_request(&mut self, mut data: Bytes) -> Result<UrpMessage> {
        let first_byte = data.get_u8();
        let function_id;

        if first_byte & 0x40 != 0 {
            // FUNCTIONID14: 2-byte header
            let high = (first_byte & 0x3F) as u16;
            if data.remaining() < 1 {
                return Err(UrpError::Protocol(
                    "short request: missing function ID low byte".into(),
                ));
            }
            let low = data.get_u8() as u16;
            function_id = (high << 8) | low;
        } else {
            // 1-byte header: bits 5..0 = function ID
            function_id = (first_byte & 0x3F) as u16;
        }

        // Uses first-level cache for type, OID, TID
        let type_name = self
            .last_type
            .clone()
            .ok_or_else(|| UrpError::Protocol("short request: no cached type".into()))?;
        let oid = self
            .last_oid
            .clone()
            .ok_or_else(|| UrpError::Protocol("short request: no cached OID".into()))?;
        let tid = self
            .last_tid
            .clone()
            .ok_or_else(|| UrpError::Protocol("short request: no cached TID".into()))?;

        Ok(UrpMessage::Request(UrpRequest {
            function_id,
            type_name,
            oid,
            tid,
            must_reply: true, // Default for short header
            synchronous: true,
            body: data,
        }))
    }

    fn decode_long_request(&mut self, mut data: Bytes) -> Result<UrpMessage> {
        let flags1 = data.get_u8();
        let mut must_reply = true;
        let mut synchronous = true;

        if flags1 & FLAG_MOREFLAGS != 0 {
            if data.remaining() < 1 {
                return Err(UrpError::Protocol(
                    "long request: missing more-flags byte".into(),
                ));
            }
            let flags2 = data.get_u8();
            must_reply = flags2 & FLAG_MUSTREPLY != 0;
            synchronous = flags2 & FLAG_SYNCHRONOUS != 0;
        }

        // Function ID
        let function_id = if flags1 & FLAG_FUNCTIONID16 != 0 {
            if data.remaining() < 2 {
                return Err(UrpError::Protocol(
                    "long request: missing function ID".into(),
                ));
            }
            data.get_u16()
        } else {
            if data.remaining() < 1 {
                return Err(UrpError::Protocol(
                    "long request: missing function ID".into(),
                ));
            }
            data.get_u8() as u16
        };

        // Type
        let type_name = if flags1 & FLAG_NEWTYPE != 0 {
            let (ty, cache_index, is_new) = marshal::read_type(&mut data)?;
            let resolved = if is_new {
                if cache_index != 0xFFFF {
                    self.type_cache[cache_index as usize] = Some(ty.clone());
                }
                ty
            } else if cache_index != 0xFFFF {
                self.type_cache[cache_index as usize]
                    .clone()
                    .ok_or_else(|| {
                        UrpError::Cache(format!("type cache miss at index {cache_index}"))
                    })?
            } else {
                ty
            };
            self.last_type = Some(resolved.clone());
            resolved
        } else {
            self.last_type
                .clone()
                .ok_or_else(|| UrpError::Protocol("long request: no cached type".into()))?
        };

        // OID
        let oid = if flags1 & FLAG_NEWOID != 0 {
            let oid_str = marshal::read_string(&mut data)?;
            if data.remaining() < 2 {
                return Err(UrpError::Protocol(
                    "long request: missing OID cache index".into(),
                ));
            }
            let cache_index = data.get_u16();

            let resolved = if oid_str.is_empty() && cache_index != 0xFFFF {
                // Read from cache
                self.oid_cache[cache_index as usize]
                    .clone()
                    .ok_or_else(|| {
                        UrpError::Cache(format!("OID cache miss at index {cache_index}"))
                    })?
            } else {
                if cache_index != 0xFFFF && !oid_str.is_empty() {
                    self.oid_cache[cache_index as usize] = Some(oid_str.clone());
                }
                oid_str
            };
            self.last_oid = Some(resolved.clone());
            resolved
        } else {
            self.last_oid
                .clone()
                .ok_or_else(|| UrpError::Protocol("long request: no cached OID".into()))?
        };

        // TID
        let tid = if flags1 & FLAG_NEWTID != 0 {
            self.read_tid(&mut data)?
        } else {
            self.last_tid
                .clone()
                .ok_or_else(|| UrpError::Protocol("long request: no cached TID".into()))?
        };

        Ok(UrpMessage::Request(UrpRequest {
            function_id,
            type_name,
            oid,
            tid,
            must_reply,
            synchronous,
            body: data,
        }))
    }

    fn decode_reply(&mut self, mut data: Bytes) -> Result<UrpMessage> {
        let flags = data.get_u8();
        let is_exception = flags & FLAG_EXCEPTION != 0;

        // TID
        let tid = if flags & FLAG_NEWTID != 0 {
            self.read_tid(&mut data)?
        } else {
            self.last_tid
                .clone()
                .ok_or_else(|| UrpError::Protocol("reply: no cached TID".into()))?
        };

        Ok(UrpMessage::Reply(UrpReply {
            tid,
            is_exception,
            body: data,
        }))
    }

    fn read_tid(&mut self, data: &mut Bytes) -> Result<Vec<u8>> {
        // TID is a Sequence<byte>: compressed length + bytes
        let len = marshal::read_compressed(data)? as usize;
        if data.remaining() < len {
            return Err(UrpError::Protocol("TID: not enough bytes".into()));
        }
        let tid_bytes = data.copy_to_bytes(len).to_vec();

        if data.remaining() < 2 {
            return Err(UrpError::Protocol("TID: missing cache index".into()));
        }
        let cache_index = data.get_u16();

        let resolved = if tid_bytes.is_empty() && cache_index != 0xFFFF {
            self.tid_cache[cache_index as usize]
                .clone()
                .ok_or_else(|| UrpError::Cache(format!("TID cache miss at index {cache_index}")))?
        } else {
            if cache_index != 0xFFFF && !tid_bytes.is_empty() {
                self.tid_cache[cache_index as usize] = Some(tid_bytes.clone());
            }
            tid_bytes
        };

        self.last_tid = Some(resolved.clone());
        Ok(resolved)
    }
}

// ============================================================================
// Writer-side cache + message encoding
// ============================================================================

/// Writer-side state for encoding outgoing messages.
pub struct WriterState {
    pub type_cache: LruCache<Type>,
    pub oid_cache: LruCache<String>,
    pub tid_cache: LruCache<Vec<u8>>,
    /// First-level cache (last-sent values)
    pub last_type: Option<Type>,
    pub last_oid: Option<String>,
    pub last_tid: Option<Vec<u8>>,
}

impl WriterState {
    pub fn new() -> Self {
        Self {
            type_cache: LruCache::new(),
            oid_cache: LruCache::new(),
            tid_cache: LruCache::new(),
            last_type: None,
            last_oid: None,
            last_tid: None,
        }
    }

    /// Encode a request message.
    pub fn encode_request(
        &mut self,
        function_id: u16,
        type_name: &Type,
        oid: &str,
        tid: &[u8],
        must_reply: bool,
        body: &[u8],
    ) -> BytesMut {
        let mut buf = BytesMut::with_capacity(256);

        // Determine what's new vs cached
        let new_type = self.last_type.as_ref() != Some(type_name);
        let new_oid = self.last_oid.as_deref() != Some(oid);
        let new_tid = self.last_tid.as_deref() != Some(tid);
        let func16 = function_id > 255;

        // Try to use short request form: if type/oid/tid are all cached,
        // function_id fits in 6 bits (or 14 bits), and must_reply is true (default),
        // we can use a compact 1-byte (or 2-byte) short header.
        if !new_type && !new_oid && !new_tid && must_reply {
            if function_id < 0x40 {
                // 1-byte short header: bits 7..6 = 00, bits 5..0 = function_id
                buf.put_u8(function_id as u8);
                buf.put_slice(body);
                return buf;
            } else if function_id < 0x4000 {
                // 2-byte short header: bit 7 = 0, bit 6 = 1, bits 13..0 = function_id
                let high = ((function_id >> 8) as u8) | 0x40;
                let low = (function_id & 0xFF) as u8;
                buf.put_u8(high);
                buf.put_u8(low);
                buf.put_slice(body);
                return buf;
            }
        }

        // Build flags byte for long header
        let mut flags1: u8 = FLAG_LONGHEADER | FLAG_REQUEST;
        if new_type {
            flags1 |= FLAG_NEWTYPE;
        }
        if new_oid {
            flags1 |= FLAG_NEWOID;
        }
        if new_tid {
            flags1 |= FLAG_NEWTID;
        }
        if func16 {
            flags1 |= FLAG_FUNCTIONID16;
        }
        // Only include MOREFLAGS when we need non-default flags
        // Default for long requests without MOREFLAGS: must_reply=true, synchronous=true
        if !must_reply {
            flags1 |= FLAG_MOREFLAGS;
        }

        buf.put_u8(flags1);

        // More-flags byte (only if MOREFLAGS is set)
        if flags1 & FLAG_MOREFLAGS != 0 {
            let mut flags2: u8 = 0;
            if must_reply {
                flags2 |= FLAG_MUSTREPLY | FLAG_SYNCHRONOUS;
            }
            buf.put_u8(flags2);
        }

        // Function ID
        if func16 {
            buf.put_u16(function_id);
        } else {
            buf.put_u8(function_id as u8);
        }

        // Type
        if new_type {
            if type_name.class.is_simple() {
                marshal::write_type(&mut buf, type_name, 0xFFFF, false);
            } else {
                let (cache_index, is_new) = self.type_cache.insert_or_get(type_name.clone());
                marshal::write_type(&mut buf, type_name, cache_index, is_new);
            }
            self.last_type = Some(type_name.clone());
        }

        // OID
        if new_oid {
            let (cache_index, is_new) = self.oid_cache.insert_or_get(oid.to_string());
            if is_new {
                marshal::write_string(&mut buf, oid);
                buf.put_u16(cache_index);
            } else {
                // Cache hit: send empty string + cache index
                marshal::write_string(&mut buf, "");
                buf.put_u16(cache_index);
            }
            self.last_oid = Some(oid.to_string());
        }

        // TID
        if new_tid {
            let tid_vec = tid.to_vec();
            let (cache_index, is_new) = self.tid_cache.insert_or_get(tid_vec);
            if is_new {
                marshal::write_compressed(&mut buf, tid.len() as u32);
                buf.put_slice(tid);
                buf.put_u16(cache_index);
            } else {
                marshal::write_compressed(&mut buf, 0);
                buf.put_u16(cache_index);
            }
            self.last_tid = Some(tid.to_vec());
        }

        // Body
        buf.put_slice(body);

        buf
    }

    /// Encode a reply message.
    pub fn encode_reply(&mut self, tid: &[u8], is_exception: bool, body: &[u8]) -> BytesMut {
        let mut buf = BytesMut::with_capacity(128);

        let new_tid = self.last_tid.as_deref() != Some(tid);

        let mut flags: u8 = FLAG_LONGHEADER; // REQUEST=0
        if is_exception {
            flags |= FLAG_EXCEPTION;
        }
        if new_tid {
            flags |= FLAG_NEWTID;
        }

        buf.put_u8(flags);

        if new_tid {
            let tid_vec = tid.to_vec();
            let (cache_index, is_new) = self.tid_cache.insert_or_get(tid_vec);
            if is_new {
                marshal::write_compressed(&mut buf, tid.len() as u32);
                buf.put_slice(tid);
                buf.put_u16(cache_index);
            } else {
                marshal::write_compressed(&mut buf, 0);
                buf.put_u16(cache_index);
            }
            self.last_tid = Some(tid.to_vec());
        }

        buf.put_slice(body);

        buf
    }
}

// ============================================================================
// Simple 256-entry LRU cache
// ============================================================================

/// A 256-entry cache with simple round-robin eviction.
/// Used by the writer side to track what values the reader has cached.
pub struct LruCache<T: Clone + PartialEq> {
    entries: [Option<T>; 256],
    next_index: u16,
}

impl<T: Clone + PartialEq> LruCache<T> {
    pub fn new() -> Self {
        Self {
            entries: std::array::from_fn(|_| None),
            next_index: 0,
        }
    }

    /// Look up a value in the cache. Returns (cache_index, is_new).
    /// If the value is already cached, returns (existing_index, false).
    /// If not cached, inserts it and returns (new_index, true).
    pub fn insert_or_get(&mut self, value: T) -> (u16, bool) {
        // Check if already cached
        for (i, entry) in self.entries.iter().enumerate() {
            if let Some(existing) = entry {
                if existing == &value {
                    return (i as u16, false);
                }
            }
        }

        // Not cached — insert at next_index (round-robin eviction)
        let index = self.next_index;
        self.entries[index as usize] = Some(value);
        self.next_index = (self.next_index + 1) % 256;
        (index, true)
    }

    /// Get a value from the cache by index.
    pub fn get(&self, index: u16) -> Option<&T> {
        if index < 256 {
            self.entries[index as usize].as_ref()
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::{Type, TypeClass};

    #[test]
    fn test_lru_cache_insert_and_hit() {
        let mut cache = LruCache::new();
        let (idx, is_new) = cache.insert_or_get("hello".to_string());
        assert_eq!(idx, 0);
        assert!(is_new);

        let (idx2, is_new2) = cache.insert_or_get("hello".to_string());
        assert_eq!(idx2, 0);
        assert!(!is_new2);
    }

    #[test]
    fn test_lru_cache_multiple() {
        let mut cache = LruCache::new();
        let (idx1, _) = cache.insert_or_get("a".to_string());
        let (idx2, _) = cache.insert_or_get("b".to_string());
        let (idx3, _) = cache.insert_or_get("c".to_string());
        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(idx3, 2);

        // "a" should still be cached at 0
        let (idx, is_new) = cache.insert_or_get("a".to_string());
        assert_eq!(idx, 0);
        assert!(!is_new);
    }

    #[test]
    fn test_encode_decode_request() {
        let ty = Type::interface("com.sun.star.uno.XInterface");
        let oid = "test-object-id";
        let tid = b"test-tid-1234";
        let body = b"body-data";

        let mut writer = WriterState::new();
        let encoded = writer.encode_request(
            0, // queryInterface
            &ty, oid, tid, true, body,
        );

        let mut reader = ReaderState::new();
        let msg = reader.decode_message(encoded.freeze()).unwrap();

        match msg {
            UrpMessage::Request(req) => {
                assert_eq!(req.function_id, 0);
                assert_eq!(req.type_name.class, TypeClass::Interface);
                assert_eq!(req.type_name.name, "com.sun.star.uno.XInterface");
                assert_eq!(req.oid, "test-object-id");
                assert_eq!(req.tid, b"test-tid-1234");
                assert!(req.must_reply);
                assert_eq!(req.body.as_ref(), b"body-data");
            }
            _ => panic!("expected Request"),
        }
    }

    #[test]
    fn test_encode_decode_reply() {
        let tid = b"test-tid-1234";
        let body = b"return-data";

        let mut writer = WriterState::new();
        let encoded = writer.encode_reply(tid, false, body);

        // Set up reader with matching TID in first-level cache
        // (In practice, the TID would have been set by an earlier request decode)
        let mut reader = ReaderState::new();
        // Pre-populate so the reply can find the TID
        // Actually, the reply includes NEWTID since writer had no last_tid
        let msg = reader.decode_message(encoded.freeze()).unwrap();

        match msg {
            UrpMessage::Reply(reply) => {
                assert!(!reply.is_exception);
                assert_eq!(reply.tid, b"test-tid-1234");
                assert_eq!(reply.body.as_ref(), b"return-data");
            }
            _ => panic!("expected Reply"),
        }
    }

    #[test]
    fn test_encode_cached_second_request() {
        let ty = Type::interface("com.sun.star.uno.XInterface");
        let oid = "test-object-id";
        let tid = b"test-tid";

        let mut writer = WriterState::new();

        // First request: everything is new
        let encoded1 = writer.encode_request(0, &ty, oid, tid, true, b"");
        let len1 = encoded1.len();

        // Second request with same type/oid/tid: should be shorter (cached)
        let encoded2 = writer.encode_request(1, &ty, oid, tid, true, b"");
        let len2 = encoded2.len();

        // The second message should be significantly shorter because type/oid/tid
        // are all cached (only function ID changes)
        assert!(
            len2 < len1,
            "cached request should be shorter: {len2} < {len1}"
        );

        // Verify both decode correctly
        let mut reader = ReaderState::new();
        let msg1 = reader.decode_message(encoded1.freeze()).unwrap();
        let msg2 = reader.decode_message(encoded2.freeze()).unwrap();

        match (msg1, msg2) {
            (UrpMessage::Request(r1), UrpMessage::Request(r2)) => {
                assert_eq!(r1.function_id, 0);
                assert_eq!(r2.function_id, 1);
                assert_eq!(r1.oid, r2.oid);
                assert_eq!(r1.type_name, r2.type_name);
                assert_eq!(r1.tid, r2.tid);
            }
            _ => panic!("expected two Requests"),
        }
    }
}
