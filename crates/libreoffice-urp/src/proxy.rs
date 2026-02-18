//! Dynamic proxy for calling methods on remote UNO objects.
//!
//! A `UnoProxy` represents a remote UNO object identified by its OID.
//! It provides methods to invoke UNO methods by encoding the parameters,
//! sending the request, and decoding the reply.

use bytes::BytesMut;

use crate::error::{Result, UrpError};
use crate::interface::MethodDef;
use crate::marshal;
use crate::protocol::LruCache;
use crate::types::{Type, TypeClass, UnoValue};

/// A proxy handle for a remote UNO object.
///
/// This is a lightweight reference — the actual communication goes through
/// the `Connection` which holds the transport and protocol state.
#[derive(Debug, Clone)]
pub struct UnoProxy {
    /// The Object IDentifier for the remote object.
    pub oid: String,
    /// The UNO interface type this proxy is typed as.
    pub interface_type: Type,
}

impl UnoProxy {
    /// Create a new proxy for a remote object.
    pub fn new(oid: String, interface_type: Type) -> Self {
        Self {
            oid,
            interface_type,
        }
    }

    /// Check if this is a null (empty) proxy.
    pub fn is_null(&self) -> bool {
        self.oid.is_empty()
    }
}

/// Serialize method parameters into the wire format body bytes.
pub fn serialize_params(method: &MethodDef, args: &[UnoValue]) -> Result<BytesMut> {
    if args.len() != method.params.len() {
        return Err(UrpError::Protocol(format!(
            "method {} expects {} params, got {}",
            method.name,
            method.params.len(),
            args.len()
        )));
    }

    let mut buf = BytesMut::with_capacity(256);
    for (arg, param_type) in args.iter().zip(method.params.iter()) {
        let ty = param_type.to_type();
        marshal::write_value(&mut buf, arg, &ty);
    }
    Ok(buf)
}

/// Serialize method parameters with OID caching for interface references.
pub fn serialize_params_cached(
    method: &MethodDef,
    args: &[UnoValue],
    oid_cache: &mut LruCache<String>,
) -> Result<BytesMut> {
    if args.len() != method.params.len() {
        return Err(UrpError::Protocol(format!(
            "method {} expects {} params, got {}",
            method.name,
            method.params.len(),
            args.len()
        )));
    }

    let mut buf = BytesMut::with_capacity(256);
    for (arg, param_type) in args.iter().zip(method.params.iter()) {
        let ty = param_type.to_type();
        marshal::write_value_cached(&mut buf, arg, &ty, Some(oid_cache));
    }
    Ok(buf)
}

/// Deserialize a return value from reply body bytes.
pub fn deserialize_return(method: &MethodDef, mut body: bytes::Bytes) -> Result<UnoValue> {
    if method.return_type.class == TypeClass::Void {
        return Ok(UnoValue::Void);
    }
    marshal::read_value(&mut body, &method.return_type)
}

/// Deserialize a return value from reply body bytes, using the shared OID cache.
pub fn deserialize_return_cached(
    method: &MethodDef,
    mut body: bytes::Bytes,
    oid_cache: &mut [Option<String>; 256],
) -> Result<UnoValue> {
    if method.return_type.class == TypeClass::Void {
        return Ok(UnoValue::Void);
    }
    marshal::read_value_cached(&mut body, &method.return_type, oid_cache)
}

/// Extract an interface OID from a return value.
///
/// UNO methods that return interfaces wrap the result in an Any containing
/// an interface reference. This helper extracts the OID from the common patterns.
pub fn extract_oid_from_return(value: &UnoValue) -> Option<String> {
    match value {
        UnoValue::Interface(oid) if !oid.is_empty() => Some(oid.clone()),
        UnoValue::Any(any) => match &any.value {
            UnoValue::Interface(oid) if !oid.is_empty() => Some(oid.clone()),
            _ => None,
        },
        _ => None,
    }
}

/// Extract an interface proxy from a queryInterface return value.
///
/// queryInterface returns an Any. If the interface is supported, the Any
/// contains an interface reference with the object's OID. If not supported,
/// the Any is void.
pub fn extract_query_interface_result(
    value: UnoValue,
    requested_type: Type,
) -> Result<Option<UnoProxy>> {
    match value {
        UnoValue::Any(any) => {
            if any.type_desc.class == TypeClass::Void {
                // Interface not supported
                Ok(None)
            } else {
                match any.value {
                    UnoValue::Interface(oid) if !oid.is_empty() => {
                        Ok(Some(UnoProxy::new(oid, requested_type)))
                    }
                    _ => Ok(None),
                }
            }
        }
        _ => Err(UrpError::Protocol(format!(
            "queryInterface returned unexpected value type: {value:?}"
        ))),
    }
}
