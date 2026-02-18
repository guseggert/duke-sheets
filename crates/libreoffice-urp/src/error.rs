//! Error types for the URP protocol implementation.

use thiserror::Error;

/// Errors that can occur during URP communication.
#[derive(Debug, Error)]
pub enum UrpError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Protocol error: {0}")]
    Protocol(String),

    #[error("Marshaling error: {0}")]
    Marshal(String),

    #[error("Connection closed")]
    ConnectionClosed,

    #[error("UNO exception from remote: {0}")]
    RemoteException(String),

    #[error("Cache error: {0}")]
    Cache(String),

    #[error("Unknown type class: {0}")]
    UnknownTypeClass(u8),

    #[error("Unknown interface: {0}")]
    UnknownInterface(String),

    #[error("Timeout waiting for reply")]
    Timeout,

    #[error("Protocol negotiation failed: {0}")]
    NegotiationFailed(String),
}

pub type Result<T> = std::result::Result<T, UrpError>;
