//! Error types for the LibreOffice bridge.

use thiserror::Error;

#[derive(Debug, Error)]
pub enum BridgeError {
    #[error("URP protocol error: {0}")]
    Urp(#[from] libreoffice_urp::UrpError),

    #[error("LibreOffice process error: {0}")]
    Process(String),

    #[error("Failed to spawn LibreOffice: {0}")]
    SpawnFailed(#[from] std::io::Error),

    #[error("LibreOffice not found. Install LibreOffice and ensure 'soffice' is in PATH.")]
    NotFound,

    #[error("Connection timeout: LibreOffice did not start within {0} seconds")]
    Timeout(u64),

    #[error("Invalid cell reference: {0}")]
    InvalidCellRef(String),

    #[error("Operation failed: {0}")]
    OperationFailed(String),
}

pub type Result<T> = std::result::Result<T, BridgeError>;
