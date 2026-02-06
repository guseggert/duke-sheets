//! XLS error types

use thiserror::Error;

/// Result type for XLS operations
pub type XlsResult<T> = std::result::Result<T, XlsError>;

/// Errors that can occur during XLS reading/writing
#[derive(Debug, Error)]
pub enum XlsError {
    /// IO error (also covers CFB errors which use std::io::Error)
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// Invalid file format
    #[error("Invalid XLS format: {0}")]
    InvalidFormat(String),

    /// Unsupported version
    #[error("Unsupported XLS version: {0}")]
    UnsupportedVersion(String),

    /// Parse error
    #[error("Parse error: {0}")]
    Parse(String),

    /// Core error
    #[error("Core error: {0}")]
    Core(#[from] duke_sheets_core::Error),
}
