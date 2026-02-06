//! XLSX error types

use thiserror::Error;

/// Result type for XLSX operations
pub type XlsxResult<T> = std::result::Result<T, XlsxError>;

/// Errors that can occur during XLSX reading/writing
#[derive(Debug, Error)]
pub enum XlsxError {
    /// IO error
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// ZIP error
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// XML error
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    /// Invalid file format
    #[error("Invalid XLSX format: {0}")]
    InvalidFormat(String),

    /// Missing required part
    #[error("Missing required part: {0}")]
    MissingPart(String),

    /// Parse error
    #[error("Parse error: {0}")]
    Parse(String),

    /// Core error
    #[error("Core error: {0}")]
    Core(#[from] duke_sheets_core::Error),
}
