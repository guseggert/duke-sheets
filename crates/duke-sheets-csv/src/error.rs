//! CSV error types

use thiserror::Error;

/// Result type for CSV operations
pub type CsvResult<T> = std::result::Result<T, CsvError>;

/// Errors that can occur during CSV operations
#[derive(Debug, Error)]
pub enum CsvError {
    /// IO error
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// CSV library error
    #[error("CSV error: {0}")]
    Csv(#[from] csv::Error),

    /// Parse error
    #[error("Parse error at row {row}, column {column}: {message}")]
    Parse {
        row: usize,
        column: usize,
        message: String,
    },

    /// Core error
    #[error("Core error: {0}")]
    Core(#[from] duke_sheets_core::Error),
}
