//! Error types for duke-sheets-core

use thiserror::Error;

/// Result type alias using [`Error`]
pub type Result<T> = std::result::Result<T, Error>;

/// Errors that can occur in duke-sheets-core
#[derive(Debug, Error)]
pub enum Error {
    /// Invalid cell address format
    #[error("Invalid cell address: {0}")]
    InvalidAddress(String),

    /// Invalid cell range format
    #[error("Invalid cell range: {0}")]
    InvalidRange(String),

    /// Row index out of bounds
    #[error("Row index {0} out of bounds (max: {1})")]
    RowOutOfBounds(u32, u32),

    /// Column index out of bounds
    #[error("Column index {0} out of bounds (max: {1})")]
    ColumnOutOfBounds(u16, u16),

    /// Sheet index out of bounds
    #[error("Sheet index {0} out of bounds (count: {1})")]
    SheetOutOfBounds(usize, usize),

    /// Sheet not found by name
    #[error("Sheet not found: {0}")]
    SheetNotFound(String),

    /// Invalid sheet name
    #[error("Invalid sheet name: {0}")]
    InvalidSheetName(String),

    /// Duplicate sheet name
    #[error("Sheet name already exists: {0}")]
    DuplicateSheetName(String),

    /// Invalid named range
    #[error("Invalid named range: {0}")]
    InvalidName(String),

    /// Invalid style index
    #[error("Invalid style index: {0}")]
    InvalidStyleIndex(u32),

    /// Invalid value type for operation
    #[error("Invalid value type: expected {expected}, got {actual}")]
    InvalidValueType {
        expected: &'static str,
        actual: &'static str,
    },

    /// Merged cell conflict
    #[error("Cell {0} is part of a merged region")]
    MergedCellConflict(String),

    /// Circular reference detected
    #[error("Circular reference detected involving cell {0}")]
    CircularReference(String),

    /// Formula parse error
    #[error("Formula parse error: {0}")]
    FormulaParse(String),

    /// Generic error with message
    #[error("{0}")]
    Other(String),
}

impl Error {
    /// Create a new "other" error with a message
    pub fn other<S: Into<String>>(msg: S) -> Self {
        Error::Other(msg.into())
    }
}
