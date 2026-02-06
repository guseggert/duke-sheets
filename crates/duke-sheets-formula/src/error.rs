//! Formula error types

use thiserror::Error;

/// Result type for formula operations
pub type FormulaResult<T> = std::result::Result<T, FormulaError>;

/// Errors that can occur during formula parsing or evaluation
#[derive(Debug, Error)]
pub enum FormulaError {
    /// Formula parse error
    #[error("Parse error: {0}")]
    Parse(String),

    /// Formula evaluation error
    #[error("Evaluation error: {0}")]
    Evaluation(String),

    /// Invalid argument
    #[error("Invalid argument: {0}")]
    Argument(String),

    /// Unknown function
    #[error("Unknown function: {0}")]
    UnknownFunction(String),

    /// Wrong number of arguments
    #[error("Wrong number of arguments for {function}: expected {expected}, got {actual}")]
    ArgumentCount {
        function: String,
        expected: String,
        actual: usize,
    },

    /// Circular reference
    #[error("Circular reference detected")]
    CircularReference,

    /// Reference to invalid cell
    #[error("Invalid reference: {0}")]
    InvalidReference(String),
}
