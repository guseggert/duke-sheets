use thiserror::Error;

pub type XlsbResult<T> = std::result::Result<T, XlsbError>;

#[derive(Debug, Error)]
pub enum XlsbError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("Invalid XLSB format: {0}")]
    InvalidFormat(String),

    #[error("Parse error: {0}")]
    Parse(String),

    #[error("Core error: {0}")]
    Core(#[from] duke_sheets_core::Error),
}
