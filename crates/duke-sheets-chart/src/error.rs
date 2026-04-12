use thiserror::Error;

#[derive(Debug, Error)]
pub enum ChartParseError {
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("Parse error: {0}")]
    Parse(String),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
}

pub type ChartParseResult<T> = Result<T, ChartParseError>;
