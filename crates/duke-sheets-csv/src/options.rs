//! CSV options

/// Options for reading CSV files
#[derive(Debug, Clone)]
pub struct CsvReadOptions {
    /// Field delimiter (default: comma)
    pub delimiter: u8,
    /// Quote character (default: double quote)
    pub quote: u8,
    /// Whether first row is header
    pub has_header: bool,
    /// Automatic type detection
    pub auto_detect_types: bool,
}

impl Default for CsvReadOptions {
    fn default() -> Self {
        Self {
            delimiter: b',',
            quote: b'"',
            has_header: true,
            auto_detect_types: true,
        }
    }
}

/// Options for writing CSV files
#[derive(Debug, Clone)]
pub struct CsvWriteOptions {
    /// Field delimiter (default: comma)
    pub delimiter: u8,
    /// Quote character (default: double quote)
    pub quote: u8,
    /// Write header row
    pub write_header: bool,
    /// Line terminator
    pub line_terminator: LineTerminator,
}

impl Default for CsvWriteOptions {
    fn default() -> Self {
        Self {
            delimiter: b',',
            quote: b'"',
            write_header: false,
            line_terminator: LineTerminator::CRLF,
        }
    }
}

/// Line terminator type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineTerminator {
    /// Unix-style (LF)
    LF,
    /// Windows-style (CRLF)
    CRLF,
    /// Mac classic (CR)
    CR,
}
