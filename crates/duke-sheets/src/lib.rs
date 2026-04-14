//! # duke-sheets
//!
//! A Rust library for reading, writing, and manipulating spreadsheets.
//!
//! Duke-sheets provides an API similar to Aspose Cells for working with Excel files
//! (XLSX, XLS) and CSV files.
//!
//! ## Features
//!
//! - Read and write XLSX files (Office Open XML)
//! - Read and write XLS files (legacy BIFF8 format) - optional
//! - Read and write CSV files
//! - Full formula evaluation
//! - Cell styling (fonts, colors, borders, etc.)
//! - Charts support
//! - Large file support via streaming APIs
//!
//! ## Example
//!
//! ```rust
//! use duke_sheets::prelude::*;
//!
//! // Create a new workbook
//! let mut workbook = Workbook::new();
//!
//! // Get the first worksheet
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! // Set cell values
//! sheet.set_cell_value("A1", "Hello").unwrap();
//! sheet.set_cell_value("B1", 42.0).unwrap();
//! sheet.set_cell_value("C1", true).unwrap();
//!
//! // Set a formula
//! sheet.set_cell_formula("D1", "=B1*2").unwrap();
//!
//! // Save to file
//! // workbook.save("output.xlsx").unwrap();
//! ```

pub mod calculation;
pub mod prelude;

// Re-export calculation types
pub use calculation::{CalculationOptions, CalculationStats, WorkbookCalculationExt};

// Re-export core types
pub use duke_sheets_core::auto_filter::{ColorFilter, DynamicFilter, DynamicFilterType};
pub use duke_sheets_core::{
    rich_text_to_plain,
    Alignment,
    AutoFilter,
    BorderEdge,
    BorderLineStyle,
    BorderStyle,
    CellAddress,
    // Comments
    CellComment,
    CellData,

    CellError,
    CellRange,
    // Cell types
    CellValue,
    CellView,
    // Conditional formatting types
    CfColorValue,
    CfOperator,
    CfRuleType,
    CfValue,
    CfValueType,
    // Chart sheet types
    ChartSheet,
    Color,
    ColumnFilter,
    ConditionalFormatRule,
    // Data validation types
    CustomFilterCondition,
    CustomFilters,
    DataValidation,
    // Error types
    Error,
    FillStyle,
    FilterColumn,
    FilterOperator,
    FontStyle,
    // Sheet-level types
    FreezePanes,
    HorizontalAlignment,
    Hyperlink,
    IconSetStyle,
    // Locale for cell formatting
    Locale,

    NumberFormat,

    PageBreak,
    PageOrientation,
    PageSetup,
    Result,

    // Rich text types
    RichTextRun,
    RunFont,
    SheetProtection,
    SheetSlot,
    // Style types
    Style,
    StylePool,
    // Table types
    Table,
    TableColumn,
    TableStyleInfo,
    TimePeriod,
    Top10Filter,
    TotalsRowFunction,

    ValidationErrorStyle,
    ValidationOperator,
    ValidationType,
    ValueFilter,
    VerticalAlignment,
    // Main types
    Workbook,
    WorkbookSettings,
    Worksheet,

    MAX_COLS,
    // Constants
    MAX_ROWS,
    MAX_SHEET_NAME_LEN,
};

// Re-export named range module (contains NamedRange, NameScope, NamedRangeCollection)
pub use duke_sheets_core::named_range;

// Re-export formula types
pub use duke_sheets_formula::{
    evaluate, parse_formula, EvaluationContext, FormulaError, FormulaExpr, FormulaResult,
    FormulaValue, ImageInfo, ImageSizing,
};

// Re-export chart types
pub use duke_sheets_chart::{
    Axis, AxisPosition, CellMarker, Chart, ChartType, DataReference, DataSeries, DrawingAnchor,
    EditAs, EmbeddedImage, ImageFormat, Legend,
};

// Re-export I/O types
pub use duke_sheets_csv::{CsvError, CsvReadOptions, CsvReader, CsvWriteOptions, CsvWriter};
#[cfg(feature = "xls")]
pub use duke_sheets_xls::{XlsError, XlsReader};
#[cfg(feature = "xlsb")]
pub use duke_sheets_xlsb::{XlsbError, XlsbReader, XlsbWriter};
pub use duke_sheets_xlsx::{XlsxError, XlsxReader, XlsxWriter};

use std::io::Cursor;
use std::path::Path;

/// Detected file format from magic bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileFormat {
    /// XLSX / Office Open XML (ZIP container starting with PK\x03\x04)
    Xlsx,
    /// XLS / BIFF8 (CFB container starting with \xD0\xCF\x11\xE0)
    Xls,
    /// XLSB / Excel Binary Workbook (ZIP container with BIFF12 internals)
    Xlsb,
    /// Unknown format
    Unknown,
}

/// Sniff the first few bytes of a buffer to determine its file format.
pub fn detect_format(bytes: &[u8]) -> FileFormat {
    if bytes.len() >= 4 && bytes[0..4] == [0x50, 0x4B, 0x03, 0x04] {
        if let Ok(archive) = zip::ZipArchive::new(Cursor::new(bytes)) {
            if archive.index_for_name("xl/workbook.bin").is_some() {
                return FileFormat::Xlsb;
            }
        }
        FileFormat::Xlsx
    } else if bytes.len() >= 8 && bytes[0..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        FileFormat::Xls
    } else {
        FileFormat::Unknown
    }
}

/// Extension trait for Workbook to add file I/O
pub trait WorkbookExt {
    /// Open a workbook from a file
    fn open<P: AsRef<Path>>(path: P) -> Result<Workbook>;

    /// Open a workbook from bytes, auto-detecting the format (XLSX or XLS)
    fn from_bytes(bytes: &[u8]) -> Result<Workbook>;

    /// Save the workbook to a file
    fn save<P: AsRef<Path>>(&self, path: P) -> Result<()>;
}

impl WorkbookExt for Workbook {
    fn open<P: AsRef<Path>>(path: P) -> Result<Workbook> {
        let path = path.as_ref();
        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_lowercase());

        match extension.as_deref() {
            Some("xlsx") | Some("xlsm") | Some("xltx") | Some("xltm") => {
                XlsxReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xls")]
            Some("xls") => XlsReader::read_file(path).map_err(|e| Error::other(e.to_string())),
            Some("csv") => {
                let worksheet = CsvReader::read_file(path, &CsvReadOptions::default())
                    .map_err(|e| Error::other(e.to_string()))?;

                let mut workbook = Workbook::empty();
                workbook.add_existing_worksheet(worksheet)?;
                Ok(workbook)
            }
            #[cfg(feature = "xlsb")]
            Some("xlsb") => XlsbReader::read_file(path).map_err(|e| Error::other(e.to_string())),
            #[cfg(not(feature = "xlsb"))]
            Some("xlsb") => Err(Error::other(
                "XLSB format detected but the 'xlsb' feature is not enabled",
            )),
            _ => Err(Error::other(format!(
                "Unsupported file format: {}",
                path.display()
            ))),
        }
    }

    fn from_bytes(bytes: &[u8]) -> Result<Workbook> {
        match detect_format(bytes) {
            FileFormat::Xlsx => {
                let cursor = Cursor::new(bytes);
                XlsxReader::read(cursor).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xls")]
            FileFormat::Xls => {
                let cursor = Cursor::new(bytes);
                XlsReader::read(cursor).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xls"))]
            FileFormat::Xls => Err(Error::other(
                "XLS format detected but the 'xls' feature is not enabled",
            )),
            #[cfg(feature = "xlsb")]
            FileFormat::Xlsb => {
                let cursor = Cursor::new(bytes);
                XlsbReader::read(cursor).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xlsb"))]
            FileFormat::Xlsb => Err(Error::other(
                "XLSB format detected but the 'xlsb' feature is not enabled",
            )),
            FileFormat::Unknown => Err(Error::other(
                "Unable to detect file format from bytes (expected XLSX or XLS magic bytes)",
            )),
        }
    }

    fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let path = path.as_ref();
        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_lowercase());

        match extension.as_deref() {
            Some("xlsx") => {
                XlsxWriter::write_file(self, path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xlsb")]
            Some("xlsb") => {
                XlsbWriter::write_file(self, path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xlsb"))]
            Some("xlsb") => Err(Error::other("XLSB writing requires the 'xlsb' feature")),
            Some("csv") => {
                if let Some(sheet) = self.worksheet(0) {
                    CsvWriter::write_file(sheet, path, &CsvWriteOptions::default())
                        .map_err(|e| Error::other(e.to_string()))
                } else {
                    Err(Error::other("No worksheets to save"))
                }
            }
            _ => Err(Error::other(format!(
                "Unsupported file format: {}",
                path.display()
            ))),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn make_xlsb_bytes() -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("xl/workbook.bin", options).unwrap();
            zip.write_all(b"fake xlsb content").unwrap();
            zip.finish().unwrap();
        }
        buf
    }

    #[test]
    fn detect_format_identifies_xlsb() {
        let buf = make_xlsb_bytes();
        assert_eq!(detect_format(&buf), FileFormat::Xlsb);
    }

    #[test]
    fn detect_format_xlsx_not_misidentified_as_xlsb() {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<workbook/>").unwrap();
            zip.finish().unwrap();
        }
        assert_eq!(detect_format(&buf), FileFormat::Xlsx);
    }

    #[test]
    #[cfg(not(feature = "xlsb"))]
    fn from_bytes_rejects_xlsb_without_feature() {
        let buf = make_xlsb_bytes();
        let result = Workbook::from_bytes(&buf);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            err_msg.contains("xlsb"),
            "Error should mention xlsb, got: {err_msg}"
        );
    }

    #[test]
    #[cfg(feature = "xlsb")]
    fn from_bytes_reads_xlsb() {
        let buf = make_xlsb_bytes();
        let result = Workbook::from_bytes(&buf);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            !err_msg.contains("Unsupported"),
            "Should attempt to parse, not reject. Got: {err_msg}"
        );
    }

    #[test]
    #[cfg(not(feature = "xlsb"))]
    fn open_rejects_xlsb_without_feature() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsb");
        std::fs::write(&path, b"anything").unwrap();
        let result = Workbook::open(&path);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            err_msg.contains("xlsb"),
            "Error should mention xlsb, got: {err_msg}"
        );
    }

    #[test]
    #[cfg(feature = "xlsb")]
    fn open_reads_xlsb_extension() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsb");
        std::fs::write(&path, b"anything").unwrap();
        let result = Workbook::open(&path);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            !err_msg.contains("Unsupported"),
            "Should attempt to parse, not reject. Got: {err_msg}"
        );
    }

    #[test]
    #[cfg(feature = "xlsb")]
    fn save_xlsb_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "hello").unwrap();
        ws.set_cell_value("B1", 42.0).unwrap();
        ws.set_cell_value("C1", true).unwrap();

        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("roundtrip.xlsb");
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("hello"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Number(42.0));
        assert_eq!(ws2.get_value_at(0, 2), CellValue::Boolean(true));
    }
}
