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
pub use duke_sheets_core::{
    Alignment,
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
    // Conditional formatting types
    CfColorValue,
    CfOperator,
    CfRuleType,
    CfValue,
    CfValueType,
    Color,
    ConditionalFormatRule,
    // Data validation types
    DataValidation,
    // Error types
    Error,
    FillStyle,
    FontStyle,
    HorizontalAlignment,
    IconSetStyle,
    NumberFormat,

    Result,

    // Style types
    Style,
    StylePool,
    TimePeriod,
    ValidationErrorStyle,
    ValidationOperator,
    ValidationType,
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

// Re-export formula types
pub use duke_sheets_formula::{
    evaluate, parse_formula, EvaluationContext, FormulaError, FormulaExpr, FormulaResult,
    FormulaValue,
};

// Re-export chart types
pub use duke_sheets_chart::{
    Axis, Chart, ChartType, DataReference, DataSeries, Legend, LegendPosition,
};

// Re-export I/O types
pub use duke_sheets_csv::{CsvError, CsvReadOptions, CsvReader, CsvWriteOptions, CsvWriter};
pub use duke_sheets_xlsx::{XlsxError, XlsxReader, XlsxWriter};

use std::path::Path;

/// Extension trait for Workbook to add file I/O
pub trait WorkbookExt {
    /// Open a workbook from a file
    fn open<P: AsRef<Path>>(path: P) -> Result<Workbook>;

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
            Some("xlsx") | Some("xlsm") => {
                XlsxReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            Some("csv") => {
                let worksheet = CsvReader::read_file(path, &CsvReadOptions::default())
                    .map_err(|e| Error::other(e.to_string()))?;

                let mut workbook = Workbook::empty();
                workbook.add_existing_worksheet(worksheet)?;
                Ok(workbook)
            }
            _ => Err(Error::other(format!(
                "Unsupported file format: {}",
                path.display()
            ))),
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
