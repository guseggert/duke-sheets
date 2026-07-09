#![cfg_attr(test, allow(clippy::approx_constant))]
//! # duke-sheets-core
//!
//! Core data structures for the duke-sheets spreadsheet library.
//!
//! This crate provides the fundamental types used throughout duke-sheets:
//! - [`CellValue`] - Represents cell values (numbers, strings, booleans, errors, formulas)
//! - [`CellAddress`] and [`CellRange`] - Cell addressing and ranges
//! - [`Style`] - Cell formatting (fonts, fills, borders, etc.)
//! - [`Workbook`], [`Worksheet`] - The main document structures
//!
//! ## Example
//!
//! ```rust
//! use duke_sheets_core::{Workbook, CellValue};
//!
//! let mut workbook = Workbook::new();
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! // Using string addresses
//! sheet.set_cell_value("A1", "Hello").unwrap();
//! sheet.set_cell_value("B1", 42.0).unwrap();
//!
//! // Or using row/column indices (0-based)
//! sheet.set_cell_value_at(1, 0, CellValue::String("World".into())).unwrap();
//! sheet.set_cell_value_at(1, 1, CellValue::Number(3.14)).unwrap();
//! ```

pub mod auto_filter;
pub mod cell;
pub mod column;
pub mod comment;
pub mod conditional_format;
pub mod error;
pub mod form_control;
pub mod hyperlink;
pub mod locale;
pub mod named_range;
pub mod range;
pub mod rich_text;
pub mod row;
pub mod style;
pub mod table;
pub mod validation;
pub mod workbook;
pub mod worksheet;

// Re-exports for convenience
pub use auto_filter::{
    AutoFilter, ColumnFilter, CustomFilterCondition, CustomFilters, FilterColumn, FilterOperator,
    Top10Filter, ValueFilter,
};
pub use cell::{
    format_cell_value, CellAddress, CellData, CellError, CellRange, CellValue, CellView, FormulaData,
};
pub use column::{Column, ColumnData};
pub use comment::CellComment;
pub use conditional_format::{
    CfColorValue, CfOperator, CfRuleType, CfValue, CfValueType, ConditionalFormatRule,
    IconSetStyle, TimePeriod,
};
pub use error::{Error, Result};
pub use form_control::{CheckState, FormControl, FormControlKind, ListSelection};
pub use hyperlink::Hyperlink;
pub use locale::Locale;
pub use rich_text::{rich_text_to_plain, RichTextRun, RunFont};
pub use table::{Table, TableColumn, TableStyleInfo, TotalsRowFunction};
pub use validation::{DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType};
pub use workbook::{ChartSheet, SheetSlot, Workbook, WorkbookSettings};
pub use worksheet::{
    FreezePanes, ImageInfo, ImageSizing, PageBreak, PageOrientation, PageSetup, Selection,
    SheetProtection, SheetVisibility, SplitPanes, Worksheet,
};

// Re-export all style types for convenience
pub use style::{
    Alignment, BorderEdge, BorderLineStyle, BorderStyle, Color, FillStyle, FontStyle,
    HorizontalAlignment, NumberFormat, Style, StylePool, VerticalAlignment,
};

/// Maximum number of rows in a worksheet (Excel limit)
pub const MAX_ROWS: u32 = 1_048_576;

/// Maximum number of columns in a worksheet (Excel limit)
pub const MAX_COLS: u16 = 16_384;

/// Maximum length of a sheet name
pub const MAX_SHEET_NAME_LEN: usize = 31;
