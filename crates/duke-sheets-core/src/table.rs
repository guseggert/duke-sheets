//! Table (ListObject) support
//!
//! This module provides support for Excel tables (ListObjects) in worksheets.
//! Tables define a named, structured range with column headers, optional
//! totals row, auto-filter, and table styling.
//!
//! ## Example
//!
//! ```rust
//! use duke_sheets_core::{Workbook, Table, TableColumn, CellRange};
//!
//! let mut workbook = Workbook::new();
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! // Create a table over A1:C5
//! let table = Table {
//!     id: 1,
//!     name: "SalesData".into(),
//!     display_name: "SalesData".into(),
//!     reference: CellRange::parse("A1:C5").unwrap(),
//!     columns: vec![
//!         TableColumn::new(1, "Product"),
//!         TableColumn::new(2, "Region"),
//!         TableColumn::new(3, "Revenue"),
//!     ],
//!     style_info: None,
//!     header_row_count: 1,
//!     totals_row_count: 0,
//!     totals_row_shown: true,
//! };
//! sheet.add_table(table);
//! ```

use crate::cell::CellRange;

/// A table (ListObject) in a worksheet.
///
/// Represents an OOXML table definition (`<table>` element in
/// `xl/tables/tableN.xml`). Each table has a unique ID and name
/// within the workbook, covers a rectangular range, and defines
/// column headers and optional totals.
#[derive(Debug, Clone, PartialEq)]
pub struct Table {
    /// Unique table ID within the workbook (1-based).
    pub id: u32,
    /// Internal name used in structured references (e.g., `Table1`).
    pub name: String,
    /// Display name (usually same as `name`).
    pub display_name: String,
    /// Table range including headers and totals (e.g., `A1:D10`).
    pub reference: CellRange,
    /// Column definitions (one per column in the range).
    pub columns: Vec<TableColumn>,
    /// Table style configuration.
    pub style_info: Option<TableStyleInfo>,
    /// Number of header rows (0 = no headers, 1 = one header row; default 1).
    pub header_row_count: u32,
    /// Number of totals rows (0 = no totals row, 1 = totals row shown; default 0).
    pub totals_row_count: u32,
    /// Whether the totals row is shown when `totals_row_count` is 0.
    /// Default: true (Excel shows totals row button in table design tab).
    pub totals_row_shown: bool,
}

impl Table {
    /// Create a new table with the given ID, name, and range.
    ///
    /// Columns must be added separately. Header row defaults to 1,
    /// totals row defaults to 0 (hidden).
    pub fn new(id: u32, name: impl Into<String>, reference: CellRange) -> Self {
        let name = name.into();
        let display_name = name.clone();
        Self {
            id,
            name,
            display_name,
            reference,
            columns: Vec::new(),
            style_info: None,
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: true,
        }
    }

    /// Whether this table has a header row.
    pub fn has_header_row(&self) -> bool {
        self.header_row_count > 0
    }

    /// Whether this table has a visible totals row.
    pub fn has_totals_row(&self) -> bool {
        self.totals_row_count > 0
    }
}

/// A column definition within a table.
#[derive(Debug, Clone, PartialEq)]
pub struct TableColumn {
    /// Column ID (1-based, unique within the table).
    pub id: u32,
    /// Column header name.
    pub name: String,
    /// Aggregate function for the totals row.
    pub totals_row_function: Option<TotalsRowFunction>,
    /// Custom formula for the totals row (when function is `Custom`).
    pub totals_row_formula: Option<String>,
    /// Custom label for the totals row (instead of a function result).
    pub totals_row_label: Option<String>,
    /// Calculated column formula (applied to all data rows).
    pub calculated_column_formula: Option<String>,
}

impl TableColumn {
    /// Create a new column with the given ID and name.
    pub fn new(id: u32, name: impl Into<String>) -> Self {
        Self {
            id,
            name: name.into(),
            totals_row_function: None,
            totals_row_formula: None,
            totals_row_label: None,
            calculated_column_formula: None,
        }
    }
}

/// Built-in aggregate functions for the totals row.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TotalsRowFunction {
    /// SUBTOTAL(109, ...) - average
    Average,
    /// SUBTOTAL(103, ...) - count non-empty
    Count,
    /// SUBTOTAL(102, ...) - count numbers
    CountNums,
    /// SUBTOTAL(104, ...) - max
    Max,
    /// SUBTOTAL(105, ...) - min
    Min,
    /// SUBTOTAL(109, ...) - sum
    Sum,
    /// SUBTOTAL(107, ...) - standard deviation
    StdDev,
    /// SUBTOTAL(110, ...) - variance
    Var,
    /// Custom formula (see `TableColumn::totals_row_formula`)
    Custom,
    /// No aggregation (blank totals cell)
    None,
}

impl TotalsRowFunction {
    /// Parse from OOXML attribute value.
    pub fn from_ooxml(s: &str) -> Option<Self> {
        match s {
            "average" => Some(Self::Average),
            "count" => Some(Self::Count),
            "countNums" => Some(Self::CountNums),
            "max" => Some(Self::Max),
            "min" => Some(Self::Min),
            "sum" => Some(Self::Sum),
            "stdDev" => Some(Self::StdDev),
            "var" => Some(Self::Var),
            "custom" => Some(Self::Custom),
            "none" => Some(Self::None),
            _ => Option::None,
        }
    }

    /// Convert to OOXML attribute value.
    pub fn to_ooxml(self) -> &'static str {
        match self {
            Self::Average => "average",
            Self::Count => "count",
            Self::CountNums => "countNums",
            Self::Max => "max",
            Self::Min => "min",
            Self::Sum => "sum",
            Self::StdDev => "stdDev",
            Self::Var => "var",
            Self::Custom => "custom",
            Self::None => "none",
        }
    }
}

/// Table style configuration (`<tableStyleInfo>` element).
#[derive(Debug, Clone, PartialEq)]
pub struct TableStyleInfo {
    /// Built-in or custom table style name (e.g., `"TableStyleMedium2"`).
    pub name: Option<String>,
    /// Show special formatting for the first column.
    pub show_first_column: bool,
    /// Show special formatting for the last column.
    pub show_last_column: bool,
    /// Show alternating row stripes.
    pub show_row_stripes: bool,
    /// Show alternating column stripes.
    pub show_column_stripes: bool,
}

impl Default for TableStyleInfo {
    fn default() -> Self {
        Self {
            name: None,
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_new() {
        let range = CellRange::parse("A1:D10").unwrap();
        let table = Table::new(1, "Sales", range.clone());
        assert_eq!(table.id, 1);
        assert_eq!(table.name, "Sales");
        assert_eq!(table.display_name, "Sales");
        assert_eq!(table.reference, range);
        assert!(table.columns.is_empty());
        assert!(table.has_header_row());
        assert!(!table.has_totals_row());
    }

    #[test]
    fn test_table_column_new() {
        let col = TableColumn::new(1, "Revenue");
        assert_eq!(col.id, 1);
        assert_eq!(col.name, "Revenue");
        assert!(col.totals_row_function.is_none());
        assert!(col.calculated_column_formula.is_none());
    }

    #[test]
    fn test_totals_row_function_roundtrip() {
        let functions = [
            TotalsRowFunction::Average,
            TotalsRowFunction::Count,
            TotalsRowFunction::CountNums,
            TotalsRowFunction::Max,
            TotalsRowFunction::Min,
            TotalsRowFunction::Sum,
            TotalsRowFunction::StdDev,
            TotalsRowFunction::Var,
            TotalsRowFunction::Custom,
            TotalsRowFunction::None,
        ];
        for f in &functions {
            let s = f.to_ooxml();
            let parsed = TotalsRowFunction::from_ooxml(s).unwrap();
            assert_eq!(*f, parsed);
        }
    }

    #[test]
    fn test_table_style_info_default() {
        let info = TableStyleInfo::default();
        assert!(info.name.is_none());
        assert!(!info.show_first_column);
        assert!(!info.show_last_column);
        assert!(info.show_row_stripes);
        assert!(!info.show_column_stripes);
    }
}
