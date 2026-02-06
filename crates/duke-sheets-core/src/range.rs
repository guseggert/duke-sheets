//! Range type for cell range operations

use crate::cell::{CellAddress, CellData, CellRange, CellValue};
use crate::error::Result;
use crate::style::Style;
use crate::worksheet::Worksheet;

/// A reference to a range of cells in a worksheet
///
/// This provides a convenient API for working with cell ranges.
pub struct Range<'a> {
    worksheet: &'a Worksheet,
    range: CellRange,
}

impl<'a> Range<'a> {
    /// Create a new range reference
    pub fn new(worksheet: &'a Worksheet, range: CellRange) -> Self {
        Self { worksheet, range }
    }

    /// Get the cell range
    pub fn range(&self) -> &CellRange {
        &self.range
    }

    /// Get the start address
    pub fn start(&self) -> CellAddress {
        self.range.start
    }

    /// Get the end address
    pub fn end(&self) -> CellAddress {
        self.range.end
    }

    /// Get the number of rows
    pub fn row_count(&self) -> u32 {
        self.range.row_count()
    }

    /// Get the number of columns
    pub fn col_count(&self) -> u16 {
        self.range.col_count()
    }

    /// Get the total number of cells
    pub fn cell_count(&self) -> u64 {
        self.range.cell_count()
    }

    /// Get a cell by relative position within the range
    pub fn cell(&self, row: u32, col: u16) -> Option<&CellData> {
        let abs_row = self.range.start.row + row;
        let abs_col = self.range.start.col + col;
        self.worksheet.cell_at(abs_row, abs_col)
    }

    /// Get a cell value by relative position
    pub fn value(&self, row: u32, col: u16) -> CellValue {
        let abs_row = self.range.start.row + row;
        let abs_col = self.range.start.col + col;
        self.worksheet.get_value_at(abs_row, abs_col)
    }

    /// Iterate over all cells in the range
    pub fn cells(&self) -> impl Iterator<Item = RangeCell<'a>> + '_ {
        self.range.cells().map(move |addr| {
            let data = self.worksheet.cell_at(addr.row, addr.col);
            RangeCell {
                address: addr,
                data,
            }
        })
    }

    /// Iterate over rows in the range
    pub fn rows(&self) -> impl Iterator<Item = RangeRow<'a>> + '_ {
        (self.range.start.row..=self.range.end.row).map(move |row| RangeRow {
            worksheet: self.worksheet,
            row,
            start_col: self.range.start.col,
            end_col: self.range.end.col,
        })
    }

    /// Get the A1-style address of this range
    pub fn address(&self) -> String {
        self.range.to_a1_string()
    }
}

/// A mutable reference to a range of cells
pub struct RangeMut<'a> {
    worksheet: &'a mut Worksheet,
    range: CellRange,
}

impl<'a> RangeMut<'a> {
    /// Create a new mutable range reference
    pub fn new(worksheet: &'a mut Worksheet, range: CellRange) -> Self {
        Self { worksheet, range }
    }

    /// Get the cell range
    pub fn range(&self) -> &CellRange {
        &self.range
    }

    /// Set a cell value by relative position
    pub fn set_value<V: Into<CellValue>>(&mut self, row: u32, col: u16, value: V) -> Result<()> {
        let abs_row = self.range.start.row + row;
        let abs_col = self.range.start.col + col;
        self.worksheet.set_cell_value_at(abs_row, abs_col, value)
    }

    /// Set all cells to the same value
    pub fn fill<V: Into<CellValue> + Clone>(&mut self, value: V) -> Result<()> {
        self.worksheet.fill_range(&self.range, value)
    }

    /// Clear all cells in the range
    pub fn clear(&mut self) {
        self.worksheet.clear_range(&self.range);
    }

    /// Set style for all cells in the range
    pub fn set_style(&mut self, style: &Style) -> Result<()> {
        for addr in self.range.cells() {
            self.worksheet
                .set_cell_style_at(addr.row, addr.col, style)?;
        }
        Ok(())
    }

    /// Merge the cells in this range
    pub fn merge(&mut self) -> Result<()> {
        self.worksheet.merge_cells(&self.range)
    }
}

/// A cell within a range iteration
pub struct RangeCell<'a> {
    /// The cell's address
    pub address: CellAddress,
    /// The cell data (if any)
    pub data: Option<&'a CellData>,
}

impl<'a> RangeCell<'a> {
    /// Get the cell value
    pub fn value(&self) -> CellValue {
        self.data
            .map(|d| d.value.clone())
            .unwrap_or(CellValue::Empty)
    }

    /// Check if the cell is empty
    pub fn is_empty(&self) -> bool {
        self.data.map(|d| d.value.is_empty()).unwrap_or(true)
    }

    /// Get the row index
    pub fn row(&self) -> u32 {
        self.address.row
    }

    /// Get the column index
    pub fn col(&self) -> u16 {
        self.address.col
    }
}

/// A row within a range iteration
pub struct RangeRow<'a> {
    worksheet: &'a Worksheet,
    row: u32,
    start_col: u16,
    end_col: u16,
}

impl<'a> RangeRow<'a> {
    /// Get the row index
    pub fn index(&self) -> u32 {
        self.row
    }

    /// Iterate over cells in this row
    pub fn cells(&self) -> impl Iterator<Item = RangeCell<'a>> + '_ {
        (self.start_col..=self.end_col).map(move |col| {
            let addr = CellAddress::new(self.row, col);
            let data = self.worksheet.cell_at(self.row, col);
            RangeCell {
                address: addr,
                data,
            }
        })
    }

    /// Get a cell by column offset within the range
    pub fn cell(&self, col_offset: u16) -> Option<&CellData> {
        let col = self.start_col + col_offset;
        if col <= self.end_col {
            self.worksheet.cell_at(self.row, col)
        } else {
            None
        }
    }
}
