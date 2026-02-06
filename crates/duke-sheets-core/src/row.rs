//! Row types

use crate::cell::CellData;

/// Row metadata
#[derive(Debug, Clone)]
pub struct Row {
    /// Row index (0-based)
    pub index: u32,
    /// Custom height (None = default)
    pub height: Option<f64>,
    /// Row is hidden
    pub hidden: bool,
    /// Outline/grouping level (0-7)
    pub outline_level: u8,
    /// Row-level style index (None = no row style)
    pub style_index: Option<u32>,
    /// Row is collapsed (in outline)
    pub collapsed: bool,
}

impl Row {
    /// Create a new row with default settings
    pub fn new(index: u32) -> Self {
        Self {
            index,
            height: None,
            hidden: false,
            outline_level: 0,
            style_index: None,
            collapsed: false,
        }
    }

    /// Check if this row has any custom settings
    pub fn has_custom_settings(&self) -> bool {
        self.height.is_some()
            || self.hidden
            || self.outline_level > 0
            || self.style_index.is_some()
            || self.collapsed
    }
}

/// Row data including cells (used during iteration)
#[derive(Debug)]
pub struct RowData<'a> {
    /// Row index
    pub index: u32,
    /// Cells in this row
    pub cells: Vec<(u16, &'a CellData)>,
}

impl<'a> RowData<'a> {
    /// Create a new row data
    pub fn new(index: u32, cells: Vec<(u16, &'a CellData)>) -> Self {
        Self { index, cells }
    }

    /// Get a cell by column index
    pub fn cell(&self, col: u16) -> Option<&CellData> {
        self.cells
            .iter()
            .find(|(c, _)| *c == col)
            .map(|(_, data)| *data)
    }

    /// Check if row has any cells
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    /// Number of cells in row
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }
}
