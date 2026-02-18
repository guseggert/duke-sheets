//! Cell storage implementation
//!
//! This module provides efficient sparse storage for spreadsheet cells.
//! Only non-empty cells are stored, using a row-based BTreeMap structure.

use std::collections::{BTreeMap, HashMap};

use super::{CellValue, StringPool};
use crate::style::StylePool;

/// Storage mode for cell data
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StorageMode {
    /// Standard in-memory storage (default)
    InMemory,
    /// Memory-optimized mode for large files
    /// Uses more CPU to reduce memory usage
    MemoryOptimized,
}

impl Default for StorageMode {
    fn default() -> Self {
        StorageMode::InMemory
    }
}

/// Complete data for a single cell
#[derive(Debug, Clone)]
pub struct CellData {
    /// The cell's value
    pub value: CellValue,
    /// Index into the style pool (0 = default style)
    pub style_index: u32,
}

impl CellData {
    /// Create a new cell with a value and default style
    pub fn new(value: CellValue) -> Self {
        Self {
            value,
            style_index: 0,
        }
    }

    /// Create a new cell with a value and style
    pub fn with_style(value: CellValue, style_index: u32) -> Self {
        Self { value, style_index }
    }

    /// Create an empty cell
    pub fn empty() -> Self {
        Self {
            value: CellValue::Empty,
            style_index: 0,
        }
    }

    /// Check if this cell is effectively empty (no value and default style)
    pub fn is_empty(&self) -> bool {
        self.value.is_empty() && self.style_index == 0
    }
}

impl Default for CellData {
    fn default() -> Self {
        Self::empty()
    }
}

/// Information about a spill source (a formula that produces an array)
#[derive(Debug, Clone)]
pub struct SpillInfo {
    /// Number of rows in the spilled array
    pub rows: u32,
    /// Number of columns in the spilled array
    pub cols: u16,
}

impl SpillInfo {
    /// Create new spill info
    pub fn new(rows: u32, cols: u16) -> Self {
        Self { rows, cols }
    }

    /// Get the spill range as (end_row, end_col) offsets from source
    pub fn end_offsets(&self) -> (u32, u16) {
        (self.rows.saturating_sub(1), self.cols.saturating_sub(1))
    }
}

/// Sparse row-based storage for worksheet cells
///
/// Design decisions:
/// - Uses BTreeMap for ordered iteration (required for streaming writes)
/// - Row-major layout matches Excel's internal structure
/// - Only stores non-empty cells (sparse)
/// - Can handle millions of cells with reasonable memory usage
///
/// Structure: `BTreeMap<row_index, BTreeMap<col_index, CellData>>`
#[derive(Debug)]
pub struct CellStorage {
    /// Row index → column map
    rows: BTreeMap<u32, BTreeMap<u16, CellData>>,

    /// Shared string pool for deduplication
    pub(crate) string_pool: StringPool,

    /// Shared style pool for deduplication
    pub(crate) style_pool: StylePool,

    /// Default row height in points (default: 15.0)
    default_row_height: f64,

    /// Default column width in characters (default: 8.43)
    default_column_width: f64,

    /// Custom row heights
    row_heights: BTreeMap<u32, f64>,

    /// Hidden rows
    hidden_rows: BTreeMap<u32, bool>,

    /// Row outline levels (for grouping)
    #[allow(dead_code)]
    row_outline_levels: BTreeMap<u32, u8>,

    /// Custom column widths
    column_widths: BTreeMap<u16, f64>,

    /// Hidden columns
    hidden_columns: BTreeMap<u16, bool>,

    /// Column outline levels
    #[allow(dead_code)]
    column_outline_levels: BTreeMap<u16, u8>,

    /// Merged cell regions
    merged_regions: Vec<crate::CellRange>,

    /// Storage mode
    mode: StorageMode,

    /// Cached bounds (invalidated on changes)
    cached_bounds: Option<CachedBounds>,

    /// Spill sources: maps source cell (row, col) to spill info
    /// This tracks which cells have active spill ranges
    spill_sources: HashMap<(u32, u16), SpillInfo>,
}

#[derive(Debug, Clone, Copy)]
struct CachedBounds {
    min_row: u32,
    max_row: u32,
    min_col: u16,
    max_col: u16,
}

impl CellStorage {
    /// Create a new empty cell storage
    pub fn new() -> Self {
        Self {
            rows: BTreeMap::new(),
            string_pool: StringPool::new(),
            style_pool: StylePool::new(),
            default_row_height: 15.0,
            default_column_width: 8.43,
            row_heights: BTreeMap::new(),
            hidden_rows: BTreeMap::new(),
            row_outline_levels: BTreeMap::new(),
            column_widths: BTreeMap::new(),
            hidden_columns: BTreeMap::new(),
            column_outline_levels: BTreeMap::new(),
            merged_regions: Vec::new(),
            mode: StorageMode::InMemory,
            cached_bounds: None,
            spill_sources: HashMap::new(),
        }
    }

    /// Create a new cell storage with specified mode
    pub fn with_mode(mode: StorageMode) -> Self {
        let mut storage = Self::new();
        storage.mode = mode;
        storage
    }

    /// Get the storage mode
    pub fn mode(&self) -> StorageMode {
        self.mode
    }

    /// Get a cell value
    pub fn get(&self, row: u32, col: u16) -> Option<&CellData> {
        self.rows.get(&row).and_then(|r| r.get(&col))
    }

    /// Get a mutable cell value
    pub fn get_mut(&mut self, row: u32, col: u16) -> Option<&mut CellData> {
        self.rows.get_mut(&row).and_then(|r| r.get_mut(&col))
    }

    /// Set a cell value
    ///
    /// If the cell data is empty (no value, default style), the cell is removed.
    pub fn set(&mut self, row: u32, col: u16, data: CellData) {
        self.invalidate_bounds();

        if data.is_empty() {
            // Remove empty cells to save memory
            if let Some(row_map) = self.rows.get_mut(&row) {
                row_map.remove(&col);
                if row_map.is_empty() {
                    self.rows.remove(&row);
                }
            }
        } else {
            self.rows.entry(row).or_default().insert(col, data);
        }
    }

    /// Set just the cell value (preserving style)
    pub fn set_value(&mut self, row: u32, col: u16, value: CellValue) {
        self.invalidate_bounds();

        if let Some(cell) = self.get_mut(row, col) {
            cell.value = value;
            // Remove if now empty
            if cell.is_empty() {
                self.set(row, col, CellData::empty());
            }
        } else if !value.is_empty() {
            self.set(row, col, CellData::new(value));
        }
    }

    /// Set just the cell style (preserving value)
    pub fn set_style(&mut self, row: u32, col: u16, style_index: u32) {
        if let Some(cell) = self.get_mut(row, col) {
            cell.style_index = style_index;
        } else if style_index != 0 {
            // Create cell with empty value but custom style
            self.set(
                row,
                col,
                CellData::with_style(CellValue::Empty, style_index),
            );
        }
    }

    /// Remove a cell
    pub fn remove(&mut self, row: u32, col: u16) -> Option<CellData> {
        self.invalidate_bounds();

        let result = self.rows.get_mut(&row).and_then(|r| r.remove(&col));

        // Clean up empty rows
        if let Some(row_map) = self.rows.get(&row) {
            if row_map.is_empty() {
                self.rows.remove(&row);
            }
        }

        result
    }

    /// Clear all cells
    pub fn clear(&mut self) {
        self.rows.clear();
        self.merged_regions.clear();
        self.invalidate_bounds();
    }

    /// Get the number of non-empty cells
    pub fn cell_count(&self) -> usize {
        self.rows.values().map(|r| r.len()).sum()
    }

    /// Check if storage is empty
    pub fn is_empty(&self) -> bool {
        self.rows.is_empty()
    }

    /// Get the bounds of used cells
    ///
    /// Returns (min_row, min_col, max_row, max_col) or None if empty
    pub fn used_bounds(&self) -> Option<(u32, u16, u32, u16)> {
        if self.rows.is_empty() {
            return None;
        }

        // Use cached bounds if available
        if let Some(bounds) = self.cached_bounds {
            return Some((
                bounds.min_row,
                bounds.min_col,
                bounds.max_row,
                bounds.max_col,
            ));
        }

        let min_row = *self.rows.keys().next()?;
        let max_row = *self.rows.keys().next_back()?;

        let mut min_col = u16::MAX;
        let mut max_col = 0u16;

        for row_data in self.rows.values() {
            if let Some(&col) = row_data.keys().next() {
                min_col = min_col.min(col);
            }
            if let Some(&col) = row_data.keys().next_back() {
                max_col = max_col.max(col);
            }
        }

        Some((min_row, min_col, max_row, max_col))
    }

    /// Iterate over all cells in row order
    pub fn iter(&self) -> impl Iterator<Item = (u32, u16, &CellData)> {
        self.rows
            .iter()
            .flat_map(|(&row, cols)| cols.iter().map(move |(&col, data)| (row, col, data)))
    }

    /// Iterate over cells in a specific row
    pub fn iter_row(&self, row: u32) -> impl Iterator<Item = (u16, &CellData)> {
        self.rows
            .get(&row)
            .into_iter()
            .flat_map(|cols| cols.iter().map(|(&col, data)| (col, data)))
    }

    /// Iterate over row indices that have data
    pub fn row_indices(&self) -> impl Iterator<Item = u32> + '_ {
        self.rows.keys().copied()
    }

    /// Get default row height
    pub fn default_row_height(&self) -> f64 {
        self.default_row_height
    }

    /// Set default row height
    pub fn set_default_row_height(&mut self, height: f64) {
        self.default_row_height = height;
    }

    /// Get row height (returns default if not customized)
    pub fn row_height(&self, row: u32) -> f64 {
        self.row_heights
            .get(&row)
            .copied()
            .unwrap_or(self.default_row_height)
    }

    /// Set custom row height
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        if (height - self.default_row_height).abs() < 0.001 {
            self.row_heights.remove(&row);
        } else {
            self.row_heights.insert(row, height);
        }
    }

    /// Check if row is hidden
    pub fn is_row_hidden(&self, row: u32) -> bool {
        self.hidden_rows.get(&row).copied().unwrap_or(false)
    }

    /// Set row hidden state
    pub fn set_row_hidden(&mut self, row: u32, hidden: bool) {
        if hidden {
            self.hidden_rows.insert(row, true);
        } else {
            self.hidden_rows.remove(&row);
        }
    }

    /// Get default column width
    pub fn default_column_width(&self) -> f64 {
        self.default_column_width
    }

    /// Set default column width
    pub fn set_default_column_width(&mut self, width: f64) {
        self.default_column_width = width;
    }

    /// Get column width (returns default if not customized)
    pub fn column_width(&self, col: u16) -> f64 {
        self.column_widths
            .get(&col)
            .copied()
            .unwrap_or(self.default_column_width)
    }

    /// Set custom column width
    pub fn set_column_width(&mut self, col: u16, width: f64) {
        if (width - self.default_column_width).abs() < 0.001 {
            self.column_widths.remove(&col);
        } else {
            self.column_widths.insert(col, width);
        }
    }

    /// Check if column is hidden
    pub fn is_column_hidden(&self, col: u16) -> bool {
        self.hidden_columns.get(&col).copied().unwrap_or(false)
    }

    /// Set column hidden state
    pub fn set_column_hidden(&mut self, col: u16, hidden: bool) {
        if hidden {
            self.hidden_columns.insert(col, true);
        } else {
            self.hidden_columns.remove(&col);
        }
    }

    /// Get all custom row heights (row index → height in points).
    pub fn custom_row_heights(&self) -> &std::collections::BTreeMap<u32, f64> {
        &self.row_heights
    }

    /// Get all hidden rows (row index → true).
    pub fn hidden_rows(&self) -> &std::collections::BTreeMap<u32, bool> {
        &self.hidden_rows
    }

    /// Get all custom column widths (column index → width in characters).
    pub fn custom_column_widths(&self) -> &std::collections::BTreeMap<u16, f64> {
        &self.column_widths
    }

    /// Get all hidden columns (column index → true).
    pub fn hidden_columns(&self) -> &std::collections::BTreeMap<u16, bool> {
        &self.hidden_columns
    }

    /// Get merged regions
    pub fn merged_regions(&self) -> &[crate::CellRange] {
        &self.merged_regions
    }

    /// Add a merged region
    pub fn add_merged_region(&mut self, range: crate::CellRange) {
        self.merged_regions.push(range);
    }

    /// Remove a merged region
    pub fn remove_merged_region(&mut self, index: usize) -> Option<crate::CellRange> {
        if index < self.merged_regions.len() {
            Some(self.merged_regions.remove(index))
        } else {
            None
        }
    }

    /// Clear all merged regions
    pub fn clear_merged_regions(&mut self) {
        self.merged_regions.clear();
    }

    /// Check if a cell is part of a merged region
    pub fn is_merged(&self, row: u32, col: u16) -> bool {
        let addr = crate::CellAddress::new(row, col);
        self.merged_regions.iter().any(|r| r.contains(&addr))
    }

    /// Get the string pool
    pub fn string_pool(&self) -> &StringPool {
        &self.string_pool
    }

    /// Get the string pool mutably
    pub fn string_pool_mut(&mut self) -> &mut StringPool {
        &mut self.string_pool
    }

    /// Get the style pool
    pub fn style_pool(&self) -> &StylePool {
        &self.style_pool
    }

    /// Get the style pool mutably
    pub fn style_pool_mut(&mut self) -> &mut StylePool {
        &mut self.style_pool
    }

    // ==================== Spill Management ====================

    /// Check if a range can be used for spilling
    ///
    /// A range can be spilled to if all cells in the range (except the source) are either:
    /// - Empty
    /// - Already a spill target from this same source
    ///
    /// Returns false if any cell would block the spill (non-empty, merged, etc.)
    pub fn can_spill_to(
        &self,
        source_row: u32,
        source_col: u16,
        num_rows: u32,
        num_cols: u16,
    ) -> bool {
        // Check each cell in the potential spill range
        for row_offset in 0..num_rows {
            for col_offset in 0..num_cols {
                let row = source_row + row_offset;
                let col = source_col + col_offset;

                // Skip the source cell itself
                if row_offset == 0 && col_offset == 0 {
                    continue;
                }

                // Check if cell exists and would block
                if let Some(cell) = self.get(row, col) {
                    match &cell.value {
                        CellValue::Empty => continue, // OK
                        CellValue::SpillTarget {
                            source_row: sr,
                            source_col: sc,
                            ..
                        } => {
                            // OK if it's from the same source
                            if *sr == source_row && *sc == source_col {
                                continue;
                            }
                            return false; // Different source - blocked
                        }
                        _ => return false, // Any other value blocks
                    }
                }

                // Check if cell is part of a merged region
                if self.is_merged(row, col) {
                    return false;
                }
            }
        }

        true
    }

    /// Register a spill source
    pub fn register_spill_source(&mut self, row: u32, col: u16, info: SpillInfo) {
        self.spill_sources.insert((row, col), info);
    }

    /// Unregister a spill source
    pub fn unregister_spill_source(&mut self, row: u32, col: u16) -> Option<SpillInfo> {
        self.spill_sources.remove(&(row, col))
    }

    /// Get spill info for a source cell
    pub fn get_spill_info(&self, row: u32, col: u16) -> Option<&SpillInfo> {
        self.spill_sources.get(&(row, col))
    }

    /// Check if a cell is a spill source
    pub fn is_spill_source(&self, row: u32, col: u16) -> bool {
        self.spill_sources.contains_key(&(row, col))
    }

    /// Clear all spill targets for a given source
    ///
    /// This removes all SpillTarget cells that reference the given source cell.
    /// Call this before recalculating a formula or when a spill source is deleted.
    pub fn clear_spill_targets(&mut self, source_row: u32, source_col: u16) {
        if let Some(info) = self.spill_sources.get(&(source_row, source_col)).cloned() {
            self.invalidate_bounds();

            // Remove all spill target cells for this source
            for row_offset in 0..info.rows {
                for col_offset in 0..info.cols {
                    let row = source_row + row_offset;
                    let col = source_col + col_offset;

                    // Skip the source cell itself
                    if row_offset == 0 && col_offset == 0 {
                        continue;
                    }

                    // Remove the spill target cell
                    if let Some(cell) = self.get(row, col) {
                        if matches!(cell.value, CellValue::SpillTarget { .. }) {
                            self.remove(row, col);
                        }
                    }
                }
            }

            // Unregister the source
            self.spill_sources.remove(&(source_row, source_col));
        }
    }

    /// Get all spill sources
    pub fn spill_sources(&self) -> impl Iterator<Item = ((u32, u16), &SpillInfo)> {
        self.spill_sources.iter().map(|(k, v)| (*k, v))
    }

    /// Invalidate cached bounds
    fn invalidate_bounds(&mut self) {
        self.cached_bounds = None;
    }
}

impl Default for CellStorage {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_basic_operations() {
        let mut storage = CellStorage::new();

        // Set and get
        storage.set(0, 0, CellData::new(CellValue::Number(42.0)));
        let cell = storage.get(0, 0).unwrap();
        assert_eq!(cell.value.as_number(), Some(42.0));

        // Get non-existent
        assert!(storage.get(1, 1).is_none());
    }

    #[test]
    fn test_empty_cells_not_stored() {
        let mut storage = CellStorage::new();

        storage.set(0, 0, CellData::new(CellValue::Number(42.0)));
        assert_eq!(storage.cell_count(), 1);

        // Setting empty removes the cell
        storage.set(0, 0, CellData::empty());
        assert_eq!(storage.cell_count(), 0);
        assert!(storage.get(0, 0).is_none());
    }

    #[test]
    fn test_used_bounds() {
        let mut storage = CellStorage::new();

        assert!(storage.used_bounds().is_none());

        storage.set(5, 3, CellData::new(CellValue::Number(1.0)));
        storage.set(10, 7, CellData::new(CellValue::Number(2.0)));
        storage.set(2, 1, CellData::new(CellValue::Number(3.0)));

        let (min_row, min_col, max_row, max_col) = storage.used_bounds().unwrap();
        assert_eq!(min_row, 2);
        assert_eq!(min_col, 1);
        assert_eq!(max_row, 10);
        assert_eq!(max_col, 7);
    }

    #[test]
    fn test_row_column_properties() {
        let mut storage = CellStorage::new();

        // Default values
        assert_eq!(storage.row_height(0), 15.0);
        assert_eq!(storage.column_width(0), 8.43);
        assert!(!storage.is_row_hidden(0));
        assert!(!storage.is_column_hidden(0));

        // Custom values
        storage.set_row_height(5, 30.0);
        storage.set_column_width(3, 20.0);
        storage.set_row_hidden(10, true);
        storage.set_column_hidden(5, true);

        assert_eq!(storage.row_height(5), 30.0);
        assert_eq!(storage.column_width(3), 20.0);
        assert!(storage.is_row_hidden(10));
        assert!(storage.is_column_hidden(5));
    }

    #[test]
    fn test_iteration() {
        let mut storage = CellStorage::new();

        storage.set(0, 0, CellData::new(CellValue::Number(1.0)));
        storage.set(0, 1, CellData::new(CellValue::Number(2.0)));
        storage.set(1, 0, CellData::new(CellValue::Number(3.0)));

        let cells: Vec<_> = storage.iter().collect();
        assert_eq!(cells.len(), 3);

        // Should be in row order
        assert_eq!(cells[0].0, 0); // row 0
        assert_eq!(cells[1].0, 0); // row 0
        assert_eq!(cells[2].0, 1); // row 1
    }
}
