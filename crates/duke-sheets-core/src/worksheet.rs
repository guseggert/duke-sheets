//! Worksheet type

use std::collections::HashMap;

use crate::cell::{CellAddress, CellData, CellRange, CellStorage, CellValue};
use crate::comment::CellComment;
use crate::conditional_format::ConditionalFormatRule;
use crate::error::{Error, Result};
use crate::style::Style;
use crate::validation::DataValidation;
use crate::{MAX_COLS, MAX_ROWS};

/// A worksheet (single sheet in a workbook)
#[derive(Debug)]
pub struct Worksheet {
    /// Sheet name
    name: String,
    /// Cell storage
    cells: CellStorage,
    /// Sheet is visible
    visible: bool,
    /// Sheet is selected
    selected: bool,
    /// Sheet protection settings
    #[allow(dead_code)]
    protection: Option<SheetProtection>,
    /// Freeze pane settings
    freeze_panes: Option<FreezePanes>,
    /// Print settings
    #[allow(dead_code)]
    page_setup: PageSetup,
    /// Tab color
    tab_color: Option<crate::style::Color>,
    /// Cell comments (keyed by (row, col))
    comments: HashMap<(u32, u16), CellComment>,
    /// Unique comment authors
    comment_authors: Vec<String>,
    /// Data validations
    data_validations: Vec<DataValidation>,
    /// Conditional formatting rules
    conditional_formats: Vec<ConditionalFormatRule>,
}

impl Worksheet {
    /// Create a new worksheet with the given name
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            cells: CellStorage::new(),
            visible: true,
            selected: false,
            protection: None,
            freeze_panes: None,
            page_setup: PageSetup::default(),
            tab_color: None,
            comments: HashMap::new(),
            comment_authors: Vec::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
        }
    }

    /// Get the sheet name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Set the sheet name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Check if the sheet is visible
    pub fn is_visible(&self) -> bool {
        self.visible
    }

    /// Set sheet visibility
    pub fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }

    /// Check if the sheet is selected
    pub fn is_selected(&self) -> bool {
        self.selected
    }

    /// Set sheet selected state
    pub fn set_selected(&mut self, selected: bool) {
        self.selected = selected;
    }

    /// Get the tab color
    pub fn tab_color(&self) -> Option<crate::style::Color> {
        self.tab_color
    }

    /// Set the tab color
    pub fn set_tab_color(&mut self, color: Option<crate::style::Color>) {
        self.tab_color = color;
    }

    // === Cell Access ===

    /// Get a cell value by address string (e.g., "A1")
    pub fn cell(&self, address: &str) -> Result<Option<&CellData>> {
        let addr = CellAddress::parse(address)?;
        Ok(self.cells.get(addr.row, addr.col))
    }

    /// Get a cell value by row and column indices
    pub fn cell_at(&self, row: u32, col: u16) -> Option<&CellData> {
        self.cells.get(row, col)
    }

    /// Get a mutable cell by row and column indices
    pub fn cell_at_mut(&mut self, row: u32, col: u16) -> Option<&mut CellData> {
        self.cells.get_mut(row, col)
    }

    /// Get cell value (convenience method)
    pub fn get_value(&self, address: &str) -> Result<CellValue> {
        let addr = CellAddress::parse(address)?;
        Ok(self
            .cells
            .get(addr.row, addr.col)
            .map(|c| c.value.clone())
            .unwrap_or(CellValue::Empty))
    }

    /// Get cell value by indices
    pub fn get_value_at(&self, row: u32, col: u16) -> CellValue {
        self.cells
            .get(row, col)
            .map(|c| c.value.clone())
            .unwrap_or(CellValue::Empty)
    }

    /// Get a cell's style index by address string.
    ///
    /// Returns 0 if the cell does not exist or has the default style.
    pub fn cell_style_index(&self, address: &str) -> Result<u32> {
        let addr = CellAddress::parse(address)?;
        Ok(self.cell_style_index_at(addr.row, addr.col))
    }

    /// Get a cell's style index by row/column.
    ///
    /// Returns 0 if the cell does not exist or has the default style.
    pub fn cell_style_index_at(&self, row: u32, col: u16) -> u32 {
        self.cells.get(row, col).map(|c| c.style_index).unwrap_or(0)
    }

    /// Get a style by its index in this worksheet's style pool.
    pub fn style_by_index(&self, style_index: u32) -> Option<&Style> {
        self.cells.style_pool().get(style_index)
    }

    /// Get the non-default style applied to a cell, if any.
    pub fn cell_style_at(&self, row: u32, col: u16) -> Option<&Style> {
        let idx = self.cell_style_index_at(row, col);
        if idx == 0 {
            None
        } else {
            self.style_by_index(idx)
        }
    }

    /// Get the non-default style applied to a cell by address, if any.
    pub fn cell_style(&self, address: &str) -> Result<Option<&Style>> {
        let addr = CellAddress::parse(address)?;
        Ok(self.cell_style_at(addr.row, addr.col))
    }

    // === Cell Modification ===

    /// Set a cell value by address string
    pub fn set_cell_value<V: Into<CellValue>>(&mut self, address: &str, value: V) -> Result<()> {
        let addr = CellAddress::parse(address)?;
        self.set_cell_value_at(addr.row, addr.col, value)
    }

    /// Set a cell value by row and column indices
    pub fn set_cell_value_at<V: Into<CellValue>>(
        &mut self,
        row: u32,
        col: u16,
        value: V,
    ) -> Result<()> {
        self.validate_cell_position(row, col)?;
        self.cells.set_value(row, col, value.into());
        Ok(())
    }

    /// Set a cell formula by address string
    pub fn set_cell_formula(&mut self, address: &str, formula: &str) -> Result<()> {
        let addr = CellAddress::parse(address)?;
        self.set_cell_formula_at(addr.row, addr.col, formula)
    }

    /// Set a cell formula by row and column indices
    pub fn set_cell_formula_at(&mut self, row: u32, col: u16, formula: &str) -> Result<()> {
        self.validate_cell_position(row, col)?;

        // Ensure formula starts with '='
        let formula = if formula.starts_with('=') {
            formula.to_string()
        } else {
            format!("={}", formula)
        };

        self.cells.set_value(row, col, CellValue::formula(formula));
        Ok(())
    }

    /// Set a cell style by address string
    pub fn set_cell_style(&mut self, address: &str, style: &Style) -> Result<()> {
        let addr = CellAddress::parse(address)?;
        self.set_cell_style_at(addr.row, addr.col, style)
    }

    /// Set a cell style by row and column indices
    pub fn set_cell_style_at(&mut self, row: u32, col: u16, style: &Style) -> Result<()> {
        self.validate_cell_position(row, col)?;
        let style_index = self.cells.style_pool_mut().get_or_insert(style.clone());
        self.cells.set_style(row, col, style_index);
        Ok(())
    }

    /// Clear a cell
    pub fn clear_cell(&mut self, address: &str) -> Result<()> {
        let addr = CellAddress::parse(address)?;
        self.cells.remove(addr.row, addr.col);
        Ok(())
    }

    /// Clear a cell by indices
    pub fn clear_cell_at(&mut self, row: u32, col: u16) {
        self.cells.remove(row, col);
    }

    // === Range Operations ===

    /// Get the used range (bounds of all non-empty cells)
    pub fn used_range(&self) -> Option<CellRange> {
        self.cells
            .used_bounds()
            .map(|(min_row, min_col, max_row, max_col)| {
                CellRange::from_indices(min_row, min_col, max_row, max_col)
            })
    }

    /// Clear all cells in a range
    pub fn clear_range(&mut self, range: &CellRange) {
        for addr in range.cells() {
            self.cells.remove(addr.row, addr.col);
        }
    }

    /// Set the same value for all cells in a range
    pub fn fill_range<V: Into<CellValue> + Clone>(
        &mut self,
        range: &CellRange,
        value: V,
    ) -> Result<()> {
        let value = value.into();
        for addr in range.cells() {
            self.validate_cell_position(addr.row, addr.col)?;
            self.cells.set_value(addr.row, addr.col, value.clone());
        }
        Ok(())
    }

    // === Row/Column Operations ===

    /// Get row height
    pub fn row_height(&self, row: u32) -> f64 {
        self.cells.row_height(row)
    }

    /// Set row height
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.cells.set_row_height(row, height);
    }

    /// Check if row is hidden
    pub fn is_row_hidden(&self, row: u32) -> bool {
        self.cells.is_row_hidden(row)
    }

    /// Set row hidden state
    pub fn set_row_hidden(&mut self, row: u32, hidden: bool) {
        self.cells.set_row_hidden(row, hidden);
    }

    /// Get column width
    pub fn column_width(&self, col: u16) -> f64 {
        self.cells.column_width(col)
    }

    /// Set column width
    pub fn set_column_width(&mut self, col: u16, width: f64) {
        self.cells.set_column_width(col, width);
    }

    /// Check if column is hidden
    pub fn is_column_hidden(&self, col: u16) -> bool {
        self.cells.is_column_hidden(col)
    }

    /// Set column hidden state
    pub fn set_column_hidden(&mut self, col: u16, hidden: bool) {
        self.cells.set_column_hidden(col, hidden);
    }

    /// Get all custom row heights (row index → height in points).
    pub fn custom_row_heights(&self) -> &std::collections::BTreeMap<u32, f64> {
        self.cells.custom_row_heights()
    }

    /// Get all hidden rows (row index → true).
    pub fn hidden_rows(&self) -> &std::collections::BTreeMap<u32, bool> {
        self.cells.hidden_rows()
    }

    /// Get all custom column widths (column index → width in characters).
    pub fn custom_column_widths(&self) -> &std::collections::BTreeMap<u16, f64> {
        self.cells.custom_column_widths()
    }

    /// Get all hidden columns (column index → true).
    pub fn hidden_columns(&self) -> &std::collections::BTreeMap<u16, bool> {
        self.cells.hidden_columns()
    }

    // === Merged Cells ===

    /// Get merged regions
    pub fn merged_regions(&self) -> &[CellRange] {
        self.cells.merged_regions()
    }

    /// Merge cells
    pub fn merge_cells(&mut self, range: &CellRange) -> Result<()> {
        // Check for overlap with existing merged regions
        for existing in self.cells.merged_regions() {
            if range.overlaps(existing) {
                return Err(Error::MergedCellConflict(range.to_string()));
            }
        }
        self.cells.add_merged_region(*range);
        Ok(())
    }

    /// Unmerge cells
    pub fn unmerge_cells(&mut self, range: &CellRange) -> bool {
        let mut found = None;
        for (i, existing) in self.cells.merged_regions().iter().enumerate() {
            if existing == range {
                found = Some(i);
                break;
            }
        }

        if let Some(i) = found {
            self.cells.remove_merged_region(i);
            true
        } else {
            false
        }
    }

    // === Freeze Panes ===

    /// Get freeze pane settings
    pub fn freeze_panes(&self) -> Option<&FreezePanes> {
        self.freeze_panes.as_ref()
    }

    /// Set freeze panes
    pub fn set_freeze_panes(&mut self, row: u32, col: u16) {
        if row == 0 && col == 0 {
            self.freeze_panes = None;
        } else {
            self.freeze_panes = Some(FreezePanes { row, col });
        }
    }

    /// Remove freeze panes
    pub fn unfreeze_panes(&mut self) {
        self.freeze_panes = None;
    }

    // === Cell Comments ===

    /// Set a comment on a cell by address string
    ///
    /// # Example
    ///
    /// ```rust
    /// use duke_sheets_core::{Worksheet, CellComment};
    ///
    /// let mut ws = Worksheet::new("Test");
    /// ws.set_comment("A1", CellComment::new("Author", "This is a note")).unwrap();
    /// ```
    pub fn set_comment(&mut self, address: &str, comment: CellComment) -> Result<()> {
        let addr = CellAddress::parse(address)?;
        self.set_comment_at(addr.row, addr.col, comment);
        Ok(())
    }

    /// Set a comment on a cell by row and column indices
    pub fn set_comment_at(&mut self, row: u32, col: u16, comment: CellComment) {
        // Track unique authors
        if !comment.author.is_empty() && !self.comment_authors.contains(&comment.author) {
            self.comment_authors.push(comment.author.clone());
        }
        self.comments.insert((row, col), comment);
    }

    /// Get a comment from a cell by address string
    pub fn comment(&self, address: &str) -> Result<Option<&CellComment>> {
        let addr = CellAddress::parse(address)?;
        Ok(self.comment_at(addr.row, addr.col))
    }

    /// Get a comment from a cell by row and column indices
    pub fn comment_at(&self, row: u32, col: u16) -> Option<&CellComment> {
        self.comments.get(&(row, col))
    }

    /// Get a mutable reference to a comment
    pub fn comment_at_mut(&mut self, row: u32, col: u16) -> Option<&mut CellComment> {
        self.comments.get_mut(&(row, col))
    }

    /// Remove a comment from a cell by address string
    pub fn remove_comment(&mut self, address: &str) -> Result<Option<CellComment>> {
        let addr = CellAddress::parse(address)?;
        Ok(self.remove_comment_at(addr.row, addr.col))
    }

    /// Remove a comment from a cell by row and column indices
    pub fn remove_comment_at(&mut self, row: u32, col: u16) -> Option<CellComment> {
        self.comments.remove(&(row, col))
    }

    /// Check if a cell has a comment
    pub fn has_comment(&self, address: &str) -> Result<bool> {
        let addr = CellAddress::parse(address)?;
        Ok(self.has_comment_at(addr.row, addr.col))
    }

    /// Check if a cell has a comment by row and column indices
    pub fn has_comment_at(&self, row: u32, col: u16) -> bool {
        self.comments.contains_key(&(row, col))
    }

    /// Get the number of comments in this worksheet
    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    /// Iterate over all comments: ((row, col), comment)
    pub fn comments(&self) -> impl Iterator<Item = ((u32, u16), &CellComment)> {
        self.comments.iter().map(|(&k, v)| (k, v))
    }

    /// Get the list of unique comment authors
    pub fn comment_authors(&self) -> &[String] {
        &self.comment_authors
    }

    /// Clear all comments from this worksheet
    pub fn clear_comments(&mut self) {
        self.comments.clear();
        self.comment_authors.clear();
    }

    // === Data Validation ===

    /// Add a data validation rule
    ///
    /// # Example
    ///
    /// ```rust
    /// use duke_sheets_core::{Worksheet, DataValidation, CellRange};
    ///
    /// let mut ws = Worksheet::new("Test");
    /// let validation = DataValidation::list("Yes,No,Maybe")
    ///     .with_range(CellRange::parse("A1:A10").unwrap());
    /// ws.add_data_validation(validation);
    /// ```
    pub fn add_data_validation(&mut self, validation: DataValidation) {
        self.data_validations.push(validation);
    }

    /// Get all data validations
    pub fn data_validations(&self) -> &[DataValidation] {
        &self.data_validations
    }

    /// Get a mutable reference to all data validations
    pub fn data_validations_mut(&mut self) -> &mut Vec<DataValidation> {
        &mut self.data_validations
    }

    /// Get data validation for a specific cell
    pub fn data_validation_at(&self, row: u32, col: u16) -> Option<&DataValidation> {
        self.data_validations
            .iter()
            .find(|v| v.applies_to(row, col))
    }

    /// Remove data validation by index
    pub fn remove_data_validation(&mut self, index: usize) -> Option<DataValidation> {
        if index < self.data_validations.len() {
            Some(self.data_validations.remove(index))
        } else {
            None
        }
    }

    /// Get the number of data validations
    pub fn data_validation_count(&self) -> usize {
        self.data_validations.len()
    }

    /// Clear all data validations
    pub fn clear_data_validations(&mut self) {
        self.data_validations.clear();
    }

    // === Conditional Formatting ===

    /// Add a conditional formatting rule
    ///
    /// # Example
    ///
    /// ```rust
    /// use duke_sheets_core::{Worksheet, ConditionalFormatRule, CellRange};
    /// use duke_sheets_core::style::{Color, Style};
    ///
    /// let mut ws = Worksheet::new("Test");
    /// let rule = ConditionalFormatRule::cell_is_greater_than("100")
    ///     .with_range(CellRange::parse("A1:A10").unwrap())
    ///     .with_format(Style::new().fill_color(Color::rgb(255, 199, 206)));
    /// ws.add_conditional_format(rule);
    /// ```
    pub fn add_conditional_format(&mut self, rule: ConditionalFormatRule) {
        self.conditional_formats.push(rule);
    }

    /// Get all conditional formatting rules
    pub fn conditional_formats(&self) -> &[ConditionalFormatRule] {
        &self.conditional_formats
    }

    /// Get a mutable reference to all conditional formatting rules
    pub fn conditional_formats_mut(&mut self) -> &mut Vec<ConditionalFormatRule> {
        &mut self.conditional_formats
    }

    /// Get conditional formatting rules for a specific cell
    pub fn conditional_formats_at(&self, row: u32, col: u16) -> Vec<&ConditionalFormatRule> {
        self.conditional_formats
            .iter()
            .filter(|r| r.applies_to(row, col))
            .collect()
    }

    /// Remove conditional formatting rule by index
    pub fn remove_conditional_format(&mut self, index: usize) -> Option<ConditionalFormatRule> {
        if index < self.conditional_formats.len() {
            Some(self.conditional_formats.remove(index))
        } else {
            None
        }
    }

    /// Get the number of conditional formatting rules
    pub fn conditional_format_count(&self) -> usize {
        self.conditional_formats.len()
    }

    /// Clear all conditional formatting rules
    pub fn clear_conditional_formats(&mut self) {
        self.conditional_formats.clear();
    }

    // === Internal ===

    /// Get cell storage (internal use)
    #[allow(dead_code)]
    pub(crate) fn cells(&self) -> &CellStorage {
        &self.cells
    }

    /// Get mutable cell storage (internal use)
    #[allow(dead_code)]
    pub(crate) fn cells_mut(&mut self) -> &mut CellStorage {
        &mut self.cells
    }

    /// Validate cell position
    fn validate_cell_position(&self, row: u32, col: u16) -> Result<()> {
        if row >= MAX_ROWS {
            return Err(Error::RowOutOfBounds(row, MAX_ROWS - 1));
        }
        if col >= MAX_COLS {
            return Err(Error::ColumnOutOfBounds(col, MAX_COLS - 1));
        }
        Ok(())
    }

    /// Get the number of non-empty cells
    pub fn cell_count(&self) -> usize {
        self.cells.cell_count()
    }

    /// Check if the worksheet is empty
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    /// Iterate over all non-empty cells
    pub fn iter_cells(&self) -> impl Iterator<Item = (u32, u16, &CellData)> {
        self.cells.iter()
    }

    // === Formula calculation support ===

    /// Iterate over all formula cells: (row, col, formula_text)
    pub fn formula_cells(&self) -> impl Iterator<Item = (u32, u16, &str)> {
        self.cells.iter().filter_map(|(row, col, cell)| {
            if let CellValue::Formula { text, .. } = &cell.value {
                Some((row, col, text.as_str()))
            } else {
                None
            }
        })
    }

    /// Get the formula text at a cell position (if it's a formula)
    pub fn get_formula_at(&self, row: u32, col: u16) -> Option<&str> {
        self.cells.get(row, col).and_then(|cell| {
            if let CellValue::Formula { text, .. } = &cell.value {
                Some(text.as_str())
            } else {
                None
            }
        })
    }

    /// Set the cached result value of a formula cell
    /// Returns Ok(()) if the cell is a formula and was updated,
    /// or an error if the cell doesn't exist or isn't a formula
    pub fn set_formula_result(&mut self, row: u32, col: u16, value: CellValue) -> Result<()> {
        let cell = self.cells.get_mut(row, col).ok_or_else(|| {
            Error::InvalidAddress(format!("Cell at ({}, {}) not found", row, col))
        })?;

        match &mut cell.value {
            CellValue::Formula { cached_value, .. } => {
                *cached_value = Some(Box::new(value));
                Ok(())
            }
            _ => Err(Error::InvalidAddress(format!(
                "Cell at ({}, {}) is not a formula",
                row, col
            ))),
        }
    }

    /// Get the cached value of a formula cell, or the cell value directly if not a formula
    pub fn get_calculated_value_at(&self, row: u32, col: u16) -> Option<&CellValue> {
        self.cells.get(row, col).map(|cell| match &cell.value {
            CellValue::Formula {
                cached_value: Some(v),
                ..
            } => v.as_ref(),
            other => other,
        })
    }

    // ==================== Dynamic Array Spill Support ====================

    /// Set the result of a dynamic array formula, spilling to adjacent cells
    ///
    /// This method:
    /// 1. Checks if the spill range is available
    /// 2. If available, writes the array values to cells
    /// 3. If blocked, returns Err with CellError::Spill
    ///
    /// # Arguments
    /// * `row` - Row of the source formula cell
    /// * `col` - Column of the source formula cell  
    /// * `array` - The array result (outer vec is rows, inner vec is columns)
    ///
    /// # Returns
    /// * `Ok(())` if the array was successfully spilled
    /// * `Err(Error)` if the spill was blocked
    pub fn set_array_formula_result(
        &mut self,
        row: u32,
        col: u16,
        array: Vec<Vec<CellValue>>,
    ) -> Result<()> {
        let num_rows = array.len() as u32;
        let num_cols = array.first().map(|r| r.len() as u16).unwrap_or(0);

        if num_rows == 0 || num_cols == 0 {
            return Err(Error::Other("Empty array result".into()));
        }

        // For single-cell results, just set the cached value normally
        if num_rows == 1 && num_cols == 1 {
            let value = array
                .into_iter()
                .next()
                .unwrap()
                .into_iter()
                .next()
                .unwrap();
            return self.set_formula_result(row, col, value);
        }

        // Clear any existing spill from this source
        self.clear_spill(row, col);

        // Check if we can spill
        if !self.cells.can_spill_to(row, col, num_rows, num_cols) {
            // Cannot spill - set the source cell to #SPILL! error
            if let Some(cell) = self.cells.get_mut(row, col) {
                if let CellValue::Formula { cached_value, .. } = &mut cell.value {
                    *cached_value = Some(Box::new(CellValue::Error(crate::CellError::Spill)));
                }
            }
            return Err(Error::Other(
                "Cannot spill: blocked by existing data".into(),
            ));
        }

        // Register the spill source
        self.cells
            .register_spill_source(row, col, crate::cell::SpillInfo::new(num_rows, num_cols));

        // Write the array values
        for (row_offset, row_values) in array.into_iter().enumerate() {
            for (col_offset, value) in row_values.into_iter().enumerate() {
                let target_row = row + row_offset as u32;
                let target_col = col + col_offset as u16;

                if row_offset == 0 && col_offset == 0 {
                    // Source cell - update the formula's cached value and array_result
                    if let Some(cell) = self.cells.get_mut(target_row, target_col) {
                        if let CellValue::Formula {
                            cached_value,
                            array_result,
                            ..
                        } = &mut cell.value
                        {
                            *cached_value = Some(Box::new(value));
                            // Note: We could store the full array here too, but it's
                            // redundant since we have the spill targets
                            *array_result = None;
                        }
                    }
                } else {
                    // Spill target cell
                    let spill_target = CellValue::SpillTarget {
                        source_row: row,
                        source_col: col,
                        offset_row: row_offset as u32,
                        offset_col: col_offset as u16,
                    };
                    self.cells.set(
                        target_row,
                        target_col,
                        crate::cell::CellData::new(spill_target),
                    );

                    // Store the actual value somewhere accessible
                    // For now, SpillTarget cells don't store the value directly -
                    // we'd need to look it up from the source. This is a simplification.
                    // A full implementation would store the value in the SpillTarget or
                    // maintain a separate cache.
                }
            }
        }

        Ok(())
    }

    /// Clear any spill targets from a source formula cell
    ///
    /// Call this before recalculating a formula or when deleting a formula cell.
    pub fn clear_spill(&mut self, row: u32, col: u16) {
        self.cells.clear_spill_targets(row, col);
    }

    /// Check if a cell is a spill target
    pub fn is_spill_target(&self, row: u32, col: u16) -> bool {
        self.cells
            .get(row, col)
            .map(|c| c.value.is_spill_target())
            .unwrap_or(false)
    }

    /// Check if a cell is a spill source (has an array formula that spills)
    pub fn is_spill_source(&self, row: u32, col: u16) -> bool {
        self.cells.is_spill_source(row, col)
    }

    /// Get the source cell coordinates for a spill target
    pub fn get_spill_source(&self, row: u32, col: u16) -> Option<(u32, u16)> {
        self.cells
            .get(row, col)
            .and_then(|c| c.value.spill_source())
    }

    /// Check if a range can be used for spilling
    pub fn can_spill_to(
        &self,
        source_row: u32,
        source_col: u16,
        num_rows: u32,
        num_cols: u16,
    ) -> bool {
        self.cells
            .can_spill_to(source_row, source_col, num_rows, num_cols)
    }
}

/// Freeze pane settings
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FreezePanes {
    /// Freeze row (first unfrozen row)
    pub row: u32,
    /// Freeze column (first unfrozen column)
    pub col: u16,
}

/// Sheet protection settings
#[derive(Debug, Clone, Default)]
pub struct SheetProtection {
    /// Sheet is protected
    pub protected: bool,
    /// Password hash
    pub password_hash: Option<u16>,
    /// Allow selecting locked cells
    pub select_locked_cells: bool,
    /// Allow selecting unlocked cells
    pub select_unlocked_cells: bool,
    /// Allow formatting cells
    pub format_cells: bool,
    /// Allow formatting columns
    pub format_columns: bool,
    /// Allow formatting rows
    pub format_rows: bool,
    /// Allow inserting columns
    pub insert_columns: bool,
    /// Allow inserting rows
    pub insert_rows: bool,
    /// Allow inserting hyperlinks
    pub insert_hyperlinks: bool,
    /// Allow deleting columns
    pub delete_columns: bool,
    /// Allow deleting rows
    pub delete_rows: bool,
    /// Allow sorting
    pub sort: bool,
    /// Allow auto filter
    pub auto_filter: bool,
    /// Allow pivot tables
    pub pivot_tables: bool,
}

/// Page setup for printing
#[derive(Debug, Clone)]
pub struct PageSetup {
    /// Paper size (e.g., 1 = Letter, 9 = A4)
    pub paper_size: u8,
    /// Orientation
    pub orientation: PageOrientation,
    /// Scale percentage (10-400)
    pub scale: u16,
    /// Fit to pages wide
    pub fit_to_width: Option<u16>,
    /// Fit to pages tall
    pub fit_to_height: Option<u16>,
    /// Top margin in inches
    pub top_margin: f64,
    /// Bottom margin in inches
    pub bottom_margin: f64,
    /// Left margin in inches
    pub left_margin: f64,
    /// Right margin in inches
    pub right_margin: f64,
    /// Header margin in inches
    pub header_margin: f64,
    /// Footer margin in inches
    pub footer_margin: f64,
    /// Print gridlines
    pub print_gridlines: bool,
    /// Print headings (row/column headers)
    pub print_headings: bool,
}

impl Default for PageSetup {
    fn default() -> Self {
        Self {
            paper_size: 1, // Letter
            orientation: PageOrientation::Portrait,
            scale: 100,
            fit_to_width: None,
            fit_to_height: None,
            top_margin: 0.75,
            bottom_margin: 0.75,
            left_margin: 0.7,
            right_margin: 0.7,
            header_margin: 0.3,
            footer_margin: 0.3,
            print_gridlines: false,
            print_headings: false,
        }
    }
}

/// Page orientation
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageOrientation {
    #[default]
    Portrait,
    Landscape,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_worksheet() {
        let ws = Worksheet::new("Test");
        assert_eq!(ws.name(), "Test");
        assert!(ws.is_visible());
        assert!(ws.is_empty());
    }

    #[test]
    fn test_set_cell_values() {
        let mut ws = Worksheet::new("Test");

        ws.set_cell_value("A1", "Hello").unwrap();
        ws.set_cell_value("B1", 42.0).unwrap();
        ws.set_cell_value("C1", true).unwrap();

        assert_eq!(ws.get_value("A1").unwrap().as_string(), Some("Hello"));
        assert_eq!(ws.get_value("B1").unwrap().as_number(), Some(42.0));
        assert_eq!(ws.get_value("C1").unwrap().as_bool(), Some(true));
    }

    #[test]
    fn test_set_cell_formula() {
        let mut ws = Worksheet::new("Test");

        ws.set_cell_formula("A1", "=SUM(B1:B10)").unwrap();

        let value = ws.get_value("A1").unwrap();
        assert!(value.is_formula());
        assert_eq!(value.formula_text(), Some("=SUM(B1:B10)"));
    }

    #[test]
    fn test_used_range() {
        let mut ws = Worksheet::new("Test");

        assert!(ws.used_range().is_none());

        ws.set_cell_value_at(5, 3, "A").unwrap();
        ws.set_cell_value_at(10, 7, "B").unwrap();

        let range = ws.used_range().unwrap();
        assert_eq!(range.start.row, 5);
        assert_eq!(range.start.col, 3);
        assert_eq!(range.end.row, 10);
        assert_eq!(range.end.col, 7);
    }

    #[test]
    fn test_row_column_dimensions() {
        let mut ws = Worksheet::new("Test");

        // Default values
        assert!((ws.row_height(0) - 15.0).abs() < 0.001);
        assert!((ws.column_width(0) - 8.43).abs() < 0.001);

        // Custom values
        ws.set_row_height(5, 30.0);
        ws.set_column_width(3, 20.0);

        assert!((ws.row_height(5) - 30.0).abs() < 0.001);
        assert!((ws.column_width(3) - 20.0).abs() < 0.001);
    }

    #[test]
    fn test_merge_cells() {
        let mut ws = Worksheet::new("Test");

        let range = CellRange::parse("A1:C3").unwrap();
        ws.merge_cells(&range).unwrap();

        assert_eq!(ws.merged_regions().len(), 1);

        // Can't merge overlapping
        let range2 = CellRange::parse("B2:D4").unwrap();
        assert!(ws.merge_cells(&range2).is_err());
    }

    #[test]
    fn test_comments() {
        use crate::CellComment;

        let mut ws = Worksheet::new("Test");

        // Initially no comments
        assert_eq!(ws.comment_count(), 0);
        assert!(!ws.has_comment("A1").unwrap());

        // Add a comment
        ws.set_comment("A1", CellComment::new("John", "Review this"))
            .unwrap();
        assert_eq!(ws.comment_count(), 1);
        assert!(ws.has_comment("A1").unwrap());

        // Get the comment
        let comment = ws.comment("A1").unwrap().unwrap();
        assert_eq!(comment.author, "John");
        assert_eq!(comment.text, "Review this");

        // Check authors
        assert_eq!(ws.comment_authors(), &["John"]);

        // Add another comment with same author
        ws.set_comment_at(1, 1, CellComment::new("John", "Another note"));
        assert_eq!(ws.comment_authors().len(), 1); // Should not duplicate

        // Add comment with different author
        ws.set_comment_at(2, 2, CellComment::new("Jane", "My note"));
        assert_eq!(ws.comment_authors().len(), 2);

        // Remove a comment
        let removed = ws.remove_comment("A1").unwrap();
        assert!(removed.is_some());
        assert!(!ws.has_comment("A1").unwrap());
        assert_eq!(ws.comment_count(), 2);

        // Clear all comments
        ws.clear_comments();
        assert_eq!(ws.comment_count(), 0);
        assert!(ws.comment_authors().is_empty());
    }

    #[test]
    fn test_data_validations() {
        use crate::{DataValidation, ValidationOperator};

        let mut ws = Worksheet::new("Test");

        // Initially no validations
        assert_eq!(ws.data_validation_count(), 0);

        // Add a list validation
        let v1 =
            DataValidation::list("Yes,No,Maybe").with_range(CellRange::parse("A1:A10").unwrap());
        ws.add_data_validation(v1);
        assert_eq!(ws.data_validation_count(), 1);

        // Add a number validation
        let v2 = DataValidation::whole_number(ValidationOperator::GreaterThan, "0")
            .with_range(CellRange::parse("B1:B10").unwrap());
        ws.add_data_validation(v2);
        assert_eq!(ws.data_validation_count(), 2);

        // Find validation for specific cell
        let v = ws.data_validation_at(0, 0); // A1
        assert!(v.is_some());

        let v = ws.data_validation_at(0, 1); // B1
        assert!(v.is_some());

        let v = ws.data_validation_at(0, 2); // C1 - no validation
        assert!(v.is_none());

        // Remove validation
        let removed = ws.remove_data_validation(0);
        assert!(removed.is_some());
        assert_eq!(ws.data_validation_count(), 1);

        // Clear all
        ws.clear_data_validations();
        assert_eq!(ws.data_validation_count(), 0);
    }

    #[test]
    fn test_conditional_formatting() {
        use crate::style::{Color, Style};
        use crate::ConditionalFormatRule;

        let mut ws = Worksheet::new("Test");

        // Initially no rules
        assert_eq!(ws.conditional_format_count(), 0);

        // Add a highlight rule
        let rule1 = ConditionalFormatRule::cell_is_greater_than("100")
            .with_range(CellRange::parse("A1:A10").unwrap())
            .with_format(Style::new().fill_color(Color::rgb(255, 199, 206)));
        ws.add_conditional_format(rule1);
        assert_eq!(ws.conditional_format_count(), 1);

        // Add a color scale
        let rule2 = ConditionalFormatRule::color_scale_3(
            Color::rgb(255, 0, 0),
            Color::rgb(255, 255, 0),
            Color::rgb(0, 255, 0),
        )
        .with_range(CellRange::parse("B1:B10").unwrap());
        ws.add_conditional_format(rule2);
        assert_eq!(ws.conditional_format_count(), 2);

        // Find rules for specific cell
        let rules = ws.conditional_formats_at(0, 0); // A1
        assert_eq!(rules.len(), 1);

        let rules = ws.conditional_formats_at(0, 1); // B1
        assert_eq!(rules.len(), 1);

        let rules = ws.conditional_formats_at(0, 2); // C1 - no rules
        assert_eq!(rules.len(), 0);

        // Remove rule
        let removed = ws.remove_conditional_format(0);
        assert!(removed.is_some());
        assert_eq!(ws.conditional_format_count(), 1);

        // Clear all
        ws.clear_conditional_formats();
        assert_eq!(ws.conditional_format_count(), 0);
    }
}
