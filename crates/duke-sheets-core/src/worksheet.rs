//! Worksheet type

use std::collections::HashMap;
use std::sync::RwLock;

use duke_sheets_chart::Chart;
use duke_sheets_chart::ChartEx;
use duke_sheets_chart::DrawingAnchor;
use duke_sheets_chart::EmbeddedImage;

use crate::auto_filter::AutoFilter;
use crate::cell::view::CellView;
use crate::cell::{CellAddress, CellData, CellRange, CellStorage, CellValue, FormulaData};
use crate::comment::CellComment;
use crate::conditional_format::ConditionalFormatRule;
use crate::drawing::{
    anchor_rect_emu_with_metrics, map_child_rect, CommentRef, DrawingKind, DrawingNodeMut,
    DrawingNodeRef, DrawingObject, DrawingPath, GroupChild, Placed, Shape,
};
use crate::error::{Error, Result};
use crate::form_control::{radio_groups, CheckState, FormControl, FormControlKind};
use crate::hyperlink::Hyperlink;
use crate::locale::Locale;
use crate::protection::{hash_legacy_protection_password, ProtectedRange};
use crate::style::Style;
use crate::table::Table;
use crate::validation::DataValidation;
use crate::{MAX_COLS, MAX_ROWS};

/// Sheet visibility state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SheetVisibility {
    /// Sheet tab is visible (default).
    #[default]
    Visible,
    /// Hidden via the UI - users can right-click to unhide.
    Hidden,
    /// Very hidden - only accessible through the VBA editor.
    VeryHidden,
}

/// A worksheet (single sheet in a workbook)
#[derive(Debug)]
pub struct Worksheet {
    /// Sheet name
    name: String,
    /// Cell storage
    cells: CellStorage,
    /// Sheet visibility
    visibility: SheetVisibility,
    /// Sheet is selected
    selected: bool,
    /// Sheet view zoom scale (percent)
    zoom_scale: Option<u16>,
    /// Sheet view selections (one per pane, up to 4 for split views)
    selections: Vec<Selection>,
    /// Sheet protection settings
    protection: Option<SheetProtection>,
    /// Protected editable ranges.
    protected_ranges: Vec<ProtectedRange>,
    /// Freeze pane settings
    freeze_panes: Option<FreezePanes>,
    /// Split pane settings
    split_panes: Option<SplitPanes>,
    /// Print settings
    page_setup: PageSetup,
    /// Tab color
    tab_color: Option<crate::style::Color>,
    /// Cell hyperlinks (keyed by cell address)
    hyperlinks: HashMap<CellAddress, Hyperlink>,
    /// Data validations
    data_validations: Vec<DataValidation>,
    /// Conditional formatting rules
    conditional_formats: Vec<ConditionalFormatRule>,
    /// Tables (ListObjects)
    tables: Vec<Table>,
    /// Drawing objects (images, charts, shapes, form controls,
    /// comments, groups, raw fragments) in z-order, back to front.
    drawings: Vec<DrawingObject>,
    /// Standalone auto-filter (dropdown filter on columns)
    auto_filter: Option<AutoFilter>,
    /// Horizontal page breaks (row breaks)
    row_breaks: Vec<PageBreak>,
    /// Vertical page breaks (column breaks)
    col_breaks: Vec<PageBreak>,
    /// Date system: false = 1900 (Windows default), true = 1904 (Mac legacy).
    /// Copied from WorkbookSettings during reading so cells can format dates.
    date_1904: bool,
    /// Locale for formatting (decimal separators, month names, currency, etc.).
    /// Defaults to en-US. Affects how built-in format IDs render; custom format
    /// strings with `[$-XXXX]` locale prefixes override this per-cell.
    locale: Locale,
    /// Cached ssfmt locale (rebuilt on set_locale).
    ssfmt_locale: ssfmt::Locale,
    /// Mutation generation counter - incremented on user-facing cell/formula edits.
    /// The calculation engine uses this to detect stale caches.
    mutation_count: u64,
    /// Topology generation counter - incremented only when formula/dependency
    /// structure changes. Used to validate cached calc plans across value-only edits.
    topology_generation: u64,
    /// Value-edit ranges since the last successful calculation.
    dirty_value_ranges: Vec<(u32, u16, u32, u16)>,
    /// Image metadata from IMAGE() formulas, populated during calculation.
    /// Behind RwLock so the evaluator can write through a shared &Worksheet reference.
    image_metadata: RwLock<HashMap<(u32, u16), ImageInfo>>,
}

impl Clone for Worksheet {
    fn clone(&self) -> Self {
        Self {
            name: self.name.clone(),
            cells: self.cells.clone(),
            visibility: self.visibility,
            selected: self.selected,
            zoom_scale: self.zoom_scale,
            selections: self.selections.clone(),
            protection: self.protection.clone(),
            protected_ranges: self.protected_ranges.clone(),
            freeze_panes: self.freeze_panes.clone(),
            split_panes: self.split_panes.clone(),
            page_setup: self.page_setup.clone(),
            tab_color: self.tab_color.clone(),
            hyperlinks: self.hyperlinks.clone(),
            data_validations: self.data_validations.clone(),
            conditional_formats: self.conditional_formats.clone(),
            tables: self.tables.clone(),
            drawings: self.drawings.clone(),
            auto_filter: self.auto_filter.clone(),
            row_breaks: self.row_breaks.clone(),
            col_breaks: self.col_breaks.clone(),
            date_1904: self.date_1904,
            locale: self.locale.clone(),
            ssfmt_locale: self.ssfmt_locale.clone(),
            mutation_count: self.mutation_count,
            topology_generation: self.topology_generation,
            dirty_value_ranges: self.dirty_value_ranges.clone(),
            image_metadata: RwLock::new(
                self.image_metadata
                    .read()
                    .unwrap_or_else(|poisoned| poisoned.into_inner())
                    .clone(),
            ),
        }
    }
}

/// Sizing mode for the IMAGE function.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageSizing {
    /// Fit the image inside the cell while preserving aspect ratio.
    FitCell,
    /// Fill the cell, allowing cropping if needed.
    FillCell,
    /// Use the image's original size.
    OriginalSize,
    /// Use an explicit custom width and height.
    Custom,
}

/// Metadata emitted by an IMAGE formula during evaluation.
#[derive(Debug, Clone, PartialEq)]
pub struct ImageInfo {
    /// Source URL or path passed to IMAGE().
    pub source: String,
    /// Alternate text passed to IMAGE().
    pub alt_text: String,
    /// Requested sizing behavior.
    pub sizing: ImageSizing,
    /// Optional custom width.
    pub width: Option<f64>,
    /// Optional custom height.
    pub height: Option<f64>,
}

impl Worksheet {
    /// Create a new worksheet with the given name
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            cells: CellStorage::new(),
            visibility: SheetVisibility::Visible,
            selected: false,
            zoom_scale: None,
            selections: Vec::new(),
            protection: None,
            protected_ranges: Vec::new(),
            freeze_panes: None,
            split_panes: None,
            page_setup: PageSetup::default(),
            tab_color: None,
            hyperlinks: HashMap::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
            tables: Vec::new(),
            drawings: Vec::new(),
            auto_filter: None,
            row_breaks: Vec::new(),
            col_breaks: Vec::new(),
            date_1904: false,
            locale: Locale::en_us(),
            ssfmt_locale: ssfmt::Locale::en_us(),
            mutation_count: 0,
            topology_generation: 0,
            dirty_value_ranges: Vec::new(),
            image_metadata: RwLock::new(HashMap::new()),
        }
    }

    /// Mutation generation counter - incremented on user-facing cell/formula edits.
    /// The calculation engine uses this to detect stale caches.
    pub fn mutation_count(&self) -> u64 {
        self.mutation_count
    }

    /// Topology generation counter - incremented on formula/layout edits that
    /// can change dependency planning, but not on ordinary value-only edits.
    pub fn topology_generation(&self) -> u64 {
        self.topology_generation
    }

    pub fn dirty_value_ranges(&self) -> &[(u32, u16, u32, u16)] {
        &self.dirty_value_ranges
    }

    pub fn clear_dirty_value_ranges(&mut self) {
        self.dirty_value_ranges.clear();
    }

    /// Get image metadata for a cell, if any.
    pub fn get_image_at(&self, row: u32, col: u16) -> Option<ImageInfo> {
        self.image_metadata
            .read()
            .unwrap()
            .get(&(row, col))
            .cloned()
    }

    /// Store image metadata for a cell (called by the evaluator through a shared reference).
    pub fn set_image_at(&self, row: u32, col: u16, info: ImageInfo) {
        self.image_metadata
            .write()
            .unwrap()
            .insert((row, col), info);
    }

    /// Clear all image metadata (called before recalculation).
    pub fn clear_image_metadata(&self) {
        self.image_metadata.write().unwrap().clear();
    }

    /// Get the sheet name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Set the sheet name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Get sheet visibility.
    pub fn visibility(&self) -> SheetVisibility {
        self.visibility
    }

    /// Set sheet visibility.
    pub fn set_visibility(&mut self, visibility: SheetVisibility) {
        self.visibility = visibility;
    }

    /// Check if the sheet is selected
    pub fn is_selected(&self) -> bool {
        self.selected
    }

    /// Set sheet selected state
    pub fn set_selected(&mut self, selected: bool) {
        self.selected = selected;
    }

    /// Get sheet zoom scale in percent.
    pub fn zoom_scale(&self) -> Option<u16> {
        self.zoom_scale
    }

    /// Set sheet zoom scale in percent.
    pub fn set_zoom_scale(&mut self, zoom_scale: Option<u16>) {
        self.zoom_scale = zoom_scale;
    }

    /// Get selected active cell in row/column coordinates.
    ///
    /// Returns the active cell from the primary (last) selection.
    pub fn selection_active_cell(&self) -> Option<(u32, u16)> {
        self.selections.last().and_then(|s| {
            s.active_cell
                .as_ref()
                .and_then(|ac| CellAddress::parse(ac).ok().map(|addr| (addr.row, addr.col)))
        })
    }

    /// Set selected active cell in row/column coordinates.
    ///
    /// Updates the primary selection or creates one if none exist.
    pub fn set_selection_active_cell(&mut self, row: u32, col: u16) {
        let ac = CellAddress::new(row, col).to_a1_string();
        if let Some(sel) = self.selections.last_mut() {
            sel.active_cell = Some(ac);
        } else {
            self.selections.push(Selection {
                pane: None,
                active_cell: Some(ac),
                sqref: None,
            });
        }
    }

    /// Get selected range in sheet view.
    ///
    /// Returns the first range from the primary (last) selection's sqref.
    pub fn selection_range(&self) -> Option<CellRange> {
        self.selections.last().and_then(|s| {
            s.sqref.as_ref().and_then(|sq| {
                sq.split_whitespace()
                    .next()
                    .and_then(|first| CellRange::parse(first).ok())
            })
        })
    }

    /// Set selected range in sheet view.
    ///
    /// Updates the primary selection or creates one if none exist.
    pub fn set_selection_range(&mut self, range: Option<CellRange>) {
        let sqref = range.map(|r| r.to_string());
        if let Some(sel) = self.selections.last_mut() {
            sel.sqref = sqref;
        } else if sqref.is_some() {
            self.selections.push(Selection {
                pane: None,
                active_cell: None,
                sqref,
            });
        }
    }

    /// Get all sheet view selections.
    pub fn selections(&self) -> &[Selection] {
        &self.selections
    }

    /// Set all sheet view selections (replaces existing).
    pub fn set_selections(&mut self, selections: Vec<Selection>) {
        self.selections = selections;
    }

    /// Add a selection to the sheet view.
    pub fn add_selection(&mut self, selection: Selection) {
        self.selections.push(selection);
    }

    /// Get the tab color
    pub fn tab_color(&self) -> Option<crate::style::Color> {
        self.tab_color
    }

    /// Set the tab color
    pub fn set_tab_color(&mut self, color: Option<crate::style::Color>) {
        self.tab_color = color;
    }

    /// Get the sheet protection settings
    pub fn protection(&self) -> Option<&SheetProtection> {
        self.protection.as_ref()
    }

    /// Set sheet protection settings
    pub fn set_protection(&mut self, protection: Option<SheetProtection>) {
        self.protection = protection;
    }

    /// Get protected ranges for this worksheet.
    pub fn protected_ranges(&self) -> &[ProtectedRange] {
        &self.protected_ranges
    }

    /// Set protected ranges for this worksheet.
    pub fn set_protected_ranges(&mut self, protected_ranges: Vec<ProtectedRange>) {
        self.protected_ranges = protected_ranges;
    }

    /// Add one protected range entry.
    pub fn add_protected_range(&mut self, protected_range: ProtectedRange) {
        self.protected_ranges.push(protected_range);
    }

    /// Get the page setup / print settings
    pub fn page_setup(&self) -> &PageSetup {
        &self.page_setup
    }

    /// Set the page setup / print settings
    pub fn set_page_setup(&mut self, page_setup: PageSetup) {
        self.page_setup = page_setup;
    }

    /// Set the print area for this worksheet.
    pub fn set_print_area(&mut self, range: CellRange) {
        let mut ps = self.page_setup.clone();
        ps.print_area = Some(range);
        self.page_setup = ps;
    }

    /// Get the print area for this worksheet.
    pub fn print_area(&self) -> Option<&CellRange> {
        self.page_setup.print_area.as_ref()
    }

    /// Set rows to repeat at top of each printed page (0-based row indices).
    pub fn set_repeat_rows(&mut self, start_row: u32, end_row: u32) {
        let mut ps = self.page_setup.clone();
        ps.repeat_rows = Some((start_row, end_row));
        self.page_setup = ps;
    }

    /// Get rows to repeat at top of each printed page.
    pub fn repeat_rows(&self) -> Option<(u32, u32)> {
        self.page_setup.repeat_rows
    }

    /// Set columns to repeat at left of each printed page (0-based column indices).
    pub fn set_repeat_cols(&mut self, start_col: u16, end_col: u16) {
        let mut ps = self.page_setup.clone();
        ps.repeat_cols = Some((start_col, end_col));
        self.page_setup = ps;
    }

    /// Get columns to repeat at left of each printed page.
    pub fn repeat_cols(&self) -> Option<(u16, u16)> {
        self.page_setup.repeat_cols
    }

    /// Get row breaks (horizontal page breaks).
    pub fn row_breaks(&self) -> &[PageBreak] {
        &self.row_breaks
    }

    /// Get column breaks (vertical page breaks).
    pub fn col_breaks(&self) -> &[PageBreak] {
        &self.col_breaks
    }

    /// Add a manual row break (full-width, after the given 0-based row).
    pub fn add_row_break(&mut self, row: u32) {
        self.row_breaks.push(PageBreak {
            id: row,
            min: 0,
            max: 16383,
            man: true,
            pt: false,
        });
    }

    /// Add a manual column break (full-height, after the given 0-based column).
    pub fn add_col_break(&mut self, col: u32) {
        self.col_breaks.push(PageBreak {
            id: col,
            min: 0,
            max: 1048575,
            man: true,
            pt: false,
        });
    }

    /// Set row breaks (replaces existing).
    pub fn set_row_breaks(&mut self, breaks: Vec<PageBreak>) {
        self.row_breaks = breaks;
    }

    /// Set column breaks (replaces existing).
    pub fn set_col_breaks(&mut self, breaks: Vec<PageBreak>) {
        self.col_breaks = breaks;
    }

    /// Get the date system (false = 1900, true = 1904).
    pub fn date_1904(&self) -> bool {
        self.date_1904
    }

    /// Set the date system (called by readers to propagate from WorkbookSettings).
    pub fn set_date_1904(&mut self, date_1904: bool) {
        self.date_1904 = date_1904;
    }

    /// Get the locale used for cell formatting.
    pub fn locale(&self) -> &Locale {
        &self.locale
    }

    /// Set the locale used for cell formatting.
    pub fn set_locale(&mut self, locale: Locale) {
        self.ssfmt_locale = locale.to_ssfmt();
        self.locale = locale;
    }

    /// Get a cell value by address string (e.g., "A1")
    pub fn cell(&self, address: &str) -> Result<Option<&CellData>> {
        let addr = CellAddress::parse(address)?;
        Ok(self.cells.get(addr.row, addr.col))
    }

    /// Get a cell value by row and column indices
    pub fn cell_at(&self, row: u32, col: u16) -> Option<&CellData> {
        self.cells.get(row, col)
    }

    /// Get a mutable cell by row and column indices.
    ///
    /// Warning: mutating `CellData.value` directly can desynchronize the cell grid
    /// from the formula side table. Prefer `set_cell_value_at`,
    /// `set_cell_formula_at`, or `set_formula_result` for formula cells.
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

    /// Get cell value by indices.
    ///
    /// SpillTarget cells are resolved transparently: the actual spilled value
    /// is returned by looking up the source formula cell's array_result.
    pub fn get_value_at(&self, row: u32, col: u16) -> CellValue {
        let cell = match self.cells.get(row, col) {
            Some(c) => c,
            None => return CellValue::Empty,
        };
        match &cell.value {
            CellValue::SpillTarget {
                source_row,
                source_col,
                offset_row,
                offset_col,
            } => self.resolve_spill_value(*source_row, *source_col, *offset_row, *offset_col),
            other => other.clone(),
        }
    }

    /// Get cell value by indices as a borrowed reference.
    ///
    /// SpillTarget cells are resolved transparently. Missing cells return a
    /// shared `CellValue::Empty` reference.
    pub fn get_value_at_ref(&self, row: u32, col: u16) -> &CellValue {
        static EMPTY: CellValue = CellValue::Empty;
        let Some(cell) = self.cells.get(row, col) else {
            return &EMPTY;
        };
        match &cell.value {
            CellValue::SpillTarget {
                source_row,
                source_col,
                offset_row,
                offset_col,
            } => self
                .resolve_spill_value_ref(*source_row, *source_col, *offset_row, *offset_col)
                .unwrap_or(&EMPTY),
            other => other,
        }
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

    /// Get a [`CellView`] for the cell at the given row and column.
    ///
    /// The view provides access to the cell's value, style, and a
    /// [`formatted()`](CellView::formatted) method that applies the cell's
    /// number format to produce a display string.
    ///
    /// Returns a view with `CellValue::Empty` for cells that don't exist.
    pub fn cell_view_at(&self, row: u32, col: u16) -> CellView<'_> {
        let (value, style) = match self.cells.get(row, col) {
            Some(data) => {
                let style = if data.style_index != 0 {
                    self.cells.style_pool().get(data.style_index)
                } else {
                    None
                };
                (&data.value, style)
            }
            None => (&CellValue::Empty, None),
        };
        CellView::new(value, style, self.date_1904, &self.ssfmt_locale)
    }

    /// Get a [`CellView`] for the cell at the given address string (e.g., "A1").
    pub fn cell_view(&self, address: &str) -> Result<CellView<'_>> {
        let addr = CellAddress::parse(address)?;
        Ok(self.cell_view_at(addr.row, addr.col))
    }

    /// Get the formatted display string for a cell at the given row and column.
    ///
    /// This is a convenience method equivalent to
    /// `self.cell_view_at(row, col).formatted()`.
    ///
    /// Numbers are formatted according to the cell's number format (percentages,
    /// dates, currencies, etc.). Strings, booleans, and errors display as-is.
    pub fn formatted_value_at(&self, row: u32, col: u16) -> String {
        self.cell_view_at(row, col).formatted()
    }

    /// Get the formatted display string for a cell at the given address string.
    pub fn formatted_value(&self, address: &str) -> Result<String> {
        let addr = CellAddress::parse(address)?;
        Ok(self.formatted_value_at(addr.row, addr.col))
    }

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
        let removed_formula = self.remove_formula_state(row, col);
        self.cells.set_value(row, col, value.into());
        self.mutation_count += 1;
        self.dirty_value_ranges.push((row, col, row, col));
        if removed_formula {
            self.topology_generation += 1;
        }
        Ok(())
    }

    /// Set a cell formula by address string
    pub fn set_cell_formula(&mut self, address: &str, formula: &str) -> Result<()> {
        let addr = CellAddress::parse(address)?;
        self.set_cell_formula_at(addr.row, addr.col, formula)
    }

    /// Set a cell formula by row and column indices
    pub fn set_cell_formula_at(&mut self, row: u32, col: u16, formula: &str) -> Result<()> {
        self.set_formula_with_cached_value_at(row, col, formula, CellValue::Empty)
    }

    /// Set a cell formula together with a cached result value.
    ///
    /// This is intended for file readers/importers that already know the cached
    /// value stored alongside the formula. Public callers that only want to set
    /// a formula should use `set_cell_formula_at`, which clears any stale cached
    /// grid value.
    #[doc(hidden)]
    pub fn set_formula_with_cached_value_at(
        &mut self,
        row: u32,
        col: u16,
        formula: &str,
        cached_value: CellValue,
    ) -> Result<()> {
        self.validate_cell_position(row, col)?;

        let formula = Self::normalize_formula_text(formula);

        self.clear_spill(row, col);
        let style_index = self
            .cells
            .get(row, col)
            .map(|cell| cell.style_index)
            .unwrap_or(0);
        self.cells
            .set(row, col, CellData::with_style(cached_value, style_index));
        self.cells.set_formula(row, col, FormulaData::new(formula));
        self.mutation_count += 1;
        self.topology_generation += 1;
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
        self.clear_cell_at(addr.row, addr.col);
        Ok(())
    }

    /// Clear a cell by indices
    pub fn clear_cell_at(&mut self, row: u32, col: u16) {
        let removed_formula = self.remove_formula_state(row, col);
        self.cells.remove(row, col);
        self.mutation_count += 1;
        self.dirty_value_ranges.push((row, col, row, col));
        if removed_formula {
            self.topology_generation += 1;
        }
    }

    /// Get the used range (bounds of all non-empty cells)
    pub fn used_range(&self) -> Option<CellRange> {
        let mut bounds = self.cells.used_bounds();

        for ((row, col), _) in self.cells.iter_formulas() {
            bounds = Some(match bounds {
                Some((min_row, min_col, max_row, max_col)) => (
                    min_row.min(row),
                    min_col.min(col),
                    max_row.max(row),
                    max_col.max(col),
                ),
                None => (row, col, row, col),
            });
        }

        bounds.map(|(min_row, min_col, max_row, max_col)| {
            CellRange::from_indices(min_row, min_col, max_row, max_col)
        })
    }

    /// Returns (row, col) pairs for all non-empty cells in [start_row, end_row] inclusive.
    /// Results are sorted by (row, col). Empty-value cells (style-only) are excluded.
    pub fn populated_cells_in_range(&self, start_row: u32, end_row: u32) -> Vec<(u32, u16)> {
        let mut coords: Vec<(u32, u16)> = self
            .cells
            .cells_map()
            .iter()
            .filter_map(|(&(row, col), data)| {
                if row >= start_row && row <= end_row && !data.value.is_empty() {
                    Some((row, col))
                } else {
                    None
                }
            })
            .collect();
        coords.sort_unstable();
        coords
    }

    /// Returns an iterator over all stored cells (including empty-value styled cells)
    /// in the given row range.
    pub fn cells_map_in_range(
        &self,
        start_row: u32,
        end_row: u32,
    ) -> impl Iterator<Item = (&(u32, u16), &CellData)> {
        self.cells
            .cells_map()
            .iter()
            .filter(move |(&(row, _), _)| row >= start_row && row <= end_row)
    }

    /// Clear all cells in a range
    pub fn clear_range(&mut self, range: &CellRange) {
        let mut removed_formula = false;
        for addr in range.cells() {
            removed_formula |= self.remove_formula_state(addr.row, addr.col);
            self.cells.remove(addr.row, addr.col);
        }
        self.mutation_count += 1;
        self.dirty_value_ranges.push((
            range.start.row,
            range.start.col,
            range.end.row,
            range.end.col,
        ));
        if removed_formula {
            self.topology_generation += 1;
        }
    }

    /// Set the same value for all cells in a range
    pub fn fill_range<V: Into<CellValue> + Clone>(
        &mut self,
        range: &CellRange,
        value: V,
    ) -> Result<()> {
        let value = value.into();
        let mut removed_formula = false;
        for addr in range.cells() {
            self.validate_cell_position(addr.row, addr.col)?;
            removed_formula |= self.remove_formula_state(addr.row, addr.col);
            self.cells.set_value(addr.row, addr.col, value.clone());
        }
        self.mutation_count += 1;
        self.dirty_value_ranges.push((
            range.start.row,
            range.start.col,
            range.end.row,
            range.end.col,
        ));
        if removed_formula {
            self.topology_generation += 1;
        }
        Ok(())
    }

    /// Get row height
    pub fn row_height(&self, row: u32) -> f64 {
        self.cells.row_height(row)
    }

    /// Get default row height in points.
    pub fn default_row_height(&self) -> f64 {
        self.cells.default_row_height()
    }

    /// Set default row height in points.
    pub fn set_default_row_height(&mut self, height: f64) {
        self.cells.set_default_row_height(height);
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

    /// Get row outline/grouping level (0-7)
    pub fn row_outline_level(&self, row: u32) -> u8 {
        self.cells.row_outline_level(row)
    }

    /// Set row outline/grouping level (clamped to 0-7)
    pub fn set_row_outline_level(&mut self, row: u32, level: u8) {
        self.cells.set_row_outline_level(row, level);
    }

    /// Check if row is collapsed in outline
    pub fn is_row_collapsed(&self, row: u32) -> bool {
        self.cells.is_row_collapsed(row)
    }

    /// Set row collapsed state in outline
    pub fn set_row_collapsed(&mut self, row: u32, collapsed: bool) {
        self.cells.set_row_collapsed(row, collapsed);
    }

    /// Get column width
    pub fn column_width(&self, col: u16) -> f64 {
        self.cells.column_width(col)
    }

    /// Get default column width in characters.
    pub fn default_column_width(&self) -> f64 {
        self.cells.default_column_width()
    }

    /// Set default column width in characters.
    pub fn set_default_column_width(&mut self, width: f64) {
        self.cells.set_default_column_width(width);
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

    /// Get column outline/grouping level (0-7)
    pub fn column_outline_level(&self, col: u16) -> u8 {
        self.cells.column_outline_level(col)
    }

    /// Set column outline/grouping level (clamped to 0-7)
    pub fn set_column_outline_level(&mut self, col: u16, level: u8) {
        self.cells.set_column_outline_level(col, level);
    }

    /// Check if column is collapsed in outline
    pub fn is_column_collapsed(&self, col: u16) -> bool {
        self.cells.is_column_collapsed(col)
    }

    /// Set column collapsed state in outline
    pub fn set_column_collapsed(&mut self, col: u16, collapsed: bool) {
        self.cells.set_column_collapsed(col, collapsed);
    }

    /// Get all custom row heights (row index → height in points).
    pub fn custom_row_heights(&self) -> &std::collections::BTreeMap<u32, f64> {
        self.cells.custom_row_heights()
    }

    /// Get all hidden rows (row index → true).
    pub fn hidden_rows(&self) -> &std::collections::BTreeMap<u32, bool> {
        self.cells.hidden_rows()
    }

    /// Get all row outline levels (row index → level).
    pub fn row_outline_levels(&self) -> &std::collections::BTreeMap<u32, u8> {
        self.cells.row_outline_levels()
    }

    /// Get all collapsed rows (row index → true).
    pub fn collapsed_rows(&self) -> &std::collections::BTreeMap<u32, bool> {
        self.cells.row_collapsed()
    }

    /// Get all custom column widths (column index → width in characters).
    pub fn custom_column_widths(&self) -> &std::collections::BTreeMap<u16, f64> {
        self.cells.custom_column_widths()
    }

    /// Get all hidden columns (column index → true).
    pub fn hidden_columns(&self) -> &std::collections::BTreeMap<u16, bool> {
        self.cells.hidden_columns()
    }

    /// Get all column outline levels (column index → level).
    pub fn column_outline_levels(&self) -> &std::collections::BTreeMap<u16, u8> {
        self.cells.column_outline_levels()
    }

    /// Get all collapsed columns (column index → true).
    pub fn collapsed_columns(&self) -> &std::collections::BTreeMap<u16, bool> {
        self.cells.column_collapsed()
    }

    /// Get merged regions
    pub fn merged_regions(&self) -> &[CellRange] {
        self.cells.merged_regions()
    }

    /// Get the merge span for a cell if it is the top-left origin of a merged region.
    ///
    /// Returns `Some((row_span, col_span))` if the cell at `(row, col)` is the
    /// top-left corner of a merged region. Returns `None` if the cell is not
    /// a merge origin (including cells that are "secondary" members of a merge).
    pub fn get_merge_span(&self, row: u32, col: u16) -> Option<(u32, u16)> {
        for region in self.cells.merged_regions() {
            if region.start.row == row && region.start.col == col {
                let row_span = region.end.row - region.start.row + 1;
                let col_span = region.end.col - region.start.col + 1;
                return Some((row_span, col_span));
            }
        }
        None
    }

    /// Check whether a cell is a non-origin ("secondary") member of a merged region.
    ///
    /// Returns `true` if the cell at `(row, col)` is covered by a merged region
    /// but is NOT the top-left origin cell. These cells should typically be
    /// skipped when rendering (the origin cell spans over them).
    pub fn is_merged_secondary(&self, row: u32, col: u16) -> bool {
        let addr = CellAddress::new(row, col);
        for region in self.cells.merged_regions() {
            if region.contains(&addr) {
                // It's in this region - check if it's NOT the origin
                if region.start.row != row || region.start.col != col {
                    return true;
                }
            }
        }
        false
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

    /// Get freeze pane settings
    pub fn freeze_panes(&self) -> Option<&FreezePanes> {
        self.freeze_panes.as_ref()
    }

    /// Get split pane settings
    pub fn split_panes(&self) -> Option<&SplitPanes> {
        self.split_panes.as_ref()
    }

    /// Set freeze panes
    pub fn set_freeze_panes(&mut self, row: u32, col: u16) {
        if row == 0 && col == 0 {
            self.freeze_panes = None;
        } else {
            self.freeze_panes = Some(FreezePanes { row, col });
            self.split_panes = None;
        }
    }

    /// Set split panes
    pub fn set_split_panes(&mut self, split_panes: Option<SplitPanes>) {
        self.split_panes = split_panes;
        if self.split_panes.is_some() {
            self.freeze_panes = None;
        }
    }

    /// Remove freeze panes
    pub fn unfreeze_panes(&mut self) {
        self.freeze_panes = None;
    }

    /// Remove split panes
    pub fn unsplit_panes(&mut self) {
        self.split_panes = None;
    }

    /// Set a hyperlink on a cell by address string.
    pub fn set_hyperlink(&mut self, cell: &str, hyperlink: Hyperlink) -> Result<()> {
        let addr = CellAddress::parse(cell)?;
        self.hyperlinks.insert(addr, hyperlink);
        Ok(())
    }

    /// Get a hyperlink from a cell by address string.
    pub fn hyperlink(&self, cell: &str) -> Option<&Hyperlink> {
        CellAddress::parse(cell)
            .ok()
            .and_then(|addr| self.hyperlinks.get(&addr))
    }

    /// Get a hyperlink from a cell by row and column indices.
    pub fn hyperlink_at(&self, row: u32, col: u16) -> Option<&Hyperlink> {
        let addr = CellAddress::new(row, col);
        self.hyperlinks.get(&addr)
    }

    /// Get a mutable reference to a hyperlink by address string.
    pub fn hyperlink_mut(&mut self, cell: &str) -> Option<&mut Hyperlink> {
        CellAddress::parse(cell)
            .ok()
            .and_then(move |addr| self.hyperlinks.get_mut(&addr))
    }

    /// Get all hyperlinks in this worksheet.
    pub fn hyperlinks(&self) -> &HashMap<CellAddress, Hyperlink> {
        &self.hyperlinks
    }

    /// Get the number of hyperlinks in this worksheet.
    pub fn hyperlink_count(&self) -> usize {
        self.hyperlinks.len()
    }

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
        self.set_comment_at(addr.row, addr.col, comment)
    }

    /// Set a comment on a cell by row and column indices.
    ///
    /// Replacing an existing comment keeps its drawing-list position
    /// and popup anchor; a new comment is appended to the drawing
    /// list with Excel's default popup placement, hidden by default.
    ///
    /// Errors when the cell is off the grid, matching the bounds
    /// validation applied by [`Self::add_drawing`].
    pub fn set_comment_at(&mut self, row: u32, col: u16, comment: CellComment) -> Result<()> {
        if row >= MAX_ROWS {
            return Err(Error::RowOutOfBounds(row, MAX_ROWS - 1));
        }
        if col >= MAX_COLS {
            return Err(Error::ColumnOutOfBounds(col, MAX_COLS - 1));
        }
        for object in &mut self.drawings {
            if let DrawingKind::Comment {
                row: r,
                col: c,
                comment: existing,
            } = &mut object.kind
            {
                if (*r, *c) == (row, col) {
                    *existing = comment;
                    return Ok(());
                }
            }
        }
        self.drawings.push(DrawingObject::comment(row, col, comment));
        Ok(())
    }

    /// Get a comment from a cell by address string
    pub fn comment(&self, address: &str) -> Result<Option<&CellComment>> {
        let addr = CellAddress::parse(address)?;
        Ok(self.comment_at(addr.row, addr.col))
    }

    /// Get a comment from a cell by row and column indices
    pub fn comment_at(&self, row: u32, col: u16) -> Option<&CellComment> {
        self.drawings.iter().find_map(|object| match &object.kind {
            DrawingKind::Comment {
                row: r,
                col: c,
                comment,
            } if (*r, *c) == (row, col) => Some(comment),
            _ => None,
        })
    }

    /// Get a mutable reference to a comment
    pub fn comment_at_mut(&mut self, row: u32, col: u16) -> Option<&mut CellComment> {
        self.drawings
            .iter_mut()
            .find_map(|object| match &mut object.kind {
                DrawingKind::Comment {
                    row: r,
                    col: c,
                    comment,
                } if (*r, *c) == (row, col) => Some(comment),
                _ => None,
            })
    }

    /// Remove a comment from a cell by address string
    pub fn remove_comment(&mut self, address: &str) -> Result<Option<CellComment>> {
        let addr = CellAddress::parse(address)?;
        Ok(self.remove_comment_at(addr.row, addr.col))
    }

    /// Remove a comment from a cell by row and column indices.
    /// Removes the comment's drawing object from the list.
    pub fn remove_comment_at(&mut self, row: u32, col: u16) -> Option<CellComment> {
        let index = self.drawings.iter().position(|object| {
            matches!(
                &object.kind,
                DrawingKind::Comment { row: r, col: c, .. } if (*r, *c) == (row, col)
            )
        })?;
        match self.drawings.remove(index).kind {
            DrawingKind::Comment { comment, .. } => Some(comment),
            _ => unreachable!("position matched a comment"),
        }
    }

    /// Check if a cell has a comment
    pub fn has_comment(&self, address: &str) -> Result<bool> {
        let addr = CellAddress::parse(address)?;
        Ok(self.has_comment_at(addr.row, addr.col))
    }

    /// Check if a cell has a comment by row and column indices
    pub fn has_comment_at(&self, row: u32, col: u16) -> bool {
        self.comment_at(row, col).is_some()
    }

    /// Get the number of comments in this worksheet
    pub fn comment_count(&self) -> usize {
        self.comments().count()
    }

    /// Iterate over all comments in drawing-list (z) order:
    /// ((row, col), comment)
    pub fn comments(&self) -> impl Iterator<Item = ((u32, u16), &CellComment)> {
        self.drawings.iter().filter_map(|object| match &object.kind {
            DrawingKind::Comment { row, col, comment } => Some(((*row, *col), comment)),
            _ => None,
        })
    }

    /// Iterate over all comments with their drawing-list positions
    /// and wrapper objects, in z-order.
    pub fn comments_drawn(&self) -> impl Iterator<Item = CommentRef<'_>> {
        self.drawings
            .iter()
            .enumerate()
            .filter_map(|(index, object)| match &object.kind {
                DrawingKind::Comment { row, col, comment } => Some(CommentRef {
                    index,
                    row: *row,
                    col: *col,
                    object,
                    comment,
                }),
                _ => None,
            })
    }

    /// Unique comment authors in first-appearance (z) order. An empty
    /// author is a distinct entry, so anonymous comments keep their
    /// own author slot instead of being attributed to author 0.
    pub fn comment_authors(&self) -> Vec<String> {
        let mut authors: Vec<String> = Vec::new();
        for (_, comment) in self.comments() {
            if !authors.iter().any(|a| *a == comment.author) {
                authors.push(comment.author.clone());
            }
        }
        authors
    }

    /// Whether a comment's popup is persistently visible.
    /// `None` when the cell has no comment.
    pub fn comment_visible(&self, row: u32, col: u16) -> Option<bool> {
        self.comment_object(row, col)
            .map(|object| !object.meta.hidden)
    }

    /// Show or hide a comment's popup persistently. Returns false
    /// when the cell has no comment.
    pub fn set_comment_visible(&mut self, row: u32, col: u16, visible: bool) -> bool {
        for object in &mut self.drawings {
            if matches!(
                &object.kind,
                DrawingKind::Comment { row: r, col: c, .. } if (*r, *c) == (row, col)
            ) {
                object.meta.hidden = !visible;
                return true;
            }
        }
        false
    }

    /// The drawing object wrapping a cell's comment.
    pub fn comment_object(&self, row: u32, col: u16) -> Option<&DrawingObject> {
        self.drawings.iter().find(|object| {
            matches!(
                &object.kind,
                DrawingKind::Comment { row: r, col: c, .. } if (*r, *c) == (row, col)
            )
        })
    }

    /// Clear all comments from this worksheet
    pub fn clear_comments(&mut self) {
        self.drawings
            .retain(|object| !matches!(object.kind, DrawingKind::Comment { .. }));
    }

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

    /// Add a table to this worksheet.
    pub fn add_table(&mut self, table: Table) {
        self.tables.push(table);
        self.mutation_count += 1;
        self.topology_generation += 1;
    }

    /// Get all tables.
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Get a mutable reference to all tables.
    pub fn tables_mut(&mut self) -> &mut Vec<Table> {
        self.mutation_count += 1;
        self.topology_generation += 1;
        &mut self.tables
    }

    /// Get a table by name.
    pub fn table_by_name(&self, name: &str) -> Option<&Table> {
        self.tables.iter().find(|t| t.name == name)
    }

    /// Get the number of tables.
    pub fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// All drawing objects, in z-order (back to front).
    pub fn drawings(&self) -> &[DrawingObject] {
        &self.drawings
    }

    /// Mutable access to the drawing list. Reordering the list is a
    /// z-order edit. Invariants (anchor bounds, one comment per cell)
    /// are not re-checked through this escape hatch; writers enforce
    /// what their formats require.
    pub fn drawings_mut(&mut self) -> &mut Vec<DrawingObject> {
        &mut self.drawings
    }

    /// Validate and append a drawing object, returning its zero-based
    /// index (= z-position). Rejects a comment for a cell that
    /// already has one anywhere in the drawing tree. Use
    /// [`Self::drawings_mut`] to append without validation, as
    /// readers do to preserve out-of-spec content from existing
    /// files.
    pub fn add_drawing(&mut self, object: DrawingObject) -> Result<usize> {
        object.validate()?;
        self.ensure_comment_cell_free(&object.kind, None)?;
        self.drawings.push(object);
        Ok(self.drawings.len() - 1)
    }

    /// Validate and insert a drawing object at `index`, shifting
    /// later objects up in z-order.
    pub fn insert_drawing(&mut self, index: usize, object: DrawingObject) -> Result<()> {
        if index > self.drawings.len() {
            return Err(Error::other(format!(
                "drawing index {index} out of bounds (count: {})",
                self.drawings.len()
            )));
        }
        object.validate()?;
        self.ensure_comment_cell_free(&object.kind, None)?;
        self.drawings.insert(index, object);
        Ok(())
    }

    /// Validate and replace the top-level drawing at `index`. A
    /// comment may keep its own cell; it only conflicts with comments
    /// elsewhere in the drawing tree.
    pub fn set_drawing(&mut self, index: usize, object: DrawingObject) -> Result<()> {
        if index >= self.drawings.len() {
            return Err(Error::other(format!(
                "drawing index {index} out of bounds (count: {})",
                self.drawings.len()
            )));
        }
        object.validate()?;
        self.ensure_comment_cell_free(&object.kind, Some(&[index]))?;
        self.drawings[index] = object;
        Ok(())
    }

    /// Validate and replace the group child at `path`: a top-level
    /// group index followed by child indices (at least two elements).
    pub fn set_group_child(&mut self, path: &[usize], child: GroupChild) -> Result<()> {
        crate::drawing::validate_group_child(&child)?;
        let (children, index) = self.group_children_mut(path)?;
        if index >= children.len() {
            return Err(Error::other(format!(
                "drawing path {path:?} out of bounds (child count: {})",
                children.len()
            )));
        }
        children[index] = child;
        Ok(())
    }

    /// Remove and return the group child at `path`: a top-level group
    /// index followed by child indices (at least two elements).
    pub fn remove_group_child(&mut self, path: &[usize]) -> Result<GroupChild> {
        let (children, index) = self.group_children_mut(path)?;
        if index >= children.len() {
            return Err(Error::other(format!(
                "drawing path {path:?} out of bounds (child count: {})",
                children.len()
            )));
        }
        Ok(children.remove(index))
    }

    /// Remove and return the drawing object at `index`.
    pub fn remove_drawing(&mut self, index: usize) -> Result<DrawingObject> {
        if index >= self.drawings.len() {
            return Err(Error::other(format!(
                "drawing index {index} out of bounds (count: {})",
                self.drawings.len()
            )));
        }
        Ok(self.drawings.remove(index))
    }

    /// Reject a comment whose cell already has one anywhere in the
    /// drawing tree (group-nested comments only arise from permissive
    /// reads), ignoring drawings at or under `exclude`.
    fn ensure_comment_cell_free(
        &self,
        kind: &DrawingKind,
        exclude: Option<&[usize]>,
    ) -> Result<()> {
        let DrawingKind::Comment { row, col, .. } = kind else {
            return Ok(());
        };
        for (path, node) in self.drawings_flat() {
            if exclude.is_some_and(|prefix| path.starts_with(prefix)) {
                continue;
            }
            if matches!(node.kind, DrawingKind::Comment { row: r, col: c, .. } if (r, c) == (row, col))
            {
                return Err(Error::other(format!(
                    "cell ({row}, {col}) already has a comment"
                )));
            }
        }
        Ok(())
    }

    /// The child list and final index addressed by a group-child
    /// `path` (at least a top-level index plus one child index).
    fn group_children_mut(&mut self, path: &[usize]) -> Result<(&mut Vec<GroupChild>, usize)> {
        let (&child_index, parent_path) = path
            .split_last()
            .ok_or_else(|| Error::other("drawing path cannot be empty"))?;
        if parent_path.is_empty() {
            return Err(Error::other(
                "group child path needs at least two elements (group, then child)",
            ));
        }
        let parent = self
            .drawing_at_path_mut(parent_path)
            .ok_or_else(|| Error::other(format!("no drawing at path {parent_path:?}")))?;
        let DrawingKind::Group(group) = parent.kind else {
            return Err(Error::other("drawing parent is not a group"));
        };
        Ok((&mut group.children, child_index))
    }

    /// Move the drawing object at `from` to position `to`, shifting
    /// objects in between. Both indices refer to current positions.
    pub fn move_drawing(&mut self, from: usize, to: usize) -> Result<()> {
        let count = self.drawings.len();
        if from >= count || to >= count {
            return Err(Error::other(format!(
                "drawing index out of bounds: move {from} -> {to} (count: {count})"
            )));
        }
        let object = self.drawings.remove(from);
        self.drawings.insert(to, object);
        Ok(())
    }

    /// The drawing node at `path`: the first path element indexes the
    /// drawing list, subsequent elements index group children.
    pub fn drawing_at_path(&self, path: &[usize]) -> Option<DrawingNodeRef<'_>> {
        let (&first, rest) = path.split_first()?;
        let object = self.drawings.get(first)?;
        let mut node = DrawingNodeRef {
            meta: &object.meta,
            kind: &object.kind,
        };
        for &index in rest {
            let DrawingKind::Group(group) = node.kind else {
                return None;
            };
            let child = group.children.get(index)?;
            node = DrawingNodeRef {
                meta: &child.meta,
                kind: &child.kind,
            };
        }
        Some(node)
    }

    /// Mutable access to the drawing node at `path`.
    pub fn drawing_at_path_mut(&mut self, path: &[usize]) -> Option<DrawingNodeMut<'_>> {
        let (&first, rest) = path.split_first()?;
        let object = self.drawings.get_mut(first)?;
        let mut node = DrawingNodeMut {
            meta: &mut object.meta,
            kind: &mut object.kind,
        };
        for &index in rest {
            let DrawingKind::Group(group) = node.kind else {
                return None;
            };
            let child = group.children.get_mut(index)?;
            node = DrawingNodeMut {
                meta: &mut child.meta,
                kind: &mut child.kind,
            };
        }
        Some(node)
    }

    /// Depth-first traversal of the drawing tree: top-level objects
    /// in z-order, each group's children before later siblings.
    pub fn drawings_flat(&self) -> Vec<(DrawingPath, DrawingNodeRef<'_>)> {
        fn walk<'a>(
            kind: &'a DrawingKind,
            path: &DrawingPath,
            out: &mut Vec<(DrawingPath, DrawingNodeRef<'a>)>,
        ) {
            if let DrawingKind::Group(group) = kind {
                for (i, child) in group.children.iter().enumerate() {
                    let mut child_path = path.clone();
                    child_path.push(i);
                    out.push((
                        child_path.clone(),
                        DrawingNodeRef {
                            meta: &child.meta,
                            kind: &child.kind,
                        },
                    ));
                    walk(&child.kind, &child_path, out);
                }
            }
        }
        let mut out = Vec::new();
        for (i, object) in self.drawings.iter().enumerate() {
            let path = vec![i];
            out.push((
                path.clone(),
                DrawingNodeRef {
                    meta: &object.meta,
                    kind: &object.kind,
                },
            ));
            walk(&object.kind, &path, &mut out);
        }
        out
    }

    /// Validate and add a chart at the given anchor. Returns the
    /// drawing index.
    pub fn add_chart(&mut self, chart: Chart, anchor: DrawingAnchor) -> Result<usize> {
        self.add_drawing(DrawingObject::chart(chart).with_anchor(anchor))
    }

    /// Every chart in the drawing tree (including inside groups), in
    /// depth-first order, with its path and resolved rectangle.
    pub fn charts(&self) -> impl Iterator<Item = Placed<'_, Chart>> {
        self.placed_nodes(|kind| match kind {
            DrawingKind::Chart(chart) => Some(chart.as_ref()),
            _ => None,
        })
        .into_iter()
    }

    /// Mutable references to top-level chart payloads, in z-order.
    pub fn charts_mut(&mut self) -> impl Iterator<Item = &mut Chart> {
        self.drawings
            .iter_mut()
            .filter_map(|object| match &mut object.kind {
                DrawingKind::Chart(chart) => Some(chart.as_mut()),
                _ => None,
            })
    }

    /// Number of charts anywhere in the drawing tree.
    pub fn chart_count(&self) -> usize {
        self.charts().count()
    }

    /// Validate and add a ChartEx chart at the given anchor. Returns
    /// the drawing index.
    pub fn add_chart_ex(&mut self, chart: ChartEx, anchor: DrawingAnchor) -> Result<usize> {
        self.add_drawing(DrawingObject::chart_ex(chart).with_anchor(anchor))
    }

    /// Every ChartEx chart in the drawing tree (including inside
    /// groups), in depth-first order.
    pub fn charts_ex(&self) -> impl Iterator<Item = Placed<'_, ChartEx>> {
        self.placed_nodes(|kind| match kind {
            DrawingKind::ChartEx(chart) => Some(chart.as_ref()),
            _ => None,
        })
        .into_iter()
    }

    /// Number of ChartEx charts anywhere in the drawing tree.
    pub fn chart_ex_count(&self) -> usize {
        self.charts_ex().count()
    }

    /// Validate and add an embedded image at the given anchor.
    /// Returns the drawing index.
    pub fn add_image(&mut self, image: EmbeddedImage, anchor: DrawingAnchor) -> Result<usize> {
        self.add_drawing(DrawingObject::image(image).with_anchor(anchor))
    }

    /// Every embedded image in the drawing tree (including inside
    /// groups), in depth-first order.
    pub fn images(&self) -> impl Iterator<Item = Placed<'_, EmbeddedImage>> {
        self.placed_nodes(|kind| match kind {
            DrawingKind::Image(image) => Some(image),
            _ => None,
        })
        .into_iter()
    }

    /// Number of embedded images anywhere in the drawing tree.
    pub fn image_count(&self) -> usize {
        self.images().count()
    }

    /// Validate and add a basic worksheet shape at the given anchor.
    /// Returns the drawing index.
    pub fn add_shape(&mut self, shape: Shape, anchor: DrawingAnchor) -> Result<usize> {
        self.add_drawing(DrawingObject::shape(shape).with_anchor(anchor))
    }

    /// Every basic worksheet shape in the drawing tree (including
    /// inside groups), in depth-first order.
    pub fn shapes(&self) -> impl Iterator<Item = Placed<'_, Shape>> {
        self.placed_nodes(|kind| match kind {
            DrawingKind::Shape(shape) => Some(shape.as_ref()),
            _ => None,
        })
        .into_iter()
    }

    /// Number of shapes anywhere in the drawing tree.
    pub fn shape_count(&self) -> usize {
        self.shapes().count()
    }

    /// Validate and add a form control at the given anchor. Returns
    /// the drawing index.
    pub fn add_form_control(&mut self, control: FormControl, anchor: DrawingAnchor) -> Result<usize> {
        self.add_drawing(DrawingObject::form_control(control).with_anchor(anchor))
    }

    /// Every form control in the drawing tree (including inside
    /// groups), in depth-first order, with its path and its absolute
    /// EMU rectangle using this worksheet's row and column metrics.
    pub fn form_controls(&self) -> impl Iterator<Item = Placed<'_, FormControl>> {
        self.placed_nodes(|kind| match kind {
            DrawingKind::FormControl(control) => Some(control),
            _ => None,
        })
        .into_iter()
    }

    /// Number of form controls anywhere in the drawing tree.
    pub fn form_control_count(&self) -> usize {
        self.form_controls().count()
    }

    /// Depth-first typed view over the drawing tree: every node whose
    /// kind projects through `project`, with its path, resolved
    /// rectangle, and metadata.
    fn placed_nodes<T: ?Sized>(
        &self,
        project: fn(&DrawingKind) -> Option<&T>,
    ) -> Vec<Placed<'_, T>> {
        fn walk<'a, T: ?Sized>(
            kind: &'a DrawingKind,
            outer: crate::drawing::CornersEmu,
            path: &DrawingPath,
            project: fn(&DrawingKind) -> Option<&T>,
            out: &mut Vec<Placed<'a, T>>,
        ) {
            if let DrawingKind::Group(group) = kind {
                for (i, child) in group.children.iter().enumerate() {
                    let mut child_path = path.clone();
                    child_path.push(i);
                    let rect = map_child_rect(outer, &group.transform, &child.transform);
                    if let Some(payload) = project(&child.kind) {
                        out.push(Placed {
                            path: child_path.clone(),
                            rect_emu: crate::drawing::RectEmu::from_corners(rect),
                            meta: &child.meta,
                            object: None,
                            payload,
                        });
                    }
                    walk(&child.kind, rect, &child_path, project, out);
                }
            }
        }
        let mut out = Vec::new();
        for (i, object) in self.drawings.iter().enumerate() {
            let path = vec![i];
            let rect = anchor_rect_emu_with_metrics(&object.anchor, self);
            if let Some(payload) = project(&object.kind) {
                out.push(Placed {
                    path: path.clone(),
                    rect_emu: crate::drawing::RectEmu::from_corners(rect),
                    meta: &object.meta,
                    object: Some(object),
                    payload,
                });
            }
            walk(&object.kind, rect, &path, project, &mut out);
        }
        out
    }

    /// The absolute EMU rectangle of the drawing at `path`, using this
    /// worksheet's row and column metrics: the anchor's rectangle for
    /// top-level objects, and the resolved on-sheet rectangle (group
    /// transform applied, rotation/flip aware) for group children.
    /// `None` when no drawing exists at `path`.
    pub fn drawing_rect_emu(&self, path: &[usize]) -> Option<crate::drawing::RectEmu> {
        let (&first, rest) = path.split_first()?;
        let object = self.drawings.get(first)?;
        let mut rect = anchor_rect_emu_with_metrics(&object.anchor, self);
        let mut kind = &object.kind;
        for &index in rest {
            let DrawingKind::Group(group) = kind else {
                return None;
            };
            let child = group.children.get(index)?;
            rect = map_child_rect(rect, &group.transform, &child.transform);
            kind = &child.kind;
        }
        Some(crate::drawing::RectEmu::from_corners(rect))
    }

    /// Mutable access to the form control at `path`.
    pub fn form_control_at_path_mut(&mut self, path: &[usize]) -> Option<&mut FormControl> {
        match self.drawing_at_path_mut(path)?.kind {
            DrawingKind::FormControl(control) => Some(control),
            _ => None,
        }
    }

    /// Apply a checkbox or option-button state change with Excel UI
    /// semantics. Checking an option button unchecks every sibling in
    /// its spatial radio group; unchecking one affects only that button.
    ///
    /// Returns the number of controls whose state changed and the
    /// paths participating in the interaction (the target, plus every
    /// radio-group sibling for option buttons - the group shares its
    /// linked-cell semantics even when only one member changes).
    ///
    /// This changes control state only. Use
    /// [`crate::Workbook::set_form_control_check_state`] when linked
    /// cells should update immediately as part of the interaction.
    pub fn set_form_control_check_state(
        &mut self,
        path: &[usize],
        new_state: CheckState,
    ) -> Result<(usize, Vec<DrawingPath>)> {
        let placed = self.form_controls().collect::<Vec<_>>();
        let target = placed
            .iter()
            .position(|placed| placed.path == path)
            .ok_or_else(|| Error::other(format!("no form control at drawing path {path:?}")))?;

        let is_checkbox = matches!(placed[target].payload.kind, FormControlKind::Checkbox { .. });
        let is_option = matches!(
            placed[target].payload.kind,
            FormControlKind::OptionButton { .. }
        );
        if !is_checkbox && !is_option {
            return Err(Error::other(
                "check state is only valid for checkboxes and option buttons",
            ));
        }
        if is_option && new_state == CheckState::Mixed {
            return Err(Error::other("option buttons cannot use the mixed state"));
        }

        let group: Vec<usize> = if is_option {
            radio_groups(&placed)
                .into_iter()
                .find(|group| group.contains(&target))
                .ok_or_else(|| Error::other("option button has no radio group"))?
        } else {
            vec![target]
        };
        let affected: Vec<DrawingPath> = group
            .iter()
            .map(|&index| placed[index].path.clone())
            .collect();
        let updates: Vec<(DrawingPath, CheckState)> = if is_option
            && new_state == CheckState::Checked
        {
            group
                .iter()
                .map(|&index| {
                    (
                        placed[index].path.clone(),
                        if index == target {
                            CheckState::Checked
                        } else {
                            CheckState::Unchecked
                        },
                    )
                })
                .collect()
        } else {
            vec![(path.to_vec(), new_state)]
        };
        drop(placed);

        let mut changed = 0;
        for (path, state) in updates {
            let control = self
                .form_control_at_path_mut(&path)
                .ok_or_else(|| Error::other("form control disappeared during state update"))?;
            let current = match &mut control.kind {
                FormControlKind::Checkbox { state, .. }
                | FormControlKind::OptionButton { state, .. } => state,
                _ => continue,
            };
            if *current != state {
                *current = state;
                changed += 1;
            }
        }
        Ok((changed, affected))
    }

    /// Set the standalone auto-filter for this worksheet.
    pub fn set_auto_filter(&mut self, auto_filter: Option<AutoFilter>) {
        self.auto_filter = auto_filter;
    }

    /// Get the standalone auto-filter.
    pub fn auto_filter(&self) -> Option<&AutoFilter> {
        self.auto_filter.as_ref()
    }

    /// Get a mutable reference to the auto-filter.
    pub fn auto_filter_mut(&mut self) -> &mut Option<AutoFilter> {
        &mut self.auto_filter
    }

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

    fn remove_formula_state(&mut self, row: u32, col: u16) -> bool {
        if self.cells.has_formula(row, col) {
            self.clear_spill(row, col);
            self.cells.remove_formula(row, col);
            true
        } else {
            false
        }
    }

    fn normalize_formula_text(formula: &str) -> String {
        if formula.starts_with('=') || formula.starts_with("{=") {
            formula.to_string()
        } else if formula.is_empty() {
            String::new()
        } else {
            format!("={}", formula)
        }
    }

    /// Get the number of non-empty cells
    pub fn cell_count(&self) -> usize {
        self.cells.cell_count()
            + self
                .cells
                .iter_formulas()
                .filter(|((row, col), _)| self.cells.get(*row, *col).is_none())
                .count()
    }

    /// Check if the worksheet is empty
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty() && self.cells.formula_count() == 0
    }

    /// Iterate over all non-empty cells
    pub fn iter_cells(&self) -> impl Iterator<Item = (u32, u16, &CellData)> {
        self.cells.iter()
    }

    /// Iterate over all formula cells: (row, col, formula_text)
    pub fn formula_cells(&self) -> impl Iterator<Item = (u32, u16, &str)> {
        self.cells
            .iter_formulas()
            .map(|((row, col), formula)| (row, col, formula.text.as_str()))
    }

    pub fn has_formula_at(&self, row: u32, col: u16) -> bool {
        self.cells.has_formula(row, col)
    }

    pub fn formula_data_at(&self, row: u32, col: u16) -> Option<&FormulaData> {
        self.cells.get_formula(row, col)
    }

    pub fn formula_data_at_mut(&mut self, row: u32, col: u16) -> Option<&mut FormulaData> {
        self.cells.get_formula_mut(row, col)
    }

    /// Get the formula text at a cell position (if it's a formula)
    pub fn get_formula_at(&self, row: u32, col: u16) -> Option<&str> {
        self.cells
            .get_formula(row, col)
            .map(|formula| formula.text.as_str())
    }

    /// Set the cached result value of a formula cell
    /// Returns Ok(()) if the cell is a formula and was updated,
    /// or an error if the cell doesn't exist or isn't a formula
    pub fn set_formula_result(&mut self, row: u32, col: u16, value: CellValue) -> Result<()> {
        if !self.cells.has_formula(row, col) {
            return Err(Error::InvalidAddress(format!(
                "Cell at ({}, {}) is not a formula",
                row, col
            )));
        }

        let style_index = self
            .cells
            .get(row, col)
            .map(|cell| cell.style_index)
            .unwrap_or(0);
        self.cells
            .set(row, col, CellData::with_style(value, style_index));
        if let Some(formula) = self.cells.get_formula_mut(row, col) {
            formula.array_result = None;
        }
        self.mutation_count += 1;
        self.dirty_value_ranges.push((row, col, row, col));
        Ok(())
    }

    /// Get the cached value of a formula cell, or the cell value directly if not a formula.
    ///
    /// SpillTarget cells are resolved: returns the actual spilled value by reference.
    pub fn get_calculated_value_at(&self, row: u32, col: u16) -> Option<&CellValue> {
        let Some(cell) = self.cells.get(row, col) else {
            return if self.cells.has_formula(row, col) {
                Some(&CellValue::Empty)
            } else {
                None
            };
        };

        match &cell.value {
            CellValue::SpillTarget {
                source_row,
                source_col,
                offset_row,
                offset_col,
            } => self.resolve_spill_value_ref(*source_row, *source_col, *offset_row, *offset_col),
            other => Some(other),
        }
    }

    /// Resolve a SpillTarget to its actual value (owned).
    ///
    /// Looks up the source formula cell's array_result and returns the value
    /// at the given offset. Returns Empty if the source cell or array is missing.
    fn resolve_spill_value(
        &self,
        source_row: u32,
        source_col: u16,
        offset_row: u32,
        offset_col: u16,
    ) -> CellValue {
        self.cells
            .get_formula(source_row, source_col)
            .and_then(|formula| formula.array_result.as_ref())
            .and_then(|arr| {
                arr.get(offset_row as usize)
                    .and_then(|r| r.get(offset_col as usize))
                    .cloned()
            })
            .unwrap_or(CellValue::Empty)
    }

    /// Resolve a SpillTarget to its actual value (by reference).
    ///
    /// Returns a reference into the source formula cell's array_result.
    fn resolve_spill_value_ref(
        &self,
        source_row: u32,
        source_col: u16,
        offset_row: u32,
        offset_col: u16,
    ) -> Option<&CellValue> {
        self.cells
            .get_formula(source_row, source_col)
            .and_then(|formula| formula.array_result.as_ref())
            .and_then(|arr| {
                arr.get(offset_row as usize)
                    .and_then(|r| r.get(offset_col as usize))
            })
    }

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
        if !self.cells.has_formula(row, col) {
            return Err(Error::InvalidAddress(format!(
                "Cell at ({}, {}) is not a formula",
                row, col
            )));
        }

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
            let _ = self.set_formula_result(row, col, CellValue::Error(crate::CellError::Spill));
            return Err(Error::Other(
                "Cannot spill: blocked by existing data".into(),
            ));
        }

        // Register the spill source
        self.cells
            .register_spill_source(row, col, crate::cell::SpillInfo::new(num_rows, num_cols));

        // Store the full array in the source cell's array_result for SpillTarget resolution.
        // SpillTarget cells only store coordinates; actual values are resolved through
        // the source cell's array_result via get_value_at() / get_calculated_value_at().
        let top_left_value = array[0][0].clone();
        let style_index = self
            .cells
            .get(row, col)
            .map(|cell| cell.style_index)
            .unwrap_or(0);
        self.cells.set(
            row,
            col,
            crate::cell::CellData::with_style(top_left_value, style_index),
        );
        if let Some(formula) = self.cells.get_formula_mut(row, col) {
            formula.array_result = Some(array);
        }

        // Write SpillTarget cells for all non-anchor positions
        for row_offset in 0..num_rows {
            for col_offset in 0..num_cols as u32 {
                if row_offset == 0 && col_offset == 0 {
                    continue; // anchor cell already handled above
                }
                let target_row = row + row_offset;
                let target_col = col + col_offset as u16;
                let spill_target = CellValue::SpillTarget {
                    source_row: row,
                    source_col: col,
                    offset_row: row_offset,
                    offset_col: col_offset as u16,
                };
                self.cells.set(
                    target_row,
                    target_col,
                    crate::cell::CellData::new(spill_target),
                );
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

    /// Get the spill info (dimensions) for a spill source cell.
    ///
    /// Returns None if the cell is not a spill source.
    pub fn get_spill_info(&self, row: u32, col: u16) -> Option<&crate::cell::SpillInfo> {
        self.cells.get_spill_info(row, col)
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

/// Hidden rows and columns render at zero extent, so they contribute
/// no width/height and no position advance.
impl duke_sheets_chart::DrawingMetrics for Worksheet {
    fn column_width_emu(&self, col: u16) -> i64 {
        if self.is_column_hidden(col) {
            return 0;
        }
        duke_sheets_chart::column_width_to_emu(self.column_width(col))
    }

    fn row_height_emu(&self, row: u32) -> i64 {
        if self.is_row_hidden(row) {
            return 0;
        }
        duke_sheets_chart::row_height_to_emu(self.row_height(row))
    }

    fn column_position_emu(&self, col: u16) -> i128 {
        let default = duke_sheets_chart::column_width_to_emu(self.default_column_width());
        let mut position = i128::from(col) * i128::from(default);
        for (_, hidden) in self.hidden_columns().range(..col) {
            if *hidden {
                position -= i128::from(default);
            }
        }
        for (index, width) in self.custom_column_widths().range(..col) {
            if !self.is_column_hidden(*index) {
                position += i128::from(duke_sheets_chart::column_width_to_emu(*width) - default);
            }
        }
        position
    }

    fn row_position_emu(&self, row: u32) -> i128 {
        let default = duke_sheets_chart::row_height_to_emu(self.default_row_height());
        let mut position = i128::from(row) * i128::from(default);
        for (_, hidden) in self.hidden_rows().range(..row) {
            if *hidden {
                position -= i128::from(default);
            }
        }
        for (index, height) in self.custom_row_heights().range(..row) {
            if !self.is_row_hidden(*index) {
                position += i128::from(duke_sheets_chart::row_height_to_emu(*height) - default);
            }
        }
        position
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

/// Split pane settings
#[derive(Debug, Clone, PartialEq)]
pub struct SplitPanes {
    /// Horizontal split position
    pub x_split: f64,
    /// Vertical split position
    pub y_split: f64,
    /// Top-left visible cell after split
    pub top_left: Option<(u32, u16)>,
    /// Active pane identifier (`topLeft`, `topRight`, `bottomLeft`, `bottomRight`)
    pub active_pane: Option<String>,
}

/// A selection within a sheet view.
///
/// Each `<selection>` element in OOXML represents a selected range in one pane.
/// A sheet view can have up to 4 selections (one per pane: topLeft, topRight,
/// bottomLeft, bottomRight). The `sqref` may contain multiple space-separated
/// ranges for non-contiguous selections.
#[derive(Debug, Clone, PartialEq)]
pub struct Selection {
    /// Which pane this selection belongs to (e.g., "bottomRight").
    /// `None` for the default pane (when no pane split exists).
    pub pane: Option<String>,
    /// The active cell (cursor position) within this selection.
    pub active_cell: Option<String>,
    /// Space-separated range references (e.g., "A1:B2 D4:E5").
    pub sqref: Option<String>,
}

/// Sheet protection settings
#[derive(Debug, Clone, Default, PartialEq, Eq)]
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

impl SheetProtection {
    /// Create sheet protection with the sheet protected.
    pub fn protected() -> Self {
        Self {
            protected: true,
            select_locked_cells: true,
            select_unlocked_cells: true,
            ..Default::default()
        }
    }

    /// Set the password from plaintext input, storing only the legacy verifier.
    pub fn with_password(mut self, password: &str) -> Self {
        self.password_hash = Some(hash_legacy_protection_password(password));
        self
    }

    /// Set a precomputed legacy password verifier.
    pub fn with_password_hash(mut self, password_hash: u16) -> Self {
        self.password_hash = Some(password_hash);
        self
    }
}

#[cfg(test)]
mod protection_tests {
    use super::*;

    #[test]
    fn protected_builder_allows_selection_by_default() {
        let protection = SheetProtection::protected();
        assert!(protection.protected);
        assert!(protection.select_locked_cells);
        assert!(protection.select_unlocked_cells);
    }
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
    /// Odd page header text (used for all pages when differentOddEven is false)
    pub odd_header: Option<String>,
    /// Odd page footer text (used for all pages when differentOddEven is false)
    pub odd_footer: Option<String>,
    /// Even page header text (only used when different_odd_even is true)
    pub even_header: Option<String>,
    /// Even page footer text (only used when different_odd_even is true)
    pub even_footer: Option<String>,
    /// First page header text (only used when different_first is true)
    pub first_header: Option<String>,
    /// First page footer text (only used when different_first is true)
    pub first_footer: Option<String>,
    /// Use different headers/footers for odd and even pages
    pub different_odd_even: bool,
    /// Use a different header/footer for the first page
    pub different_first: bool,
    /// Scale header/footer with document scaling (default: true)
    pub scale_with_doc: bool,
    /// Align header/footer margins with page margins (default: true)
    pub align_with_margins: bool,
    /// Print area (the range that will be printed)
    pub print_area: Option<CellRange>,
    /// Repeat rows at top of each printed page (start_row, end_row), 0-based
    pub repeat_rows: Option<(u32, u32)>,
    /// Repeat columns at left of each printed page (start_col, end_col), 0-based
    pub repeat_cols: Option<(u16, u16)>,
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
            odd_header: None,
            odd_footer: None,
            even_header: None,
            even_footer: None,
            first_header: None,
            first_footer: None,
            different_odd_even: false,
            different_first: false,
            scale_with_doc: true,
            align_with_margins: true,
            print_area: None,
            repeat_rows: None,
            repeat_cols: None,
        }
    }
}

/// A manual page break (row or column).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageBreak {
    /// Row index (for row breaks) or column index (for col breaks).
    /// This is the 0-based index; the XLSX format uses 1-based `id` attributes.
    pub id: u32,
    /// Minimum column (row breaks) or row (col breaks), 0-based. Default 0.
    pub min: u32,
    /// Maximum column (row breaks) or row (col breaks), 0-based.
    /// Row breaks: 16383 for full-width. Col breaks: 1048575 for full-height.
    pub max: u32,
    /// Whether this is a manual break (true) or automatic (false).
    pub man: bool,
    /// Whether this break was created by a PivotTable.
    pub pt: bool,
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
    fn drawing_rect_emu_matches_form_controls() {
        use crate::drawing::{ChildTransform, Group, GroupChild, GroupTransform};
        use crate::drawing::{DrawingMeta, DrawingObject};
        use crate::{CheckState, FormControl, FormControlKind};
        use duke_sheets_chart::{CellMarker, DrawingAnchor};

        let checkbox = || {
            FormControl::new(FormControlKind::Checkbox {
                caption: "cb".into(),
                state: CheckState::Unchecked,
                cell_link: None,
                no_3d: false,
            })
        };

        let mut ws = Worksheet::new("Test");
        ws.add_form_control(
            checkbox(),
            DrawingAnchor::TwoCell {
                from: CellMarker {
                    col: 1,
                    col_offset_emu: 1000,
                    row: 1,
                    row_offset_emu: 2000,
                },
                to: CellMarker {
                    col: 3,
                    col_offset_emu: 0,
                    row: 4,
                    row_offset_emu: 0,
                },
                edit_as: None,
            },
        ).unwrap();
        ws.add_drawing(DrawingObject::group(Group {
            transform: GroupTransform {
                x_emu: 100_000,
                y_emu: 50_000,
                cx_emu: 400_000,
                cy_emu: 200_000,
                child_x_emu: 0,
                child_y_emu: 0,
                child_cx_emu: 800_000,
                child_cy_emu: 400_000,
                ..GroupTransform::default()
            },
            children: vec![GroupChild {
                meta: DrawingMeta::default(),
                transform: ChildTransform {
                    x_emu: 200_000,
                    y_emu: 100_000,
                    cx_emu: 400_000,
                    cy_emu: 200_000,
                    rotation: 45,
                    ..ChildTransform::default()
                },
                kind: DrawingKind::FormControl(checkbox()),
            }],
        })).unwrap();

        let placed = ws.form_controls().collect::<Vec<_>>();
        assert_eq!(placed.len(), 2);
        for control in &placed {
            assert_eq!(
                ws.drawing_rect_emu(&control.path),
                Some(control.rect_emu),
                "path {:?} must resolve to the placed rectangle",
                control.path
            );
        }
        // The group child's resolved rect is scaled by the group frame,
        // not the raw child transform.
        assert_ne!(
            ws.drawing_rect_emu(&placed[1].path).unwrap(),
            crate::drawing::RectEmu {
                x_emu: 200_000,
                y_emu: 100_000,
                width_emu: 400_000,
                height_emu: 200_000,
            },
        );
        assert_eq!(ws.drawing_rect_emu(&[7]), None);
        assert_eq!(ws.drawing_rect_emu(&[1, 3]), None);
        assert_eq!(ws.drawing_rect_emu(&[]), None);
    }

    fn drawing_test_button() -> crate::drawing::DrawingObject {
        use crate::{FormControl, FormControlKind};
        crate::drawing::DrawingObject::form_control(FormControl::new(FormControlKind::Button {
            caption: "b".into(),
        }))
    }

    fn reversed_anchor_object() -> crate::drawing::DrawingObject {
        use duke_sheets_chart::{CellMarker, DrawingAnchor};
        drawing_test_button().with_anchor(DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 4,
                col_offset_emu: 0,
                row: 4,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 1,
                row_offset_emu: 0,
            },
            edit_as: None,
        })
    }

    fn group_with_button() -> crate::drawing::DrawingObject {
        use crate::drawing::{
            ChildTransform, DrawingMeta, DrawingObject, Group, GroupChild, GroupTransform,
        };
        DrawingObject::group(Group {
            transform: GroupTransform::default(),
            children: vec![GroupChild {
                meta: DrawingMeta::default(),
                transform: ChildTransform::default(),
                kind: drawing_test_button().kind,
            }],
        })
    }

    #[test]
    fn kind_views_are_recursive_with_paths_and_rects() {
        use crate::drawing::{
            ChildTransform, DrawingKind, DrawingMeta, DrawingObject, Group, GroupChild,
            GroupTransform,
        };
        use duke_sheets_chart::{EmbeddedImage, ImageFormat};

        let image = || EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 100,
            height_emu: 100,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: vec![1, 2, 3],
            svg_data: None,
        };
        let mut ws = Worksheet::new("Test");
        ws.add_drawing(DrawingObject::image(image())).unwrap();
        ws.add_drawing(DrawingObject::group(Group {
            transform: GroupTransform::default(),
            children: vec![
                GroupChild {
                    meta: DrawingMeta {
                        name: Some("nested image".to_string()),
                        ..DrawingMeta::default()
                    },
                    transform: ChildTransform::default(),
                    kind: DrawingKind::Image(image()),
                },
                GroupChild {
                    meta: DrawingMeta::default(),
                    transform: ChildTransform::default(),
                    kind: drawing_test_button().kind,
                },
            ],
        }))
        .unwrap();

        let images: Vec<_> = ws.images().collect();
        assert_eq!(ws.image_count(), 2, "counts include group children");
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].path, vec![0]);
        assert!(
            images[0].object.is_some(),
            "top-level nodes expose their wrapper object"
        );
        assert_eq!(images[1].path, vec![1, 0]);
        assert_eq!(images[1].meta.name.as_deref(), Some("nested image"));
        assert!(images[1].object.is_none(), "group children have no wrapper");
        assert_eq!(
            Some(images[1].rect_emu),
            ws.drawing_rect_emu(&[1, 0]),
            "view rects match the per-path resolution"
        );

        let controls: Vec<_> = ws.form_controls().collect();
        assert_eq!(ws.form_control_count(), 1);
        assert_eq!(controls[0].path, vec![1, 1]);
        assert_eq!(ws.chart_count(), 0);
        assert!(ws.charts().next().is_none());
    }

    #[test]
    fn add_drawing_validates_and_enforces_comment_uniqueness() {
        use crate::drawing::{DrawingMeta, DrawingObject, Group, GroupChild};
        use crate::CellComment;

        let mut ws = Worksheet::new("Test");
        assert!(ws.add_drawing(reversed_anchor_object()).is_err());
        assert!(ws.drawings().is_empty());

        let comment = |row, col| {
            DrawingObject::comment(row, col, CellComment::new("a", "t"))
                .with_anchor(crate::drawing::default_comment_anchor(row, col))
        };
        assert_eq!(ws.add_drawing(comment(1, 1)).unwrap(), 0);
        assert!(ws
            .add_drawing(comment(1, 1))
            .unwrap_err()
            .to_string()
            .contains("already has a comment"));

        // A group-nested comment (only permissive reads produce them)
        // still blocks its cell.
        let hostile = DrawingObject::group(Group {
            transform: Default::default(),
            children: vec![GroupChild {
                meta: DrawingMeta::default(),
                transform: Default::default(),
                kind: comment(5, 5).kind,
            }],
        });
        ws.drawings_mut().push(hostile);
        assert!(ws.add_drawing(comment(5, 5)).is_err());
        assert!(ws.insert_drawing(0, comment(5, 5)).is_err());
    }

    #[test]
    fn set_drawing_replaces_top_level_with_validation() {
        use crate::drawing::{DrawingKind, DrawingObject};
        use crate::CellComment;

        let mut ws = Worksheet::new("Test");
        let comment = |row, col| {
            DrawingObject::comment(row, col, CellComment::new("a", "t"))
                .with_anchor(crate::drawing::default_comment_anchor(row, col))
        };
        ws.add_drawing(comment(1, 1)).unwrap();
        ws.add_drawing(comment(2, 2)).unwrap();

        assert!(ws.set_drawing(5, drawing_test_button()).is_err());
        assert!(ws.set_drawing(0, reversed_anchor_object()).is_err());
        // Conflicts with the *other* comment's cell.
        assert!(ws.set_drawing(0, comment(2, 2)).is_err());
        // Replacing a comment with one on the same cell excludes itself.
        ws.set_drawing(0, comment(1, 1)).unwrap();
        // Moving the comment to a free cell works.
        ws.set_drawing(0, comment(3, 3)).unwrap();
        let DrawingKind::Comment { row, col, .. } = ws.drawings()[0].kind else {
            panic!("expected a comment at index 0");
        };
        assert_eq!((row, col), (3, 3));
    }

    #[test]
    fn set_group_child_replaces_nested_children_with_validation() {
        use crate::drawing::{
            ChildTransform, DrawingKind, DrawingMeta, GroupChild, RawDrawing,
        };

        let mut ws = Worksheet::new("Test");
        ws.add_drawing(group_with_button()).unwrap();

        let group_child = || GroupChild {
            meta: DrawingMeta::default(),
            transform: ChildTransform::default(),
            kind: DrawingKind::Group(Box::new(crate::drawing::Group::default())),
        };
        ws.set_group_child(&[0, 0], group_child()).unwrap();
        assert!(matches!(
            ws.drawing_at_path(&[0, 0]).unwrap().kind,
            DrawingKind::Group(_)
        ));

        // Raw payloads cannot be group children.
        let raw_child = GroupChild {
            meta: DrawingMeta::default(),
            transform: ChildTransform::default(),
            kind: DrawingKind::Raw(RawDrawing::default()),
        };
        assert!(ws.set_group_child(&[0, 0], raw_child).is_err());

        assert!(ws.set_group_child(&[0, 7], group_child()).is_err());
        assert!(ws.set_group_child(&[3, 0], group_child()).is_err());
        assert!(ws.set_group_child(&[0], group_child()).is_err());
        // Parent is not a group.
        ws.add_drawing(drawing_test_button()).unwrap();
        assert!(ws.set_group_child(&[1, 0], group_child()).is_err());
    }

    #[test]
    fn remove_group_child_removes_nested_children() {
        use crate::drawing::DrawingKind;

        let mut ws = Worksheet::new("Test");
        ws.add_drawing(group_with_button()).unwrap();

        assert!(ws.remove_group_child(&[0, 7]).is_err());
        assert!(ws.remove_group_child(&[0]).is_err());
        let removed = ws.remove_group_child(&[0, 0]).unwrap();
        assert!(matches!(removed.kind, DrawingKind::FormControl(_)));
        let DrawingKind::Group(group) = &ws.drawings()[0].kind else {
            panic!("expected the group to remain");
        };
        assert!(group.children.is_empty());
    }

    #[test]
    fn test_new_worksheet() {
        let ws = Worksheet::new("Test");
        assert_eq!(ws.name(), "Test");
        assert_eq!(ws.visibility(), SheetVisibility::Visible);
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
        assert_eq!(value, CellValue::Empty);
        assert_eq!(ws.get_formula_at(0, 0), Some("=SUM(B1:B10)"));
    }

    #[test]
    fn test_set_cell_formula_clears_stale_cached_value() {
        let mut ws = Worksheet::new("Test");

        ws.set_cell_value("A1", 42.0).unwrap();
        ws.set_cell_formula("A1", "=1+1").unwrap();

        assert_eq!(ws.get_value("A1").unwrap(), CellValue::Empty);
        assert_eq!(ws.get_formula_at(0, 0), Some("=1+1"));
    }

    #[test]
    fn test_set_formula_with_cached_value_replaces_formula_and_value_atomically() {
        let mut ws = Worksheet::new("Test");

        ws.set_formula_with_cached_value_at(0, 0, "=1+1", CellValue::Number(2.0))
            .unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Number(2.0));
        assert_eq!(ws.get_formula_at(0, 0), Some("=1+1"));

        ws.set_formula_with_cached_value_at(0, 0, "=2+2", CellValue::Number(4.0))
            .unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Number(4.0));
        assert_eq!(ws.get_formula_at(0, 0), Some("=2+2"));
    }

    #[test]
    fn test_formula_only_cells_count_toward_used_range_and_cell_count() {
        let mut ws = Worksheet::new("Test");

        ws.set_cell_formula("C5", "=1+1").unwrap();

        let range = ws.used_range().unwrap();
        assert_eq!(range.start.row, 4);
        assert_eq!(range.start.col, 2);
        assert_eq!(range.end.row, 4);
        assert_eq!(range.end.col, 2);
        assert_eq!(ws.cell_count(), 1);
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
    fn test_populated_cells_in_range_skips_style_only_cells() {
        let mut ws = Worksheet::new("Test");

        ws.set_cell_value_at(2, 5, "A").unwrap();
        ws.set_cell_style_at(3, 1, &Style::new().bold(true))
            .unwrap();
        ws.set_cell_formula_at(3, 2, "=1+1").unwrap();
        ws.set_formula_with_cached_value_at(4, 1, "=2+2", CellValue::Number(4.0))
            .unwrap();
        ws.set_cell_value_at(4, 3, "B").unwrap();

        assert_eq!(
            ws.populated_cells_in_range(2, 4),
            vec![(2, 5), (4, 1), (4, 3)]
        );
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
        assert_eq!(comment.plain_text(), "Review this");

        // Check authors
        assert_eq!(ws.comment_authors(), &["John"]);

        // Add another comment with same author
        ws.set_comment_at(1, 1, CellComment::new("John", "Another note"))
            .unwrap();
        assert_eq!(ws.comment_authors().len(), 1); // Should not duplicate

        // Add comment with different author
        ws.set_comment_at(2, 2, CellComment::new("Jane", "My note"))
            .unwrap();
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
    fn set_comment_at_rejects_out_of_bounds_cells() {
        use crate::CellComment;

        let mut ws = Worksheet::new("Test");
        assert!(ws
            .set_comment_at(MAX_ROWS, 0, CellComment::new("a", "row off grid"))
            .is_err());
        assert!(ws
            .set_comment_at(0, MAX_COLS, CellComment::new("a", "col off grid"))
            .is_err());
        assert_eq!(ws.comment_count(), 0, "off-grid comments must be rejected");

        // The last grid cell is still valid.
        ws.set_comment_at(MAX_ROWS - 1, MAX_COLS - 1, CellComment::new("a", "edge"))
            .unwrap();
        assert_eq!(ws.comment_count(), 1);
    }

    #[test]
    fn test_hyperlinks() {
        use crate::Hyperlink;

        let mut ws = Worksheet::new("Test");
        let link = Hyperlink {
            target: "https://example.com".to_string(),
            display: Some("Example".to_string()),
            tooltip: Some("Visit site".to_string()),
            location: None,
        };

        ws.set_hyperlink("A1", link.clone()).unwrap();
        assert_eq!(ws.hyperlink_count(), 1);
        assert_eq!(ws.hyperlink("A1"), Some(&link));
        assert!(ws.hyperlink("B2").is_none());
        assert!(ws.set_hyperlink("", link).is_err());
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
