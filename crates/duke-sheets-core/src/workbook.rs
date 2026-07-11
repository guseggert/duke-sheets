//! Workbook type - the main document structure

use std::any::Any;
use std::collections::BTreeMap;
use std::sync::atomic::{AtomicU64, Ordering};

use crate::cell::{CellAddress, CellError, CellValue};
use crate::error::{Error, Result};
use crate::form_control::{radio_groups, CheckState, FormControlKind, ListSelection};
use crate::named_range::{NameScope, NamedRange, NamedRangeCollection};
use crate::worksheet::{SheetVisibility, Worksheet};
use crate::MAX_SHEET_NAME_LEN;
use duke_sheets_chart::Chart;

static NEXT_WORKBOOK_NONCE: AtomicU64 = AtomicU64::new(1);

/// A slot in the workbook tab bar, referencing either a worksheet or a chartsheet by index.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetSlot {
    Worksheet(usize),
    ChartSheet(usize),
}

/// A workbook (spreadsheet document)
///
/// A workbook contains one or more worksheets and global settings.
#[derive(Debug)]
pub struct Workbook {
    /// Worksheets in the workbook
    worksheets: Vec<Worksheet>,
    /// Chart sheets in the workbook
    chartsheets: Vec<ChartSheet>,
    /// Tab-bar ordering of worksheets and chartsheets.
    /// Empty means "use default order: all worksheets first, then all chartsheets".
    sheet_order: Vec<SheetSlot>,
    /// Workbook settings
    settings: WorkbookSettings,
    /// Active sheet index
    active_sheet: usize,
    /// Named ranges (defined names)
    named_ranges: NamedRangeCollection,
    /// Opaque calculation cache, populated and consumed by the calculation engine.
    /// Stored as type-erased `Box<dyn Any>` so the core crate needs no dependency
    /// on `duke-sheets-formula`.
    calc_cache: Option<Box<dyn Any + Send + Sync>>,
    /// Structural generation counter - incremented when sheets are added, removed,
    /// reordered, or renamed. The calculation engine uses this to detect stale caches.
    structural_generation: u64,
    /// Unique identity for roundtrip state lookup (not persisted).
    nonce: u64,
}

impl Workbook {
    /// Create a new empty workbook with one worksheet
    pub fn new() -> Self {
        let mut wb = Self {
            worksheets: Vec::new(),
            chartsheets: Vec::new(),
            sheet_order: Vec::new(),
            settings: WorkbookSettings::default(),
            active_sheet: 0,
            named_ranges: NamedRangeCollection::new(),
            calc_cache: None,
            structural_generation: 0,
            nonce: NEXT_WORKBOOK_NONCE.fetch_add(1, Ordering::Relaxed),
        };
        wb.add_worksheet_with_name("Sheet1").unwrap();
        wb
    }

    /// Create an empty workbook with no worksheets
    pub fn empty() -> Self {
        Self {
            worksheets: Vec::new(),
            chartsheets: Vec::new(),
            sheet_order: Vec::new(),
            settings: WorkbookSettings::default(),
            active_sheet: 0,
            named_ranges: NamedRangeCollection::new(),
            calc_cache: None,
            structural_generation: 0,
            nonce: NEXT_WORKBOOK_NONCE.fetch_add(1, Ordering::Relaxed),
        }
    }

    /// Get the number of worksheets
    pub fn sheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Get the tab-bar order as a slice of SheetSlot entries.
    /// Empty means the default order (all worksheets first, then all chartsheets).
    pub fn sheet_order(&self) -> &[SheetSlot] {
        &self.sheet_order
    }

    /// Get a mutable reference to the tab-bar order.
    pub fn sheet_order_mut(&mut self) -> &mut Vec<SheetSlot> {
        &mut self.sheet_order
    }

    /// Total number of tabs (worksheets + chartsheets).
    pub fn total_sheet_count(&self) -> usize {
        self.worksheets.len() + self.chartsheets.len()
    }

    /// Check if the workbook has no worksheets and no chartsheets
    pub fn is_empty(&self) -> bool {
        self.worksheets.is_empty() && self.chartsheets.is_empty()
    }

    /// Get a worksheet by index
    pub fn worksheet(&self, index: usize) -> Option<&Worksheet> {
        self.worksheets.get(index)
    }

    /// Get a mutable worksheet by index
    pub fn worksheet_mut(&mut self, index: usize) -> Option<&mut Worksheet> {
        self.worksheets.get_mut(index)
    }

    /// Get a worksheet by name
    pub fn worksheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.worksheets.iter().find(|ws| ws.name() == name)
    }

    /// Get a mutable worksheet by name
    pub fn worksheet_by_name_mut(&mut self, name: &str) -> Option<&mut Worksheet> {
        self.worksheets.iter_mut().find(|ws| ws.name() == name)
    }

    /// Get the index of a worksheet by name
    pub fn sheet_index(&self, name: &str) -> Option<usize> {
        self.worksheets.iter().position(|ws| ws.name() == name)
    }

    /// Iterate over all worksheets
    pub fn worksheets(&self) -> impl Iterator<Item = &Worksheet> {
        self.worksheets.iter()
    }

    /// Iterate over all worksheets mutably
    pub fn worksheets_mut(&mut self) -> impl Iterator<Item = &mut Worksheet> {
        self.worksheets.iter_mut()
    }

    /// Synchronize form-control state into each control's linked cell.
    ///
    /// This mirrors Excel's runtime behavior: checkboxes write booleans
    /// (`#N/A` for mixed), single-select lists and dropdowns write a
    /// one-based item index, option-button groups write their one-based
    /// selected index, and scrollbars/spinners write their numeric value.
    /// Fresh unchecked/no-selection/multi-select links preserve a blank cell;
    /// an existing linked value is reset to `FALSE` or `0`. Existing formulas
    /// are replaced. Malformed, external-workbook, and unknown-sheet links are
    /// left unchanged. If multiple controls target one cell, the last control
    /// in worksheet order wins.
    ///
    /// Returns the number of distinct linked cells updated.
    pub fn sync_form_control_links(&mut self) -> usize {
        let mut updates: BTreeMap<(usize, u32, u16), CellValue> = BTreeMap::new();

        for source_sheet in 0..self.worksheets.len() {
            let controls = self.worksheets[source_sheet].form_controls();
            let mut radio_values = vec![None; controls.len()];
            for group in radio_groups(controls) {
                let value = group
                    .iter()
                    .position(|&index| {
                        matches!(
                            controls[index].kind,
                            FormControlKind::OptionButton {
                                state: CheckState::Checked,
                                ..
                            }
                        )
                    })
                    .map(|index| CellValue::Number(index as f64 + 1.0))
                    .unwrap_or(CellValue::Empty);
                for index in group {
                    radio_values[index] = Some(value.clone());
                }
            }

            for (index, control) in controls.iter().enumerate() {
                let (link, value) = match &control.kind {
                    FormControlKind::Checkbox {
                        state, cell_link, ..
                    } => {
                        let value = match state {
                            CheckState::Unchecked => CellValue::Boolean(false),
                            CheckState::Checked => CellValue::Boolean(true),
                            CheckState::Mixed => CellValue::Error(CellError::Na),
                        };
                        (cell_link.as_deref(), value)
                    }
                    FormControlKind::ListBox {
                        cell_link,
                        selection,
                        selected,
                        ..
                    } => {
                        let value = if *selection == ListSelection::Single {
                            selected
                                .first()
                                .map(|index| CellValue::Number(f64::from(*index) + 1.0))
                                .unwrap_or(CellValue::Empty)
                        } else {
                            CellValue::Empty
                        };
                        (cell_link.as_deref(), value)
                    }
                    FormControlKind::Dropdown {
                        cell_link,
                        selected,
                        ..
                    } => (
                        cell_link.as_deref(),
                        selected
                            .map(|index| CellValue::Number(f64::from(index) + 1.0))
                            .unwrap_or(CellValue::Empty),
                    ),
                    FormControlKind::Scrollbar {
                        cell_link, value, ..
                    }
                    | FormControlKind::Spinner {
                        cell_link, value, ..
                    } => (cell_link.as_deref(), CellValue::Number(f64::from(*value))),
                    FormControlKind::OptionButton { cell_link, .. } => (
                        cell_link.as_deref(),
                        radio_values[index].clone().unwrap_or(CellValue::Empty),
                    ),
                    FormControlKind::Button { .. }
                    | FormControlKind::Label { .. }
                    | FormControlKind::GroupBox { .. } => continue,
                };

                if let Some((sheet, address)) =
                    link.and_then(|link| self.resolve_control_link(source_sheet, link))
                {
                    let existing = self.worksheets[sheet].get_value_at(address.row, address.col);
                    let fresh_blank = existing == CellValue::Empty
                        && !self.worksheets[sheet].has_formula_at(address.row, address.col);
                    let value = if matches!(
                        &control.kind,
                        FormControlKind::Checkbox {
                            state: CheckState::Unchecked,
                            ..
                        }
                    ) && fresh_blank
                    {
                        CellValue::Empty
                    } else if value == CellValue::Empty && !fresh_blank {
                        CellValue::Number(0.0)
                    } else {
                        value
                    };
                    updates.insert((sheet, address.row, address.col), value);
                }
            }
        }

        let mut updated = 0;
        for ((sheet, row, col), value) in updates {
            if self.worksheets[sheet]
                .set_cell_value_at(row, col, value)
                .is_ok()
            {
                updated += 1;
            }
        }
        updated
    }

    /// Create a serialization snapshot with form-control linked cells
    /// synchronized. Calculation caches are omitted; persisted workbook state
    /// and the roundtrip nonce are retained.
    #[doc(hidden)]
    pub fn synchronized_for_save(&self) -> Option<Self> {
        if !self
            .worksheets
            .iter()
            .flat_map(|sheet| sheet.form_controls())
            .any(|control| control.cell_link().is_some())
        {
            return None;
        }
        let mut snapshot = Self {
            worksheets: self.worksheets.clone(),
            chartsheets: self.chartsheets.clone(),
            sheet_order: self.sheet_order.clone(),
            settings: self.settings.clone(),
            active_sheet: self.active_sheet,
            named_ranges: self.named_ranges.clone(),
            calc_cache: None,
            structural_generation: self.structural_generation,
            nonce: self.nonce,
        };
        snapshot.sync_form_control_links();
        Some(snapshot)
    }

    fn resolve_control_link(
        &self,
        source_sheet: usize,
        link: &str,
    ) -> Option<(usize, CellAddress)> {
        // Names may refer to other names; Excel resolves the chain.
        self.resolve_control_link_depth(source_sheet, link, 4)
    }

    fn resolve_control_link_depth(
        &self,
        source_sheet: usize,
        link: &str,
        depth: u8,
    ) -> Option<(usize, CellAddress)> {
        let depth = depth.checked_sub(1)?;
        let link = link.trim().strip_prefix('=').unwrap_or(link.trim());
        let (sheet, address) = match link.rsplit_once('!') {
            Some((sheet, address)) => {
                let sheet = parse_link_sheet_name(sheet)?;
                let folded = sheet.to_lowercase();
                let index = self
                    .worksheets
                    .iter()
                    .position(|worksheet| worksheet.name().to_lowercase() == folded)?;
                (index, address)
            }
            None => (source_sheet, link),
        };
        if let Ok(address) = CellAddress::parse(address) {
            return Some((sheet, address));
        }
        // Excel also accepts a defined name as a cell link.
        let name = self.named_ranges.get(address, sheet)?;
        self.resolve_control_link_depth(sheet, name.expression(), depth)
    }

    /// Structural generation counter - incremented when sheets are added,
    /// removed, reordered, or renamed.
    pub fn structural_generation(&self) -> u64 {
        self.structural_generation
    }

    /// Unique nonce for roundtrip state lookup.
    pub fn nonce(&self) -> u64 {
        self.nonce
    }

    /// Take the calculation cache (moves it out of the workbook).
    pub fn take_calc_cache(&mut self) -> Option<Box<dyn Any + Send + Sync>> {
        self.calc_cache.take()
    }

    /// Store a calculation cache on the workbook.
    pub fn set_calc_cache(&mut self, cache: Box<dyn Any + Send + Sync>) {
        self.calc_cache = Some(cache);
    }

    /// Add a new worksheet with default name
    pub fn add_worksheet(&mut self) -> Result<usize> {
        let name = self.generate_sheet_name();
        self.add_worksheet_with_name(&name)
    }

    /// Add a new worksheet with specified name
    pub fn add_worksheet_with_name(&mut self, name: &str) -> Result<usize> {
        self.validate_sheet_name(name)?;

        let index = self.worksheets.len();
        let worksheet = Worksheet::new(name);
        self.worksheets.push(worksheet);
        if !self.sheet_order.is_empty() {
            self.sheet_order.push(SheetSlot::Worksheet(index));
        }
        self.structural_generation += 1;
        Ok(index)
    }

    /// Add a new worksheet with the exact name from a file, skipping validation.
    ///
    /// Real-world files can have empty names or names >31 chars. Readers must
    /// preserve them as-is; validation only applies to user-facing APIs.
    pub fn add_worksheet_with_name_unchecked(&mut self, name: &str) -> usize {
        let index = self.worksheets.len();
        let worksheet = Worksheet::new(name);
        self.worksheets.push(worksheet);
        self.structural_generation += 1;
        index
    }

    /// Insert a worksheet at a specific index
    pub fn insert_worksheet(&mut self, index: usize, name: &str) -> Result<()> {
        if index > self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }

        self.validate_sheet_name(name)?;

        let worksheet = Worksheet::new(name);
        self.worksheets.insert(index, worksheet);

        // Update sheet_order: increment indices >= the insertion point
        if !self.sheet_order.is_empty() {
            for slot in &mut self.sheet_order {
                if let SheetSlot::Worksheet(ref mut idx) = slot {
                    if *idx >= index {
                        *idx += 1;
                    }
                }
            }
            self.sheet_order.push(SheetSlot::Worksheet(index));
        }

        // Adjust active sheet index if needed
        if self.active_sheet >= index && !self.worksheets.is_empty() {
            self.active_sheet = self.active_sheet.saturating_add(1);
        }
        self.structural_generation += 1;
        Ok(())
    }

    /// Add an existing worksheet to the workbook
    pub fn add_existing_worksheet(&mut self, worksheet: Worksheet) -> Result<usize> {
        self.validate_sheet_name(worksheet.name())?;
        let index = self.worksheets.len();
        self.worksheets.push(worksheet);
        if !self.sheet_order.is_empty() {
            self.sheet_order.push(SheetSlot::Worksheet(index));
        }
        self.structural_generation += 1;
        Ok(index)
    }

    /// Remove a worksheet by index
    pub fn remove_worksheet(&mut self, index: usize) -> Result<Worksheet> {
        if index >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }

        let worksheet = self.worksheets.remove(index);

        // Update sheet_order: remove the entry and decrement indices above it
        if !self.sheet_order.is_empty() {
            self.sheet_order.retain(|slot| *slot != SheetSlot::Worksheet(index));
            for slot in &mut self.sheet_order {
                if let SheetSlot::Worksheet(ref mut idx) = slot {
                    if *idx > index {
                        *idx -= 1;
                    }
                }
            }
        }

        // Adjust active sheet index
        if !self.worksheets.is_empty() {
            if self.active_sheet >= self.worksheets.len() {
                self.active_sheet = self.worksheets.len() - 1;
            }
        } else {
            self.active_sheet = 0;
        }
        self.structural_generation += 1;
        Ok(worksheet)
    }

    /// Move a worksheet to a new position
    pub fn move_worksheet(&mut self, from: usize, to: usize) -> Result<()> {
        if from >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(from, self.worksheets.len()));
        }
        if to >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(to, self.worksheets.len()));
        }

        let worksheet = self.worksheets.remove(from);
        self.worksheets.insert(to, worksheet);

        // Update sheet_order indices to reflect the move
        if !self.sheet_order.is_empty() {
            for slot in &mut self.sheet_order {
                if let SheetSlot::Worksheet(ref mut idx) = slot {
                    if *idx == from {
                        *idx = to;
                    } else if from < to && *idx > from && *idx <= to {
                        *idx -= 1;
                    } else if from > to && *idx >= to && *idx < from {
                        *idx += 1;
                    }
                }
            }
        }

        // Adjust active sheet if needed
        if self.active_sheet == from {
            self.active_sheet = to;
        } else if from < self.active_sheet && to >= self.active_sheet {
            self.active_sheet = self.active_sheet.saturating_sub(1);
        } else if from > self.active_sheet && to <= self.active_sheet {
            self.active_sheet = self.active_sheet.saturating_add(1);
        }
        self.structural_generation += 1;
        Ok(())
    }

    /// Rename a worksheet
    pub fn rename_worksheet(&mut self, index: usize, new_name: &str) -> Result<()> {
        // Check index first
        if index >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }

        // Validate the new name (excluding current sheet from duplicate check)
        self.validate_sheet_name_excluding(new_name, Some(index))?;

        self.worksheets[index].set_name(new_name);
        self.structural_generation += 1;
        Ok(())
    }

    /// Get the active sheet index
    pub fn active_sheet(&self) -> usize {
        self.active_sheet
    }

    /// Set the active sheet index
    pub fn set_active_sheet(&mut self, index: usize) -> Result<()> {
        if index >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }
        self.active_sheet = index;
        Ok(())
    }

    /// Get workbook settings
    pub fn settings(&self) -> &WorkbookSettings {
        &self.settings
    }

    /// Get mutable workbook settings
    pub fn settings_mut(&mut self) -> &mut WorkbookSettings {
        &mut self.settings
    }

    /// Define a new workbook-scoped named range
    ///
    /// # Example
    /// ```
    /// use duke_sheets_core::Workbook;
    ///
    /// let mut wb = Workbook::new();
    /// wb.define_name("TaxRate", "Sheet1!$B$1").unwrap();
    /// ```
    pub fn define_name(&mut self, name: &str, refers_to: &str) -> Result<()> {
        self.define_name_with_scope(name, refers_to, NameScope::Workbook)
    }

    /// Define a named range with a specific scope
    pub fn define_name_with_scope(
        &mut self,
        name: &str,
        refers_to: &str,
        scope: NameScope,
    ) -> Result<()> {
        let range = NamedRange::new(name, refers_to, scope);
        self.named_ranges
            .define(range)
            .map_err(Error::InvalidName)?;
        self.structural_generation += 1;
        Ok(())
    }

    /// Define a sheet-scoped named range
    pub fn define_name_for_sheet(
        &mut self,
        name: &str,
        refers_to: &str,
        sheet_index: usize,
    ) -> Result<()> {
        self.define_name_with_scope(name, refers_to, NameScope::Sheet(sheet_index))
    }

    /// Get a named range by name, following Excel's scoping rules
    ///
    /// Looks for sheet-scoped name first (for the given sheet), then workbook-scoped.
    pub fn get_named_range(&self, name: &str, current_sheet: usize) -> Option<&NamedRange> {
        self.named_ranges.get(name, current_sheet)
    }

    /// Remove a workbook-scoped named range
    pub fn remove_name(&mut self, name: &str) -> Option<NamedRange> {
        let result = self.named_ranges.remove(name, &NameScope::Workbook);
        if result.is_some() {
            self.structural_generation += 1;
        }
        result
    }

    /// Remove a sheet-scoped named range
    pub fn remove_name_from_sheet(&mut self, name: &str, sheet_index: usize) -> Option<NamedRange> {
        let result = self.named_ranges
            .remove(name, &NameScope::Sheet(sheet_index));
        if result.is_some() {
            self.structural_generation += 1;
        }
        result
    }

    /// Get the named range collection (read-only)
    pub fn named_ranges(&self) -> &NamedRangeCollection {
        &self.named_ranges
    }

    /// Get the named range collection (mutable)
    pub fn named_ranges_mut(&mut self) -> &mut NamedRangeCollection {
        &mut self.named_ranges
    }

    /// Validate a sheet name
    fn validate_sheet_name(&self, name: &str) -> Result<()> {
        self.validate_sheet_name_excluding(name, None)
    }

    /// Validate a sheet name, optionally excluding a sheet from duplicate check
    fn validate_sheet_name_excluding(
        &self,
        name: &str,
        exclude_index: Option<usize>,
    ) -> Result<()> {
        // Check length
        if name.is_empty() {
            return Err(Error::InvalidSheetName("Sheet name cannot be empty".into()));
        }
        if name.len() > MAX_SHEET_NAME_LEN {
            return Err(Error::InvalidSheetName(format!(
                "Sheet name too long (max {} characters)",
                MAX_SHEET_NAME_LEN
            )));
        }

        // Check for invalid characters
        const INVALID_CHARS: &[char] = &[':', '\\', '/', '?', '*', '[', ']'];
        for c in INVALID_CHARS {
            if name.contains(*c) {
                return Err(Error::InvalidSheetName(format!(
                    "Sheet name cannot contain '{}'",
                    c
                )));
            }
        }

        // Check for duplicate names against worksheets (case-insensitive)
        let name_lower = name.to_lowercase();
        for (i, ws) in self.worksheets.iter().enumerate() {
            if Some(i) != exclude_index && ws.name().to_lowercase() == name_lower {
                return Err(Error::DuplicateSheetName(name.into()));
            }
        }

        // Check for duplicate names against chartsheets (case-insensitive)
        for cs in &self.chartsheets {
            if cs.name.to_lowercase() == name_lower {
                return Err(Error::DuplicateSheetName(name.into()));
            }
        }

        Ok(())
    }

    /// Generate a unique sheet name
    fn generate_sheet_name(&self) -> String {
        let mut n = self.worksheets.len() + 1;
        loop {
            let name = format!("Sheet{}", n);
            if self.validate_sheet_name(&name).is_ok() {
                return name;
            }
            n += 1;
        }
    }

    /// Add a chart sheet to the workbook.
    ///
    /// Returns an error if the name is empty or duplicates an existing
    /// worksheet or chartsheet name.
    pub fn add_chartsheet(&mut self, sheet: ChartSheet) -> Result<usize> {
        self.validate_sheet_name(&sheet.name)?;
        let index = self.chartsheets.len();
        self.chartsheets.push(sheet);
        if !self.sheet_order.is_empty() {
            self.sheet_order.push(SheetSlot::ChartSheet(index));
        }
        self.structural_generation += 1;
        Ok(index)
    }

    /// Add a chart sheet without name validation (for the reader).
    ///
    /// Real-world files can have names that violate validation rules.
    /// Readers must preserve them as-is.
    pub fn add_chartsheet_unchecked(&mut self, sheet: ChartSheet) -> usize {
        let index = self.chartsheets.len();
        self.chartsheets.push(sheet);
        self.structural_generation += 1;
        index
    }

    /// Get all chart sheets.
    pub fn chartsheets(&self) -> &[ChartSheet] {
        &self.chartsheets
    }

    /// Get a chart sheet by index.
    pub fn chartsheet(&self, index: usize) -> Option<&ChartSheet> {
        self.chartsheets.get(index)
    }

    /// Get the number of chart sheets.
    pub fn chartsheet_count(&self) -> usize {
        self.chartsheets.len()
    }
}

fn parse_link_sheet_name(value: &str) -> Option<String> {
    let value = value.trim();
    if value.contains(['[', ']']) {
        return None;
    }
    if let Some(quoted) = value.strip_prefix('\'') {
        let quoted = quoted.strip_suffix('\'')?;
        let mut name = String::with_capacity(quoted.len());
        let mut chars = quoted.chars().peekable();
        while let Some(ch) = chars.next() {
            if ch == '\'' {
                if chars.next() != Some('\'') {
                    return None;
                }
                name.push('\'');
            } else {
                name.push(ch);
            }
        }
        (!name.is_empty()).then_some(name)
    } else if value.is_empty() || value.contains('\'') {
        None
    } else {
        Some(value.to_string())
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

/// Workbook-level settings
#[derive(Debug, Clone)]
pub struct WorkbookSettings {
    /// Date system: false = 1900 (Windows), true = 1904 (Mac)
    pub date_1904: bool,
    /// Workbook is protected
    pub protected: bool,
    /// Password hash for protection (if protected)
    pub password_hash: Option<u16>,
    /// Calculate formulas on open
    pub calc_on_open: bool,
    /// Default theme name
    pub theme: Option<String>,
}

impl Default for WorkbookSettings {
    fn default() -> Self {
        Self {
            date_1904: false,
            protected: false,
            password_hash: None,
            calc_on_open: true,
            theme: None,
        }
    }
}

/// A chart sheet - a sheet that contains only a chart, no cell data.
#[derive(Debug, Clone)]
pub struct ChartSheet {
    /// Sheet name (as it appears on the tab)
    pub name: String,
    /// The chart displayed in this sheet
    pub chart: Chart,
    /// Sheet visibility
    pub visibility: SheetVisibility,
    /// Raw XML fragments for non-chart drawing anchors, preserved for roundtrip.
    #[doc(hidden)]
    pub raw_drawing_objects: Vec<Vec<u8>>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{CellError, CellValue, CheckState, FormControl, FormControlKind, ListSelection};

    #[test]
    fn test_new_workbook() {
        let wb = Workbook::new();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.worksheet(0).unwrap().name(), "Sheet1");
    }

    #[test]
    fn test_add_worksheets() {
        let mut wb = Workbook::new();

        let idx = wb.add_worksheet().unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_count(), 2);

        let idx = wb.add_worksheet_with_name("Data").unwrap();
        assert_eq!(idx, 2);
        assert_eq!(wb.worksheet(2).unwrap().name(), "Data");
    }

    #[test]
    fn test_duplicate_name() {
        let mut wb = Workbook::new();

        // Case-insensitive duplicate check
        assert!(wb.add_worksheet_with_name("SHEET1").is_err());
        assert!(wb.add_worksheet_with_name("sheet1").is_err());
    }

    #[test]
    fn test_invalid_sheet_name() {
        let mut wb = Workbook::new();

        assert!(wb.add_worksheet_with_name("").is_err());
        assert!(wb.add_worksheet_with_name("Sheet/1").is_err());
        assert!(wb.add_worksheet_with_name("Sheet:1").is_err());
        assert!(wb.add_worksheet_with_name("Sheet[1]").is_err());

        // Too long
        let long_name = "A".repeat(MAX_SHEET_NAME_LEN + 1);
        assert!(wb.add_worksheet_with_name(&long_name).is_err());
    }

    #[test]
    fn test_move_worksheet() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("A").unwrap();
        wb.add_worksheet_with_name("B").unwrap();
        wb.add_worksheet_with_name("C").unwrap();

        // Move C to position 1
        wb.move_worksheet(3, 1).unwrap();

        assert_eq!(wb.worksheet(0).unwrap().name(), "Sheet1");
        assert_eq!(wb.worksheet(1).unwrap().name(), "C");
        assert_eq!(wb.worksheet(2).unwrap().name(), "A");
        assert_eq!(wb.worksheet(3).unwrap().name(), "B");
    }

    #[test]
    fn test_worksheet_by_name() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Data").unwrap();

        assert!(wb.worksheet_by_name("Data").is_some());
        assert!(wb.worksheet_by_name("NonExistent").is_none());
    }

    #[test]
    fn test_insert_worksheet_updates_sheet_order() {
        let mut wb = Workbook::new();
        // wb starts with Sheet1 at worksheets[0], sheet_order is empty
        // Manually populate sheet_order to simulate a read workbook:
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        // Now insert at position 0
        wb.insert_worksheet(0, "Inserted").unwrap();
        // Expected: sheet_order should have [Worksheet(0), Worksheet(1)]
        // The original Worksheet(0) became Worksheet(1) because it shifted right
        // The new sheet at index 0 was pushed to the end of sheet_order
        assert_eq!(wb.sheet_order().len(), 2);
        assert!(wb.sheet_order().contains(&SheetSlot::Worksheet(0)));
        assert!(wb.sheet_order().contains(&SheetSlot::Worksheet(1)));
        // Verify the original slot was renumbered:
        assert_eq!(wb.sheet_order()[0], SheetSlot::Worksheet(1)); // was 0, now 1
        assert_eq!(wb.sheet_order()[1], SheetSlot::Worksheet(0)); // newly inserted, pushed to end
    }

    #[test]
    fn test_remove_worksheet_updates_sheet_order() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Sheet2").unwrap();
        // worksheets: [Sheet1(0), Sheet2(1)]
        // Manually set sheet_order:
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        wb.sheet_order_mut().push(SheetSlot::Worksheet(1));
        // Remove worksheet 0 (Sheet1)
        wb.remove_worksheet(0).unwrap();
        // Expected: sheet_order should have [Worksheet(0)] - the old index 1 became 0
        assert_eq!(wb.sheet_order().len(), 1);
        assert_eq!(wb.sheet_order()[0], SheetSlot::Worksheet(0));
        assert_eq!(wb.worksheet(0).unwrap().name(), "Sheet2");
    }

    #[test]
    fn test_move_worksheet_updates_sheet_order() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Sheet2").unwrap();
        wb.add_worksheet_with_name("Sheet3").unwrap();
        // worksheets: [Sheet1(0), Sheet2(1), Sheet3(2)]
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        wb.sheet_order_mut().push(SheetSlot::Worksheet(1));
        wb.sheet_order_mut().push(SheetSlot::Worksheet(2));
        // Move worksheet from index 0 to index 2
        wb.move_worksheet(0, 2).unwrap();
        // After move: worksheets = [Sheet2(0), Sheet3(1), Sheet1(2)]
        // sheet_order indices should be remapped:
        // Old 0→new 2, old 1→new 0, old 2→new 1
        assert_eq!(wb.sheet_order()[0], SheetSlot::Worksheet(2)); // was 0, moved to 2
        assert_eq!(wb.sheet_order()[1], SheetSlot::Worksheet(0)); // was 1, shifted to 0
        assert_eq!(wb.sheet_order()[2], SheetSlot::Worksheet(1)); // was 2, shifted to 1
    }

    #[test]
    fn test_add_worksheet_pushes_to_nonempty_sheet_order() {
        let mut wb = Workbook::new();
        // Simulate reading: manually populate sheet_order
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        // Now add a new worksheet via the API
        wb.add_worksheet_with_name("Sheet2").unwrap();
        // Since sheet_order was non-empty, the new sheet should have been appended
        assert_eq!(wb.sheet_order().len(), 2);
        assert_eq!(wb.sheet_order()[1], SheetSlot::Worksheet(1));
    }

    #[test]
    fn test_add_chartsheet_pushes_to_nonempty_sheet_order() {
        use duke_sheets_chart::ChartType;

        let mut wb = Workbook::new();
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        let cs = ChartSheet {
            name: "Chart1".into(),
            chart: Chart::new(ChartType::Pie),
            visibility: SheetVisibility::Visible,
            raw_drawing_objects: Vec::new(),
        };
        wb.add_chartsheet(cs).unwrap();
        assert_eq!(wb.sheet_order().len(), 2);
        assert_eq!(wb.sheet_order()[1], SheetSlot::ChartSheet(0));
    }

    #[test]
    fn sync_form_control_links_matches_excel_semantics() {
        let mut wb = Workbook::new();
        wb.rename_worksheet(0, "Controls").unwrap();
        wb.add_worksheet_with_name("Linked Data").unwrap();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_formula("A8", "=FALSE").unwrap();
        ws.set_cell_formula("A11", "=0").unwrap();
        ws.set_cell_value("A3", 99.0).unwrap();
        ws.set_cell_value("A10", true).unwrap();

        let kinds = [
            FormControlKind::Checkbox {
                caption: "mixed".into(),
                state: CheckState::Mixed,
                cell_link: Some("$A$1".into()),
                no_3d: false,
            },
            FormControlKind::ListBox {
                input_range: None,
                cell_link: Some("$A$2".into()),
                selection: ListSelection::Single,
                selected: vec![2],
                no_3d: false,
            },
            FormControlKind::ListBox {
                input_range: None,
                cell_link: Some("$A$3".into()),
                selection: ListSelection::Multi,
                selected: vec![0, 2, 3],
                no_3d: false,
            },
            FormControlKind::Dropdown {
                input_range: None,
                cell_link: Some("$A$4".into()),
                selected: None,
                lines: 8,
                no_3d: false,
            },
            FormControlKind::Scrollbar {
                value: 55,
                min: 0,
                max: 100,
                increment: 1,
                page: 10,
                horizontal: false,
                cell_link: Some("$A$5".into()),
            },
            FormControlKind::Spinner {
                value: 18,
                min: 0,
                max: 30,
                increment: 1,
                cell_link: Some("'Linked Data'!$A$1".into()),
            },
            FormControlKind::OptionButton {
                caption: "one".into(),
                state: CheckState::Unchecked,
                cell_link: Some("$A$7".into()),
                first_in_group: false,
                no_3d: false,
            },
            FormControlKind::OptionButton {
                caption: "two".into(),
                state: CheckState::Checked,
                cell_link: Some("$A$7".into()),
                first_in_group: false,
                no_3d: false,
            },
            FormControlKind::Checkbox {
                caption: "formula overwrite".into(),
                state: CheckState::Checked,
                cell_link: Some("$A$8".into()),
                no_3d: false,
            },
            FormControlKind::Checkbox {
                caption: "fresh unchecked".into(),
                state: CheckState::Unchecked,
                cell_link: Some("$A$9".into()),
                no_3d: false,
            },
            FormControlKind::Checkbox {
                caption: "changed unchecked".into(),
                state: CheckState::Unchecked,
                cell_link: Some("$A$10".into()),
                no_3d: false,
            },
            FormControlKind::Dropdown {
                input_range: None,
                cell_link: Some("$A$11".into()),
                selected: None,
                lines: 8,
                no_3d: false,
            },
        ];
        for kind in kinds {
            ws.add_form_control(FormControl::new(kind));
        }

        assert_eq!(wb.sync_form_control_links(), 11);
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value("A1").unwrap(), CellValue::Error(CellError::Na));
        assert_eq!(ws.get_value("A2").unwrap(), CellValue::Number(3.0));
        assert_eq!(ws.get_value("A3").unwrap(), CellValue::Number(0.0));
        assert_eq!(ws.get_value("A4").unwrap(), CellValue::Empty);
        assert_eq!(ws.get_value("A5").unwrap(), CellValue::Number(55.0));
        assert_eq!(ws.get_value("A7").unwrap(), CellValue::Number(2.0));
        assert_eq!(ws.get_value("A8").unwrap(), CellValue::Boolean(true));
        assert_eq!(ws.get_value("A9").unwrap(), CellValue::Empty);
        assert_eq!(ws.get_value("A10").unwrap(), CellValue::Boolean(false));
        assert_eq!(ws.get_value("A11").unwrap(), CellValue::Number(0.0));
        assert!(!ws.has_formula_at(10, 0));
        assert!(!ws.has_formula_at(7, 0));
        assert_eq!(
            wb.worksheet(1).unwrap().get_value("A1").unwrap(),
            CellValue::Number(18.0)
        );
    }

    #[test]
    fn sync_form_control_links_resolves_quoted_sheets_and_skips_unsupported_links() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("It's A! Sheet").unwrap();
        let ws = wb.worksheet_mut(0).unwrap();
        for link in [
            "='It''s A! Sheet'!$B$2",
            "SUM(A1:A3)",
            "[Other.xlsx]Sheet1!$A$1",
            "Missing!$A$1",
            "=''!$A$1",
            "='Bob's'!$A$1",
        ] {
            ws.add_form_control(FormControl::new(FormControlKind::Checkbox {
                caption: link.into(),
                state: CheckState::Checked,
                cell_link: Some(link.into()),
                no_3d: false,
            }));
        }

        assert_eq!(wb.sync_form_control_links(), 1);
        assert_eq!(
            wb.worksheet(1).unwrap().get_value("B2").unwrap(),
            CellValue::Boolean(true)
        );
    }

    #[test]
    fn sync_form_control_links_resolves_defined_names() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Data").unwrap();
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("Target", "Data!$B$2"));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("Alias", "=Target"));
        wb.named_ranges_mut().define_or_update(NamedRange::new(
            "Local",
            "$C$3",
            NameScope::Sheet(0),
        ));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("Wide", "Data!$A$1:$A$4"));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("LoopA", "LoopB"));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("LoopB", "LoopA"));

        let ws = wb.worksheet_mut(0).unwrap();
        for (link, value) in [
            ("Target", 11),
            ("Alias", 22),
            ("Local", 33),
            ("Wide", 44),
            ("LoopA", 55),
            ("NoSuchName", 66),
        ] {
            ws.add_form_control(FormControl::new(FormControlKind::Spinner {
                value,
                min: 0,
                max: 100,
                increment: 1,
                cell_link: Some(link.into()),
            }));
        }

        assert_eq!(wb.sync_form_control_links(), 2);
        assert_eq!(
            wb.worksheet(1).unwrap().get_value("B2").unwrap(),
            CellValue::Number(22.0),
            "workbook-scoped name resolves; the aliased control targets it too"
        );
        assert_eq!(
            wb.worksheet(0).unwrap().get_value("C3").unwrap(),
            CellValue::Number(33.0),
            "sheet-scoped name resolves against the control's sheet"
        );
    }

    #[test]
    fn sync_form_control_links_uses_control_order_for_duplicate_targets() {
        let mut wb = Workbook::new();
        assert!(wb.synchronized_for_save().is_none());
        let ws = wb.worksheet_mut(0).unwrap();
        ws.add_form_control(FormControl::new(FormControlKind::Spinner {
            value: 7,
            min: 0,
            max: 10,
            increment: 1,
            cell_link: Some("$A$1".into()),
        }));
        ws.add_form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "later".into(),
            state: CheckState::Checked,
            cell_link: Some("$A$1".into()),
            no_3d: false,
        }));

        assert_eq!(wb.sync_form_control_links(), 1);
        assert_eq!(
            wb.worksheet(0).unwrap().get_value("A1").unwrap(),
            CellValue::Boolean(true)
        );
    }
}
