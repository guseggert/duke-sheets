//! Workbook type - the main document structure

use std::sync::atomic::{AtomicU64, Ordering};

use std::any::Any;
use crate::error::{Error, Result};
use crate::named_range::{NameScope, NamedRange, NamedRangeCollection};
use crate::worksheet::{SheetVisibility, Worksheet};
use duke_sheets_chart::Chart;
use crate::MAX_SHEET_NAME_LEN;

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
    /// Opaque pivot runtime cache, populated and consumed by the pivot engine.
    /// This is intentionally separate from file-format pivot cache records and
    /// from the formula calculation cache.
    pivot_runtime_cache: Option<Box<dyn Any + Send + Sync>>,
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
            pivot_runtime_cache: None,
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
            pivot_runtime_cache: None,
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

    /// Take the pivot runtime cache (moves it out of the workbook).
    #[doc(hidden)]
    pub fn take_pivot_runtime_cache(&mut self) -> Option<Box<dyn Any + Send + Sync>> {
        self.pivot_runtime_cache.take()
    }

    /// Store a pivot runtime cache on the workbook.
    #[doc(hidden)]
    pub fn set_pivot_runtime_cache(&mut self, cache: Box<dyn Any + Send + Sync>) {
        self.pivot_runtime_cache = Some(cache);
    }

    /// Clear any pivot runtime cache stored on the workbook.
    #[doc(hidden)]
    pub fn clear_pivot_runtime_cache(&mut self) {
        self.pivot_runtime_cache = None;
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
}
