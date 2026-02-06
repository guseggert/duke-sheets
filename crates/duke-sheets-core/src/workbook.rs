//! Workbook type - the main document structure

use crate::error::{Error, Result};
use crate::named_range::{NameScope, NamedRange, NamedRangeCollection};
use crate::worksheet::Worksheet;
use crate::MAX_SHEET_NAME_LEN;

/// A workbook (spreadsheet document)
///
/// A workbook contains one or more worksheets and global settings.
#[derive(Debug)]
pub struct Workbook {
    /// Worksheets in the workbook
    worksheets: Vec<Worksheet>,
    /// Workbook settings
    settings: WorkbookSettings,
    /// Active sheet index
    active_sheet: usize,
    /// Named ranges (defined names)
    named_ranges: NamedRangeCollection,
}

impl Workbook {
    /// Create a new empty workbook with one worksheet
    pub fn new() -> Self {
        let mut wb = Self {
            worksheets: Vec::new(),
            settings: WorkbookSettings::default(),
            active_sheet: 0,
            named_ranges: NamedRangeCollection::new(),
        };
        wb.add_worksheet_with_name("Sheet1").unwrap();
        wb
    }

    /// Create an empty workbook with no worksheets
    pub fn empty() -> Self {
        Self {
            worksheets: Vec::new(),
            settings: WorkbookSettings::default(),
            active_sheet: 0,
            named_ranges: NamedRangeCollection::new(),
        }
    }

    /// Get the number of worksheets
    pub fn sheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Check if the workbook has no worksheets
    pub fn is_empty(&self) -> bool {
        self.worksheets.is_empty()
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

        Ok(index)
    }

    /// Insert a worksheet at a specific index
    pub fn insert_worksheet(&mut self, index: usize, name: &str) -> Result<()> {
        if index > self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }

        self.validate_sheet_name(name)?;

        let worksheet = Worksheet::new(name);
        self.worksheets.insert(index, worksheet);

        // Adjust active sheet index if needed
        if self.active_sheet >= index && !self.worksheets.is_empty() {
            self.active_sheet = self.active_sheet.saturating_add(1);
        }

        Ok(())
    }

    /// Add an existing worksheet to the workbook
    pub fn add_existing_worksheet(&mut self, worksheet: Worksheet) -> Result<usize> {
        self.validate_sheet_name(worksheet.name())?;
        let index = self.worksheets.len();
        self.worksheets.push(worksheet);
        Ok(index)
    }

    /// Remove a worksheet by index
    pub fn remove_worksheet(&mut self, index: usize) -> Result<Worksheet> {
        if index >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }

        let worksheet = self.worksheets.remove(index);

        // Adjust active sheet index
        if !self.worksheets.is_empty() {
            if self.active_sheet >= self.worksheets.len() {
                self.active_sheet = self.worksheets.len() - 1;
            }
        } else {
            self.active_sheet = 0;
        }

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

        // Adjust active sheet if needed
        if self.active_sheet == from {
            self.active_sheet = to;
        } else if from < self.active_sheet && to >= self.active_sheet {
            self.active_sheet = self.active_sheet.saturating_sub(1);
        } else if from > self.active_sheet && to <= self.active_sheet {
            self.active_sheet = self.active_sheet.saturating_add(1);
        }

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

    // ==================== Named Ranges ====================

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
            .map_err(|e| Error::InvalidName(e))
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
        self.named_ranges.remove(name, &NameScope::Workbook)
    }

    /// Remove a sheet-scoped named range
    pub fn remove_name_from_sheet(&mut self, name: &str, sheet_index: usize) -> Option<NamedRange> {
        self.named_ranges
            .remove(name, &NameScope::Sheet(sheet_index))
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

        // Check for duplicate names (case-insensitive)
        let name_lower = name.to_lowercase();
        for (i, ws) in self.worksheets.iter().enumerate() {
            if Some(i) != exclude_index && ws.name().to_lowercase() == name_lower {
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
}
