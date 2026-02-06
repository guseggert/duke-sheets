//! Named range definitions
//!
//! Named ranges allow users to assign meaningful names to cells or ranges of cells,
//! making formulas easier to read and maintain.
//!
//! # Example
//!
//! ```text
//! // Define a named range "TaxRate" that refers to cell B1
//! workbook.define_name("TaxRate", "Sheet1!$B$1", NameScope::Workbook)?;
//!
//! // Use it in a formula
//! =Price * TaxRate
//! ```

use std::collections::HashMap;

/// Scope of a named range
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NameScope {
    /// Available throughout the workbook (global)
    Workbook,
    /// Scoped to a specific sheet (local)
    Sheet(usize),
}

/// A named range definition
///
/// Named ranges can refer to:
/// - A single cell: `Sheet1!$A$1`
/// - A range of cells: `Sheet1!$A$1:$D$10`
/// - A constant value: `0.0725` (for things like tax rates)
/// - A formula expression: `=SUM(Sales)` (rare, but supported)
#[derive(Debug, Clone)]
pub struct NamedRange {
    /// The name (e.g., "SalesData", "TaxRate")
    /// Names are case-insensitive in Excel
    pub name: String,
    /// Scope of this name (workbook-wide or sheet-specific)
    pub scope: NameScope,
    /// What the name refers to
    /// This is stored as a string that can be parsed as a formula/reference
    /// Examples:
    /// - "Sheet1!$A$1" - single cell
    /// - "Sheet1!$A$1:$D$10" - range
    /// - "0.0725" - constant
    /// - "=SUM(A1:A10)" - formula (starts with =)
    pub refers_to: String,
    /// Optional comment/description for documentation
    pub comment: Option<String>,
    /// Whether this name is hidden from the UI
    pub hidden: bool,
}

impl NamedRange {
    /// Create a new named range
    pub fn new(name: impl Into<String>, refers_to: impl Into<String>, scope: NameScope) -> Self {
        Self {
            name: name.into(),
            scope,
            refers_to: refers_to.into(),
            comment: None,
            hidden: false,
        }
    }

    /// Create a workbook-scoped named range
    pub fn workbook_scope(name: impl Into<String>, refers_to: impl Into<String>) -> Self {
        Self::new(name, refers_to, NameScope::Workbook)
    }

    /// Create a sheet-scoped named range
    pub fn sheet_scope(
        name: impl Into<String>,
        refers_to: impl Into<String>,
        sheet_index: usize,
    ) -> Self {
        Self::new(name, refers_to, NameScope::Sheet(sheet_index))
    }

    /// Set a comment for this named range
    pub fn with_comment(mut self, comment: impl Into<String>) -> Self {
        self.comment = Some(comment.into());
        self
    }

    /// Mark this named range as hidden
    pub fn hidden(mut self) -> Self {
        self.hidden = true;
        self
    }

    /// Check if the refers_to is a formula (starts with =)
    pub fn is_formula(&self) -> bool {
        self.refers_to.starts_with('=')
    }

    /// Get the refers_to expression without the leading = if it's a formula
    pub fn expression(&self) -> &str {
        if self.refers_to.starts_with('=') {
            &self.refers_to[1..]
        } else {
            &self.refers_to
        }
    }
}

/// Collection of named ranges with efficient lookup
#[derive(Debug, Default, Clone)]
pub struct NamedRangeCollection {
    /// Named ranges stored by lowercase name for case-insensitive lookup
    /// The key is (name_lowercase, scope_key) where scope_key is:
    /// - "" for workbook scope
    /// - "sheet:{index}" for sheet scope
    ranges: HashMap<String, NamedRange>,
}

impl NamedRangeCollection {
    /// Create a new empty collection
    pub fn new() -> Self {
        Self::default()
    }

    /// Generate the storage key for a named range
    fn make_key(name: &str, scope: &NameScope) -> String {
        let name_lower = name.to_lowercase();
        match scope {
            NameScope::Workbook => name_lower,
            NameScope::Sheet(idx) => format!("{}:sheet:{}", name_lower, idx),
        }
    }

    /// Define a new named range
    ///
    /// Returns an error if a name with the same scope already exists
    pub fn define(&mut self, range: NamedRange) -> Result<(), String> {
        let key = Self::make_key(&range.name, &range.scope);

        if self.ranges.contains_key(&key) {
            return Err(format!(
                "Named range '{}' already exists in this scope",
                range.name
            ));
        }

        self.ranges.insert(key, range);
        Ok(())
    }

    /// Define or update a named range
    pub fn define_or_update(&mut self, range: NamedRange) {
        let key = Self::make_key(&range.name, &range.scope);
        self.ranges.insert(key, range);
    }

    /// Get a named range by name and current sheet context
    ///
    /// This follows Excel's scoping rules:
    /// 1. First look for a sheet-scoped name matching the current sheet
    /// 2. Then look for a workbook-scoped name
    pub fn get(&self, name: &str, current_sheet: usize) -> Option<&NamedRange> {
        // First try sheet-scoped
        let sheet_key = Self::make_key(name, &NameScope::Sheet(current_sheet));
        if let Some(range) = self.ranges.get(&sheet_key) {
            return Some(range);
        }

        // Then try workbook-scoped
        let workbook_key = Self::make_key(name, &NameScope::Workbook);
        self.ranges.get(&workbook_key)
    }

    /// Get a named range by exact scope
    pub fn get_exact(&self, name: &str, scope: &NameScope) -> Option<&NamedRange> {
        let key = Self::make_key(name, scope);
        self.ranges.get(&key)
    }

    /// Remove a named range
    pub fn remove(&mut self, name: &str, scope: &NameScope) -> Option<NamedRange> {
        let key = Self::make_key(name, scope);
        self.ranges.remove(&key)
    }

    /// Check if a name exists in the given scope
    pub fn contains(&self, name: &str, scope: &NameScope) -> bool {
        let key = Self::make_key(name, scope);
        self.ranges.contains_key(&key)
    }

    /// Iterate over all named ranges
    pub fn iter(&self) -> impl Iterator<Item = &NamedRange> {
        self.ranges.values()
    }

    /// Get the number of named ranges
    pub fn len(&self) -> usize {
        self.ranges.len()
    }

    /// Check if the collection is empty
    pub fn is_empty(&self) -> bool {
        self.ranges.is_empty()
    }

    /// Get all workbook-scoped names
    pub fn workbook_names(&self) -> impl Iterator<Item = &NamedRange> {
        self.ranges
            .values()
            .filter(|r| matches!(r.scope, NameScope::Workbook))
    }

    /// Get all names scoped to a specific sheet
    pub fn sheet_names(&self, sheet_index: usize) -> impl Iterator<Item = &NamedRange> {
        self.ranges
            .values()
            .filter(move |r| matches!(r.scope, NameScope::Sheet(idx) if idx == sheet_index))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_named_range_creation() {
        let nr = NamedRange::workbook_scope("TaxRate", "Sheet1!$B$1");
        assert_eq!(nr.name, "TaxRate");
        assert_eq!(nr.refers_to, "Sheet1!$B$1");
        assert_eq!(nr.scope, NameScope::Workbook);
        assert!(!nr.is_formula());
    }

    #[test]
    fn test_named_range_formula() {
        let nr = NamedRange::workbook_scope("Total", "=SUM(A1:A10)");
        assert!(nr.is_formula());
        assert_eq!(nr.expression(), "SUM(A1:A10)");
    }

    #[test]
    fn test_collection_scope_lookup() {
        let mut coll = NamedRangeCollection::new();

        // Add workbook-scoped name
        coll.define(NamedRange::workbook_scope("Rate", "0.05"))
            .unwrap();

        // Add sheet-scoped name with same name
        coll.define(NamedRange::sheet_scope("Rate", "0.08", 0))
            .unwrap();

        // Sheet 0 should find the sheet-scoped version
        let found = coll.get("Rate", 0).unwrap();
        assert_eq!(found.refers_to, "0.08");

        // Sheet 1 should find the workbook-scoped version
        let found = coll.get("Rate", 1).unwrap();
        assert_eq!(found.refers_to, "0.05");
    }

    #[test]
    fn test_case_insensitive() {
        let mut coll = NamedRangeCollection::new();
        coll.define(NamedRange::workbook_scope("TaxRate", "0.05"))
            .unwrap();

        // Should find regardless of case
        assert!(coll.get("taxrate", 0).is_some());
        assert!(coll.get("TAXRATE", 0).is_some());
        assert!(coll.get("TaxRate", 0).is_some());

        // Should not allow duplicate with different case
        assert!(coll
            .define(NamedRange::workbook_scope("TAXRATE", "0.10"))
            .is_err());
    }
}
