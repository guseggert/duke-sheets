//! Cell value types

use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

use crate::rich_text::{rich_text_to_plain, RichTextRun};

/// Formula data stored in a side table, separate from cell values.
///
/// Formula cells store their cached result directly in the cell grid as a
/// regular `CellValue` (Number, String, Error, etc.). The formula text and
/// array results live here, indexed by `(row, col)` in `CellStorage::formulas`.
#[derive(Debug, Clone)]
pub struct FormulaData {
    /// Original formula text (e.g., "=SUM(A1:A10)")
    pub text: String,
    /// If this formula produces a dynamic array, the full result lives here.
    /// The outer Vec is rows, inner Vec is columns.
    pub array_result: Option<Vec<Vec<CellValue>>>,
}

impl FormulaData {
    /// Create a new formula with no array result.
    pub fn new<S: Into<String>>(text: S) -> Self {
        Self {
            text: text.into(),
            array_result: None,
        }
    }

    /// Check if this formula has a dynamic array result.
    pub fn is_array_formula(&self) -> bool {
        self.array_result.is_some()
    }
}

/// Represents the value stored in a cell.
///
/// This enum holds only leaf values - no formula text or metadata.
/// Formulas are stored separately in [`CellStorage`](super::CellStorage).
#[derive(Debug, Clone, PartialEq, Default)]
pub enum CellValue {
    /// Empty cell (no value)
    #[default]
    Empty,

    /// Boolean value (TRUE/FALSE)
    Boolean(bool),

    /// Numeric value (all numbers stored as f64, including dates)
    Number(f64),

    /// String value
    String(SharedString),

    /// Error value (#VALUE!, #REF!, etc.)
    Error(CellError),

    /// A cell that receives a spilled value from a dynamic array formula.
    /// This cell cannot be edited directly - it displays a value from the source formula.
    SpillTarget {
        /// Row of the source formula cell
        source_row: u32,
        /// Column of the source formula cell
        source_col: u16,
        /// Row offset from source (0 for first row of spill)
        offset_row: u32,
        /// Column offset from source (0 for first column of spill)
        offset_col: u16,
    },

    /// Rich text with per-run character formatting (boxed to keep enum small)
    RichText(Box<Vec<RichTextRun>>),
}

impl CellValue {
    /// Create a new string value
    pub fn string<S: Into<String>>(s: S) -> Self {
        CellValue::String(SharedString::new(s.into()))
    }

    /// Create a new rich text value from runs.
    pub fn rich_text(runs: Vec<RichTextRun>) -> Self {
        CellValue::RichText(Box::new(runs))
    }

    /// Check if the cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Check if the cell is blank (empty, empty string, or empty rich text)
    pub fn is_blank(&self) -> bool {
        match self {
            CellValue::Empty => true,
            CellValue::String(s) => s.as_str().is_empty(),
            CellValue::RichText(runs) => rich_text_to_plain(runs).is_empty(),
            _ => false,
        }
    }

    /// Check if the cell contains rich text
    pub fn is_rich_text(&self) -> bool {
        matches!(self, CellValue::RichText(_))
    }

    /// Check if the cell contains an error
    pub fn is_error(&self) -> bool {
        matches!(self, CellValue::Error(_))
    }

    /// Check if the cell is a spill target
    pub fn is_spill_target(&self) -> bool {
        matches!(self, CellValue::SpillTarget { .. })
    }

    /// Get the spill source coordinates if this is a spill target
    pub fn spill_source(&self) -> Option<(u32, u16)> {
        match self {
            CellValue::SpillTarget {
                source_row,
                source_col,
                ..
            } => Some((*source_row, *source_col)),
            _ => None,
        }
    }

    /// Try to get the value as a number
    pub fn as_number(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            CellValue::Boolean(true) => Some(1.0),
            CellValue::Boolean(false) => Some(0.0),
            _ => None,
        }
    }

    /// Try to get the value as a boolean
    pub fn as_bool(&self) -> Option<bool> {
        match self {
            CellValue::Boolean(b) => Some(*b),
            CellValue::Number(n) => Some(*n != 0.0),
            _ => None,
        }
    }

    /// Try to get the value as a string
    pub fn as_string(&self) -> Option<&str> {
        match self {
            CellValue::String(s) => Some(s.as_str()),
            _ => None,
        }
    }

    /// Get the effective value - for leaf values, returns self.
    ///
    /// (Formulas are now stored separately; this is identity for all variants.)
    pub fn effective_value(&self) -> &CellValue {
        self
    }

    /// Get the type name for error messages
    pub fn type_name(&self) -> &'static str {
        match self {
            CellValue::Empty => "empty",
            CellValue::Boolean(_) => "boolean",
            CellValue::Number(_) => "number",
            CellValue::String(_) => "string",
            CellValue::RichText(_) => "rich_text",
            CellValue::Error(_) => "error",
            CellValue::SpillTarget { .. } => "spill_target",
        }
    }
}

impl fmt::Display for CellValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CellValue::Empty => write!(f, ""),
            CellValue::Boolean(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            CellValue::Number(n) => write!(f, "{}", n),
            CellValue::String(s) => write!(f, "{}", s.as_str()),
            CellValue::RichText(runs) => write!(f, "{}", rich_text_to_plain(runs)),
            CellValue::Error(e) => write!(f, "{}", e),
            // SpillTarget shows as empty - the actual value comes from looking up the source
            CellValue::SpillTarget { .. } => write!(f, ""),
        }
    }
}

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Boolean(b)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<i64> for CellValue {
    fn from(n: i64) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::string(s)
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::string(s)
    }
}

impl From<CellError> for CellValue {
    fn from(e: CellError) -> Self {
        CellValue::Error(e)
    }
}

/// Excel error values
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CellError {
    /// #NULL! - Incorrect range operator
    Null,
    /// #DIV/0! - Division by zero
    Div0,
    /// #VALUE! - Wrong type of argument or operand
    Value,
    /// #REF! - Invalid cell reference
    Ref,
    /// #NAME? - Unrecognized formula name
    Name,
    /// #NUM! - Invalid numeric value
    Num,
    /// #N/A - Value not available
    Na,
    /// #GETTING_DATA - External data is loading
    GettingData,
    /// #SPILL! - Dynamic array cannot spill
    Spill,
    /// #CALC! - Calculation error
    Calc,
}

impl CellError {
    /// Get the display string for this error
    pub fn as_str(&self) -> &'static str {
        match self {
            CellError::Null => "#NULL!",
            CellError::Div0 => "#DIV/0!",
            CellError::Value => "#VALUE!",
            CellError::Ref => "#REF!",
            CellError::Name => "#NAME?",
            CellError::Num => "#NUM!",
            CellError::Na => "#N/A",
            CellError::GettingData => "#GETTING_DATA",
            CellError::Spill => "#SPILL!",
            CellError::Calc => "#CALC!",
        }
    }

    /// Parse an error string into a `CellError`.
    pub fn parse(s: &str) -> Option<Self> {
        match s.to_uppercase().as_str() {
            "#NULL!" => Some(CellError::Null),
            "#DIV/0!" => Some(CellError::Div0),
            "#VALUE!" => Some(CellError::Value),
            "#REF!" => Some(CellError::Ref),
            "#NAME?" => Some(CellError::Name),
            "#NUM!" => Some(CellError::Num),
            "#N/A" => Some(CellError::Na),
            "#GETTING_DATA" => Some(CellError::GettingData),
            "#SPILL!" => Some(CellError::Spill),
            "#CALC!" => Some(CellError::Calc),
            _ => None,
        }
    }

    /// Get the numeric error code (for BIFF format)
    pub fn code(&self) -> u8 {
        match self {
            CellError::Null => 0x00,
            CellError::Div0 => 0x07,
            CellError::Value => 0x0F,
            CellError::Ref => 0x17,
            CellError::Name => 0x1D,
            CellError::Num => 0x24,
            CellError::Na => 0x2A,
            CellError::GettingData => 0x2B,
            CellError::Spill => 0x2C,
            CellError::Calc => 0x2D,
        }
    }
}

impl fmt::Display for CellError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

/// Interned string for memory efficiency
///
/// Strings are often repeated across cells (e.g., "Yes", "No", dates).
/// Using Arc<str> allows sharing the same string data across multiple cells.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct SharedString(Arc<str>);

impl SharedString {
    /// Create a new shared string
    pub fn new<S: AsRef<str>>(s: S) -> Self {
        SharedString(Arc::from(s.as_ref()))
    }

    /// Get the string slice
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Get the length of the string
    pub fn len(&self) -> usize {
        self.0.len()
    }

    /// Check if the string is empty
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

impl fmt::Debug for SharedString {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:?}", self.0)
    }
}

impl fmt::Display for SharedString {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl AsRef<str> for SharedString {
    fn as_ref(&self) -> &str {
        &self.0
    }
}

impl From<&str> for SharedString {
    fn from(s: &str) -> Self {
        SharedString::new(s)
    }
}

impl From<String> for SharedString {
    fn from(s: String) -> Self {
        SharedString::new(s)
    }
}

/// String pool for deduplicating strings
///
/// When reading large spreadsheets, many cells often contain the same string values.
/// The string pool ensures each unique string is stored only once in memory.
#[derive(Debug, Default, Clone)]
pub struct StringPool {
    strings: HashMap<Arc<str>, SharedString>,
}

impl StringPool {
    /// Create a new empty string pool
    pub fn new() -> Self {
        Self::default()
    }

    /// Get or create a shared string
    ///
    /// If the string already exists in the pool, returns a clone of the existing SharedString.
    /// Otherwise, creates a new SharedString and adds it to the pool.
    pub fn intern<S: AsRef<str>>(&mut self, s: S) -> SharedString {
        let s = s.as_ref();
        if let Some(shared) = self.strings.get(s) {
            shared.clone()
        } else {
            let arc: Arc<str> = Arc::from(s);
            let shared = SharedString(arc.clone());
            self.strings.insert(arc, shared.clone());
            shared
        }
    }

    /// Get the number of unique strings in the pool
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Check if the pool is empty
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Clear all strings from the pool
    pub fn clear(&mut self) {
        self.strings.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_value_conversions() {
        assert_eq!(CellValue::from(42), CellValue::Number(42.0));
        assert_eq!(CellValue::from(3.14), CellValue::Number(3.14));
        assert_eq!(CellValue::from(true), CellValue::Boolean(true));

        let s = CellValue::from("hello");
        assert_eq!(s.as_string(), Some("hello"));
    }

    #[test]
    fn test_cell_value_as_number() {
        assert_eq!(CellValue::Number(42.0).as_number(), Some(42.0));
        assert_eq!(CellValue::Boolean(true).as_number(), Some(1.0));
        assert_eq!(CellValue::Boolean(false).as_number(), Some(0.0));
        assert_eq!(CellValue::string("hello").as_number(), None);
        assert_eq!(CellValue::Empty.as_number(), None);
    }

    #[test]
    fn test_cell_value_is_blank() {
        assert!(CellValue::Empty.is_blank());
        assert!(CellValue::Empty.is_empty());

        assert!(CellValue::string("").is_blank());
        assert!(!CellValue::string("").is_empty());

        assert!(!CellValue::string("hello").is_blank());
        assert!(!CellValue::Number(0.0).is_blank());
        assert!(!CellValue::Boolean(false).is_blank());
        assert!(CellValue::rich_text(vec![]).is_blank());
    }

    #[test]
    fn test_cell_error_display() {
        assert_eq!(CellError::Div0.to_string(), "#DIV/0!");
        assert_eq!(CellError::Value.to_string(), "#VALUE!");
        assert_eq!(CellError::Na.to_string(), "#N/A");
    }

    #[test]
    fn test_cell_error_parse() {
        assert_eq!(CellError::parse("#DIV/0!"), Some(CellError::Div0));
        assert_eq!(CellError::parse("#VALUE!"), Some(CellError::Value));
        assert_eq!(CellError::parse("#n/a"), Some(CellError::Na)); // Case insensitive
        assert_eq!(CellError::parse("invalid"), None);
    }

    #[test]
    fn test_string_pool() {
        let mut pool = StringPool::new();

        let s1 = pool.intern("hello");
        let s2 = pool.intern("hello");
        let s3 = pool.intern("world");

        // Same string should return same SharedString
        assert!(Arc::ptr_eq(&s1.0, &s2.0));

        // Different strings should be different
        assert!(!Arc::ptr_eq(&s1.0, &s3.0));

        assert_eq!(pool.len(), 2);
    }

    #[test]
    fn test_formula_data() {
        let f = FormulaData::new("=SUM(A1:A10)");
        assert_eq!(f.text, "=SUM(A1:A10)");
        assert!(!f.is_array_formula());

        let mut f2 = FormulaData::new("=A1:A3");
        f2.array_result = Some(vec![
            vec![CellValue::Number(1.0)],
            vec![CellValue::Number(2.0)],
        ]);
        assert!(f2.is_array_formula());
    }
}
