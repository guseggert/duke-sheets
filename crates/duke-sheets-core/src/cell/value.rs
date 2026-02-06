//! Cell value types

use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

/// Represents the value stored in a cell
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell (no value)
    Empty,

    /// Boolean value (TRUE/FALSE)
    Boolean(bool),

    /// Numeric value (all numbers stored as f64, including dates)
    Number(f64),

    /// String value
    String(SharedString),

    /// Error value (#VALUE!, #REF!, etc.)
    Error(CellError),

    /// Formula with cached result
    Formula {
        /// Original formula text (e.g., "=SUM(A1:A10)")
        text: String,
        /// Last calculated value (if any)
        /// For dynamic array formulas, this contains the top-left value
        cached_value: Option<Box<CellValue>>,
        /// If this formula produces an array, this contains all values
        /// The outer Vec is rows, inner Vec is columns
        array_result: Option<Vec<Vec<CellValue>>>,
    },

    /// A cell that receives a spilled value from a dynamic array formula
    /// This cell cannot be edited directly - it displays a value from the source formula
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
}

impl CellValue {
    /// Create a new string value
    pub fn string<S: Into<String>>(s: S) -> Self {
        CellValue::String(SharedString::new(s.into()))
    }

    /// Create a new formula value
    pub fn formula<S: Into<String>>(text: S) -> Self {
        CellValue::Formula {
            text: text.into(),
            cached_value: None,
            array_result: None,
        }
    }

    /// Check if the cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Check if the cell contains a formula
    pub fn is_formula(&self) -> bool {
        matches!(self, CellValue::Formula { .. })
    }

    /// Check if the cell contains an error
    pub fn is_error(&self) -> bool {
        matches!(self, CellValue::Error(_))
    }

    /// Check if the cell is a spill target
    pub fn is_spill_target(&self) -> bool {
        matches!(self, CellValue::SpillTarget { .. })
    }

    /// Check if the cell contains a dynamic array formula
    pub fn is_array_formula(&self) -> bool {
        matches!(
            self,
            CellValue::Formula {
                array_result: Some(_),
                ..
            }
        )
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
            CellValue::Formula {
                cached_value: Some(v),
                ..
            } => v.as_number(),
            _ => None,
        }
    }

    /// Try to get the value as a boolean
    pub fn as_bool(&self) -> Option<bool> {
        match self {
            CellValue::Boolean(b) => Some(*b),
            CellValue::Number(n) => Some(*n != 0.0),
            CellValue::Formula {
                cached_value: Some(v),
                ..
            } => v.as_bool(),
            _ => None,
        }
    }

    /// Try to get the value as a string
    pub fn as_string(&self) -> Option<&str> {
        match self {
            CellValue::String(s) => Some(s.as_str()),
            CellValue::Formula {
                cached_value: Some(v),
                ..
            } => v.as_string(),
            _ => None,
        }
    }

    /// Get the formula text if this is a formula cell
    pub fn formula_text(&self) -> Option<&str> {
        match self {
            CellValue::Formula { text, .. } => Some(text),
            _ => None,
        }
    }

    /// Get the effective value (cached value for formulas, value otherwise)
    pub fn effective_value(&self) -> &CellValue {
        match self {
            CellValue::Formula {
                cached_value: Some(v),
                ..
            } => v.effective_value(),
            _ => self,
        }
    }

    /// Get the type name for error messages
    pub fn type_name(&self) -> &'static str {
        match self {
            CellValue::Empty => "empty",
            CellValue::Boolean(_) => "boolean",
            CellValue::Number(_) => "number",
            CellValue::String(_) => "string",
            CellValue::Error(_) => "error",
            CellValue::Formula { .. } => "formula",
            CellValue::SpillTarget { .. } => "spill_target",
        }
    }
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

impl fmt::Display for CellValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CellValue::Empty => write!(f, ""),
            CellValue::Boolean(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            CellValue::Number(n) => write!(f, "{}", n),
            CellValue::String(s) => write!(f, "{}", s.as_str()),
            CellValue::Error(e) => write!(f, "{}", e),
            CellValue::Formula {
                cached_value: Some(v),
                ..
            } => write!(f, "{}", v),
            CellValue::Formula { text, .. } => write!(f, "{}", text),
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

    /// Parse an error string
    pub fn from_str(s: &str) -> Option<Self> {
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
#[derive(Debug, Default)]
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
    fn test_cell_error_display() {
        assert_eq!(CellError::Div0.to_string(), "#DIV/0!");
        assert_eq!(CellError::Value.to_string(), "#VALUE!");
        assert_eq!(CellError::Na.to_string(), "#N/A");
    }

    #[test]
    fn test_cell_error_parse() {
        assert_eq!(CellError::from_str("#DIV/0!"), Some(CellError::Div0));
        assert_eq!(CellError::from_str("#VALUE!"), Some(CellError::Value));
        assert_eq!(CellError::from_str("#n/a"), Some(CellError::Na)); // Case insensitive
        assert_eq!(CellError::from_str("invalid"), None);
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
}
