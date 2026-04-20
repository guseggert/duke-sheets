//! Evaluation cache for a single `calculate()` pass.
//!
//! Provides three layers of caching:
//! - **Tier 1**: Range materialization cache - avoids re-reading cells for the same range
//! - **Tier 2**: Lookup hash index - O(1) exact-match lookups for MATCH/VLOOKUP/XLOOKUP
//! - **Tier 3**: Sheet name → index cache - avoids linear scan for cross-sheet references

use std::sync::Arc;

use ahash::AHashMap;
use dashmap::DashMap;

use crate::evaluator::FormulaValue;
use duke_sheets_core::{CellError, CellValue};

/// Key for a materialized range: (sheet_idx, start_row, start_col, end_row, end_col).
pub type RangeKey = (usize, u32, u16, u32, u16);

/// Key for a lookup index: range key + column offset within that range.
pub type LookupIndexKey = (RangeKey, u16);

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum LookupSourceKey {
    Column {
        sheet: usize,
        row_start: u32,
        row_end: u32,
        col: u16,
    },
    Row {
        sheet: usize,
        row: u32,
        col_start: u16,
        col_end: u16,
    },
}

/// Shared cache for a single `calculate()` pass.
///
/// Created once at the start of `execute_eval_plan()`, passed by reference
/// into every `EvaluationContext`. Dropped when the eval pass completes.
///
/// All fields use `DashMap` for lock-free concurrent access from the
/// parallel (rayon) evaluation path.
pub struct EvalCache {
    /// Tier 1: Materialized range values.
    /// First call to `get_range_values()` for a given range key materializes
    /// and stores the data; subsequent calls return an `Arc` clone.
    pub ranges: DashMap<RangeKey, Arc<Vec<Vec<FormulaValue>>>>,

    /// Tier 2: Hash index on a 1D lookup column for exact-match lookups.
    /// Key is (range_key, column_offset). Built lazily on first exact-match
    /// MATCH/VLOOKUP/XLOOKUP against that column.
    pub lookup_indexes: DashMap<LookupIndexKey, Arc<LookupIndex>>,

    /// Tier 2b: Hash index built directly from worksheet source coordinates.
    /// Used by evaluator fast paths to avoid materializing lookup vectors.
    pub source_lookup_indexes: DashMap<LookupSourceKey, Arc<LookupIndex>>,

    /// Tier 3: Sheet name → index mapping.
    /// Pre-built before the eval loop; immutable during evaluation.
    pub sheet_names: AHashMap<String, usize>,
}

impl EvalCache {
    /// Create a new cache with a pre-built sheet name index.
    pub fn new(sheet_names: AHashMap<String, usize>) -> Self {
        Self {
            ranges: DashMap::new(),
            lookup_indexes: DashMap::new(),
            source_lookup_indexes: DashMap::new(),
            sheet_names,
        }
    }
}

/// Hash index on a 1D vector of FormulaValues for O(1) exact-match lookup.
///
/// Maps normalized lookup keys to the index of their *first* occurrence
/// (matching MATCH's "return first match" semantics).
pub struct LookupIndex {
    index: AHashMap<LookupKey, usize>,
}

impl LookupIndex {
    /// Build a hash index from a 1D slice of FormulaValues.
    /// Stores the position of the *first* occurrence of each distinct value.
    pub fn build(values: &[FormulaValue]) -> Self {
        let mut index = AHashMap::with_capacity(values.len());
        for (i, v) in values.iter().enumerate() {
            let key = LookupKey::from(v);
            // Only insert if not present - first occurrence wins
            index.entry(key).or_insert(i);
        }
        Self { index }
    }

    pub fn build_from_refs<'a, I>(values: I) -> Self
    where
        I: IntoIterator<Item = &'a FormulaValue>,
    {
        let mut index = AHashMap::new();
        for (i, v) in values.into_iter().enumerate() {
            let key = LookupKey::from(v);
            index.entry(key).or_insert(i);
        }
        Self { index }
    }

    pub fn build_from_cell_refs<'a, I>(values: I) -> Self
    where
        I: IntoIterator<Item = &'a CellValue>,
    {
        let mut index = AHashMap::new();
        for (i, v) in values.into_iter().enumerate() {
            let key = LookupKey::from_cell_value(v);
            index.entry(key).or_insert(i);
        }
        Self { index }
    }

    /// Look up a value and return its 0-based index, or None if not found.
    pub fn find(&self, value: &FormulaValue) -> Option<usize> {
        let key = LookupKey::from(value);
        self.index.get(&key).copied()
    }
}

/// Hashable key that matches `values_equal()` semantics.
///
/// - Numbers: f64 bits with -0.0 normalized to +0.0, NaN to a sentinel
/// - Strings: lowercased for case-insensitive matching
/// - Booleans, Errors, Empty: direct equality
#[derive(Clone, Debug)]
enum LookupKey {
    Number(u64),
    Text(String),
    Boolean(bool),
    Error(CellError),
    Empty,
}

impl LookupKey {
    fn from(value: &FormulaValue) -> Self {
        match value {
            FormulaValue::Number(n) => {
                let normalized = if *n == 0.0 { 0.0f64 } else { *n };
                if normalized.is_nan() {
                    // All NaNs hash to the same sentinel
                    LookupKey::Number(u64::MAX)
                } else {
                    LookupKey::Number(normalized.to_bits())
                }
            }
            FormulaValue::String(s) => LookupKey::Text(s.to_lowercase()),
            FormulaValue::Boolean(b) => LookupKey::Boolean(*b),
            FormulaValue::Error(e) => LookupKey::Error(*e),
            FormulaValue::Empty => LookupKey::Empty,
            FormulaValue::Array { .. } => LookupKey::Empty, // arrays in lookup columns shouldn't happen
        }
    }

    fn from_cell_value(value: &CellValue) -> Self {
        match value {
            CellValue::Number(n) => {
                let normalized = if *n == 0.0 { 0.0f64 } else { *n };
                if normalized.is_nan() {
                    LookupKey::Number(u64::MAX)
                } else {
                    LookupKey::Number(normalized.to_bits())
                }
            }
            CellValue::String(s) => LookupKey::Text(s.as_str().to_lowercase()),
            CellValue::Boolean(b) => LookupKey::Boolean(*b),
            CellValue::Error(e) => LookupKey::Error(*e),
            CellValue::Empty | CellValue::SpillTarget { .. } | CellValue::RichText(_) => {
                LookupKey::Empty
            }
        }
    }
}

impl PartialEq for LookupKey {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (LookupKey::Number(a), LookupKey::Number(b)) => a == b,
            (LookupKey::Text(a), LookupKey::Text(b)) => a == b,
            (LookupKey::Boolean(a), LookupKey::Boolean(b)) => a == b,
            (LookupKey::Error(a), LookupKey::Error(b)) => a == b,
            (LookupKey::Empty, LookupKey::Empty) => true,
            _ => false,
        }
    }
}

impl Eq for LookupKey {}

impl std::hash::Hash for LookupKey {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        std::mem::discriminant(self).hash(state);
        match self {
            LookupKey::Number(bits) => bits.hash(state),
            LookupKey::Text(s) => s.hash(state),
            LookupKey::Boolean(b) => b.hash(state),
            LookupKey::Error(e) => format!("{:?}", e).hash(state),
            LookupKey::Empty => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lookup_key_case_insensitive() {
        let a = LookupKey::from(&FormulaValue::String("Hello".to_string()));
        let b = LookupKey::from(&FormulaValue::String("HELLO".to_string()));
        assert_eq!(a, b);
    }

    #[test]
    fn lookup_key_neg_zero() {
        let a = LookupKey::from(&FormulaValue::Number(0.0));
        let b = LookupKey::from(&FormulaValue::Number(-0.0));
        assert_eq!(a, b);
    }

    #[test]
    fn lookup_key_nan_consistent() {
        let a = LookupKey::from(&FormulaValue::Number(f64::NAN));
        let b = LookupKey::from(&FormulaValue::Number(f64::NAN));
        assert_eq!(a, b);
    }

    #[test]
    fn lookup_key_different_types() {
        let num = LookupKey::from(&FormulaValue::Number(42.0));
        let text = LookupKey::from(&FormulaValue::String("42".to_string()));
        assert_ne!(num, text);
    }

    #[test]
    fn lookup_index_first_occurrence_wins() {
        let values = vec![
            FormulaValue::String("a".to_string()),
            FormulaValue::String("b".to_string()),
            FormulaValue::String("a".to_string()), // duplicate
            FormulaValue::String("c".to_string()),
        ];
        let idx = LookupIndex::build(&values);
        assert_eq!(idx.find(&FormulaValue::String("a".to_string())), Some(0));
        assert_eq!(idx.find(&FormulaValue::String("A".to_string())), Some(0)); // case-insensitive
        assert_eq!(idx.find(&FormulaValue::String("b".to_string())), Some(1));
        assert_eq!(idx.find(&FormulaValue::String("c".to_string())), Some(3));
        assert_eq!(idx.find(&FormulaValue::String("d".to_string())), None);
    }

    #[test]
    fn lookup_index_mixed_types() {
        let values = vec![
            FormulaValue::Number(1.0),
            FormulaValue::String("text".to_string()),
            FormulaValue::Boolean(true),
            FormulaValue::Empty,
        ];
        let idx = LookupIndex::build(&values);
        assert_eq!(idx.find(&FormulaValue::Number(1.0)), Some(0));
        assert_eq!(idx.find(&FormulaValue::String("TEXT".to_string())), Some(1));
        assert_eq!(idx.find(&FormulaValue::Boolean(true)), Some(2));
        assert_eq!(idx.find(&FormulaValue::Empty), Some(3));
    }

    #[test]
    fn lookup_index_empty_input() {
        let idx = LookupIndex::build(&[]);
        assert_eq!(idx.find(&FormulaValue::Number(1.0)), None);
    }
}
