//! Criteria matching for SUMIF, COUNTIF, AVERAGEIF and related functions
//!
//! Excel criteria can be:
//! - A number: exact match (e.g., 5)
//! - A text string: case-insensitive match (e.g., "apple")
//! - A comparison expression: ">5", ">=10", "<100", "<=50", "<>0", "=5"
//! - Wildcards: "*" matches any characters, "?" matches single character
//! - Empty string: matches empty cells

use crate::evaluator::FormulaValue;

/// Criteria matcher for SUMIF/COUNTIF/AVERAGEIF and related functions
/// Handles comparison operators, wildcards, and exact matching
#[derive(Debug)]
pub struct CriteriaMatcher {
    criteria_type: CriteriaType,
}

#[derive(Debug)]
enum CriteriaType {
    /// Exact number match
    Number(f64),
    /// Comparison with number (operator, value)
    Comparison(ComparisonOp, f64),
    /// Text match (case-insensitive, with wildcards)
    Text(String),
    /// Match empty values
    Empty,
}

#[derive(Debug, Clone, Copy)]
enum ComparisonOp {
    Equal,
    NotEqual,
    LessThan,
    LessEqual,
    GreaterThan,
    GreaterEqual,
}

impl CriteriaMatcher {
    /// Create a new criteria matcher from a FormulaValue
    pub fn new(criteria: &FormulaValue) -> Self {
        let criteria_type = match criteria {
            FormulaValue::Number(n) => CriteriaType::Number(*n),
            FormulaValue::Boolean(b) => CriteriaType::Number(if *b { 1.0 } else { 0.0 }),
            FormulaValue::String(s) => Self::parse_string_criteria(s),
            FormulaValue::Empty => CriteriaType::Empty,
            FormulaValue::Error(_) => CriteriaType::Empty, // Errors don't match anything
            FormulaValue::Array(_) => CriteriaType::Empty, // Arrays as criteria not supported
        };

        Self { criteria_type }
    }

    fn parse_string_criteria(s: &str) -> CriteriaType {
        let s = s.trim();

        if s.is_empty() {
            return CriteriaType::Empty;
        }

        // Try to parse as comparison operator + number
        if let Some(ct) = Self::try_parse_comparison(s) {
            return ct;
        }

        // Try to parse as plain number
        if let Ok(n) = s.parse::<f64>() {
            return CriteriaType::Number(n);
        }

        // Text match (case-insensitive, supports wildcards)
        CriteriaType::Text(s.to_lowercase())
    }

    fn try_parse_comparison(s: &str) -> Option<CriteriaType> {
        // Check for comparison operators (order matters - check longer ones first)
        let (op, rest) = if s.starts_with(">=") {
            (ComparisonOp::GreaterEqual, &s[2..])
        } else if s.starts_with("<=") {
            (ComparisonOp::LessEqual, &s[2..])
        } else if s.starts_with("<>") {
            (ComparisonOp::NotEqual, &s[2..])
        } else if s.starts_with('>') {
            (ComparisonOp::GreaterThan, &s[1..])
        } else if s.starts_with('<') {
            (ComparisonOp::LessThan, &s[1..])
        } else if s.starts_with('=') {
            (ComparisonOp::Equal, &s[1..])
        } else {
            return None;
        };

        // Try to parse the number part
        let rest = rest.trim();
        if let Ok(n) = rest.parse::<f64>() {
            Some(CriteriaType::Comparison(op, n))
        } else {
            // Could be text comparison like ">A" - treat as text
            None
        }
    }

    /// Check if a value matches the criteria
    pub fn matches(&self, value: &FormulaValue) -> bool {
        match &self.criteria_type {
            CriteriaType::Number(criteria_num) => {
                // Only match actual numeric values, not strings that look like numbers
                // This matches Excel behavior where SUMIF(A:A, 5) won't match text "5"
                match value {
                    FormulaValue::Number(n) => (n - criteria_num).abs() < 1e-10,
                    FormulaValue::Boolean(b) => {
                        let n = if *b { 1.0 } else { 0.0 };
                        (n - criteria_num).abs() < 1e-10
                    }
                    _ => false,
                }
            }

            CriteriaType::Comparison(op, criteria_num) => {
                // Comparisons also only work on numeric values
                let n = match value {
                    FormulaValue::Number(n) => *n,
                    FormulaValue::Boolean(b) => {
                        if *b {
                            1.0
                        } else {
                            0.0
                        }
                    }
                    _ => return false,
                };
                match op {
                    ComparisonOp::Equal => (n - criteria_num).abs() < 1e-10,
                    ComparisonOp::NotEqual => (n - criteria_num).abs() >= 1e-10,
                    ComparisonOp::LessThan => n < *criteria_num,
                    ComparisonOp::LessEqual => n <= *criteria_num,
                    ComparisonOp::GreaterThan => n > *criteria_num,
                    ComparisonOp::GreaterEqual => n >= *criteria_num,
                }
            }

            CriteriaType::Text(pattern) => {
                let text = value.as_string().to_lowercase();
                Self::wildcard_match(pattern, &text)
            }

            CriteriaType::Empty => {
                matches!(value, FormulaValue::Empty)
                    || matches!(value, FormulaValue::String(s) if s.is_empty())
            }
        }
    }

    /// Match with wildcards: * = any characters, ? = single character
    fn wildcard_match(pattern: &str, text: &str) -> bool {
        // If no wildcards, do exact match
        if !pattern.contains('*') && !pattern.contains('?') {
            return pattern == text;
        }

        // Wildcard matching using iterative approach with backtracking
        let pattern_chars: Vec<char> = pattern.chars().collect();
        let text_chars: Vec<char> = text.chars().collect();

        Self::wildcard_match_impl(&pattern_chars, &text_chars)
    }

    fn wildcard_match_impl(pattern: &[char], text: &[char]) -> bool {
        let mut pi = 0; // pattern index
        let mut ti = 0; // text index
        let mut star_pi = None; // position of last * in pattern
        let mut star_ti = 0; // position in text when we matched last *

        while ti < text.len() {
            if pi < pattern.len() && (pattern[pi] == '?' || pattern[pi] == text[ti]) {
                // Characters match or pattern has ?
                pi += 1;
                ti += 1;
            } else if pi < pattern.len() && pattern[pi] == '*' {
                // Star found, record position
                star_pi = Some(pi);
                star_ti = ti;
                pi += 1; // Try matching * with empty string first
            } else if let Some(sp) = star_pi {
                // Mismatch, but we have a star to backtrack to
                pi = sp + 1;
                star_ti += 1;
                ti = star_ti;
            } else {
                // No match and no star to backtrack to
                return false;
            }
        }

        // Check remaining pattern characters (must all be *)
        while pi < pattern.len() && pattern[pi] == '*' {
            pi += 1;
        }

        pi == pattern.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number_criteria() {
        let matcher = CriteriaMatcher::new(&FormulaValue::Number(5.0));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));
        assert!(!matcher.matches(&FormulaValue::String("5".into())));
    }

    #[test]
    fn test_comparison_criteria() {
        // Greater than
        let matcher = CriteriaMatcher::new(&FormulaValue::String(">5".into()));
        assert!(matcher.matches(&FormulaValue::Number(6.0)));
        assert!(!matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));

        // Greater than or equal
        let matcher = CriteriaMatcher::new(&FormulaValue::String(">=5".into()));
        assert!(matcher.matches(&FormulaValue::Number(6.0)));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));

        // Less than
        let matcher = CriteriaMatcher::new(&FormulaValue::String("<5".into()));
        assert!(!matcher.matches(&FormulaValue::Number(6.0)));
        assert!(!matcher.matches(&FormulaValue::Number(5.0)));
        assert!(matcher.matches(&FormulaValue::Number(4.0)));

        // Less than or equal
        let matcher = CriteriaMatcher::new(&FormulaValue::String("<=5".into()));
        assert!(!matcher.matches(&FormulaValue::Number(6.0)));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(matcher.matches(&FormulaValue::Number(4.0)));

        // Not equal
        let matcher = CriteriaMatcher::new(&FormulaValue::String("<>5".into()));
        assert!(matcher.matches(&FormulaValue::Number(6.0)));
        assert!(!matcher.matches(&FormulaValue::Number(5.0)));
        assert!(matcher.matches(&FormulaValue::Number(4.0)));

        // Equal
        let matcher = CriteriaMatcher::new(&FormulaValue::String("=5".into()));
        assert!(!matcher.matches(&FormulaValue::Number(6.0)));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));
    }

    #[test]
    fn test_text_criteria() {
        // Case insensitive
        let matcher = CriteriaMatcher::new(&FormulaValue::String("apple".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(matcher.matches(&FormulaValue::String("APPLE".into())));
        assert!(matcher.matches(&FormulaValue::String("Apple".into())));
        assert!(!matcher.matches(&FormulaValue::String("banana".into())));
    }

    #[test]
    fn test_wildcard_criteria() {
        // Asterisk wildcard
        let matcher = CriteriaMatcher::new(&FormulaValue::String("a*".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(matcher.matches(&FormulaValue::String("a".into())));
        assert!(!matcher.matches(&FormulaValue::String("banana".into())));

        // Asterisk in middle
        let matcher = CriteriaMatcher::new(&FormulaValue::String("a*e".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(matcher.matches(&FormulaValue::String("ae".into())));
        assert!(!matcher.matches(&FormulaValue::String("apples".into())));

        // Question mark wildcard
        let matcher = CriteriaMatcher::new(&FormulaValue::String("a?ple".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(!matcher.matches(&FormulaValue::String("aple".into())));
        assert!(!matcher.matches(&FormulaValue::String("axxple".into())));

        // Combined wildcards
        let matcher = CriteriaMatcher::new(&FormulaValue::String("a?p*".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(matcher.matches(&FormulaValue::String("app".into())));
        assert!(!matcher.matches(&FormulaValue::String("ap".into())));
    }

    #[test]
    fn test_empty_criteria() {
        let matcher = CriteriaMatcher::new(&FormulaValue::String("".into()));
        assert!(matcher.matches(&FormulaValue::Empty));
        assert!(matcher.matches(&FormulaValue::String("".into())));
        assert!(!matcher.matches(&FormulaValue::String("text".into())));
        assert!(!matcher.matches(&FormulaValue::Number(0.0)));
    }
}
