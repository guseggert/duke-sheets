//! Data validation
//!
//! This module provides support for data validation rules in worksheets.
//! Data validation restricts the type and range of data that users can enter
//! in cells.
//!
//! ## Example
//!
//! ```rust
//! use duke_sheets_core::{Workbook, DataValidation, ValidationType, CellRange};
//!
//! let mut workbook = Workbook::new();
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! // Add a dropdown list validation to A1:A10
//! let validation = DataValidation::list("Yes,No,Maybe")
//!     .with_range(CellRange::parse("A1:A10").unwrap())
//!     .with_error_message("Invalid value", "Please select from the list");
//!
//! sheet.add_data_validation(validation);
//! ```

use crate::cell::CellRange;

/// Data validation rule for cells
///
/// Controls what data can be entered into cells, and can display
/// input messages and error alerts.
#[derive(Debug, Clone, PartialEq)]
pub struct DataValidation {
    /// Type of validation
    pub validation_type: ValidationType,
    /// Cell ranges this validation applies to
    pub ranges: Vec<CellRange>,
    /// Allow blank/empty cells
    pub allow_blank: bool,
    /// Show dropdown for list validation
    pub show_dropdown: bool,

    // Input message (shown when cell is selected)
    /// Show input message when cell is selected
    pub show_input_message: bool,
    /// Input message title
    pub input_title: Option<String>,
    /// Input message text
    pub input_message: Option<String>,

    // Error alert (shown when invalid data entered)
    /// Show error alert when invalid data entered
    pub show_error_alert: bool,
    /// Error alert style
    pub error_style: ValidationErrorStyle,
    /// Error alert title
    pub error_title: Option<String>,
    /// Error alert message
    pub error_message: Option<String>,
}

impl Default for DataValidation {
    fn default() -> Self {
        Self {
            validation_type: ValidationType::None,
            ranges: Vec::new(),
            allow_blank: true,
            show_dropdown: true,
            show_input_message: false,
            input_title: None,
            input_message: None,
            show_error_alert: true,
            error_style: ValidationErrorStyle::Stop,
            error_title: None,
            error_message: None,
        }
    }
}

impl DataValidation {
    /// Create a new data validation with no restrictions
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a list validation (dropdown)
    ///
    /// # Arguments
    ///
    /// * `source` - Either a comma-separated list of values (e.g., "Yes,No,Maybe")
    ///              or a range reference (e.g., "Sheet1!$A$1:$A$10")
    ///
    /// # Example
    ///
    /// ```rust
    /// use duke_sheets_core::DataValidation;
    ///
    /// // Inline list
    /// let v1 = DataValidation::list("Red,Green,Blue");
    ///
    /// // Range reference
    /// let v2 = DataValidation::list("=Sheet1!$A$1:$A$5");
    /// ```
    pub fn list(source: impl Into<String>) -> Self {
        Self {
            validation_type: ValidationType::List {
                source: source.into(),
            },
            ..Self::default()
        }
    }

    /// Create a whole number validation
    pub fn whole_number(operator: ValidationOperator, value1: impl Into<String>) -> Self {
        Self {
            validation_type: ValidationType::Whole {
                operator,
                value1: value1.into(),
                value2: None,
            },
            ..Self::default()
        }
    }

    /// Create a whole number validation with between/not between operator
    pub fn whole_number_between(
        operator: ValidationOperator,
        value1: impl Into<String>,
        value2: impl Into<String>,
    ) -> Self {
        Self {
            validation_type: ValidationType::Whole {
                operator,
                value1: value1.into(),
                value2: Some(value2.into()),
            },
            ..Self::default()
        }
    }

    /// Create a decimal number validation
    pub fn decimal(operator: ValidationOperator, value1: impl Into<String>) -> Self {
        Self {
            validation_type: ValidationType::Decimal {
                operator,
                value1: value1.into(),
                value2: None,
            },
            ..Self::default()
        }
    }

    /// Create a decimal number validation with between/not between operator
    pub fn decimal_between(
        operator: ValidationOperator,
        value1: impl Into<String>,
        value2: impl Into<String>,
    ) -> Self {
        Self {
            validation_type: ValidationType::Decimal {
                operator,
                value1: value1.into(),
                value2: Some(value2.into()),
            },
            ..Self::default()
        }
    }

    /// Create a date validation
    pub fn date(operator: ValidationOperator, value1: impl Into<String>) -> Self {
        Self {
            validation_type: ValidationType::Date {
                operator,
                value1: value1.into(),
                value2: None,
            },
            ..Self::default()
        }
    }

    /// Create a time validation
    pub fn time(operator: ValidationOperator, value1: impl Into<String>) -> Self {
        Self {
            validation_type: ValidationType::Time {
                operator,
                value1: value1.into(),
                value2: None,
            },
            ..Self::default()
        }
    }

    /// Create a text length validation
    pub fn text_length(operator: ValidationOperator, value1: impl Into<String>) -> Self {
        Self {
            validation_type: ValidationType::TextLength {
                operator,
                value1: value1.into(),
                value2: None,
            },
            ..Self::default()
        }
    }

    /// Create a custom formula validation
    ///
    /// # Arguments
    ///
    /// * `formula` - A formula that returns TRUE for valid values, FALSE for invalid
    ///
    /// # Example
    ///
    /// ```rust
    /// use duke_sheets_core::DataValidation;
    ///
    /// // Only allow values that are multiples of 5
    /// let v = DataValidation::custom("=MOD(A1,5)=0");
    /// ```
    pub fn custom(formula: impl Into<String>) -> Self {
        Self {
            validation_type: ValidationType::Custom {
                formula: formula.into(),
            },
            ..Self::default()
        }
    }

    /// Add a cell range to this validation
    pub fn with_range(mut self, range: CellRange) -> Self {
        self.ranges.push(range);
        self
    }

    /// Set the cell ranges for this validation
    pub fn with_ranges(mut self, ranges: Vec<CellRange>) -> Self {
        self.ranges = ranges;
        self
    }

    /// Set whether blank cells are allowed
    pub fn with_allow_blank(mut self, allow: bool) -> Self {
        self.allow_blank = allow;
        self
    }

    /// Set whether to show dropdown for list validation
    pub fn with_dropdown(mut self, show: bool) -> Self {
        self.show_dropdown = show;
        self
    }

    /// Set an input message (shown when cell is selected)
    pub fn with_input_message(
        mut self,
        title: impl Into<String>,
        message: impl Into<String>,
    ) -> Self {
        self.show_input_message = true;
        self.input_title = Some(title.into());
        self.input_message = Some(message.into());
        self
    }

    /// Set an error message (shown when invalid data entered)
    pub fn with_error_message(
        mut self,
        title: impl Into<String>,
        message: impl Into<String>,
    ) -> Self {
        self.show_error_alert = true;
        self.error_title = Some(title.into());
        self.error_message = Some(message.into());
        self
    }

    /// Set the error style
    pub fn with_error_style(mut self, style: ValidationErrorStyle) -> Self {
        self.error_style = style;
        self
    }

    /// Check if this validation applies to a specific cell
    pub fn applies_to(&self, row: u32, col: u16) -> bool {
        self.ranges.iter().any(|r| {
            row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col
        })
    }
}

/// Types of data validation
#[derive(Debug, Clone, PartialEq)]
pub enum ValidationType {
    /// No validation (any value allowed)
    None,

    /// Must be a whole number
    Whole {
        operator: ValidationOperator,
        value1: String,
        value2: Option<String>,
    },

    /// Must be a decimal number
    Decimal {
        operator: ValidationOperator,
        value1: String,
        value2: Option<String>,
    },

    /// Must be from a list
    List {
        /// Either comma-separated values or a range reference
        source: String,
    },

    /// Must be a date
    Date {
        operator: ValidationOperator,
        value1: String,
        value2: Option<String>,
    },

    /// Must be a time
    Time {
        operator: ValidationOperator,
        value1: String,
        value2: Option<String>,
    },

    /// Text length constraint
    TextLength {
        operator: ValidationOperator,
        value1: String,
        value2: Option<String>,
    },

    /// Custom formula validation
    Custom {
        /// Formula that returns TRUE/FALSE
        formula: String,
    },
}

impl Default for ValidationType {
    fn default() -> Self {
        Self::None
    }
}

impl ValidationType {
    /// Get the XLSX type string for this validation type
    pub fn xlsx_type(&self) -> &'static str {
        match self {
            ValidationType::None => "none",
            ValidationType::Whole { .. } => "whole",
            ValidationType::Decimal { .. } => "decimal",
            ValidationType::List { .. } => "list",
            ValidationType::Date { .. } => "date",
            ValidationType::Time { .. } => "time",
            ValidationType::TextLength { .. } => "textLength",
            ValidationType::Custom { .. } => "custom",
        }
    }
}

/// Comparison operators for validation
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationOperator {
    /// Value must be between value1 and value2
    #[default]
    Between,
    /// Value must NOT be between value1 and value2
    NotBetween,
    /// Value must equal value1
    Equal,
    /// Value must NOT equal value1
    NotEqual,
    /// Value must be greater than value1
    GreaterThan,
    /// Value must be less than value1
    LessThan,
    /// Value must be greater than or equal to value1
    GreaterThanOrEqual,
    /// Value must be less than or equal to value1
    LessThanOrEqual,
}

impl ValidationOperator {
    /// Get the XLSX operator string
    pub fn xlsx_operator(&self) -> &'static str {
        match self {
            ValidationOperator::Between => "between",
            ValidationOperator::NotBetween => "notBetween",
            ValidationOperator::Equal => "equal",
            ValidationOperator::NotEqual => "notEqual",
            ValidationOperator::GreaterThan => "greaterThan",
            ValidationOperator::LessThan => "lessThan",
            ValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
            ValidationOperator::LessThanOrEqual => "lessThanOrEqual",
        }
    }

    /// Parse from XLSX operator string
    pub fn from_xlsx(s: &str) -> Option<Self> {
        match s {
            "between" => Some(ValidationOperator::Between),
            "notBetween" => Some(ValidationOperator::NotBetween),
            "equal" => Some(ValidationOperator::Equal),
            "notEqual" => Some(ValidationOperator::NotEqual),
            "greaterThan" => Some(ValidationOperator::GreaterThan),
            "lessThan" => Some(ValidationOperator::LessThan),
            "greaterThanOrEqual" => Some(ValidationOperator::GreaterThanOrEqual),
            "lessThanOrEqual" => Some(ValidationOperator::LessThanOrEqual),
            _ => None,
        }
    }

    /// Check if this operator requires two values
    pub fn requires_two_values(&self) -> bool {
        matches!(
            self,
            ValidationOperator::Between | ValidationOperator::NotBetween
        )
    }
}

/// Error alert styles
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationErrorStyle {
    /// Reject invalid data (default)
    #[default]
    Stop,
    /// Warn but allow
    Warning,
    /// Just inform
    Information,
}

impl ValidationErrorStyle {
    /// Get the XLSX error style string
    pub fn xlsx_style(&self) -> &'static str {
        match self {
            ValidationErrorStyle::Stop => "stop",
            ValidationErrorStyle::Warning => "warning",
            ValidationErrorStyle::Information => "information",
        }
    }

    /// Parse from XLSX style string
    pub fn from_xlsx(s: &str) -> Option<Self> {
        match s {
            "stop" => Some(ValidationErrorStyle::Stop),
            "warning" => Some(ValidationErrorStyle::Warning),
            "information" => Some(ValidationErrorStyle::Information),
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_list_validation() {
        let v = DataValidation::list("Yes,No,Maybe");
        assert!(matches!(v.validation_type, ValidationType::List { .. }));
        if let ValidationType::List { source } = &v.validation_type {
            assert_eq!(source, "Yes,No,Maybe");
        }
    }

    #[test]
    fn test_whole_number_validation() {
        let v = DataValidation::whole_number(ValidationOperator::GreaterThan, "0");
        if let ValidationType::Whole {
            operator, value1, ..
        } = &v.validation_type
        {
            assert_eq!(*operator, ValidationOperator::GreaterThan);
            assert_eq!(value1, "0");
        } else {
            panic!("Expected Whole validation type");
        }
    }

    #[test]
    fn test_between_validation() {
        let v = DataValidation::whole_number_between(ValidationOperator::Between, "1", "100");
        if let ValidationType::Whole {
            operator,
            value1,
            value2,
        } = &v.validation_type
        {
            assert_eq!(*operator, ValidationOperator::Between);
            assert_eq!(value1, "1");
            assert_eq!(value2.as_deref(), Some("100"));
        } else {
            panic!("Expected Whole validation type");
        }
    }

    #[test]
    fn test_with_messages() {
        let v = DataValidation::list("A,B,C")
            .with_input_message("Choose", "Select a value from the list")
            .with_error_message("Error", "Invalid selection");

        assert!(v.show_input_message);
        assert_eq!(v.input_title.as_deref(), Some("Choose"));
        assert_eq!(
            v.input_message.as_deref(),
            Some("Select a value from the list")
        );
        assert!(v.show_error_alert);
        assert_eq!(v.error_title.as_deref(), Some("Error"));
        assert_eq!(v.error_message.as_deref(), Some("Invalid selection"));
    }

    #[test]
    fn test_applies_to() {
        let v = DataValidation::list("A,B").with_range(CellRange::parse("A1:C10").unwrap());

        assert!(v.applies_to(0, 0)); // A1
        assert!(v.applies_to(5, 2)); // C6
        assert!(v.applies_to(9, 0)); // A10
        assert!(!v.applies_to(10, 0)); // A11 - out of range
        assert!(!v.applies_to(0, 3)); // D1 - out of range
    }

    #[test]
    fn test_operator_xlsx_strings() {
        assert_eq!(ValidationOperator::Between.xlsx_operator(), "between");
        assert_eq!(
            ValidationOperator::GreaterThan.xlsx_operator(),
            "greaterThan"
        );
        assert_eq!(
            ValidationOperator::from_xlsx("lessThanOrEqual"),
            Some(ValidationOperator::LessThanOrEqual)
        );
    }
}
