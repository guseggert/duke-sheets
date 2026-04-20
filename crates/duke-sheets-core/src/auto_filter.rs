//! Auto-filter support
//!
//! This module provides support for the auto-filter (dropdown filter)
//! feature on worksheet columns. A standalone auto-filter is attached
//! to a worksheet range and lets users filter rows by column values.
//!
//! Tables embed their own auto-filter; this module covers the
//! standalone `<autoFilter>` element in the worksheet XML.

use crate::cell::CellRange;

/// A standalone auto-filter attached to a worksheet range.
///
/// Corresponds to the `<autoFilter>` element in worksheet XML.
#[derive(Debug, Clone, PartialEq)]
pub struct AutoFilter {
    /// The range the auto-filter covers (including header row).
    pub range: CellRange,
    /// Per-column filter criteria (keyed by 0-based column offset).
    pub filter_columns: Vec<FilterColumn>,
}

impl AutoFilter {
    /// Create a new auto-filter for the given range with no column filters.
    pub fn new(range: CellRange) -> Self {
        Self {
            range,
            filter_columns: Vec::new(),
        }
    }
}

/// A filter applied to a single column within the auto-filter range.
///
/// Corresponds to the `<filterColumn>` element. At most one filter
/// type child is allowed per column.
#[derive(Debug, Clone, PartialEq)]
pub struct FilterColumn {
    /// 0-based column offset from the start of the auto-filter range.
    pub col_id: u32,
    /// Whether the dropdown button is hidden.
    pub hidden_button: bool,
    /// Whether the dropdown button is shown (default true).
    pub show_button: bool,
    /// The filter criterion for this column.
    pub filter: ColumnFilter,
}

impl FilterColumn {
    /// Create a filter column with a given column ID and filter.
    pub fn new(col_id: u32, filter: ColumnFilter) -> Self {
        Self {
            col_id,
            hidden_button: false,
            show_button: true,
            filter,
        }
    }
}

/// The type of filter applied to a column.
#[derive(Debug, Clone, PartialEq)]
pub enum ColumnFilter {
    /// Discrete value filter (`<filters>` element).
    /// Rows are shown only if the cell value matches one of the listed values.
    Values(ValueFilter),
    /// Custom filter with one or two operator conditions (`<customFilters>`).
    Custom(CustomFilters),
    /// Top-N or bottom-N filter (`<top10>`).
    Top10(Top10Filter),
    /// Dynamic filter based on date/value ranges (`<dynamicFilter>`).
    Dynamic(DynamicFilter),
    /// Color-based filter (`<colorFilter>`).
    Color(ColorFilter),
}

/// Discrete value filter - show rows matching any of the listed values.
#[derive(Debug, Clone, PartialEq)]
pub struct ValueFilter {
    /// The allowed values. Rows not matching any value are hidden.
    pub values: Vec<String>,
    /// Whether blank cells are included in the filter.
    pub blank: bool,
}

/// Custom filter with one or two conditions and optional AND/OR join.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomFilters {
    /// If true, both conditions must match (AND). If false, either suffices (OR).
    /// Only meaningful when two conditions are present.
    pub and: bool,
    /// Filter conditions (1 or 2).
    pub conditions: Vec<CustomFilterCondition>,
}

/// A single custom filter condition.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomFilterCondition {
    /// Comparison operator.
    pub operator: FilterOperator,
    /// The value to compare against.
    pub value: String,
}

/// Comparison operators for custom filters.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterOperator {
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

impl FilterOperator {
    /// Parse from OOXML attribute value.
    pub fn from_ooxml(s: &str) -> Option<Self> {
        match s {
            "equal" => Some(Self::Equal),
            "notEqual" => Some(Self::NotEqual),
            "greaterThan" => Some(Self::GreaterThan),
            "greaterThanOrEqual" => Some(Self::GreaterThanOrEqual),
            "lessThan" => Some(Self::LessThan),
            "lessThanOrEqual" => Some(Self::LessThanOrEqual),
            _ => None,
        }
    }

    /// Convert to OOXML attribute value.
    pub fn to_ooxml(self) -> &'static str {
        match self {
            Self::Equal => "equal",
            Self::NotEqual => "notEqual",
            Self::GreaterThan => "greaterThan",
            Self::GreaterThanOrEqual => "greaterThanOrEqual",
            Self::LessThan => "lessThan",
            Self::LessThanOrEqual => "lessThanOrEqual",
        }
    }
}

/// Top-N or bottom-N filter.
#[derive(Debug, Clone, PartialEq)]
pub struct Top10Filter {
    /// If true, filter top values; if false, filter bottom values.
    pub top: bool,
    /// If true, `val` is a percentage; if false, it's an item count.
    pub percent: bool,
    /// The count or percentage value.
    pub val: f64,
    /// The actual cutoff value computed by Excel (optional, for reader).
    pub filter_val: Option<f64>,
}

/// Dynamic filter (e.g., "today", "this month", "above average").
#[derive(Debug, Clone, PartialEq)]
pub struct DynamicFilter {
    /// The dynamic filter type.
    pub filter_type: DynamicFilterType,
    /// Optional value computed by the application.
    pub val: Option<f64>,
    /// Optional max value for range-based dynamic filters.
    pub max_val: Option<f64>,
}

/// Dynamic filter types (OOXML §18.18.26).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DynamicFilterType {
    Null,
    AboveAverage,
    BelowAverage,
    Tomorrow,
    Today,
    Yesterday,
    NextWeek,
    ThisWeek,
    LastWeek,
    NextMonth,
    ThisMonth,
    LastMonth,
    NextQuarter,
    ThisQuarter,
    LastQuarter,
    NextYear,
    ThisYear,
    LastYear,
    YearToDate,
    Q1,
    Q2,
    Q3,
    Q4,
    M1,
    M2,
    M3,
    M4,
    M5,
    M6,
    M7,
    M8,
    M9,
    M10,
    M11,
    M12,
}

impl DynamicFilterType {
    /// Parse from OOXML attribute value.
    pub fn from_ooxml(s: &str) -> Option<Self> {
        match s {
            "null" => Some(Self::Null),
            "aboveAverage" => Some(Self::AboveAverage),
            "belowAverage" => Some(Self::BelowAverage),
            "tomorrow" => Some(Self::Tomorrow),
            "today" => Some(Self::Today),
            "yesterday" => Some(Self::Yesterday),
            "nextWeek" => Some(Self::NextWeek),
            "thisWeek" => Some(Self::ThisWeek),
            "lastWeek" => Some(Self::LastWeek),
            "nextMonth" => Some(Self::NextMonth),
            "thisMonth" => Some(Self::ThisMonth),
            "lastMonth" => Some(Self::LastMonth),
            "nextQuarter" => Some(Self::NextQuarter),
            "thisQuarter" => Some(Self::ThisQuarter),
            "lastQuarter" => Some(Self::LastQuarter),
            "nextYear" => Some(Self::NextYear),
            "thisYear" => Some(Self::ThisYear),
            "lastYear" => Some(Self::LastYear),
            "yearToDate" => Some(Self::YearToDate),
            "Q1" => Some(Self::Q1),
            "Q2" => Some(Self::Q2),
            "Q3" => Some(Self::Q3),
            "Q4" => Some(Self::Q4),
            "M1" => Some(Self::M1),
            "M2" => Some(Self::M2),
            "M3" => Some(Self::M3),
            "M4" => Some(Self::M4),
            "M5" => Some(Self::M5),
            "M6" => Some(Self::M6),
            "M7" => Some(Self::M7),
            "M8" => Some(Self::M8),
            "M9" => Some(Self::M9),
            "M10" => Some(Self::M10),
            "M11" => Some(Self::M11),
            "M12" => Some(Self::M12),
            _ => None,
        }
    }

    /// Convert to OOXML attribute value.
    pub fn to_ooxml(self) -> &'static str {
        match self {
            Self::Null => "null",
            Self::AboveAverage => "aboveAverage",
            Self::BelowAverage => "belowAverage",
            Self::Tomorrow => "tomorrow",
            Self::Today => "today",
            Self::Yesterday => "yesterday",
            Self::NextWeek => "nextWeek",
            Self::ThisWeek => "thisWeek",
            Self::LastWeek => "lastWeek",
            Self::NextMonth => "nextMonth",
            Self::ThisMonth => "thisMonth",
            Self::LastMonth => "lastMonth",
            Self::NextQuarter => "nextQuarter",
            Self::ThisQuarter => "thisQuarter",
            Self::LastQuarter => "lastQuarter",
            Self::NextYear => "nextYear",
            Self::ThisYear => "thisYear",
            Self::LastYear => "lastYear",
            Self::YearToDate => "yearToDate",
            Self::Q1 => "Q1",
            Self::Q2 => "Q2",
            Self::Q3 => "Q3",
            Self::Q4 => "Q4",
            Self::M1 => "M1",
            Self::M2 => "M2",
            Self::M3 => "M3",
            Self::M4 => "M4",
            Self::M5 => "M5",
            Self::M6 => "M6",
            Self::M7 => "M7",
            Self::M8 => "M8",
            Self::M9 => "M9",
            Self::M10 => "M10",
            Self::M11 => "M11",
            Self::M12 => "M12",
        }
    }
}

/// Color-based filter.
#[derive(Debug, Clone, PartialEq)]
pub struct ColorFilter {
    /// DXF ID for the color to filter on.
    pub dxf_id: Option<u32>,
    /// If true, filter by cell color; if false, filter by font color.
    pub cell_color: bool,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_auto_filter_new() {
        let range = CellRange::parse("A1:D10").unwrap();
        let af = AutoFilter::new(range.clone());
        assert_eq!(af.range, range);
        assert!(af.filter_columns.is_empty());
    }

    #[test]
    fn test_filter_operator_roundtrip() {
        let ops = [
            FilterOperator::Equal,
            FilterOperator::NotEqual,
            FilterOperator::GreaterThan,
            FilterOperator::GreaterThanOrEqual,
            FilterOperator::LessThan,
            FilterOperator::LessThanOrEqual,
        ];
        for op in &ops {
            let s = op.to_ooxml();
            let parsed = FilterOperator::from_ooxml(s).unwrap();
            assert_eq!(*op, parsed);
        }
    }

    #[test]
    fn test_dynamic_filter_type_roundtrip() {
        let types = [
            DynamicFilterType::AboveAverage,
            DynamicFilterType::Today,
            DynamicFilterType::ThisMonth,
            DynamicFilterType::Q1,
            DynamicFilterType::M12,
        ];
        for t in &types {
            let s = t.to_ooxml();
            let parsed = DynamicFilterType::from_ooxml(s).unwrap();
            assert_eq!(*t, parsed);
        }
    }

    #[test]
    fn test_filter_column_new() {
        let fc = FilterColumn::new(
            2,
            ColumnFilter::Values(ValueFilter {
                values: vec!["Alice".into(), "Bob".into()],
                blank: false,
            }),
        );
        assert_eq!(fc.col_id, 2);
        assert!(!fc.hidden_button);
        assert!(fc.show_button);
    }
}
