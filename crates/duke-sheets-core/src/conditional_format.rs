//! Conditional formatting
//!
//! This module provides support for conditional formatting rules in worksheets.
//! Conditional formatting allows you to apply formatting to cells based on their values.
//!
//! ## Example
//!
//! ```rust
//! use duke_sheets_core::{Workbook, ConditionalFormatRule, CellRange};
//! use duke_sheets_core::style::{Color, Style};
//!
//! let mut workbook = Workbook::new();
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! // Highlight cells greater than 100
//! let rule = ConditionalFormatRule::cell_is_greater_than("100")
//!     .with_range(CellRange::parse("A1:A10").unwrap())
//!     .with_format(Style::new().fill_color(Color::rgb(255, 199, 206))); // Light red
//!
//! sheet.add_conditional_format(rule);
//! ```

use crate::cell::CellRange;
use crate::style::{Color, Style};

/// A conditional formatting rule
#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalFormatRule {
    /// Rule type
    pub rule_type: CfRuleType,
    /// Cell ranges this rule applies to
    pub ranges: Vec<CellRange>,
    /// Priority (lower = higher priority)
    pub priority: u32,
    /// Stop processing further rules if this one matches
    pub stop_if_true: bool,
    /// Format to apply when rule matches (for simple rules)
    pub format: Option<Style>,
    /// Differential format index (for XLSX - index into dxf table)
    pub dxf_id: Option<u32>,
}

impl Default for ConditionalFormatRule {
    fn default() -> Self {
        Self {
            rule_type: CfRuleType::Expression {
                formula: String::new(),
            },
            ranges: Vec::new(),
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        }
    }
}

impl ConditionalFormatRule {
    /// Create a new conditional format rule
    pub fn new(rule_type: CfRuleType) -> Self {
        Self {
            rule_type,
            ..Self::default()
        }
    }

    // === Cell Is rules ===

    /// Highlight cells greater than a value
    pub fn cell_is_greater_than(value: impl Into<String>) -> Self {
        Self::new(CfRuleType::CellIs {
            operator: CfOperator::GreaterThan,
            formula1: value.into(),
            formula2: None,
        })
    }

    /// Highlight cells less than a value
    pub fn cell_is_less_than(value: impl Into<String>) -> Self {
        Self::new(CfRuleType::CellIs {
            operator: CfOperator::LessThan,
            formula1: value.into(),
            formula2: None,
        })
    }

    /// Highlight cells equal to a value
    pub fn cell_is_equal_to(value: impl Into<String>) -> Self {
        Self::new(CfRuleType::CellIs {
            operator: CfOperator::Equal,
            formula1: value.into(),
            formula2: None,
        })
    }

    /// Highlight cells between two values
    pub fn cell_is_between(value1: impl Into<String>, value2: impl Into<String>) -> Self {
        Self::new(CfRuleType::CellIs {
            operator: CfOperator::Between,
            formula1: value1.into(),
            formula2: Some(value2.into()),
        })
    }

    // === Expression rule ===

    /// Highlight cells where formula evaluates to TRUE
    pub fn expression(formula: impl Into<String>) -> Self {
        Self::new(CfRuleType::Expression {
            formula: formula.into(),
        })
    }

    // === Color Scale rules ===

    /// Create a 2-color scale (min to max)
    pub fn color_scale_2(min_color: Color, max_color: Color) -> Self {
        Self::new(CfRuleType::ColorScale {
            colors: vec![
                CfColorValue::new(CfValueType::Min, None, min_color),
                CfColorValue::new(CfValueType::Max, None, max_color),
            ],
        })
    }

    /// Create a 3-color scale (min, mid, max)
    pub fn color_scale_3(min_color: Color, mid_color: Color, max_color: Color) -> Self {
        Self::new(CfRuleType::ColorScale {
            colors: vec![
                CfColorValue::new(CfValueType::Min, None, min_color),
                CfColorValue::new(CfValueType::Percentile, Some("50".to_string()), mid_color),
                CfColorValue::new(CfValueType::Max, None, max_color),
            ],
        })
    }

    // === Data Bar rule ===

    /// Create a data bar
    pub fn data_bar(color: Color) -> Self {
        Self::new(CfRuleType::DataBar {
            min_value: CfValue::new(CfValueType::Min, None),
            max_value: CfValue::new(CfValueType::Max, None),
            color,
            show_value: true,
            gradient: true,
            border_color: None,
            negative_color: None,
        })
    }

    // === Icon Set rule ===

    /// Create an icon set
    pub fn icon_set(style: IconSetStyle) -> Self {
        // Default thresholds for icon sets
        let values = match style.icon_count() {
            3 => vec![
                CfValue::new(CfValueType::Percent, Some("0".to_string())),
                CfValue::new(CfValueType::Percent, Some("33".to_string())),
                CfValue::new(CfValueType::Percent, Some("67".to_string())),
            ],
            4 => vec![
                CfValue::new(CfValueType::Percent, Some("0".to_string())),
                CfValue::new(CfValueType::Percent, Some("25".to_string())),
                CfValue::new(CfValueType::Percent, Some("50".to_string())),
                CfValue::new(CfValueType::Percent, Some("75".to_string())),
            ],
            5 => vec![
                CfValue::new(CfValueType::Percent, Some("0".to_string())),
                CfValue::new(CfValueType::Percent, Some("20".to_string())),
                CfValue::new(CfValueType::Percent, Some("40".to_string())),
                CfValue::new(CfValueType::Percent, Some("60".to_string())),
                CfValue::new(CfValueType::Percent, Some("80".to_string())),
            ],
            _ => vec![],
        };

        Self::new(CfRuleType::IconSet {
            icon_style: style,
            values,
            reverse: false,
            show_value: true,
        })
    }

    // === Top/Bottom rules ===

    /// Highlight top N values
    pub fn top_n(n: u32) -> Self {
        Self::new(CfRuleType::Top10 {
            rank: n,
            percent: false,
            bottom: false,
        })
    }

    /// Highlight bottom N values
    pub fn bottom_n(n: u32) -> Self {
        Self::new(CfRuleType::Top10 {
            rank: n,
            percent: false,
            bottom: true,
        })
    }

    /// Highlight top N percent
    pub fn top_percent(n: u32) -> Self {
        Self::new(CfRuleType::Top10 {
            rank: n,
            percent: true,
            bottom: false,
        })
    }

    // === Above/Below Average rules ===

    /// Highlight cells above average
    pub fn above_average() -> Self {
        Self::new(CfRuleType::AboveAverage {
            above: true,
            equal_average: false,
            std_dev: None,
        })
    }

    /// Highlight cells below average
    pub fn below_average() -> Self {
        Self::new(CfRuleType::AboveAverage {
            above: false,
            equal_average: false,
            std_dev: None,
        })
    }

    // === Text rules ===

    /// Highlight cells containing text
    pub fn contains_text(text: impl Into<String>) -> Self {
        Self::new(CfRuleType::ContainsText { text: text.into() })
    }

    /// Highlight cells beginning with text
    pub fn begins_with(text: impl Into<String>) -> Self {
        Self::new(CfRuleType::BeginsWith { text: text.into() })
    }

    /// Highlight cells ending with text
    pub fn ends_with(text: impl Into<String>) -> Self {
        Self::new(CfRuleType::EndsWith { text: text.into() })
    }

    // === Duplicate/Unique rules ===

    /// Highlight duplicate values
    pub fn duplicate_values() -> Self {
        Self::new(CfRuleType::DuplicateValues)
    }

    /// Highlight unique values
    pub fn unique_values() -> Self {
        Self::new(CfRuleType::UniqueValues)
    }

    // === Blanks/Errors rules ===

    /// Highlight blank cells
    pub fn contains_blanks() -> Self {
        Self::new(CfRuleType::ContainsBlanks)
    }

    /// Highlight cells containing errors
    pub fn contains_errors() -> Self {
        Self::new(CfRuleType::ContainsErrors)
    }

    // === Builder methods ===

    /// Add a cell range to this rule
    pub fn with_range(mut self, range: CellRange) -> Self {
        self.ranges.push(range);
        self
    }

    /// Set the cell ranges for this rule
    pub fn with_ranges(mut self, ranges: Vec<CellRange>) -> Self {
        self.ranges = ranges;
        self
    }

    /// Set the format to apply when rule matches
    pub fn with_format(mut self, style: Style) -> Self {
        self.format = Some(style);
        self
    }

    /// Set the priority (lower = higher priority)
    pub fn with_priority(mut self, priority: u32) -> Self {
        self.priority = priority;
        self
    }

    /// Set whether to stop processing further rules if this one matches
    pub fn with_stop_if_true(mut self, stop: bool) -> Self {
        self.stop_if_true = stop;
        self
    }

    /// Check if this rule applies to a specific cell
    pub fn applies_to(&self, row: u32, col: u16) -> bool {
        self.ranges.iter().any(|r| {
            row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col
        })
    }
}

/// Types of conditional formatting rules
#[derive(Debug, Clone, PartialEq)]
pub enum CfRuleType {
    /// Cell value comparison (e.g., "greater than 100")
    CellIs {
        operator: CfOperator,
        formula1: String,
        formula2: Option<String>,
    },

    /// Formula evaluates to TRUE
    Expression { formula: String },

    /// Color scale (2 or 3 color gradient)
    ColorScale { colors: Vec<CfColorValue> },

    /// Data bar (in-cell bar chart)
    DataBar {
        min_value: CfValue,
        max_value: CfValue,
        color: Color,
        show_value: bool,
        gradient: bool,
        border_color: Option<Color>,
        negative_color: Option<Color>,
    },

    /// Icon set (arrows, traffic lights, etc.)
    IconSet {
        icon_style: IconSetStyle,
        values: Vec<CfValue>,
        reverse: bool,
        show_value: bool,
    },

    /// Top/bottom N values
    Top10 {
        rank: u32,
        percent: bool,
        bottom: bool,
    },

    /// Above/below average
    AboveAverage {
        above: bool,
        equal_average: bool,
        std_dev: Option<u32>,
    },

    /// Contains text
    ContainsText { text: String },

    /// Begins with text
    BeginsWith { text: String },

    /// Ends with text
    EndsWith { text: String },

    /// Duplicate values
    DuplicateValues,

    /// Unique values
    UniqueValues,

    /// Blank cells
    ContainsBlanks,

    /// Non-blank cells
    NotContainsBlanks,

    /// Cells containing errors
    ContainsErrors,

    /// Cells not containing errors
    NotContainsErrors,

    /// Time period (today, yesterday, this week, etc.)
    TimePeriod { period: TimePeriod },
}

impl CfRuleType {
    /// Get the XLSX type string for this rule type
    pub fn xlsx_type(&self) -> &'static str {
        match self {
            CfRuleType::CellIs { .. } => "cellIs",
            CfRuleType::Expression { .. } => "expression",
            CfRuleType::ColorScale { .. } => "colorScale",
            CfRuleType::DataBar { .. } => "dataBar",
            CfRuleType::IconSet { .. } => "iconSet",
            CfRuleType::Top10 { .. } => "top10",
            CfRuleType::AboveAverage { .. } => "aboveAverage",
            CfRuleType::ContainsText { .. } => "containsText",
            CfRuleType::BeginsWith { .. } => "beginsWith",
            CfRuleType::EndsWith { .. } => "endsWith",
            CfRuleType::DuplicateValues => "duplicateValues",
            CfRuleType::UniqueValues => "uniqueValues",
            CfRuleType::ContainsBlanks => "containsBlanks",
            CfRuleType::NotContainsBlanks => "notContainsBlanks",
            CfRuleType::ContainsErrors => "containsErrors",
            CfRuleType::NotContainsErrors => "notContainsErrors",
            CfRuleType::TimePeriod { .. } => "timePeriod",
        }
    }
}

/// Operators for CellIs rules
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CfOperator {
    /// Value is between formula1 and formula2
    #[default]
    Between,
    /// Value is NOT between formula1 and formula2
    NotBetween,
    /// Value equals formula1
    Equal,
    /// Value does NOT equal formula1
    NotEqual,
    /// Value is greater than formula1
    GreaterThan,
    /// Value is less than formula1
    LessThan,
    /// Value is greater than or equal to formula1
    GreaterThanOrEqual,
    /// Value is less than or equal to formula1
    LessThanOrEqual,
}

impl CfOperator {
    /// Get the XLSX operator string
    pub fn xlsx_operator(&self) -> &'static str {
        match self {
            CfOperator::Between => "between",
            CfOperator::NotBetween => "notBetween",
            CfOperator::Equal => "equal",
            CfOperator::NotEqual => "notEqual",
            CfOperator::GreaterThan => "greaterThan",
            CfOperator::LessThan => "lessThan",
            CfOperator::GreaterThanOrEqual => "greaterThanOrEqual",
            CfOperator::LessThanOrEqual => "lessThanOrEqual",
        }
    }

    /// Parse from XLSX operator string
    pub fn from_xlsx(s: &str) -> Option<Self> {
        match s {
            "between" => Some(CfOperator::Between),
            "notBetween" => Some(CfOperator::NotBetween),
            "equal" => Some(CfOperator::Equal),
            "notEqual" => Some(CfOperator::NotEqual),
            "greaterThan" => Some(CfOperator::GreaterThan),
            "lessThan" => Some(CfOperator::LessThan),
            "greaterThanOrEqual" => Some(CfOperator::GreaterThanOrEqual),
            "lessThanOrEqual" => Some(CfOperator::LessThanOrEqual),
            _ => None,
        }
    }
}

/// Value specification for color scales, data bars, icon sets
#[derive(Debug, Clone, PartialEq)]
pub struct CfValue {
    /// How to interpret the value
    pub value_type: CfValueType,
    /// The value (if applicable)
    pub value: Option<String>,
}

impl CfValue {
    /// Create a new conditional format value
    pub fn new(value_type: CfValueType, value: Option<String>) -> Self {
        Self { value_type, value }
    }

    /// Create a min value
    pub fn min() -> Self {
        Self::new(CfValueType::Min, None)
    }

    /// Create a max value
    pub fn max() -> Self {
        Self::new(CfValueType::Max, None)
    }

    /// Create a number value
    pub fn number(n: impl Into<String>) -> Self {
        Self::new(CfValueType::Num, Some(n.into()))
    }

    /// Create a percent value
    pub fn percent(p: impl Into<String>) -> Self {
        Self::new(CfValueType::Percent, Some(p.into()))
    }

    /// Create a percentile value
    pub fn percentile(p: impl Into<String>) -> Self {
        Self::new(CfValueType::Percentile, Some(p.into()))
    }

    /// Create a formula value
    pub fn formula(f: impl Into<String>) -> Self {
        Self::new(CfValueType::Formula, Some(f.into()))
    }
}

/// Color with value threshold for color scales
#[derive(Debug, Clone, PartialEq)]
pub struct CfColorValue {
    /// How to interpret the value
    pub value_type: CfValueType,
    /// The value (if applicable)
    pub value: Option<String>,
    /// Color at this threshold
    pub color: Color,
}

impl CfColorValue {
    /// Create a new color value
    pub fn new(value_type: CfValueType, value: Option<String>, color: Color) -> Self {
        Self {
            value_type,
            value,
            color,
        }
    }
}

/// Value types for conditional format thresholds
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CfValueType {
    /// Minimum value in range
    #[default]
    Min,
    /// Maximum value in range
    Max,
    /// Specific number
    Num,
    /// Percentage (0-100)
    Percent,
    /// Percentile (0-100)
    Percentile,
    /// Formula result
    Formula,
    /// Automatic minimum
    AutoMin,
    /// Automatic maximum
    AutoMax,
}

impl CfValueType {
    /// Get the XLSX type string
    pub fn xlsx_type(&self) -> &'static str {
        match self {
            CfValueType::Min => "min",
            CfValueType::Max => "max",
            CfValueType::Num => "num",
            CfValueType::Percent => "percent",
            CfValueType::Percentile => "percentile",
            CfValueType::Formula => "formula",
            CfValueType::AutoMin => "autoMin",
            CfValueType::AutoMax => "autoMax",
        }
    }

    /// Parse from XLSX type string
    pub fn from_xlsx(s: &str) -> Option<Self> {
        match s {
            "min" => Some(CfValueType::Min),
            "max" => Some(CfValueType::Max),
            "num" => Some(CfValueType::Num),
            "percent" => Some(CfValueType::Percent),
            "percentile" => Some(CfValueType::Percentile),
            "formula" => Some(CfValueType::Formula),
            "autoMin" => Some(CfValueType::AutoMin),
            "autoMax" => Some(CfValueType::AutoMax),
            _ => None,
        }
    }
}

/// Icon set styles
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum IconSetStyle {
    /// 3 arrows (up, right, down)
    #[default]
    Arrows3,
    /// 3 gray arrows
    Arrows3Gray,
    /// 3 flags
    Flags3,
    /// 3 traffic lights (unrimmed)
    TrafficLights3,
    /// 3 traffic lights (rimmed)
    TrafficLights3Black,
    /// 3 signs
    Signs3,
    /// 3 symbols (checkmark, exclamation, X)
    Symbols3,
    /// 3 symbols circled
    Symbols3Circled,
    /// 3 stars
    Stars3,
    /// 3 triangles
    Triangles3,
    /// 4 arrows
    Arrows4,
    /// 4 gray arrows
    Arrows4Gray,
    /// 4 circles (red to black)
    RedToBlack4,
    /// 4 ratings
    Rating4,
    /// 4 traffic lights
    TrafficLights4,
    /// 5 arrows
    Arrows5,
    /// 5 gray arrows
    Arrows5Gray,
    /// 5 ratings
    Rating5,
    /// 5 quarters
    Quarters5,
    /// 5 boxes
    Boxes5,
}

impl IconSetStyle {
    /// Get the XLSX icon set name
    pub fn xlsx_name(&self) -> &'static str {
        match self {
            IconSetStyle::Arrows3 => "3Arrows",
            IconSetStyle::Arrows3Gray => "3ArrowsGray",
            IconSetStyle::Flags3 => "3Flags",
            IconSetStyle::TrafficLights3 => "3TrafficLights1",
            IconSetStyle::TrafficLights3Black => "3TrafficLights2",
            IconSetStyle::Signs3 => "3Signs",
            IconSetStyle::Symbols3 => "3Symbols",
            IconSetStyle::Symbols3Circled => "3Symbols2",
            IconSetStyle::Stars3 => "3Stars",
            IconSetStyle::Triangles3 => "3Triangles",
            IconSetStyle::Arrows4 => "4Arrows",
            IconSetStyle::Arrows4Gray => "4ArrowsGray",
            IconSetStyle::RedToBlack4 => "4RedToBlack",
            IconSetStyle::Rating4 => "4Rating",
            IconSetStyle::TrafficLights4 => "4TrafficLights",
            IconSetStyle::Arrows5 => "5Arrows",
            IconSetStyle::Arrows5Gray => "5ArrowsGray",
            IconSetStyle::Rating5 => "5Rating",
            IconSetStyle::Quarters5 => "5Quarters",
            IconSetStyle::Boxes5 => "5Boxes",
        }
    }

    /// Parse from XLSX icon set name
    pub fn from_xlsx(s: &str) -> Option<Self> {
        match s {
            "3Arrows" => Some(IconSetStyle::Arrows3),
            "3ArrowsGray" => Some(IconSetStyle::Arrows3Gray),
            "3Flags" => Some(IconSetStyle::Flags3),
            "3TrafficLights1" => Some(IconSetStyle::TrafficLights3),
            "3TrafficLights2" => Some(IconSetStyle::TrafficLights3Black),
            "3Signs" => Some(IconSetStyle::Signs3),
            "3Symbols" => Some(IconSetStyle::Symbols3),
            "3Symbols2" => Some(IconSetStyle::Symbols3Circled),
            "3Stars" => Some(IconSetStyle::Stars3),
            "3Triangles" => Some(IconSetStyle::Triangles3),
            "4Arrows" => Some(IconSetStyle::Arrows4),
            "4ArrowsGray" => Some(IconSetStyle::Arrows4Gray),
            "4RedToBlack" => Some(IconSetStyle::RedToBlack4),
            "4Rating" => Some(IconSetStyle::Rating4),
            "4TrafficLights" => Some(IconSetStyle::TrafficLights4),
            "5Arrows" => Some(IconSetStyle::Arrows5),
            "5ArrowsGray" => Some(IconSetStyle::Arrows5Gray),
            "5Rating" => Some(IconSetStyle::Rating5),
            "5Quarters" => Some(IconSetStyle::Quarters5),
            "5Boxes" => Some(IconSetStyle::Boxes5),
            _ => None,
        }
    }

    /// Get the number of icons in this set
    pub fn icon_count(&self) -> usize {
        match self {
            IconSetStyle::Arrows3
            | IconSetStyle::Arrows3Gray
            | IconSetStyle::Flags3
            | IconSetStyle::TrafficLights3
            | IconSetStyle::TrafficLights3Black
            | IconSetStyle::Signs3
            | IconSetStyle::Symbols3
            | IconSetStyle::Symbols3Circled
            | IconSetStyle::Stars3
            | IconSetStyle::Triangles3 => 3,

            IconSetStyle::Arrows4
            | IconSetStyle::Arrows4Gray
            | IconSetStyle::RedToBlack4
            | IconSetStyle::Rating4
            | IconSetStyle::TrafficLights4 => 4,

            IconSetStyle::Arrows5
            | IconSetStyle::Arrows5Gray
            | IconSetStyle::Rating5
            | IconSetStyle::Quarters5
            | IconSetStyle::Boxes5 => 5,
        }
    }
}

/// Time periods for time-based conditional formatting
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TimePeriod {
    /// Today
    #[default]
    Today,
    /// Yesterday
    Yesterday,
    /// Tomorrow
    Tomorrow,
    /// Last 7 days
    Last7Days,
    /// This week
    ThisWeek,
    /// Last week
    LastWeek,
    /// Next week
    NextWeek,
    /// This month
    ThisMonth,
    /// Last month
    LastMonth,
    /// Next month
    NextMonth,
}

impl TimePeriod {
    /// Get the XLSX time period string
    pub fn xlsx_period(&self) -> &'static str {
        match self {
            TimePeriod::Today => "today",
            TimePeriod::Yesterday => "yesterday",
            TimePeriod::Tomorrow => "tomorrow",
            TimePeriod::Last7Days => "last7Days",
            TimePeriod::ThisWeek => "thisWeek",
            TimePeriod::LastWeek => "lastWeek",
            TimePeriod::NextWeek => "nextWeek",
            TimePeriod::ThisMonth => "thisMonth",
            TimePeriod::LastMonth => "lastMonth",
            TimePeriod::NextMonth => "nextMonth",
        }
    }

    /// Parse from XLSX time period string
    pub fn from_xlsx(s: &str) -> Option<Self> {
        match s {
            "today" => Some(TimePeriod::Today),
            "yesterday" => Some(TimePeriod::Yesterday),
            "tomorrow" => Some(TimePeriod::Tomorrow),
            "last7Days" => Some(TimePeriod::Last7Days),
            "thisWeek" => Some(TimePeriod::ThisWeek),
            "lastWeek" => Some(TimePeriod::LastWeek),
            "nextWeek" => Some(TimePeriod::NextWeek),
            "thisMonth" => Some(TimePeriod::ThisMonth),
            "lastMonth" => Some(TimePeriod::LastMonth),
            "nextMonth" => Some(TimePeriod::NextMonth),
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_is_rule() {
        let rule = ConditionalFormatRule::cell_is_greater_than("100")
            .with_range(CellRange::parse("A1:A10").unwrap());

        assert!(matches!(
            rule.rule_type,
            CfRuleType::CellIs {
                operator: CfOperator::GreaterThan,
                ..
            }
        ));
        assert_eq!(rule.ranges.len(), 1);
    }

    #[test]
    fn test_color_scale() {
        let rule = ConditionalFormatRule::color_scale_3(
            Color::rgb(255, 0, 0),   // Red
            Color::rgb(255, 255, 0), // Yellow
            Color::rgb(0, 255, 0),   // Green
        );

        if let CfRuleType::ColorScale { colors } = &rule.rule_type {
            assert_eq!(colors.len(), 3);
        } else {
            panic!("Expected ColorScale rule type");
        }
    }

    #[test]
    fn test_data_bar() {
        let rule = ConditionalFormatRule::data_bar(Color::rgb(99, 142, 198));

        if let CfRuleType::DataBar { color, .. } = &rule.rule_type {
            assert_eq!(*color, Color::rgb(99, 142, 198));
        } else {
            panic!("Expected DataBar rule type");
        }
    }

    #[test]
    fn test_icon_set() {
        let rule = ConditionalFormatRule::icon_set(IconSetStyle::TrafficLights3);

        if let CfRuleType::IconSet {
            icon_style, values, ..
        } = &rule.rule_type
        {
            assert_eq!(*icon_style, IconSetStyle::TrafficLights3);
            assert_eq!(values.len(), 3);
        } else {
            panic!("Expected IconSet rule type");
        }
    }

    #[test]
    fn test_top_n() {
        let rule = ConditionalFormatRule::top_n(10);

        if let CfRuleType::Top10 {
            rank,
            percent,
            bottom,
        } = &rule.rule_type
        {
            assert_eq!(*rank, 10);
            assert!(!percent);
            assert!(!bottom);
        } else {
            panic!("Expected Top10 rule type");
        }
    }

    #[test]
    fn test_applies_to() {
        let rule = ConditionalFormatRule::cell_is_greater_than("0")
            .with_range(CellRange::parse("A1:C10").unwrap());

        assert!(rule.applies_to(0, 0)); // A1
        assert!(rule.applies_to(5, 2)); // C6
        assert!(!rule.applies_to(10, 0)); // A11 - out of range
        assert!(!rule.applies_to(0, 3)); // D1 - out of range
    }

    #[test]
    fn test_with_format() {
        let rule = ConditionalFormatRule::cell_is_greater_than("0")
            .with_format(Style::new().fill_color(Color::rgb(255, 199, 206)));

        assert!(rule.format.is_some());
    }

    #[test]
    fn test_xlsx_type_strings() {
        assert_eq!(CfOperator::GreaterThan.xlsx_operator(), "greaterThan");
        assert_eq!(
            CfOperator::from_xlsx("lessThanOrEqual"),
            Some(CfOperator::LessThanOrEqual)
        );

        assert_eq!(IconSetStyle::TrafficLights3.xlsx_name(), "3TrafficLights1");
        assert_eq!(
            IconSetStyle::from_xlsx("4Rating"),
            Some(IconSetStyle::Rating4)
        );
    }
}
