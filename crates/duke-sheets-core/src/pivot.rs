//! Pivot table model.
//!
//! The public pivot API is semantic: callers describe the source data, axes,
//! measures, filters, layout, and refresh policy they want. File-format cache
//! records are intentionally not part of the authoring API; readers and writers
//! can preserve or synthesize those artifacts internally.

use std::fmt;
use std::hash::{Hash, Hasher};

use crate::cell::{CellAddress, CellError, CellRange, CellValue};
use crate::error::{Error, Result};
use crate::rich_text::rich_text_to_plain;

/// A scalar value used for grouping and filtering pivot data.
///
/// Numeric values compare and hash by their exact IEEE-754 bit pattern so the
/// value can be used safely as a hash-map key. This keeps the model lossless for
/// file round-tripping; engines may normalize values internally when desired.
#[derive(Debug, Clone)]
pub enum PivotValue {
    /// Blank or empty source value.
    Blank,
    /// Boolean value.
    Boolean(bool),
    /// Numeric value.
    Number(f64),
    /// Text value.
    String(String),
    /// Spreadsheet error value.
    Error(CellError),
}

impl PivotValue {
    /// Convert a cell value into a pivot grouping/filtering value.
    pub fn from_cell_value(value: &CellValue) -> Self {
        match value {
            CellValue::Empty => Self::Blank,
            CellValue::Boolean(value) => Self::Boolean(*value),
            CellValue::Number(value) => Self::Number(*value),
            CellValue::String(value) => Self::String(value.as_str().to_string()),
            CellValue::Error(value) => Self::Error(*value),
            CellValue::RichText(runs) => Self::String(rich_text_to_plain(runs)),
            CellValue::SpillTarget { .. } => Self::Blank,
        }
    }

    /// Convert this pivot value into a worksheet cell value.
    pub fn to_cell_value(&self) -> CellValue {
        match self {
            Self::Blank => CellValue::Empty,
            Self::Boolean(value) => CellValue::Boolean(*value),
            Self::Number(value) => CellValue::Number(*value),
            Self::String(value) => CellValue::string(value),
            Self::Error(value) => CellValue::Error(*value),
        }
    }

    /// Whether the pivot value is blank.
    pub fn is_blank(&self) -> bool {
        matches!(self, Self::Blank)
    }
}

impl PartialEq for PivotValue {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Blank, Self::Blank) => true,
            (Self::Boolean(a), Self::Boolean(b)) => a == b,
            (Self::Number(a), Self::Number(b)) => a.to_bits() == b.to_bits(),
            (Self::String(a), Self::String(b)) => a == b,
            (Self::Error(a), Self::Error(b)) => a == b,
            _ => false,
        }
    }
}

impl Eq for PivotValue {}

impl Hash for PivotValue {
    fn hash<H: Hasher>(&self, state: &mut H) {
        std::mem::discriminant(self).hash(state);
        match self {
            Self::Blank => {}
            Self::Boolean(value) => value.hash(state),
            Self::Number(value) => value.to_bits().hash(state),
            Self::String(value) => value.hash(state),
            Self::Error(value) => value.hash(state),
        }
    }
}

impl fmt::Display for PivotValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Blank => Ok(()),
            Self::Boolean(value) => write!(f, "{}", if *value { "TRUE" } else { "FALSE" }),
            Self::Number(value) => write!(f, "{value}"),
            Self::String(value) => write!(f, "{value}"),
            Self::Error(value) => write!(f, "{value}"),
        }
    }
}

impl From<&CellValue> for PivotValue {
    fn from(value: &CellValue) -> Self {
        Self::from_cell_value(value)
    }
}

impl From<&str> for PivotValue {
    fn from(value: &str) -> Self {
        Self::String(value.to_string())
    }
}

impl From<String> for PivotValue {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}

impl From<f64> for PivotValue {
    fn from(value: f64) -> Self {
        Self::Number(value)
    }
}

impl From<bool> for PivotValue {
    fn from(value: bool) -> Self {
        Self::Boolean(value)
    }
}

/// The source data for a pivot table.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotSource {
    /// A rectangular worksheet range. If `sheet` is `None`, the sheet containing
    /// the pivot table is used.
    WorksheetRange {
        /// Optional source sheet name.
        sheet: Option<String>,
        /// Source range including its header row.
        range: CellRange,
    },
    /// A workbook table/list object by name.
    Table {
        /// Table name.
        name: String,
    },
    /// External connection-backed source preserved by readers/writers.
    External {
        /// Connection name or identifier.
        connection_name: String,
        /// Optional command/query text.
        command_text: Option<String>,
    },
    /// Multiple worksheet ranges consolidated into one pivot source.
    Consolidation {
        /// Source ranges.
        ranges: Vec<PivotSourceRange>,
    },
    /// Scenario-manager source preserved by readers/writers.
    Scenario {
        /// Scenario name.
        name: String,
    },
    /// OLAP/data-model source preserved by readers/writers.
    Olap {
        /// Connection name or identifier.
        connection_name: String,
        /// Cube name.
        cube: Option<String>,
        /// Optional command/query text.
        command_text: Option<String>,
    },
}

impl PivotSource {
    /// Create a worksheet-range source on the pivot table's sheet.
    pub fn range(range: CellRange) -> Self {
        Self::WorksheetRange { sheet: None, range }
    }

    /// Create a worksheet-range source on a named sheet.
    pub fn range_on_sheet(sheet: impl Into<String>, range: CellRange) -> Self {
        Self::WorksheetRange {
            sheet: Some(sheet.into()),
            range,
        }
    }

    /// Create a table source by table name.
    pub fn table(name: impl Into<String>) -> Self {
        Self::Table { name: name.into() }
    }
}

/// A worksheet range used by a consolidated pivot source.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotSourceRange {
    /// Source sheet name.
    pub sheet: String,
    /// Source range.
    pub range: CellRange,
    /// Optional display name for the source range.
    pub name: Option<String>,
    /// Consolidation page-item labels for this range, ordered by page field.
    pub page_items: Vec<String>,
}

impl PivotSourceRange {
    /// Create a source range.
    pub fn new(sheet: impl Into<String>, range: CellRange) -> Self {
        Self {
            sheet: sheet.into(),
            range,
            name: None,
            page_items: Vec::new(),
        }
    }

    /// Set the display name for this source range.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set consolidation page-item labels for this source range.
    pub fn with_page_items<I, S>(mut self, page_items: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.page_items = page_items.into_iter().map(Into::into).collect();
        self
    }
}

/// A reference to a pivot source field by display/header name.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotFieldRef {
    /// Source field name.
    pub name: String,
}

impl PivotFieldRef {
    /// Create a field reference.
    pub fn new(name: impl Into<String>) -> Self {
        Self { name: name.into() }
    }
}

impl From<&str> for PivotFieldRef {
    fn from(value: &str) -> Self {
        Self::new(value)
    }
}

impl From<String> for PivotFieldRef {
    fn from(value: String) -> Self {
        Self::new(value)
    }
}

/// A field placed on a pivot axis.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotField {
    /// Source field reference.
    pub field: PivotFieldRef,
    /// Sort behavior for this field.
    pub sort: PivotSort,
    /// Subtotal behavior for this field.
    pub subtotal: PivotSubtotal,
    /// Explicit subtotal functions for formats that support multiple subtotals.
    ///
    /// When empty, [`PivotField::subtotal`] is used for compatibility with the
    /// simple single-subtotal API.
    pub subtotals: Vec<PivotSubtotal>,
    /// Items whose details are collapsed for this axis field.
    ///
    /// SpreadsheetML stores this as `CT_Item@sd="0"` on item records; the
    /// semantic model keeps the item values here so callers do not need to
    /// manage file-format cache indexes.
    pub collapsed_items: Vec<PivotValue>,
    /// Whether items with no data should be shown.
    pub show_empty_items: bool,
    /// Show the field dropdown/filter control where supported.
    pub show_drop_downs: bool,
    /// Place subtotals above grouped items.
    pub subtotal_top: bool,
    /// Insert a blank row after each item group.
    pub insert_blank_row: bool,
    /// Insert a page break after each item group.
    pub insert_page_break: bool,
    /// Include new source items in existing filters.
    pub include_new_items_in_filter: bool,
    /// Number of items shown per page in item filter menus.
    pub item_page_count: u32,
}

impl PivotField {
    /// Create an axis field with default sorting and subtotals.
    pub fn new(field: impl Into<PivotFieldRef>) -> Self {
        Self {
            field: field.into(),
            sort: PivotSort::Ascending,
            subtotal: PivotSubtotal::Automatic,
            subtotals: Vec::new(),
            collapsed_items: Vec::new(),
            show_empty_items: false,
            show_drop_downs: true,
            subtotal_top: true,
            insert_blank_row: false,
            insert_page_break: false,
            include_new_items_in_filter: false,
            item_page_count: 10,
        }
    }

    /// Set explicit subtotal functions for this axis field.
    pub fn with_subtotals<I>(mut self, subtotals: I) -> Self
    where
        I: IntoIterator<Item = PivotSubtotal>,
    {
        self.subtotals = subtotals.into_iter().collect();
        if let Some(subtotal) = self
            .subtotals
            .iter()
            .copied()
            .find(|subtotal| subtotal.is_custom_function())
        {
            self.subtotal = subtotal;
        } else if self
            .subtotals
            .iter()
            .any(|subtotal| matches!(subtotal, PivotSubtotal::None))
        {
            self.subtotal = PivotSubtotal::None;
        }
        self
    }

    /// Set items whose details are collapsed for this axis field.
    pub fn with_collapsed_items<I, V>(mut self, items: I) -> Self
    where
        I: IntoIterator<Item = V>,
        V: Into<PivotValue>,
    {
        self.collapsed_items = items.into_iter().map(Into::into).collect();
        self
    }

    /// Primary subtotal used by semantic refresh when multiple subtotal
    /// functions are present.
    pub fn primary_subtotal(&self) -> PivotSubtotal {
        self.subtotals
            .iter()
            .copied()
            .find(|subtotal| subtotal.is_custom_function())
            .unwrap_or(self.subtotal)
    }

    /// Whether this field has any subtotal enabled.
    pub fn has_enabled_subtotal(&self) -> bool {
        if self.subtotals.is_empty() {
            !matches!(self.subtotal, PivotSubtotal::None)
        } else {
            self.subtotals
                .iter()
                .any(|subtotal| !matches!(subtotal, PivotSubtotal::None))
        }
    }
}

impl From<&str> for PivotField {
    fn from(value: &str) -> Self {
        Self::new(value)
    }
}

impl From<String> for PivotField {
    fn from(value: String) -> Self {
        Self::new(value)
    }
}

impl From<PivotFieldRef> for PivotField {
    fn from(value: PivotFieldRef) -> Self {
        Self::new(value)
    }
}

/// Pivot axis identity.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotAxis {
    /// Row labels.
    Rows,
    /// Column labels.
    Columns,
    /// Report/page filters.
    Filters,
    /// Values/data fields.
    Values,
}

/// Axis used to place multiple pivot value fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotValuesAxis {
    /// Render value fields across columns.
    Columns,
    /// Render value fields down rows.
    Rows,
}

/// Sort behavior for pivot items.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotSort {
    /// Preserve source order.
    None,
    /// Sort ascending.
    Ascending,
    /// Sort descending.
    Descending,
}

/// Subtotal behavior for an axis field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotSubtotal {
    /// Let the engine/file format choose the default subtotal behavior.
    Automatic,
    /// No subtotals for this field.
    None,
    /// Sum subtotal.
    Sum,
    /// Count subtotal.
    Count,
    /// Count numeric values subtotal.
    CountNumbers,
    /// Average subtotal.
    Average,
    /// Minimum subtotal.
    Min,
    /// Maximum subtotal.
    Max,
    /// Product subtotal.
    Product,
    /// Sample standard deviation subtotal.
    StdDev,
    /// Population standard deviation subtotal.
    StdDevP,
    /// Sample variance subtotal.
    Var,
    /// Population variance subtotal.
    VarP,
}

impl PivotSubtotal {
    /// Whether this is an explicit subtotal function.
    pub fn is_custom_function(self) -> bool {
        !matches!(self, Self::Automatic | Self::None)
    }
}

/// Aggregation function for a pivot measure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotAggregate {
    /// Sum numeric values.
    Sum,
    /// Count non-blank values.
    Count,
    /// Count numeric values.
    CountNumbers,
    /// Average numeric values.
    Average,
    /// Maximum numeric value.
    Max,
    /// Minimum numeric value.
    Min,
    /// Product of numeric values.
    Product,
    /// Sample standard deviation.
    StdDev,
    /// Population standard deviation.
    StdDevP,
    /// Sample variance.
    Var,
    /// Population variance.
    VarP,
}

impl PivotAggregate {
    /// Default caption fragment for this aggregate.
    pub fn caption(self) -> &'static str {
        match self {
            Self::Sum => "Sum",
            Self::Count => "Count",
            Self::CountNumbers => "Count",
            Self::Average => "Average",
            Self::Max => "Max",
            Self::Min => "Min",
            Self::Product => "Product",
            Self::StdDev => "StdDev",
            Self::StdDevP => "StdDevP",
            Self::Var => "Var",
            Self::VarP => "VarP",
        }
    }
}

/// A value field/measure in a pivot table.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotMeasure {
    /// Source field to aggregate.
    pub field: PivotFieldRef,
    /// Aggregation function.
    pub aggregate: PivotAggregate,
    /// Optional display caption.
    pub name: Option<String>,
    /// Post-aggregation display transformation.
    pub show_as: PivotShowAs,
    /// Optional number format for rendered cells.
    pub number_format: Option<String>,
}

impl PivotMeasure {
    /// Create a measure.
    pub fn new(field: impl Into<PivotFieldRef>, aggregate: PivotAggregate) -> Self {
        Self {
            field: field.into(),
            aggregate,
            name: None,
            show_as: PivotShowAs::Normal,
            number_format: None,
        }
    }

    /// Set the display caption.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set the post-aggregation display transformation.
    pub fn with_show_as(mut self, show_as: PivotShowAs) -> Self {
        self.show_as = show_as;
        self
    }

    /// Set the display number format for rendered aggregate values.
    pub fn with_number_format(mut self, number_format: impl Into<String>) -> Self {
        self.number_format = Some(number_format.into());
        self
    }

    /// The display caption for this measure.
    pub fn caption(&self) -> String {
        self.name
            .clone()
            .unwrap_or_else(|| format!("{} of {}", self.aggregate.caption(), self.field.name))
    }
}

/// A calculated field materialized from source-row values before aggregation.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotCalculatedField {
    /// Field name exposed to axes, filters, grouping, and measures.
    pub name: String,
    /// Formula text. The formula may reference source fields by name.
    pub formula: String,
}

impl PivotCalculatedField {
    /// Create a calculated field.
    pub fn new(name: impl Into<String>, formula: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            formula: formula.into(),
        }
    }
}

/// A calculated item attached to a source field item.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotCalculatedItem {
    /// Source field containing the virtual item.
    pub field: PivotFieldRef,
    /// Virtual item name/value within the source field.
    pub item: PivotValue,
    /// Formula text. The formula may reference other items in the field.
    pub formula: String,
}

impl PivotCalculatedItem {
    /// Create a calculated item.
    pub fn new(
        field: impl Into<PivotFieldRef>,
        item: impl Into<PivotValue>,
        formula: impl Into<String>,
    ) -> Self {
        Self {
            field: field.into(),
            item: item.into(),
            formula: formula.into(),
        }
    }
}

/// Post-aggregation display transformation.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum PivotShowAs {
    /// Show the raw aggregate value.
    Normal,
    /// Percent of the grand total.
    PercentOfGrandTotal,
    /// Percent of each row total.
    PercentOfRowTotal,
    /// Percent of each column total.
    PercentOfColumnTotal,
    /// Index contribution relative to row, column, and grand totals.
    Index,
    /// Running total within a base field.
    RunningTotal {
        /// Base field.
        base_field: PivotFieldRef,
    },
    /// Difference from a base item in a base field.
    DifferenceFrom {
        /// Base field.
        base_field: PivotFieldRef,
        /// Base item.
        base_item: PivotValue,
    },
    /// Percent difference from a base item in a base field.
    PercentDifferenceFrom {
        /// Base field.
        base_field: PivotFieldRef,
        /// Base item.
        base_item: PivotValue,
    },
    /// Rank smallest to largest.
    RankAscending {
        /// Base field.
        base_field: PivotFieldRef,
    },
    /// Rank largest to smallest.
    RankDescending {
        /// Base field.
        base_field: PivotFieldRef,
    },
}

/// A pivot filter.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotFilter {
    /// Keep only the listed field items.
    FieldItems {
        /// Field to filter.
        field: PivotFieldRef,
        /// Allowed item values.
        allowed_items: Vec<PivotValue>,
    },
    /// Label/text filter.
    Label {
        /// Field to filter.
        field: PivotFieldRef,
        /// Operator.
        operator: PivotFilterOperator,
        /// String operand.
        value: String,
    },
    /// Label/text range filter.
    LabelBetween {
        /// Field to filter.
        field: PivotFieldRef,
        /// Lower bound.
        start: String,
        /// Upper bound.
        end: String,
        /// Invert the range.
        not_between: bool,
    },
    /// Date/serial-date filter.
    Date {
        /// Field to filter.
        field: PivotFieldRef,
        /// Operator.
        operator: PivotFilterOperator,
        /// Date operand as an Excel serial number.
        value: f64,
    },
    /// Date/serial-date range filter.
    DateBetween {
        /// Field to filter.
        field: PivotFieldRef,
        /// Lower bound as an Excel serial number.
        start: f64,
        /// Upper bound as an Excel serial number.
        end: f64,
        /// Invert the range.
        not_between: bool,
    },
    /// Date period filter.
    DatePeriod {
        /// Field to filter.
        field: PivotFieldRef,
        /// Period to include.
        period: PivotDatePeriod,
    },
    /// Value filter against an aggregate measure.
    Value {
        /// Field to filter.
        field: PivotFieldRef,
        /// Measure to compare.
        measure: PivotMeasure,
        /// Operator.
        operator: PivotFilterOperator,
        /// Numeric operand.
        value: f64,
    },
    /// Value range filter against an aggregate measure.
    ValueBetween {
        /// Field to filter.
        field: PivotFieldRef,
        /// Measure to compare.
        measure: PivotMeasure,
        /// Lower bound.
        start: f64,
        /// Upper bound.
        end: f64,
        /// Invert the range.
        not_between: bool,
    },
    /// Top/bottom item filter.
    TopN {
        /// Field to filter.
        field: PivotFieldRef,
        /// Measure to rank by.
        measure: PivotMeasure,
        /// Count/percent threshold.
        n: u32,
        /// Whether this is a top or bottom filter.
        top: bool,
        /// Whether `n` is a percentage.
        percent: bool,
    },
    /// Unsupported but preserved filter.
    Unsupported {
        /// Filter kind.
        kind: String,
        /// Human-readable details.
        detail: Option<String>,
    },
}

impl PivotFilter {
    /// Create an item filter.
    pub fn field_items<I, V>(field: impl Into<PivotFieldRef>, allowed_items: I) -> Self
    where
        I: IntoIterator<Item = V>,
        V: Into<PivotValue>,
    {
        Self::FieldItems {
            field: field.into(),
            allowed_items: allowed_items.into_iter().map(Into::into).collect(),
        }
    }
}

/// Date periods used by pivot date filters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotDatePeriod {
    /// Tomorrow relative to the refresh date.
    Tomorrow,
    /// Today relative to the refresh date.
    Today,
    /// Yesterday relative to the refresh date.
    Yesterday,
    /// Next week relative to the refresh date.
    NextWeek,
    /// Current week relative to the refresh date.
    ThisWeek,
    /// Previous week relative to the refresh date.
    LastWeek,
    /// Next month relative to the refresh date.
    NextMonth,
    /// Current month relative to the refresh date.
    ThisMonth,
    /// Previous month relative to the refresh date.
    LastMonth,
    /// Next quarter relative to the refresh date.
    NextQuarter,
    /// Current quarter relative to the refresh date.
    ThisQuarter,
    /// Previous quarter relative to the refresh date.
    LastQuarter,
    /// Next year relative to the refresh date.
    NextYear,
    /// Current year relative to the refresh date.
    ThisYear,
    /// Previous year relative to the refresh date.
    LastYear,
    /// From the first day of the refresh date's year through the refresh date.
    YearToDate,
    /// Any date in the given quarter, 1 through 4.
    Quarter(u8),
    /// Any date in the given month, 1 through 12.
    Month(u8),
}

/// Operators used by pivot label and value filters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotFilterOperator {
    /// Equal to.
    Equals,
    /// Not equal to.
    NotEquals,
    /// Less than.
    LessThan,
    /// Less than or equal to.
    LessThanOrEqual,
    /// Greater than.
    GreaterThan,
    /// Greater than or equal to.
    GreaterThanOrEqual,
    /// Begins with.
    BeginsWith,
    /// Does not begin with.
    DoesNotBeginWith,
    /// Ends with.
    EndsWith,
    /// Does not end with.
    DoesNotEndWith,
    /// Contains.
    Contains,
    /// Does not contain.
    DoesNotContain,
}

/// Field grouping configuration.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotGrouping {
    /// Numeric binning.
    Number {
        /// Field to group.
        field: PivotFieldRef,
        /// Optional start value.
        start: Option<f64>,
        /// Optional end value.
        end: Option<f64>,
        /// Bin interval.
        interval: f64,
    },
    /// Date/time grouping.
    Date {
        /// Field to group.
        field: PivotFieldRef,
        /// Grouping units.
        units: Vec<PivotDateGroupUnit>,
    },
    /// Manual item grouping.
    Manual {
        /// Field to group.
        field: PivotFieldRef,
        /// Item groups to apply. Source items not listed in any group remain
        /// visible under their original value.
        groups: Vec<PivotManualGroup>,
    },
}

/// A manually named group of source items in one pivot field.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotManualGroup {
    /// Display caption for the grouped items.
    pub name: String,
    /// Source item values that should roll up under `name`.
    pub members: Vec<PivotValue>,
}

impl PivotManualGroup {
    /// Create a manual item group.
    pub fn new<I, V>(name: impl Into<String>, members: I) -> Self
    where
        I: IntoIterator<Item = V>,
        V: Into<PivotValue>,
    {
        Self {
            name: name.into(),
            members: members.into_iter().map(Into::into).collect(),
        }
    }
}

/// Date grouping unit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotDateGroupUnit {
    /// Seconds.
    Seconds,
    /// Minutes.
    Minutes,
    /// Hours.
    Hours,
    /// Days.
    Days,
    /// Months.
    Months,
    /// Quarters.
    Quarters,
    /// Years.
    Years,
}

/// Pivot layout settings.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotLayout {
    /// Axis layout style.
    pub kind: PivotLayoutKind,
    /// Show grand totals for rows.
    pub show_row_grand_totals: bool,
    /// Show grand totals for columns.
    pub show_column_grand_totals: bool,
    /// Show field headers/dropdowns.
    pub show_field_headers: bool,
    /// Repeat item labels in tabular/outline layouts.
    pub repeat_item_labels: bool,
    /// Show expand/collapse buttons.
    pub show_expand_collapse: bool,
    /// Print expand/collapse buttons.
    pub print_drill_indicators: bool,
    /// Repeat pivot item labels as print titles.
    pub item_print_titles: bool,
    /// Repeat pivot field labels as print titles.
    pub field_print_titles: bool,
    /// Number of report/page fields before wrapping the filter area. Zero means no wrapping.
    pub page_wrap: u32,
    /// Lay wrapped report/page fields across rows before columns.
    pub page_over_then_down: bool,
    /// Merge repeated item labels in the rendered pivot layout when supported.
    pub merge_item_labels: bool,
    /// Caption used for the aggregate-values field.
    pub data_caption: String,
    /// Axis used for the aggregate-values field when multiple measures are present.
    pub values_axis: PivotValuesAxis,
    /// Optional zero-based position of the aggregate-values field on its axis.
    pub values_axis_position: Option<u32>,
    /// Optional caption for grand total labels.
    pub grand_total_caption: Option<String>,
    /// Optional display text for error values.
    pub error_caption: Option<String>,
    /// Whether error values should be replaced by [`Self::error_caption`].
    pub show_error: bool,
    /// Optional display text for missing values.
    pub missing_caption: Option<String>,
    /// Whether missing values should be replaced by [`Self::missing_caption`].
    pub show_missing: bool,
    /// Show asterisks beside visual total captions.
    pub asterisk_totals: bool,
    /// Show items with no data.
    pub show_items: bool,
    /// Allow users to edit pivot table values.
    pub edit_data: bool,
    /// Disable the field list UI for this pivot.
    pub disable_field_list: bool,
    /// Show calculated members.
    pub show_calculated_members: bool,
    /// Use visual totals for filtered items.
    pub visual_totals: bool,
    /// Show the multiple-items label for page fields.
    pub show_multiple_label: bool,
    /// Show the data field dropdown.
    pub show_data_drop_down: bool,
    /// Show member property tooltips.
    pub show_member_property_tips: bool,
    /// Show data tooltips.
    pub show_data_tips: bool,
    /// Enable the classic pivot table wizard.
    pub enable_wizard: bool,
    /// Enable drill actions.
    pub enable_drill: bool,
    /// Enable field property editing.
    pub enable_field_properties: bool,
    /// Include hidden items when calculating subtotals.
    pub subtotal_hidden_items: bool,
    /// Show PivotChart drop zones.
    pub show_drop_zones: bool,
    /// Row field indentation level in compact/outline layouts.
    pub indent: u32,
    /// Show empty rows.
    pub show_empty_rows: bool,
    /// Show empty columns.
    pub show_empty_columns: bool,
}

impl Default for PivotLayout {
    fn default() -> Self {
        Self {
            kind: PivotLayoutKind::Compact,
            show_row_grand_totals: true,
            show_column_grand_totals: true,
            show_field_headers: true,
            repeat_item_labels: false,
            show_expand_collapse: true,
            print_drill_indicators: false,
            item_print_titles: false,
            field_print_titles: false,
            page_wrap: 0,
            page_over_then_down: false,
            merge_item_labels: false,
            data_caption: "Values".to_string(),
            values_axis: PivotValuesAxis::Columns,
            values_axis_position: None,
            grand_total_caption: None,
            error_caption: None,
            show_error: false,
            missing_caption: None,
            show_missing: true,
            asterisk_totals: false,
            show_items: true,
            edit_data: false,
            disable_field_list: false,
            show_calculated_members: true,
            visual_totals: true,
            show_multiple_label: true,
            show_data_drop_down: true,
            show_member_property_tips: true,
            show_data_tips: true,
            enable_wizard: true,
            enable_drill: true,
            enable_field_properties: true,
            subtotal_hidden_items: false,
            show_drop_zones: true,
            indent: 1,
            show_empty_rows: false,
            show_empty_columns: false,
        }
    }
}

/// Pivot layout kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotLayoutKind {
    /// Excel compact form.
    Compact,
    /// Outline form.
    Outline,
    /// Tabular form.
    Tabular,
}

/// Pivot style settings.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotStyle {
    /// Built-in or custom style name.
    pub name: Option<String>,
    /// Show row header styling.
    pub show_row_headers: bool,
    /// Show column header styling.
    pub show_column_headers: bool,
    /// Show alternating row stripes.
    pub show_row_stripes: bool,
    /// Show alternating column stripes.
    pub show_column_stripes: bool,
    /// Show last-column styling.
    pub show_last_column: bool,
}

impl Default for PivotStyle {
    fn default() -> Self {
        Self {
            name: Some("PivotStyleMedium9".to_string()),
            show_row_headers: true,
            show_column_headers: true,
            show_row_stripes: false,
            show_column_stripes: false,
            show_last_column: false,
        }
    }
}

/// Pivot refresh behavior.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotRefreshPolicy {
    /// Refresh when the workbook is opened.
    pub refresh_on_open: bool,
    /// Preserve user formatting when refreshed.
    pub preserve_formatting: bool,
    /// Whether refresh may be performed asynchronously by external engines.
    pub background_query: bool,
    /// Optional cap for retained missing items.
    pub missing_items_limit: Option<u32>,
}

impl Default for PivotRefreshPolicy {
    fn default() -> Self {
        Self {
            refresh_on_open: false,
            preserve_formatting: true,
            background_query: false,
            missing_items_limit: None,
        }
    }
}

/// What to do with existing cells in the pivot output area.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotOverwritePolicy {
    /// Clear the previously-rendered pivot range, then write the new result.
    ClearOwnedRange,
    /// Overwrite the target range without checking existing cells.
    Overwrite,
    /// Fail if the output would write over a non-empty cell outside the previous
    /// pivot output range.
    FailOnOccupied,
}

impl Default for PivotOverwritePolicy {
    fn default() -> Self {
        Self::ClearOwnedRange
    }
}

/// Extension payload for format-specific pivot metadata.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotExtension {
    /// Namespace or extension URI.
    pub uri: String,
    /// Raw payload bytes in the source format.
    ///
    /// For XLSX pivot tables this is the complete `<ext>` element so namespace
    /// declarations and unknown extension content can be preserved losslessly.
    pub payload: Vec<u8>,
}

/// A collection of extension payloads.
pub type PivotExtensions = Vec<PivotExtension>;

/// Read-only diagnostic information about the format-level cache backing a pivot.
///
/// This is not an authoring API. It lets readers report what was present in a
/// file while pivot generation remains driven by [`PivotTable`].
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotCacheInfo {
    /// Format-level cache ID, if known.
    pub cache_id: u32,
    /// Cache source kind.
    pub source_kind: PivotCacheSourceKind,
    /// Number of cache records, if known.
    pub record_count: Option<u64>,
    /// Application version that last refreshed the cache, if known.
    pub refreshed_version: Option<String>,
    /// Latest refresh status.
    pub refresh_status: PivotRefreshStatus,
}

/// Pivot cache source kind reported by a reader.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotCacheSourceKind {
    /// Worksheet/table source.
    Worksheet,
    /// External connection source.
    External,
    /// Consolidation source.
    Consolidation,
    /// Scenario source.
    Scenario,
    /// OLAP/data-model source.
    Olap,
    /// Unknown source.
    Unknown,
}

/// Latest known pivot refresh status.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum PivotRefreshStatus {
    /// No refresh has been attempted in this process.
    NotRefreshed,
    /// Refresh succeeded.
    Succeeded,
    /// Refresh failed.
    Failed {
        /// Error details.
        message: String,
    },
    /// Refresh is handled by an external engine.
    External,
}

impl Default for PivotRefreshStatus {
    fn default() -> Self {
        Self::NotRefreshed
    }
}

/// A pivot table definition attached to a worksheet.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotTable {
    /// Unique pivot table ID within the workbook. `0` means "assign on insert".
    pub id: u32,
    /// Pivot table name.
    pub name: String,
    /// Source data.
    pub source: PivotSource,
    /// Top-left output cell.
    pub target: CellAddress,
    /// Row axis fields.
    pub rows: Vec<PivotField>,
    /// Column axis fields.
    pub columns: Vec<PivotField>,
    /// Report/page filter axis fields.
    pub page_fields: Vec<PivotField>,
    /// Filter criteria.
    pub filters: Vec<PivotFilter>,
    /// Calculated fields materialized before aggregation.
    pub calculated_fields: Vec<PivotCalculatedField>,
    /// Calculated items preserved in the pivot cache definition.
    pub calculated_items: Vec<PivotCalculatedItem>,
    /// Value fields.
    pub measures: Vec<PivotMeasure>,
    /// Grouping definitions.
    pub groupings: Vec<PivotGrouping>,
    /// Layout settings.
    pub layout: PivotLayout,
    /// Style settings.
    pub style: PivotStyle,
    /// Refresh policy.
    pub refresh_policy: PivotRefreshPolicy,
    /// Output overwrite policy.
    pub overwrite_policy: PivotOverwritePolicy,
    /// Last rendered output range, if this pivot has been refreshed by an engine.
    pub rendered_range: Option<CellRange>,
    /// Latest semantic refresh status.
    pub refresh_status: PivotRefreshStatus,
    cache_info: Option<PivotCacheInfo>,
    /// Format-specific extension payloads.
    pub extensions: PivotExtensions,
}

impl PivotTable {
    /// Create a pivot table with default layout, style, and refresh policy.
    pub fn new(id: u32, name: impl Into<String>, source: PivotSource, target: CellAddress) -> Self {
        Self {
            id,
            name: name.into(),
            source,
            target,
            rows: Vec::new(),
            columns: Vec::new(),
            page_fields: Vec::new(),
            filters: Vec::new(),
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            measures: Vec::new(),
            groupings: Vec::new(),
            layout: PivotLayout::default(),
            style: PivotStyle::default(),
            refresh_policy: PivotRefreshPolicy::default(),
            overwrite_policy: PivotOverwritePolicy::default(),
            rendered_range: None,
            refresh_status: PivotRefreshStatus::NotRefreshed,
            cache_info: None,
            extensions: Vec::new(),
        }
    }

    /// Start building a pivot table.
    pub fn builder(name: impl Into<String>) -> PivotTableBuilder {
        PivotTableBuilder::new(name)
    }

    /// Reader-provided cache diagnostics, if the source file contained them.
    ///
    /// This is intentionally read-only. Format-level pivot caches are internal
    /// reader/writer artifacts; authoring should continue to use semantic
    /// fields, measures, filters, grouping, layout, and refresh policy.
    pub fn cache_info(&self) -> Option<&PivotCacheInfo> {
        self.cache_info.as_ref()
    }

    /// Set reader-provided cache diagnostics.
    #[doc(hidden)]
    pub fn set_cache_info(&mut self, cache_info: Option<PivotCacheInfo>) {
        self.cache_info = cache_info;
    }

    /// Update the cache diagnostic refresh status when diagnostics exist.
    #[doc(hidden)]
    pub fn set_cache_refresh_status(&mut self, refresh_status: PivotRefreshStatus) {
        if let Some(cache_info) = &mut self.cache_info {
            cache_info.refresh_status = refresh_status;
        }
    }
}

/// Builder for [`PivotTable`].
#[derive(Debug, Clone)]
pub struct PivotTableBuilder {
    id: u32,
    name: String,
    source: Option<PivotSource>,
    target: Option<CellAddress>,
    rows: Vec<PivotField>,
    columns: Vec<PivotField>,
    page_fields: Vec<PivotField>,
    filters: Vec<PivotFilter>,
    calculated_fields: Vec<PivotCalculatedField>,
    calculated_items: Vec<PivotCalculatedItem>,
    measures: Vec<PivotMeasure>,
    groupings: Vec<PivotGrouping>,
    layout: PivotLayout,
    style: PivotStyle,
    refresh_policy: PivotRefreshPolicy,
    overwrite_policy: PivotOverwritePolicy,
}

impl PivotTableBuilder {
    /// Create a builder with the given pivot name.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            id: 0,
            name: name.into(),
            source: None,
            target: None,
            rows: Vec::new(),
            columns: Vec::new(),
            page_fields: Vec::new(),
            filters: Vec::new(),
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            measures: Vec::new(),
            groupings: Vec::new(),
            layout: PivotLayout::default(),
            style: PivotStyle::default(),
            refresh_policy: PivotRefreshPolicy::default(),
            overwrite_policy: PivotOverwritePolicy::default(),
        }
    }

    /// Set the pivot ID. `0` lets the worksheet assign an ID on insert.
    pub fn id(mut self, id: u32) -> Self {
        self.id = id;
        self
    }

    /// Set the source.
    pub fn source(mut self, source: PivotSource) -> Self {
        self.source = Some(source);
        self
    }

    /// Set a worksheet-range source on the pivot table's sheet.
    pub fn source_range(mut self, range: CellRange) -> Self {
        self.source = Some(PivotSource::range(range));
        self
    }

    /// Set a worksheet-range source on a named sheet.
    pub fn source_range_on_sheet(mut self, sheet: impl Into<String>, range: CellRange) -> Self {
        self.source = Some(PivotSource::range_on_sheet(sheet, range));
        self
    }

    /// Set a table source by name.
    pub fn table_source(mut self, name: impl Into<String>) -> Self {
        self.source = Some(PivotSource::table(name));
        self
    }

    /// Set the top-left target cell.
    pub fn target_cell(mut self, target: CellAddress) -> Self {
        self.target = Some(target);
        self
    }

    /// Set the top-left target cell from A1 notation.
    pub fn target_address(mut self, target: &str) -> Result<Self> {
        self.target = Some(CellAddress::parse(target)?);
        Ok(self)
    }

    /// Add a row-axis field.
    pub fn row(mut self, field: impl Into<PivotField>) -> Self {
        self.rows.push(field.into());
        self
    }

    /// Add a column-axis field.
    pub fn column(mut self, field: impl Into<PivotField>) -> Self {
        self.columns.push(field.into());
        self
    }

    /// Add a report/page filter-axis field.
    pub fn page(mut self, field: impl Into<PivotField>) -> Self {
        self.page_fields.push(field.into());
        self
    }

    /// Add a value field.
    pub fn measure(mut self, field: impl Into<PivotFieldRef>, aggregate: PivotAggregate) -> Self {
        self.measures.push(PivotMeasure::new(field, aggregate));
        self
    }

    /// Add a value field with a custom caption.
    pub fn named_measure(
        mut self,
        field: impl Into<PivotFieldRef>,
        aggregate: PivotAggregate,
        name: impl Into<String>,
    ) -> Self {
        self.measures
            .push(PivotMeasure::new(field, aggregate).with_name(name));
        self
    }

    /// Add a fully configured measure.
    pub fn pivot_measure(mut self, measure: PivotMeasure) -> Self {
        self.measures.push(measure);
        self
    }

    /// Add a filter.
    pub fn filter(mut self, filter: PivotFilter) -> Self {
        self.filters.push(filter);
        self
    }

    /// Add a calculated field.
    pub fn calculated_field(mut self, name: impl Into<String>, formula: impl Into<String>) -> Self {
        self.calculated_fields
            .push(PivotCalculatedField::new(name, formula));
        self
    }

    /// Add a fully configured calculated field.
    pub fn pivot_calculated_field(mut self, field: PivotCalculatedField) -> Self {
        self.calculated_fields.push(field);
        self
    }

    /// Add a calculated item.
    pub fn calculated_item(
        mut self,
        field: impl Into<PivotFieldRef>,
        item: impl Into<PivotValue>,
        formula: impl Into<String>,
    ) -> Self {
        self.calculated_items
            .push(PivotCalculatedItem::new(field, item, formula));
        self
    }

    /// Add a fully configured calculated item.
    pub fn pivot_calculated_item(mut self, item: PivotCalculatedItem) -> Self {
        self.calculated_items.push(item);
        self
    }

    /// Add a grouping definition.
    pub fn grouping(mut self, grouping: PivotGrouping) -> Self {
        self.groupings.push(grouping);
        self
    }

    /// Set layout settings.
    pub fn layout(mut self, layout: PivotLayout) -> Self {
        self.layout = layout;
        self
    }

    /// Set style settings.
    pub fn style(mut self, style: PivotStyle) -> Self {
        self.style = style;
        self
    }

    /// Set refresh policy.
    pub fn refresh_policy(mut self, refresh_policy: PivotRefreshPolicy) -> Self {
        self.refresh_policy = refresh_policy;
        self
    }

    /// Set output overwrite policy.
    pub fn overwrite_policy(mut self, overwrite_policy: PivotOverwritePolicy) -> Self {
        self.overwrite_policy = overwrite_policy;
        self
    }

    /// Build the pivot table.
    pub fn build(self) -> Result<PivotTable> {
        if self.name.trim().is_empty() {
            return Err(Error::other("pivot table name cannot be empty"));
        }
        if self.measures.is_empty() {
            return Err(Error::other(
                "pivot table must contain at least one measure",
            ));
        }

        let source = self
            .source
            .ok_or_else(|| Error::other("pivot table source is required"))?;
        let target = self
            .target
            .ok_or_else(|| Error::other("pivot table target is required"))?;

        let mut table = PivotTable::new(self.id, self.name, source, target);
        table.rows = self.rows;
        table.columns = self.columns;
        table.page_fields = self.page_fields;
        table.filters = self.filters;
        table.calculated_fields = self.calculated_fields;
        table.calculated_items = self.calculated_items;
        table.measures = self.measures;
        table.groupings = self.groupings;
        table.layout = self.layout;
        table.style = self.style;
        table.refresh_policy = self.refresh_policy;
        table.overwrite_policy = self.overwrite_policy;
        Ok(table)
    }
}

#[cfg(test)]
mod tests {
    use std::collections::HashSet;

    use super::*;

    #[test]
    fn pivot_value_hashes_numbers_by_bits() {
        let mut values = HashSet::new();
        values.insert(PivotValue::Number(0.0));
        values.insert(PivotValue::Number(-0.0));

        assert_eq!(values.len(), 2);
    }

    #[test]
    fn builder_requires_source_target_and_measure() {
        let err = PivotTable::builder("Sales").build().unwrap_err();
        assert!(err.to_string().contains("measure"));

        let err = PivotTable::builder("Sales")
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap_err();
        assert!(err.to_string().contains("source"));

        let err = PivotTable::builder("Sales")
            .source_range(CellRange::parse("A1:C10").unwrap())
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap_err();
        assert!(err.to_string().contains("target"));
    }

    #[test]
    fn builder_creates_semantic_pivot() {
        let pivot = PivotTable::builder("Sales")
            .source_range(CellRange::parse("A1:C10").unwrap())
            .target_address("E3")
            .unwrap()
            .row("Region")
            .column("Quarter")
            .page("Salesperson")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .filter(PivotFilter::field_items("Region", ["East", "West"]))
            .build()
            .unwrap();

        assert_eq!(pivot.name, "Sales");
        assert_eq!(pivot.target, CellAddress::parse("E3").unwrap());
        assert_eq!(pivot.rows[0].field.name, "Region");
        assert_eq!(pivot.columns[0].field.name, "Quarter");
        assert_eq!(pivot.page_fields[0].field.name, "Salesperson");
        assert_eq!(pivot.measures[0].caption(), "Revenue");
    }
}
