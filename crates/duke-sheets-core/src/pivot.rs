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
}

impl PivotSourceRange {
    /// Create a source range.
    pub fn new(sheet: impl Into<String>, range: CellRange) -> Self {
        Self {
            sheet: sheet.into(),
            range,
        }
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
    /// Whether items with no data should be shown.
    pub show_empty_items: bool,
}

impl PivotField {
    /// Create an axis field with default sorting and subtotals.
    pub fn new(field: impl Into<PivotFieldRef>) -> Self {
        Self {
            field: field.into(),
            sort: PivotSort::Ascending,
            subtotal: PivotSubtotal::Automatic,
            show_empty_items: false,
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
}

impl Default for PivotStyle {
    fn default() -> Self {
        Self {
            name: Some("PivotStyleMedium9".to_string()),
            show_row_headers: true,
            show_column_headers: true,
            show_row_stripes: false,
            show_column_stripes: false,
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
    /// Reader-provided cache diagnostics, if the source file contained them.
    pub cache_info: Option<PivotCacheInfo>,
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
