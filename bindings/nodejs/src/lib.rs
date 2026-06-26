//! Node.js/TypeScript bindings for duke-sheets
//!
//! This module provides NAPI-RS-based native Node.js bindings for the duke-sheets
//! library, allowing JavaScript/TypeScript code to read, write, and manipulate
//! Excel files with native performance.

use napi::bindgen_prelude::*;
use napi_derive::napi;

use std::io::Cursor;
use std::path::PathBuf;
use std::sync::{Arc, RwLock};

use duke_sheets::{
    CalculationOptions, CalculationStats as CoreCalculationStats, FormulaValue, ImageSizing,
    WorkbookCalculationExt, WorkbookExt, WorkbookPivotExt,
};
use duke_sheets_core::{
    CellAddress, CellError, CellRange, CellValue as CoreCellValue, PivotAggregate,
    PivotDateGroupUnit, PivotField, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayout,
    PivotLayoutKind, PivotManualGroup, PivotMeasure, PivotOverwritePolicy, PivotRefreshPolicy,
    PivotShowAs, PivotSort, PivotStyle, PivotSubtotal, PivotTable, PivotValue,
    Workbook as CoreWorkbook,
};

fn to_napi_err(e: impl std::fmt::Display) -> napi::Error {
    napi::Error::from_reason(e.to_string())
}

fn parse_encryption_profile(
    profile: Option<&str>,
    key_bits: Option<u32>,
    spin_count: Option<u32>,
) -> std::result::Result<duke_sheets::EncryptionProfile, String> {
    use duke_sheets::EncryptionProfile;
    let normalized = profile.map(|p| p.to_lowercase());
    Ok(match normalized.as_deref() {
        None | Some("default") => EncryptionProfile::Default,
        Some("agile") | Some("ooxml-agile") => EncryptionProfile::OoxmlAgile {
            key_bits: key_bits.unwrap_or(256),
            spin_count: spin_count.unwrap_or(100_000),
        },
        Some("standard") | Some("ooxml-standard") => EncryptionProfile::OoxmlStandard {
            key_bits: key_bits.unwrap_or(128),
        },
        Some("rc4-cryptoapi") | Some("xls-rc4-cryptoapi") => EncryptionProfile::XlsRc4CryptoApi {
            key_bits: key_bits.unwrap_or(128),
        },
        Some("rc4-legacy") | Some("xls-rc4-legacy") => EncryptionProfile::XlsRc4Legacy,
        Some("xor") | Some("xls-xor") => EncryptionProfile::XlsXor,
        Some(other) => return Err(format!("unknown encryption profile: {other:?}")),
    })
}

/// Catch Rust panics at the FFI boundary and convert them to napi::Error.
///
/// Without this, a panic (e.g. integer overflow, index out of bounds) would
/// unwind across the FFI boundary into Node.js, which is undefined behavior
/// and kills the process. With this wrapper, panics become JS exceptions that
/// callers can catch normally.
pub(crate) fn catch_panic<T>(f: impl FnOnce() -> napi::Result<T>) -> napi::Result<T> {
    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(f)) {
        Ok(result) => result,
        Err(payload) => {
            let msg = if let Some(s) = payload.downcast_ref::<&str>() {
                s.to_string()
            } else if let Some(s) = payload.downcast_ref::<String>() {
                s.clone()
            } else {
                "unknown error".to_string()
            };
            Err(napi::Error::from_reason(format!("Internal error: {}", msg)))
        }
    }
}

mod types;
pub use types::*;
mod workbook_read;
mod worksheet_read;

fn cell_error_to_string(e: &CellError) -> &'static str {
    match e {
        CellError::Div0 => "#DIV/0!",
        CellError::Na => "#N/A",
        CellError::Name => "#NAME?",
        CellError::Null => "#NULL!",
        CellError::Num => "#NUM!",
        CellError::Ref => "#REF!",
        CellError::Value => "#VALUE!",
        CellError::GettingData => "#GETTING_DATA",
        CellError::Spill => "#SPILL!",
        CellError::Calc => "#CALC!",
    }
}

/// Extract an optional property from a JS object.
/// Returns `Ok(None)` if the property is missing, null, or undefined.
fn try_get_property<'a, T: FromNapiValue + ValidateNapiValue>(
    obj: &Object<'a>,
    key: &str,
) -> napi::Result<Option<T>> {
    if !obj.has_named_property(key)? {
        return Ok(None);
    }
    let val: Unknown<'_> = obj.get_named_property(key)?;
    let val_type = val.get_type()?;
    if val_type == ValueType::Null || val_type == ValueType::Undefined {
        return Ok(None);
    }
    let v = obj.value();
    unsafe { T::from_napi_value(v.env, val.raw()).map(Some) }
}

/// Represents a cell value in a spreadsheet.
///
/// Cell values can be one of several types:
/// - Empty (null/undefined)
/// - Number
/// - Text (string)
/// - Boolean
/// - Error (like "#DIV/0!")
/// - Formula cached results are exposed as regular cell values; formula text lives on Worksheet accessors
#[napi]
pub struct CellValue {
    inner: CoreCellValue,
}

#[napi]
impl CellValue {
    /// Check if the cell is empty
    #[napi(getter)]
    pub fn is_empty(&self) -> bool {
        matches!(self.inner, CoreCellValue::Empty)
    }

    /// Check if the cell contains a number
    #[napi(getter)]
    pub fn is_number(&self) -> bool {
        matches!(self.inner, CoreCellValue::Number(_))
    }

    /// Check if the cell contains text
    #[napi(getter)]
    pub fn is_text(&self) -> bool {
        matches!(self.inner, CoreCellValue::String(_))
    }

    /// Check if the cell contains a boolean
    #[napi(getter)]
    pub fn is_boolean(&self) -> bool {
        matches!(self.inner, CoreCellValue::Boolean(_))
    }

    /// Check if the cell contains an error
    #[napi(getter)]
    pub fn is_error(&self) -> bool {
        matches!(self.inner, CoreCellValue::Error(_))
    }

    /// Get the value as a number, or null if not a number
    #[napi]
    pub fn as_number(&self) -> Option<f64> {
        match &self.inner {
            CoreCellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    /// Get the value as text, or null if not text
    #[napi]
    pub fn as_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::String(s) => Some(s.to_string()),
            _ => None,
        }
    }

    /// Get the value as a boolean, or null if not a boolean
    #[napi]
    pub fn as_boolean(&self) -> Option<bool> {
        match &self.inner {
            CoreCellValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    /// Get the error string, or null if not an error
    #[napi]
    pub fn as_error(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::Error(e) => Some(cell_error_to_string(e).to_string()),
            _ => None,
        }
    }

    /// Convert to a JavaScript native value (number, string, boolean, or null)
    #[napi(js_name = "toJs")]
    pub fn to_js(&self) -> Either4<f64, String, bool, Null> {
        match &self.inner {
            CoreCellValue::Empty => Either4::D(Null),
            CoreCellValue::Number(n) => Either4::A(*n),
            CoreCellValue::String(s) => Either4::B(s.to_string()),
            CoreCellValue::Boolean(b) => Either4::C(*b),
            CoreCellValue::Error(e) => Either4::B(cell_error_to_string(e).to_string()),
            _ => Either4::D(Null),
        }
    }

    /// Get string representation of the cell value
    #[napi(js_name = "toString")]
    pub fn to_string_js(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => String::new(),
            CoreCellValue::Number(n) => n.to_string(),
            CoreCellValue::String(s) => s.to_string(),
            CoreCellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CoreCellValue::Error(e) => cell_error_to_string(e).to_string(),
            _ => String::new(),
        }
    }
}

/// Options for workbook calculation.
///
/// All fields are optional and default to sensible values.
#[napi(object)]
pub struct JsCalculationOptions {
    /// Enable iterative calculation for circular references (default: false)
    pub iterative: Option<bool>,
    /// Maximum iterations for circular references (default: 100)
    pub max_iterations: Option<u32>,
    /// Maximum change threshold for convergence (default: 0.001)
    pub max_change: Option<f64>,
    /// Force recalculation of all cells, even if not dirty (default: true)
    pub force_full_calculation: Option<bool>,
    /// Include volatile functions in calculation (NOW, TODAY, RAND, etc.) (default: true)
    pub calculate_volatile: Option<bool>,
    /// Only calculate these sheet indices (and their transitive cross-sheet dependencies).
    /// If empty or omitted, calculate all sheets.
    pub sheets: Option<Vec<u32>>,
    /// Maximum number of threads for parallel evaluation.
    /// - null/undefined: use all available cores
    /// - 1: force serial evaluation
    /// - n: use at most n threads
    pub max_threads: Option<u32>,
}

impl JsCalculationOptions {
    fn into_core(self) -> CalculationOptions {
        CalculationOptions {
            iterative: self.iterative.unwrap_or(false),
            max_iterations: self.max_iterations.unwrap_or(100),
            max_change: self.max_change.unwrap_or(0.001),
            force_full_calculation: self.force_full_calculation.unwrap_or(true),
            calculate_volatile: self.calculate_volatile.unwrap_or(true),
            sheets: self
                .sheets
                .unwrap_or_default()
                .into_iter()
                .map(|i| i as usize)
                .collect(),
            max_threads: self.max_threads.map(|n| n as usize),
            web_service_fn: None,
            rtd_fn: None,
            external_fn: None,
        }
    }
}

// JsImageInfo moved to types.rs

/// Statistics from calculating a workbook.
#[napi]
pub struct CalculationStats {
    inner: CoreCalculationStats,
}

#[napi]
impl CalculationStats {
    /// Number of formulas found
    #[napi(getter)]
    pub fn formula_count(&self) -> u32 {
        self.inner.formula_count as u32
    }

    /// Number of cells calculated
    #[napi(getter)]
    pub fn cells_calculated(&self) -> u32 {
        self.inner.cells_calculated as u32
    }

    /// Number of errors encountered
    #[napi(getter)]
    pub fn errors(&self) -> u32 {
        self.inner.errors as u32
    }

    /// Number of circular references detected
    #[napi(getter)]
    pub fn circular_references(&self) -> u32 {
        self.inner.circular_references as u32
    }

    /// Number of volatile cells (e.g., NOW(), RAND())
    #[napi(getter)]
    pub fn volatile_cells(&self) -> u32 {
        self.inner.volatile_cells as u32
    }

    /// Whether iterative calculation converged
    #[napi(getter)]
    pub fn converged(&self) -> bool {
        self.inner.converged
    }

    /// Number of iterations performed
    #[napi(getter)]
    pub fn iterations(&self) -> u32 {
        self.inner.iterations as u32
    }
}

/// The used range of a worksheet, describing the bounding box of all cells
/// that contain data.
#[napi(object)]
pub struct UsedRange {
    pub min_row: u32,
    pub min_col: u32,
    pub max_row: u32,
    pub max_col: u32,
}

#[napi(object)]
pub struct JsPivotMeasureOptions {
    pub field: String,
    pub aggregate: Option<String>,
    pub name: Option<String>,
    pub show_as: Option<String>,
    pub base_field: Option<String>,
    pub base_item: Option<Either3<f64, String, bool>>,
    pub number_format: Option<String>,
}

#[napi(object)]
pub struct JsPivotFilterOptions {
    pub kind: Option<String>,
    pub field: String,
    pub items: Option<Vec<String>>,
    pub operator: Option<String>,
    pub text: Option<String>,
    pub measure: Option<JsPivotMeasureOptions>,
    pub value: Option<f64>,
    pub n: Option<u32>,
    pub top: Option<bool>,
    pub percent: Option<bool>,
}

#[napi(object)]
pub struct JsPivotCalculatedFieldOptions {
    pub name: String,
    pub formula: String,
}

#[napi(object)]
pub struct JsPivotManualGroupOptions {
    pub name: String,
    pub members: Vec<Either3<f64, String, bool>>,
}

#[napi(object)]
pub struct JsPivotGroupingOptions {
    pub field: String,
    pub kind: String,
    pub start: Option<f64>,
    pub end: Option<f64>,
    pub interval: Option<f64>,
    pub units: Option<Vec<String>>,
    pub groups: Option<Vec<JsPivotManualGroupOptions>>,
}

#[napi(object)]
pub struct JsPivotFieldOptions {
    pub field: String,
    pub sort: Option<String>,
    pub subtotal: Option<String>,
    pub show_empty_items: Option<bool>,
}

#[napi(object)]
pub struct JsPivotRefreshPolicyOptions {
    pub refresh_on_open: Option<bool>,
    pub preserve_formatting: Option<bool>,
    pub background_query: Option<bool>,
    pub missing_items_limit: Option<u32>,
}

#[napi(object)]
pub struct JsPivotLayoutOptions {
    pub kind: Option<String>,
    pub show_row_grand_totals: Option<bool>,
    pub show_column_grand_totals: Option<bool>,
    pub show_field_headers: Option<bool>,
    pub repeat_item_labels: Option<bool>,
    pub show_expand_collapse: Option<bool>,
    pub print_drill_indicators: Option<bool>,
    pub item_print_titles: Option<bool>,
    pub field_print_titles: Option<bool>,
}

#[napi(object)]
pub struct JsPivotStyleOptions {
    pub name: Option<String>,
    pub show_row_headers: Option<bool>,
    pub show_column_headers: Option<bool>,
    pub show_row_stripes: Option<bool>,
    pub show_column_stripes: Option<bool>,
    pub show_last_column: Option<bool>,
}

#[napi(object)]
pub struct JsPivotTableOptions {
    pub name: String,
    pub source_range: Option<String>,
    pub source_sheet: Option<String>,
    pub table_name: Option<String>,
    pub target: String,
    pub rows: Option<Vec<String>>,
    pub columns: Option<Vec<String>>,
    pub pages: Option<Vec<String>>,
    pub row_fields: Option<Vec<JsPivotFieldOptions>>,
    pub column_fields: Option<Vec<JsPivotFieldOptions>>,
    pub page_fields: Option<Vec<JsPivotFieldOptions>>,
    pub measures: Vec<JsPivotMeasureOptions>,
    pub filters: Option<Vec<JsPivotFilterOptions>>,
    pub calculated_fields: Option<Vec<JsPivotCalculatedFieldOptions>>,
    pub groupings: Option<Vec<JsPivotGroupingOptions>>,
    pub refresh_policy: Option<JsPivotRefreshPolicyOptions>,
    pub layout: Option<JsPivotLayoutOptions>,
    pub style: Option<JsPivotStyleOptions>,
    pub overwrite_policy: Option<String>,
}

#[napi(object)]
pub struct JsPivotRefreshStats {
    pub pivot_count: u32,
    pub pivots_refreshed: u32,
    pub source_rows: u32,
    pub output_cells: u32,
    pub cache_hits: u32,
    pub cache_misses: u32,
}

impl TryFrom<duke_sheets::PivotRefreshStats> for JsPivotRefreshStats {
    type Error = napi::Error;

    fn try_from(stats: duke_sheets::PivotRefreshStats) -> Result<Self> {
        Ok(Self {
            pivot_count: u32::try_from(stats.pivot_count).map_err(to_napi_err)?,
            pivots_refreshed: u32::try_from(stats.pivots_refreshed).map_err(to_napi_err)?,
            source_rows: u32::try_from(stats.source_rows).map_err(to_napi_err)?,
            output_cells: u32::try_from(stats.output_cells).map_err(to_napi_err)?,
            cache_hits: u32::try_from(stats.cache_hits).map_err(to_napi_err)?,
            cache_misses: u32::try_from(stats.cache_misses).map_err(to_napi_err)?,
        })
    }
}

fn build_pivot_table_from_js(options: JsPivotTableOptions) -> Result<PivotTable> {
    let mut builder = PivotTable::builder(options.name);
    match (options.table_name, options.source_range) {
        (Some(table_name), None) => {
            builder = builder.table_source(table_name);
        }
        (None, Some(source_range)) => {
            let range = CellRange::parse(&source_range).map_err(|e| {
                napi::Error::from_reason(format!("Invalid pivot source range: {e}"))
            })?;
            builder = if let Some(sheet) = options.source_sheet {
                builder.source_range_on_sheet(sheet, range)
            } else {
                builder.source_range(range)
            };
        }
        (Some(_), Some(_)) => {
            return Err(napi::Error::from_reason(
                "Pivot options must use either tableName or sourceRange, not both",
            ));
        }
        (None, None) => {
            return Err(napi::Error::from_reason(
                "Pivot options require tableName or sourceRange",
            ));
        }
    }

    builder = builder
        .target_address(&options.target)
        .map_err(|e| napi::Error::from_reason(format!("Invalid pivot target: {e}")))?;
    for field in options.rows.unwrap_or_default() {
        builder = builder.row(field);
    }
    for field in options.columns.unwrap_or_default() {
        builder = builder.column(field);
    }
    for field in options.pages.unwrap_or_default() {
        builder = builder.page(field);
    }
    for field in options.row_fields.unwrap_or_default() {
        builder = builder.row(build_pivot_field_from_js(field)?);
    }
    for field in options.column_fields.unwrap_or_default() {
        builder = builder.column(build_pivot_field_from_js(field)?);
    }
    for field in options.page_fields.unwrap_or_default() {
        builder = builder.page(build_pivot_field_from_js(field)?);
    }
    for measure in options.measures {
        builder = builder.pivot_measure(build_pivot_measure_from_js(measure)?);
    }
    for filter in options.filters.unwrap_or_default() {
        builder = builder.filter(build_pivot_filter_from_js(filter)?);
    }
    for calculated_field in options.calculated_fields.unwrap_or_default() {
        builder = builder.calculated_field(calculated_field.name, calculated_field.formula);
    }
    for grouping in options.groupings.unwrap_or_default() {
        builder = builder.grouping(build_pivot_grouping_from_js(grouping)?);
    }
    if let Some(refresh_policy) = options.refresh_policy {
        builder = builder.refresh_policy(build_pivot_refresh_policy_from_js(refresh_policy));
    }
    if let Some(layout) = options.layout {
        builder = builder.layout(build_pivot_layout_from_js(layout)?);
    }
    if let Some(style) = options.style {
        builder = builder.style(build_pivot_style_from_js(style));
    }
    if let Some(overwrite_policy) = options.overwrite_policy {
        builder = builder.overwrite_policy(parse_pivot_overwrite_policy(&overwrite_policy)?);
    }

    builder.build().map_err(to_napi_err)
}

fn build_pivot_field_from_js(options: JsPivotFieldOptions) -> Result<PivotField> {
    let mut field = PivotField::new(options.field);
    if let Some(sort) = options.sort {
        field.sort = parse_pivot_sort(&sort)?;
    }
    if let Some(subtotal) = options.subtotal {
        field.subtotal = parse_pivot_subtotal(&subtotal)?;
    }
    if let Some(show_empty_items) = options.show_empty_items {
        field.show_empty_items = show_empty_items;
    }
    Ok(field)
}

fn build_pivot_refresh_policy_from_js(options: JsPivotRefreshPolicyOptions) -> PivotRefreshPolicy {
    let mut policy = PivotRefreshPolicy::default();
    if let Some(value) = options.refresh_on_open {
        policy.refresh_on_open = value;
    }
    if let Some(value) = options.preserve_formatting {
        policy.preserve_formatting = value;
    }
    if let Some(value) = options.background_query {
        policy.background_query = value;
    }
    policy.missing_items_limit = options.missing_items_limit;
    policy
}

fn build_pivot_layout_from_js(options: JsPivotLayoutOptions) -> Result<PivotLayout> {
    let mut layout = PivotLayout::default();
    if let Some(kind) = options.kind {
        layout.kind = parse_pivot_layout_kind(&kind)?;
    }
    if let Some(value) = options.show_row_grand_totals {
        layout.show_row_grand_totals = value;
    }
    if let Some(value) = options.show_column_grand_totals {
        layout.show_column_grand_totals = value;
    }
    if let Some(value) = options.show_field_headers {
        layout.show_field_headers = value;
    }
    if let Some(value) = options.repeat_item_labels {
        layout.repeat_item_labels = value;
    }
    if let Some(value) = options.show_expand_collapse {
        layout.show_expand_collapse = value;
    }
    if let Some(value) = options.print_drill_indicators {
        layout.print_drill_indicators = value;
    }
    if let Some(value) = options.item_print_titles {
        layout.item_print_titles = value;
    }
    if let Some(value) = options.field_print_titles {
        layout.field_print_titles = value;
    }
    Ok(layout)
}

fn build_pivot_style_from_js(options: JsPivotStyleOptions) -> PivotStyle {
    let mut style = PivotStyle::default();
    if let Some(name) = options.name {
        style.name = if name.is_empty() { None } else { Some(name) };
    }
    if let Some(value) = options.show_row_headers {
        style.show_row_headers = value;
    }
    if let Some(value) = options.show_column_headers {
        style.show_column_headers = value;
    }
    if let Some(value) = options.show_row_stripes {
        style.show_row_stripes = value;
    }
    if let Some(value) = options.show_column_stripes {
        style.show_column_stripes = value;
    }
    if let Some(value) = options.show_last_column {
        style.show_last_column = value;
    }
    style
}

fn parse_pivot_layout_kind(value: &str) -> Result<PivotLayoutKind> {
    Ok(match value {
        "compact" => PivotLayoutKind::Compact,
        "outline" => PivotLayoutKind::Outline,
        "tabular" => PivotLayoutKind::Tabular,
        other => {
            return Err(napi::Error::from_reason(format!(
                "Unsupported pivot layout kind: {other}"
            )));
        }
    })
}

fn parse_pivot_overwrite_policy(value: &str) -> Result<PivotOverwritePolicy> {
    Ok(match value {
        "clearOwnedRange" | "clear_owned_range" | "clear" => PivotOverwritePolicy::ClearOwnedRange,
        "overwrite" => PivotOverwritePolicy::Overwrite,
        "failOnOccupied" | "fail_on_occupied" => PivotOverwritePolicy::FailOnOccupied,
        other => {
            return Err(napi::Error::from_reason(format!(
                "Unsupported pivot overwrite policy: {other}"
            )));
        }
    })
}

fn parse_pivot_sort(value: &str) -> Result<PivotSort> {
    Ok(match value {
        "none" | "manual" => PivotSort::None,
        "ascending" | "asc" => PivotSort::Ascending,
        "descending" | "desc" => PivotSort::Descending,
        other => {
            return Err(napi::Error::from_reason(format!(
                "Unsupported pivot sort: {other}"
            )))
        }
    })
}

fn parse_pivot_subtotal(value: &str) -> Result<PivotSubtotal> {
    Ok(match value {
        "automatic" | "auto" => PivotSubtotal::Automatic,
        "none" => PivotSubtotal::None,
        "sum" => PivotSubtotal::Sum,
        "count" => PivotSubtotal::Count,
        "count_numbers" | "countNumbers" | "countnumbers" | "count_nums" | "countNums"
        | "countnums" => PivotSubtotal::CountNumbers,
        "average" | "avg" => PivotSubtotal::Average,
        "min" => PivotSubtotal::Min,
        "max" => PivotSubtotal::Max,
        "product" => PivotSubtotal::Product,
        "std_dev" | "stdDev" | "stddev" => PivotSubtotal::StdDev,
        "std_dev_p" | "stdDevP" | "stddevp" => PivotSubtotal::StdDevP,
        "var" | "variance" => PivotSubtotal::Var,
        "var_p" | "varP" | "varp" | "variance_p" | "varianceP" => PivotSubtotal::VarP,
        other => {
            return Err(napi::Error::from_reason(format!(
                "Unsupported pivot subtotal: {other}"
            )));
        }
    })
}

fn build_pivot_measure_from_js(options: JsPivotMeasureOptions) -> Result<PivotMeasure> {
    let aggregate = parse_pivot_aggregate(options.aggregate.as_deref())?;
    let mut measure = PivotMeasure::new(options.field, aggregate);
    if let Some(name) = options.name {
        measure = measure.with_name(name);
    }
    if let Some(show_as) = options.show_as {
        measure = measure.with_show_as(parse_pivot_show_as(
            &show_as,
            options.base_field,
            options.base_item,
        )?);
    }
    if let Some(number_format) = options.number_format {
        measure = measure.with_number_format(number_format);
    }
    Ok(measure)
}

fn build_pivot_filter_from_js(options: JsPivotFilterOptions) -> Result<PivotFilter> {
    let kind = options.kind.unwrap_or_else(|| {
        if options.items.is_some() {
            "items".to_string()
        } else {
            "item".to_string()
        }
    });
    match kind.as_str() {
        "item" | "items" | "fieldItems" | "field_items" => {
            let items = options
                .items
                .ok_or_else(|| napi::Error::from_reason("Pivot item filter requires items"))?;
            Ok(PivotFilter::field_items(
                options.field,
                items.into_iter().map(PivotValue::from).collect::<Vec<_>>(),
            ))
        }
        "label" => Ok(PivotFilter::Label {
            field: options.field.into(),
            operator: parse_pivot_filter_operator(options.operator.as_deref())?,
            value: options
                .text
                .ok_or_else(|| napi::Error::from_reason("Pivot label filter requires text"))?,
        }),
        "value" => {
            Ok(PivotFilter::Value {
                field: options.field.into(),
                measure: build_pivot_measure_from_js(options.measure.ok_or_else(|| {
                    napi::Error::from_reason("Pivot value filter requires measure")
                })?)?,
                operator: parse_pivot_filter_operator(options.operator.as_deref())?,
                value: options
                    .value
                    .ok_or_else(|| napi::Error::from_reason("Pivot value filter requires value"))?,
            })
        }
        "topN" | "top_n" | "top" => {
            Ok(PivotFilter::TopN {
                field: options.field.into(),
                measure: build_pivot_measure_from_js(options.measure.ok_or_else(|| {
                    napi::Error::from_reason("Pivot top-N filter requires measure")
                })?)?,
                n: options
                    .n
                    .ok_or_else(|| napi::Error::from_reason("Pivot top-N filter requires n"))?,
                top: options.top.unwrap_or(true),
                percent: options.percent.unwrap_or(false),
            })
        }
        other => Err(napi::Error::from_reason(format!(
            "Unsupported pivot filter kind: {other}"
        ))),
    }
}

fn build_pivot_grouping_from_js(options: JsPivotGroupingOptions) -> Result<PivotGrouping> {
    match options.kind.as_str() {
        "number" | "numeric" => Ok(PivotGrouping::Number {
            field: options.field.into(),
            start: options.start,
            end: options.end,
            interval: options.interval.ok_or_else(|| {
                napi::Error::from_reason("Numeric pivot grouping requires interval")
            })?,
        }),
        "date" => {
            let units = options
                .units
                .ok_or_else(|| napi::Error::from_reason("Date pivot grouping requires units"))?
                .iter()
                .map(|unit| parse_pivot_date_group_unit(unit))
                .collect::<Result<Vec<_>>>()?;
            Ok(PivotGrouping::Date {
                field: options.field.into(),
                units,
            })
        }
        "manual" | "items" | "item" => Ok(PivotGrouping::Manual {
            field: options.field.into(),
            groups: options
                .groups
                .ok_or_else(|| napi::Error::from_reason("Manual pivot grouping requires groups"))?
                .into_iter()
                .map(|group| PivotManualGroup {
                    name: group.name,
                    members: group.members.into_iter().map(pivot_value_from_js).collect(),
                })
                .collect(),
        }),
        other => Err(napi::Error::from_reason(format!(
            "Unsupported pivot grouping kind: {other}"
        ))),
    }
}

fn parse_pivot_aggregate(value: Option<&str>) -> Result<PivotAggregate> {
    let Some(value) = value else {
        return Ok(PivotAggregate::Sum);
    };
    Ok(match value {
        "sum" => PivotAggregate::Sum,
        "count" => PivotAggregate::Count,
        "countNumbers" | "countNums" => PivotAggregate::CountNumbers,
        "average" | "avg" => PivotAggregate::Average,
        "max" => PivotAggregate::Max,
        "min" => PivotAggregate::Min,
        "product" => PivotAggregate::Product,
        "stdDev" => PivotAggregate::StdDev,
        "stdDevP" | "stdDevp" => PivotAggregate::StdDevP,
        "var" => PivotAggregate::Var,
        "varP" | "varp" => PivotAggregate::VarP,
        other => {
            return Err(napi::Error::from_reason(format!(
                "Unsupported pivot aggregate: {other}"
            )));
        }
    })
}

fn parse_pivot_filter_operator(value: Option<&str>) -> Result<PivotFilterOperator> {
    let value = value.ok_or_else(|| napi::Error::from_reason("Pivot filter requires operator"))?;
    Ok(match value {
        "equals" | "equal" | "eq" => PivotFilterOperator::Equals,
        "notEquals" | "notEqual" | "ne" => PivotFilterOperator::NotEquals,
        "lessThan" | "lt" => PivotFilterOperator::LessThan,
        "lessThanOrEqual" | "lte" => PivotFilterOperator::LessThanOrEqual,
        "greaterThan" | "gt" => PivotFilterOperator::GreaterThan,
        "greaterThanOrEqual" | "gte" => PivotFilterOperator::GreaterThanOrEqual,
        "beginsWith" => PivotFilterOperator::BeginsWith,
        "doesNotBeginWith" | "notBeginsWith" => PivotFilterOperator::DoesNotBeginWith,
        "endsWith" => PivotFilterOperator::EndsWith,
        "doesNotEndWith" | "notEndsWith" => PivotFilterOperator::DoesNotEndWith,
        "contains" => PivotFilterOperator::Contains,
        "doesNotContain" | "notContains" => PivotFilterOperator::DoesNotContain,
        other => {
            return Err(napi::Error::from_reason(format!(
                "Unsupported pivot filter operator: {other}"
            )));
        }
    })
}

fn parse_pivot_date_group_unit(value: &str) -> Result<PivotDateGroupUnit> {
    Ok(match value {
        "seconds" => PivotDateGroupUnit::Seconds,
        "minutes" => PivotDateGroupUnit::Minutes,
        "hours" => PivotDateGroupUnit::Hours,
        "days" => PivotDateGroupUnit::Days,
        "months" => PivotDateGroupUnit::Months,
        "quarters" => PivotDateGroupUnit::Quarters,
        "years" => PivotDateGroupUnit::Years,
        other => {
            return Err(napi::Error::from_reason(format!(
                "Unsupported pivot date grouping unit: {other}"
            )));
        }
    })
}

fn parse_pivot_show_as(
    value: &str,
    base_field: Option<String>,
    base_item: Option<Either3<f64, String, bool>>,
) -> Result<PivotShowAs> {
    Ok(match value {
        "normal" => PivotShowAs::Normal,
        "percentOfGrandTotal" | "percentOfTotal" => PivotShowAs::PercentOfGrandTotal,
        "percentOfRowTotal" | "percentOfRow" => PivotShowAs::PercentOfRowTotal,
        "percentOfColumnTotal" | "percentOfCol" => PivotShowAs::PercentOfColumnTotal,
        "index" => PivotShowAs::Index,
        "runningTotal" | "runTotal" => PivotShowAs::RunningTotal {
            base_field: require_pivot_base_field(value, base_field)?.into(),
        },
        "differenceFrom" | "difference" => PivotShowAs::DifferenceFrom {
            base_field: require_pivot_base_field(value, base_field)?.into(),
            base_item: require_pivot_base_item(value, base_item)?,
        },
        "percentDifferenceFrom" | "percentDiff" => PivotShowAs::PercentDifferenceFrom {
            base_field: require_pivot_base_field(value, base_field)?.into(),
            base_item: require_pivot_base_item(value, base_item)?,
        },
        "rankAscending" => PivotShowAs::RankAscending {
            base_field: require_pivot_base_field(value, base_field)?.into(),
        },
        "rankDescending" => PivotShowAs::RankDescending {
            base_field: require_pivot_base_field(value, base_field)?.into(),
        },
        other => {
            return Err(napi::Error::from_reason(format!(
                "Unsupported pivot showAs mode: {other}"
            )));
        }
    })
}

fn require_pivot_base_field(value: &str, base_field: Option<String>) -> Result<String> {
    base_field.ok_or_else(|| {
        napi::Error::from_reason(format!("pivot showAs mode {value} requires baseField"))
    })
}

fn require_pivot_base_item(
    value: &str,
    base_item: Option<Either3<f64, String, bool>>,
) -> Result<PivotValue> {
    base_item.map(pivot_value_from_js).ok_or_else(|| {
        napi::Error::from_reason(format!("pivot showAs mode {value} requires baseItem"))
    })
}

fn pivot_value_from_js(value: Either3<f64, String, bool>) -> PivotValue {
    match value {
        Either3::A(number) => PivotValue::Number(number),
        Either3::B(string) => PivotValue::String(string),
        Either3::C(boolean) => PivotValue::Boolean(boolean),
    }
}

/// A worksheet within a workbook.
///
/// Worksheets contain cells organized in rows and columns. Each cell can
/// contain a value (number, text, boolean) or a formula.
#[napi]
pub struct Worksheet {
    workbook: Arc<RwLock<CoreWorkbook>>,
    sheet_index: usize,
}

#[napi]
impl Worksheet {
    /// Get the worksheet name
    #[napi(getter)]
    pub fn name(&self) -> Result<String> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            wb.worksheet(self.sheet_index)
                .map(|ws| ws.name().to_string())
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))
        })
    }

    /// Set a cell value by address (e.g., "A1", "B2")
    ///
    /// The value can be:
    /// - `null` or `undefined` (clears the cell)
    /// - `number` (numeric value)
    /// - `string` (text)
    /// - `boolean`
    #[napi]
    pub fn set_cell(
        &self,
        address: String,
        value: Option<Either3<f64, String, bool>>,
    ) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let cell_value = if let Some(value) = value {
                match value {
                    Either3::A(n) => CoreCellValue::Number(n),
                    Either3::B(s) => CoreCellValue::string(s),
                    Either3::C(b) => CoreCellValue::Boolean(b),
                }
            } else {
                CoreCellValue::Empty
            };

            let addr = CellAddress::parse(&address)
                .map_err(|e| napi::Error::from_reason(format!("Invalid cell address: {}", e)))?;

            ws.set_cell_value_at(addr.row, addr.col, cell_value)
                .map_err(to_napi_err)
        })
    }

    /// Set a formula in a cell
    ///
    /// @param address - Cell address (e.g., "A1")
    /// @param formula - Formula string (e.g., "=SUM(A1:A10)")
    #[napi]
    pub fn set_formula(&self, address: String, formula: String) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            ws.set_cell_formula(&address, &formula).map_err(to_napi_err)
        })
    }

    /// Set or update a cell style by address.
    ///
    /// A full style object returned by `getCellStyle()` can be used to copy a
    /// style. Partial objects update only the provided top-level components.
    #[napi]
    pub fn set_cell_style(&self, address: String, style: JsStylePatch) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let addr = CellAddress::parse(&address)
                .map_err(|e| napi::Error::from_reason(format!("Invalid cell address: {}", e)))?;

            let mut core_style = ws
                .cell_style_at(addr.row, addr.col)
                .cloned()
                .unwrap_or_default();
            style.apply_to_core_style(&mut core_style)?;
            ws.set_cell_style_at(addr.row, addr.col, &core_style)
                .map_err(to_napi_err)
        })
    }

    /// Set or update a cell style by row/col (0-based).
    #[napi]
    pub fn set_cell_style_at(&self, row: u32, col: u32, style: JsStylePatch) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let mut core_style = ws
                .cell_style_at(row, col as u16)
                .cloned()
                .unwrap_or_default();
            style.apply_to_core_style(&mut core_style)?;
            ws.set_cell_style_at(row, col as u16, &core_style)
                .map_err(to_napi_err)
        })
    }

    /// Set or update the style for all cells in a range (e.g. "A1:C3").
    #[napi]
    pub fn set_range_style(&self, range_str: String, style: JsStylePatch) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let range = CellRange::parse(&range_str)
                .map_err(|e| napi::Error::from_reason(format!("Invalid range: {}", e)))?;
            for addr in range.cells() {
                let mut core_style = ws
                    .cell_style_at(addr.row, addr.col)
                    .cloned()
                    .unwrap_or_default();
                style.apply_to_core_style(&mut core_style)?;
                ws.set_cell_style_at(addr.row, addr.col, &core_style)
                    .map_err(to_napi_err)?;
            }
            Ok(())
        })
    }

    /// Get the raw cell value (not calculated)
    #[napi]
    pub fn get_cell(&self, address: String) -> Result<CellValue> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let addr = CellAddress::parse(&address)
                .map_err(|e| napi::Error::from_reason(format!("Invalid cell address: {}", e)))?;

            let value = ws.get_value_at(addr.row, addr.col);
            Ok(CellValue { inner: value })
        })
    }

    /// Get the raw cell value by row/col (0-based).
    #[napi]
    pub fn get_cell_at(&self, row: u32, col: u32) -> Result<CellValue> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let value = ws.get_value_at(row, col as u16);
            Ok(CellValue { inner: value })
        })
    }

    /// Get the calculated value of a cell
    ///
    /// For formulas, this returns the computed result.
    /// For regular values, returns the value itself.
    #[napi]
    pub fn get_calculated_value(&self, address: String) -> Result<CellValue> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let addr = CellAddress::parse(&address)
                .map_err(|e| napi::Error::from_reason(format!("Invalid cell address: {}", e)))?;

            let value = ws
                .get_calculated_value_at(addr.row, addr.col)
                .cloned()
                .unwrap_or(CoreCellValue::Empty);

            Ok(CellValue { inner: value })
        })
    }

    /// Get the calculated value of a cell by row/col (0-based).
    #[napi]
    pub fn get_calculated_value_at(&self, row: u32, col: u32) -> Result<CellValue> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let value = ws
                .get_calculated_value_at(row, col as u16)
                .cloned()
                .unwrap_or(CoreCellValue::Empty);

            Ok(CellValue { inner: value })
        })
    }

    /// Get the used range as `{ minRow, minCol, maxRow, maxCol }` or null
    /// if the worksheet is empty.
    #[napi(getter)]
    pub fn used_range(&self) -> Result<Option<UsedRange>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            Ok(ws.used_range().map(|r| UsedRange {
                min_row: r.start.row,
                min_col: r.start.col as u32,
                max_row: r.end.row,
                max_col: r.end.col as u32,
            }))
        })
    }

    /// Get IMAGE() metadata for a cell, or null if no image.
    #[napi(js_name = "getImageAt")]
    pub fn get_image_at(&self, row: u32, col: u32) -> Result<Option<JsImageInfo>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            Ok(ws.get_image_at(row, col as u16).map(|info| JsImageInfo {
                source: info.source,
                alt_text: info.alt_text,
                sizing: match info.sizing {
                    ImageSizing::FitCell => 0,
                    ImageSizing::FillCell => 1,
                    ImageSizing::OriginalSize => 2,
                    ImageSizing::Custom => 3,
                },
                width: info.width,
                height: info.height,
            }))
        })
    }

    /// Number of pivot tables on the worksheet.
    #[napi(getter)]
    pub fn pivot_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            u32::try_from(ws.pivot_table_count()).map_err(to_napi_err)
        })
    }

    /// Pivot table names on the worksheet.
    #[napi(getter)]
    pub fn pivot_table_names(&self) -> Result<Vec<String>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .pivot_tables()
                .iter()
                .map(|pivot| pivot.name.clone())
                .collect())
        })
    }

    /// Add a semantic pivot table definition to the worksheet.
    #[napi]
    pub fn add_pivot_table(&self, options: JsPivotTableOptions) -> Result<()> {
        catch_panic(|| {
            let pivot = build_pivot_table_from_js(options)?;
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            ws.add_pivot_table(pivot).map_err(to_napi_err)
        })
    }

    /// Set the height of a row in points
    #[napi]
    pub fn set_row_height(&self, row: u32, height: f64) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            ws.set_row_height(row, height);
            Ok(())
        })
    }

    /// Set the width of a column in character units
    #[napi]
    pub fn set_column_width(&self, col: u32, width: f64) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            ws.set_column_width(col as u16, width);
            Ok(())
        })
    }

    /// Get the row height in points, or null if not explicitly set
    #[napi]
    pub fn get_row_height(&self, row: u32) -> Result<Option<f64>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            Ok(ws.custom_row_heights().get(&row).copied())
        })
    }

    /// Get the column width in character units, or null if not explicitly set
    #[napi]
    pub fn get_column_width(&self, col: u32) -> Result<Option<f64>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            Ok(ws.custom_column_widths().get(&(col as u16)).copied())
        })
    }

    /// Merge cells in a range (e.g., "A1:C3")
    #[napi]
    pub fn merge_cells(&self, range_str: String) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let range = CellRange::parse(&range_str)
                .map_err(|e| napi::Error::from_reason(format!("Invalid range: {}", e)))?;
            ws.merge_cells(&range).map_err(to_napi_err)
        })
    }

    /// Unmerge cells in a range
    #[napi]
    pub fn unmerge_cells(&self, range_str: String) -> Result<bool> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let range = CellRange::parse(&range_str)
                .map_err(|e| napi::Error::from_reason(format!("Invalid range: {}", e)))?;
            Ok(ws.unmerge_cells(&range))
        })
    }
}

/// A workbook containing one or more worksheets.
///
/// This is the main entry point for working with spreadsheet files.
///
/// @example
/// ```typescript
/// const wb = new Workbook();
/// const sheet = wb.getSheet(0);
/// sheet.setCell("A1", 10);
/// sheet.setCell("A2", 20);
/// sheet.setFormula("A3", "=A1+A2");
/// wb.calculate();
/// console.log(sheet.getCalculatedValue("A3").asNumber()); // 30
/// ```
#[napi]
pub struct Workbook {
    inner: Arc<RwLock<CoreWorkbook>>,
}

#[napi]
impl Workbook {
    /// Create a new empty workbook with one worksheet
    #[napi(constructor)]
    pub fn new() -> Self {
        Self {
            inner: Arc::new(RwLock::new(CoreWorkbook::new())),
        }
    }

    /// Open a workbook from a file
    ///
    /// Supported formats:
    /// - `.xlsx` (Excel 2007+)
    /// - `.xls` (Legacy Excel)
    /// - `.csv` (Comma-separated values)
    ///
    /// @param path - Path to the file
    #[napi(factory)]
    pub fn open(path: String) -> Result<Self> {
        catch_panic(|| {
            let path = PathBuf::from(path);
            let wb = CoreWorkbook::open(&path)
                .map_err(|e| napi::Error::from_reason(format!("Failed to open file: {}", e)))?;

            Ok(Self {
                inner: Arc::new(RwLock::new(wb)),
            })
        })
    }

    /// Load a workbook from bytes (Buffer/Uint8Array), auto-detecting the format.
    ///
    /// Supports XLSX and XLS formats. The format is detected from magic bytes.
    ///
    /// @param data - The file content as a Buffer
    #[napi(factory)]
    pub fn from_bytes(data: Buffer) -> Result<Self> {
        catch_panic(|| {
            use duke_sheets::WorkbookExt;
            let wb = duke_sheets_core::Workbook::from_bytes(data.as_ref())
                .map_err(|e| napi::Error::from_reason(format!("Failed to read file: {}", e)))?;

            Ok(Self {
                inner: Arc::new(RwLock::new(wb)),
            })
        })
    }

    /// Load a workbook from a CSV string
    ///
    /// @param csv - The CSV content as a string
    #[napi(factory)]
    pub fn from_csv_string(csv: String) -> Result<Self> {
        catch_panic(|| {
            let reader = Cursor::new(csv.into_bytes());
            let ws = duke_sheets_csv::CsvReader::read(
                reader,
                &duke_sheets_csv::CsvReadOptions::default(),
            )
            .map_err(|e| napi::Error::from_reason(format!("Failed to read CSV: {}", e)))?;

            let mut wb = CoreWorkbook::empty();
            wb.add_existing_worksheet(ws).map_err(to_napi_err)?;

            Ok(Self {
                inner: Arc::new(RwLock::new(wb)),
            })
        })
    }

    /// Save the workbook to a file
    ///
    /// The format is determined by the file extension:
    /// - `.xlsx` for Excel format
    /// - `.xls` for legacy Excel binary format
    /// - `.csv` for CSV format (first sheet only)
    ///
    /// @param path - Path to save to
    #[napi]
    pub fn save(&self, path: String) -> Result<()> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            let path = PathBuf::from(path);
            wb.save(&path)
                .map_err(|e| napi::Error::from_reason(format!("Failed to save: {}", e)))
        })
    }

    /// Save the workbook to a password-protected file. The encryption
    /// variant is selected via `profile`:
    ///
    /// - `"default"` (or null) - Agile-256 for .xlsx, RC4 CryptoAPI 128 for .xls
    /// - `"agile"` - OOXML Agile (AES-CBC + HMAC-SHA*); pass keyBits to override
    /// - `"standard"` - OOXML Standard Encryption (AES-ECB)
    /// - `"rc4-cryptoapi"` - XLS RC4 CryptoAPI; keyBits 40 or 128
    /// - `"rc4-legacy"` - XLS legacy RC4 (MD5 KDF)
    /// - `"xor"` - XLS XOR Obfuscation; round-trips via duke-sheets but
    ///   does not interoperate with modern Excel
    ///
    /// @param path - Path to save to
    /// @param password - Password to encrypt with
    /// @param profile - Optional encryption variant (see above)
    /// @param keyBits - Optional key size override (Agile / RC4 CryptoAPI)
    /// @param spinCount - Optional iteration count (Agile only; default 100,000)
    #[napi]
    pub fn save_with_password(
        &self,
        path: String,
        password: String,
        profile: Option<String>,
        key_bits: Option<u32>,
        spin_count: Option<u32>,
    ) -> Result<()> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            let encryption = parse_encryption_profile(profile.as_deref(), key_bits, spin_count)
                .map_err(napi::Error::from_reason)?;
            let opts = duke_sheets::WorkbookSaveOptions::default()
                .password(&password)
                .encryption(encryption);
            wb.save_with(&PathBuf::from(path), &opts)
                .map_err(|e| napi::Error::from_reason(format!("Failed to save: {}", e)))
        })
    }

    /// Open a password-protected workbook.
    ///
    /// @param path - File path
    /// @param password - Password to attempt
    /// @param skipIntegrityCheck - If true, skip the HMAC integrity
    ///   check on Agile-encrypted files (matches Office behaviour).
    ///   Default false.
    #[napi(factory)]
    pub fn open_with_password(
        path: String,
        password: String,
        skip_integrity_check: Option<bool>,
    ) -> Result<Self> {
        catch_panic(|| {
            let mut opts = duke_sheets::WorkbookOpenOptions::default().password(&password);
            if skip_integrity_check.unwrap_or(false) {
                opts = opts.skip_integrity_check();
            }
            let wb = CoreWorkbook::open_with(&PathBuf::from(path), &opts)
                .map_err(|e| napi::Error::from_reason(format!("Failed to open file: {}", e)))?;
            Ok(Self {
                inner: Arc::new(RwLock::new(wb)),
            })
        })
    }

    /// Save the workbook as a CSV string (first sheet only)
    #[napi]
    pub fn save_csv_string(&self) -> Result<String> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(0)
                .ok_or_else(|| napi::Error::from_reason("No worksheets to save"))?;

            let mut buf = Vec::new();
            duke_sheets_csv::CsvWriter::write(
                ws,
                &mut buf,
                &duke_sheets_csv::CsvWriteOptions::default(),
            )
            .map_err(|e| napi::Error::from_reason(format!("Failed to write CSV: {}", e)))?;

            String::from_utf8(buf)
                .map_err(|e| napi::Error::from_reason(format!("Invalid UTF-8: {}", e)))
        })
    }

    /// Get the number of worksheets
    #[napi(getter)]
    pub fn sheet_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            let count = u32::try_from(wb.sheet_count()).map_err(to_napi_err)?;
            Ok(count)
        })
    }

    /// Get a list of all worksheet names
    #[napi(getter)]
    pub fn sheet_names(&self) -> Result<Vec<String>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok((0..wb.sheet_count())
                .filter_map(|i| wb.worksheet(i).map(|ws| ws.name().to_string()))
                .collect())
        })
    }

    /// Get a worksheet by index (number) or name (string)
    ///
    /// @param indexOrName - Either a zero-based index or a sheet name
    /// @throws Error if index out of range or name not found
    #[napi]
    pub fn get_sheet(&self, index_or_name: Either<u32, String>) -> Result<Worksheet> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;

            let sheet_index = match index_or_name {
                Either::A(idx) => {
                    let idx = idx as usize;
                    if idx >= wb.sheet_count() {
                        return Err(napi::Error::from_reason(format!(
                            "Sheet index {} out of range (0..{})",
                            idx,
                            wb.sheet_count()
                        )));
                    }
                    idx
                }
                Either::B(name) => wb.sheet_index(&name).ok_or_else(|| {
                    napi::Error::from_reason(format!("Sheet '{}' not found", name))
                })?,
            };

            drop(wb);

            Ok(Worksheet {
                workbook: Arc::clone(&self.inner),
                sheet_index,
            })
        })
    }

    /// Add a new worksheet with the given name
    ///
    /// @param name - Name for the new worksheet
    /// @returns Index of the new worksheet
    #[napi]
    pub fn add_sheet(&self, name: String) -> Result<u32> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            wb.add_worksheet_with_name(&name)
                .map(|idx| idx as u32)
                .map_err(to_napi_err)
        })
    }

    /// Remove a worksheet by index
    ///
    /// @param index - Zero-based index of the worksheet to remove
    #[napi]
    pub fn remove_sheet(&self, index: u32) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            wb.remove_worksheet(index as usize)
                .map(|_| ())
                .map_err(to_napi_err)
        })
    }

    /// Calculate all formulas in the workbook.
    ///
    /// Optionally accepts calculation options. Callbacks (`webServiceFn`, `rtdFn`, `externalFn`)
    /// are only supported on the async path via `calculateAsync`.
    ///
    /// @param options - Optional calculation options
    /// @returns Statistics about the calculation
    #[napi]
    pub fn calculate(&self, options: Option<JsCalculationOptions>) -> Result<CalculationStats> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            let stats = if let Some(opts) = options {
                wb.calculate_with_options(&opts.into_core())
                    .map_err(to_napi_err)?
            } else {
                wb.calculate().map_err(to_napi_err)?
            };
            Ok(CalculationStats { inner: stats })
        })
    }

    /// Refresh all pivot tables in the workbook.
    #[napi]
    pub fn refresh_pivots(&self) -> Result<JsPivotRefreshStats> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            wb.refresh_pivots()
                .map_err(to_napi_err)
                .and_then(JsPivotRefreshStats::try_from)
        })
    }

    /// Define a named range
    ///
    /// @param name - Name for the range (e.g., "TaxRate")
    /// @param refersTo - What the name refers to (e.g., "Sheet1!$A$1" or "0.05")
    #[napi]
    pub fn define_name(&self, name: String, refers_to: String) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            wb.define_name(&name, &refers_to).map_err(to_napi_err)
        })
    }

    /// Get a named range definition
    ///
    /// @param name - Name to look up
    /// @returns The refers_to string, or null if not found
    #[napi]
    pub fn get_named_range(&self, name: String) -> Result<Option<String>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.get_named_range(&name, 0).map(|nr| nr.refers_to.clone()))
        })
    }
}

pub struct OpenTask {
    path: PathBuf,
}

impl Task for OpenTask {
    type Output = CoreWorkbook;
    type JsValue = Workbook;

    fn compute(&mut self) -> Result<Self::Output> {
        catch_panic(|| {
            CoreWorkbook::open(&self.path)
                .map_err(|e| napi::Error::from_reason(format!("Failed to open file: {}", e)))
        })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook {
            inner: Arc::new(RwLock::new(output)),
        })
    }
}

pub struct OpenBytesTask {
    data: Vec<u8>,
}

impl Task for OpenBytesTask {
    type Output = CoreWorkbook;
    type JsValue = Workbook;

    fn compute(&mut self) -> Result<Self::Output> {
        catch_panic(|| {
            use duke_sheets::WorkbookExt;
            CoreWorkbook::from_bytes(&self.data)
                .map_err(|e| napi::Error::from_reason(format!("Failed to read file: {}", e)))
        })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook {
            inner: Arc::new(RwLock::new(output)),
        })
    }
}

pub struct SaveTask {
    workbook: Arc<RwLock<CoreWorkbook>>,
    path: PathBuf,
}

impl Task for SaveTask {
    type Output = ();
    type JsValue = ();

    fn compute(&mut self) -> Result<Self::Output> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            wb.save(&self.path)
                .map_err(|e| napi::Error::from_reason(format!("Failed to save: {}", e)))
        })
    }

    fn resolve(&mut self, _env: Env, _output: Self::Output) -> Result<Self::JsValue> {
        Ok(())
    }
}

pub struct CalculateTask {
    workbook: Arc<RwLock<CoreWorkbook>>,
    options: Option<CalculationOptions>,
}

impl Task for CalculateTask {
    type Output = CoreCalculationStats;
    type JsValue = CalculationStats;

    fn compute(&mut self) -> Result<Self::Output> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            if let Some(opts) = &self.options {
                wb.calculate_with_options(opts).map_err(to_napi_err)
            } else {
                wb.calculate().map_err(to_napi_err)
            }
        })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(CalculationStats { inner: output })
    }
}

/// Open a workbook from a file asynchronously (non-blocking).
///
/// Runs file I/O and parsing on the libuv thread pool so the
/// Node.js event loop is not blocked.
///
/// @param path - Path to the file
/// @returns Promise<Workbook>
#[napi(ts_return_type = "Promise<Workbook>")]
pub fn open_async(path: String) -> AsyncTask<OpenTask> {
    AsyncTask::new(OpenTask {
        path: PathBuf::from(path),
    })
}

/// Load a workbook from bytes asynchronously (non-blocking).
///
/// Auto-detects format (XLSX or XLS) from magic bytes.
///
/// @param data - The file content as a Buffer
/// @returns Promise<Workbook>
#[napi(ts_return_type = "Promise<Workbook>")]
pub fn from_bytes_async(data: Buffer) -> AsyncTask<OpenBytesTask> {
    AsyncTask::new(OpenBytesTask {
        data: data.to_vec(),
    })
}

#[napi]
impl Workbook {
    /// Save the workbook to a file asynchronously (non-blocking).
    ///
    /// @param path - Path to save to
    /// @returns Promise<void>
    #[napi(ts_return_type = "Promise<void>")]
    pub fn save_async(&self, path: String) -> AsyncTask<SaveTask> {
        AsyncTask::new(SaveTask {
            workbook: Arc::clone(&self.inner),
            path: PathBuf::from(path),
        })
    }

    /// Calculate all formulas asynchronously (non-blocking).
    ///
    /// Optionally accepts calculation options with callback functions:
    /// - `webServiceFn`: called for each `WEBSERVICE(url)` evaluation
    /// - `rtdFn`: called for each `RTD(progId, server, ...topics)` evaluation
    ///
    /// Callbacks must return a Promise. Use `async (url) => ...` or
    /// wrap synchronous results: `async (url) => mySyncFn(url)`.
    ///
    /// @param options - Optional calculation options with optional callbacks
    /// @returns Promise<CalculationStats>
    #[napi(
        ts_args_type = "options?: JsCalculationOptions & { webServiceFn?: (url: string) => Promise<string | null | undefined>; rtdFn?: (progId: string, server: string, topics: string[]) => Promise<string | null | undefined>; externalFn?: (book: string, name: string, args: string[]) => Promise<string | null | undefined>; externalFnFn?: (book: string, name: string, args: string[]) => Promise<string | null | undefined> }",
        ts_return_type = "Promise<CalculationStats>"
    )]
    pub fn calculate_async<'env>(
        &self,
        options: Option<Object<'env>>,
    ) -> Result<AsyncTask<CalculateTask>> {
        let opts = if let Some(options) = options {
            // Extract basic options from the JS object
            let iterative: Option<bool> = try_get_property(&options, "iterative")?;
            let max_iterations: Option<u32> = try_get_property(&options, "maxIterations")?;
            let max_change: Option<f64> = try_get_property(&options, "maxChange")?;
            let force_full_calculation: Option<bool> =
                try_get_property(&options, "forceFullCalculation")?;
            let calculate_volatile: Option<bool> = try_get_property(&options, "calculateVolatile")?;
            let sheets: Option<Vec<u32>> = try_get_property(&options, "sheets")?;
            let max_threads: Option<u32> = try_get_property(&options, "maxThreads")?;

            // Extract callback functions and build ThreadsafeFunctions
            let web_service_js_fn: Option<Function<'env, String, Promise<Option<String>>>> =
                try_get_property(&options, "webServiceFn")?;
            let rtd_js_fn: Option<
                Function<'env, (String, String, Vec<String>), Promise<Option<String>>>,
            > = try_get_property(&options, "rtdFn")?;
            let external_js_fn: Option<
                Function<'env, (String, String, Vec<String>), Promise<Option<String>>>,
            > = match try_get_property(&options, "externalFn")? {
                Some(f) => Some(f),
                None => try_get_property(&options, "externalFnFn")?,
            };

            let web_service_fn: Option<Arc<dyn Fn(&str) -> Option<String> + Send + Sync>> =
                if let Some(js_fn) = web_service_js_fn {
                    let tsfn = js_fn.build_threadsafe_function::<String>().build()?;
                    let tsfn = Arc::new(tsfn);
                    Some(Arc::new(move |url: &str| -> Option<String> {
                        let url = url.to_string();
                        let tsfn = Arc::clone(&tsfn);
                        napi::bindgen_prelude::block_on(async move {
                            let promise = tsfn.call_async(url).await.ok()?;
                            promise.await.ok().flatten()
                        })
                    }))
                } else {
                    None
                };

            let external_fn: Option<
                Arc<dyn Fn(&str, &str, &[String]) -> Option<FormulaValue> + Send + Sync>,
            > = if let Some(js_fn) = external_js_fn {
                let tsfn = js_fn
                    .build_threadsafe_function::<(String, String, Vec<String>)>()
                    .build_callback(|ctx| Ok(FnArgs { data: ctx.value }))?;
                let tsfn = Arc::new(tsfn);
                Some(Arc::new(
                    move |book: &str, name: &str, args: &[String]| -> Option<FormulaValue> {
                        let args = (book.to_string(), name.to_string(), args.to_vec());
                        let tsfn = Arc::clone(&tsfn);
                        napi::bindgen_prelude::block_on(async move {
                            let promise = tsfn.call_async(args).await.ok()?;
                            promise.await.ok().flatten().map(FormulaValue::String)
                        })
                    },
                ))
            } else {
                None
            };

            let rtd_fn: Option<Arc<dyn Fn(&str, &str, &[String]) -> Option<String> + Send + Sync>> =
                if let Some(js_fn) = rtd_js_fn {
                    let tsfn = js_fn
                        .build_threadsafe_function::<(String, String, Vec<String>)>()
                        .build_callback(|ctx| Ok(FnArgs { data: ctx.value }))?;
                    let tsfn = Arc::new(tsfn);
                    Some(Arc::new(
                        move |prog_id: &str, server: &str, topics: &[String]| -> Option<String> {
                            let args = (prog_id.to_string(), server.to_string(), topics.to_vec());
                            let tsfn = Arc::clone(&tsfn);
                            napi::bindgen_prelude::block_on(async move {
                                let promise = tsfn.call_async(args).await.ok()?;
                                promise.await.ok().flatten()
                            })
                        },
                    ))
                } else {
                    None
                };

            Some(CalculationOptions {
                iterative: iterative.unwrap_or(false),
                max_iterations: max_iterations.unwrap_or(100),
                max_change: max_change.unwrap_or(0.001),
                force_full_calculation: force_full_calculation.unwrap_or(true),
                calculate_volatile: calculate_volatile.unwrap_or(true),
                sheets: sheets
                    .unwrap_or_default()
                    .into_iter()
                    .map(|i| i as usize)
                    .collect(),
                max_threads: max_threads.map(|n| n as usize),
                web_service_fn,
                rtd_fn,
                external_fn,
            })
        } else {
            None
        };

        Ok(AsyncTask::new(CalculateTask {
            workbook: Arc::clone(&self.inner),
            options: opts,
        }))
    }
}
