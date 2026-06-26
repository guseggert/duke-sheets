//! Python bindings for duke-sheets
//!
//! This module provides PyO3-based Python bindings for the duke-sheets library,
//! allowing Python code to read, write, and manipulate Excel files.

use pyo3::exceptions::{PyIOError, PyIndexError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use std::path::PathBuf;
use std::sync::{Arc, RwLock};

use duke_sheets::prelude::*;
use duke_sheets::{
    CalculationOptions, FormulaValue, ImageSizing, WorkbookCalculationExt, WorkbookPivotExt,
};
use duke_sheets_core::{
    CellError, CellValue as CoreCellValue, PivotAggregate, PivotDateGroupUnit, PivotField,
    PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayout, PivotLayoutKind,
    PivotManualGroup, PivotMeasure, PivotOverwritePolicy, PivotRefreshPolicy, PivotShowAs,
    PivotSort, PivotSource, PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable, PivotValue,
    WorkbookConnection, WorkbookConnectionKind,
};

mod types;
pub use types::*;
mod workbook_read;
mod worksheet_read;
pub use worksheet_read::PyRowIterator;

// Error Conversion

pub(crate) fn to_py_err(e: impl std::fmt::Display) -> PyErr {
    PyRuntimeError::new_err(e.to_string())
}

fn build_pivot_table_from_py(options: &Bound<'_, PyAny>) -> PyResult<PivotTable> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("pivot options must be a dict"))?;
    let mut builder = PivotTable::builder(required_string(dict, &["name"])?);

    let table_name = optional_string(dict, &["table_name", "tableName"])?;
    let source_range = optional_string(dict, &["source_range", "sourceRange"])?;
    let external_connection_name = optional_string(
        dict,
        &["external_connection_name", "externalConnectionName"],
    )?;
    let external_command_text =
        optional_string(dict, &["external_command_text", "externalCommandText"])?;
    let consolidation_ranges_value =
        optional_any(dict, &["consolidation_ranges", "consolidationRanges"])?;
    if external_command_text.is_some() && external_connection_name.is_none() {
        return Err(PyValueError::new_err(
            "Pivot options require external_connection_name/externalConnectionName when external_command_text/externalCommandText is set",
        ));
    }
    let source_count = usize::from(table_name.is_some())
        + usize::from(source_range.is_some())
        + usize::from(external_connection_name.is_some())
        + usize::from(consolidation_ranges_value.is_some());
    if source_count != 1 {
        return Err(PyValueError::new_err(
            "Pivot options require exactly one of table_name/tableName, source_range/sourceRange, external_connection_name/externalConnectionName, or consolidation_ranges/consolidationRanges",
        ));
    }

    match (
        table_name,
        source_range,
        external_connection_name,
        consolidation_ranges_value,
    ) {
        (Some(table_name), None, None, None) => {
            builder = builder.table_source(table_name);
        }
        (None, Some(source_range), None, None) => {
            let range = CellRange::parse(&source_range)
                .map_err(|e| PyValueError::new_err(format!("Invalid pivot source range: {e}")))?;
            builder = if let Some(sheet) = optional_string(dict, &["source_sheet", "sourceSheet"])?
            {
                builder.source_range_on_sheet(sheet, range)
            } else {
                builder.source_range(range)
            };
        }
        (None, None, Some(connection_name), None) => {
            builder = builder.source(PivotSource::External {
                connection_name,
                command_text: external_command_text,
            });
        }
        (None, None, None, Some(consolidation_ranges)) => {
            builder = builder.source(PivotSource::Consolidation {
                ranges: build_pivot_consolidation_ranges_from_py(&consolidation_ranges)?,
            });
        }
        _ => unreachable!("source_count validation accepts exactly one source"),
    }

    builder = builder
        .target_address(&required_string(dict, &["target"])?)
        .map_err(|e| PyValueError::new_err(format!("Invalid pivot target: {e}")))?;
    for field in optional_string_vec(dict, &["rows"])?.unwrap_or_default() {
        builder = builder.row(field);
    }
    for field in optional_string_vec(dict, &["columns"])?.unwrap_or_default() {
        builder = builder.column(field);
    }
    for field in optional_string_vec(dict, &["pages"])?.unwrap_or_default() {
        builder = builder.page(field);
    }
    if let Some(row_fields) = optional_any(dict, &["row_fields", "rowFields"])? {
        let row_fields = row_fields
            .downcast::<PyList>()
            .map_err(|_| PyValueError::new_err("pivot row_fields must be a list"))?;
        for field in row_fields.iter() {
            builder = builder.row(build_pivot_field_from_py(&field)?);
        }
    }
    if let Some(column_fields) = optional_any(dict, &["column_fields", "columnFields"])? {
        let column_fields = column_fields
            .downcast::<PyList>()
            .map_err(|_| PyValueError::new_err("pivot column_fields must be a list"))?;
        for field in column_fields.iter() {
            builder = builder.column(build_pivot_field_from_py(&field)?);
        }
    }
    if let Some(page_fields) = optional_any(dict, &["page_fields", "pageFields"])? {
        let page_fields = page_fields
            .downcast::<PyList>()
            .map_err(|_| PyValueError::new_err("pivot page_fields must be a list"))?;
        for field in page_fields.iter() {
            builder = builder.page(build_pivot_field_from_py(&field)?);
        }
    }

    let measures_value = dict
        .get_item("measures")?
        .ok_or_else(|| PyValueError::new_err("pivot options require measures"))?;
    let measures = measures_value
        .downcast::<PyList>()
        .map_err(|_| PyValueError::new_err("pivot measures must be a list"))?;
    for measure in measures.iter() {
        builder = builder.pivot_measure(build_pivot_measure_from_py(&measure)?);
    }

    if let Some(filters_value) = optional_any(dict, &["filters"])? {
        let filters = filters_value
            .downcast::<PyList>()
            .map_err(|_| PyValueError::new_err("pivot filters must be a list"))?;
        for filter in filters.iter() {
            builder = builder.filter(build_pivot_filter_from_py(&filter)?);
        }
    }
    if let Some(calculated_fields_value) =
        optional_any(dict, &["calculated_fields", "calculatedFields"])?
    {
        let calculated_fields = calculated_fields_value
            .downcast::<PyList>()
            .map_err(|_| PyValueError::new_err("pivot calculated_fields must be a list"))?;
        for calculated_field in calculated_fields.iter() {
            let calculated_field = calculated_field
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("pivot calculated field must be a dict"))?;
            builder = builder.calculated_field(
                required_string(calculated_field, &["name"])?,
                required_string(calculated_field, &["formula"])?,
            );
        }
    }
    if let Some(groupings_value) = optional_any(dict, &["groupings"])? {
        let groupings = groupings_value
            .downcast::<PyList>()
            .map_err(|_| PyValueError::new_err("pivot groupings must be a list"))?;
        for grouping in groupings.iter() {
            builder = builder.grouping(build_pivot_grouping_from_py(&grouping)?);
        }
    }
    if let Some(refresh_policy) = optional_any(dict, &["refresh_policy", "refreshPolicy"])? {
        builder = builder.refresh_policy(build_pivot_refresh_policy_from_py(&refresh_policy)?);
    }
    if let Some(layout) = optional_any(dict, &["layout"])? {
        builder = builder.layout(build_pivot_layout_from_py(&layout)?);
    }
    if let Some(style) = optional_any(dict, &["style"])? {
        builder = builder.style(build_pivot_style_from_py(&style)?);
    }
    if let Some(overwrite_policy) = optional_string(dict, &["overwrite_policy", "overwritePolicy"])?
    {
        builder = builder.overwrite_policy(parse_pivot_overwrite_policy(&overwrite_policy)?);
    }

    builder.build().map_err(to_py_err)
}

fn build_pivot_consolidation_ranges_from_py(
    ranges_value: &Bound<'_, PyAny>,
) -> PyResult<Vec<PivotSourceRange>> {
    let ranges = ranges_value
        .downcast::<PyList>()
        .map_err(|_| PyValueError::new_err("pivot consolidation_ranges must be a list"))?;
    if ranges.is_empty() {
        return Err(PyValueError::new_err(
            "pivot consolidation_ranges must contain at least one range",
        ));
    }

    ranges
        .iter()
        .map(|range| {
            let dict = range
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("pivot consolidation range must be a dict"))?;
            let range_ref = required_string(dict, &["range"])?;
            let parsed = CellRange::parse(&range_ref).map_err(|e| {
                PyValueError::new_err(format!("Invalid pivot consolidation range: {e}"))
            })?;
            let mut source_range =
                PivotSourceRange::new(required_string(dict, &["sheet"])?, parsed);
            if let Some(name) = optional_string(dict, &["name"])? {
                source_range = source_range.with_name(name);
            }
            if let Some(page_items) = optional_string_vec(dict, &["page_items", "pageItems"])? {
                source_range = source_range.with_page_items(page_items);
            }
            Ok(source_range)
        })
        .collect()
}

fn build_pivot_field_from_py(options: &Bound<'_, PyAny>) -> PyResult<PivotField> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("pivot field options must be a dict"))?;
    let mut field = PivotField::new(required_string(dict, &["field"])?);
    if let Some(sort) = optional_string(dict, &["sort"])? {
        field.sort = parse_pivot_sort(&sort)?;
    }
    if let Some(subtotal) = optional_string(dict, &["subtotal"])? {
        field.subtotal = parse_pivot_subtotal(&subtotal)?;
    }
    if let Some(show_empty_items) = optional_bool(dict, &["show_empty_items", "showEmptyItems"])? {
        field.show_empty_items = show_empty_items;
    }
    Ok(field)
}

fn build_pivot_refresh_policy_from_py(options: &Bound<'_, PyAny>) -> PyResult<PivotRefreshPolicy> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("pivot refresh_policy must be a dict"))?;
    let mut policy = PivotRefreshPolicy::default();
    if let Some(value) = optional_bool(dict, &["refresh_on_open", "refreshOnOpen"])? {
        policy.refresh_on_open = value;
    }
    if let Some(value) = optional_bool(dict, &["preserve_formatting", "preserveFormatting"])? {
        policy.preserve_formatting = value;
    }
    if let Some(value) = optional_bool(dict, &["background_query", "backgroundQuery"])? {
        policy.background_query = value;
    }
    policy.missing_items_limit = optional_u32(dict, &["missing_items_limit", "missingItemsLimit"])?;
    Ok(policy)
}

fn build_pivot_layout_from_py(options: &Bound<'_, PyAny>) -> PyResult<PivotLayout> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("pivot layout must be a dict"))?;
    let mut layout = PivotLayout::default();
    if let Some(kind) = optional_string(dict, &["kind"])? {
        layout.kind = parse_pivot_layout_kind(&kind)?;
    }
    if let Some(value) = optional_bool(dict, &["show_row_grand_totals", "showRowGrandTotals"])? {
        layout.show_row_grand_totals = value;
    }
    if let Some(value) =
        optional_bool(dict, &["show_column_grand_totals", "showColumnGrandTotals"])?
    {
        layout.show_column_grand_totals = value;
    }
    if let Some(value) = optional_bool(dict, &["show_field_headers", "showFieldHeaders"])? {
        layout.show_field_headers = value;
    }
    if let Some(value) = optional_bool(dict, &["repeat_item_labels", "repeatItemLabels"])? {
        layout.repeat_item_labels = value;
    }
    if let Some(value) = optional_bool(dict, &["show_expand_collapse", "showExpandCollapse"])? {
        layout.show_expand_collapse = value;
    }
    if let Some(value) = optional_bool(dict, &["print_drill_indicators", "printDrillIndicators"])? {
        layout.print_drill_indicators = value;
    }
    if let Some(value) = optional_bool(dict, &["item_print_titles", "itemPrintTitles"])? {
        layout.item_print_titles = value;
    }
    if let Some(value) = optional_bool(dict, &["field_print_titles", "fieldPrintTitles"])? {
        layout.field_print_titles = value;
    }
    Ok(layout)
}

fn build_pivot_style_from_py(options: &Bound<'_, PyAny>) -> PyResult<PivotStyle> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("pivot style must be a dict"))?;
    let mut style = PivotStyle::default();
    if let Some(name) = optional_string(dict, &["name"])? {
        style.name = if name.is_empty() { None } else { Some(name) };
    }
    if let Some(value) = optional_bool(dict, &["show_row_headers", "showRowHeaders"])? {
        style.show_row_headers = value;
    }
    if let Some(value) = optional_bool(dict, &["show_column_headers", "showColumnHeaders"])? {
        style.show_column_headers = value;
    }
    if let Some(value) = optional_bool(dict, &["show_row_stripes", "showRowStripes"])? {
        style.show_row_stripes = value;
    }
    if let Some(value) = optional_bool(dict, &["show_column_stripes", "showColumnStripes"])? {
        style.show_column_stripes = value;
    }
    if let Some(value) = optional_bool(dict, &["show_last_column", "showLastColumn"])? {
        style.show_last_column = value;
    }
    Ok(style)
}

fn build_pivot_measure_from_py(options: &Bound<'_, PyAny>) -> PyResult<PivotMeasure> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("pivot measure must be a dict"))?;
    let aggregate = parse_pivot_aggregate(optional_string(dict, &["aggregate"])?.as_deref())?;
    let mut measure = PivotMeasure::new(required_string(dict, &["field"])?, aggregate);
    if let Some(name) = optional_string(dict, &["name"])? {
        measure = measure.with_name(name);
    }
    if let Some(show_as) = optional_string(dict, &["show_as", "showAs"])? {
        measure = measure.with_show_as(parse_pivot_show_as(
            &show_as,
            optional_string(dict, &["base_field", "baseField"])?,
            optional_pivot_value(dict, &["base_item", "baseItem"])?,
        )?);
    }
    if let Some(number_format) = optional_string(dict, &["number_format", "numberFormat"])? {
        measure = measure.with_number_format(number_format);
    }
    Ok(measure)
}

fn build_pivot_filter_from_py(options: &Bound<'_, PyAny>) -> PyResult<PivotFilter> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("pivot filter must be a dict"))?;
    let has_items = optional_any(dict, &["items"])?.is_some();
    let kind = optional_string(dict, &["kind"])?.unwrap_or_else(|| {
        if has_items {
            "items".to_string()
        } else {
            "item".to_string()
        }
    });
    let field = required_string(dict, &["field"])?;
    match kind.as_str() {
        "item" | "items" | "fieldItems" | "field_items" => {
            let items = required_string_vec(dict, &["items"])?;
            Ok(PivotFilter::field_items(
                field,
                items.into_iter().map(PivotValue::from).collect::<Vec<_>>(),
            ))
        }
        "label" => Ok(PivotFilter::Label {
            field: field.into(),
            operator: parse_pivot_filter_operator(
                optional_string(dict, &["operator"])?.as_deref(),
            )?,
            value: required_string(dict, &["text", "value"])?,
        }),
        "value" => {
            let measure = optional_any(dict, &["measure"])?
                .ok_or_else(|| PyValueError::new_err("pivot value filter requires measure"))?;
            Ok(PivotFilter::Value {
                field: field.into(),
                measure: build_pivot_measure_from_py(&measure)?,
                operator: parse_pivot_filter_operator(
                    optional_string(dict, &["operator"])?.as_deref(),
                )?,
                value: optional_f64(dict, &["value"])?
                    .ok_or_else(|| PyValueError::new_err("pivot value filter requires value"))?,
            })
        }
        "topN" | "top_n" | "top" => {
            let measure = optional_any(dict, &["measure"])?
                .ok_or_else(|| PyValueError::new_err("pivot top-N filter requires measure"))?;
            Ok(PivotFilter::TopN {
                field: field.into(),
                measure: build_pivot_measure_from_py(&measure)?,
                n: optional_u32(dict, &["n"])?
                    .ok_or_else(|| PyValueError::new_err("pivot top-N filter requires n"))?,
                top: optional_bool(dict, &["top"])?.unwrap_or(true),
                percent: optional_bool(dict, &["percent"])?.unwrap_or(false),
            })
        }
        other => Err(PyValueError::new_err(format!(
            "Unsupported pivot filter kind: {other}"
        ))),
    }
}

fn build_pivot_grouping_from_py(options: &Bound<'_, PyAny>) -> PyResult<PivotGrouping> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("pivot grouping must be a dict"))?;
    let field = required_string(dict, &["field"])?;
    match required_string(dict, &["kind"])?.as_str() {
        "number" | "numeric" => Ok(PivotGrouping::Number {
            field: field.into(),
            start: optional_f64(dict, &["start"])?,
            end: optional_f64(dict, &["end"])?,
            interval: optional_f64(dict, &["interval"])?
                .ok_or_else(|| PyValueError::new_err("numeric pivot grouping requires interval"))?,
        }),
        "date" => Ok(PivotGrouping::Date {
            field: field.into(),
            units: required_string_vec(dict, &["units"])?
                .iter()
                .map(|unit| parse_pivot_date_group_unit(unit))
                .collect::<PyResult<Vec<_>>>()?,
        }),
        "manual" | "items" | "item" => Ok(PivotGrouping::Manual {
            field: field.into(),
            groups: required_manual_groups(dict)?,
        }),
        other => Err(PyValueError::new_err(format!(
            "Unsupported pivot grouping kind: {other}"
        ))),
    }
}

fn required_manual_groups(dict: &Bound<'_, PyDict>) -> PyResult<Vec<PivotManualGroup>> {
    let groups_value = optional_any(dict, &["groups"])?
        .ok_or_else(|| PyValueError::new_err("manual pivot grouping requires groups"))?;
    let groups = groups_value
        .downcast::<PyList>()
        .map_err(|_| PyValueError::new_err("manual pivot grouping groups must be a list"))?;

    groups
        .iter()
        .map(|group| {
            let group_dict = group
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("manual pivot group must be a dict"))?;
            let members_value = optional_any(group_dict, &["members"])?
                .ok_or_else(|| PyValueError::new_err("manual pivot group requires members"))?;
            let members = members_value
                .downcast::<PyList>()
                .map_err(|_| PyValueError::new_err("manual pivot group members must be a list"))?
                .iter()
                .map(|member| pivot_value_from_py(&member))
                .collect::<PyResult<Vec<_>>>()?;
            Ok(PivotManualGroup {
                name: required_string(group_dict, &["name"])?,
                members,
            })
        })
        .collect()
}

fn optional_any<'py>(
    dict: &Bound<'py, PyDict>,
    keys: &[&str],
) -> PyResult<Option<Bound<'py, PyAny>>> {
    for key in keys {
        if let Some(value) = dict.get_item(*key)? {
            if !value.is_none() {
                return Ok(Some(value));
            }
        }
    }
    Ok(None)
}

fn optional_string(dict: &Bound<'_, PyDict>, keys: &[&str]) -> PyResult<Option<String>> {
    optional_any(dict, keys)?
        .map(|value| value.extract::<String>())
        .transpose()
}

fn required_string(dict: &Bound<'_, PyDict>, keys: &[&str]) -> PyResult<String> {
    optional_string(dict, keys)?
        .ok_or_else(|| PyValueError::new_err(format!("missing {}", keys[0])))
}

fn optional_string_vec(dict: &Bound<'_, PyDict>, keys: &[&str]) -> PyResult<Option<Vec<String>>> {
    optional_any(dict, keys)?
        .map(|value| value.extract::<Vec<String>>())
        .transpose()
}

fn required_string_vec(dict: &Bound<'_, PyDict>, keys: &[&str]) -> PyResult<Vec<String>> {
    optional_string_vec(dict, keys)?
        .ok_or_else(|| PyValueError::new_err(format!("missing {}", keys[0])))
}

fn optional_f64(dict: &Bound<'_, PyDict>, keys: &[&str]) -> PyResult<Option<f64>> {
    optional_any(dict, keys)?
        .map(|value| value.extract::<f64>())
        .transpose()
}

fn optional_bool(dict: &Bound<'_, PyDict>, keys: &[&str]) -> PyResult<Option<bool>> {
    optional_any(dict, keys)?
        .map(|value| value.extract::<bool>())
        .transpose()
}

fn optional_u32(dict: &Bound<'_, PyDict>, keys: &[&str]) -> PyResult<Option<u32>> {
    optional_any(dict, keys)?
        .map(|value| value.extract::<u32>())
        .transpose()
}

fn optional_pivot_value(dict: &Bound<'_, PyDict>, keys: &[&str]) -> PyResult<Option<PivotValue>> {
    optional_any(dict, keys)?
        .map(|value| pivot_value_from_py(&value))
        .transpose()
}

fn pivot_value_from_py(value: &Bound<'_, PyAny>) -> PyResult<PivotValue> {
    if let Ok(value) = value.extract::<bool>() {
        return Ok(PivotValue::Boolean(value));
    }
    if let Ok(value) = value.extract::<f64>() {
        return Ok(PivotValue::Number(value));
    }
    if let Ok(value) = value.extract::<String>() {
        return Ok(PivotValue::String(value));
    }
    Err(PyValueError::new_err(
        "pivot value must be a string, number, or boolean",
    ))
}

fn parse_pivot_layout_kind(value: &str) -> PyResult<PivotLayoutKind> {
    Ok(match value {
        "compact" => PivotLayoutKind::Compact,
        "outline" => PivotLayoutKind::Outline,
        "tabular" => PivotLayoutKind::Tabular,
        other => {
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot layout kind: {other}"
            )));
        }
    })
}

fn parse_pivot_overwrite_policy(value: &str) -> PyResult<PivotOverwritePolicy> {
    Ok(match value {
        "clearOwnedRange" | "clear_owned_range" | "clear" => PivotOverwritePolicy::ClearOwnedRange,
        "overwrite" => PivotOverwritePolicy::Overwrite,
        "failOnOccupied" | "fail_on_occupied" => PivotOverwritePolicy::FailOnOccupied,
        other => {
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot overwrite policy: {other}"
            )));
        }
    })
}

fn parse_pivot_aggregate(value: Option<&str>) -> PyResult<PivotAggregate> {
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
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot aggregate: {other}"
            )))
        }
    })
}

fn parse_pivot_filter_operator(value: Option<&str>) -> PyResult<PivotFilterOperator> {
    let value = value.ok_or_else(|| PyValueError::new_err("pivot filter requires operator"))?;
    Ok(match value {
        "equals" | "equal" | "eq" => PivotFilterOperator::Equals,
        "notEquals" | "notEqual" | "ne" => PivotFilterOperator::NotEquals,
        "lessThan" | "lt" => PivotFilterOperator::LessThan,
        "lessThanOrEqual" | "lte" => PivotFilterOperator::LessThanOrEqual,
        "greaterThan" | "gt" => PivotFilterOperator::GreaterThan,
        "greaterThanOrEqual" | "gte" => PivotFilterOperator::GreaterThanOrEqual,
        "beginsWith" | "begins_with" => PivotFilterOperator::BeginsWith,
        "doesNotBeginWith" | "does_not_begin_with" | "notBeginsWith" | "not_begins_with" => {
            PivotFilterOperator::DoesNotBeginWith
        }
        "endsWith" | "ends_with" => PivotFilterOperator::EndsWith,
        "doesNotEndWith" | "does_not_end_with" | "notEndsWith" | "not_ends_with" => {
            PivotFilterOperator::DoesNotEndWith
        }
        "contains" => PivotFilterOperator::Contains,
        "doesNotContain" | "does_not_contain" | "notContains" | "not_contains" => {
            PivotFilterOperator::DoesNotContain
        }
        other => {
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot filter operator: {other}"
            )))
        }
    })
}

fn parse_pivot_sort(value: &str) -> PyResult<PivotSort> {
    Ok(match value {
        "none" | "manual" => PivotSort::None,
        "ascending" | "asc" => PivotSort::Ascending,
        "descending" | "desc" => PivotSort::Descending,
        other => {
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot sort: {other}"
            )));
        }
    })
}

fn parse_pivot_subtotal(value: &str) -> PyResult<PivotSubtotal> {
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
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot subtotal: {other}"
            )));
        }
    })
}

fn parse_pivot_date_group_unit(value: &str) -> PyResult<PivotDateGroupUnit> {
    Ok(match value {
        "seconds" => PivotDateGroupUnit::Seconds,
        "minutes" => PivotDateGroupUnit::Minutes,
        "hours" => PivotDateGroupUnit::Hours,
        "days" => PivotDateGroupUnit::Days,
        "months" => PivotDateGroupUnit::Months,
        "quarters" => PivotDateGroupUnit::Quarters,
        "years" => PivotDateGroupUnit::Years,
        other => {
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot date grouping unit: {other}"
            )));
        }
    })
}

fn parse_pivot_show_as(
    value: &str,
    base_field: Option<String>,
    base_item: Option<PivotValue>,
) -> PyResult<PivotShowAs> {
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
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot show_as mode: {other}"
            )))
        }
    })
}

fn require_pivot_base_field(value: &str, base_field: Option<String>) -> PyResult<String> {
    base_field.ok_or_else(|| {
        PyValueError::new_err(format!("pivot show_as mode {value} requires base_field"))
    })
}

fn require_pivot_base_item(value: &str, base_item: Option<PivotValue>) -> PyResult<PivotValue> {
    base_item.ok_or_else(|| {
        PyValueError::new_err(format!("pivot show_as mode {value} requires base_item"))
    })
}

fn pivot_refresh_stats_to_py(
    py: Python<'_>,
    stats: duke_sheets::PivotRefreshStats,
) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("pivot_count", stats.pivot_count)?;
    dict.set_item("pivots_refreshed", stats.pivots_refreshed)?;
    dict.set_item("source_rows", stats.source_rows)?;
    dict.set_item("output_cells", stats.output_cells)?;
    dict.set_item("cache_hits", stats.cache_hits)?;
    dict.set_item("cache_misses", stats.cache_misses)?;
    Ok(dict.into_any().unbind())
}

fn build_workbook_connection_from_py(options: &Bound<'_, PyAny>) -> PyResult<WorkbookConnection> {
    let dict = options
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("data connection options must be a dict"))?;
    let id = optional_u32(dict, &["id"])?
        .ok_or_else(|| PyValueError::new_err("data connection options require id"))?;
    let name = required_string(dict, &["name"])?;
    let kind = optional_string(dict, &["kind"])?
        .unwrap_or_else(|| "database".to_string())
        .to_ascii_lowercase();
    let mut connection = match kind.as_str() {
        "database" | "db" => {
            WorkbookConnection::database(id, name, required_string(dict, &["connection"])?)
        }
        "web" => {
            let mut connection = WorkbookConnection::web(
                id,
                name,
                optional_string(dict, &["url"])?.unwrap_or_default(),
            );
            connection.kind = WorkbookConnectionKind::Web {
                url: optional_string(dict, &["url"])?,
                xml: optional_bool(dict, &["xml"])?.unwrap_or(false),
                source_data: optional_bool(dict, &["source_data", "sourceData"])?.unwrap_or(false),
                html_tables: optional_bool(dict, &["html_tables", "htmlTables"])?.unwrap_or(false),
                html_format: optional_string(dict, &["html_format", "htmlFormat"])?,
                post: optional_string(dict, &["post"])?,
                edit_page: optional_string(dict, &["edit_page", "editPage"])?,
            };
            connection
        }
        "text" => {
            let source_file = optional_string(dict, &["source_file", "sourceFile"])?;
            let mut connection =
                WorkbookConnection::text(id, name, source_file.clone().unwrap_or_default());
            connection.kind = WorkbookConnectionKind::Text {
                source_file,
                delimiter: optional_string(dict, &["delimiter"])?,
                first_row: optional_u32(dict, &["first_row", "firstRow"])?.unwrap_or(1),
                delimited: optional_bool(dict, &["delimited"])?.unwrap_or(true),
                decimal: optional_string(dict, &["decimal"])?,
                thousands: optional_string(dict, &["thousands"])?,
            };
            connection
        }
        "olap" => {
            let mut connection = WorkbookConnection::olap(id, name);
            connection.kind = WorkbookConnectionKind::Olap {
                local: optional_bool(dict, &["local"])?.unwrap_or(false),
                local_connection: optional_string(dict, &["local_connection", "localConnection"])?,
                local_refresh: optional_bool(dict, &["local_refresh", "localRefresh"])?
                    .unwrap_or(true),
                send_locale: optional_bool(dict, &["send_locale", "sendLocale"])?.unwrap_or(false),
                row_drill_count: optional_u32(dict, &["row_drill_count", "rowDrillCount"])?,
            };
            connection
        }
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown data connection kind: {other}"
            )))
        }
    };
    if let Some(command) = optional_string(dict, &["command"])? {
        connection = connection.with_command(command);
    }
    if let Some(command_type) = optional_u32(dict, &["command_type", "commandType"])? {
        connection = connection.with_command_type(command_type);
    }
    if let Some(refresh_on_load) = optional_bool(dict, &["refresh_on_load", "refreshOnLoad"])? {
        connection = connection.with_refresh_on_load(refresh_on_load);
    }
    if let Some(background) = optional_bool(dict, &["background"])? {
        connection = connection.with_background(background);
    }
    if let Some(save_data) = optional_bool(dict, &["save_data", "saveData"])? {
        connection = connection.with_save_data(save_data);
    }
    Ok(connection)
}

fn parse_encryption_profile(
    profile: Option<&str>,
    key_bits: Option<u32>,
    spin_count: Option<u32>,
) -> PyResult<duke_sheets::EncryptionProfile> {
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
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown encryption profile: {other:?}"
            )));
        }
    })
}

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

fn image_sizing_to_python(sizing: ImageSizing) -> &'static str {
    match sizing {
        ImageSizing::FitCell => "fit_cell",
        ImageSizing::FillCell => "fill_cell",
        ImageSizing::OriginalSize => "original_size",
        ImageSizing::Custom => "custom",
    }
}

// CellValue - Python wrapper for cell values

/// Represents a cell value in a spreadsheet.
///
/// Cell values can be one of several types:
/// - Empty (None)
/// - Number (float)
/// - Text (str)
/// - Boolean (bool)
/// - Error (str like "#DIV/0!")
/// - Formula cached results are exposed as regular cell values; formula text lives on Worksheet accessors
#[pyclass(name = "CellValue")]
#[derive(Clone)]
pub struct PyCellValue {
    inner: CoreCellValue,
}

#[pymethods]
impl PyCellValue {
    /// Check if the cell is empty
    #[getter]
    fn is_empty(&self) -> bool {
        matches!(self.inner, CoreCellValue::Empty)
    }

    /// Check if the cell contains a number
    #[getter]
    fn is_number(&self) -> bool {
        matches!(self.inner, CoreCellValue::Number(_))
    }

    /// Check if the cell contains text
    #[getter]
    fn is_text(&self) -> bool {
        matches!(self.inner, CoreCellValue::String(_))
    }

    /// Check if the cell contains a boolean
    #[getter]
    fn is_boolean(&self) -> bool {
        matches!(self.inner, CoreCellValue::Boolean(_))
    }

    /// Check if the cell contains an error
    #[getter]
    fn is_error(&self) -> bool {
        matches!(self.inner, CoreCellValue::Error(_))
    }

    /// Get the value as a number, or None if not a number
    fn as_number(&self) -> Option<f64> {
        match &self.inner {
            CoreCellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    /// Get the value as text, or None if not text
    fn as_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::String(s) => Some(s.to_string()),
            _ => None,
        }
    }

    /// Get the value as a boolean, or None if not a boolean
    fn as_boolean(&self) -> Option<bool> {
        match &self.inner {
            CoreCellValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    /// Get the error string, or None if not an error
    fn as_error(&self) -> Option<&'static str> {
        match &self.inner {
            CoreCellValue::Error(e) => Some(cell_error_to_string(e)),
            _ => None,
        }
    }

    /// Convert to a Python object (None, float, str, bool)
    fn to_python(&self, py: Python<'_>) -> PyObject {
        match &self.inner {
            CoreCellValue::Empty => py.None(),
            CoreCellValue::Number(n) => n.into_py(py),
            CoreCellValue::String(s) => s.to_string().into_py(py),
            CoreCellValue::Boolean(b) => b.into_py(py),
            CoreCellValue::Error(e) => cell_error_to_string(e).into_py(py),
            CoreCellValue::RichText(runs) => runs
                .iter()
                .map(|r| r.text.as_str())
                .collect::<String>()
                .into_py(py),
            CoreCellValue::SpillTarget { .. } => py.None(), // Spill targets appear empty
        }
    }

    fn __repr__(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => "CellValue(Empty)".to_string(),
            CoreCellValue::Number(n) => format!("CellValue(Number({}))", n),
            CoreCellValue::String(s) => format!("CellValue(Text({:?}))", s.to_string()),
            CoreCellValue::Boolean(b) => format!("CellValue(Boolean({}))", b),
            CoreCellValue::Error(e) => format!("CellValue(Error({}))", cell_error_to_string(e)),
            CoreCellValue::RichText(runs) => {
                let text = runs.iter().map(|r| r.text.as_str()).collect::<String>();
                format!("CellValue(RichText({:?}))", text)
            }
            CoreCellValue::SpillTarget { .. } => "CellValue(SpillTarget)".to_string(),
        }
    }

    fn __str__(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => "".to_string(),
            CoreCellValue::Number(n) => n.to_string(),
            CoreCellValue::String(s) => s.to_string(),
            CoreCellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CoreCellValue::Error(e) => cell_error_to_string(e).to_string(),
            CoreCellValue::RichText(runs) => runs.iter().map(|r| r.text.as_str()).collect(),
            CoreCellValue::SpillTarget { .. } => "".to_string(),
        }
    }
}

// CalculationStats - Statistics from workbook calculation

/// IMAGE() metadata captured during workbook calculation.
#[pyclass(name = "CalculationImage")]
#[derive(Clone)]
pub struct PyCalculationImage {
    /// IMAGE source URL or path.
    #[pyo3(get)]
    pub source: String,
    /// IMAGE alternate text.
    #[pyo3(get)]
    pub alt_text: String,
    /// IMAGE sizing mode.
    #[pyo3(get)]
    pub sizing: String,
    /// Optional custom width.
    #[pyo3(get)]
    pub width: Option<f64>,
    /// Optional custom height.
    #[pyo3(get)]
    pub height: Option<f64>,
}

#[pymethods]
impl PyCalculationImage {
    fn __repr__(&self) -> String {
        format!("CalculationImage(source={:?})", self.source)
    }
}

/// Statistics from calculating a workbook.
#[pyclass(name = "CalculationStats")]
#[derive(Clone)]
pub struct PyCalculationStats {
    /// Number of formulas found
    #[pyo3(get)]
    pub formula_count: usize,
    /// Number of cells calculated
    #[pyo3(get)]
    pub cells_calculated: usize,
    /// Number of errors encountered
    #[pyo3(get)]
    pub errors: usize,
    /// Number of circular references detected
    #[pyo3(get)]
    pub circular_references: usize,
    /// Number of volatile cells (e.g., NOW(), RAND())
    #[pyo3(get)]
    pub volatile_cells: usize,
    /// Whether iterative calculation converged
    #[pyo3(get)]
    pub converged: bool,
    /// Number of iterations performed
    #[pyo3(get)]
    pub iterations: usize,
}

#[pymethods]
impl PyCalculationStats {
    fn __repr__(&self) -> String {
        format!(
            "CalculationStats(formulas={}, calculated={}, errors={}, circular={}, converged={})",
            self.formula_count,
            self.cells_calculated,
            self.errors,
            self.circular_references,
            self.converged
        )
    }
}

impl From<CalculationStats> for PyCalculationStats {
    fn from(stats: CalculationStats) -> Self {
        Self {
            formula_count: stats.formula_count,
            cells_calculated: stats.cells_calculated,
            errors: stats.errors,
            circular_references: stats.circular_references,
            volatile_cells: stats.volatile_cells,
            converged: stats.converged,
            iterations: stats.iterations as usize,
        }
    }
}

// Worksheet - Python wrapper

/// A worksheet within a workbook.
///
/// Worksheets contain cells organized in rows and columns. Each cell can
/// contain a value (number, text, boolean) or a formula.
#[pyclass(name = "Worksheet")]
pub struct PyWorksheet {
    workbook: Arc<RwLock<Workbook>>,
    sheet_index: usize,
}

#[pymethods]
impl PyWorksheet {
    /// Get the worksheet name
    #[getter]
    fn name(&self) -> PyResult<String> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        wb.worksheet(self.sheet_index)
            .map(|ws| ws.name().to_string())
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))
    }

    /// Set a cell value by address (e.g., "A1", "B2")
    ///
    /// The value can be:
    /// - None (clears the cell)
    /// - int or float (number)
    /// - str (text)
    /// - bool (boolean)
    #[pyo3(signature = (address, value))]
    fn set_cell(&self, address: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let cell_value = python_to_cell_value(value)?;

        // Parse address
        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| PyValueError::new_err(format!("Invalid cell address: {}", e)))?;

        ws.set_cell_value_at(addr.row, addr.col, cell_value)
            .map_err(to_py_err)
    }

    /// Set a formula in a cell
    ///
    /// Args:
    ///     address: Cell address (e.g., "A1")
    ///     formula: Formula string (e.g., "=SUM(A1:A10)")
    #[pyo3(signature = (address, formula))]
    fn set_formula(&self, address: &str, formula: &str) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        ws.set_cell_formula(address, formula).map_err(to_py_err)
    }

    /// Set or update a cell style by address.
    ///
    /// Accepts either a Style object returned by get_cell_style() or a dict patch.
    #[pyo3(signature = (address, style))]
    fn set_cell_style(&self, address: &str, style: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| PyValueError::new_err(format!("Invalid cell address: {}", e)))?;
        let mut core_style = ws
            .cell_style_at(addr.row, addr.col)
            .cloned()
            .unwrap_or_default();
        types::apply_style_input_to_core(style, &mut core_style)?;
        ws.set_cell_style_at(addr.row, addr.col, &core_style)
            .map_err(to_py_err)
    }

    /// Set or update a cell style by row/col (0-based).
    #[pyo3(signature = (row, col, style))]
    fn set_cell_style_at(&self, row: u32, col: u32, style: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let mut core_style = ws
            .cell_style_at(row, col as u16)
            .cloned()
            .unwrap_or_default();
        types::apply_style_input_to_core(style, &mut core_style)?;
        ws.set_cell_style_at(row, col as u16, &core_style)
            .map_err(to_py_err)
    }

    /// Set or update the style for all cells in a range (e.g. "A1:C3").
    #[pyo3(signature = (range_str, style))]
    fn set_range_style(&self, range_str: &str, style: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let range = duke_sheets_core::CellRange::parse(range_str)
            .map_err(|e| PyValueError::new_err(format!("Invalid range: {}", e)))?;
        for addr in range.cells() {
            let mut core_style = ws
                .cell_style_at(addr.row, addr.col)
                .cloned()
                .unwrap_or_default();
            types::apply_style_input_to_core(style, &mut core_style)?;
            ws.set_cell_style_at(addr.row, addr.col, &core_style)
                .map_err(to_py_err)?;
        }
        Ok(())
    }

    /// Get the raw cell value (not calculated)
    #[pyo3(signature = (address))]
    fn get_cell(&self, address: &str) -> PyResult<PyCellValue> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| PyValueError::new_err(format!("Invalid cell address: {}", e)))?;

        let value = ws.get_value_at(addr.row, addr.col);

        Ok(PyCellValue { inner: value })
    }

    /// Get the raw cell value by row/col (0-based).
    #[pyo3(signature = (row, col))]
    fn get_cell_at(&self, row: u32, col: u32) -> PyResult<PyCellValue> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let value = ws.get_value_at(row, col as u16);
        Ok(PyCellValue { inner: value })
    }

    /// Get the calculated value of a cell
    ///
    /// For formulas, this returns the computed result.
    /// For regular values, returns the value itself.
    #[pyo3(signature = (address))]
    fn get_calculated_value(&self, address: &str) -> PyResult<PyCellValue> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| PyValueError::new_err(format!("Invalid cell address: {}", e)))?;

        let value = ws
            .get_calculated_value_at(addr.row, addr.col)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);

        Ok(PyCellValue { inner: value })
    }

    /// Get the calculated value of a cell by row/col (0-based).
    #[pyo3(signature = (row, col))]
    fn get_calculated_value_at(&self, row: u32, col: u32) -> PyResult<PyCellValue> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let value = ws
            .get_calculated_value_at(row, col as u16)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);

        Ok(PyCellValue { inner: value })
    }

    /// Get the used range as (min_row, min_col, max_row, max_col)
    ///
    /// Returns None if the worksheet is empty.
    #[getter]
    fn used_range(&self) -> PyResult<Option<(u32, u16, u32, u16)>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        Ok(ws
            .used_range()
            .map(|r| (r.start.row, r.start.col, r.end.row, r.end.col)))
    }

    /// Set the height of a row in points
    #[pyo3(signature = (row, height))]
    fn set_row_height(&self, row: u32, height: f64) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        ws.set_row_height(row, height);
        Ok(())
    }

    /// Set the width of a column in character units
    #[pyo3(signature = (col, width))]
    fn set_column_width(&self, col: u16, width: f64) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        ws.set_column_width(col, width);
        Ok(())
    }

    /// Merge cells in a range
    ///
    /// Args:
    ///     range_str: Range to merge (e.g., "A1:C3")
    #[pyo3(signature = (range_str))]
    fn merge_cells(&self, range_str: &str) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let range = duke_sheets_core::CellRange::parse(range_str)
            .map_err(|e| PyValueError::new_err(format!("Invalid range: {}", e)))?;
        ws.merge_cells(&range).map_err(to_py_err)
    }

    /// Unmerge cells in a range
    #[pyo3(signature = (range_str))]
    fn unmerge_cells(&self, range_str: &str) -> PyResult<bool> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let range = duke_sheets_core::CellRange::parse(range_str)
            .map_err(|e| PyValueError::new_err(format!("Invalid range: {}", e)))?;
        Ok(ws.unmerge_cells(&range))
    }

    /// Get the row height in points, or None if not explicitly set
    #[pyo3(signature = (row))]
    fn get_row_height(&self, row: u32) -> PyResult<Option<f64>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        Ok(ws.custom_row_heights().get(&row).copied())
    }

    /// Get the column width in character units, or None if not explicitly set
    #[pyo3(signature = (col))]
    fn get_column_width(&self, col: u16) -> PyResult<Option<f64>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        Ok(ws.custom_column_widths().get(&col).copied())
    }

    /// Get IMAGE() metadata for a cell, or None if no image.
    ///
    /// Args:
    ///     row: Zero-based row index
    ///     col: Zero-based column index
    ///
    /// Returns:
    ///     Dict with keys: source, alt_text, sizing, width, height - or None
    #[pyo3(signature = (row, col))]
    fn get_image_at(&self, row: u32, col: u32, py: Python<'_>) -> PyResult<Option<PyObject>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        match ws.get_image_at(row, col as u16) {
            Some(info) => {
                let dict = PyDict::new_bound(py);
                dict.set_item("source", &info.source)?;
                dict.set_item("alt_text", &info.alt_text)?;
                dict.set_item("sizing", image_sizing_to_python(info.sizing))?;
                dict.set_item("width", info.width)?;
                dict.set_item("height", info.height)?;
                Ok(Some(dict.into_any().unbind()))
            }
            None => Ok(None),
        }
    }

    /// Number of pivot tables on the worksheet.
    #[getter]
    fn pivot_count(&self) -> PyResult<usize> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.pivot_table_count())
    }

    /// Pivot table names on the worksheet.
    #[getter]
    fn pivot_table_names(&self) -> PyResult<Vec<String>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .pivot_tables()
            .iter()
            .map(|pivot| pivot.name.clone())
            .collect())
    }

    /// Add a semantic pivot table definition to the worksheet.
    #[pyo3(signature = (options))]
    fn add_pivot_table(&self, options: &Bound<'_, PyAny>) -> PyResult<()> {
        let pivot = build_pivot_table_from_py(options)?;
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        ws.add_pivot_table(pivot).map_err(to_py_err)
    }

    fn __repr__(&self) -> PyResult<String> {
        let name = self.name()?;
        Ok(format!("Worksheet({:?})", name))
    }
}

// Workbook - Python wrapper

/// A workbook containing one or more worksheets.
///
/// This is the main entry point for working with spreadsheet files.
///
/// Example:
///     >>> wb = Workbook()
///     >>> sheet = wb.get_sheet(0)
///     >>> sheet.set_cell("A1", 10)
///     >>> sheet.set_cell("A2", 20)
///     >>> sheet.set_formula("A3", "=A1+A2")
///     >>> wb.calculate()
///     >>> sheet.get_calculated_value("A3").as_number()
///     30.0
#[pyclass(name = "Workbook")]
pub struct PyWorkbook {
    inner: Arc<RwLock<Workbook>>,
}

#[pymethods]
impl PyWorkbook {
    /// Create a new empty workbook with one worksheet
    #[new]
    fn new() -> Self {
        Self {
            inner: Arc::new(RwLock::new(Workbook::new())),
        }
    }

    /// Open a workbook from a file
    ///
    /// Supported formats:
    /// - .xlsx (Excel 2007+)
    /// - .csv (Comma-separated values)
    ///
    /// Args:
    ///     path: Path to the file
    ///
    /// Returns:
    ///     Workbook instance
    #[staticmethod]
    #[pyo3(signature = (path))]
    fn open(path: &str) -> PyResult<Self> {
        use duke_sheets::WorkbookExt;
        let path = PathBuf::from(path);

        let wb = Workbook::open(&path).map_err(|e| PyIOError::new_err(e.to_string()))?;

        Ok(Self {
            inner: Arc::new(RwLock::new(wb)),
        })
    }

    /// Save the workbook to a file
    ///
    /// The format is determined by the file extension:
    /// - .xlsx for Excel format
    /// - .xls for legacy Excel binary format
    /// - .csv for CSV format (first sheet only)
    ///
    /// Args:
    ///     path: Path to save to
    #[pyo3(signature = (path))]
    fn save(&self, path: &str) -> PyResult<()> {
        use duke_sheets::WorkbookExt;
        let wb = self.inner.read().map_err(to_py_err)?;
        let path = PathBuf::from(path);

        wb.save(&path)
            .map_err(|e| PyIOError::new_err(e.to_string()))
    }

    /// Save the workbook to a password-protected file.
    ///
    /// The encryption variant is chosen by `profile`:
    /// - "default" (or None): Agile AES-256 for .xlsx, RC4 CryptoAPI 128
    ///   for .xls.
    /// - "agile": OOXML Agile (AES-CBC + HMAC-SHA*); pass key_bits to
    ///   override the default 256.
    /// - "standard": OOXML Standard Encryption (AES-ECB).
    /// - "rc4-cryptoapi": XLS RC4 CryptoAPI; key_bits 40 or 128.
    /// - "rc4-legacy": XLS legacy RC4 (MD5 KDF).
    /// - "xor": XLS XOR Obfuscation. Round-trips via duke-sheets but
    ///   does not interoperate with modern Excel; use only for legacy
    ///   reader compatibility.
    ///
    /// Args:
    ///     path: Path to save to
    ///     password: Password to encrypt with
    ///     profile: Optional encryption variant string (see above)
    ///     key_bits: Optional key size override (Agile / RC4 CryptoAPI)
    ///     spin_count: Optional iteration count (Agile only; default 100,000)
    #[pyo3(signature = (path, password, profile=None, key_bits=None, spin_count=None))]
    fn save_with_password(
        &self,
        path: &str,
        password: &str,
        profile: Option<&str>,
        key_bits: Option<u32>,
        spin_count: Option<u32>,
    ) -> PyResult<()> {
        use duke_sheets::{WorkbookExt, WorkbookSaveOptions};
        let wb = self.inner.read().map_err(to_py_err)?;
        let encryption = parse_encryption_profile(profile, key_bits, spin_count)?;
        let opts = WorkbookSaveOptions::default()
            .password(password)
            .encryption(encryption);
        wb.save_with(&PathBuf::from(path), &opts)
            .map_err(|e| PyIOError::new_err(e.to_string()))
    }

    /// Open a password-protected workbook.
    ///
    /// Args:
    ///     path: File path
    ///     password: Password to attempt
    ///     skip_integrity_check: If True, skip the HMAC integrity check
    ///         on Agile-encrypted files (matches Office behaviour).
    ///         Default False.
    #[staticmethod]
    #[pyo3(signature = (path, password, skip_integrity_check=false))]
    fn open_with_password(
        path: &str,
        password: &str,
        skip_integrity_check: bool,
    ) -> PyResult<Self> {
        use duke_sheets::{WorkbookExt, WorkbookOpenOptions};
        let mut opts = WorkbookOpenOptions::default().password(password);
        if skip_integrity_check {
            opts = opts.skip_integrity_check();
        }
        let wb = Workbook::open_with(&PathBuf::from(path), &opts)
            .map_err(|e| PyIOError::new_err(e.to_string()))?;
        Ok(Self {
            inner: Arc::new(RwLock::new(wb)),
        })
    }

    /// Get the number of worksheets
    #[getter]
    fn sheet_count(&self) -> PyResult<usize> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.sheet_count())
    }

    /// Get a list of all worksheet names
    #[getter]
    fn sheet_names(&self) -> PyResult<Vec<String>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok((0..wb.sheet_count())
            .filter_map(|i| wb.worksheet(i).map(|ws| ws.name().to_string()))
            .collect())
    }

    /// Get a worksheet by index or name
    ///
    /// Args:
    ///     index_or_name: Either an integer index or a string name
    ///
    /// Returns:
    ///     Worksheet instance
    ///
    /// Raises:
    ///     IndexError: If the index is out of range or name not found
    #[pyo3(signature = (index_or_name))]
    fn get_sheet(&self, index_or_name: &Bound<'_, PyAny>) -> PyResult<PyWorksheet> {
        let wb = self.inner.read().map_err(to_py_err)?;

        let sheet_index = if let Ok(idx) = index_or_name.extract::<usize>() {
            if idx >= wb.sheet_count() {
                return Err(PyIndexError::new_err(format!(
                    "Sheet index {} out of range (0..{})",
                    idx,
                    wb.sheet_count()
                )));
            }
            idx
        } else if let Ok(name) = index_or_name.extract::<String>() {
            wb.sheet_index(&name)
                .ok_or_else(|| PyIndexError::new_err(format!("Sheet '{}' not found", name)))?
        } else {
            return Err(PyValueError::new_err(
                "Expected int or str for sheet index/name",
            ));
        };

        drop(wb); // Release the lock

        Ok(PyWorksheet {
            workbook: Arc::clone(&self.inner),
            sheet_index,
        })
    }

    /// Add a new worksheet with the given name
    ///
    /// Args:
    ///     name: Name for the new worksheet
    ///
    /// Returns:
    ///     Index of the new worksheet
    #[pyo3(signature = (name))]
    fn add_sheet(&self, name: &str) -> PyResult<usize> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        wb.add_worksheet_with_name(name).map_err(to_py_err)
    }

    /// Remove a worksheet by index
    ///
    /// Args:
    ///     index: Index of the worksheet to remove
    #[pyo3(signature = (index))]
    fn remove_sheet(&self, index: usize) -> PyResult<()> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        wb.remove_worksheet(index).map(|_| ()).map_err(to_py_err)
    }

    /// Calculate all formulas in the workbook.
    ///
    /// Args:
    ///     iterative: Enable iterative calculation for circular references (default: False)
    ///     max_iterations: Maximum number of iterations (default: 100)
    ///     max_change: Convergence threshold (default: 0.001)
    ///     force_full_calculation: Force recalculation of all cells (default: True)
    ///     calculate_volatile: Include volatile functions like NOW(), RAND() (default: True)
    ///     sheets: Only calculate these sheet indices. Empty list means all sheets (default: [])
    ///     max_threads: Maximum threads for parallel evaluation. None means all cores (default: None)
    ///     web_service_fn: Optional callable(url: str) -> Optional[str] for WEBSERVICE evaluation
    ///     rtd_fn: Optional callable(prog_id: str, server: str, topics: list[str]) -> Optional[str] for RTD evaluation
    ///     external_fn: Optional callable(book: str, name: str, args: list[str]) -> Optional[str|float|bool]
    ///
    /// Returns:
    ///     CalculationStats with information about the calculation
    #[pyo3(signature = (*, iterative=false, max_iterations=100, max_change=0.001, force_full_calculation=true, calculate_volatile=true, sheets=vec![], max_threads=None, web_service_fn=None, rtd_fn=None, external_fn=None))]
    fn calculate(
        &self,
        iterative: bool,
        max_iterations: u32,
        max_change: f64,
        force_full_calculation: bool,
        calculate_volatile: bool,
        sheets: Vec<usize>,
        max_threads: Option<usize>,
        web_service_fn: Option<PyObject>,
        rtd_fn: Option<PyObject>,
        external_fn: Option<PyObject>,
    ) -> PyResult<PyCalculationStats> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        let web_service_fn_arc = web_service_fn.map(|py_fn| {
            Arc::new(move |url: &str| -> Option<String> {
                Python::with_gil(|py| {
                    let result = py_fn.call1(py, (url,)).ok()?;
                    if result.is_none(py) {
                        return None;
                    }
                    result.extract::<String>(py).ok()
                })
            }) as Arc<dyn Fn(&str) -> Option<String> + Send + Sync>
        });
        let rtd_fn_arc = rtd_fn.map(|py_fn| {
            Arc::new(
                move |prog_id: &str, server: &str, topics: &[String]| -> Option<String> {
                    Python::with_gil(|py| {
                        let topics_vec: Vec<String> = topics.to_vec();
                        let result = py_fn.call1(py, (prog_id, server, topics_vec)).ok()?;
                        if result.is_none(py) {
                            return None;
                        }
                        result.extract::<String>(py).ok()
                    })
                },
            ) as Arc<dyn Fn(&str, &str, &[String]) -> Option<String> + Send + Sync>
        });
        let external_fn_arc = external_fn.map(|py_fn| {
            Arc::new(
                move |book: &str, name: &str, args: &[String]| -> Option<FormulaValue> {
                    Python::with_gil(|py| {
                        let args_vec: Vec<String> = args.to_vec();
                        let result = py_fn.call1(py, (book, name, args_vec)).ok()?;
                        if result.is_none(py) {
                            return None;
                        }
                        if let Ok(b) = result.extract::<bool>(py) {
                            return Some(FormulaValue::Boolean(b));
                        }
                        if let Ok(n) = result.extract::<f64>(py) {
                            return Some(FormulaValue::Number(n));
                        }
                        if let Ok(s) = result.extract::<String>(py) {
                            return Some(FormulaValue::String(s));
                        }
                        None
                    })
                },
            )
                as Arc<dyn Fn(&str, &str, &[String]) -> Option<FormulaValue> + Send + Sync>
        });
        let options = CalculationOptions {
            iterative,
            max_iterations,
            max_change,
            force_full_calculation,
            calculate_volatile,
            sheets,
            max_threads,
            web_service_fn: web_service_fn_arc,
            rtd_fn: rtd_fn_arc,
            external_fn: external_fn_arc,
        };
        let stats = wb.calculate_with_options(&options).map_err(to_py_err)?;
        Ok(stats.into())
    }

    /// Refresh all pivot tables in the workbook.
    fn refresh_pivots(&self, py: Python<'_>) -> PyResult<PyObject> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        let stats = wb.refresh_pivots().map_err(to_py_err)?;
        pivot_refresh_stats_to_py(py, stats)
    }

    /// Add a workbook-level database connection.
    fn add_data_connection(&self, options: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        wb.add_data_connection(build_workbook_connection_from_py(options)?)
            .map_err(to_py_err)
    }

    /// Number of workbook-level data connections.
    #[getter]
    fn data_connection_count(&self) -> PyResult<usize> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.data_connections().len())
    }

    /// Workbook-level data connection names.
    #[getter]
    fn data_connection_names(&self) -> PyResult<Vec<String>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb
            .data_connections()
            .iter()
            .map(|connection| connection.name.clone())
            .collect())
    }

    /// Define a named range
    ///
    /// Args:
    ///     name: Name for the range (e.g., "TaxRate")
    ///     refers_to: What the name refers to (e.g., "Sheet1!$A$1" or "0.05")
    #[pyo3(signature = (name, refers_to))]
    fn define_name(&self, name: &str, refers_to: &str) -> PyResult<()> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        wb.define_name(name, refers_to).map_err(to_py_err)
    }

    /// Get a named range definition
    ///
    /// Args:
    ///     name: Name to look up
    ///
    /// Returns:
    ///     The refers_to string, or None if not found
    #[pyo3(signature = (name))]
    fn get_named_range(&self, name: &str) -> PyResult<Option<String>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.get_named_range(name, 0).map(|nr| nr.refers_to.clone()))
    }

    fn __repr__(&self) -> PyResult<String> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(format!("Workbook(sheets={})", wb.sheet_count()))
    }
}

// Helper functions

/// Convert a Python value to a CellValue
fn python_to_cell_value(value: &Bound<'_, PyAny>) -> PyResult<CoreCellValue> {
    if value.is_none() {
        Ok(CoreCellValue::Empty)
    } else if let Ok(b) = value.extract::<bool>() {
        Ok(CoreCellValue::Boolean(b))
    } else if let Ok(n) = value.extract::<f64>() {
        Ok(CoreCellValue::Number(n))
    } else if let Ok(n) = value.extract::<i64>() {
        Ok(CoreCellValue::Number(n as f64))
    } else if let Ok(s) = value.extract::<String>() {
        Ok(CoreCellValue::string(s))
    } else {
        Err(PyValueError::new_err(
            "Cell value must be None, bool, int, float, or str",
        ))
    }
}

// Module definition

/// duke_sheets - High-performance Excel file library for Python
///
/// This module provides fast, memory-efficient access to Excel files (.xlsx)
/// and CSV files, with full formula calculation support.
///
/// Example:
///     >>> import duke_sheets
///     >>> wb = duke_sheets.Workbook()
///     >>> sheet = wb.get_sheet(0)
///     >>> sheet.set_cell("A1", 10)
///     >>> sheet.set_formula("A2", "=A1*2")
///     >>> wb.calculate()
///     >>> print(sheet.get_calculated_value("A2").as_number())
///     20.0
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<PyCellValue>()?;
    m.add_class::<PyCalculationImage>()?;
    m.add_class::<PyCalculationStats>()?;
    m.add_class::<PyColor>()?;
    m.add_class::<PyFontStyle>()?;
    m.add_class::<PyGradientStop>()?;
    m.add_class::<PyFillStyle>()?;
    m.add_class::<PyBorderEdge>()?;
    m.add_class::<PyBorderStyle>()?;
    m.add_class::<PyAlignment>()?;
    m.add_class::<PyNumberFormat>()?;
    m.add_class::<PyCellProtection>()?;
    m.add_class::<PyStyle>()?;
    m.add_class::<PyHyperlink>()?;
    m.add_class::<PyComment>()?;
    m.add_class::<PyCommentEntry>()?;
    m.add_class::<PyFreezePanes>()?;
    m.add_class::<PySplitPanes>()?;
    m.add_class::<PySelection>()?;
    m.add_class::<PySheetProtection>()?;
    m.add_class::<PyPageSetup>()?;
    m.add_class::<PyPageBreak>()?;
    m.add_class::<PyWorkbookSettings>()?;
    m.add_class::<PyNamedRange>()?;
    m.add_class::<PyTable>()?;
    m.add_class::<PyTableColumn>()?;
    m.add_class::<PyTableStyleInfo>()?;
    m.add_class::<PyAutoFilter>()?;
    m.add_class::<PyFilterColumn>()?;
    m.add_class::<PyDataValidation>()?;
    m.add_class::<PyConditionalFormatRule>()?;
    m.add_class::<PyRichTextRun>()?;
    m.add_class::<PyRunFont>()?;
    m.add_class::<PyHyperlinkEntry>()?;
    m.add_class::<PyRowCell>()?;
    m.add_class::<PyRow>()?;
    m.add_class::<PyRowIterator>()?;
    m.add_class::<PyFormulaCell>()?;
    m.add_class::<PySpillSource>()?;
    m.add_class::<PyMergedRegion>()?;
    m.add_class::<PyMergeSpan>()?;
    m.add_class::<PyChart>()?;
    m.add_class::<PyPivotChartSource>()?;
    m.add_class::<PyDrawingAnchor>()?;
    m.add_class::<PyDataSeries>()?;
    m.add_class::<PyDataReference>()?;
    m.add_class::<PyAxis>()?;
    m.add_class::<PyLegend>()?;
    m.add_class::<PyChartNumberFormat>()?;
    m.add_class::<PyChartShapeProperties>()?;
    m.add_class::<PyDataLabels>()?;
    m.add_class::<PyTrendline>()?;
    m.add_class::<PyErrorBars>()?;
    m.add_class::<PyMarker>()?;
    m.add_class::<PyDataPoint>()?;
    m.add_class::<PyView3D>()?;
    m.add_class::<PyChartDataTable>()?;
    m.add_class::<PyManualLayout>()?;
    m.add_class::<PyLayout>()?;
    m.add_class::<PyChartTypeGroup>()?;
    m.add_class::<PyChartAxis>()?;
    m.add_class::<PyChartLines>()?;
    m.add_class::<PyUpDownBars>()?;
    m.add_class::<PyChartSheet>()?;
    m.add_class::<PySheetSlot>()?;
    m.add_class::<PyChartEx>()?;
    m.add_class::<PyChartExSeries>()?;
    m.add_class::<PyChartExData>()?;
    m.add_class::<PyChartExDimension>()?;
    m.add_class::<PyChartExAxis>()?;
    m.add_class::<PyChartExLegend>()?;
    m.add_class::<PyChartExDataLabels>()?;
    m.add_class::<PyChartExTitle>()?;
    m.add_class::<PyChartExLayoutPr>()?;
    m.add_class::<PyChartExOffset>()?;
    m.add_class::<PyChartExText>()?;
    m.add_class::<PyChartExColorPosition>()?;
    m.add_class::<PyChartExValueColorPositions>()?;
    m.add_class::<PyChartExScaling>()?;
    m.add_class::<PyChartExAxisTitle>()?;
    m.add_class::<PyChartExAxisUnits>()?;
    m.add_class::<PyChartExSeriesVisibility>()?;
    m.add_class::<PyChartExBinning>()?;
    m.add_class::<PyChartExGeography>()?;
    m.add_class::<PyChartExStatistics>()?;
    m.add_class::<PyChartExDataPoint>()?;
    m.add_class::<PyChartExDataLabel>()?;
    m.add_class::<PyChartExFormatOverride>()?;
    m.add_class::<PyChartExHeaderFooter>()?;
    m.add_class::<PyChartExPageMargins>()?;
    m.add_class::<PyChartExPageSetup>()?;
    m.add_class::<PyChartExPrintSettings>()?;
    m.add_class::<PyChartExPlotArea>()?;
    m.add_class::<PyEmbeddedImage>()?;
    Ok(())
}
