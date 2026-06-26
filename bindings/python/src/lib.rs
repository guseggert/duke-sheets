//! Python bindings for duke-sheets
//!
//! This module provides PyO3-based Python bindings for the duke-sheets library,
//! allowing Python code to read, write, and manipulate Excel files.

use pyo3::exceptions::{PyIOError, PyIndexError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList};

use std::path::PathBuf;
use std::sync::{Arc, RwLock};

use duke_sheets::prelude::*;
use duke_sheets::{
    CalculationOptions, ChartType, FormulaValue, ImageSizing, PivotRefreshOptions,
    WorkbookCalculationExt, WorkbookPivotExt,
};
use duke_sheets_core::{
    CellError, CellValue as CoreCellValue, PivotAggregate, PivotCalculatedField,
    PivotCalculatedItem, PivotDateGroupUnit, PivotDatePeriod, PivotField, PivotFilter,
    PivotFilterOperator, PivotGrouping, PivotLayout, PivotLayoutKind, PivotManualGroup,
    PivotMeasure, PivotOverwritePolicy, PivotRefreshPolicy, PivotRefreshStatus, PivotShowAs,
    PivotSort, PivotSource, PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable, PivotValue,
    PivotValuesAxis, WorkbookConnection, WorkbookConnectionKind, WorkbookExtension,
    WorkbookExtensionPart,
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
    let olap_connection_name =
        optional_string(dict, &["olap_connection_name", "olapConnectionName"])?;
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
        + usize::from(olap_connection_name.is_some())
        + usize::from(consolidation_ranges_value.is_some());
    if source_count != 1 {
        return Err(PyValueError::new_err(
            "Pivot options require exactly one of table_name/tableName, source_range/sourceRange, external_connection_name/externalConnectionName, olap_connection_name/olapConnectionName, or consolidation_ranges/consolidationRanges",
        ));
    }

    match (
        table_name,
        source_range,
        external_connection_name,
        olap_connection_name,
        consolidation_ranges_value,
    ) {
        (Some(table_name), None, None, None, None) => {
            builder = builder.table_source(table_name);
        }
        (None, Some(source_range), None, None, None) => {
            let range = CellRange::parse(&source_range)
                .map_err(|e| PyValueError::new_err(format!("Invalid pivot source range: {e}")))?;
            builder = if let Some(sheet) = optional_string(dict, &["source_sheet", "sourceSheet"])?
            {
                builder.source_range_on_sheet(sheet, range)
            } else {
                builder.source_range(range)
            };
        }
        (None, None, Some(connection_name), None, None) => {
            builder = builder.source(PivotSource::External {
                connection_name,
                command_text: external_command_text,
            });
        }
        (None, None, None, Some(connection_name), None) => {
            builder = builder.source(PivotSource::Olap {
                connection_name,
                cube: None,
                command_text: None,
            });
        }
        (None, None, None, None, Some(consolidation_ranges)) => {
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
    if let Some(calculated_items_value) =
        optional_any(dict, &["calculated_items", "calculatedItems"])?
    {
        let calculated_items = calculated_items_value
            .downcast::<PyList>()
            .map_err(|_| PyValueError::new_err("pivot calculated_items must be a list"))?;
        for calculated_item in calculated_items.iter() {
            let calculated_item = calculated_item
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("pivot calculated item must be a dict"))?;
            let item_value = calculated_item
                .get_item("item")?
                .ok_or_else(|| PyValueError::new_err("pivot calculated item requires item"))?;
            builder = builder.calculated_item(
                required_string(calculated_item, &["field"])?,
                pivot_value_from_py(&item_value)?,
                required_string(calculated_item, &["formula"])?,
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
    field.subtotal_caption =
        optional_string(dict, &["subtotal_caption", "subtotalCaption"])?;
    if let Some(subtotals) = optional_string_vec(dict, &["subtotals"])? {
        let subtotals = subtotals
            .into_iter()
            .map(|subtotal| parse_pivot_subtotal(&subtotal))
            .collect::<PyResult<Vec<_>>>()?;
        field = field.with_subtotals(subtotals);
    }
    if let Some(values) = optional_pivot_values(dict, &["collapsed_items", "collapsedItems"])? {
        field.collapsed_items = values;
    }
    if let Some(show_empty_items) = optional_bool(dict, &["show_empty_items", "showEmptyItems"])? {
        field.show_empty_items = show_empty_items;
    }
    if let Some(value) = optional_bool(dict, &["show_drop_downs", "showDropDowns"])? {
        field.show_drop_downs = value;
    }
    if let Some(value) = optional_bool(dict, &["subtotal_top", "subtotalTop"])? {
        field.subtotal_top = value;
    }
    if let Some(value) = optional_bool(dict, &["insert_blank_row", "insertBlankRow"])? {
        field.insert_blank_row = value;
    }
    if let Some(value) = optional_bool(dict, &["insert_page_break", "insertPageBreak"])? {
        field.insert_page_break = value;
    }
    if let Some(value) = optional_bool(
        dict,
        &["include_new_items_in_filter", "includeNewItemsInFilter"],
    )? {
        field.include_new_items_in_filter = value;
    }
    if let Some(value) = optional_u32(dict, &["item_page_count", "itemPageCount"])? {
        field.item_page_count = value;
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
    if let Some(value) = optional_u32(dict, &["page_wrap", "pageWrap"])? {
        layout.page_wrap = value;
    }
    if let Some(value) = optional_bool(dict, &["page_over_then_down", "pageOverThenDown"])? {
        layout.page_over_then_down = value;
    }
    if let Some(value) = optional_bool(dict, &["merge_item_labels", "mergeItemLabels"])? {
        layout.merge_item_labels = value;
    }
    if let Some(value) = optional_string(dict, &["data_caption", "dataCaption"])? {
        layout.data_caption = value;
    }
    if let Some(value) = optional_string(dict, &["values_axis", "valuesAxis"])? {
        layout.values_axis = parse_pivot_values_axis(&value)?;
    }
    layout.values_axis_position =
        optional_u32(dict, &["values_axis_position", "valuesAxisPosition"])?;
    if let Some(value) = optional_string(dict, &["grand_total_caption", "grandTotalCaption"])? {
        layout.grand_total_caption = Some(value);
    }
    if let Some(value) = optional_string(dict, &["error_caption", "errorCaption"])? {
        layout.error_caption = Some(value);
    }
    if let Some(value) = optional_bool(dict, &["show_error", "showError"])? {
        layout.show_error = value;
    }
    if let Some(value) = optional_string(dict, &["missing_caption", "missingCaption"])? {
        layout.missing_caption = Some(value);
    }
    if let Some(value) = optional_bool(dict, &["show_missing", "showMissing"])? {
        layout.show_missing = value;
    }
    if let Some(value) = optional_bool(dict, &["asterisk_totals", "asteriskTotals"])? {
        layout.asterisk_totals = value;
    }
    if let Some(value) = optional_bool(dict, &["show_items", "showItems"])? {
        layout.show_items = value;
    }
    if let Some(value) = optional_bool(dict, &["edit_data", "editData"])? {
        layout.edit_data = value;
    }
    if let Some(value) = optional_bool(dict, &["disable_field_list", "disableFieldList"])? {
        layout.disable_field_list = value;
    }
    if let Some(value) = optional_bool(dict, &["show_calculated_members", "showCalculatedMembers"])?
    {
        layout.show_calculated_members = value;
    }
    if let Some(value) = optional_bool(dict, &["visual_totals", "visualTotals"])? {
        layout.visual_totals = value;
    }
    if let Some(value) = optional_bool(dict, &["show_multiple_label", "showMultipleLabel"])? {
        layout.show_multiple_label = value;
    }
    if let Some(value) = optional_bool(dict, &["show_data_drop_down", "showDataDropDown"])? {
        layout.show_data_drop_down = value;
    }
    if let Some(value) = optional_bool(
        dict,
        &["show_member_property_tips", "showMemberPropertyTips"],
    )? {
        layout.show_member_property_tips = value;
    }
    if let Some(value) = optional_bool(dict, &["show_data_tips", "showDataTips"])? {
        layout.show_data_tips = value;
    }
    if let Some(value) = optional_bool(dict, &["enable_wizard", "enableWizard"])? {
        layout.enable_wizard = value;
    }
    if let Some(value) = optional_bool(dict, &["enable_drill", "enableDrill"])? {
        layout.enable_drill = value;
    }
    if let Some(value) = optional_bool(dict, &["enable_field_properties", "enableFieldProperties"])?
    {
        layout.enable_field_properties = value;
    }
    if let Some(value) = optional_bool(dict, &["subtotal_hidden_items", "subtotalHiddenItems"])? {
        layout.subtotal_hidden_items = value;
    }
    if let Some(value) = optional_bool(dict, &["show_drop_zones", "showDropZones"])? {
        layout.show_drop_zones = value;
    }
    if let Some(value) = optional_u32(dict, &["indent"])? {
        layout.indent = value;
    }
    if let Some(value) = optional_bool(dict, &["show_empty_rows", "showEmptyRows"])? {
        layout.show_empty_rows = value;
    }
    if let Some(value) = optional_bool(dict, &["show_empty_columns", "showEmptyColumns"])? {
        layout.show_empty_columns = value;
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
        "labelBetween" | "label_between" | "captionBetween" | "caption_between" => {
            Ok(PivotFilter::LabelBetween {
                field: field.into(),
                start: optional_string(dict, &["start_text", "startText", "text"])?
                    .ok_or_else(|| {
                        PyValueError::new_err("pivot label-between filter requires start_text")
                    })?,
                end: optional_string(dict, &["end_text", "endText"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot label-between filter requires end_text")
                })?,
                not_between: false,
            })
        }
        "labelNotBetween" | "label_not_between" | "captionNotBetween" | "caption_not_between" => {
            Ok(PivotFilter::LabelBetween {
                field: field.into(),
                start: optional_string(dict, &["start_text", "startText", "text"])?
                    .ok_or_else(|| {
                        PyValueError::new_err("pivot label-not-between filter requires start_text")
                    })?,
                end: optional_string(dict, &["end_text", "endText"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot label-not-between filter requires end_text")
                })?,
                not_between: true,
            })
        }
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
        "valueBetween" | "value_between" | "valueRange" | "value_range" => {
            let measure = optional_any(dict, &["measure"])?.ok_or_else(|| {
                PyValueError::new_err("pivot value-between filter requires measure")
            })?;
            Ok(PivotFilter::ValueBetween {
                field: field.into(),
                measure: build_pivot_measure_from_py(&measure)?,
                start: optional_f64(dict, &["start", "value"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot value-between filter requires start")
                })?,
                end: optional_f64(dict, &["end"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot value-between filter requires end")
                })?,
                not_between: false,
            })
        }
        "valueNotBetween" | "value_not_between" | "valueNotRange" | "value_not_range" => {
            let measure = optional_any(dict, &["measure"])?.ok_or_else(|| {
                PyValueError::new_err("pivot value-not-between filter requires measure")
            })?;
            Ok(PivotFilter::ValueBetween {
                field: field.into(),
                measure: build_pivot_measure_from_py(&measure)?,
                start: optional_f64(dict, &["start", "value"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot value-not-between filter requires start")
                })?,
                end: optional_f64(dict, &["end"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot value-not-between filter requires end")
                })?,
                not_between: true,
            })
        }
        "date" => Ok(PivotFilter::Date {
            field: field.into(),
            operator: parse_pivot_filter_operator(
                optional_string(dict, &["operator"])?.as_deref(),
            )?,
            value: optional_f64(dict, &["value", "start"])?
                .ok_or_else(|| PyValueError::new_err("pivot date filter requires value"))?,
        }),
        "dateBetween" | "date_between" | "dateRange" | "date_range" => {
            Ok(PivotFilter::DateBetween {
                field: field.into(),
                start: optional_f64(dict, &["start", "value"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot date-between filter requires start")
                })?,
                end: optional_f64(dict, &["end"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot date-between filter requires end")
                })?,
                not_between: false,
            })
        }
        "dateNotBetween" | "date_not_between" | "dateNotRange" | "date_not_range" => {
            Ok(PivotFilter::DateBetween {
                field: field.into(),
                start: optional_f64(dict, &["start", "value"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot date-not-between filter requires start")
                })?,
                end: optional_f64(dict, &["end"])?.ok_or_else(|| {
                    PyValueError::new_err("pivot date-not-between filter requires end")
                })?,
                not_between: true,
            })
        }
        "datePeriod" | "date_period" | "period" => Ok(PivotFilter::DatePeriod {
            field: field.into(),
            period: parse_pivot_date_period(
                optional_string(dict, &["period"])?
                    .as_deref()
                    .ok_or_else(|| {
                        PyValueError::new_err("pivot date-period filter requires period")
                    })?,
            )?,
        }),
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

fn parse_pivot_date_period(value: &str) -> PyResult<PivotDatePeriod> {
    Ok(match value {
        "tomorrow" => PivotDatePeriod::Tomorrow,
        "today" => PivotDatePeriod::Today,
        "yesterday" => PivotDatePeriod::Yesterday,
        "nextWeek" | "next_week" => PivotDatePeriod::NextWeek,
        "thisWeek" | "this_week" => PivotDatePeriod::ThisWeek,
        "lastWeek" | "last_week" => PivotDatePeriod::LastWeek,
        "nextMonth" | "next_month" => PivotDatePeriod::NextMonth,
        "thisMonth" | "this_month" => PivotDatePeriod::ThisMonth,
        "lastMonth" | "last_month" => PivotDatePeriod::LastMonth,
        "nextQuarter" | "next_quarter" => PivotDatePeriod::NextQuarter,
        "thisQuarter" | "this_quarter" => PivotDatePeriod::ThisQuarter,
        "lastQuarter" | "last_quarter" => PivotDatePeriod::LastQuarter,
        "nextYear" | "next_year" => PivotDatePeriod::NextYear,
        "thisYear" | "this_year" => PivotDatePeriod::ThisYear,
        "lastYear" | "last_year" => PivotDatePeriod::LastYear,
        "yearToDate" | "year_to_date" => PivotDatePeriod::YearToDate,
        "Q1" | "quarter1" | "quarter_1" => PivotDatePeriod::Quarter(1),
        "Q2" | "quarter2" | "quarter_2" => PivotDatePeriod::Quarter(2),
        "Q3" | "quarter3" | "quarter_3" => PivotDatePeriod::Quarter(3),
        "Q4" | "quarter4" | "quarter_4" => PivotDatePeriod::Quarter(4),
        "M1" | "month1" | "month_1" => PivotDatePeriod::Month(1),
        "M2" | "month2" | "month_2" => PivotDatePeriod::Month(2),
        "M3" | "month3" | "month_3" => PivotDatePeriod::Month(3),
        "M4" | "month4" | "month_4" => PivotDatePeriod::Month(4),
        "M5" | "month5" | "month_5" => PivotDatePeriod::Month(5),
        "M6" | "month6" | "month_6" => PivotDatePeriod::Month(6),
        "M7" | "month7" | "month_7" => PivotDatePeriod::Month(7),
        "M8" | "month8" | "month_8" => PivotDatePeriod::Month(8),
        "M9" | "month9" | "month_9" => PivotDatePeriod::Month(9),
        "M10" | "month10" | "month_10" => PivotDatePeriod::Month(10),
        "M11" | "month11" | "month_11" => PivotDatePeriod::Month(11),
        "M12" | "month12" | "month_12" => PivotDatePeriod::Month(12),
        other => {
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot date period: {other}"
            )))
        }
    })
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

fn optional_pivot_values(
    dict: &Bound<'_, PyDict>,
    keys: &[&str],
) -> PyResult<Option<Vec<PivotValue>>> {
    optional_any(dict, keys)?
        .map(|value| {
            let values = value
                .downcast::<PyList>()
                .map_err(|_| PyValueError::new_err("pivot value list must be a list"))?;
            values
                .iter()
                .map(|item| pivot_value_from_py(&item))
                .collect()
        })
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

fn parse_pivot_values_axis(value: &str) -> PyResult<PivotValuesAxis> {
    Ok(match value {
        "columns" | "column" | "cols" => PivotValuesAxis::Columns,
        "rows" | "row" => PivotValuesAxis::Rows,
        other => {
            return Err(PyValueError::new_err(format!(
                "Unsupported pivot values axis: {other}"
            )));
        }
    })
}

fn parse_chart_type(value: Option<&str>) -> PyResult<ChartType> {
    match value {
        Some(value) => ChartType::from_name(value)
            .ok_or_else(|| PyValueError::new_err(format!("Unsupported chart type: {value}"))),
        None => Ok(ChartType::ColumnClustered),
    }
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
        "percentOfParentRowTotal" | "percentOfParentRow" | "percent_of_parent_row_total"
        | "percent_of_parent_row" => PivotShowAs::PercentOfParentRowTotal,
        "percentOfParentColumnTotal" | "percentOfParentCol"
        | "percent_of_parent_column_total" | "percent_of_parent_col" => {
            PivotShowAs::PercentOfParentColumnTotal
        }
        "percentOfParentTotal" | "percentOfParent" | "percent_of_parent_total"
        | "percent_of_parent" => PivotShowAs::PercentOfParentTotal {
            base_field: require_pivot_base_field(value, base_field)?.into(),
        },
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

fn pivot_table_to_py(py: Python<'_>, pivot: &PivotTable) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("id", pivot.id)?;
    dict.set_item("name", &pivot.name)?;
    dict.set_item("source", pivot_source_to_py(py, &pivot.source)?)?;
    dict.set_item("target", pivot.target.to_string())?;
    dict.set_item("rows", pivot_fields_to_py(py, &pivot.rows)?)?;
    dict.set_item("columns", pivot_fields_to_py(py, &pivot.columns)?)?;
    dict.set_item("page_fields", pivot_fields_to_py(py, &pivot.page_fields)?)?;
    dict.set_item("filters", pivot_filters_to_py(py, &pivot.filters)?)?;
    dict.set_item(
        "calculated_fields",
        pivot_calculated_fields_to_py(py, &pivot.calculated_fields)?,
    )?;
    dict.set_item(
        "calculated_items",
        pivot_calculated_items_to_py(py, &pivot.calculated_items)?,
    )?;
    dict.set_item("measures", pivot_measures_to_py(py, &pivot.measures)?)?;
    dict.set_item("groupings", pivot_groupings_to_py(py, &pivot.groupings)?)?;
    dict.set_item("layout", pivot_layout_to_py(py, &pivot.layout)?)?;
    dict.set_item("style", pivot_style_to_py(py, &pivot.style)?)?;
    dict.set_item(
        "refresh_policy",
        pivot_refresh_policy_to_py(py, &pivot.refresh_policy)?,
    )?;
    dict.set_item(
        "overwrite_policy",
        pivot_overwrite_policy_to_python(pivot.overwrite_policy),
    )?;
    dict.set_item(
        "rendered_range",
        pivot.rendered_range.map(|range| range.to_string()),
    )?;
    dict.set_item(
        "refresh_status",
        pivot_refresh_status_to_py(py, &pivot.refresh_status)?,
    )?;
    dict.set_item("extension_count", pivot.extensions.len())?;
    Ok(dict.into_any().unbind())
}

fn pivot_source_to_py(py: Python<'_>, source: &PivotSource) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    match source {
        PivotSource::WorksheetRange { sheet, range } => {
            dict.set_item("kind", "worksheet_range")?;
            dict.set_item("sheet", sheet)?;
            dict.set_item("range", range.to_string())?;
        }
        PivotSource::Table { name } => {
            dict.set_item("kind", "table")?;
            dict.set_item("table_name", name)?;
        }
        PivotSource::External {
            connection_name,
            command_text,
        } => {
            dict.set_item("kind", "external")?;
            dict.set_item("connection_name", connection_name)?;
            dict.set_item("command_text", command_text)?;
        }
        PivotSource::Consolidation { ranges } => {
            dict.set_item("kind", "consolidation")?;
            let items = PyList::empty_bound(py);
            for range in ranges {
                items.append(pivot_source_range_to_py(py, range)?)?;
            }
            dict.set_item("ranges", items)?;
        }
        PivotSource::Scenario { name } => {
            dict.set_item("kind", "scenario")?;
            dict.set_item("scenario_name", name)?;
        }
        PivotSource::Olap {
            connection_name,
            cube,
            command_text,
        } => {
            dict.set_item("kind", "olap")?;
            dict.set_item("connection_name", connection_name)?;
            dict.set_item("cube", cube)?;
            dict.set_item("command_text", command_text)?;
        }
    }
    Ok(dict.into_any().unbind())
}

fn pivot_source_range_to_py(py: Python<'_>, range: &PivotSourceRange) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("sheet", &range.sheet)?;
    dict.set_item("range", range.range.to_string())?;
    dict.set_item("name", &range.name)?;
    dict.set_item("page_items", &range.page_items)?;
    Ok(dict.into_any().unbind())
}

fn pivot_fields_to_py(py: Python<'_>, fields: &[PivotField]) -> PyResult<PyObject> {
    let items = PyList::empty_bound(py);
    for field in fields {
        items.append(pivot_field_to_py(py, field)?)?;
    }
    Ok(items.into_any().unbind())
}

fn pivot_field_to_py(py: Python<'_>, field: &PivotField) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("field", &field.field.name)?;
    dict.set_item("sort", pivot_sort_to_python(field.sort))?;
    dict.set_item("subtotal", pivot_subtotal_to_python(field.subtotal))?;
    dict.set_item("subtotal_caption", &field.subtotal_caption)?;
    let subtotals = field
        .subtotals
        .iter()
        .map(|subtotal| pivot_subtotal_to_python(*subtotal))
        .collect::<Vec<_>>();
    dict.set_item("subtotals", subtotals)?;
    let collapsed_items = PyList::empty_bound(py);
    for item in &field.collapsed_items {
        collapsed_items.append(pivot_value_to_py(py, item)?)?;
    }
    dict.set_item("collapsed_items", collapsed_items)?;
    dict.set_item("show_empty_items", field.show_empty_items)?;
    dict.set_item("show_drop_downs", field.show_drop_downs)?;
    dict.set_item("subtotal_top", field.subtotal_top)?;
    dict.set_item("insert_blank_row", field.insert_blank_row)?;
    dict.set_item("insert_page_break", field.insert_page_break)?;
    dict.set_item(
        "include_new_items_in_filter",
        field.include_new_items_in_filter,
    )?;
    dict.set_item("item_page_count", field.item_page_count)?;
    Ok(dict.into_any().unbind())
}

fn pivot_measures_to_py(py: Python<'_>, measures: &[PivotMeasure]) -> PyResult<PyObject> {
    let items = PyList::empty_bound(py);
    for measure in measures {
        items.append(pivot_measure_to_py(py, measure)?)?;
    }
    Ok(items.into_any().unbind())
}

fn pivot_measure_to_py(py: Python<'_>, measure: &PivotMeasure) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("field", &measure.field.name)?;
    dict.set_item("aggregate", pivot_aggregate_to_python(measure.aggregate))?;
    dict.set_item("name", &measure.name)?;
    dict.set_item("caption", measure.caption())?;
    dict.set_item("show_as", pivot_show_as_to_py(py, &measure.show_as)?)?;
    dict.set_item("number_format", &measure.number_format)?;
    Ok(dict.into_any().unbind())
}

fn pivot_show_as_to_py(py: Python<'_>, show_as: &PivotShowAs) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    match show_as {
        PivotShowAs::Normal => {
            dict.set_item("kind", "normal")?;
        }
        PivotShowAs::PercentOfGrandTotal => {
            dict.set_item("kind", "percent_of_grand_total")?;
        }
        PivotShowAs::PercentOfRowTotal => {
            dict.set_item("kind", "percent_of_row_total")?;
        }
        PivotShowAs::PercentOfColumnTotal => {
            dict.set_item("kind", "percent_of_column_total")?;
        }
        PivotShowAs::PercentOfParentRowTotal => {
            dict.set_item("kind", "percent_of_parent_row_total")?;
        }
        PivotShowAs::PercentOfParentColumnTotal => {
            dict.set_item("kind", "percent_of_parent_column_total")?;
        }
        PivotShowAs::PercentOfParentTotal { base_field } => {
            dict.set_item("kind", "percent_of_parent_total")?;
            dict.set_item("base_field", &base_field.name)?;
        }
        PivotShowAs::Index => {
            dict.set_item("kind", "index")?;
        }
        PivotShowAs::RunningTotal { base_field } => {
            dict.set_item("kind", "running_total")?;
            dict.set_item("base_field", &base_field.name)?;
        }
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        } => {
            dict.set_item("kind", "difference_from")?;
            dict.set_item("base_field", &base_field.name)?;
            dict.set_item("base_item", pivot_value_to_py(py, base_item)?)?;
        }
        PivotShowAs::PercentDifferenceFrom {
            base_field,
            base_item,
        } => {
            dict.set_item("kind", "percent_difference_from")?;
            dict.set_item("base_field", &base_field.name)?;
            dict.set_item("base_item", pivot_value_to_py(py, base_item)?)?;
        }
        PivotShowAs::RankAscending { base_field } => {
            dict.set_item("kind", "rank_ascending")?;
            dict.set_item("base_field", &base_field.name)?;
        }
        PivotShowAs::RankDescending { base_field } => {
            dict.set_item("kind", "rank_descending")?;
            dict.set_item("base_field", &base_field.name)?;
        }
    }
    Ok(dict.into_any().unbind())
}

fn pivot_value_to_py(py: Python<'_>, value: &PivotValue) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    match value {
        PivotValue::Blank => {
            dict.set_item("kind", "blank")?;
        }
        PivotValue::Boolean(value) => {
            dict.set_item("kind", "boolean")?;
            dict.set_item("boolean", value)?;
        }
        PivotValue::Number(value) => {
            dict.set_item("kind", "number")?;
            dict.set_item("number", value)?;
        }
        PivotValue::String(value) => {
            dict.set_item("kind", "string")?;
            dict.set_item("text", value)?;
        }
        PivotValue::Error(value) => {
            dict.set_item("kind", "error")?;
            dict.set_item("error", value.to_string())?;
        }
    }
    Ok(dict.into_any().unbind())
}

fn pivot_filters_to_py(py: Python<'_>, filters: &[PivotFilter]) -> PyResult<PyObject> {
    let items = PyList::empty_bound(py);
    for filter in filters {
        items.append(pivot_filter_to_py(py, filter)?)?;
    }
    Ok(items.into_any().unbind())
}

fn pivot_filter_to_py(py: Python<'_>, filter: &PivotFilter) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    match filter {
        PivotFilter::FieldItems {
            field,
            allowed_items,
        } => {
            dict.set_item("kind", "field_items")?;
            dict.set_item("field", &field.name)?;
            let items = PyList::empty_bound(py);
            for item in allowed_items {
                items.append(pivot_value_to_py(py, item)?)?;
            }
            dict.set_item("items", items)?;
        }
        PivotFilter::Label {
            field,
            operator,
            value,
        } => {
            dict.set_item("kind", "label")?;
            dict.set_item("field", &field.name)?;
            dict.set_item("operator", pivot_filter_operator_to_python(*operator))?;
            dict.set_item("text", value)?;
        }
        PivotFilter::LabelBetween {
            field,
            start,
            end,
            not_between,
        } => {
            dict.set_item(
                "kind",
                if *not_between {
                    "label_not_between"
                } else {
                    "label_between"
                },
            )?;
            dict.set_item("field", &field.name)?;
            dict.set_item("start_text", start)?;
            dict.set_item("end_text", end)?;
        }
        PivotFilter::Value {
            field,
            measure,
            operator,
            value,
        } => {
            dict.set_item("kind", "value")?;
            dict.set_item("field", &field.name)?;
            dict.set_item("measure", pivot_measure_to_py(py, measure)?)?;
            dict.set_item("operator", pivot_filter_operator_to_python(*operator))?;
            dict.set_item("value", value)?;
        }
        PivotFilter::ValueBetween {
            field,
            measure,
            start,
            end,
            not_between,
        } => {
            dict.set_item(
                "kind",
                if *not_between {
                    "value_not_between"
                } else {
                    "value_between"
                },
            )?;
            dict.set_item("field", &field.name)?;
            dict.set_item("measure", pivot_measure_to_py(py, measure)?)?;
            dict.set_item("start", start)?;
            dict.set_item("end", end)?;
        }
        PivotFilter::Date {
            field,
            operator,
            value,
        } => {
            dict.set_item("kind", "date")?;
            dict.set_item("field", &field.name)?;
            dict.set_item("operator", pivot_filter_operator_to_python(*operator))?;
            dict.set_item("value", value)?;
        }
        PivotFilter::DateBetween {
            field,
            start,
            end,
            not_between,
        } => {
            dict.set_item(
                "kind",
                if *not_between {
                    "date_not_between"
                } else {
                    "date_between"
                },
            )?;
            dict.set_item("field", &field.name)?;
            dict.set_item("start", start)?;
            dict.set_item("end", end)?;
        }
        PivotFilter::DatePeriod { field, period } => {
            dict.set_item("kind", "date_period")?;
            dict.set_item("field", &field.name)?;
            dict.set_item("period", pivot_date_period_to_python(*period))?;
        }
        PivotFilter::TopN {
            field,
            measure,
            n,
            top,
            percent,
        } => {
            dict.set_item("kind", "top_n")?;
            dict.set_item("field", &field.name)?;
            dict.set_item("measure", pivot_measure_to_py(py, measure)?)?;
            dict.set_item("n", n)?;
            dict.set_item("top", top)?;
            dict.set_item("percent", percent)?;
        }
        PivotFilter::Unsupported { kind, detail } => {
            dict.set_item("kind", kind)?;
            dict.set_item("detail", detail)?;
        }
    }
    Ok(dict.into_any().unbind())
}

fn pivot_date_period_to_python(period: PivotDatePeriod) -> &'static str {
    match period {
        PivotDatePeriod::Tomorrow => "tomorrow",
        PivotDatePeriod::Today => "today",
        PivotDatePeriod::Yesterday => "yesterday",
        PivotDatePeriod::NextWeek => "next_week",
        PivotDatePeriod::ThisWeek => "this_week",
        PivotDatePeriod::LastWeek => "last_week",
        PivotDatePeriod::NextMonth => "next_month",
        PivotDatePeriod::ThisMonth => "this_month",
        PivotDatePeriod::LastMonth => "last_month",
        PivotDatePeriod::NextQuarter => "next_quarter",
        PivotDatePeriod::ThisQuarter => "this_quarter",
        PivotDatePeriod::LastQuarter => "last_quarter",
        PivotDatePeriod::NextYear => "next_year",
        PivotDatePeriod::ThisYear => "this_year",
        PivotDatePeriod::LastYear => "last_year",
        PivotDatePeriod::YearToDate => "year_to_date",
        PivotDatePeriod::Quarter(1) => "Q1",
        PivotDatePeriod::Quarter(2) => "Q2",
        PivotDatePeriod::Quarter(3) => "Q3",
        PivotDatePeriod::Quarter(4) => "Q4",
        PivotDatePeriod::Month(1) => "M1",
        PivotDatePeriod::Month(2) => "M2",
        PivotDatePeriod::Month(3) => "M3",
        PivotDatePeriod::Month(4) => "M4",
        PivotDatePeriod::Month(5) => "M5",
        PivotDatePeriod::Month(6) => "M6",
        PivotDatePeriod::Month(7) => "M7",
        PivotDatePeriod::Month(8) => "M8",
        PivotDatePeriod::Month(9) => "M9",
        PivotDatePeriod::Month(10) => "M10",
        PivotDatePeriod::Month(11) => "M11",
        PivotDatePeriod::Month(12) => "M12",
        PivotDatePeriod::Month(_) | PivotDatePeriod::Quarter(_) => "unknown",
    }
}

fn pivot_calculated_fields_to_py(
    py: Python<'_>,
    fields: &[PivotCalculatedField],
) -> PyResult<PyObject> {
    let items = PyList::empty_bound(py);
    for field in fields {
        let dict = PyDict::new_bound(py);
        dict.set_item("name", &field.name)?;
        dict.set_item("formula", &field.formula)?;
        items.append(dict)?;
    }
    Ok(items.into_any().unbind())
}

fn pivot_calculated_items_to_py(
    py: Python<'_>,
    calculated_items: &[PivotCalculatedItem],
) -> PyResult<PyObject> {
    let items = PyList::empty_bound(py);
    for item in calculated_items {
        let dict = PyDict::new_bound(py);
        dict.set_item("field", &item.field.name)?;
        dict.set_item("item", pivot_value_to_py(py, &item.item)?)?;
        dict.set_item("formula", &item.formula)?;
        items.append(dict)?;
    }
    Ok(items.into_any().unbind())
}

fn pivot_groupings_to_py(py: Python<'_>, groupings: &[PivotGrouping]) -> PyResult<PyObject> {
    let items = PyList::empty_bound(py);
    for grouping in groupings {
        items.append(pivot_grouping_to_py(py, grouping)?)?;
    }
    Ok(items.into_any().unbind())
}

fn pivot_grouping_to_py(py: Python<'_>, grouping: &PivotGrouping) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    match grouping {
        PivotGrouping::Number {
            field,
            start,
            end,
            interval,
        } => {
            dict.set_item("kind", "number")?;
            dict.set_item("field", &field.name)?;
            dict.set_item("start", start)?;
            dict.set_item("end", end)?;
            dict.set_item("interval", interval)?;
        }
        PivotGrouping::Date { field, units } => {
            dict.set_item("kind", "date")?;
            dict.set_item("field", &field.name)?;
            let unit_names = units
                .iter()
                .map(|unit| pivot_date_group_unit_to_python(*unit))
                .collect::<Vec<_>>();
            dict.set_item("units", unit_names)?;
        }
        PivotGrouping::Manual { field, groups } => {
            dict.set_item("kind", "manual")?;
            dict.set_item("field", &field.name)?;
            let items = PyList::empty_bound(py);
            for group in groups {
                items.append(pivot_manual_group_to_py(py, group)?)?;
            }
            dict.set_item("groups", items)?;
        }
    }
    Ok(dict.into_any().unbind())
}

fn pivot_manual_group_to_py(py: Python<'_>, group: &PivotManualGroup) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("name", &group.name)?;
    let members = PyList::empty_bound(py);
    for member in &group.members {
        members.append(pivot_value_to_py(py, member)?)?;
    }
    dict.set_item("members", members)?;
    Ok(dict.into_any().unbind())
}

fn pivot_layout_to_py(py: Python<'_>, layout: &PivotLayout) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("kind", pivot_layout_kind_to_python(layout.kind))?;
    dict.set_item("show_row_grand_totals", layout.show_row_grand_totals)?;
    dict.set_item("show_column_grand_totals", layout.show_column_grand_totals)?;
    dict.set_item("show_field_headers", layout.show_field_headers)?;
    dict.set_item("repeat_item_labels", layout.repeat_item_labels)?;
    dict.set_item("show_expand_collapse", layout.show_expand_collapse)?;
    dict.set_item("print_drill_indicators", layout.print_drill_indicators)?;
    dict.set_item("item_print_titles", layout.item_print_titles)?;
    dict.set_item("field_print_titles", layout.field_print_titles)?;
    dict.set_item("page_wrap", layout.page_wrap)?;
    dict.set_item("page_over_then_down", layout.page_over_then_down)?;
    dict.set_item("merge_item_labels", layout.merge_item_labels)?;
    dict.set_item("data_caption", &layout.data_caption)?;
    dict.set_item(
        "values_axis",
        pivot_values_axis_to_python(layout.values_axis),
    )?;
    dict.set_item("values_axis_position", layout.values_axis_position)?;
    dict.set_item("grand_total_caption", &layout.grand_total_caption)?;
    dict.set_item("error_caption", &layout.error_caption)?;
    dict.set_item("show_error", layout.show_error)?;
    dict.set_item("missing_caption", &layout.missing_caption)?;
    dict.set_item("show_missing", layout.show_missing)?;
    dict.set_item("asterisk_totals", layout.asterisk_totals)?;
    dict.set_item("show_items", layout.show_items)?;
    dict.set_item("edit_data", layout.edit_data)?;
    dict.set_item("disable_field_list", layout.disable_field_list)?;
    dict.set_item("show_calculated_members", layout.show_calculated_members)?;
    dict.set_item("visual_totals", layout.visual_totals)?;
    dict.set_item("show_multiple_label", layout.show_multiple_label)?;
    dict.set_item("show_data_drop_down", layout.show_data_drop_down)?;
    dict.set_item(
        "show_member_property_tips",
        layout.show_member_property_tips,
    )?;
    dict.set_item("show_data_tips", layout.show_data_tips)?;
    dict.set_item("enable_wizard", layout.enable_wizard)?;
    dict.set_item("enable_drill", layout.enable_drill)?;
    dict.set_item("enable_field_properties", layout.enable_field_properties)?;
    dict.set_item("subtotal_hidden_items", layout.subtotal_hidden_items)?;
    dict.set_item("show_drop_zones", layout.show_drop_zones)?;
    dict.set_item("indent", layout.indent)?;
    dict.set_item("show_empty_rows", layout.show_empty_rows)?;
    dict.set_item("show_empty_columns", layout.show_empty_columns)?;
    Ok(dict.into_any().unbind())
}

fn pivot_style_to_py(py: Python<'_>, style: &PivotStyle) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("name", &style.name)?;
    dict.set_item("show_row_headers", style.show_row_headers)?;
    dict.set_item("show_column_headers", style.show_column_headers)?;
    dict.set_item("show_row_stripes", style.show_row_stripes)?;
    dict.set_item("show_column_stripes", style.show_column_stripes)?;
    dict.set_item("show_last_column", style.show_last_column)?;
    Ok(dict.into_any().unbind())
}

fn pivot_refresh_policy_to_py(py: Python<'_>, policy: &PivotRefreshPolicy) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("refresh_on_open", policy.refresh_on_open)?;
    dict.set_item("preserve_formatting", policy.preserve_formatting)?;
    dict.set_item("background_query", policy.background_query)?;
    dict.set_item("missing_items_limit", policy.missing_items_limit)?;
    Ok(dict.into_any().unbind())
}

fn pivot_refresh_status_to_py(py: Python<'_>, status: &PivotRefreshStatus) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    match status {
        PivotRefreshStatus::NotRefreshed => {
            dict.set_item("kind", "not_refreshed")?;
        }
        PivotRefreshStatus::Succeeded => {
            dict.set_item("kind", "succeeded")?;
        }
        PivotRefreshStatus::Failed { message } => {
            dict.set_item("kind", "failed")?;
            dict.set_item("message", message)?;
        }
        PivotRefreshStatus::External => {
            dict.set_item("kind", "external")?;
        }
    }
    Ok(dict.into_any().unbind())
}

fn pivot_sort_to_python(sort: PivotSort) -> &'static str {
    match sort {
        PivotSort::None => "none",
        PivotSort::Ascending => "ascending",
        PivotSort::Descending => "descending",
    }
}

fn pivot_subtotal_to_python(subtotal: PivotSubtotal) -> &'static str {
    match subtotal {
        PivotSubtotal::Automatic => "automatic",
        PivotSubtotal::None => "none",
        PivotSubtotal::Sum => "sum",
        PivotSubtotal::Count => "count",
        PivotSubtotal::CountNumbers => "count_numbers",
        PivotSubtotal::Average => "average",
        PivotSubtotal::Min => "min",
        PivotSubtotal::Max => "max",
        PivotSubtotal::Product => "product",
        PivotSubtotal::StdDev => "std_dev",
        PivotSubtotal::StdDevP => "std_dev_p",
        PivotSubtotal::Var => "var",
        PivotSubtotal::VarP => "var_p",
    }
}

fn pivot_aggregate_to_python(aggregate: PivotAggregate) -> &'static str {
    match aggregate {
        PivotAggregate::Sum => "sum",
        PivotAggregate::Count => "count",
        PivotAggregate::CountNumbers => "count_numbers",
        PivotAggregate::Average => "average",
        PivotAggregate::Max => "max",
        PivotAggregate::Min => "min",
        PivotAggregate::Product => "product",
        PivotAggregate::StdDev => "std_dev",
        PivotAggregate::StdDevP => "std_dev_p",
        PivotAggregate::Var => "var",
        PivotAggregate::VarP => "var_p",
    }
}

fn pivot_filter_operator_to_python(operator: PivotFilterOperator) -> &'static str {
    match operator {
        PivotFilterOperator::Equals => "equals",
        PivotFilterOperator::NotEquals => "not_equals",
        PivotFilterOperator::LessThan => "less_than",
        PivotFilterOperator::LessThanOrEqual => "less_than_or_equal",
        PivotFilterOperator::GreaterThan => "greater_than",
        PivotFilterOperator::GreaterThanOrEqual => "greater_than_or_equal",
        PivotFilterOperator::BeginsWith => "begins_with",
        PivotFilterOperator::DoesNotBeginWith => "does_not_begin_with",
        PivotFilterOperator::EndsWith => "ends_with",
        PivotFilterOperator::DoesNotEndWith => "does_not_end_with",
        PivotFilterOperator::Contains => "contains",
        PivotFilterOperator::DoesNotContain => "does_not_contain",
    }
}

fn pivot_date_group_unit_to_python(unit: PivotDateGroupUnit) -> &'static str {
    match unit {
        PivotDateGroupUnit::Seconds => "seconds",
        PivotDateGroupUnit::Minutes => "minutes",
        PivotDateGroupUnit::Hours => "hours",
        PivotDateGroupUnit::Days => "days",
        PivotDateGroupUnit::Months => "months",
        PivotDateGroupUnit::Quarters => "quarters",
        PivotDateGroupUnit::Years => "years",
    }
}

fn pivot_layout_kind_to_python(kind: PivotLayoutKind) -> &'static str {
    match kind {
        PivotLayoutKind::Compact => "compact",
        PivotLayoutKind::Outline => "outline",
        PivotLayoutKind::Tabular => "tabular",
    }
}

fn pivot_values_axis_to_python(axis: PivotValuesAxis) -> &'static str {
    match axis {
        PivotValuesAxis::Columns => "columns",
        PivotValuesAxis::Rows => "rows",
    }
}

fn pivot_overwrite_policy_to_python(policy: PivotOverwritePolicy) -> &'static str {
    match policy {
        PivotOverwritePolicy::ClearOwnedRange => "clear_owned_range",
        PivotOverwritePolicy::Overwrite => "overwrite",
        PivotOverwritePolicy::FailOnOccupied => "fail_on_occupied",
    }
}

fn workbook_connection_to_py(
    py: Python<'_>,
    connection: &WorkbookConnection,
) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("id", connection.id)?;
    dict.set_item("name", &connection.name)?;
    dict.set_item("kind", workbook_connection_kind_to_python(&connection.kind))?;
    dict.set_item("refreshed_version", connection.refreshed_version)?;
    dict.set_item("refresh_on_load", connection.refresh_on_load)?;
    dict.set_item("background", connection.background)?;
    dict.set_item("save_data", connection.save_data)?;
    match &connection.kind {
        WorkbookConnectionKind::Database {
            connection,
            command,
            command_type,
        } => {
            dict.set_item("connection", connection)?;
            dict.set_item("command", command)?;
            dict.set_item("command_type", command_type)?;
        }
        WorkbookConnectionKind::Olap {
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        } => {
            dict.set_item("local", local)?;
            dict.set_item("local_connection", local_connection)?;
            dict.set_item("local_refresh", local_refresh)?;
            dict.set_item("send_locale", send_locale)?;
            dict.set_item("row_drill_count", row_drill_count)?;
        }
        WorkbookConnectionKind::Web {
            url,
            xml,
            source_data,
            html_tables,
            html_format,
            post,
            edit_page,
        } => {
            dict.set_item("url", url)?;
            dict.set_item("xml", xml)?;
            dict.set_item("source_data", source_data)?;
            dict.set_item("html_tables", html_tables)?;
            dict.set_item("html_format", html_format)?;
            dict.set_item("post", post)?;
            dict.set_item("edit_page", edit_page)?;
        }
        WorkbookConnectionKind::Text {
            source_file,
            delimiter,
            first_row,
            delimited,
            decimal,
            thousands,
        } => {
            dict.set_item("source_file", source_file)?;
            dict.set_item("delimiter", delimiter)?;
            dict.set_item("first_row", first_row)?;
            dict.set_item("delimited", delimited)?;
            dict.set_item("decimal", decimal)?;
            dict.set_item("thousands", thousands)?;
        }
    }
    Ok(dict.into_any().unbind())
}

fn workbook_connection_kind_to_python(kind: &WorkbookConnectionKind) -> &'static str {
    match kind {
        WorkbookConnectionKind::Database { .. } => "database",
        WorkbookConnectionKind::Olap { .. } => "olap",
        WorkbookConnectionKind::Web { .. } => "web",
        WorkbookConnectionKind::Text { .. } => "text",
    }
}

fn workbook_extension_to_py(py: Python<'_>, extension: &WorkbookExtension) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("uri", &extension.uri)?;
    dict.set_item("payload", PyBytes::new_bound(py, &extension.payload))?;
    Ok(dict.into_any().unbind())
}

fn workbook_extension_part_to_py(
    py: Python<'_>,
    part: &WorkbookExtensionPart,
) -> PyResult<PyObject> {
    let dict = PyDict::new_bound(py);
    dict.set_item("path", &part.path)?;
    dict.set_item("content_type", &part.content_type)?;
    dict.set_item("relationship_type", &part.relationship_type)?;
    dict.set_item("relationship_id", &part.relationship_id)?;
    dict.set_item("payload", PyBytes::new_bound(py, &part.payload))?;
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

    /// Pivot table definitions on the worksheet.
    #[getter]
    fn pivot_tables(&self, py: Python<'_>) -> PyResult<Vec<PyObject>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        ws.pivot_tables()
            .iter()
            .map(|pivot| pivot_table_to_py(py, pivot))
            .collect()
    }

    /// Get a pivot table definition by name.
    #[pyo3(signature = (name))]
    fn get_pivot_table(&self, name: &str, py: Python<'_>) -> PyResult<Option<PyObject>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        ws.pivot_table_by_name(name)
            .map(|pivot| pivot_table_to_py(py, pivot))
            .transpose()
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

    /// Generate and add a PivotChart from a rendered pivot table.
    #[pyo3(signature = (pivot_name, chart_type=None))]
    fn add_pivot_chart(&self, pivot_name: &str, chart_type: Option<&str>) -> PyResult<PyChart> {
        let chart_type = parse_chart_type(chart_type)?;
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let chart = ws
            .build_pivot_chart(
                pivot_name,
                chart_type,
                duke_sheets::DrawingAnchor::default(),
            )
            .map_err(to_py_err)?;
        ws.add_chart(chart.clone());
        Ok(PyChart::from(&chart))
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
    ///
    /// Args:
    ///     max_threads: Maximum worker threads for parallel refresh. None uses the active pool.
    ///     today: Excel serial date used to evaluate relative date-period filters.
    #[pyo3(signature = (*, max_threads=None, today=None))]
    fn refresh_pivots(
        &self,
        py: Python<'_>,
        max_threads: Option<usize>,
        today: Option<f64>,
    ) -> PyResult<PyObject> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        let stats = wb
            .refresh_pivots_with_options(&PivotRefreshOptions { max_threads, today })
            .map_err(to_py_err)?;
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

    /// Workbook-level data connection definitions.
    #[getter]
    fn data_connections(&self, py: Python<'_>) -> PyResult<Vec<PyObject>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        wb.data_connections()
            .iter()
            .map(|connection| workbook_connection_to_py(py, connection))
            .collect()
    }

    /// Number of raw workbook extension elements preserved from the package.
    #[getter]
    fn workbook_extension_count(&self) -> PyResult<usize> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.workbook_extensions().len())
    }

    /// Raw workbook extension elements preserved from workbook.xml.
    #[getter]
    fn workbook_extensions(&self, py: Python<'_>) -> PyResult<Vec<PyObject>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        wb.workbook_extensions()
            .iter()
            .map(|extension| workbook_extension_to_py(py, extension))
            .collect()
    }

    /// Number of raw workbook-related extension package parts.
    #[getter]
    fn workbook_extension_part_count(&self) -> PyResult<usize> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.workbook_extension_parts().len())
    }

    /// Raw workbook-related extension package parts.
    #[getter]
    fn workbook_extension_parts(&self, py: Python<'_>) -> PyResult<Vec<PyObject>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        wb.workbook_extension_parts()
            .iter()
            .map(|part| workbook_extension_part_to_py(py, part))
            .collect()
    }

    /// Get a raw workbook extension package part by package path.
    #[pyo3(signature = (path))]
    fn get_workbook_extension_part(
        &self,
        path: &str,
        py: Python<'_>,
    ) -> PyResult<Option<PyObject>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        wb.workbook_extension_parts()
            .iter()
            .find(|part| part.path == path)
            .map(|part| workbook_extension_part_to_py(py, part))
            .transpose()
    }

    /// Get a raw workbook extension package part by workbook relationship id.
    #[pyo3(signature = (relationship_id))]
    fn get_workbook_extension_part_by_relationship_id(
        &self,
        relationship_id: &str,
        py: Python<'_>,
    ) -> PyResult<Option<PyObject>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        wb.workbook_extension_parts()
            .iter()
            .find(|part| part.relationship_id.as_deref() == Some(relationship_id))
            .map(|part| workbook_extension_part_to_py(py, part))
            .transpose()
    }

    /// Get a workbook-level data connection by name.
    #[pyo3(signature = (name))]
    fn get_data_connection(&self, name: &str, py: Python<'_>) -> PyResult<Option<PyObject>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        wb.data_connection_by_name(name)
            .map(|connection| workbook_connection_to_py(py, connection))
            .transpose()
    }

    /// Get a workbook-level data connection by id.
    #[pyo3(signature = (id))]
    fn get_data_connection_by_id(&self, id: u32, py: Python<'_>) -> PyResult<Option<PyObject>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        wb.data_connection_by_id(id)
            .map(|connection| workbook_connection_to_py(py, connection))
            .transpose()
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
