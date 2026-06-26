//! NAPI object types for the Node.js binding.
//!
//! These are plain JS objects (`#[napi(object)]`) used as return types
//! from the read-only API. Each has a `From` impl to convert from the
//! corresponding core Rust type.

use napi::{Error as NapiError, Result as NapiResult};
use napi_derive::napi;

use duke_sheets_core::{
    self as core,
    style::{
        Alignment as CoreAlignment, BorderEdge as CoreBorderEdge,
        BorderLineStyle as CoreBorderLineStyle, BorderStyle as CoreBorderStyle, Color as CoreColor,
        DiagonalDirection, FillStyle as CoreFillStyle, FontStyle as CoreFontStyle,
        FontVerticalAlign, GradientType, HorizontalAlignment, NumberFormat as CoreNumberFormat,
        PatternType, ReadingOrder, Style as CoreStyle, Underline, VerticalAlignment,
    },
};

/// A single cell within a sparse row.
#[napi(object)]
pub struct JsRowCell {
    /// Column index (0-based).
    pub col: u32,
    /// String representation of the cell value.
    pub value: String,
    /// Cell style (when includeStyles is set).
    pub style: Option<JsStyle>,
    /// Merge span for merge-origin cells (when includeMergeInfo is set).
    pub merge_span: Option<JsMergeSpan>,
    /// Whether this cell is a non-origin member of a merge (when includeMergeInfo is set).
    pub is_merged_secondary: Option<bool>,
    /// Hyperlink (when includeHyperlinks is set).
    pub hyperlink: Option<JsHyperlink>,
    /// Comment (when includeComments is set).
    pub comment: Option<JsComment>,
    /// Formula text (when includeFormulas is set).
    pub formula: Option<String>,
    /// IMAGE() metadata (when includeImages is set).
    pub image: Option<JsImageInfo>,
}

/// A sparse row containing only non-empty cells.
#[napi(object)]
pub struct JsRow {
    /// Row index (0-based).
    pub index: u32,
    /// Non-empty cells in this row, sorted by column.
    pub cells: Vec<JsRowCell>,
}

/// Options for row iteration.
#[napi(object)]
pub struct JsRowsOptions {
    /// Use display-formatted values (e.g., "$1,234.56") instead of raw values.
    pub use_formatted_values: Option<bool>,
    /// Use calculated values for formula cells (requires prior calculate() call).
    pub use_calculated_values: Option<bool>,
    /// Include cell styles.
    pub include_styles: Option<bool>,
    /// Include merge info (mergeSpan + isMergedSecondary).
    pub include_merge_info: Option<bool>,
    /// Include hyperlinks.
    pub include_hyperlinks: Option<bool>,
    /// Include comments.
    pub include_comments: Option<bool>,
    /// Include formula text.
    pub include_formulas: Option<bool>,
    /// Include IMAGE() metadata.
    pub include_images: Option<bool>,
    /// Skip cells whose raw value is empty (CellValue::Empty).
    pub skip_empty_values: Option<bool>,
    /// Skip cells whose raw value is blank (empty, empty string, or empty rich text).
    pub skip_blank_values: Option<bool>,
}

#[napi(object)]
pub struct JsPivotValue {
    pub kind: String,
    pub number: Option<f64>,
    pub text: Option<String>,
    pub boolean: Option<bool>,
    pub error: Option<String>,
}

impl From<&core::PivotValue> for JsPivotValue {
    fn from(value: &core::PivotValue) -> Self {
        match value {
            core::PivotValue::Blank => Self {
                kind: "blank".into(),
                number: None,
                text: None,
                boolean: None,
                error: None,
            },
            core::PivotValue::Boolean(value) => Self {
                kind: "boolean".into(),
                number: None,
                text: None,
                boolean: Some(*value),
                error: None,
            },
            core::PivotValue::Number(value) => Self {
                kind: "number".into(),
                number: Some(*value),
                text: None,
                boolean: None,
                error: None,
            },
            core::PivotValue::String(value) => Self {
                kind: "string".into(),
                number: None,
                text: Some(value.clone()),
                boolean: None,
                error: None,
            },
            core::PivotValue::Error(value) => Self {
                kind: "error".into(),
                number: None,
                text: None,
                boolean: None,
                error: Some(value.to_string()),
            },
        }
    }
}

#[napi(object)]
pub struct JsPivotSourceRangeDefinition {
    pub sheet: String,
    pub range: String,
    pub name: Option<String>,
    pub page_items: Vec<String>,
}

impl From<&core::PivotSourceRange> for JsPivotSourceRangeDefinition {
    fn from(range: &core::PivotSourceRange) -> Self {
        Self {
            sheet: range.sheet.clone(),
            range: range.range.to_string(),
            name: range.name.clone(),
            page_items: range.page_items.clone(),
        }
    }
}

#[napi(object)]
pub struct JsPivotSourceDefinition {
    pub kind: String,
    pub sheet: Option<String>,
    pub range: Option<String>,
    pub table_name: Option<String>,
    pub connection_name: Option<String>,
    pub command_text: Option<String>,
    pub ranges: Option<Vec<JsPivotSourceRangeDefinition>>,
    pub scenario_name: Option<String>,
    pub cube: Option<String>,
}

impl From<&core::PivotSource> for JsPivotSourceDefinition {
    fn from(source: &core::PivotSource) -> Self {
        match source {
            core::PivotSource::WorksheetRange { sheet, range } => Self {
                kind: "worksheetRange".into(),
                sheet: sheet.clone(),
                range: Some(range.to_string()),
                table_name: None,
                connection_name: None,
                command_text: None,
                ranges: None,
                scenario_name: None,
                cube: None,
            },
            core::PivotSource::Table { name } => Self {
                kind: "table".into(),
                sheet: None,
                range: None,
                table_name: Some(name.clone()),
                connection_name: None,
                command_text: None,
                ranges: None,
                scenario_name: None,
                cube: None,
            },
            core::PivotSource::External {
                connection_name,
                command_text,
            } => Self {
                kind: "external".into(),
                sheet: None,
                range: None,
                table_name: None,
                connection_name: Some(connection_name.clone()),
                command_text: command_text.clone(),
                ranges: None,
                scenario_name: None,
                cube: None,
            },
            core::PivotSource::Consolidation { ranges } => Self {
                kind: "consolidation".into(),
                sheet: None,
                range: None,
                table_name: None,
                connection_name: None,
                command_text: None,
                ranges: Some(ranges.iter().map(Into::into).collect()),
                scenario_name: None,
                cube: None,
            },
            core::PivotSource::Scenario { name } => Self {
                kind: "scenario".into(),
                sheet: None,
                range: None,
                table_name: None,
                connection_name: None,
                command_text: None,
                ranges: None,
                scenario_name: Some(name.clone()),
                cube: None,
            },
            core::PivotSource::Olap {
                connection_name,
                cube,
                command_text,
            } => Self {
                kind: "olap".into(),
                sheet: None,
                range: None,
                table_name: None,
                connection_name: Some(connection_name.clone()),
                command_text: command_text.clone(),
                ranges: None,
                scenario_name: None,
                cube: cube.clone(),
            },
        }
    }
}

#[napi(object)]
pub struct JsPivotFieldDefinition {
    pub field: String,
    pub caption: Option<String>,
    pub sort: String,
    pub subtotal: String,
    pub subtotal_caption: Option<String>,
    pub subtotals: Vec<String>,
    pub collapsed_items: Vec<JsPivotValue>,
    pub show_empty_items: bool,
    pub show_drop_downs: bool,
    pub subtotal_top: bool,
    pub insert_blank_row: bool,
    pub insert_page_break: bool,
    pub include_new_items_in_filter: bool,
    pub item_page_count: u32,
}

impl From<&core::PivotField> for JsPivotFieldDefinition {
    fn from(field: &core::PivotField) -> Self {
        Self {
            field: field.field.name.clone(),
            caption: field.caption.clone(),
            sort: pivot_sort_to_string(field.sort).into(),
            subtotal: pivot_subtotal_to_string(field.subtotal).into(),
            subtotal_caption: field.subtotal_caption.clone(),
            subtotals: field
                .subtotals
                .iter()
                .map(|subtotal| pivot_subtotal_to_string(*subtotal).into())
                .collect(),
            collapsed_items: field
                .collapsed_items
                .iter()
                .map(JsPivotValue::from)
                .collect(),
            show_empty_items: field.show_empty_items,
            show_drop_downs: field.show_drop_downs,
            subtotal_top: field.subtotal_top,
            insert_blank_row: field.insert_blank_row,
            insert_page_break: field.insert_page_break,
            include_new_items_in_filter: field.include_new_items_in_filter,
            item_page_count: field.item_page_count,
        }
    }
}

#[napi(object)]
pub struct JsPivotShowAsDefinition {
    pub kind: String,
    pub base_field: Option<String>,
    pub base_item: Option<JsPivotValue>,
}

impl From<&core::PivotShowAs> for JsPivotShowAsDefinition {
    fn from(show_as: &core::PivotShowAs) -> Self {
        match show_as {
            core::PivotShowAs::Normal => Self {
                kind: "normal".into(),
                base_field: None,
                base_item: None,
            },
            core::PivotShowAs::PercentOfGrandTotal => Self {
                kind: "percentOfGrandTotal".into(),
                base_field: None,
                base_item: None,
            },
            core::PivotShowAs::PercentOfRowTotal => Self {
                kind: "percentOfRowTotal".into(),
                base_field: None,
                base_item: None,
            },
            core::PivotShowAs::PercentOfColumnTotal => Self {
                kind: "percentOfColumnTotal".into(),
                base_field: None,
                base_item: None,
            },
            core::PivotShowAs::PercentOfParentRowTotal => Self {
                kind: "percentOfParentRowTotal".into(),
                base_field: None,
                base_item: None,
            },
            core::PivotShowAs::PercentOfParentColumnTotal => Self {
                kind: "percentOfParentColumnTotal".into(),
                base_field: None,
                base_item: None,
            },
            core::PivotShowAs::PercentOfParentTotal { base_field } => Self {
                kind: "percentOfParentTotal".into(),
                base_field: Some(base_field.name.clone()),
                base_item: None,
            },
            core::PivotShowAs::Index => Self {
                kind: "index".into(),
                base_field: None,
                base_item: None,
            },
            core::PivotShowAs::RunningTotal { base_field } => Self {
                kind: "runningTotal".into(),
                base_field: Some(base_field.name.clone()),
                base_item: None,
            },
            core::PivotShowAs::DifferenceFrom {
                base_field,
                base_item,
            } => Self {
                kind: "differenceFrom".into(),
                base_field: Some(base_field.name.clone()),
                base_item: Some(base_item.into()),
            },
            core::PivotShowAs::PercentDifferenceFrom {
                base_field,
                base_item,
            } => Self {
                kind: "percentDifferenceFrom".into(),
                base_field: Some(base_field.name.clone()),
                base_item: Some(base_item.into()),
            },
            core::PivotShowAs::RankAscending { base_field } => Self {
                kind: "rankAscending".into(),
                base_field: Some(base_field.name.clone()),
                base_item: None,
            },
            core::PivotShowAs::RankDescending { base_field } => Self {
                kind: "rankDescending".into(),
                base_field: Some(base_field.name.clone()),
                base_item: None,
            },
        }
    }
}

#[napi(object)]
pub struct JsPivotMeasureDefinition {
    pub field: String,
    pub aggregate: String,
    pub name: Option<String>,
    pub caption: String,
    pub show_as: JsPivotShowAsDefinition,
    pub number_format: Option<String>,
}

impl From<&core::PivotMeasure> for JsPivotMeasureDefinition {
    fn from(measure: &core::PivotMeasure) -> Self {
        Self {
            field: measure.field.name.clone(),
            aggregate: pivot_aggregate_to_string(measure.aggregate).into(),
            name: measure.name.clone(),
            caption: measure.caption(),
            show_as: JsPivotShowAsDefinition::from(&measure.show_as),
            number_format: measure.number_format.clone(),
        }
    }
}

#[napi(object)]
pub struct JsPivotFilterDefinition {
    pub kind: String,
    pub field: Option<String>,
    pub items: Option<Vec<JsPivotValue>>,
    pub operator: Option<String>,
    pub text: Option<String>,
    pub start_text: Option<String>,
    pub end_text: Option<String>,
    pub period: Option<String>,
    pub measure: Option<JsPivotMeasureDefinition>,
    pub value: Option<f64>,
    pub start: Option<f64>,
    pub end: Option<f64>,
    pub n: Option<u32>,
    pub top: Option<bool>,
    pub percent: Option<bool>,
    pub detail: Option<String>,
}

impl From<&core::PivotFilter> for JsPivotFilterDefinition {
    fn from(filter: &core::PivotFilter) -> Self {
        match filter {
            core::PivotFilter::FieldItems {
                field,
                allowed_items,
            } => Self {
                kind: "fieldItems".into(),
                field: Some(field.name.clone()),
                items: Some(allowed_items.iter().map(Into::into).collect()),
                operator: None,
                text: None,
                start_text: None,
                end_text: None,
                period: None,
                measure: None,
                value: None,
                start: None,
                end: None,
                n: None,
                top: None,
                percent: None,
                detail: None,
            },
            core::PivotFilter::Label {
                field,
                operator,
                value,
            } => Self {
                kind: "label".into(),
                field: Some(field.name.clone()),
                items: None,
                operator: Some(pivot_filter_operator_to_string(*operator).into()),
                text: Some(value.clone()),
                start_text: None,
                end_text: None,
                period: None,
                measure: None,
                value: None,
                start: None,
                end: None,
                n: None,
                top: None,
                percent: None,
                detail: None,
            },
            core::PivotFilter::LabelBetween {
                field,
                start,
                end,
                not_between,
            } => Self {
                kind: if *not_between {
                    "labelNotBetween".into()
                } else {
                    "labelBetween".into()
                },
                field: Some(field.name.clone()),
                items: None,
                operator: None,
                text: None,
                start_text: Some(start.clone()),
                end_text: Some(end.clone()),
                period: None,
                measure: None,
                value: None,
                start: None,
                end: None,
                n: None,
                top: None,
                percent: None,
                detail: None,
            },
            core::PivotFilter::Date {
                field,
                operator,
                value,
            } => Self {
                kind: "date".into(),
                field: Some(field.name.clone()),
                items: None,
                operator: Some(pivot_filter_operator_to_string(*operator).into()),
                text: None,
                start_text: None,
                end_text: None,
                period: None,
                measure: None,
                value: Some(*value),
                start: None,
                end: None,
                n: None,
                top: None,
                percent: None,
                detail: None,
            },
            core::PivotFilter::DateBetween {
                field,
                start,
                end,
                not_between,
            } => Self {
                kind: if *not_between {
                    "dateNotBetween".into()
                } else {
                    "dateBetween".into()
                },
                field: Some(field.name.clone()),
                items: None,
                operator: None,
                text: None,
                start_text: None,
                end_text: None,
                period: None,
                measure: None,
                value: None,
                start: Some(*start),
                end: Some(*end),
                n: None,
                top: None,
                percent: None,
                detail: None,
            },
            core::PivotFilter::DatePeriod { field, period } => Self {
                kind: "datePeriod".into(),
                field: Some(field.name.clone()),
                items: None,
                operator: None,
                text: None,
                start_text: None,
                end_text: None,
                period: Some(pivot_date_period_to_string(*period).into()),
                measure: None,
                value: None,
                start: None,
                end: None,
                n: None,
                top: None,
                percent: None,
                detail: None,
            },
            core::PivotFilter::Value {
                field,
                measure,
                operator,
                value,
            } => Self {
                kind: "value".into(),
                field: Some(field.name.clone()),
                items: None,
                operator: Some(pivot_filter_operator_to_string(*operator).into()),
                text: None,
                start_text: None,
                end_text: None,
                period: None,
                measure: Some(measure.into()),
                value: Some(*value),
                start: None,
                end: None,
                n: None,
                top: None,
                percent: None,
                detail: None,
            },
            core::PivotFilter::ValueBetween {
                field,
                measure,
                start,
                end,
                not_between,
            } => Self {
                kind: if *not_between {
                    "valueNotBetween".into()
                } else {
                    "valueBetween".into()
                },
                field: Some(field.name.clone()),
                items: None,
                operator: None,
                text: None,
                start_text: None,
                end_text: None,
                period: None,
                measure: Some(measure.into()),
                value: None,
                start: Some(*start),
                end: Some(*end),
                n: None,
                top: None,
                percent: None,
                detail: None,
            },
            core::PivotFilter::TopN {
                field,
                measure,
                n,
                top,
                percent,
            } => Self {
                kind: "topN".into(),
                field: Some(field.name.clone()),
                items: None,
                operator: None,
                text: None,
                start_text: None,
                end_text: None,
                period: None,
                measure: Some(measure.into()),
                value: None,
                start: None,
                end: None,
                n: Some(*n),
                top: Some(*top),
                percent: Some(*percent),
                detail: None,
            },
            core::PivotFilter::Unsupported { kind, detail } => Self {
                kind: kind.clone(),
                field: None,
                items: None,
                operator: None,
                text: None,
                start_text: None,
                end_text: None,
                period: None,
                measure: None,
                value: None,
                start: None,
                end: None,
                n: None,
                top: None,
                percent: None,
                detail: detail.clone(),
            },
        }
    }
}

#[napi(object)]
pub struct JsPivotCalculatedFieldDefinition {
    pub name: String,
    pub formula: String,
}

impl From<&core::PivotCalculatedField> for JsPivotCalculatedFieldDefinition {
    fn from(field: &core::PivotCalculatedField) -> Self {
        Self {
            name: field.name.clone(),
            formula: field.formula.clone(),
        }
    }
}

#[napi(object)]
pub struct JsPivotCalculatedItemDefinition {
    pub field: String,
    pub item: JsPivotValue,
    pub formula: String,
}

impl From<&core::PivotCalculatedItem> for JsPivotCalculatedItemDefinition {
    fn from(item: &core::PivotCalculatedItem) -> Self {
        Self {
            field: item.field.name.clone(),
            item: JsPivotValue::from(&item.item),
            formula: item.formula.clone(),
        }
    }
}

#[napi(object)]
pub struct JsPivotManualGroupDefinition {
    pub name: String,
    pub members: Vec<JsPivotValue>,
}

impl From<&core::PivotManualGroup> for JsPivotManualGroupDefinition {
    fn from(group: &core::PivotManualGroup) -> Self {
        Self {
            name: group.name.clone(),
            members: group.members.iter().map(Into::into).collect(),
        }
    }
}

#[napi(object)]
pub struct JsPivotGroupingDefinition {
    pub kind: String,
    pub field: String,
    pub start: Option<f64>,
    pub end: Option<f64>,
    pub interval: Option<f64>,
    pub units: Option<Vec<String>>,
    pub groups: Option<Vec<JsPivotManualGroupDefinition>>,
}

impl From<&core::PivotGrouping> for JsPivotGroupingDefinition {
    fn from(grouping: &core::PivotGrouping) -> Self {
        match grouping {
            core::PivotGrouping::Number {
                field,
                start,
                end,
                interval,
            } => Self {
                kind: "number".into(),
                field: field.name.clone(),
                start: *start,
                end: *end,
                interval: Some(*interval),
                units: None,
                groups: None,
            },
            core::PivotGrouping::Date { field, units } => Self {
                kind: "date".into(),
                field: field.name.clone(),
                start: None,
                end: None,
                interval: None,
                units: Some(
                    units
                        .iter()
                        .map(|unit| pivot_date_group_unit_to_string(*unit).into())
                        .collect(),
                ),
                groups: None,
            },
            core::PivotGrouping::Manual { field, groups } => Self {
                kind: "manual".into(),
                field: field.name.clone(),
                start: None,
                end: None,
                interval: None,
                units: None,
                groups: Some(groups.iter().map(Into::into).collect()),
            },
        }
    }
}

#[napi(object)]
pub struct JsPivotLayoutDefinition {
    pub kind: String,
    pub show_row_grand_totals: bool,
    pub show_column_grand_totals: bool,
    pub show_field_headers: bool,
    pub repeat_item_labels: bool,
    pub show_expand_collapse: bool,
    pub print_drill_indicators: bool,
    pub item_print_titles: bool,
    pub field_print_titles: bool,
    pub page_wrap: u32,
    pub page_over_then_down: bool,
    pub merge_item_labels: bool,
    pub data_caption: String,
    pub values_axis: String,
    pub values_axis_position: Option<u32>,
    pub grand_total_caption: Option<String>,
    pub error_caption: Option<String>,
    pub show_error: bool,
    pub missing_caption: Option<String>,
    pub show_missing: bool,
    pub asterisk_totals: bool,
    pub show_items: bool,
    pub edit_data: bool,
    pub disable_field_list: bool,
    pub show_calculated_members: bool,
    pub visual_totals: bool,
    pub show_multiple_label: bool,
    pub show_data_drop_down: bool,
    pub show_member_property_tips: bool,
    pub show_data_tips: bool,
    pub enable_wizard: bool,
    pub enable_drill: bool,
    pub enable_field_properties: bool,
    pub subtotal_hidden_items: bool,
    pub show_drop_zones: bool,
    pub indent: u32,
    pub show_empty_rows: bool,
    pub show_empty_columns: bool,
}

impl From<&core::PivotLayout> for JsPivotLayoutDefinition {
    fn from(layout: &core::PivotLayout) -> Self {
        Self {
            kind: pivot_layout_kind_to_string(layout.kind).into(),
            show_row_grand_totals: layout.show_row_grand_totals,
            show_column_grand_totals: layout.show_column_grand_totals,
            show_field_headers: layout.show_field_headers,
            repeat_item_labels: layout.repeat_item_labels,
            show_expand_collapse: layout.show_expand_collapse,
            print_drill_indicators: layout.print_drill_indicators,
            item_print_titles: layout.item_print_titles,
            field_print_titles: layout.field_print_titles,
            page_wrap: layout.page_wrap,
            page_over_then_down: layout.page_over_then_down,
            merge_item_labels: layout.merge_item_labels,
            data_caption: layout.data_caption.clone(),
            values_axis: pivot_values_axis_to_string(layout.values_axis).into(),
            values_axis_position: layout.values_axis_position,
            grand_total_caption: layout.grand_total_caption.clone(),
            error_caption: layout.error_caption.clone(),
            show_error: layout.show_error,
            missing_caption: layout.missing_caption.clone(),
            show_missing: layout.show_missing,
            asterisk_totals: layout.asterisk_totals,
            show_items: layout.show_items,
            edit_data: layout.edit_data,
            disable_field_list: layout.disable_field_list,
            show_calculated_members: layout.show_calculated_members,
            visual_totals: layout.visual_totals,
            show_multiple_label: layout.show_multiple_label,
            show_data_drop_down: layout.show_data_drop_down,
            show_member_property_tips: layout.show_member_property_tips,
            show_data_tips: layout.show_data_tips,
            enable_wizard: layout.enable_wizard,
            enable_drill: layout.enable_drill,
            enable_field_properties: layout.enable_field_properties,
            subtotal_hidden_items: layout.subtotal_hidden_items,
            show_drop_zones: layout.show_drop_zones,
            indent: layout.indent,
            show_empty_rows: layout.show_empty_rows,
            show_empty_columns: layout.show_empty_columns,
        }
    }
}

#[napi(object)]
pub struct JsPivotStyleDefinition {
    pub name: Option<String>,
    pub show_row_headers: bool,
    pub show_column_headers: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
    pub show_last_column: bool,
}

impl From<&core::PivotStyle> for JsPivotStyleDefinition {
    fn from(style: &core::PivotStyle) -> Self {
        Self {
            name: style.name.clone(),
            show_row_headers: style.show_row_headers,
            show_column_headers: style.show_column_headers,
            show_row_stripes: style.show_row_stripes,
            show_column_stripes: style.show_column_stripes,
            show_last_column: style.show_last_column,
        }
    }
}

#[napi(object)]
pub struct JsPivotRefreshPolicyDefinition {
    pub refresh_on_open: bool,
    pub preserve_formatting: bool,
    pub background_query: bool,
    pub missing_items_limit: Option<u32>,
}

impl From<&core::PivotRefreshPolicy> for JsPivotRefreshPolicyDefinition {
    fn from(policy: &core::PivotRefreshPolicy) -> Self {
        Self {
            refresh_on_open: policy.refresh_on_open,
            preserve_formatting: policy.preserve_formatting,
            background_query: policy.background_query,
            missing_items_limit: policy.missing_items_limit,
        }
    }
}

#[napi(object)]
pub struct JsPivotRefreshStatusDefinition {
    pub kind: String,
    pub message: Option<String>,
}

impl From<&core::PivotRefreshStatus> for JsPivotRefreshStatusDefinition {
    fn from(status: &core::PivotRefreshStatus) -> Self {
        match status {
            core::PivotRefreshStatus::NotRefreshed => Self {
                kind: "notRefreshed".into(),
                message: None,
            },
            core::PivotRefreshStatus::Succeeded => Self {
                kind: "succeeded".into(),
                message: None,
            },
            core::PivotRefreshStatus::Failed { message } => Self {
                kind: "failed".into(),
                message: Some(message.clone()),
            },
            core::PivotRefreshStatus::External => Self {
                kind: "external".into(),
                message: None,
            },
        }
    }
}

#[napi(object)]
pub struct JsPivotTableDefinition {
    pub id: u32,
    pub name: String,
    pub source: JsPivotSourceDefinition,
    pub target: String,
    pub rows: Vec<JsPivotFieldDefinition>,
    pub columns: Vec<JsPivotFieldDefinition>,
    pub page_fields: Vec<JsPivotFieldDefinition>,
    pub filters: Vec<JsPivotFilterDefinition>,
    pub calculated_fields: Vec<JsPivotCalculatedFieldDefinition>,
    pub calculated_items: Vec<JsPivotCalculatedItemDefinition>,
    pub measures: Vec<JsPivotMeasureDefinition>,
    pub groupings: Vec<JsPivotGroupingDefinition>,
    pub layout: JsPivotLayoutDefinition,
    pub style: JsPivotStyleDefinition,
    pub refresh_policy: JsPivotRefreshPolicyDefinition,
    pub overwrite_policy: String,
    pub rendered_range: Option<String>,
    pub refresh_status: JsPivotRefreshStatusDefinition,
    pub extension_count: u32,
}

impl From<&core::PivotTable> for JsPivotTableDefinition {
    fn from(pivot: &core::PivotTable) -> Self {
        Self {
            id: pivot.id,
            name: pivot.name.clone(),
            source: JsPivotSourceDefinition::from(&pivot.source),
            target: pivot.target.to_string(),
            rows: pivot.rows.iter().map(Into::into).collect(),
            columns: pivot.columns.iter().map(Into::into).collect(),
            page_fields: pivot.page_fields.iter().map(Into::into).collect(),
            filters: pivot.filters.iter().map(Into::into).collect(),
            calculated_fields: pivot.calculated_fields.iter().map(Into::into).collect(),
            calculated_items: pivot.calculated_items.iter().map(Into::into).collect(),
            measures: pivot.measures.iter().map(Into::into).collect(),
            groupings: pivot.groupings.iter().map(Into::into).collect(),
            layout: JsPivotLayoutDefinition::from(&pivot.layout),
            style: JsPivotStyleDefinition::from(&pivot.style),
            refresh_policy: JsPivotRefreshPolicyDefinition::from(&pivot.refresh_policy),
            overwrite_policy: pivot_overwrite_policy_to_string(pivot.overwrite_policy).into(),
            rendered_range: pivot.rendered_range.map(|range| range.to_string()),
            refresh_status: JsPivotRefreshStatusDefinition::from(&pivot.refresh_status),
            extension_count: pivot.extensions.len() as u32,
        }
    }
}

fn pivot_sort_to_string(sort: core::PivotSort) -> &'static str {
    match sort {
        core::PivotSort::None => "none",
        core::PivotSort::Ascending => "ascending",
        core::PivotSort::Descending => "descending",
    }
}

fn pivot_subtotal_to_string(subtotal: core::PivotSubtotal) -> &'static str {
    match subtotal {
        core::PivotSubtotal::Automatic => "automatic",
        core::PivotSubtotal::None => "none",
        core::PivotSubtotal::Sum => "sum",
        core::PivotSubtotal::Count => "count",
        core::PivotSubtotal::CountNumbers => "countNumbers",
        core::PivotSubtotal::Average => "average",
        core::PivotSubtotal::Min => "min",
        core::PivotSubtotal::Max => "max",
        core::PivotSubtotal::Product => "product",
        core::PivotSubtotal::StdDev => "stdDev",
        core::PivotSubtotal::StdDevP => "stdDevP",
        core::PivotSubtotal::Var => "var",
        core::PivotSubtotal::VarP => "varP",
    }
}

fn pivot_aggregate_to_string(aggregate: core::PivotAggregate) -> &'static str {
    match aggregate {
        core::PivotAggregate::Sum => "sum",
        core::PivotAggregate::Count => "count",
        core::PivotAggregate::CountNumbers => "countNumbers",
        core::PivotAggregate::Average => "average",
        core::PivotAggregate::Max => "max",
        core::PivotAggregate::Min => "min",
        core::PivotAggregate::Product => "product",
        core::PivotAggregate::StdDev => "stdDev",
        core::PivotAggregate::StdDevP => "stdDevP",
        core::PivotAggregate::Var => "var",
        core::PivotAggregate::VarP => "varP",
    }
}

fn pivot_filter_operator_to_string(operator: core::PivotFilterOperator) -> &'static str {
    match operator {
        core::PivotFilterOperator::Equals => "equals",
        core::PivotFilterOperator::NotEquals => "notEquals",
        core::PivotFilterOperator::LessThan => "lessThan",
        core::PivotFilterOperator::LessThanOrEqual => "lessThanOrEqual",
        core::PivotFilterOperator::GreaterThan => "greaterThan",
        core::PivotFilterOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        core::PivotFilterOperator::BeginsWith => "beginsWith",
        core::PivotFilterOperator::DoesNotBeginWith => "doesNotBeginWith",
        core::PivotFilterOperator::EndsWith => "endsWith",
        core::PivotFilterOperator::DoesNotEndWith => "doesNotEndWith",
        core::PivotFilterOperator::Contains => "contains",
        core::PivotFilterOperator::DoesNotContain => "doesNotContain",
    }
}

fn pivot_date_period_to_string(period: core::PivotDatePeriod) -> &'static str {
    match period {
        core::PivotDatePeriod::Tomorrow => "tomorrow",
        core::PivotDatePeriod::Today => "today",
        core::PivotDatePeriod::Yesterday => "yesterday",
        core::PivotDatePeriod::NextWeek => "nextWeek",
        core::PivotDatePeriod::ThisWeek => "thisWeek",
        core::PivotDatePeriod::LastWeek => "lastWeek",
        core::PivotDatePeriod::NextMonth => "nextMonth",
        core::PivotDatePeriod::ThisMonth => "thisMonth",
        core::PivotDatePeriod::LastMonth => "lastMonth",
        core::PivotDatePeriod::NextQuarter => "nextQuarter",
        core::PivotDatePeriod::ThisQuarter => "thisQuarter",
        core::PivotDatePeriod::LastQuarter => "lastQuarter",
        core::PivotDatePeriod::NextYear => "nextYear",
        core::PivotDatePeriod::ThisYear => "thisYear",
        core::PivotDatePeriod::LastYear => "lastYear",
        core::PivotDatePeriod::YearToDate => "yearToDate",
        core::PivotDatePeriod::Quarter(1) => "Q1",
        core::PivotDatePeriod::Quarter(2) => "Q2",
        core::PivotDatePeriod::Quarter(3) => "Q3",
        core::PivotDatePeriod::Quarter(4) => "Q4",
        core::PivotDatePeriod::Month(1) => "M1",
        core::PivotDatePeriod::Month(2) => "M2",
        core::PivotDatePeriod::Month(3) => "M3",
        core::PivotDatePeriod::Month(4) => "M4",
        core::PivotDatePeriod::Month(5) => "M5",
        core::PivotDatePeriod::Month(6) => "M6",
        core::PivotDatePeriod::Month(7) => "M7",
        core::PivotDatePeriod::Month(8) => "M8",
        core::PivotDatePeriod::Month(9) => "M9",
        core::PivotDatePeriod::Month(10) => "M10",
        core::PivotDatePeriod::Month(11) => "M11",
        core::PivotDatePeriod::Month(12) => "M12",
        core::PivotDatePeriod::Month(_) | core::PivotDatePeriod::Quarter(_) => "unknown",
    }
}

fn pivot_date_group_unit_to_string(unit: core::PivotDateGroupUnit) -> &'static str {
    match unit {
        core::PivotDateGroupUnit::Seconds => "seconds",
        core::PivotDateGroupUnit::Minutes => "minutes",
        core::PivotDateGroupUnit::Hours => "hours",
        core::PivotDateGroupUnit::Days => "days",
        core::PivotDateGroupUnit::Months => "months",
        core::PivotDateGroupUnit::Quarters => "quarters",
        core::PivotDateGroupUnit::Years => "years",
    }
}

fn pivot_layout_kind_to_string(kind: core::PivotLayoutKind) -> &'static str {
    match kind {
        core::PivotLayoutKind::Compact => "compact",
        core::PivotLayoutKind::Outline => "outline",
        core::PivotLayoutKind::Tabular => "tabular",
    }
}

fn pivot_values_axis_to_string(axis: core::PivotValuesAxis) -> &'static str {
    match axis {
        core::PivotValuesAxis::Columns => "columns",
        core::PivotValuesAxis::Rows => "rows",
    }
}

fn pivot_overwrite_policy_to_string(policy: core::PivotOverwritePolicy) -> &'static str {
    match policy {
        core::PivotOverwritePolicy::ClearOwnedRange => "clearOwnedRange",
        core::PivotOverwritePolicy::Overwrite => "overwrite",
        core::PivotOverwritePolicy::FailOnOccupied => "failOnOccupied",
    }
}

/// Color representation. The `colorType` field indicates the variant:
/// `"auto"`, `"rgb"`, `"argb"`, `"theme"`, or `"indexed"`.
/// The `hex` field always contains the resolved 6- or 8-char hex string.
#[napi(object)]
pub struct JsColor {
    pub color_type: String,
    /// Resolved hex string (6 or 8 chars, no `#` prefix).
    pub hex: String,
    pub r: Option<u32>,
    pub g: Option<u32>,
    pub b: Option<u32>,
    pub a: Option<u32>,
    /// Theme color index (0-9), present when `colorType === "theme"`.
    pub theme_index: Option<u32>,
    /// Tint percentage (-100 to 100), present when `colorType === "theme"`.
    pub tint: Option<i32>,
    /// Palette index, present when `colorType === "indexed"`.
    pub palette_index: Option<u32>,
}

impl From<&CoreColor> for JsColor {
    fn from(c: &CoreColor) -> Self {
        let hex = c.to_hex();
        match c {
            CoreColor::Auto => JsColor {
                color_type: "auto".into(),
                hex,
                r: None,
                g: None,
                b: None,
                a: None,
                theme_index: None,
                tint: None,
                palette_index: None,
            },
            CoreColor::Rgb { r, g, b } => JsColor {
                color_type: "rgb".into(),
                hex,
                r: Some(*r as u32),
                g: Some(*g as u32),
                b: Some(*b as u32),
                a: None,
                theme_index: None,
                tint: None,
                palette_index: None,
            },
            CoreColor::Argb { a, r, g, b } => JsColor {
                color_type: "argb".into(),
                hex,
                r: Some(*r as u32),
                g: Some(*g as u32),
                b: Some(*b as u32),
                a: Some(*a as u32),
                theme_index: None,
                tint: None,
                palette_index: None,
            },
            CoreColor::Theme { index, tint } => JsColor {
                color_type: "theme".into(),
                hex,
                r: None,
                g: None,
                b: None,
                a: None,
                theme_index: Some(*index as u32),
                tint: Some(*tint as i32),
                palette_index: None,
            },
            CoreColor::Indexed(i) => JsColor {
                color_type: "indexed".into(),
                hex,
                r: None,
                g: None,
                b: None,
                a: None,
                theme_index: None,
                tint: None,
                palette_index: Some(*i as u32),
            },
        }
    }
}

/// Font style settings.
#[napi(object)]
pub struct JsFontStyle {
    pub name: String,
    pub size: f64,
    pub bold: bool,
    pub italic: bool,
    /// One of: `"none"`, `"single"`, `"double"`, `"singleAccounting"`, `"doubleAccounting"`.
    pub underline: String,
    pub strikethrough: bool,
    pub color: JsColor,
    /// One of: `"baseline"`, `"superscript"`, `"subscript"`.
    pub vertical_align: String,
    pub family: Option<u32>,
    pub charset: Option<u32>,
    pub scheme: Option<String>,
}

fn underline_to_string(u: &Underline) -> &'static str {
    match u {
        Underline::None => "none",
        Underline::Single => "single",
        Underline::Double => "double",
        Underline::SingleAccounting => "singleAccounting",
        Underline::DoubleAccounting => "doubleAccounting",
    }
}

fn font_valign_to_string(v: &FontVerticalAlign) -> &'static str {
    match v {
        FontVerticalAlign::Baseline => "baseline",
        FontVerticalAlign::Superscript => "superscript",
        FontVerticalAlign::Subscript => "subscript",
    }
}

impl From<&CoreFontStyle> for JsFontStyle {
    fn from(f: &CoreFontStyle) -> Self {
        JsFontStyle {
            name: f.name.clone(),
            size: f.size,
            bold: f.bold,
            italic: f.italic,
            underline: underline_to_string(&f.underline).into(),
            strikethrough: f.strikethrough,
            color: JsColor::from(&f.color),
            vertical_align: font_valign_to_string(&f.vertical_align).into(),
            family: f.family.map(|v| v as u32),
            charset: f.charset.map(|v| v as u32),
            scheme: f.scheme.clone(),
        }
    }
}

/// Gradient color stop.
#[napi(object)]
pub struct JsGradientStop {
    pub position: f64,
    pub color: JsColor,
}

fn pattern_type_to_string(p: &PatternType) -> &'static str {
    match p {
        PatternType::None => "none",
        PatternType::Solid => "solid",
        PatternType::MediumGray => "mediumGray",
        PatternType::DarkGray => "darkGray",
        PatternType::LightGray => "lightGray",
        PatternType::DarkHorizontal => "darkHorizontal",
        PatternType::DarkVertical => "darkVertical",
        PatternType::DarkDown => "darkDown",
        PatternType::DarkUp => "darkUp",
        PatternType::DarkGrid => "darkGrid",
        PatternType::DarkTrellis => "darkTrellis",
        PatternType::LightHorizontal => "lightHorizontal",
        PatternType::LightVertical => "lightVertical",
        PatternType::LightDown => "lightDown",
        PatternType::LightUp => "lightUp",
        PatternType::LightGrid => "lightGrid",
        PatternType::LightTrellis => "lightTrellis",
        PatternType::Gray125 => "gray125",
        PatternType::Gray0625 => "gray0625",
    }
}

/// Fill/background style. The `fillType` field indicates the variant:
/// `"none"`, `"solid"`, `"pattern"`, or `"gradient"`.
#[napi(object)]
pub struct JsFillStyle {
    pub fill_type: String,
    /// Solid fill color (present when `fillType === "solid"`).
    pub color: Option<JsColor>,
    /// Pattern type string (present when `fillType === "pattern"`).
    pub pattern: Option<String>,
    /// Pattern foreground color.
    pub foreground: Option<JsColor>,
    /// Pattern background color.
    pub background: Option<JsColor>,
    /// Gradient type: `"linear"` or `"path"` (present when `fillType === "gradient"`).
    pub gradient_type: Option<String>,
    /// Gradient angle in degrees.
    pub angle: Option<f64>,
    /// Gradient color stops.
    pub stops: Option<Vec<JsGradientStop>>,
}

impl From<&CoreFillStyle> for JsFillStyle {
    fn from(f: &CoreFillStyle) -> Self {
        match f {
            CoreFillStyle::None => JsFillStyle {
                fill_type: "none".into(),
                color: None,
                pattern: None,
                foreground: None,
                background: None,
                gradient_type: None,
                angle: None,
                stops: None,
            },
            CoreFillStyle::Solid { color } => JsFillStyle {
                fill_type: "solid".into(),
                color: Some(JsColor::from(color)),
                pattern: None,
                foreground: None,
                background: None,
                gradient_type: None,
                angle: None,
                stops: None,
            },
            CoreFillStyle::Pattern {
                pattern,
                foreground,
                background,
            } => JsFillStyle {
                fill_type: "pattern".into(),
                color: None,
                pattern: Some(pattern_type_to_string(pattern).into()),
                foreground: Some(JsColor::from(foreground)),
                background: Some(JsColor::from(background)),
                gradient_type: None,
                angle: None,
                stops: None,
            },
            CoreFillStyle::Gradient {
                gradient_type,
                angle,
                stops,
            } => JsFillStyle {
                fill_type: "gradient".into(),
                color: None,
                pattern: None,
                foreground: None,
                background: None,
                gradient_type: Some(
                    match gradient_type {
                        GradientType::Linear => "linear",
                        GradientType::Path => "path",
                    }
                    .into(),
                ),
                angle: Some(*angle),
                stops: Some(
                    stops
                        .iter()
                        .map(|s| JsGradientStop {
                            position: s.position,
                            color: JsColor::from(&s.color),
                        })
                        .collect(),
                ),
            },
        }
    }
}

fn border_line_style_to_string(s: &CoreBorderLineStyle) -> &'static str {
    match s {
        CoreBorderLineStyle::None => "none",
        CoreBorderLineStyle::Thin => "thin",
        CoreBorderLineStyle::Medium => "medium",
        CoreBorderLineStyle::Thick => "thick",
        CoreBorderLineStyle::Dashed => "dashed",
        CoreBorderLineStyle::Dotted => "dotted",
        CoreBorderLineStyle::Double => "double",
        CoreBorderLineStyle::Hair => "hair",
        CoreBorderLineStyle::MediumDashed => "mediumDashed",
        CoreBorderLineStyle::DashDot => "dashDot",
        CoreBorderLineStyle::MediumDashDot => "mediumDashDot",
        CoreBorderLineStyle::DashDotDot => "dashDotDot",
        CoreBorderLineStyle::MediumDashDotDot => "mediumDashDotDot",
        CoreBorderLineStyle::SlantDashDot => "slantDashDot",
    }
}

/// A single border edge (line style + color).
#[napi(object)]
pub struct JsBorderEdge {
    /// One of: `"none"`, `"thin"`, `"medium"`, `"thick"`, `"dashed"`, `"dotted"`,
    /// `"double"`, `"hair"`, `"mediumDashed"`, `"dashDot"`, `"mediumDashDot"`,
    /// `"dashDotDot"`, `"mediumDashDotDot"`, `"slantDashDot"`.
    pub style: String,
    pub color: JsColor,
}

impl From<&CoreBorderEdge> for JsBorderEdge {
    fn from(e: &CoreBorderEdge) -> Self {
        JsBorderEdge {
            style: border_line_style_to_string(&e.style).into(),
            color: JsColor::from(&e.color),
        }
    }
}

/// Cell border style.
#[napi(object)]
pub struct JsBorderStyle {
    pub left: Option<JsBorderEdge>,
    pub right: Option<JsBorderEdge>,
    pub top: Option<JsBorderEdge>,
    pub bottom: Option<JsBorderEdge>,
    pub diagonal: Option<JsBorderEdge>,
    /// One of: `"none"`, `"down"`, `"up"`, `"both"`.
    pub diagonal_direction: String,
}

impl From<&CoreBorderStyle> for JsBorderStyle {
    fn from(b: &CoreBorderStyle) -> Self {
        JsBorderStyle {
            left: b.left.as_ref().map(JsBorderEdge::from),
            right: b.right.as_ref().map(JsBorderEdge::from),
            top: b.top.as_ref().map(JsBorderEdge::from),
            bottom: b.bottom.as_ref().map(JsBorderEdge::from),
            diagonal: b.diagonal.as_ref().map(JsBorderEdge::from),
            diagonal_direction: match b.diagonal_direction {
                DiagonalDirection::None => "none",
                DiagonalDirection::Down => "down",
                DiagonalDirection::Up => "up",
                DiagonalDirection::Both => "both",
            }
            .into(),
        }
    }
}

fn horizontal_alignment_to_string(a: &HorizontalAlignment) -> &'static str {
    match a {
        HorizontalAlignment::General => "general",
        HorizontalAlignment::Left => "left",
        HorizontalAlignment::Center => "center",
        HorizontalAlignment::Right => "right",
        HorizontalAlignment::Fill => "fill",
        HorizontalAlignment::Justify => "justify",
        HorizontalAlignment::CenterContinuous => "centerContinuous",
        HorizontalAlignment::Distributed => "distributed",
    }
}

fn vertical_alignment_to_string(a: &VerticalAlignment) -> &'static str {
    match a {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Bottom => "bottom",
        VerticalAlignment::Justify => "justify",
        VerticalAlignment::Distributed => "distributed",
    }
}

fn reading_order_to_string(r: &ReadingOrder) -> &'static str {
    match r {
        ReadingOrder::ContextDependent => "contextDependent",
        ReadingOrder::LeftToRight => "leftToRight",
        ReadingOrder::RightToLeft => "rightToLeft",
    }
}

/// Text alignment settings.
#[napi(object)]
pub struct JsAlignment {
    pub horizontal: String,
    pub vertical: String,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
    pub indent: u32,
    /// Rotation in degrees (-90 to 90, or 255 for vertical text).
    pub rotation: i32,
    pub reading_order: String,
}

impl From<&CoreAlignment> for JsAlignment {
    fn from(a: &CoreAlignment) -> Self {
        JsAlignment {
            horizontal: horizontal_alignment_to_string(&a.horizontal).into(),
            vertical: vertical_alignment_to_string(&a.vertical).into(),
            wrap_text: a.wrap_text,
            shrink_to_fit: a.shrink_to_fit,
            indent: a.indent as u32,
            rotation: a.rotation as i32,
            reading_order: reading_order_to_string(&a.reading_order).into(),
        }
    }
}

/// Number format. The `formatType` field indicates the variant:
/// `"general"`, `"builtin"`, or `"custom"`.
#[napi(object)]
pub struct JsNumberFormat {
    pub format_type: String,
    /// Built-in format ID (present when `formatType === "builtin"`).
    pub id: Option<u32>,
    /// The resolved format string (always present).
    pub format_string: String,
    /// Whether this format represents a date/time.
    pub is_date_format: bool,
}

impl From<&CoreNumberFormat> for JsNumberFormat {
    fn from(n: &CoreNumberFormat) -> Self {
        JsNumberFormat {
            format_type: match n {
                CoreNumberFormat::General => "general",
                CoreNumberFormat::BuiltIn(_) => "builtin",
                CoreNumberFormat::Custom(_) => "custom",
            }
            .into(),
            id: match n {
                CoreNumberFormat::BuiltIn(id) => Some(*id),
                _ => None,
            },
            format_string: n.format_string().to_string(),
            is_date_format: n.is_date_format(),
        }
    }
}

/// Cell protection settings.
#[napi(object)]
pub struct JsCellProtection {
    pub locked: bool,
    pub hidden: bool,
}

/// Complete cell style including font, fill, border, alignment, number format,
/// and protection settings.
#[napi(object)]
pub struct JsStyle {
    pub font: JsFontStyle,
    pub fill: JsFillStyle,
    pub border: JsBorderStyle,
    pub alignment: JsAlignment,
    pub number_format: JsNumberFormat,
    pub protection: JsCellProtection,
}

impl From<&CoreStyle> for JsStyle {
    fn from(s: &CoreStyle) -> Self {
        JsStyle {
            font: JsFontStyle::from(&s.font),
            fill: JsFillStyle::from(&s.fill),
            border: JsBorderStyle::from(&s.border),
            alignment: JsAlignment::from(&s.alignment),
            number_format: JsNumberFormat::from(&s.number_format),
            protection: JsCellProtection {
                locked: s.protection.locked,
                hidden: s.protection.hidden,
            },
        }
    }
}

/// Color input for style setters. Mirrors `JsColor`, but all fields are optional
/// so callers can pass either a returned color object or a compact patch.
#[napi(object)]
pub struct JsColorInput {
    pub color_type: Option<String>,
    pub hex: Option<String>,
    pub r: Option<u32>,
    pub g: Option<u32>,
    pub b: Option<u32>,
    pub a: Option<u32>,
    pub theme_index: Option<u32>,
    pub tint: Option<i32>,
    pub palette_index: Option<u32>,
}

/// Font input for style setters. Missing fields leave the existing font setting unchanged.
#[napi(object)]
pub struct JsFontStylePatch {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<String>,
    pub strikethrough: Option<bool>,
    pub color: Option<JsColorInput>,
    pub vertical_align: Option<String>,
    pub family: Option<u32>,
    pub charset: Option<u32>,
    pub scheme: Option<String>,
}

/// Gradient color stop input for style setters.
#[napi(object)]
pub struct JsGradientStopInput {
    pub position: f64,
    pub color: JsColorInput,
}

/// Fill/background input for style setters.
#[napi(object)]
pub struct JsFillStylePatch {
    pub fill_type: Option<String>,
    pub color: Option<JsColorInput>,
    pub pattern: Option<String>,
    pub foreground: Option<JsColorInput>,
    pub background: Option<JsColorInput>,
    pub gradient_type: Option<String>,
    pub angle: Option<f64>,
    pub stops: Option<Vec<JsGradientStopInput>>,
}

/// Border edge input for style setters.
#[napi(object)]
pub struct JsBorderEdgePatch {
    pub style: Option<String>,
    pub color: Option<JsColorInput>,
}

/// Border input for style setters.
#[napi(object)]
pub struct JsBorderStylePatch {
    pub left: Option<JsBorderEdgePatch>,
    pub right: Option<JsBorderEdgePatch>,
    pub top: Option<JsBorderEdgePatch>,
    pub bottom: Option<JsBorderEdgePatch>,
    pub diagonal: Option<JsBorderEdgePatch>,
    pub diagonal_direction: Option<String>,
}

/// Alignment input for style setters.
#[napi(object)]
pub struct JsAlignmentPatch {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: Option<bool>,
    pub shrink_to_fit: Option<bool>,
    pub indent: Option<u32>,
    pub rotation: Option<i32>,
    pub reading_order: Option<String>,
}

/// Number format input for style setters.
#[napi(object)]
pub struct JsNumberFormatPatch {
    pub format_type: Option<String>,
    pub id: Option<u32>,
    pub format_string: Option<String>,
}

/// Cell protection input for style setters.
#[napi(object)]
pub struct JsCellProtectionPatch {
    pub locked: Option<bool>,
    pub hidden: Option<bool>,
}

/// Cell style input for style setters. A complete `JsStyle` returned from
/// `getCellStyle()` is assignable to this type; partial objects act as patches.
#[napi(object)]
pub struct JsStylePatch {
    pub font: Option<JsFontStylePatch>,
    pub fill: Option<JsFillStylePatch>,
    pub border: Option<JsBorderStylePatch>,
    pub alignment: Option<JsAlignmentPatch>,
    pub number_format: Option<JsNumberFormatPatch>,
    pub protection: Option<JsCellProtectionPatch>,
}

fn style_input_error(message: impl Into<String>) -> NapiError {
    NapiError::from_reason(message.into())
}

fn u32_to_u8(value: u32, field: &str) -> NapiResult<u8> {
    u8::try_from(value).map_err(|_| style_input_error(format!("{field} must be between 0 and 255")))
}

fn i32_to_i8(value: i32, field: &str) -> NapiResult<i8> {
    i8::try_from(value).map_err(|_| style_input_error(format!("{field} must be between -128 and 127")))
}

fn parse_color_hex(hex: &str) -> NapiResult<CoreColor> {
    CoreColor::from_hex(hex).ok_or_else(|| {
        style_input_error("color hex must be 6 or 8 hexadecimal characters, with optional # prefix")
    })
}

fn parse_rgb_hex(hex: &str) -> NapiResult<CoreColor> {
    match parse_color_hex(hex)? {
        CoreColor::Rgb { r, g, b } => Ok(CoreColor::Rgb { r, g, b }),
        CoreColor::Argb { r, g, b, .. } => Ok(CoreColor::Rgb { r, g, b }),
        other => Ok(other),
    }
}

fn parse_argb_hex(hex: &str) -> NapiResult<CoreColor> {
    match parse_color_hex(hex)? {
        CoreColor::Rgb { r, g, b } => Ok(CoreColor::Argb { a: 255, r, g, b }),
        CoreColor::Argb { a, r, g, b } => Ok(CoreColor::Argb { a, r, g, b }),
        other => Ok(other),
    }
}

impl JsColorInput {
    pub fn to_core_color(&self) -> NapiResult<CoreColor> {
        match self.color_type.as_deref() {
            Some("auto") => Ok(CoreColor::Auto),
            Some("rgb") => {
                if let Some(hex) = &self.hex {
                    parse_rgb_hex(hex)
                } else {
                    Ok(CoreColor::Rgb {
                        r: u32_to_u8(self.r.ok_or_else(|| style_input_error("rgb color requires r"))?, "r")?,
                        g: u32_to_u8(self.g.ok_or_else(|| style_input_error("rgb color requires g"))?, "g")?,
                        b: u32_to_u8(self.b.ok_or_else(|| style_input_error("rgb color requires b"))?, "b")?,
                    })
                }
            }
            Some("argb") => {
                if let Some(hex) = &self.hex {
                    parse_argb_hex(hex)
                } else {
                    Ok(CoreColor::Argb {
                        a: u32_to_u8(self.a.unwrap_or(255), "a")?,
                        r: u32_to_u8(self.r.ok_or_else(|| style_input_error("argb color requires r"))?, "r")?,
                        g: u32_to_u8(self.g.ok_or_else(|| style_input_error("argb color requires g"))?, "g")?,
                        b: u32_to_u8(self.b.ok_or_else(|| style_input_error("argb color requires b"))?, "b")?,
                    })
                }
            }
            Some("theme") => Ok(CoreColor::Theme {
                index: u32_to_u8(
                    self.theme_index
                        .ok_or_else(|| style_input_error("theme color requires themeIndex"))?,
                    "themeIndex",
                )?,
                tint: i32_to_i8(self.tint.unwrap_or(0), "tint")?,
            }),
            Some("indexed") => Ok(CoreColor::Indexed(u32_to_u8(
                self.palette_index
                    .ok_or_else(|| style_input_error("indexed color requires paletteIndex"))?,
                "paletteIndex",
            )?)),
            Some(other) => Err(style_input_error(format!("unknown colorType {other:?}"))),
            None => {
                if let Some(hex) = &self.hex {
                    parse_color_hex(hex)
                } else if self.r.is_some() || self.g.is_some() || self.b.is_some() {
                    Ok(CoreColor::Rgb {
                        r: u32_to_u8(self.r.ok_or_else(|| style_input_error("rgb color requires r"))?, "r")?,
                        g: u32_to_u8(self.g.ok_or_else(|| style_input_error("rgb color requires g"))?, "g")?,
                        b: u32_to_u8(self.b.ok_or_else(|| style_input_error("rgb color requires b"))?, "b")?,
                    })
                } else if let Some(theme_index) = self.theme_index {
                    Ok(CoreColor::Theme {
                        index: u32_to_u8(theme_index, "themeIndex")?,
                        tint: i32_to_i8(self.tint.unwrap_or(0), "tint")?,
                    })
                } else if let Some(palette_index) = self.palette_index {
                    Ok(CoreColor::Indexed(u32_to_u8(palette_index, "paletteIndex")?))
                } else {
                    Err(style_input_error("color requires colorType, hex, rgb, themeIndex, or paletteIndex"))
                }
            }
        }
    }
}

fn parse_underline(value: &str) -> NapiResult<Underline> {
    match value {
        "none" => Ok(Underline::None),
        "single" => Ok(Underline::Single),
        "double" => Ok(Underline::Double),
        "singleAccounting" => Ok(Underline::SingleAccounting),
        "doubleAccounting" => Ok(Underline::DoubleAccounting),
        other => Err(style_input_error(format!("unknown underline {other:?}"))),
    }
}

fn parse_font_vertical_align(value: &str) -> NapiResult<FontVerticalAlign> {
    match value {
        "baseline" => Ok(FontVerticalAlign::Baseline),
        "superscript" => Ok(FontVerticalAlign::Superscript),
        "subscript" => Ok(FontVerticalAlign::Subscript),
        other => Err(style_input_error(format!("unknown verticalAlign {other:?}"))),
    }
}

impl JsFontStylePatch {
    fn is_full_font(&self) -> bool {
        self.name.is_some()
            && self.size.is_some()
            && self.bold.is_some()
            && self.italic.is_some()
            && self.underline.is_some()
            && self.strikethrough.is_some()
            && self.color.is_some()
            && self.vertical_align.is_some()
    }

    fn apply_to_core_font(&self, font: &mut CoreFontStyle) -> NapiResult<()> {
        if let Some(name) = &self.name {
            font.name = name.clone();
        }
        if let Some(size) = self.size {
            font.size = size;
        }
        if let Some(bold) = self.bold {
            font.bold = bold;
        }
        if let Some(italic) = self.italic {
            font.italic = italic;
        }
        if let Some(underline) = &self.underline {
            font.underline = parse_underline(underline)?;
        }
        if let Some(strikethrough) = self.strikethrough {
            font.strikethrough = strikethrough;
        }
        if let Some(color) = &self.color {
            font.color = color.to_core_color()?;
        }
        if let Some(vertical_align) = &self.vertical_align {
            font.vertical_align = parse_font_vertical_align(vertical_align)?;
        }
        if let Some(family) = self.family {
            font.family = Some(u32_to_u8(family, "family")?);
        }
        if let Some(charset) = self.charset {
            font.charset = Some(u32_to_u8(charset, "charset")?);
        }
        if let Some(scheme) = &self.scheme {
            font.scheme = Some(scheme.clone());
        }
        Ok(())
    }
}

fn parse_pattern_type(value: &str) -> NapiResult<PatternType> {
    match value {
        "none" => Ok(PatternType::None),
        "solid" => Ok(PatternType::Solid),
        "mediumGray" => Ok(PatternType::MediumGray),
        "darkGray" => Ok(PatternType::DarkGray),
        "lightGray" => Ok(PatternType::LightGray),
        "darkHorizontal" => Ok(PatternType::DarkHorizontal),
        "darkVertical" => Ok(PatternType::DarkVertical),
        "darkDown" => Ok(PatternType::DarkDown),
        "darkUp" => Ok(PatternType::DarkUp),
        "darkGrid" => Ok(PatternType::DarkGrid),
        "darkTrellis" => Ok(PatternType::DarkTrellis),
        "lightHorizontal" => Ok(PatternType::LightHorizontal),
        "lightVertical" => Ok(PatternType::LightVertical),
        "lightDown" => Ok(PatternType::LightDown),
        "lightUp" => Ok(PatternType::LightUp),
        "lightGrid" => Ok(PatternType::LightGrid),
        "lightTrellis" => Ok(PatternType::LightTrellis),
        "gray125" => Ok(PatternType::Gray125),
        "gray0625" => Ok(PatternType::Gray0625),
        other => Err(style_input_error(format!("unknown fill pattern {other:?}"))),
    }
}

fn parse_gradient_type(value: &str) -> NapiResult<GradientType> {
    match value {
        "linear" => Ok(GradientType::Linear),
        "path" => Ok(GradientType::Path),
        other => Err(style_input_error(format!("unknown gradientType {other:?}"))),
    }
}

impl JsFillStylePatch {
    fn to_core_fill(&self) -> NapiResult<CoreFillStyle> {
        match self.fill_type.as_deref() {
            Some("none") => Ok(CoreFillStyle::None),
            Some("solid") | None if self.color.is_some() => Ok(CoreFillStyle::Solid {
                color: self
                    .color
                    .as_ref()
                    .ok_or_else(|| style_input_error("solid fill requires color"))?
                    .to_core_color()?,
            }),
            Some("pattern") => Ok(CoreFillStyle::Pattern {
                pattern: parse_pattern_type(
                    self.pattern
                        .as_deref()
                        .ok_or_else(|| style_input_error("pattern fill requires pattern"))?,
                )?,
                foreground: self
                    .foreground
                    .as_ref()
                    .ok_or_else(|| style_input_error("pattern fill requires foreground"))?
                    .to_core_color()?,
                background: self
                    .background
                    .as_ref()
                    .ok_or_else(|| style_input_error("pattern fill requires background"))?
                    .to_core_color()?,
            }),
            Some("gradient") => Ok(CoreFillStyle::Gradient {
                gradient_type: parse_gradient_type(self.gradient_type.as_deref().unwrap_or("linear"))?,
                angle: self.angle.unwrap_or(0.0),
                stops: self
                    .stops
                    .as_ref()
                    .ok_or_else(|| style_input_error("gradient fill requires stops"))?
                    .iter()
                    .map(|stop| {
                        Ok(duke_sheets_core::style::GradientStop {
                            position: stop.position,
                            color: stop.color.to_core_color()?,
                        })
                    })
                    .collect::<NapiResult<Vec<_>>>()?,
            }),
            Some(other) => Err(style_input_error(format!("unknown fillType {other:?}"))),
            None => Err(style_input_error("fill patch requires fillType or color")),
        }
    }
}

fn parse_border_line_style(value: &str) -> NapiResult<CoreBorderLineStyle> {
    match value {
        "none" => Ok(CoreBorderLineStyle::None),
        "thin" => Ok(CoreBorderLineStyle::Thin),
        "medium" => Ok(CoreBorderLineStyle::Medium),
        "thick" => Ok(CoreBorderLineStyle::Thick),
        "dashed" => Ok(CoreBorderLineStyle::Dashed),
        "dotted" => Ok(CoreBorderLineStyle::Dotted),
        "double" => Ok(CoreBorderLineStyle::Double),
        "hair" => Ok(CoreBorderLineStyle::Hair),
        "mediumDashed" => Ok(CoreBorderLineStyle::MediumDashed),
        "dashDot" => Ok(CoreBorderLineStyle::DashDot),
        "mediumDashDot" => Ok(CoreBorderLineStyle::MediumDashDot),
        "dashDotDot" => Ok(CoreBorderLineStyle::DashDotDot),
        "mediumDashDotDot" => Ok(CoreBorderLineStyle::MediumDashDotDot),
        "slantDashDot" => Ok(CoreBorderLineStyle::SlantDashDot),
        other => Err(style_input_error(format!("unknown border style {other:?}"))),
    }
}

fn parse_diagonal_direction(value: &str) -> NapiResult<DiagonalDirection> {
    match value {
        "none" => Ok(DiagonalDirection::None),
        "down" => Ok(DiagonalDirection::Down),
        "up" => Ok(DiagonalDirection::Up),
        "both" => Ok(DiagonalDirection::Both),
        other => Err(style_input_error(format!("unknown diagonalDirection {other:?}"))),
    }
}

impl JsBorderEdgePatch {
    fn apply_to_edge(&self, existing: Option<&CoreBorderEdge>) -> NapiResult<Option<CoreBorderEdge>> {
        let parsed_style = self
            .style
            .as_deref()
            .map(parse_border_line_style)
            .transpose()?;
        if parsed_style == Some(CoreBorderLineStyle::None) {
            return Ok(None);
        }

        let mut edge = existing
            .cloned()
            .unwrap_or_else(|| CoreBorderEdge::new(CoreBorderLineStyle::Thin, CoreColor::BLACK));
        if let Some(style) = parsed_style {
            edge.style = style;
        }
        if let Some(color) = &self.color {
            edge.color = color.to_core_color()?;
        }
        Ok(Some(edge))
    }
}

impl JsBorderStylePatch {
    fn is_full_border(&self) -> bool {
        self.diagonal_direction.is_some()
    }

    fn apply_to_core_border(&self, border: &mut CoreBorderStyle) -> NapiResult<()> {
        if let Some(edge) = &self.left {
            border.left = edge.apply_to_edge(border.left.as_ref())?;
        }
        if let Some(edge) = &self.right {
            border.right = edge.apply_to_edge(border.right.as_ref())?;
        }
        if let Some(edge) = &self.top {
            border.top = edge.apply_to_edge(border.top.as_ref())?;
        }
        if let Some(edge) = &self.bottom {
            border.bottom = edge.apply_to_edge(border.bottom.as_ref())?;
        }
        if let Some(edge) = &self.diagonal {
            border.diagonal = edge.apply_to_edge(border.diagonal.as_ref())?;
        }
        if let Some(direction) = &self.diagonal_direction {
            border.diagonal_direction = parse_diagonal_direction(direction)?;
        }
        Ok(())
    }
}

fn parse_horizontal_alignment(value: &str) -> NapiResult<HorizontalAlignment> {
    match value {
        "general" => Ok(HorizontalAlignment::General),
        "left" => Ok(HorizontalAlignment::Left),
        "center" => Ok(HorizontalAlignment::Center),
        "right" => Ok(HorizontalAlignment::Right),
        "fill" => Ok(HorizontalAlignment::Fill),
        "justify" => Ok(HorizontalAlignment::Justify),
        "centerContinuous" => Ok(HorizontalAlignment::CenterContinuous),
        "distributed" => Ok(HorizontalAlignment::Distributed),
        other => Err(style_input_error(format!("unknown horizontal alignment {other:?}"))),
    }
}

fn parse_vertical_alignment(value: &str) -> NapiResult<VerticalAlignment> {
    match value {
        "top" => Ok(VerticalAlignment::Top),
        "center" => Ok(VerticalAlignment::Center),
        "bottom" => Ok(VerticalAlignment::Bottom),
        "justify" => Ok(VerticalAlignment::Justify),
        "distributed" => Ok(VerticalAlignment::Distributed),
        other => Err(style_input_error(format!("unknown vertical alignment {other:?}"))),
    }
}

fn parse_reading_order(value: &str) -> NapiResult<ReadingOrder> {
    match value {
        "contextDependent" => Ok(ReadingOrder::ContextDependent),
        "leftToRight" => Ok(ReadingOrder::LeftToRight),
        "rightToLeft" => Ok(ReadingOrder::RightToLeft),
        other => Err(style_input_error(format!("unknown readingOrder {other:?}"))),
    }
}

impl JsAlignmentPatch {
    fn is_full_alignment(&self) -> bool {
        self.horizontal.is_some()
            && self.vertical.is_some()
            && self.wrap_text.is_some()
            && self.shrink_to_fit.is_some()
            && self.indent.is_some()
            && self.rotation.is_some()
            && self.reading_order.is_some()
    }

    fn apply_to_core_alignment(&self, alignment: &mut CoreAlignment) -> NapiResult<()> {
        if let Some(horizontal) = &self.horizontal {
            alignment.horizontal = parse_horizontal_alignment(horizontal)?;
        }
        if let Some(vertical) = &self.vertical {
            alignment.vertical = parse_vertical_alignment(vertical)?;
        }
        if let Some(wrap_text) = self.wrap_text {
            alignment.wrap_text = wrap_text;
        }
        if let Some(shrink_to_fit) = self.shrink_to_fit {
            alignment.shrink_to_fit = shrink_to_fit;
        }
        if let Some(indent) = self.indent {
            alignment.indent = u32_to_u8(indent, "indent")?;
        }
        if let Some(rotation) = self.rotation {
            if !((-90..=90).contains(&rotation) || rotation == 255) {
                return Err(style_input_error("rotation must be between -90 and 90, or 255"));
            }
            alignment.rotation = rotation as i16;
        }
        if let Some(reading_order) = &self.reading_order {
            alignment.reading_order = parse_reading_order(reading_order)?;
        }
        Ok(())
    }
}

impl JsNumberFormatPatch {
    fn to_core_number_format(&self) -> NapiResult<CoreNumberFormat> {
        match self.format_type.as_deref() {
            Some("general") => Ok(CoreNumberFormat::General),
            Some("builtin") => Ok(CoreNumberFormat::BuiltIn(
                self.id
                    .ok_or_else(|| style_input_error("builtin number format requires id"))?,
            )),
            Some("custom") => Ok(CoreNumberFormat::Custom(
                self.format_string
                    .clone()
                    .ok_or_else(|| style_input_error("custom number format requires formatString"))?,
            )),
            Some(other) => Err(style_input_error(format!("unknown formatType {other:?}"))),
            None if self.id.is_some() => Ok(CoreNumberFormat::BuiltIn(self.id.unwrap())),
            None if self.format_string.is_some() => Ok(CoreNumberFormat::Custom(
                self.format_string.clone().unwrap(),
            )),
            None => Err(style_input_error("numberFormat requires formatType, id, or formatString")),
        }
    }
}

impl JsCellProtectionPatch {
    fn apply_to_core_protection(&self, protection: &mut duke_sheets_core::style::Protection) {
        if let Some(locked) = self.locked {
            protection.locked = locked;
        }
        if let Some(hidden) = self.hidden {
            protection.hidden = hidden;
        }
    }
}

impl JsStylePatch {
    pub fn apply_to_core_style(&self, style: &mut CoreStyle) -> NapiResult<()> {
        if let Some(font_patch) = &self.font {
            if font_patch.is_full_font() {
                let mut font = CoreFontStyle::default();
                font_patch.apply_to_core_font(&mut font)?;
                style.font = font;
            } else {
                font_patch.apply_to_core_font(&mut style.font)?;
            }
        }

        if let Some(fill_patch) = &self.fill {
            style.fill = fill_patch.to_core_fill()?;
        }

        if let Some(border_patch) = &self.border {
            if border_patch.is_full_border() {
                let mut border = CoreBorderStyle::default();
                border_patch.apply_to_core_border(&mut border)?;
                style.border = border;
            } else {
                border_patch.apply_to_core_border(&mut style.border)?;
            }
        }

        if let Some(alignment_patch) = &self.alignment {
            if alignment_patch.is_full_alignment() {
                let mut alignment = CoreAlignment::default();
                alignment_patch.apply_to_core_alignment(&mut alignment)?;
                style.alignment = alignment;
            } else {
                alignment_patch.apply_to_core_alignment(&mut style.alignment)?;
            }
        }

        if let Some(number_format_patch) = &self.number_format {
            style.number_format = number_format_patch.to_core_number_format()?;
        }

        if let Some(protection_patch) = &self.protection {
            protection_patch.apply_to_core_protection(&mut style.protection);
        }

        Ok(())
    }
}

/// A hyperlink attached to a cell.
#[napi(object)]
pub struct JsHyperlink {
    /// Target URL (external) or cell reference (internal).
    pub target: String,
    /// Display text (shown in cell; `null` means cell value is used).
    pub display: Option<String>,
    /// Tooltip shown on hover.
    pub tooltip: Option<String>,
    /// Location within target (e.g., sheet reference for internal links).
    pub location: Option<String>,
}

impl From<&core::Hyperlink> for JsHyperlink {
    fn from(h: &core::Hyperlink) -> Self {
        JsHyperlink {
            target: h.target.clone(),
            display: h.display.clone(),
            tooltip: h.tooltip.clone(),
            location: h.location.clone(),
        }
    }
}

/// A cell comment/note.
#[napi(object)]
pub struct JsComment {
    pub author: String,
    pub text: String,
    pub visible: bool,
}

impl From<&core::CellComment> for JsComment {
    fn from(c: &core::CellComment) -> Self {
        JsComment {
            author: c.author.clone(),
            text: c.text.clone(),
            visible: c.visible,
        }
    }
}

/// A comment with its cell address.
#[napi(object)]
pub struct JsCommentEntry {
    pub row: u32,
    pub col: u32,
    pub comment: JsComment,
}

/// Freeze pane settings.
#[napi(object)]
pub struct JsFreezePanes {
    /// First unfrozen row.
    pub row: u32,
    /// First unfrozen column.
    pub col: u32,
}

impl From<&core::FreezePanes> for JsFreezePanes {
    fn from(f: &core::FreezePanes) -> Self {
        JsFreezePanes {
            row: f.row,
            col: f.col as u32,
        }
    }
}

/// Split pane settings.
#[napi(object)]
pub struct JsSplitPanes {
    pub x_split: f64,
    pub y_split: f64,
    pub top_left_row: Option<u32>,
    pub top_left_col: Option<u32>,
    pub active_pane: Option<String>,
}

impl From<&core::SplitPanes> for JsSplitPanes {
    fn from(s: &core::SplitPanes) -> Self {
        JsSplitPanes {
            x_split: s.x_split,
            y_split: s.y_split,
            top_left_row: s.top_left.map(|(r, _)| r),
            top_left_col: s.top_left.map(|(_, c)| c as u32),
            active_pane: s.active_pane.clone(),
        }
    }
}

/// A selection within a sheet view.
#[napi(object)]
pub struct JsSelection {
    pub pane: Option<String>,
    pub active_cell: Option<String>,
    pub sqref: Option<String>,
}

impl From<&core::Selection> for JsSelection {
    fn from(s: &core::Selection) -> Self {
        JsSelection {
            pane: s.pane.clone(),
            active_cell: s.active_cell.clone(),
            sqref: s.sqref.clone(),
        }
    }
}

/// Sheet protection settings.
#[napi(object)]
pub struct JsSheetProtection {
    pub protected: bool,
    pub select_locked_cells: bool,
    pub select_unlocked_cells: bool,
    pub format_cells: bool,
    pub format_columns: bool,
    pub format_rows: bool,
    pub insert_columns: bool,
    pub insert_rows: bool,
    pub insert_hyperlinks: bool,
    pub delete_columns: bool,
    pub delete_rows: bool,
    pub sort: bool,
    pub auto_filter: bool,
    pub pivot_tables: bool,
}

impl From<&core::SheetProtection> for JsSheetProtection {
    fn from(p: &core::SheetProtection) -> Self {
        JsSheetProtection {
            protected: p.protected,
            select_locked_cells: p.select_locked_cells,
            select_unlocked_cells: p.select_unlocked_cells,
            format_cells: p.format_cells,
            format_columns: p.format_columns,
            format_rows: p.format_rows,
            insert_columns: p.insert_columns,
            insert_rows: p.insert_rows,
            insert_hyperlinks: p.insert_hyperlinks,
            delete_columns: p.delete_columns,
            delete_rows: p.delete_rows,
            sort: p.sort,
            auto_filter: p.auto_filter,
            pivot_tables: p.pivot_tables,
        }
    }
}

/// Page setup / print settings.
#[napi(object)]
pub struct JsPageSetup {
    /// Paper size (1 = Letter, 9 = A4, etc.).
    pub paper_size: u32,
    /// `"portrait"` or `"landscape"`.
    pub orientation: String,
    /// Scale percentage (10-400).
    pub scale: u32,
    pub fit_to_width: Option<u32>,
    pub fit_to_height: Option<u32>,
    pub top_margin: f64,
    pub bottom_margin: f64,
    pub left_margin: f64,
    pub right_margin: f64,
    pub header_margin: f64,
    pub footer_margin: f64,
    pub print_gridlines: bool,
    pub print_headings: bool,
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
    pub different_odd_even: bool,
    pub different_first: bool,
    pub scale_with_doc: bool,
    pub align_with_margins: bool,
}

impl From<&core::PageSetup> for JsPageSetup {
    fn from(p: &core::PageSetup) -> Self {
        JsPageSetup {
            paper_size: p.paper_size as u32,
            orientation: match p.orientation {
                core::PageOrientation::Portrait => "portrait",
                core::PageOrientation::Landscape => "landscape",
            }
            .into(),
            scale: p.scale as u32,
            fit_to_width: p.fit_to_width.map(|v| v as u32),
            fit_to_height: p.fit_to_height.map(|v| v as u32),
            top_margin: p.top_margin,
            bottom_margin: p.bottom_margin,
            left_margin: p.left_margin,
            right_margin: p.right_margin,
            header_margin: p.header_margin,
            footer_margin: p.footer_margin,
            print_gridlines: p.print_gridlines,
            print_headings: p.print_headings,
            odd_header: p.odd_header.clone(),
            odd_footer: p.odd_footer.clone(),
            even_header: p.even_header.clone(),
            even_footer: p.even_footer.clone(),
            first_header: p.first_header.clone(),
            first_footer: p.first_footer.clone(),
            different_odd_even: p.different_odd_even,
            different_first: p.different_first,
            scale_with_doc: p.scale_with_doc,
            align_with_margins: p.align_with_margins,
        }
    }
}

/// A manual page break (row or column).
#[napi(object)]
pub struct JsPageBreak {
    /// Row index (for row breaks) or column index (for col breaks), 0-based.
    pub id: u32,
    pub min: u32,
    pub max: u32,
    /// Whether this is a manual break.
    pub manual: bool,
}

impl From<&core::PageBreak> for JsPageBreak {
    fn from(b: &core::PageBreak) -> Self {
        JsPageBreak {
            id: b.id,
            min: b.min,
            max: b.max,
            manual: b.man,
        }
    }
}

/// Workbook-level settings.
#[napi(object)]
pub struct JsWorkbookSettings {
    /// Whether the 1904 date system is used (macOS default).
    pub date_1904: bool,
    /// Whether the workbook structure is protected.
    pub protected: bool,
    /// Calculate formulas on open.
    pub calc_on_open: bool,
    pub theme: Option<String>,
}

impl From<&core::WorkbookSettings> for JsWorkbookSettings {
    fn from(s: &core::WorkbookSettings) -> Self {
        JsWorkbookSettings {
            date_1904: s.date_1904,
            protected: s.protected,
            calc_on_open: s.calc_on_open,
            theme: s.theme.clone(),
        }
    }
}

/// A named range definition.
#[napi(object)]
pub struct JsNamedRange {
    pub name: String,
    /// `"workbook"` or `"sheet"`.
    pub scope: String,
    /// Sheet index when `scope === "sheet"`.
    pub sheet_index: Option<u32>,
    /// The formula/reference the name refers to.
    pub refers_to: String,
    pub comment: Option<String>,
    pub hidden: bool,
}

/// An Excel table (ListObject).
#[napi(object)]
pub struct JsTable {
    pub id: u32,
    pub name: String,
    pub display_name: String,
    /// Range string (e.g., `"A1:D10"`).
    pub reference: String,
    pub columns: Vec<JsTableColumn>,
    pub style_info: Option<JsTableStyleInfo>,
    pub header_row_count: u32,
    pub totals_row_count: u32,
    pub totals_row_shown: bool,
}

impl From<&core::Table> for JsTable {
    fn from(t: &core::Table) -> Self {
        JsTable {
            id: t.id,
            name: t.name.clone(),
            display_name: t.display_name.clone(),
            reference: t.reference.to_string(),
            columns: t.columns.iter().map(JsTableColumn::from).collect(),
            style_info: t.style_info.as_ref().map(JsTableStyleInfo::from),
            header_row_count: t.header_row_count,
            totals_row_count: t.totals_row_count,
            totals_row_shown: t.totals_row_shown,
        }
    }
}

/// A column within a table.
#[napi(object)]
pub struct JsTableColumn {
    pub id: u32,
    pub name: String,
    /// One of: `"average"`, `"count"`, `"countNums"`, `"max"`, `"min"`,
    /// `"sum"`, `"stdDev"`, `"var"`, `"custom"`, `"none"`, or `null`.
    pub totals_row_function: Option<String>,
    pub totals_row_formula: Option<String>,
    pub totals_row_label: Option<String>,
    pub calculated_column_formula: Option<String>,
}

impl From<&core::TableColumn> for JsTableColumn {
    fn from(c: &core::TableColumn) -> Self {
        JsTableColumn {
            id: c.id,
            name: c.name.clone(),
            totals_row_function: c.totals_row_function.as_ref().map(|f| f.to_ooxml().into()),
            totals_row_formula: c.totals_row_formula.clone(),
            totals_row_label: c.totals_row_label.clone(),
            calculated_column_formula: c.calculated_column_formula.clone(),
        }
    }
}

/// Table style configuration.
#[napi(object)]
pub struct JsTableStyleInfo {
    pub name: Option<String>,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
}

impl From<&core::TableStyleInfo> for JsTableStyleInfo {
    fn from(s: &core::TableStyleInfo) -> Self {
        JsTableStyleInfo {
            name: s.name.clone(),
            show_first_column: s.show_first_column,
            show_last_column: s.show_last_column,
            show_row_stripes: s.show_row_stripes,
            show_column_stripes: s.show_column_stripes,
        }
    }
}

/// A standalone auto-filter on a worksheet.
#[napi(object)]
pub struct JsAutoFilter {
    /// Range string the filter covers (e.g., `"A1:D10"`).
    pub range: String,
    pub filter_columns: Vec<JsFilterColumn>,
}

impl From<&core::AutoFilter> for JsAutoFilter {
    fn from(af: &core::AutoFilter) -> Self {
        JsAutoFilter {
            range: af.range.to_string(),
            filter_columns: af.filter_columns.iter().map(JsFilterColumn::from).collect(),
        }
    }
}

/// A filter on a single column.
#[napi(object)]
pub struct JsFilterColumn {
    pub col_id: u32,
    pub hidden_button: bool,
    pub show_button: bool,
    /// The type of filter: `"values"`, `"custom"`, `"top10"`, `"dynamic"`, or `"color"`.
    pub filter_type: String,
    /// Values for a discrete value filter (present when `filterType === "values"`).
    pub values: Option<Vec<String>>,
    /// Include blanks (present when `filterType === "values"`).
    pub blank: Option<bool>,
}

impl From<&core::FilterColumn> for JsFilterColumn {
    fn from(fc: &core::FilterColumn) -> Self {
        let (filter_type, values, blank) = match &fc.filter {
            core::ColumnFilter::Values(vf) => ("values", Some(vf.values.clone()), Some(vf.blank)),
            core::ColumnFilter::Custom(_) => ("custom", None, None),
            core::ColumnFilter::Top10(_) => ("top10", None, None),
            core::ColumnFilter::Dynamic(_) => ("dynamic", None, None),
            core::ColumnFilter::Color(_) => ("color", None, None),
        };
        JsFilterColumn {
            col_id: fc.col_id,
            hidden_button: fc.hidden_button,
            show_button: fc.show_button,
            filter_type: filter_type.into(),
            values,
            blank,
        }
    }
}

/// A data validation rule.
#[napi(object)]
pub struct JsDataValidation {
    /// The type of validation: `"none"`, `"whole"`, `"decimal"`, `"list"`,
    /// `"date"`, `"time"`, `"textLength"`, `"custom"`.
    pub validation_type: String,
    /// Ranges this validation applies to (as range strings).
    pub ranges: Vec<String>,
    pub allow_blank: bool,
    pub show_dropdown: bool,
    pub show_input_message: bool,
    pub input_title: Option<String>,
    pub input_message: Option<String>,
    pub show_error_alert: bool,
    /// `"stop"`, `"warning"`, or `"information"`.
    pub error_style: String,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
    /// Operator (present for numeric/date/time/textLength validations):
    /// `"between"`, `"notBetween"`, `"equal"`, `"notEqual"`, `"greaterThan"`,
    /// `"lessThan"`, `"greaterThanOrEqual"`, `"lessThanOrEqual"`.
    pub operator: Option<String>,
    /// First value/formula (present for most validation types).
    pub value1: Option<String>,
    /// Second value/formula (present for `between`/`notBetween`).
    pub value2: Option<String>,
    /// List source string (present when `validationType === "list"`).
    pub list_source: Option<String>,
    /// Custom formula (present when `validationType === "custom"`).
    pub formula: Option<String>,
}

fn validation_operator_to_string(op: &core::ValidationOperator) -> &'static str {
    match op {
        core::ValidationOperator::Between => "between",
        core::ValidationOperator::NotBetween => "notBetween",
        core::ValidationOperator::Equal => "equal",
        core::ValidationOperator::NotEqual => "notEqual",
        core::ValidationOperator::GreaterThan => "greaterThan",
        core::ValidationOperator::LessThan => "lessThan",
        core::ValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        core::ValidationOperator::LessThanOrEqual => "lessThanOrEqual",
    }
}

impl From<&core::DataValidation> for JsDataValidation {
    fn from(dv: &core::DataValidation) -> Self {
        let (vtype, operator, value1, value2, list_source, formula) = match &dv.validation_type {
            core::ValidationType::None => ("none", None, None, None, None, None),
            core::ValidationType::Whole {
                operator,
                value1,
                value2,
            } => (
                "whole",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::Decimal {
                operator,
                value1,
                value2,
            } => (
                "decimal",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::List { source } => {
                ("list", None, None, None, Some(source.clone()), None)
            }
            core::ValidationType::Date {
                operator,
                value1,
                value2,
            } => (
                "date",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::Time {
                operator,
                value1,
                value2,
            } => (
                "time",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::TextLength {
                operator,
                value1,
                value2,
            } => (
                "textLength",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::Custom { formula } => {
                ("custom", None, None, None, None, Some(formula.clone()))
            }
        };

        JsDataValidation {
            validation_type: vtype.into(),
            ranges: dv.ranges.iter().map(|r| r.to_string()).collect(),
            allow_blank: dv.allow_blank,
            show_dropdown: dv.show_dropdown,
            show_input_message: dv.show_input_message,
            input_title: dv.input_title.clone(),
            input_message: dv.input_message.clone(),
            show_error_alert: dv.show_error_alert,
            error_style: match dv.error_style {
                core::ValidationErrorStyle::Stop => "stop",
                core::ValidationErrorStyle::Warning => "warning",
                core::ValidationErrorStyle::Information => "information",
            }
            .into(),
            error_title: dv.error_title.clone(),
            error_message: dv.error_message.clone(),
            operator: operator.map(Into::into),
            value1,
            value2,
            list_source,
            formula,
        }
    }
}

/// A conditional formatting rule.
#[napi(object)]
pub struct JsConditionalFormatRule {
    /// The type of rule: `"cellIs"`, `"expression"`, `"colorScale"`, `"dataBar"`,
    /// `"iconSet"`, `"top10"`, `"aboveAverage"`, `"containsText"`, `"beginsWith"`,
    /// `"endsWith"`, `"duplicateValues"`, `"uniqueValues"`, `"containsBlanks"`,
    /// `"notContainsBlanks"`, `"containsErrors"`, `"notContainsErrors"`, `"timePeriod"`.
    pub rule_type: String,
    /// Ranges this rule applies to (as range strings).
    pub ranges: Vec<String>,
    /// Lower number = higher priority.
    pub priority: u32,
    pub stop_if_true: bool,
    /// Operator (present for `cellIs` rules).
    pub operator: Option<String>,
    /// Formula/value (present for `cellIs`, `expression` rules).
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    /// Text value (present for `containsText`, `beginsWith`, `endsWith`).
    pub text: Option<String>,
    /// Rank for top/bottom N rules.
    pub rank: Option<u32>,
    /// Whether the top/bottom N rule uses percentages.
    pub percent: Option<bool>,
    /// Whether it's a "bottom N" rule (vs top N).
    pub bottom: Option<bool>,
}

impl From<&core::ConditionalFormatRule> for JsConditionalFormatRule {
    fn from(r: &core::ConditionalFormatRule) -> Self {
        let (rule_type, operator, formula1, formula2, text, rank, percent, bottom) =
            match &r.rule_type {
                core::CfRuleType::CellIs {
                    operator,
                    formula1,
                    formula2,
                } => {
                    let op = match operator {
                        core::CfOperator::Between => "between",
                        core::CfOperator::NotBetween => "notBetween",
                        core::CfOperator::Equal => "equal",
                        core::CfOperator::NotEqual => "notEqual",
                        core::CfOperator::GreaterThan => "greaterThan",
                        core::CfOperator::LessThan => "lessThan",
                        core::CfOperator::GreaterThanOrEqual => "greaterThanOrEqual",
                        core::CfOperator::LessThanOrEqual => "lessThanOrEqual",
                    };
                    (
                        "cellIs",
                        Some(op),
                        Some(formula1.clone()),
                        formula2.clone(),
                        None,
                        None,
                        None,
                        None,
                    )
                }
                core::CfRuleType::Expression { formula } => (
                    "expression",
                    None,
                    Some(formula.clone()),
                    None,
                    None,
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::ColorScale { .. } => {
                    ("colorScale", None, None, None, None, None, None, None)
                }
                core::CfRuleType::DataBar { .. } => {
                    ("dataBar", None, None, None, None, None, None, None)
                }
                core::CfRuleType::IconSet { .. } => {
                    ("iconSet", None, None, None, None, None, None, None)
                }
                core::CfRuleType::Top10 {
                    rank,
                    percent,
                    bottom,
                } => (
                    "top10",
                    None,
                    None,
                    None,
                    None,
                    Some(*rank),
                    Some(*percent),
                    Some(*bottom),
                ),
                core::CfRuleType::AboveAverage { .. } => {
                    ("aboveAverage", None, None, None, None, None, None, None)
                }
                core::CfRuleType::ContainsText { text } => (
                    "containsText",
                    None,
                    None,
                    None,
                    Some(text.clone()),
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::BeginsWith { text } => (
                    "beginsWith",
                    None,
                    None,
                    None,
                    Some(text.clone()),
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::EndsWith { text } => (
                    "endsWith",
                    None,
                    None,
                    None,
                    Some(text.clone()),
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::DuplicateValues => {
                    ("duplicateValues", None, None, None, None, None, None, None)
                }
                core::CfRuleType::UniqueValues => {
                    ("uniqueValues", None, None, None, None, None, None, None)
                }
                core::CfRuleType::ContainsBlanks => {
                    ("containsBlanks", None, None, None, None, None, None, None)
                }
                core::CfRuleType::NotContainsBlanks => (
                    "notContainsBlanks",
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::ContainsErrors => {
                    ("containsErrors", None, None, None, None, None, None, None)
                }
                core::CfRuleType::NotContainsErrors => (
                    "notContainsErrors",
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::TimePeriod { .. } => {
                    ("timePeriod", None, None, None, None, None, None, None)
                }
            };

        JsConditionalFormatRule {
            rule_type: rule_type.into(),
            ranges: r.ranges.iter().map(|rng| rng.to_string()).collect(),
            priority: r.priority,
            stop_if_true: r.stop_if_true,
            operator: operator.map(Into::into),
            formula1,
            formula2,
            text,
            rank,
            percent,
            bottom,
        }
    }
}

/// A single run of rich text.
#[napi(object)]
pub struct JsRichTextRun {
    pub text: String,
    pub font: Option<JsRunFont>,
}

impl From<&core::RichTextRun> for JsRichTextRun {
    fn from(r: &core::RichTextRun) -> Self {
        JsRichTextRun {
            text: r.text.clone(),
            font: r.font.as_ref().map(JsRunFont::from),
        }
    }
}

/// Font properties for a rich text run (all fields optional - unset inherits cell style).
#[napi(object)]
pub struct JsRunFont {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub size: Option<f64>,
    pub color: Option<JsColor>,
    pub name: Option<String>,
    pub underline: Option<String>,
    pub strikethrough: Option<bool>,
    pub vertical_align: Option<String>,
}

impl From<&core::RunFont> for JsRunFont {
    fn from(f: &core::RunFont) -> Self {
        JsRunFont {
            bold: f.bold,
            italic: f.italic,
            size: f.size,
            color: f.color.as_ref().map(JsColor::from),
            name: f.name.clone(),
            underline: f.underline.as_ref().map(|u| underline_to_string(u).into()),
            strikethrough: f.strikethrough,
            vertical_align: f
                .vertical_align
                .as_ref()
                .map(|v| font_valign_to_string(v).into()),
        }
    }
}

/// A hyperlink with its cell address.
#[napi(object)]
pub struct JsHyperlinkEntry {
    pub address: String,
    pub hyperlink: JsHyperlink,
}

/// A formula cell with address.
#[napi(object)]
pub struct JsFormulaCell {
    pub row: u32,
    pub col: u32,
    pub formula: String,
}

/// A cell with address and value.
#[napi(object)]
pub struct JsSpillSource {
    pub row: u32,
    pub col: u32,
}

/// A merged cell region with structured coordinates.
#[napi(object)]
pub struct JsMergedRegion {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
    /// The range as an A1-style string (e.g., "A1:C3").
    pub range: String,
}

/// The row/column span of a merged region's origin cell.
#[napi(object)]
pub struct JsMergeSpan {
    pub row_span: u32,
    pub col_span: u32,
}

/// IMAGE() metadata captured during calculation.
#[napi(object)]
pub struct JsImageInfo {
    /// IMAGE source URL or path.
    pub source: String,
    /// IMAGE alternate text.
    pub alt_text: String,
    /// 0=FitCell, 1=FillCell, 2=OriginalSize, 3=Custom
    pub sizing: u32,
    /// Optional custom width.
    pub width: Option<f64>,
    /// Optional custom height.
    pub height: Option<f64>,
}


/// Chart anchor position in a worksheet.
#[napi(object)]
pub struct JsDrawingAnchor {
    pub from_col: u32,
    pub from_row: u32,
    pub from_col_offset: i64,
    pub from_row_offset: i64,
    pub to_col: u32,
    pub to_row: u32,
    pub to_col_offset: i64,
    pub to_row_offset: i64,
}

impl From<&duke_sheets_chart::DrawingAnchor> for JsDrawingAnchor {
    fn from(a: &duke_sheets_chart::DrawingAnchor) -> Self {
        match a {
            duke_sheets_chart::DrawingAnchor::TwoCell { from, to, .. } => JsDrawingAnchor {
                from_col: from.col as u32,
                from_row: from.row,
                from_col_offset: from.col_offset_emu,
                from_row_offset: from.row_offset_emu,
                to_col: to.col as u32,
                to_row: to.row,
                to_col_offset: to.col_offset_emu,
                to_row_offset: to.row_offset_emu,
            },
            _ => JsDrawingAnchor {
                from_col: 0, from_row: 0, from_col_offset: 0, from_row_offset: 0,
                to_col: 0, to_row: 0, to_col_offset: 0, to_row_offset: 0,
            },
        }
    }
}

/// Reference to chart data.
#[napi(object)]
pub struct JsDataReference {
    /// One of: `"formula"`, `"numbers"`, `"strings"`.
    pub ref_type: String,
    pub formula: Option<String>,
    pub numbers: Option<Vec<f64>>,
    pub strings: Option<Vec<String>>,
}

impl From<&duke_sheets_chart::DataReference> for JsDataReference {
    fn from(r: &duke_sheets_chart::DataReference) -> Self {
        match r {
            duke_sheets_chart::DataReference::Formula(f) => JsDataReference {
                ref_type: "formula".into(),
                formula: Some(f.clone()),
                numbers: None,
                strings: None,
            },
            duke_sheets_chart::DataReference::Numbers(ns) => JsDataReference {
                ref_type: "numbers".into(),
                formula: None,
                numbers: Some(ns.clone()),
                strings: None,
            },
            duke_sheets_chart::DataReference::Strings(ss) => JsDataReference {
                ref_type: "strings".into(),
                formula: None,
                numbers: None,
                strings: Some(ss.clone()),
            },
        }
    }
}

/// A chart number format.
#[napi(object)]
pub struct JsChartNumberFormat {
    pub format_code: String,
    pub source_linked: Option<bool>,
}

impl From<&duke_sheets_chart::NumberFormat> for JsChartNumberFormat {
    fn from(n: &duke_sheets_chart::NumberFormat) -> Self {
        JsChartNumberFormat {
            format_code: n.format_code.clone(),
            source_linked: n.source_linked,
        }
    }
}

/// Shape properties for chart elements.
#[napi(object)]
pub struct JsChartShapeProperties {
    pub solid_fill_hex: Option<String>,
    pub no_fill: bool,
    pub line_width: Option<i64>,
    pub line_color_hex: Option<String>,
    pub line_no_fill: bool,
    pub line_dash_style: Option<String>,
}

impl From<&duke_sheets_chart::ChartShapeProperties> for JsChartShapeProperties {
    fn from(sp: &duke_sheets_chart::ChartShapeProperties) -> Self {
        JsChartShapeProperties {
            solid_fill_hex: sp.solid_fill.as_ref().map(|c| c.hex.clone()),
            no_fill: sp.no_fill,
            line_width: sp.line.as_ref().and_then(|l| l.width),
            line_color_hex: sp.line.as_ref().and_then(|l| l.solid_fill.as_ref().map(|c| c.hex.clone())),
            line_no_fill: sp.line.as_ref().map(|l| l.no_fill).unwrap_or(false),
            line_dash_style: sp.line.as_ref().and_then(|l| l.dash_style.clone()),
        }
    }
}

/// Data labels configuration.
#[napi(object)]
pub struct JsDataLabels {
    pub show_legend_key: Option<bool>,
    pub show_value: Option<bool>,
    pub show_category_name: Option<bool>,
    pub show_series_name: Option<bool>,
    pub show_percent: Option<bool>,
    pub show_bubble_size: Option<bool>,
    pub separator: Option<String>,
    pub position: Option<String>,
    pub number_format: Option<JsChartNumberFormat>,
    pub show_leader_lines: Option<bool>,
}

impl From<&duke_sheets_chart::DataLabels> for JsDataLabels {
    fn from(d: &duke_sheets_chart::DataLabels) -> Self {
        JsDataLabels {
            show_legend_key: d.show_legend_key,
            show_value: d.show_value,
            show_category_name: d.show_category_name,
            show_series_name: d.show_series_name,
            show_percent: d.show_percent,
            show_bubble_size: d.show_bubble_size,
            separator: d.separator.clone(),
            position: d.position.as_ref().map(|p| format!("{:?}", p)),
            number_format: d.number_format.as_ref().map(JsChartNumberFormat::from),
            show_leader_lines: d.show_leader_lines,
        }
    }
}

/// A trendline attached to a data series.
#[napi(object)]
pub struct JsTrendline {
    pub trendline_type: String,
    pub name: Option<String>,
    pub order: Option<u32>,
    pub period: Option<u32>,
    pub forward: Option<f64>,
    pub backward: Option<f64>,
    pub intercept: Option<f64>,
    pub display_r_squared: Option<bool>,
    pub display_equation: Option<bool>,
}

impl From<&duke_sheets_chart::Trendline> for JsTrendline {
    fn from(t: &duke_sheets_chart::Trendline) -> Self {
        JsTrendline {
            trendline_type: format!("{:?}", t.trendline_type),
            name: t.name.clone(),
            order: t.order,
            period: t.period,
            forward: t.forward,
            backward: t.backward,
            intercept: t.intercept,
            display_r_squared: t.display_r_squared,
            display_equation: t.display_equation,
        }
    }
}

/// Error bars attached to a data series.
#[napi(object)]
pub struct JsErrorBars {
    pub direction: String,
    pub bar_type: String,
    pub value_type: String,
    pub value: Option<f64>,
    pub no_end_cap: Option<bool>,
}

impl From<&duke_sheets_chart::ErrorBars> for JsErrorBars {
    fn from(e: &duke_sheets_chart::ErrorBars) -> Self {
        JsErrorBars {
            direction: format!("{:?}", e.direction),
            bar_type: format!("{:?}", e.bar_type),
            value_type: format!("{:?}", e.value_type),
            value: e.value,
            no_end_cap: e.no_end_cap,
        }
    }
}

/// Marker for a data point.
#[napi(object)]
pub struct JsMarker {
    pub symbol: Option<String>,
    pub size: Option<u32>,
}

impl From<&duke_sheets_chart::Marker> for JsMarker {
    fn from(m: &duke_sheets_chart::Marker) -> Self {
        JsMarker {
            symbol: m.symbol.as_ref().map(|s| format!("{:?}", s)),
            size: m.size.map(|s| s as u32),
        }
    }
}

/// An individual data point override.
#[napi(object)]
pub struct JsDataPoint {
    pub index: u32,
    pub marker: Option<JsMarker>,
    pub explosion: Option<u32>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::DataPoint> for JsDataPoint {
    fn from(p: &duke_sheets_chart::DataPoint) -> Self {
        JsDataPoint {
            index: p.index,
            marker: p.marker.as_ref().map(JsMarker::from),
            explosion: p.explosion,
            shape_properties: p.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

/// A chart data series.
#[napi(object)]
pub struct JsDataSeries {
    pub name: Option<String>,
    pub values: JsDataReference,
    pub categories: Option<JsDataReference>,
    pub data_labels: Option<JsDataLabels>,
    pub trendline: Option<JsTrendline>,
    pub error_bars: Option<JsErrorBars>,
    pub marker: Option<JsMarker>,
    pub data_points: Vec<JsDataPoint>,
    pub smooth: Option<bool>,
    pub explosion: Option<u32>,
    pub invert_if_negative: Option<bool>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::DataSeries> for JsDataSeries {
    fn from(s: &duke_sheets_chart::DataSeries) -> Self {
        JsDataSeries {
            name: s.name.clone(),
            values: JsDataReference::from(&s.values),
            categories: s.categories.as_ref().map(JsDataReference::from),
            data_labels: s.data_labels.as_ref().map(JsDataLabels::from),
            trendline: s.trendline.as_ref().map(JsTrendline::from),
            error_bars: s.error_bars.as_ref().map(JsErrorBars::from),
            marker: s.marker.as_ref().map(JsMarker::from),
            data_points: s.data_points.iter().map(JsDataPoint::from).collect(),
            smooth: s.smooth,
            explosion: s.explosion,
            invert_if_negative: s.invert_if_negative,
            shape_properties: s.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

/// A chart axis.
#[napi(object)]
pub struct JsAxis {
    pub title: Option<String>,
    pub minimum: Option<f64>,
    pub maximum: Option<f64>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
    /// One of: `"Bottom"`, `"Top"`, `"Left"`, `"Right"`.
    pub position: String,
    pub number_format: Option<JsChartNumberFormat>,
    pub major_gridlines: bool,
    pub minor_gridlines: bool,
    pub major_gridlines_shape_properties: Option<JsChartShapeProperties>,
    pub minor_gridlines_shape_properties: Option<JsChartShapeProperties>,
    pub major_tick_mark: Option<String>,
    pub minor_tick_mark: Option<String>,
    pub label_position: Option<String>,
    pub delete: Option<bool>,
    pub crosses: Option<String>,
    pub cross_between: Option<String>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::Axis> for JsAxis {
    fn from(a: &duke_sheets_chart::Axis) -> Self {
        JsAxis {
            title: a.title.clone(),
            minimum: a.minimum,
            maximum: a.maximum,
            major_unit: a.major_unit,
            minor_unit: a.minor_unit,
            position: format!("{:?}", a.position),
            number_format: a.number_format.as_ref().map(JsChartNumberFormat::from),
            major_gridlines: a.major_gridlines,
            minor_gridlines: a.minor_gridlines,
            major_gridlines_shape_properties: a.major_gridlines_shape_properties.as_ref().map(JsChartShapeProperties::from),
            minor_gridlines_shape_properties: a.minor_gridlines_shape_properties.as_ref().map(JsChartShapeProperties::from),
            major_tick_mark: a.major_tick_mark.as_ref().map(|t| format!("{:?}", t)),
            minor_tick_mark: a.minor_tick_mark.as_ref().map(|t| format!("{:?}", t)),
            label_position: a.label_position.as_ref().map(|p| format!("{:?}", p)),
            delete: a.delete,
            crosses: a.crosses.as_ref().map(|c| format!("{:?}", c)),
            cross_between: a.cross_between.as_ref().map(|c| format!("{:?}", c)),
            shape_properties: a.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

/// A chart legend.
#[napi(object)]
pub struct JsLegend {
    /// One of: `"Right"`, `"Top"`, `"Bottom"`, `"Left"`, `"TopRight"`.
    pub position: String,
    pub overlay: bool,
}

impl From<&duke_sheets_chart::Legend> for JsLegend {
    fn from(l: &duke_sheets_chart::Legend) -> Self {
        JsLegend {
            position: format!("{:?}", l.position),
            overlay: l.overlay,
        }
    }
}

/// 3D view settings.
#[napi(object)]
pub struct JsView3D {
    pub rotate_x: Option<i32>,
    pub rotate_y: Option<i32>,
    pub depth_percent: Option<u32>,
    pub height_percent: Option<u32>,
    pub perspective: Option<u32>,
    pub right_angle_axes: Option<bool>,
}

impl From<&duke_sheets_chart::View3D> for JsView3D {
    fn from(v: &duke_sheets_chart::View3D) -> Self {
        JsView3D {
            rotate_x: v.rotate_x,
            rotate_y: v.rotate_y,
            depth_percent: v.depth_percent,
            height_percent: v.height_percent,
            perspective: v.perspective,
            right_angle_axes: v.right_angle_axes,
        }
    }
}

/// Data table displayed beneath the chart.
#[napi(object)]
pub struct JsChartDataTable {
    pub show_horizontal_border: Option<bool>,
    pub show_vertical_border: Option<bool>,
    pub show_outline: Option<bool>,
    pub show_keys: Option<bool>,
}

impl From<&duke_sheets_chart::ChartDataTable> for JsChartDataTable {
    fn from(t: &duke_sheets_chart::ChartDataTable) -> Self {
        JsChartDataTable {
            show_horizontal_border: t.show_horizontal_border,
            show_vertical_border: t.show_vertical_border,
            show_outline: t.show_outline,
            show_keys: t.show_keys,
        }
    }
}

/// Manual layout positioning.
#[napi(object)]
pub struct JsManualLayout {
    pub x: Option<f64>,
    pub y: Option<f64>,
    pub width: Option<f64>,
    pub height: Option<f64>,
}

impl From<&duke_sheets_chart::ManualLayout> for JsManualLayout {
    fn from(m: &duke_sheets_chart::ManualLayout) -> Self {
        JsManualLayout {
            x: m.x,
            y: m.y,
            width: m.width,
            height: m.height,
        }
    }
}

/// Layout container.
#[napi(object)]
pub struct JsLayout {
    pub manual_layout: Option<JsManualLayout>,
}

impl From<&duke_sheets_chart::Layout> for JsLayout {
    fn from(l: &duke_sheets_chart::Layout) -> Self {
        JsLayout {
            manual_layout: l.manual_layout.as_ref().map(JsManualLayout::from),
        }
    }
}

#[napi(object)]
pub struct JsChartTypeGroup {
    pub chart_type: String,
    pub is_3d: bool,
    pub series: Vec<JsDataSeries>,
    pub data_labels: Option<JsDataLabels>,
    pub vary_colors: Option<bool>,
    pub gap_width: Option<u32>,
    pub overlap: Option<i32>,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
    pub bubble_scale: Option<u32>,
    pub show_negative_bubbles: Option<bool>,
    pub radar_style: Option<String>,
    pub wireframe: Option<bool>,
    pub axis_ids: Vec<u32>,
    pub drop_lines: Option<JsChartLines>,
    pub high_low_lines: Option<JsChartLines>,
    pub series_lines: Option<JsChartLines>,
    pub up_down_bars: Option<JsUpDownBars>,
}

impl From<&duke_sheets_chart::ChartTypeGroup> for JsChartTypeGroup {
    fn from(g: &duke_sheets_chart::ChartTypeGroup) -> Self {
        Self {
            chart_type: format!("{:?}", g.chart_type),
            is_3d: g.is_3d,
            series: g.series.iter().map(JsDataSeries::from).collect(),
            data_labels: g.data_labels.as_ref().map(JsDataLabels::from),
            vary_colors: g.vary_colors,
            gap_width: g.gap_width,
            overlap: g.overlap,
            first_slice_angle: g.first_slice_angle,
            hole_size: g.hole_size,
            bubble_scale: g.bubble_scale,
            show_negative_bubbles: g.show_negative_bubbles,
            radar_style: g.radar_style.clone(),
            wireframe: g.wireframe,
            axis_ids: g.axis_ids.clone(),
            drop_lines: g.drop_lines.as_ref().map(JsChartLines::from),
            high_low_lines: g.high_low_lines.as_ref().map(JsChartLines::from),
            series_lines: g.series_lines.as_ref().map(JsChartLines::from),
            up_down_bars: g.up_down_bars.as_ref().map(JsUpDownBars::from),
        }
    }
}

#[napi(object)]
pub struct JsChartAxis {
    pub id: u32,
    pub cross_id: u32,
    pub axis: JsAxis,
}

impl From<&duke_sheets_chart::ChartAxis> for JsChartAxis {
    fn from(a: &duke_sheets_chart::ChartAxis) -> Self {
        Self {
            id: a.id,
            cross_id: a.cross_id,
            axis: JsAxis::from(&a.axis),
        }
    }
}

#[napi(object)]
pub struct JsPivotChartSource {
    pub name: String,
    pub format_id: u32,
}

impl From<&duke_sheets_chart::PivotChartSource> for JsPivotChartSource {
    fn from(s: &duke_sheets_chart::PivotChartSource) -> Self {
        Self {
            name: s.name.clone(),
            format_id: s.format_id,
        }
    }
}

/// A chart embedded in a worksheet.
#[napi(object)]
pub struct JsChart {
    /// Chart type string, e.g. `"ColumnClustered"`, `"Line"`, `"Pie"`.
    pub chart_type: String,
    pub title: Option<String>,
    pub series: Vec<JsDataSeries>,
    pub category_axis: Option<JsAxis>,
    pub value_axis: Option<JsAxis>,
    pub legend: Option<JsLegend>,
    pub anchor: JsDrawingAnchor,
    pub data_labels: Option<JsDataLabels>,
    pub view_3d: Option<JsView3D>,
    pub data_table: Option<JsChartDataTable>,
    pub display_blanks_as: Option<String>,
    pub plot_visible_only: Option<bool>,
    pub layout: Option<JsLayout>,
    pub shape_properties: Option<JsChartShapeProperties>,
    pub is_3d: bool,
    pub vary_colors: Option<bool>,
    pub gap_width: Option<u32>,
    pub overlap: Option<i32>,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
    pub bubble_scale: Option<u32>,
    pub show_negative_bubbles: Option<bool>,
    pub auto_title_deleted: Option<bool>,
    pub rounded_corners: Option<bool>,
    pub pivot_source: Option<JsPivotChartSource>,
    pub show_dlbls_over_max: Option<bool>,
    pub wireframe: Option<bool>,
    pub radar_style: Option<String>,
    pub type_groups: Vec<JsChartTypeGroup>,
    pub axes: Vec<JsChartAxis>,
    pub drop_lines: Option<JsChartLines>,
    pub high_low_lines: Option<JsChartLines>,
    pub series_lines: Option<JsChartLines>,
    pub up_down_bars: Option<JsUpDownBars>,
}

impl From<&duke_sheets_chart::Chart> for JsChart {
    fn from(c: &duke_sheets_chart::Chart) -> Self {
        let chart_type = match &c.chart_type {
            duke_sheets_chart::ChartType::Unsupported(tag) => format!("Unsupported({})", tag),
            other => format!("{:?}", other),
        };
        JsChart {
            chart_type,
            title: c.title.clone(),
            series: c.series.iter().map(JsDataSeries::from).collect(),
            category_axis: c.category_axis.as_ref().map(JsAxis::from),
            value_axis: c.value_axis.as_ref().map(JsAxis::from),
            legend: c.legend.as_ref().map(JsLegend::from),
            anchor: JsDrawingAnchor::from(&c.anchor),
            data_labels: c.data_labels.as_ref().map(JsDataLabels::from),
            view_3d: c.view_3d.as_ref().map(JsView3D::from),
            data_table: c.data_table.as_ref().map(JsChartDataTable::from),
            display_blanks_as: c.display_blanks_as.as_ref().map(|d| format!("{:?}", d)),
            plot_visible_only: c.plot_visible_only,
            layout: c.layout.as_ref().map(JsLayout::from),
            shape_properties: c.shape_properties.as_ref().map(JsChartShapeProperties::from),
            is_3d: c.is_3d,
            vary_colors: c.vary_colors,
            gap_width: c.gap_width,
            overlap: c.overlap,
            first_slice_angle: c.first_slice_angle,
            hole_size: c.hole_size,
            bubble_scale: c.bubble_scale,
            show_negative_bubbles: c.show_negative_bubbles,
            auto_title_deleted: c.auto_title_deleted,
            rounded_corners: c.rounded_corners,
            pivot_source: c.pivot_source.as_ref().map(JsPivotChartSource::from),
            show_dlbls_over_max: c.show_dlbls_over_max,
            wireframe: c.wireframe,
            radar_style: c.radar_style.clone(),
            type_groups: c.type_groups.iter().map(JsChartTypeGroup::from).collect(),
            axes: c.axes.iter().map(JsChartAxis::from).collect(),
            drop_lines: c.drop_lines.as_ref().map(JsChartLines::from),
            high_low_lines: c.high_low_lines.as_ref().map(JsChartLines::from),
            series_lines: c.series_lines.as_ref().map(JsChartLines::from),
            up_down_bars: c.up_down_bars.as_ref().map(JsUpDownBars::from),
        }
    }
}

/// Chart line overlay (drop lines, high-low lines, series lines).
#[napi(object)]
pub struct JsChartLines {
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartLines> for JsChartLines {
    fn from(cl: &duke_sheets_chart::ChartLines) -> Self {
        Self {
            shape_properties: cl.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

/// Up-down bars (stock charts).
#[napi(object)]
pub struct JsUpDownBars {
    pub gap_width: Option<u32>,
    pub up_bars: Option<JsChartLines>,
    pub down_bars: Option<JsChartLines>,
}

impl From<&duke_sheets_chart::UpDownBars> for JsUpDownBars {
    fn from(ud: &duke_sheets_chart::UpDownBars) -> Self {
        Self {
            gap_width: ud.gap_width,
            up_bars: ud.up_bars.as_ref().map(JsChartLines::from),
            down_bars: ud.down_bars.as_ref().map(JsChartLines::from),
        }
    }
}

/// A chart sheet - a sheet that contains only a chart.
#[napi(object)]
pub struct JsChartSheet {
    pub name: String,
    pub chart: JsChart,
    pub visibility: String,
}

impl From<&core::ChartSheet> for JsChartSheet {
    fn from(cs: &core::ChartSheet) -> Self {
        Self {
            name: cs.name.clone(),
            chart: JsChart::from(&cs.chart),
            visibility: match cs.visibility {
                core::worksheet::SheetVisibility::Visible => "visible",
                core::worksheet::SheetVisibility::Hidden => "hidden",
                core::worksheet::SheetVisibility::VeryHidden => "veryHidden",
            }
            .into(),
        }
    }
}

/// A slot in the workbook tab bar.
#[napi(object)]
pub struct JsSheetSlot {
    /// `"worksheet"` or `"chartsheet"`.
    pub slot_type: String,
    /// Index into the respective collection.
    pub index: u32,
}

impl From<&core::SheetSlot> for JsSheetSlot {
    fn from(slot: &core::SheetSlot) -> Self {
        match slot {
            core::SheetSlot::Worksheet(idx) => Self {
                slot_type: "worksheet".into(),
                index: *idx as u32,
            },
            core::SheetSlot::ChartSheet(idx) => Self {
                slot_type: "chartsheet".into(),
                index: *idx as u32,
            },
        }
    }
}

fn chart_ex_layout_to_string(layout: &duke_sheets_chart::ChartExLayout) -> &'static str {
    match layout {
        duke_sheets_chart::ChartExLayout::Waterfall => "waterfall",
        duke_sheets_chart::ChartExLayout::Treemap => "treemap",
        duke_sheets_chart::ChartExLayout::Sunburst => "sunburst",
        duke_sheets_chart::ChartExLayout::Funnel => "funnel",
        duke_sheets_chart::ChartExLayout::Histogram => "histogram",
        duke_sheets_chart::ChartExLayout::BoxWhisker => "boxWhisker",
        duke_sheets_chart::ChartExLayout::ParetoLine => "paretoLine",
        duke_sheets_chart::ChartExLayout::RegionMap => "regionMap",
        duke_sheets_chart::ChartExLayout::ClusteredColumn => "clusteredColumn",
        duke_sheets_chart::ChartExLayout::Unknown(_) => "unknown",
    }
}

#[napi(object)]
pub struct JsChartExOffset {
    pub top: Option<f64>,
    pub left: Option<f64>,
}

impl From<&duke_sheets_chart::ChartExOffset> for JsChartExOffset {
    fn from(o: &duke_sheets_chart::ChartExOffset) -> Self {
        Self {
            top: o.top,
            left: o.left,
        }
    }
}

#[napi(object)]
pub struct JsChartExText {
    pub formula: Option<String>,
    pub value: Option<String>,
}

impl From<&duke_sheets_chart::ChartExText> for JsChartExText {
    fn from(t: &duke_sheets_chart::ChartExText) -> Self {
        Self {
            formula: t.data.as_ref().and_then(|d| d.formula.clone()),
            value: t.data.as_ref().and_then(|d| d.value.clone()),
        }
    }
}

#[napi(object)]
pub struct JsChartExColorPosition {
    pub position_type: String,
    pub value: Option<f64>,
}

impl From<&duke_sheets_chart::ChartExColorPosition> for JsChartExColorPosition {
    fn from(p: &duke_sheets_chart::ChartExColorPosition) -> Self {
        match p {
            duke_sheets_chart::ChartExColorPosition::ExtremeValue => Self {
                position_type: "extremeValue".into(),
                value: None,
            },
            duke_sheets_chart::ChartExColorPosition::Number(v) => Self {
                position_type: "number".into(),
                value: Some(*v),
            },
            duke_sheets_chart::ChartExColorPosition::Percent(v) => Self {
                position_type: "percent".into(),
                value: Some(*v),
            },
        }
    }
}

#[napi(object)]
pub struct JsChartExValueColorPositions {
    pub count: Option<u32>,
    pub min: Option<JsChartExColorPosition>,
    pub mid: Option<JsChartExColorPosition>,
    pub max: Option<JsChartExColorPosition>,
}

impl From<&duke_sheets_chart::ChartExValueColorPositions> for JsChartExValueColorPositions {
    fn from(p: &duke_sheets_chart::ChartExValueColorPositions) -> Self {
        Self {
            count: p.count,
            min: p.min.as_ref().map(JsChartExColorPosition::from),
            mid: p.mid.as_ref().map(JsChartExColorPosition::from),
            max: p.max.as_ref().map(JsChartExColorPosition::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExScaling {
    pub scaling_type: String,
    pub gap_width: Option<f64>,
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
}

impl From<&duke_sheets_chart::ChartExScaling> for JsChartExScaling {
    fn from(s: &duke_sheets_chart::ChartExScaling) -> Self {
        match s {
            duke_sheets_chart::ChartExScaling::Category { gap_width } => Self {
                scaling_type: "category".into(),
                gap_width: *gap_width,
                min: None,
                max: None,
                major_unit: None,
                minor_unit: None,
            },
            duke_sheets_chart::ChartExScaling::Value { min, max, major_unit, minor_unit } => Self {
                scaling_type: "value".into(),
                gap_width: None,
                min: *min,
                max: *max,
                major_unit: *major_unit,
                minor_unit: *minor_unit,
            },
        }
    }
}

#[napi(object)]
pub struct JsChartExAxisTitle {
    pub text: Option<String>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExAxisTitle> for JsChartExAxisTitle {
    fn from(t: &duke_sheets_chart::ChartExAxisTitle) -> Self {
        Self {
            text: t.text.as_ref().and_then(|tx| {
                tx.data.as_ref().and_then(|d| d.value.clone().or_else(|| d.formula.clone()))
            }),
            shape_properties: t.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExAxisUnits {
    pub unit: Option<String>,
}

impl From<&duke_sheets_chart::ChartExAxisUnits> for JsChartExAxisUnits {
    fn from(u: &duke_sheets_chart::ChartExAxisUnits) -> Self {
        Self {
            unit: u.unit.clone(),
        }
    }
}

#[napi(object)]
pub struct JsChartExSeriesVisibility {
    pub connector_lines: Option<bool>,
    pub mean_line: Option<bool>,
    pub mean_marker: Option<bool>,
    pub nonoutliers: Option<bool>,
    pub outliers: Option<bool>,
}

impl From<&duke_sheets_chart::ChartExSeriesVisibility> for JsChartExSeriesVisibility {
    fn from(v: &duke_sheets_chart::ChartExSeriesVisibility) -> Self {
        Self {
            connector_lines: v.connector_lines,
            mean_line: v.mean_line,
            mean_marker: v.mean_marker,
            nonoutliers: v.nonoutliers,
            outliers: v.outliers,
        }
    }
}

#[napi(object)]
pub struct JsChartExBinning {
    pub interval_closed: Option<String>,
    pub underflow: Option<String>,
    pub overflow: Option<String>,
    pub bin_size: Option<f64>,
    pub bin_count: Option<u32>,
}

impl From<&duke_sheets_chart::ChartExBinning> for JsChartExBinning {
    fn from(b: &duke_sheets_chart::ChartExBinning) -> Self {
        Self {
            interval_closed: b.interval_closed.clone(),
            underflow: b.underflow.clone(),
            overflow: b.overflow.clone(),
            bin_size: b.bin_size,
            bin_count: b.bin_count,
        }
    }
}

#[napi(object)]
pub struct JsChartExGeography {
    pub projection_type: Option<String>,
    pub viewed_region_type: Option<String>,
    pub culture_language: Option<String>,
    pub culture_region: Option<String>,
    pub attribution: Option<String>,
}

impl From<&duke_sheets_chart::ChartExGeography> for JsChartExGeography {
    fn from(g: &duke_sheets_chart::ChartExGeography) -> Self {
        Self {
            projection_type: g.projection_type.clone(),
            viewed_region_type: g.viewed_region_type.clone(),
            culture_language: g.culture_language.clone(),
            culture_region: g.culture_region.clone(),
            attribution: g.attribution.clone(),
        }
    }
}

#[napi(object)]
pub struct JsChartExStatistics {
    pub quartile_method: Option<String>,
}

impl From<&duke_sheets_chart::ChartExStatistics> for JsChartExStatistics {
    fn from(s: &duke_sheets_chart::ChartExStatistics) -> Self {
        Self {
            quartile_method: s.quartile_method.clone(),
        }
    }
}

#[napi(object)]
pub struct JsChartExDataPoint {
    pub idx: u32,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExDataPoint> for JsChartExDataPoint {
    fn from(p: &duke_sheets_chart::ChartExDataPoint) -> Self {
        Self {
            idx: p.idx,
            shape_properties: p.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExDataLabel {
    pub idx: u32,
    pub position: Option<String>,
    pub visibility_series_name: Option<bool>,
    pub visibility_category_name: Option<bool>,
    pub visibility_value: Option<bool>,
    pub number_format: Option<JsChartNumberFormat>,
    pub separator: Option<String>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExDataLabel> for JsChartExDataLabel {
    fn from(l: &duke_sheets_chart::ChartExDataLabel) -> Self {
        Self {
            idx: l.idx,
            position: l.position.clone(),
            visibility_series_name: l.visibility_series_name,
            visibility_category_name: l.visibility_category_name,
            visibility_value: l.visibility_value,
            number_format: l.number_format.as_ref().map(JsChartNumberFormat::from),
            separator: l.separator.clone(),
            shape_properties: l.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExFormatOverride {
    pub idx: u32,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExFormatOverride> for JsChartExFormatOverride {
    fn from(o: &duke_sheets_chart::ChartExFormatOverride) -> Self {
        Self {
            idx: o.idx,
            shape_properties: o.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExHeaderFooter {
    pub align_with_margins: Option<bool>,
    pub different_odd_even: Option<bool>,
    pub different_first: Option<bool>,
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
}

impl From<&duke_sheets_chart::ChartExHeaderFooter> for JsChartExHeaderFooter {
    fn from(h: &duke_sheets_chart::ChartExHeaderFooter) -> Self {
        Self {
            align_with_margins: h.align_with_margins,
            different_odd_even: h.different_odd_even,
            different_first: h.different_first,
            odd_header: h.odd_header.clone(),
            odd_footer: h.odd_footer.clone(),
            even_header: h.even_header.clone(),
            even_footer: h.even_footer.clone(),
            first_header: h.first_header.clone(),
            first_footer: h.first_footer.clone(),
        }
    }
}

#[napi(object)]
pub struct JsChartExPageMargins {
    pub left: Option<f64>,
    pub right: Option<f64>,
    pub top: Option<f64>,
    pub bottom: Option<f64>,
    pub header: Option<f64>,
    pub footer: Option<f64>,
}

impl From<&duke_sheets_chart::ChartExPageMargins> for JsChartExPageMargins {
    fn from(m: &duke_sheets_chart::ChartExPageMargins) -> Self {
        Self {
            left: m.left,
            right: m.right,
            top: m.top,
            bottom: m.bottom,
            header: m.header,
            footer: m.footer,
        }
    }
}

#[napi(object)]
pub struct JsChartExPageSetup {
    pub paper_size: Option<u32>,
    pub first_page_number: Option<u32>,
    pub orientation: Option<String>,
    pub black_and_white: Option<bool>,
    pub draft: Option<bool>,
    pub use_first_page_number: Option<bool>,
    pub horizontal_dpi: Option<u32>,
    pub vertical_dpi: Option<u32>,
    pub copies: Option<u32>,
}

impl From<&duke_sheets_chart::ChartExPageSetup> for JsChartExPageSetup {
    fn from(p: &duke_sheets_chart::ChartExPageSetup) -> Self {
        Self {
            paper_size: p.paper_size,
            first_page_number: p.first_page_number,
            orientation: p.orientation.clone(),
            black_and_white: p.black_and_white,
            draft: p.draft,
            use_first_page_number: p.use_first_page_number,
            horizontal_dpi: p.horizontal_dpi,
            vertical_dpi: p.vertical_dpi,
            copies: p.copies,
        }
    }
}

#[napi(object)]
pub struct JsChartExPrintSettings {
    pub header_footer: Option<JsChartExHeaderFooter>,
    pub page_margins: Option<JsChartExPageMargins>,
    pub page_setup: Option<JsChartExPageSetup>,
}

impl From<&duke_sheets_chart::ChartExPrintSettings> for JsChartExPrintSettings {
    fn from(p: &duke_sheets_chart::ChartExPrintSettings) -> Self {
        Self {
            header_footer: p.header_footer.as_ref().map(JsChartExHeaderFooter::from),
            page_margins: p.page_margins.as_ref().map(JsChartExPageMargins::from),
            page_setup: p.page_setup.as_ref().map(JsChartExPageSetup::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExPlotArea {
    pub plot_surface: Option<JsChartShapeProperties>,
    pub series: Vec<JsChartExSeries>,
    pub axes: Vec<JsChartExAxis>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExPlotArea> for JsChartExPlotArea {
    fn from(p: &duke_sheets_chart::ChartExPlotArea) -> Self {
        Self {
            plot_surface: p.plot_surface.as_ref().map(JsChartShapeProperties::from),
            series: p.series.iter().map(JsChartExSeries::from).collect(),
            axes: p.axes.iter().map(JsChartExAxis::from).collect(),
            shape_properties: p.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExDimension {
    pub dim_type: String,
    pub formula: Option<String>,
    pub nf_formula: Option<String>,
}

impl From<&duke_sheets_chart::ChartExDimension> for JsChartExDimension {
    fn from(d: &duke_sheets_chart::ChartExDimension) -> Self {
        match d {
            duke_sheets_chart::ChartExDimension::String { dim_type, formula, nf_formula, .. } => {
                Self {
                    dim_type: match dim_type {
                        duke_sheets_chart::StringDimType::Cat => "cat".into(),
                        duke_sheets_chart::StringDimType::ColorStr => "colorStr".into(),
                        duke_sheets_chart::StringDimType::EntityId => "entityId".into(),
                    },
                    formula: formula.clone(),
                    nf_formula: nf_formula.clone(),
                }
            }
            duke_sheets_chart::ChartExDimension::Numeric { dim_type, formula, nf_formula, .. } => {
                Self {
                    dim_type: match dim_type {
                        duke_sheets_chart::NumericDimType::Val => "val".into(),
                        duke_sheets_chart::NumericDimType::X => "x".into(),
                        duke_sheets_chart::NumericDimType::Y => "y".into(),
                        duke_sheets_chart::NumericDimType::Size => "size".into(),
                        duke_sheets_chart::NumericDimType::ColorVal => "colorVal".into(),
                    },
                    formula: formula.clone(),
                    nf_formula: nf_formula.clone(),
                }
            }
        }
    }
}

#[napi(object)]
pub struct JsChartExData {
    pub id: u32,
    pub dimensions: Vec<JsChartExDimension>,
}

impl From<&duke_sheets_chart::ChartExData> for JsChartExData {
    fn from(d: &duke_sheets_chart::ChartExData) -> Self {
        Self {
            id: d.id,
            dimensions: d.dimensions.iter().map(JsChartExDimension::from).collect(),
        }
    }
}

#[napi(object)]
pub struct JsChartExDataLabels {
    pub position: Option<String>,
    pub visibility_series_name: Option<bool>,
    pub visibility_category_name: Option<bool>,
    pub visibility_value: Option<bool>,
    pub number_format: Option<JsChartNumberFormat>,
    pub separator: Option<String>,
    pub shape_properties: Option<JsChartShapeProperties>,
    pub overrides: Vec<JsChartExDataLabel>,
    pub hidden_labels: Vec<u32>,
}

impl From<&duke_sheets_chart::ChartExDataLabels> for JsChartExDataLabels {
    fn from(l: &duke_sheets_chart::ChartExDataLabels) -> Self {
        Self {
            position: l.position.clone(),
            visibility_series_name: l.visibility_series_name,
            visibility_category_name: l.visibility_category_name,
            visibility_value: l.visibility_value,
            number_format: l.number_format.as_ref().map(JsChartNumberFormat::from),
            separator: l.separator.clone(),
            shape_properties: l.shape_properties.as_ref().map(JsChartShapeProperties::from),
            overrides: l.overrides.iter().map(JsChartExDataLabel::from).collect(),
            hidden_labels: l.hidden_labels.clone(),
        }
    }
}

#[napi(object)]
pub struct JsChartExTitle {
    pub text: Option<String>,
    pub position: Option<String>,
    pub align: Option<String>,
    pub overlay: Option<bool>,
    pub offset: Option<JsChartExOffset>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExTitle> for JsChartExTitle {
    fn from(t: &duke_sheets_chart::ChartExTitle) -> Self {
        Self {
            text: t.text.clone(),
            position: t.position.clone(),
            align: t.align.clone(),
            overlay: t.overlay,
            offset: t.offset.as_ref().map(JsChartExOffset::from),
            shape_properties: t.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExLegend {
    pub position: Option<String>,
    pub align: Option<String>,
    pub overlay: Option<bool>,
    pub offset: Option<JsChartExOffset>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExLegend> for JsChartExLegend {
    fn from(l: &duke_sheets_chart::ChartExLegend) -> Self {
        Self {
            position: l.position.clone(),
            align: l.align.clone(),
            overlay: l.overlay,
            offset: l.offset.as_ref().map(JsChartExOffset::from),
            shape_properties: l.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExLayoutPr {
    pub parent_label_layout: Option<String>,
    pub region_label_layout: Option<String>,
    pub visibility: Option<JsChartExSeriesVisibility>,
    pub aggregation: bool,
    pub binning: Option<JsChartExBinning>,
    pub geography: Option<JsChartExGeography>,
    pub statistics: Option<JsChartExStatistics>,
    pub subtotals: Vec<u32>,
}

impl From<&duke_sheets_chart::ChartExLayoutPr> for JsChartExLayoutPr {
    fn from(l: &duke_sheets_chart::ChartExLayoutPr) -> Self {
        Self {
            parent_label_layout: l.parent_label_layout.clone(),
            region_label_layout: l.region_label_layout.clone(),
            visibility: l.visibility.as_ref().map(JsChartExSeriesVisibility::from),
            aggregation: l.aggregation,
            binning: l.binning.as_ref().map(JsChartExBinning::from),
            geography: l.geography.as_ref().map(JsChartExGeography::from),
            statistics: l.statistics.as_ref().map(JsChartExStatistics::from),
            subtotals: l.subtotals.clone(),
        }
    }
}

#[napi(object)]
pub struct JsChartExAxis {
    pub id: u32,
    pub hidden: Option<bool>,
    pub scaling: JsChartExScaling,
    pub title: Option<JsChartExAxisTitle>,
    pub units: Option<JsChartExAxisUnits>,
    pub major_gridlines: Option<JsChartShapeProperties>,
    pub minor_gridlines: Option<JsChartShapeProperties>,
    pub major_tick_marks: Option<String>,
    pub minor_tick_marks: Option<String>,
    pub tick_labels: bool,
    pub number_format: Option<JsChartNumberFormat>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExAxis> for JsChartExAxis {
    fn from(a: &duke_sheets_chart::ChartExAxis) -> Self {
        Self {
            id: a.id,
            hidden: a.hidden,
            scaling: JsChartExScaling::from(&a.scaling),
            title: a.title.as_ref().map(JsChartExAxisTitle::from),
            units: a.units.as_ref().map(JsChartExAxisUnits::from),
            major_gridlines: a.major_gridlines.as_ref().map(JsChartShapeProperties::from),
            minor_gridlines: a.minor_gridlines.as_ref().map(JsChartShapeProperties::from),
            major_tick_marks: a.major_tick_marks.clone(),
            minor_tick_marks: a.minor_tick_marks.clone(),
            tick_labels: a.tick_labels,
            number_format: a.number_format.as_ref().map(JsChartNumberFormat::from),
            shape_properties: a.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartExSeries {
    pub layout: String,
    pub data_id: u32,
    pub unique_id: Option<String>,
    pub hidden: Option<bool>,
    pub owner_idx: Option<u32>,
    pub format_idx: Option<u32>,
    pub text: Option<JsChartExText>,
    pub data_labels: Option<JsChartExDataLabels>,
    pub data_points: Vec<JsChartExDataPoint>,
    pub layout_properties: Option<JsChartExLayoutPr>,
    pub axis_ids: Vec<u32>,
    pub value_colors: bool,
    pub value_color_positions: Option<JsChartExValueColorPositions>,
    pub shape_properties: Option<JsChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExSeries> for JsChartExSeries {
    fn from(s: &duke_sheets_chart::ChartExSeries) -> Self {
        Self {
            layout: chart_ex_layout_to_string(&s.layout).into(),
            data_id: s.data_id,
            unique_id: s.unique_id.clone(),
            hidden: s.hidden,
            owner_idx: s.owner_idx,
            format_idx: s.format_idx,
            text: s.text.as_ref().map(JsChartExText::from),
            data_labels: s.data_labels.as_ref().map(JsChartExDataLabels::from),
            data_points: s.data_points.iter().map(JsChartExDataPoint::from).collect(),
            layout_properties: s.layout_properties.as_ref().map(JsChartExLayoutPr::from),
            axis_ids: s.axis_ids.clone(),
            value_colors: s.value_colors.is_some(),
            value_color_positions: s.value_color_positions.as_ref().map(JsChartExValueColorPositions::from),
            shape_properties: s.shape_properties.as_ref().map(JsChartShapeProperties::from),
        }
    }
}

#[napi(object)]
pub struct JsChartEx {
    pub layout: String,
    pub version: Option<String>,
    pub feature_list: Option<String>,
    pub fallback_img: Option<String>,
    pub title: Option<JsChartExTitle>,
    pub data: Vec<JsChartExData>,
    pub plot_area: JsChartExPlotArea,
    pub legend: Option<JsChartExLegend>,
    pub anchor: JsDrawingAnchor,
    pub shape_properties: Option<JsChartShapeProperties>,
    pub format_overrides: Vec<JsChartExFormatOverride>,
    pub print_settings: Option<JsChartExPrintSettings>,
    pub external_data_rel_id: Option<String>,
    pub external_data_auto_update: Option<bool>,
}

impl From<&duke_sheets_chart::ChartEx> for JsChartEx {
    fn from(c: &duke_sheets_chart::ChartEx) -> Self {
        let layout = c
            .plot_area
            .series
            .first()
            .map(|s| chart_ex_layout_to_string(&s.layout))
            .unwrap_or("unknown");
        Self {
            layout: layout.into(),
            version: c.version.clone(),
            feature_list: c.feature_list.clone(),
            fallback_img: c.fallback_img.clone(),
            title: c.title.as_ref().map(JsChartExTitle::from),
            data: c.data.iter().map(JsChartExData::from).collect(),
            plot_area: JsChartExPlotArea::from(&c.plot_area),
            legend: c.legend.as_ref().map(JsChartExLegend::from),
            anchor: JsDrawingAnchor::from(&c.anchor),
            shape_properties: c.shape_properties.as_ref().map(JsChartShapeProperties::from),
            format_overrides: c.format_overrides.iter().map(JsChartExFormatOverride::from).collect(),
            print_settings: c.print_settings.as_ref().map(JsChartExPrintSettings::from),
            external_data_rel_id: c.external_data.as_ref().map(|e| e.rel_id.clone()),
            external_data_auto_update: c.external_data.as_ref().and_then(|e| e.auto_update),
        }
    }
}

#[napi(object)]
pub struct JsEmbeddedImage {
    pub id: u32,
    pub name: String,
    pub description: Option<String>,
    pub anchor: JsDrawingAnchor,
    pub format: String,
    pub media_path: String,
    pub svg_media_path: Option<String>,
    pub width_emu: i64,
    pub height_emu: i64,
    pub rotation: Option<i32>,
    pub flip_h: bool,
    pub flip_v: bool,
    pub data: napi::bindgen_prelude::Buffer,
    pub svg_data: Option<napi::bindgen_prelude::Buffer>,
}

impl From<&duke_sheets_chart::EmbeddedImage> for JsEmbeddedImage {
    fn from(img: &duke_sheets_chart::EmbeddedImage) -> Self {
        JsEmbeddedImage {
            id: img.id,
            name: img.name.clone(),
            description: img.description.clone(),
            anchor: JsDrawingAnchor::from(&img.anchor),
            format: img.format.as_str().to_string(),
            media_path: img.media_path.clone(),
            svg_media_path: img.svg_media_path.clone(),
            width_emu: img.width_emu,
            height_emu: img.height_emu,
            rotation: img.rotation,
            flip_h: img.flip_h,
            flip_v: img.flip_v,
            data: img.data().to_vec().into(),
            svg_data: img.svg_data().map(|b| b.to_vec().into()),
        }
    }
}
