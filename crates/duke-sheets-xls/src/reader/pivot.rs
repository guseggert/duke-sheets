//! BIFF8 pivot reader using the SX* and DCON* record layouts in [MS-XLS] section 2.4.

use std::collections::{BTreeMap, HashSet};
use std::io::Cursor;

use duke_sheets_core::style::NumberFormat;
use duke_sheets_core::{
    CellAddress, CellRange, PivotAggregate, PivotCacheInfo, PivotCacheSourceKind,
    PivotCalculatedField, PivotCalculatedItem, PivotDateGroupUnit, PivotDatePeriod, PivotField,
    PivotFieldRef, PivotFilter, PivotFilterOperator, PivotGrouping, PivotManualGroup, PivotMeasure,
    PivotShowAs, PivotSort, PivotSource, PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable,
    PivotValue, PivotValuesAxis, Worksheet,
};
use duke_sheets_formula::decompile::{
    decompile_pivot_formula as decompile_biff_pivot_formula, PivotFormulaHooks,
    PivotVariableArgCount,
};
use ssfmt::{date_serial::date_to_serial, DateSystem};

use super::XlsReader;
use crate::biff::parser::read_u8;
use crate::biff::records;
use crate::biff::strings::{read_character_data, read_unicode_string};
use crate::biff::{self, BiffRecord};
use crate::error::{XlsError, XlsResult};

#[derive(Default)]
pub(super) struct PivotCaches(std::collections::HashMap<u16, XlsPivotCache>);

pub(super) fn parse_cache_source(record_type: u16, data: &[u8]) -> Option<PivotSource> {
    if record_type == records::DCONREF {
        XlsReader::parse_dconref(data)
    } else {
        XlsReader::parse_dconname(data)
    }
}

pub(super) fn parse_consolidation_page_refs(data: &[u8]) -> Vec<u16> {
    XlsReader::parse_sxtbpg(data)
}

pub(super) fn apply_consolidation_page_items(
    sources: &mut std::collections::HashMap<u16, Vec<PivotSource>>,
    page_refs: &std::collections::HashMap<u16, Vec<Vec<u16>>>,
    page_names: &std::collections::HashMap<u16, Vec<String>>,
) {
    XlsReader::apply_pivot_consolidation_page_items(sources, page_refs, page_names);
}

pub(super) fn read_caches(
    cfb: &crate::cfb::CompoundFile,
    cache_ids: &std::collections::BTreeSet<u16>,
    sources: &std::collections::HashMap<u16, Vec<PivotSource>>,
    date_system: DateSystem,
) -> XlsResult<PivotCaches> {
    XlsReader::read_pivot_caches(cfb, cache_ids, sources, date_system).map(PivotCaches)
}

pub(super) fn parse_sheet_tables(
    records: &[&BiffRecord],
    ws: &mut Worksheet,
    caches: &PivotCaches,
    num_fmts: &std::collections::HashMap<u16, String>,
) -> XlsResult<()> {
    XlsReader::parse_sheet_pivot_tables(records, ws, &caches.0, num_fmts)
}

fn xls_date_group_unit_from_flags(flags: u16) -> Option<PivotDateGroupUnit> {
    Some(match ((flags & 0x001C) >> 2) as u8 {
        0x01 => PivotDateGroupUnit::Seconds,
        0x02 => PivotDateGroupUnit::Minutes,
        0x03 => PivotDateGroupUnit::Hours,
        0x04 => PivotDateGroupUnit::Days,
        0x05 => PivotDateGroupUnit::Months,
        0x06 => PivotDateGroupUnit::Quarters,
        0x07 => PivotDateGroupUnit::Years,
        _ => return None,
    })
}

fn pivot_grouping_field_name(grouping: &PivotGrouping) -> &str {
    match grouping {
        PivotGrouping::Number { field, .. }
        | PivotGrouping::Date { field, .. }
        | PivotGrouping::Manual { field, .. } => &field.name,
    }
}

#[derive(Debug, Clone)]
struct XlsPivotCache {
    cache_id: u16,
    source: PivotSource,
    source_kind: PivotCacheSourceKind,
    fields: Vec<XlsPivotCacheField>,
    calculated_item_formulas: Vec<XlsPivotCalculatedItemFormula>,
    record_count: Option<u64>,
}

#[derive(Debug, Clone)]
struct XlsPivotCacheField {
    name: String,
    formula: Option<String>,
    grouping: Option<PivotGrouping>,
    shared_items: Vec<PivotValue>,
    manual_group_item_ids: Vec<u32>,
    group_parent_field: Option<usize>,
    group_base_field: Option<usize>,
}

#[derive(Debug, Clone)]
struct PendingPivotRangeGroup {
    flags: u16,
    numbers: Vec<f64>,
}

#[derive(Debug, Clone)]
struct XlsPivotFormulaPending {
    ptgs: Vec<u8>,
    expected_sx_names: usize,
    sx_name_field_indexes: Vec<usize>,
}

#[derive(Debug, Clone)]
struct XlsPivotCalculatedItemFormula {
    field_index: usize,
    formula: String,
}

#[derive(Debug, Clone)]
struct XlsPivotCalculatedItemPending {
    ptgs: Vec<u8>,
    expected_sx_names: usize,
    item_refs: Vec<(usize, usize)>,
}

struct XlsPivotFormulaHooks<'a> {
    indexes: &'a [usize],
    names: &'a [String],
    formatter: fn(&str) -> String,
}

impl PivotFormulaHooks for XlsPivotFormulaHooks<'_> {
    fn resolve_name(&mut self, index: u32) -> Option<String> {
        let field_index = *self.indexes.get(index as usize)?;
        self.names
            .get(field_index)
            .map(|name| (self.formatter)(name))
    }

    fn variable_arg_count(&self) -> PivotVariableArgCount {
        PivotVariableArgCount::Biff8
    }
}

#[derive(Debug, Clone)]
struct XlsPivotViewBuilder {
    name: String,
    cache_id: u16,
    target: CellAddress,
    rendered_range: Option<CellRange>,
    field_options: Vec<XlsPivotFieldOptions>,
    pending_sxaddl_field_name: Option<String>,
    row_axis_count: usize,
    column_axis_count: usize,
    page_axis_count: usize,
    row_axis: Vec<i16>,
    column_axis: Vec<i16>,
    page_fields: Vec<(usize, u16)>,
    measures: Vec<PivotMeasure>,
    layout: duke_sheets_core::PivotLayout,
    preserve_formatting: bool,
    style: PivotStyle,
    pending_sxaddl_label_filter_field_index: Option<usize>,
    pending_sxaddl_label_filter_operator: Option<PivotFilterOperator>,
    pending_sxaddl_label_filter_type: Option<u32>,
    pending_sxaddl_label_filter_values: Vec<String>,
    pending_sxaddl_value_filter: Option<XlsPendingValueFilter>,
    pending_sxaddl_date_filter: Option<XlsPendingDateFilter>,
}

fn pivot_page_area_rows(layout: &duke_sheets_core::PivotLayout, count: usize) -> u32 {
    if count == 0 {
        return 0;
    }
    let wrap = layout.page_wrap as usize;
    let rows = if wrap == 0 {
        count
    } else if layout.page_over_then_down {
        count.div_ceil(wrap)
    } else {
        wrap.min(count)
    };
    rows as u32
}

fn pivot_target_from_body(
    body: CellAddress,
    layout: &duke_sheets_core::PivotLayout,
    page_axis_count: usize,
) -> CellAddress {
    let page_rows = pivot_page_area_rows(layout, page_axis_count);
    let row = if page_rows > 0 {
        body.row.saturating_sub(page_rows.saturating_add(1))
    } else {
        body.row
    };
    CellAddress::new(row, body.col)
}

#[derive(Debug, Clone)]
struct XlsPivotFieldOptions {
    caption: Option<String>,
    sort: PivotSort,
    sort_by_measure_index: Option<usize>,
    subtotal: PivotSubtotal,
    subtotal_caption: Option<String>,
    subtotals: Vec<PivotSubtotal>,
    show_empty_items: bool,
    show_drop_downs: bool,
    subtotal_top: bool,
    insert_blank_row: bool,
    insert_page_break: bool,
    include_new_items_in_filter: bool,
    item_page_count: u32,
    top_n_filter: Option<XlsTopNFilter>,
    label_filter: Option<XlsLabelFilter>,
    value_filter: Option<XlsValueFilter>,
    date_filter: Option<XlsDateFilter>,
    hidden_items: Vec<usize>,
    calculated_items: Vec<usize>,
}

#[derive(Debug, Clone, Copy)]
struct XlsTopNFilter {
    n: u32,
    top: bool,
    percent: bool,
    measure_index: usize,
}

#[derive(Debug, Clone)]
struct XlsLabelFilter {
    kind: XlsLabelFilterKind,
}

#[derive(Debug, Clone)]
enum XlsLabelFilterKind {
    Comparison {
        operator: PivotFilterOperator,
        value: String,
    },
    Between {
        start: String,
        end: String,
        not_between: bool,
    },
}

#[derive(Debug, Clone, Copy)]
struct XlsValueFilter {
    measure_index: usize,
    kind: XlsValueFilterKind,
}

#[derive(Debug, Clone, Copy)]
enum XlsValueFilterKind {
    Comparison {
        operator: PivotFilterOperator,
        value: f64,
    },
    Between {
        start: f64,
        end: f64,
        not_between: bool,
    },
}

#[derive(Debug, Clone, Copy)]
struct XlsDateFilter {
    kind: XlsDateFilterKind,
}

#[derive(Debug, Clone, Copy)]
enum XlsDateFilterKind {
    Comparison {
        operator: PivotFilterOperator,
        value: f64,
    },
    Between {
        start: f64,
        end: f64,
        not_between: bool,
    },
    Period(PivotDatePeriod),
}

#[derive(Debug, Clone, Copy)]
struct XlsPendingValueFilter {
    field_index: usize,
    measure_index: usize,
    filter_type: u32,
}

#[derive(Debug, Clone, Copy)]
struct XlsPendingDateFilter {
    field_index: usize,
    filter_type: u32,
}

impl Default for XlsPivotFieldOptions {
    fn default() -> Self {
        Self {
            caption: None,
            sort: PivotSort::None,
            sort_by_measure_index: None,
            subtotal: PivotSubtotal::Automatic,
            subtotal_caption: None,
            subtotals: Vec::new(),
            show_empty_items: false,
            show_drop_downs: true,
            subtotal_top: true,
            insert_blank_row: false,
            insert_page_break: false,
            include_new_items_in_filter: false,
            item_page_count: 10,
            top_n_filter: None,
            label_filter: None,
            value_filter: None,
            date_filter: None,
            hidden_items: Vec::new(),
            calculated_items: Vec::new(),
        }
    }
}

impl XlsReader {
    fn read_pivot_caches(
        cfb: &crate::cfb::CompoundFile,
        cache_ids: &std::collections::BTreeSet<u16>,
        sources: &std::collections::HashMap<u16, Vec<PivotSource>>,
        date_system: DateSystem,
    ) -> XlsResult<std::collections::HashMap<u16, XlsPivotCache>> {
        let mut caches = std::collections::HashMap::new();
        for cache_id in cache_ids {
            let stream_path = format!("/_SX_DB_CUR/{:04}", cache_id.saturating_add(1));
            let Ok(stream) = cfb.read_stream(&stream_path) else {
                continue;
            };
            let records = biff::read_all_records(&mut Cursor::new(stream))?;
            let source = sources.get(cache_id).cloned().unwrap_or_default();
            if let Some(cache) =
                Self::parse_pivot_cache_stream(*cache_id, source, &records, date_system)?
            {
                caches.insert(*cache_id, cache);
            }
        }
        Ok(caches)
    }

    fn parse_pivot_cache_stream(
        cache_id: u16,
        sources: Vec<PivotSource>,
        records: &[BiffRecord],
        date_system: DateSystem,
    ) -> XlsResult<Option<XlsPivotCache>> {
        let mut fields = Vec::new();
        let mut current_field: Option<XlsPivotCacheField> = None;
        let mut pending_formula: Option<XlsPivotFormulaPending> = None;
        let mut pending_calculated_item: Option<XlsPivotCalculatedItemPending> = None;
        let mut pending_calculated_item_formulas = Vec::new();
        let mut pending_group_range: Option<PendingPivotRangeGroup> = None;
        let mut record_count = None;
        let mut source_kind = PivotCacheSourceKind::Worksheet;

        for rec in records {
            match rec.record_type {
                records::SXDB => {
                    if rec.data.len() >= 4 {
                        record_count = Some(u32::from_le_bytes([
                            rec.data[0],
                            rec.data[1],
                            rec.data[2],
                            rec.data[3],
                        ]) as u64);
                    }
                    if rec.data.len() >= 18 {
                        source_kind =
                            Self::pivot_cache_source_kind_from_sxdb(u16::from_le_bytes([
                                rec.data[16],
                                rec.data[17],
                            ]));
                    }
                }
                records::SXFDB => {
                    Self::attach_pending_pivot_grouping(
                        &mut current_field,
                        &mut pending_group_range,
                        &fields,
                    );
                    Self::attach_pending_pivot_formula(
                        &mut current_field,
                        &mut pending_formula,
                        &fields,
                    );
                    if let Some(field) = current_field.take() {
                        fields.push(field);
                    }
                    current_field = Some(XlsPivotCacheField {
                        name: Self::parse_sxfdb_name(&rec.data)?,
                        formula: None,
                        grouping: None,
                        shared_items: Vec::new(),
                        manual_group_item_ids: Vec::new(),
                        group_parent_field: Self::parse_sxfdb_field_index(&rec.data, 2),
                        group_base_field: Self::parse_sxfdb_field_index(&rec.data, 4),
                    });
                }
                0x00F9 => {
                    pending_formula = Self::parse_sxfmla(&rec.data);
                    Self::attach_pending_pivot_formula(
                        &mut current_field,
                        &mut pending_formula,
                        &fields,
                    );
                }
                0x00F6 => {
                    if let Some(field_index) = Self::parse_sxname(&rec.data) {
                        if let Some(pending) = &mut pending_formula {
                            pending.sx_name_field_indexes.push(field_index);
                        }
                    } else if Self::parse_sxname_item_pair_count(&rec.data).is_some()
                        && pending_calculated_item.is_none()
                    {
                        if let Some(pending) = pending_formula.take() {
                            pending_calculated_item = Some(XlsPivotCalculatedItemPending {
                                ptgs: pending.ptgs,
                                expected_sx_names: pending.expected_sx_names,
                                item_refs: Vec::with_capacity(pending.expected_sx_names),
                            });
                        }
                    }
                    Self::attach_pending_pivot_formula(
                        &mut current_field,
                        &mut pending_formula,
                        &fields,
                    );
                }
                0x00F8 => {
                    if let Some(pending) = &mut pending_calculated_item {
                        if let Some(item_ref) = Self::parse_sxpair(&rec.data) {
                            pending.item_refs.push(item_ref);
                        }
                    }
                }
                0x0103 => {
                    if let Some(pending) = pending_calculated_item.take() {
                        if pending.item_refs.len() >= pending.expected_sx_names {
                            pending_calculated_item_formulas.push(pending);
                        }
                    }
                }
                records::SXSTRING => {
                    if let Some(field) = &mut current_field {
                        field.shared_items.push(Self::parse_sxstring(&rec.data)?);
                    }
                }
                records::SXBOOL => {
                    if let Some(field) = &mut current_field {
                        field.shared_items.push(PivotValue::Boolean(
                            rec.data.first().copied().unwrap_or(0) != 0,
                        ));
                    }
                }
                records::SXERR => {
                    if let Some(field) = &mut current_field {
                        field.shared_items.push(PivotValue::Error(
                            Self::cell_error_from_biff_code(
                                rec.data.first().copied().unwrap_or(0x0F),
                            ),
                        ));
                    }
                }
                records::SXINT => {
                    if rec.data.len() >= 2 {
                        let value = i16::from_le_bytes([rec.data[0], rec.data[1]]) as f64;
                        if let Some(pending) = &mut pending_group_range {
                            pending.numbers.push(value);
                            Self::attach_pending_pivot_grouping(
                                &mut current_field,
                                &mut pending_group_range,
                                &fields,
                            );
                        } else if let Some(field) = &mut current_field {
                            field.shared_items.push(PivotValue::Number(value));
                        }
                    }
                }
                records::SXDTR => {
                    if let Some(value) = Self::parse_sxdtr(&rec.data, date_system) {
                        if let Some(pending) = &mut pending_group_range {
                            pending.numbers.push(value);
                            Self::attach_pending_pivot_grouping(
                                &mut current_field,
                                &mut pending_group_range,
                                &fields,
                            );
                        } else if let Some(field) = &mut current_field {
                            field.shared_items.push(PivotValue::Number(value));
                        }
                    }
                }
                records::SXNIL => {
                    if let Some(field) = &mut current_field {
                        field.shared_items.push(PivotValue::Blank);
                    }
                }
                records::SXIDSTM => {
                    if let Some(field) = &mut current_field {
                        field.manual_group_item_ids = rec
                            .data
                            .chunks_exact(2)
                            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]) as u32)
                            .collect();
                    }
                }
                records::SXRNG => {
                    pending_group_range = Self::parse_sxrng(&rec.data);
                }
                records::SXDBB => {
                    Self::attach_pending_pivot_grouping(
                        &mut current_field,
                        &mut pending_group_range,
                        &fields,
                    );
                    Self::attach_pending_pivot_formula(
                        &mut current_field,
                        &mut pending_formula,
                        &fields,
                    );
                    if let Some(field) = current_field.take() {
                        fields.push(field);
                    }
                }
                records::SXNUM => {
                    if let Some(pending) = &mut pending_group_range {
                        if rec.data.len() >= 8 {
                            pending
                                .numbers
                                .push(f64::from_le_bytes(rec.data[0..8].try_into().unwrap()));
                        }
                        Self::attach_pending_pivot_grouping(
                            &mut current_field,
                            &mut pending_group_range,
                            &fields,
                        );
                    } else if let Some(field) = &mut current_field {
                        if rec.data.len() >= 8 {
                            field
                                .shared_items
                                .push(PivotValue::Number(f64::from_le_bytes(
                                    rec.data[0..8].try_into().unwrap(),
                                )));
                        }
                    }
                }
                records::EOF => break,
                _ => {}
            }
        }

        Self::attach_pending_pivot_grouping(&mut current_field, &mut pending_group_range, &fields);
        Self::attach_pending_pivot_formula(&mut current_field, &mut pending_formula, &fields);
        if let Some(field) = current_field.take() {
            fields.push(field);
        }
        if fields.is_empty() {
            return Ok(None);
        }

        let calculated_item_formulas = pending_calculated_item_formulas
            .iter()
            .filter_map(|pending| Self::calculated_item_formula_from_pending(pending, &fields))
            .collect();
        let Some(source) = Self::pivot_source_for_sxdb_kind(source_kind, sources) else {
            return Ok(None);
        };

        Ok(Some(XlsPivotCache {
            cache_id,
            source,
            source_kind,
            fields,
            calculated_item_formulas,
            record_count,
        }))
    }

    fn parse_sheet_pivot_tables(
        records: &[&BiffRecord],
        ws: &mut duke_sheets_core::Worksheet,
        caches: &std::collections::HashMap<u16, XlsPivotCache>,
        num_fmts: &std::collections::HashMap<u16, String>,
    ) -> XlsResult<()> {
        let mut current: Option<XlsPivotViewBuilder> = None;

        for rec in records {
            match rec.record_type {
                records::SXVIEW => {
                    if let Some(builder) = current.take() {
                        if let Some(pivot) = Self::finish_pivot_view(builder, caches) {
                            ws.add_pivot_table(pivot).map_err(|e| {
                                XlsError::InvalidFormat(format!("invalid XLS pivot table: {e}"))
                            })?;
                        }
                    }
                    current = Self::parse_sxview(&rec.data)?;
                }
                records::SXVD => {
                    if let Some(builder) = &mut current {
                        builder.field_options.push(Self::parse_sxvd(&rec.data));
                    }
                }
                records::SXVI => {
                    if let Some(builder) = &mut current {
                        if let Some(options) = builder.field_options.last_mut() {
                            Self::apply_sxvi(options, &rec.data);
                        }
                    }
                }
                records::SXIVD => {
                    if let Some(builder) = &mut current {
                        let mut declaration = Self::parse_sxivd(&rec.data).into_iter();
                        while builder.row_axis.len() < builder.row_axis_count {
                            let Some(index) = declaration.next() else {
                                break;
                            };
                            builder.row_axis.push(index);
                        }
                        while builder.column_axis.len() < builder.column_axis_count {
                            let Some(index) = declaration.next() else {
                                break;
                            };
                            builder.column_axis.push(index);
                        }
                    }
                }
                0x0100 => {
                    if let Some(builder) = &mut current {
                        if let Some(options) = builder.field_options.last_mut() {
                            Self::apply_sxvdex(options, &rec.data);
                        }
                    }
                }
                0x00B6 => {
                    if let Some(builder) = &mut current {
                        if let Some(page_field) = Self::parse_sxpi(&rec.data) {
                            builder.page_fields.push(page_field);
                        }
                    }
                }
                0x00C5 => {
                    if let Some(builder) = &mut current {
                        if let Some(measure) =
                            Self::parse_sxdi(&rec.data, caches, builder.cache_id, num_fmts)?
                        {
                            builder.measures.push(measure);
                        }
                    }
                }
                0x00F1 => {
                    if let Some(builder) = &mut current {
                        Self::apply_sxex(builder, &rec.data);
                    }
                }
                0x0864 => {
                    if let Some(builder) = &mut current {
                        Self::apply_pivot_frt0864_field_options(builder, caches, &rec.data);
                        if let Some(style) = Self::parse_pivot_frt0864_style(&rec.data) {
                            builder.style = style;
                        }
                    }
                }
                0x0810 => {
                    if let Some(builder) = &mut current {
                        Self::apply_sxviewex9(builder, &rec.data);
                    }
                }
                _ => {}
            }
        }

        if let Some(builder) = current.take() {
            if let Some(pivot) = Self::finish_pivot_view(builder, caches) {
                ws.add_pivot_table(pivot).map_err(|e| {
                    XlsError::InvalidFormat(format!("invalid XLS pivot table: {e}"))
                })?;
            }
        }
        Ok(())
    }

    fn finish_pivot_view(
        builder: XlsPivotViewBuilder,
        caches: &std::collections::HashMap<u16, XlsPivotCache>,
    ) -> Option<PivotTable> {
        let cache = caches.get(&builder.cache_id)?;
        let mut layout = builder.layout;
        let mut rows = Vec::new();
        let mut columns = Vec::new();
        let row_field_indexes = Self::axis_field_indexes(&builder.row_axis);
        let column_field_indexes = Self::axis_field_indexes(&builder.column_axis);

        if !builder.row_axis.is_empty() {
            Self::push_pivot_axis_fields(
                &builder.row_axis,
                cache,
                &builder.field_options,
                &mut rows,
                PivotValuesAxis::Rows,
                &mut layout,
            );
        }
        if !builder.column_axis.is_empty() {
            Self::push_pivot_axis_fields(
                &builder.column_axis,
                cache,
                &builder.field_options,
                &mut columns,
                PivotValuesAxis::Columns,
                &mut layout,
            );
        }

        let mut page_fields = Vec::new();
        let mut filters = Vec::new();
        for (field_index, selected_item) in builder.page_fields {
            let Some(field) = cache.fields.get(field_index) else {
                continue;
            };
            let field_name = Self::pivot_cache_field_semantic_name_by_index(cache, field_index);
            if let Some(field) =
                Self::pivot_axis_field(cache, field_index, builder.field_options.get(field_index))
            {
                Self::push_axis_field(&mut page_fields, field);
            }
            if selected_item != 0xFFFF {
                if let Some(item) = field.shared_items.get(selected_item as usize) {
                    filters.push(PivotFilter::FieldItems {
                        field: duke_sheets_core::PivotFieldRef::new(field_name),
                        allowed_items: vec![item.clone()],
                    });
                } else if selected_item == 0x7FFD {
                    if let Some(options) = builder.field_options.get(field_index) {
                        if !options.hidden_items.is_empty() {
                            let allowed_items = field
                                .shared_items
                                .iter()
                                .enumerate()
                                .filter(|(index, _)| !options.hidden_items.contains(index))
                                .map(|(_, item)| item.clone())
                                .collect::<Vec<_>>();
                            if !allowed_items.is_empty()
                                && allowed_items.len() < field.shared_items.len()
                            {
                                filters.push(PivotFilter::FieldItems {
                                    field: duke_sheets_core::PivotFieldRef::new(field_name),
                                    allowed_items,
                                });
                            }
                        }
                    }
                }
            }
        }
        Self::push_hidden_axis_field_item_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &row_field_indexes,
        );
        Self::push_hidden_axis_field_item_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &column_field_indexes,
        );
        let all_field_indexes = (0..builder.field_options.len()).collect::<Vec<_>>();
        Self::push_hidden_axis_field_item_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &all_field_indexes,
        );
        Self::push_label_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &row_field_indexes,
        );
        Self::push_label_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &column_field_indexes,
        );
        Self::push_value_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &row_field_indexes,
            &builder.measures,
        );
        Self::push_value_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &column_field_indexes,
            &builder.measures,
        );
        Self::push_date_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &row_field_indexes,
        );
        Self::push_date_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &column_field_indexes,
        );
        Self::push_top_n_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &row_field_indexes,
            &builder.measures,
        );
        Self::push_top_n_filters(
            &mut filters,
            cache,
            &builder.field_options,
            &column_field_indexes,
            &builder.measures,
        );
        Self::apply_include_new_items_options(
            &mut rows,
            &mut columns,
            &mut page_fields,
            cache,
            &builder.field_options,
            &filters,
        );
        Self::apply_sort_by_measure_options(
            &mut rows,
            &mut columns,
            &mut page_fields,
            cache,
            &builder.field_options,
            &builder.measures,
        );

        let source = cache.source.clone();
        let source_kind = cache.source_kind;

        let target = pivot_target_from_body(builder.target, &layout, builder.page_axis_count);
        let mut pivot = PivotTable::new(0, builder.name, source, target);
        pivot.rows = rows;
        pivot.columns = columns;
        pivot.page_fields = page_fields;
        pivot.measures = builder.measures;
        pivot.calculated_fields = cache
            .fields
            .iter()
            .filter_map(|field| {
                Some(PivotCalculatedField::new(
                    field.name.clone(),
                    field.formula.clone()?,
                ))
            })
            .collect();
        pivot.calculated_items =
            Self::pivot_calculated_items_from_cache(cache, &builder.field_options);
        pivot.groupings = Self::semantic_groupings_from_pivot_cache_fields(&cache.fields);
        pivot.filters = filters;
        pivot.layout = layout;
        pivot.refresh_policy.preserve_formatting = builder.preserve_formatting;
        pivot.style = builder.style;
        pivot.rendered_range = builder.rendered_range;
        pivot.set_cache_info(Some(PivotCacheInfo {
            cache_id: u32::from(cache.cache_id),
            source_kind,
            record_count: cache.record_count,
            refreshed_version: None,
        }));
        Some(pivot)
    }

    fn pivot_cache_source_kind_from_sxdb(vs_type: u16) -> PivotCacheSourceKind {
        match vs_type {
            0x01 => PivotCacheSourceKind::Worksheet,
            0x02 => PivotCacheSourceKind::External,
            0x04 => PivotCacheSourceKind::Consolidation,
            0x08 => PivotCacheSourceKind::Scenario,
            _ => PivotCacheSourceKind::Unknown,
        }
    }

    fn pivot_source_for_sxdb_kind(
        source_kind: PivotCacheSourceKind,
        sources: Vec<PivotSource>,
    ) -> Option<PivotSource> {
        match source_kind {
            PivotCacheSourceKind::Worksheet => sources.into_iter().next().or_else(|| {
                log::warn!("skipping XLS worksheet pivot cache without DCONREF/DCONNAME source");
                None
            }),
            PivotCacheSourceKind::External => Some(
                sources
                    .into_iter()
                    .find(|source| matches!(source, PivotSource::External { .. }))
                    .unwrap_or_else(|| PivotSource::External {
                        connection_name: String::new(),
                        command_text: None,
                    }),
            ),
            PivotCacheSourceKind::Consolidation => {
                let ranges = sources
                    .into_iter()
                    .flat_map(|source| match source {
                        PivotSource::Consolidation { ranges } => ranges,
                        PivotSource::WorksheetRange { sheet, range } => {
                            vec![PivotSourceRange::new(sheet.unwrap_or_default(), range)]
                        }
                        PivotSource::Table { name } => vec![PivotSourceRange::named(name)],
                        _ => Vec::new(),
                    })
                    .collect();
                Some(PivotSource::Consolidation { ranges })
            }
            PivotCacheSourceKind::Scenario => Some(PivotSource::Scenario {
                name: String::new(),
            }),
            PivotCacheSourceKind::Olap => Some(PivotSource::Olap {
                connection_name: String::new(),
                cube: None,
                command_text: None,
            }),
            PivotCacheSourceKind::Unknown => Some(sources.into_iter().next().unwrap_or_else(
                || PivotSource::External {
                    connection_name: String::new(),
                    command_text: None,
                },
            )),
        }
    }

    fn apply_pivot_consolidation_page_items(
        sources_by_cache: &mut std::collections::HashMap<u16, Vec<PivotSource>>,
        page_refs_by_cache: &std::collections::HashMap<u16, Vec<Vec<u16>>>,
        page_names_by_cache: &std::collections::HashMap<u16, Vec<String>>,
    ) {
        for (cache_id, page_refs) in page_refs_by_cache {
            let Some(sources) = sources_by_cache.get_mut(cache_id) else {
                continue;
            };
            let page_names = page_names_by_cache
                .get(cache_id)
                .map(Vec::as_slice)
                .unwrap_or(&[]);
            for (source, indexes) in sources.iter_mut().zip(page_refs.iter()) {
                let page_items = indexes
                    .iter()
                    .filter_map(|index| page_names.get(*index as usize).cloned())
                    .collect::<Vec<_>>();
                if !page_items.is_empty() {
                    Self::set_pivot_source_page_items(source, page_items);
                }
            }
        }
    }

    fn set_pivot_source_page_items(source: &mut PivotSource, page_items: Vec<String>) {
        match source {
            PivotSource::Consolidation { ranges } => {
                for range in ranges {
                    range.page_items = page_items.clone();
                }
            }
            PivotSource::WorksheetRange { sheet, range } => {
                let sheet = sheet.clone().unwrap_or_default();
                let range = *range;
                *source = PivotSource::Consolidation {
                    ranges: vec![PivotSourceRange::new(sheet, range).with_page_items(page_items)],
                };
            }
            PivotSource::Table { name } => {
                let name = name.clone();
                *source = PivotSource::Consolidation {
                    ranges: vec![PivotSourceRange::named(name).with_page_items(page_items)],
                };
            }
            _ => {}
        }
    }

    fn pivot_calculated_items_from_cache(
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
    ) -> Vec<PivotCalculatedItem> {
        let mut used_formulas_by_field = vec![0usize; cache.fields.len()];
        let mut calculated_items = Vec::new();

        for (field_index, options) in field_options.iter().enumerate() {
            let Some(field) = cache.fields.get(field_index) else {
                continue;
            };
            for item_index in &options.calculated_items {
                let Some(item) = field.shared_items.get(*item_index).cloned() else {
                    continue;
                };
                let formula_index = used_formulas_by_field
                    .get(field_index)
                    .copied()
                    .unwrap_or(0);
                let Some(formula) = cache
                    .calculated_item_formulas
                    .iter()
                    .filter(|formula| formula.field_index == field_index)
                    .nth(formula_index)
                else {
                    continue;
                };
                if let Some(slot) = used_formulas_by_field.get_mut(field_index) {
                    *slot += 1;
                }
                calculated_items.push(PivotCalculatedItem::new(
                    field.name.clone(),
                    item,
                    formula.formula.clone(),
                ));
            }
        }

        calculated_items
    }

    fn push_pivot_axis_fields(
        indexes: &[i16],
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
        fields: &mut Vec<PivotField>,
        axis: PivotValuesAxis,
        layout: &mut duke_sheets_core::PivotLayout,
    ) {
        for index in indexes {
            if *index == -2 {
                layout.values_axis = axis;
                layout
                    .values_axis_position
                    .get_or_insert(fields.len() as u32);
            } else if *index >= 0 {
                let field_index = *index as usize;
                if let Some(field) =
                    Self::pivot_axis_field(cache, field_index, field_options.get(field_index))
                {
                    Self::push_axis_field(fields, field);
                }
            }
        }
    }

    fn axis_field_indexes(indexes: &[i16]) -> Vec<usize> {
        indexes
            .iter()
            .filter_map(|index| (*index >= 0).then_some(*index as usize))
            .collect()
    }

    fn push_hidden_axis_field_item_filters(
        filters: &mut Vec<PivotFilter>,
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
        field_indexes: &[usize],
    ) {
        for &field_index in field_indexes {
            let Some(field) = cache.fields.get(field_index) else {
                continue;
            };
            let Some(options) = field_options.get(field_index) else {
                continue;
            };
            if options.hidden_items.is_empty() || field.shared_items.is_empty() {
                continue;
            }

            let field_name = Self::pivot_cache_field_semantic_name_by_index(cache, field_index);
            if filters.iter().any(|filter| {
                matches!(
                    filter,
                    PivotFilter::FieldItems { field, .. }
                        if field.name.eq_ignore_ascii_case(&field_name)
                )
            }) {
                continue;
            }

            let hidden_items = options.hidden_items.iter().copied().collect::<HashSet<_>>();
            let allowed_items = field
                .shared_items
                .iter()
                .enumerate()
                .filter(|(index, _)| !hidden_items.contains(index))
                .map(|(_, item)| item.clone())
                .collect::<Vec<_>>();
            if !allowed_items.is_empty() && allowed_items.len() < field.shared_items.len() {
                filters.push(PivotFilter::FieldItems {
                    field: PivotFieldRef::new(field_name),
                    allowed_items,
                });
            }
        }
    }

    fn push_top_n_filters(
        filters: &mut Vec<PivotFilter>,
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
        field_indexes: &[usize],
        measures: &[PivotMeasure],
    ) {
        for &field_index in field_indexes {
            let Some(options) = field_options.get(field_index) else {
                continue;
            };
            let Some(top_n) = options.top_n_filter else {
                continue;
            };
            let Some(measure) = measures.get(top_n.measure_index).cloned() else {
                continue;
            };
            filters.push(PivotFilter::TopN {
                field: PivotFieldRef::new(Self::pivot_cache_field_semantic_name_by_index(
                    cache,
                    field_index,
                )),
                measure,
                n: top_n.n,
                top: top_n.top,
                percent: top_n.percent,
            });
        }
    }

    fn push_value_filters(
        filters: &mut Vec<PivotFilter>,
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
        field_indexes: &[usize],
        measures: &[PivotMeasure],
    ) {
        for &field_index in field_indexes {
            let Some(options) = field_options.get(field_index) else {
                continue;
            };
            let Some(value_filter) = options.value_filter else {
                continue;
            };
            let Some(measure) = measures.get(value_filter.measure_index).cloned() else {
                continue;
            };
            let field = PivotFieldRef::new(Self::pivot_cache_field_semantic_name_by_index(
                cache,
                field_index,
            ));
            match value_filter.kind {
                XlsValueFilterKind::Comparison { operator, value } => {
                    filters.push(PivotFilter::Value {
                        field,
                        measure,
                        operator,
                        value,
                    });
                }
                XlsValueFilterKind::Between {
                    start,
                    end,
                    not_between,
                } => {
                    filters.push(PivotFilter::ValueBetween {
                        field,
                        measure,
                        start,
                        end,
                        not_between,
                    });
                }
            }
        }
    }

    fn push_date_filters(
        filters: &mut Vec<PivotFilter>,
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
        field_indexes: &[usize],
    ) {
        for &field_index in field_indexes {
            let Some(options) = field_options.get(field_index) else {
                continue;
            };
            let Some(date_filter) = options.date_filter else {
                continue;
            };
            let field = PivotFieldRef::new(Self::pivot_cache_field_semantic_name_by_index(
                cache,
                field_index,
            ));
            match date_filter.kind {
                XlsDateFilterKind::Comparison { operator, value } => {
                    filters.push(PivotFilter::Date {
                        field,
                        operator,
                        value,
                    });
                }
                XlsDateFilterKind::Between {
                    start,
                    end,
                    not_between,
                } => {
                    filters.push(PivotFilter::DateBetween {
                        field,
                        start,
                        end,
                        not_between,
                    });
                }
                XlsDateFilterKind::Period(period) => {
                    filters.push(PivotFilter::DatePeriod { field, period });
                }
            }
        }
    }

    fn push_label_filters(
        filters: &mut Vec<PivotFilter>,
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
        field_indexes: &[usize],
    ) {
        for &field_index in field_indexes {
            let Some(options) = field_options.get(field_index) else {
                continue;
            };
            let Some(label_filter) = options.label_filter.as_ref() else {
                continue;
            };
            let field = PivotFieldRef::new(Self::pivot_cache_field_semantic_name_by_index(
                cache,
                field_index,
            ));
            match &label_filter.kind {
                XlsLabelFilterKind::Comparison { operator, value } => {
                    filters.push(PivotFilter::Label {
                        field,
                        operator: *operator,
                        value: value.clone(),
                    });
                }
                XlsLabelFilterKind::Between {
                    start,
                    end,
                    not_between,
                } => {
                    filters.push(PivotFilter::LabelBetween {
                        field,
                        start: start.clone(),
                        end: end.clone(),
                        not_between: *not_between,
                    });
                }
            }
        }
    }

    fn pivot_axis_field(
        cache: &XlsPivotCache,
        field_index: usize,
        options: Option<&XlsPivotFieldOptions>,
    ) -> Option<PivotField> {
        cache.fields.get(field_index)?;
        let options = options.cloned().unwrap_or_default();
        let mut field = PivotField::new(Self::pivot_cache_field_semantic_name_by_index(
            cache,
            field_index,
        ));
        field.caption = options.caption;
        field.sort = options.sort;
        field.subtotal = options.subtotal;
        field.subtotal_caption = options.subtotal_caption;
        field.subtotals = options.subtotals;
        field.show_empty_items = options.show_empty_items;
        field.show_drop_downs = options.show_drop_downs;
        field.subtotal_top = options.subtotal_top;
        field.insert_blank_row = options.insert_blank_row;
        field.insert_page_break = options.insert_page_break;
        field.item_page_count = options.item_page_count;
        Some(field)
    }

    fn apply_include_new_items_options(
        rows: &mut [PivotField],
        columns: &mut [PivotField],
        page_fields: &mut [PivotField],
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
        filters: &[PivotFilter],
    ) {
        for filter in filters {
            let PivotFilter::FieldItems { field, .. } = filter else {
                continue;
            };
            let Some((field_index, _)) = cache.fields.iter().enumerate().find(|(index, _)| {
                Self::pivot_cache_field_semantic_name_by_index(cache, *index)
                    .eq_ignore_ascii_case(&field.name)
            }) else {
                continue;
            };
            let Some(options) = field_options.get(field_index) else {
                continue;
            };
            if !options.include_new_items_in_filter {
                continue;
            }
            for axis_field in rows
                .iter_mut()
                .chain(columns.iter_mut())
                .chain(page_fields.iter_mut())
            {
                if axis_field.field.name.eq_ignore_ascii_case(&field.name) {
                    axis_field.include_new_items_in_filter = true;
                }
            }
        }
    }

    fn apply_sort_by_measure_options(
        rows: &mut [PivotField],
        columns: &mut [PivotField],
        page_fields: &mut [PivotField],
        cache: &XlsPivotCache,
        field_options: &[XlsPivotFieldOptions],
        measures: &[PivotMeasure],
    ) {
        for field in rows
            .iter_mut()
            .chain(columns.iter_mut())
            .chain(page_fields.iter_mut())
        {
            let Some((_, options)) = field_options.iter().enumerate().find(|(field_index, _)| {
                Self::pivot_cache_field_semantic_name_by_index(cache, *field_index)
                    .eq_ignore_ascii_case(&field.field.name)
            }) else {
                continue;
            };
            if let Some(measure_index) = options.sort_by_measure_index {
                if let Some(measure) = measures.get(measure_index) {
                    field.sort_by_measure = Some(measure.clone());
                }
            }
        }
    }

    fn semantic_groupings_from_pivot_cache_fields(
        fields: &[XlsPivotCacheField],
    ) -> Vec<PivotGrouping> {
        let mut groupings = Vec::new();
        let mut manual_bases = HashSet::new();
        let date_units_by_base = Self::date_units_by_grouping_base(fields);
        let mut emitted_date_bases = HashSet::new();
        for (index, field) in fields.iter().enumerate() {
            if let Some(indexed_units) = date_units_by_base.get(&index) {
                if let Some(grouping) =
                    Self::date_grouping_from_indexed_units(fields, index, indexed_units)
                {
                    groupings.push(grouping);
                    emitted_date_bases.insert(index);
                }
            }

            if let Some(grouping) = Self::manual_grouping_from_cache_field(fields, index, field) {
                if let PivotGrouping::Manual { field, .. } = &grouping {
                    manual_bases.insert(field.name.to_lowercase());
                }
                groupings.push(grouping);
                continue;
            }

            if let Some(grouping) = &field.grouping {
                match grouping {
                    PivotGrouping::Date { .. } => {
                        let base = field.group_base_field.unwrap_or(index);
                        if !emitted_date_bases.contains(&base) {
                            if let Some(indexed_units) = date_units_by_base.get(&base) {
                                if let Some(grouping) = Self::date_grouping_from_indexed_units(
                                    fields,
                                    base,
                                    indexed_units,
                                ) {
                                    groupings.push(grouping);
                                    emitted_date_bases.insert(base);
                                }
                            }
                        }
                    }
                    _ => {
                        let field_name = pivot_grouping_field_name(grouping).to_lowercase();
                        if !manual_bases.contains(&field_name) {
                            groupings.push(grouping.clone());
                        }
                    }
                }
            }
        }

        groupings
    }

    fn date_units_by_grouping_base(
        fields: &[XlsPivotCacheField],
    ) -> BTreeMap<usize, Vec<(usize, PivotDateGroupUnit)>> {
        let mut date_units_by_base: BTreeMap<usize, Vec<(usize, PivotDateGroupUnit)>> =
            BTreeMap::new();
        for (index, field) in fields.iter().enumerate() {
            if let Some(PivotGrouping::Date { units, .. }) = &field.grouping {
                if let Some(unit) = units.first().copied() {
                    date_units_by_base
                        .entry(field.group_base_field.unwrap_or(index))
                        .or_default()
                        .push((index, unit));
                }
            }
        }
        date_units_by_base
    }

    fn date_grouping_from_indexed_units(
        fields: &[XlsPivotCacheField],
        base: usize,
        indexed_units: &[(usize, PivotDateGroupUnit)],
    ) -> Option<PivotGrouping> {
        let field = fields.get(base)?;
        let mut indexed_units = indexed_units.to_vec();
        if let Some(parent_index) = field.group_parent_field {
            indexed_units
                .sort_by_key(|(index, _)| (*index != parent_index, std::cmp::Reverse(*index)));
        } else {
            indexed_units.sort_by_key(|(index, _)| std::cmp::Reverse(*index));
        }
        let units = indexed_units
            .into_iter()
            .map(|(_, unit)| unit)
            .collect::<Vec<_>>();
        Some(PivotGrouping::Date {
            field: PivotFieldRef::new(field.name.clone()),
            units,
        })
    }

    fn manual_grouping_from_cache_field(
        fields: &[XlsPivotCacheField],
        field_index: usize,
        field: &XlsPivotCacheField,
    ) -> Option<PivotGrouping> {
        if field.grouping.is_some() || field.shared_items.is_empty() {
            return None;
        }
        let base_index = field.group_base_field?;
        let base_field = fields.get(base_index)?;
        if base_field.group_parent_field != Some(field_index) || base_field.shared_items.is_empty()
        {
            return None;
        }

        let mut members_by_group: BTreeMap<u32, Vec<PivotValue>> = BTreeMap::new();
        if field.manual_group_item_ids.is_empty() {
            let source_items = base_field
                .shared_items
                .iter()
                .cloned()
                .collect::<HashSet<_>>();
            let ungrouped_items = field
                .shared_items
                .iter()
                .filter(|item| source_items.contains(*item))
                .cloned()
                .collect::<HashSet<_>>();
            let group_names = field
                .shared_items
                .iter()
                .enumerate()
                .filter(|(_, item)| !source_items.contains(*item))
                .map(|(index, _)| index as u32)
                .collect::<Vec<_>>();
            if let [group_index] = group_names.as_slice() {
                for member in base_field
                    .shared_items
                    .iter()
                    .filter(|item| !ungrouped_items.contains(*item))
                    .cloned()
                {
                    members_by_group
                        .entry(*group_index)
                        .or_default()
                        .push(member);
                }
            }
        } else {
            for (base_index, group_index) in field.manual_group_item_ids.iter().copied().enumerate()
            {
                let Some(member) = base_field.shared_items.get(base_index).cloned() else {
                    continue;
                };
                members_by_group
                    .entry(group_index)
                    .or_default()
                    .push(member);
            }
        }

        let groups = members_by_group
            .into_iter()
            .filter_map(|(group_index, members)| {
                let group_value = field.shared_items.get(group_index as usize)?;
                let is_renamed_single_item = members
                    .first()
                    .is_some_and(|member| members.len() == 1 && member != group_value);
                if members.len() <= 1 && !is_renamed_single_item {
                    return None;
                }

                Some(PivotManualGroup {
                    name: group_value.to_string(),
                    members,
                })
            })
            .collect::<Vec<_>>();

        (!groups.is_empty()).then(|| PivotGrouping::Manual {
            field: PivotFieldRef::new(base_field.name.clone()),
            groups,
        })
    }

    fn pivot_cache_field_semantic_name_by_index(cache: &XlsPivotCache, index: usize) -> String {
        let Some(field) = cache.fields.get(index) else {
            return String::new();
        };
        match &field.grouping {
            Some(PivotGrouping::Date {
                field: source_field,
                ..
            }) => source_field.name.clone(),
            None => field
                .group_base_field
                .and_then(|base| cache.fields.get(base))
                .filter(|_| field.group_parent_field.is_none())
                .map(|source| source.name.clone())
                .unwrap_or_else(|| field.name.clone()),
            _ => field.name.clone(),
        }
    }

    fn parse_sxview(data: &[u8]) -> XlsResult<Option<XlsPivotViewBuilder>> {
        if data.len() < 44 {
            return Ok(None);
        }

        let first_row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let last_row = u16::from_le_bytes([data[2], data[3]]) as u32;
        let first_col = u16::from_le_bytes([data[4], data[5]]);
        let last_col = u16::from_le_bytes([data[6], data[7]]);
        let cache_id = u16::from_le_bytes([data[14], data[15]]);
        let page_axis_count = u16::from_le_bytes([data[28], data[29]]);
        let row_axis_count = u16::from_le_bytes([data[24], data[25]]) as usize;
        let column_axis_count = u16::from_le_bytes([data[26], data[27]]) as usize;
        let grbit = u16::from_le_bytes([data[36], data[37]]);
        let name_len = u16::from_le_bytes([data[40], data[41]]);
        let data_caption_len = u16::from_le_bytes([data[42], data[43]]);
        let mut offset = 44usize;
        let name = Self::read_xlunicode_no_cch(data, &mut offset, name_len)
            .unwrap_or_else(|_| format!("PivotTable{}", cache_id.saturating_add(1)));
        let data_caption = if data_caption_len > 0 {
            Self::read_xlunicode_no_cch(data, &mut offset, data_caption_len).ok()
        } else {
            None
        };

        let mut layout = duke_sheets_core::PivotLayout::default();
        layout.show_row_grand_totals = grbit & 0x0001 != 0;
        layout.show_column_grand_totals = grbit & 0x0002 != 0;
        if let Some(data_caption) = data_caption.filter(|caption| !caption.is_empty()) {
            layout.data_caption = data_caption;
        }

        Ok(Some(XlsPivotViewBuilder {
            name,
            cache_id,
            target: CellAddress::new(first_row, first_col),
            rendered_range: Some(CellRange::from_indices(
                first_row, first_col, last_row, last_col,
            )),
            field_options: Vec::new(),
            pending_sxaddl_field_name: None,
            row_axis_count,
            column_axis_count,
            page_axis_count: usize::from(page_axis_count),
            row_axis: Vec::with_capacity(row_axis_count),
            column_axis: Vec::with_capacity(column_axis_count),
            page_fields: Vec::new(),
            measures: Vec::new(),
            layout,
            preserve_formatting: true,
            style: PivotStyle::default(),
            pending_sxaddl_label_filter_field_index: None,
            pending_sxaddl_label_filter_operator: None,
            pending_sxaddl_label_filter_type: None,
            pending_sxaddl_label_filter_values: Vec::new(),
            pending_sxaddl_value_filter: None,
            pending_sxaddl_date_filter: None,
        }))
    }

    fn apply_sxex(builder: &mut XlsPivotViewBuilder, data: &[u8]) {
        if data.len() < 24 {
            return;
        }

        let error_len = u16::from_le_bytes([data[2], data[3]]);
        let missing_len = u16::from_le_bytes([data[4], data[5]]);
        let tag_len = u16::from_le_bytes([data[6], data[7]]);
        let grbit1 = u16::from_le_bytes([data[14], data[15]]);
        let grbit2 = u16::from_le_bytes([data[16], data[17]]);
        let page_field_style_len = u16::from_le_bytes([data[18], data[19]]);
        let table_style_len = u16::from_le_bytes([data[20], data[21]]);
        let vacate_style_len = u16::from_le_bytes([data[22], data[23]]);

        builder.layout.page_over_then_down = grbit1 & 0x0001 != 0;
        builder.layout.page_wrap = u32::from((grbit1 & 0x01FE) >> 1);
        builder.layout.enable_wizard = grbit2 & 0x0001 != 0;
        builder.layout.enable_drill = grbit2 & 0x0002 != 0;
        builder.layout.enable_field_properties = grbit2 & 0x0004 != 0;
        builder.preserve_formatting = grbit2 & 0x0008 != 0;
        builder.layout.merge_item_labels = grbit2 & 0x0010 != 0;
        builder.layout.show_error = grbit2 & 0x0020 != 0;
        builder.layout.show_missing = grbit2 & 0x0040 != 0;
        builder.layout.subtotal_hidden_items = grbit2 & 0x0080 != 0;
        builder.layout.edit_data = grbit2 & 0x0200 != 0;
        builder.layout.disable_field_list = grbit2 & 0x0400 != 0;

        let mut offset = 24usize;
        if let Some(value) = Self::read_optional_xlunicode_no_cch(data, &mut offset, error_len) {
            builder.layout.error_caption = Some(value);
        }
        if let Some(value) = Self::read_optional_xlunicode_no_cch(data, &mut offset, missing_len) {
            builder.layout.missing_caption = Some(value);
        }
        let _ = Self::read_optional_xlunicode_no_cch(data, &mut offset, tag_len);
        let _ = Self::read_optional_xlunicode_no_cch(data, &mut offset, page_field_style_len);
        let _ = Self::read_optional_xlunicode_no_cch(data, &mut offset, table_style_len);
        let _ = Self::read_optional_xlunicode_no_cch(data, &mut offset, vacate_style_len);
    }

    fn apply_sxviewex9(builder: &mut XlsPivotViewBuilder, data: &[u8]) {
        if data.len() < 14 {
            return;
        }
        let grbit = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        builder.layout.field_print_titles = grbit & 0x0002 != 0;
        builder.layout.item_print_titles = grbit & 0x0020 != 0;

        if data.len() > 14 {
            let mut offset = 14usize;
            if let Ok(caption) = read_unicode_string(data, &mut offset) {
                if !caption.is_empty() {
                    builder.layout.grand_total_caption = Some(caption);
                }
            }
        }
    }

    fn read_optional_xlunicode_no_cch(
        data: &[u8],
        offset: &mut usize,
        char_count: u16,
    ) -> Option<String> {
        if char_count == 0xFFFF {
            return None;
        }
        Self::read_xlunicode_no_cch(data, offset, char_count).ok()
    }

    fn parse_sxivd(data: &[u8]) -> Vec<i16> {
        data.chunks_exact(2)
            .map(|chunk| i16::from_le_bytes([chunk[0], chunk[1]]))
            .collect()
    }

    fn parse_sxvd(data: &[u8]) -> XlsPivotFieldOptions {
        if data.len() < 10 {
            return XlsPivotFieldOptions::default();
        }

        let subtotal_count = u16::from_le_bytes([data[2], data[3]]);
        let subtotal_flags = u16::from_le_bytes([data[4], data[5]]);
        let caption_len = u16::from_le_bytes([data[8], data[9]]);
        let mut offset = 10usize;
        let caption = if caption_len != 0xFFFF && caption_len > 0 {
            Self::read_xlunicode_no_cch(data, &mut offset, caption_len).ok()
        } else {
            None
        };
        let subtotals = Self::parse_pivot_subtotals(subtotal_flags);
        let subtotal = subtotals.first().copied().unwrap_or_else(|| {
            if subtotal_count == 0 || subtotal_flags == 0 {
                PivotSubtotal::None
            } else if subtotal_flags & 0x0001 != 0 {
                PivotSubtotal::Automatic
            } else {
                PivotSubtotal::None
            }
        });

        XlsPivotFieldOptions {
            caption,
            subtotal,
            subtotals,
            ..XlsPivotFieldOptions::default()
        }
    }

    fn apply_sxvdex(options: &mut XlsPivotFieldOptions, data: &[u8]) {
        if data.len() < 14 {
            return;
        }

        let grbit1 = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        let sort_index = i16::from_le_bytes([data[4], data[5]]);
        let auto_show_measure_index = i16::from_le_bytes([data[6], data[7]]);

        options.show_empty_items = grbit1 & 0x0001 != 0;
        options.sort = Self::parse_pivot_sort(grbit1);
        options.sort_by_measure_index = (sort_index >= 0).then_some(sort_index as usize);
        options.insert_page_break = grbit1 & 0x4000 != 0;
        options.include_new_items_in_filter = grbit1 & 0x8000 == 0;
        options.insert_blank_row = grbit1 & 0x0040_0000 != 0;
        options.subtotal_top = grbit1 & 0x0080_0000 != 0;
        let auto_show_count = grbit1 >> 24;
        if grbit1 & 0x0800 != 0 && auto_show_count != 0 && auto_show_measure_index >= 0 {
            options.top_n_filter = Some(XlsTopNFilter {
                n: auto_show_count,
                top: grbit1 & 0x1000 != 0,
                percent: false,
                measure_index: auto_show_measure_index as usize,
            });
        } else if auto_show_count != 0 {
            options.item_page_count = auto_show_count;
        }

        if data.len() < 20 {
            return;
        }

        let subtotal_caption_len = u16::from_le_bytes([data[10], data[11]]);

        if subtotal_caption_len != 0xFFFF && subtotal_caption_len > 0 {
            let mut offset = 20usize;
            options.subtotal_caption =
                Self::read_xlunicode_no_cch(data, &mut offset, subtotal_caption_len).ok();
        }
    }

    fn apply_sxvi(options: &mut XlsPivotFieldOptions, data: &[u8]) {
        if data.len() < 6 {
            return;
        }
        let flags = u16::from_le_bytes([data[2], data[3]]);
        let item_index = i16::from_le_bytes([data[4], data[5]]);
        if flags & 0x0001 != 0 && item_index >= 0 {
            options.hidden_items.push(item_index as usize);
        }
        if flags & 0x0008 != 0 && item_index >= 0 {
            options.calculated_items.push(item_index as usize);
        }
    }

    fn parse_pivot_sort(grbit1: u32) -> PivotSort {
        if grbit1 & 0x0200 == 0 {
            PivotSort::None
        } else if grbit1 & 0x0400 != 0 {
            PivotSort::Ascending
        } else {
            PivotSort::Descending
        }
    }

    fn parse_pivot_subtotals(flags: u16) -> Vec<PivotSubtotal> {
        let mut subtotals = Vec::new();
        for (mask, subtotal) in [
            (0x0002, PivotSubtotal::Sum),
            (0x0004, PivotSubtotal::Count),
            (0x0008, PivotSubtotal::Average),
            (0x0010, PivotSubtotal::Max),
            (0x0020, PivotSubtotal::Min),
            (0x0040, PivotSubtotal::Product),
            (0x0080, PivotSubtotal::CountNumbers),
            (0x0100, PivotSubtotal::StdDev),
            (0x0200, PivotSubtotal::StdDevP),
            (0x0400, PivotSubtotal::Var),
            (0x0800, PivotSubtotal::VarP),
        ] {
            if flags & mask != 0 {
                subtotals.push(subtotal);
            }
        }
        subtotals
    }

    fn parse_sxpi(data: &[u8]) -> Option<(usize, u16)> {
        if data.len() < 4 {
            return None;
        }
        Some((
            u16::from_le_bytes([data[0], data[1]]) as usize,
            u16::from_le_bytes([data[2], data[3]]),
        ))
    }

    fn parse_sxdi(
        data: &[u8],
        caches: &std::collections::HashMap<u16, XlsPivotCache>,
        cache_id: u16,
        num_fmts: &std::collections::HashMap<u16, String>,
    ) -> XlsResult<Option<PivotMeasure>> {
        if data.len() < 12 {
            return Ok(None);
        }
        let field_index = u16::from_le_bytes([data[0], data[1]]) as usize;
        let aggregate = Self::parse_pivot_aggregate(u16::from_le_bytes([data[2], data[3]]));
        let show_as = u16::from_le_bytes([data[4], data[5]]);
        let base_field = u16::from_le_bytes([data[6], data[7]]) as usize;
        let base_item = u16::from_le_bytes([data[8], data[9]]) as usize;
        let num_format = u16::from_le_bytes([data[10], data[11]]) as u32;
        let cache = match caches.get(&cache_id) {
            Some(cache) => cache,
            None => return Ok(None),
        };
        let Some(field) = cache.fields.get(field_index) else {
            return Ok(None);
        };
        let mut offset = 12usize;
        let caption = read_unicode_string(data, &mut offset).unwrap_or_default();
        let mut measure = PivotMeasure::new(field.name.clone(), aggregate);
        if !caption.is_empty() {
            measure.name = Some(caption);
        }
        measure.show_as = Self::parse_data_field_show_as(show_as, base_field, base_item, cache)
            .unwrap_or(PivotShowAs::Normal);
        measure.number_format = Self::pivot_number_format_code(num_format, num_fmts);
        Ok(Some(measure))
    }

    fn parse_data_field_show_as(
        code: u16,
        base_field_index: usize,
        base_item_index: usize,
        cache: &XlsPivotCache,
    ) -> Option<PivotShowAs> {
        let base_field = || {
            cache
                .fields
                .get(base_field_index)
                .map(|field| PivotFieldRef::new(field.name.clone()))
        };
        let base_item = || {
            cache
                .fields
                .get(base_field_index)?
                .shared_items
                .get(base_item_index)
                .cloned()
        };
        Some(match code {
            0 => PivotShowAs::Normal,
            1 => PivotShowAs::DifferenceFrom {
                base_field: base_field()?,
                base_item: base_item()?,
            },
            3 => PivotShowAs::PercentDifferenceFrom {
                base_field: base_field()?,
                base_item: base_item()?,
            },
            4 => PivotShowAs::RunningTotal {
                base_field: base_field()?,
            },
            5 => PivotShowAs::PercentOfRowTotal,
            6 => PivotShowAs::PercentOfColumnTotal,
            7 => PivotShowAs::PercentOfGrandTotal,
            8 => PivotShowAs::Index,
            _ => return None,
        })
    }

    fn pivot_number_format_code(
        num_fmt_id: u32,
        num_fmts: &std::collections::HashMap<u16, String>,
    ) -> Option<String> {
        if num_fmt_id == NumberFormat::ID_GENERAL {
            return None;
        }
        if let Ok(num_fmt_id_u16) = u16::try_from(num_fmt_id) {
            if let Some(code) = num_fmts.get(&num_fmt_id_u16) {
                return Some(code.clone());
            }
        }
        let format = NumberFormat::BuiltIn(num_fmt_id);
        let code = format.format_string();
        (code != "General").then(|| code.to_string())
    }

    fn parse_pivot_aggregate(code: u16) -> PivotAggregate {
        match code {
            1 => PivotAggregate::Count,
            2 => PivotAggregate::Average,
            3 => PivotAggregate::Max,
            4 => PivotAggregate::Min,
            5 => PivotAggregate::Product,
            6 => PivotAggregate::CountNumbers,
            7 => PivotAggregate::StdDev,
            8 => PivotAggregate::StdDevP,
            9 => PivotAggregate::Var,
            10 => PivotAggregate::VarP,
            _ => PivotAggregate::Sum,
        }
    }

    fn parse_pivot_frt0864_style(data: &[u8]) -> Option<PivotStyle> {
        if data.len() < 16 {
            return None;
        }
        let subtype = u16::from_le_bytes([data[4], data[5]]);
        if subtype != 0x001E && subtype != 0x1E00 {
            return None;
        }
        let flags = u16::from_le_bytes([data[12], data[13]]);
        let char_count = u16::from_le_bytes([data[14], data[15]]) as usize;
        let byte_len = char_count.checked_mul(2)?;
        let end = 16usize.checked_add(byte_len)?;
        if end > data.len() {
            return None;
        }
        let units = data[16..end]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        let name = String::from_utf16_lossy(&units);
        let mut style = PivotStyle::default();
        style.show_last_column = (flags & 0x02) != 0;
        style.show_row_stripes = (flags & 0x04) != 0;
        style.show_column_stripes = (flags & 0x08) != 0;
        style.show_row_headers = (flags & 0x10) != 0;
        style.show_column_headers = (flags & 0x20) != 0;
        if !name.is_empty() {
            style.name = Some(name);
        }
        Some(style)
    }

    fn apply_pivot_frt0864_field_options(
        builder: &mut XlsPivotViewBuilder,
        caches: &std::collections::HashMap<u16, XlsPivotCache>,
        data: &[u8],
    ) {
        if data.len() < 12 {
            return;
        }
        match (data[4], data[5]) {
            (0x01, 0x00) => {
                let mut offset = 12usize;
                builder.pending_sxaddl_field_name = read_unicode_string(data, &mut offset).ok();
            }
            (0x17, 0x00) => {
                let mut offset = 12usize;
                builder.pending_sxaddl_field_name = read_unicode_string(data, &mut offset).ok();
            }
            (0x01, 0x02) => {
                let Some(field_name) = builder.pending_sxaddl_field_name.as_deref() else {
                    return;
                };
                let Some(cache) = caches.get(&builder.cache_id) else {
                    return;
                };
                let flags = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);
                if let Some((field_index, _)) =
                    builder
                        .field_options
                        .iter()
                        .enumerate()
                        .find(|(field_index, _)| {
                            Self::pivot_cache_field_semantic_name_by_index(cache, *field_index)
                                .eq_ignore_ascii_case(field_name)
                        })
                {
                    if let Some(options) = builder.field_options.get_mut(field_index) {
                        options.show_drop_downs = flags & 0x0000_0001 == 0;
                    }
                }
            }
            (0x17, 0x19) => {
                let Some(field_name) = builder.pending_sxaddl_field_name.as_deref() else {
                    return;
                };
                let Some(cache) = caches.get(&builder.cache_id) else {
                    return;
                };
                let flags = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);
                if let Some((field_index, _)) =
                    builder
                        .field_options
                        .iter()
                        .enumerate()
                        .find(|(field_index, _)| {
                            Self::pivot_cache_field_semantic_name_by_index(cache, *field_index)
                                .eq_ignore_ascii_case(field_name)
                        })
                {
                    if let Some(options) = builder.field_options.get_mut(field_index) {
                        options.include_new_items_in_filter = flags & 0x0000_0020 == 0;
                    }
                }
            }
            (0x17, 0x37) => {
                let Some(field_name) = builder.pending_sxaddl_field_name.as_deref() else {
                    return;
                };
                let Some(cache) = caches.get(&builder.cache_id) else {
                    return;
                };
                let count = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);
                if let Some((field_index, _)) =
                    builder
                        .field_options
                        .iter()
                        .enumerate()
                        .find(|(field_index, _)| {
                            Self::pivot_cache_field_semantic_name_by_index(cache, *field_index)
                                .eq_ignore_ascii_case(field_name)
                        })
                {
                    if let Some(options) = builder.field_options.get_mut(field_index) {
                        if let Some(filter) = options.top_n_filter.as_mut() {
                            filter.n = count;
                        }
                    }
                }
            }
            (0x1D, 0x38) => Self::apply_pivot_frt0864_filter_collection(builder, data),
            (0x1D, 0x3C) => {
                Self::apply_pivot_frt0864_top_n_filter(builder, caches, data);
                Self::apply_pivot_frt0864_label_filter_header(builder, data);
                Self::apply_pivot_frt0864_value_filter(builder, data);
                Self::apply_pivot_frt0864_date_filter(builder, data);
            }
            (0x1D, 0x3D) | (0x1D, 0x3E) => {
                Self::apply_pivot_frt0864_label_filter_value(builder, data)
            }
            (0x01, 0xFF) | (0x17, 0xFF) => builder.pending_sxaddl_field_name = None,
            (0x1D, 0xFF) | (0x1C, 0xFF) => {
                builder.pending_sxaddl_label_filter_field_index = None;
                builder.pending_sxaddl_label_filter_operator = None;
                builder.pending_sxaddl_label_filter_type = None;
                builder.pending_sxaddl_label_filter_values.clear();
                builder.pending_sxaddl_value_filter = None;
                builder.pending_sxaddl_date_filter = None;
            }
            _ => {}
        }
    }

    fn apply_pivot_frt0864_filter_collection(builder: &mut XlsPivotViewBuilder, data: &[u8]) {
        builder.pending_sxaddl_label_filter_field_index = None;
        builder.pending_sxaddl_label_filter_operator = None;
        builder.pending_sxaddl_label_filter_type = None;
        builder.pending_sxaddl_label_filter_values.clear();
        builder.pending_sxaddl_value_filter = None;
        builder.pending_sxaddl_date_filter = None;
        if data.len() < 24 {
            return;
        }
        let field_index = u32::from_le_bytes([data[12], data[13], data[14], data[15]]) as usize;
        let filter_type = u32::from_le_bytes([data[20], data[21], data[22], data[23]]);
        let operator = match filter_type {
            4 => Some(PivotFilterOperator::Equals),
            5 => Some(PivotFilterOperator::NotEquals),
            6 => Some(PivotFilterOperator::BeginsWith),
            7 => Some(PivotFilterOperator::DoesNotBeginWith),
            8 => Some(PivotFilterOperator::EndsWith),
            9 => Some(PivotFilterOperator::DoesNotEndWith),
            10 => Some(PivotFilterOperator::Contains),
            11 => Some(PivotFilterOperator::DoesNotContain),
            12 => Some(PivotFilterOperator::GreaterThan),
            13 => Some(PivotFilterOperator::GreaterThanOrEqual),
            14 => Some(PivotFilterOperator::LessThan),
            15 => Some(PivotFilterOperator::LessThanOrEqual),
            16 | 17 => None,
            _ => {
                if matches!(filter_type, 18..=25) && data.len() >= 32 {
                    let measure_index =
                        i32::from_le_bytes([data[28], data[29], data[30], data[31]]);
                    if field_index < builder.field_options.len() && measure_index >= 0 {
                        builder.pending_sxaddl_value_filter = Some(XlsPendingValueFilter {
                            field_index,
                            measure_index: measure_index as usize,
                            filter_type,
                        });
                    }
                } else if Self::is_xls_date_filter_type(filter_type)
                    && field_index < builder.field_options.len()
                {
                    builder.pending_sxaddl_date_filter = Some(XlsPendingDateFilter {
                        field_index,
                        filter_type,
                    });
                }
                return;
            }
        };
        if field_index < builder.field_options.len() {
            builder.pending_sxaddl_label_filter_field_index = Some(field_index);
            builder.pending_sxaddl_label_filter_operator = operator;
            builder.pending_sxaddl_label_filter_type = Some(filter_type);
        }
    }

    fn apply_pivot_frt0864_top_n_filter(
        builder: &mut XlsPivotViewBuilder,
        caches: &std::collections::HashMap<u16, XlsPivotCache>,
        data: &[u8],
    ) {
        if data.len() < 40 {
            return;
        }
        let Some(cache) = caches.get(&builder.cache_id) else {
            return;
        };
        let filter_type = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
        let (top, percent) = match filter_type {
            1 => (true, false),
            2 => (false, false),
            3 => (true, true),
            4 => (false, true),
            _ => return,
        };
        let field_index = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
        let measure_field_index = u32::from_le_bytes([data[20], data[21], data[22], data[23]]);
        if field_index == 0 || measure_field_index == 0 {
            return;
        }
        let field_index = (field_index - 1) as usize;
        let measure_field_index = (measure_field_index - 1) as usize;
        if field_index >= builder.field_options.len() {
            return;
        }
        let measure_field_name =
            Self::pivot_cache_field_semantic_name_by_index(cache, measure_field_index);
        let Some(measure_index) = builder
            .measures
            .iter()
            .position(|measure| measure.field.name.eq_ignore_ascii_case(&measure_field_name))
        else {
            return;
        };
        let value = Self::parse_pivot_xnum(&data[32..40]);
        if !value.is_finite() || value < 1.0 || value > u32::MAX as f64 {
            return;
        }
        builder.field_options[field_index].top_n_filter = Some(XlsTopNFilter {
            n: value.round() as u32,
            top,
            percent,
            measure_index,
        });
    }

    fn apply_pivot_frt0864_label_filter_header(builder: &mut XlsPivotViewBuilder, data: &[u8]) {
        if builder.pending_sxaddl_label_filter_field_index.is_none() {
            return;
        }
        let Some(filter_type) = builder.pending_sxaddl_label_filter_type else {
            builder.pending_sxaddl_label_filter_field_index = None;
            builder.pending_sxaddl_label_filter_operator = None;
            builder.pending_sxaddl_label_filter_values.clear();
            return;
        };
        if matches!(filter_type, 16 | 17) {
            if data.len() < 44 {
                builder.pending_sxaddl_label_filter_field_index = None;
                builder.pending_sxaddl_label_filter_type = None;
                builder.pending_sxaddl_label_filter_values.clear();
                return;
            }
            let count = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
            let first = &data[20..24];
            let second = &data[30..34];
            let flag = u32::from_le_bytes([data[40], data[41], data[42], data[43]]);
            let valid = match filter_type {
                16 => {
                    count == 2
                        && first == [0x06, 0x06, 0x01, 0x00]
                        && second == [0x06, 0x03, 0x01, 0x00]
                        && flag == 1
                }
                17 => {
                    count == 2
                        && first == [0x06, 0x01, 0x01, 0x00]
                        && second == [0x06, 0x04, 0x01, 0x00]
                        && flag == 2
                }
                _ => false,
            };
            if !valid {
                builder.pending_sxaddl_label_filter_field_index = None;
                builder.pending_sxaddl_label_filter_type = None;
                builder.pending_sxaddl_label_filter_values.clear();
            }
            return;
        }
        if data.len() < 24 {
            builder.pending_sxaddl_label_filter_field_index = None;
            builder.pending_sxaddl_label_filter_type = None;
            builder.pending_sxaddl_label_filter_values.clear();
            return;
        }
        let Some(operator) = builder.pending_sxaddl_label_filter_operator else {
            builder.pending_sxaddl_label_filter_field_index = None;
            builder.pending_sxaddl_label_filter_type = None;
            builder.pending_sxaddl_label_filter_values.clear();
            return;
        };
        let (expected_custom_operator_code, expected_discriminator) = match operator {
            PivotFilterOperator::Equals => (0x02, 0x01),
            PivotFilterOperator::BeginsWith
            | PivotFilterOperator::EndsWith
            | PivotFilterOperator::Contains => (0x02, 0x00),
            PivotFilterOperator::NotEquals => (0x05, 0x01),
            PivotFilterOperator::DoesNotBeginWith
            | PivotFilterOperator::DoesNotEndWith
            | PivotFilterOperator::DoesNotContain => (0x05, 0x00),
            PivotFilterOperator::LessThan => (0x01, 0x01),
            PivotFilterOperator::LessThanOrEqual => (0x03, 0x01),
            PivotFilterOperator::GreaterThan => (0x04, 0x01),
            PivotFilterOperator::GreaterThanOrEqual => (0x06, 0x01),
        };
        if data[20] != 0x06
            || data[21] != expected_custom_operator_code
            || data[22] != expected_discriminator
        {
            builder.pending_sxaddl_label_filter_field_index = None;
            builder.pending_sxaddl_label_filter_operator = None;
            builder.pending_sxaddl_label_filter_type = None;
            builder.pending_sxaddl_label_filter_values.clear();
        }
    }

    fn apply_pivot_frt0864_label_filter_value(builder: &mut XlsPivotViewBuilder, data: &[u8]) {
        let Some(field_index) = builder.pending_sxaddl_label_filter_field_index else {
            return;
        };
        if data.len() < 12 {
            builder.pending_sxaddl_label_filter_field_index = None;
            builder.pending_sxaddl_label_filter_operator = None;
            builder.pending_sxaddl_label_filter_type = None;
            builder.pending_sxaddl_label_filter_values.clear();
            return;
        }
        let mut offset = 12usize;
        let Ok(raw_value) = read_unicode_string(data, &mut offset) else {
            builder.pending_sxaddl_label_filter_field_index = None;
            builder.pending_sxaddl_label_filter_operator = None;
            builder.pending_sxaddl_label_filter_type = None;
            builder.pending_sxaddl_label_filter_values.clear();
            return;
        };
        if matches!(builder.pending_sxaddl_label_filter_type, Some(16 | 17)) {
            if raw_value.is_empty() {
                builder.pending_sxaddl_label_filter_field_index = None;
                builder.pending_sxaddl_label_filter_type = None;
                builder.pending_sxaddl_label_filter_values.clear();
                return;
            }
            builder.pending_sxaddl_label_filter_values.push(raw_value);
            if builder.pending_sxaddl_label_filter_values.len() >= 2 {
                let start = builder.pending_sxaddl_label_filter_values[0].clone();
                let end = builder.pending_sxaddl_label_filter_values[1].clone();
                let not_between = builder.pending_sxaddl_label_filter_type == Some(17);
                if let Some(options) = builder.field_options.get_mut(field_index) {
                    options.label_filter = Some(XlsLabelFilter {
                        kind: XlsLabelFilterKind::Between {
                            start,
                            end,
                            not_between,
                        },
                    });
                }
                builder.pending_sxaddl_label_filter_field_index = None;
                builder.pending_sxaddl_label_filter_operator = None;
                builder.pending_sxaddl_label_filter_type = None;
                builder.pending_sxaddl_label_filter_values.clear();
            }
            return;
        }
        let Some(operator) = builder.pending_sxaddl_label_filter_operator else {
            builder.pending_sxaddl_label_filter_field_index = None;
            builder.pending_sxaddl_label_filter_type = None;
            builder.pending_sxaddl_label_filter_values.clear();
            return;
        };
        let value = match operator {
            PivotFilterOperator::NotEquals => raw_value,
            PivotFilterOperator::LessThan
            | PivotFilterOperator::LessThanOrEqual
            | PivotFilterOperator::GreaterThan
            | PivotFilterOperator::GreaterThanOrEqual => raw_value,
            PivotFilterOperator::BeginsWith => {
                let Some(value) = raw_value.strip_suffix('*') else {
                    builder.pending_sxaddl_label_filter_field_index = None;
                    builder.pending_sxaddl_label_filter_operator = None;
                    builder.pending_sxaddl_label_filter_type = None;
                    builder.pending_sxaddl_label_filter_values.clear();
                    return;
                };
                value.to_string()
            }
            PivotFilterOperator::DoesNotBeginWith => {
                let Some(value) = raw_value.strip_suffix('*') else {
                    builder.pending_sxaddl_label_filter_field_index = None;
                    builder.pending_sxaddl_label_filter_operator = None;
                    builder.pending_sxaddl_label_filter_type = None;
                    builder.pending_sxaddl_label_filter_values.clear();
                    return;
                };
                value.to_string()
            }
            PivotFilterOperator::EndsWith => {
                let Some(value) = raw_value.strip_prefix('*') else {
                    builder.pending_sxaddl_label_filter_field_index = None;
                    builder.pending_sxaddl_label_filter_operator = None;
                    builder.pending_sxaddl_label_filter_type = None;
                    builder.pending_sxaddl_label_filter_values.clear();
                    return;
                };
                value.to_string()
            }
            PivotFilterOperator::DoesNotEndWith => {
                let Some(value) = raw_value.strip_prefix('*') else {
                    builder.pending_sxaddl_label_filter_field_index = None;
                    builder.pending_sxaddl_label_filter_operator = None;
                    builder.pending_sxaddl_label_filter_type = None;
                    builder.pending_sxaddl_label_filter_values.clear();
                    return;
                };
                value.to_string()
            }
            PivotFilterOperator::Contains => {
                let Some(value) = raw_value
                    .strip_prefix('*')
                    .and_then(|value| value.strip_suffix('*'))
                else {
                    builder.pending_sxaddl_label_filter_field_index = None;
                    builder.pending_sxaddl_label_filter_operator = None;
                    builder.pending_sxaddl_label_filter_type = None;
                    builder.pending_sxaddl_label_filter_values.clear();
                    return;
                };
                value.to_string()
            }
            PivotFilterOperator::DoesNotContain => {
                let Some(value) = raw_value
                    .strip_prefix('*')
                    .and_then(|value| value.strip_suffix('*'))
                else {
                    builder.pending_sxaddl_label_filter_field_index = None;
                    builder.pending_sxaddl_label_filter_operator = None;
                    builder.pending_sxaddl_label_filter_type = None;
                    builder.pending_sxaddl_label_filter_values.clear();
                    return;
                };
                value.to_string()
            }
            _ => raw_value,
        };
        if value.is_empty() {
            builder.pending_sxaddl_label_filter_field_index = None;
            builder.pending_sxaddl_label_filter_operator = None;
            builder.pending_sxaddl_label_filter_type = None;
            builder.pending_sxaddl_label_filter_values.clear();
            return;
        }
        if let Some(options) = builder.field_options.get_mut(field_index) {
            options.label_filter = Some(XlsLabelFilter {
                kind: XlsLabelFilterKind::Comparison { operator, value },
            });
        }
        builder.pending_sxaddl_label_filter_field_index = None;
        builder.pending_sxaddl_label_filter_operator = None;
        builder.pending_sxaddl_label_filter_type = None;
        builder.pending_sxaddl_label_filter_values.clear();
    }

    fn apply_pivot_frt0864_value_filter(builder: &mut XlsPivotViewBuilder, data: &[u8]) {
        let Some(pending) = builder.pending_sxaddl_value_filter else {
            return;
        };
        let Some(kind) = Self::parse_pivot_frt0864_value_filter_kind(pending.filter_type, data)
        else {
            builder.pending_sxaddl_value_filter = None;
            return;
        };
        if let Some(options) = builder.field_options.get_mut(pending.field_index) {
            options.value_filter = Some(XlsValueFilter {
                measure_index: pending.measure_index,
                kind,
            });
        }
        builder.pending_sxaddl_value_filter = None;
    }

    fn apply_pivot_frt0864_date_filter(builder: &mut XlsPivotViewBuilder, data: &[u8]) {
        let Some(pending) = builder.pending_sxaddl_date_filter else {
            return;
        };
        let Some(kind) = Self::parse_pivot_frt0864_date_filter_kind(pending.filter_type, data)
        else {
            builder.pending_sxaddl_date_filter = None;
            return;
        };
        if let Some(options) = builder.field_options.get_mut(pending.field_index) {
            options.date_filter = Some(XlsDateFilter { kind });
        }
        builder.pending_sxaddl_date_filter = None;
    }

    fn parse_pivot_frt0864_value_filter_kind(
        filter_type: u32,
        data: &[u8],
    ) -> Option<XlsValueFilterKind> {
        match filter_type {
            18..=23 => {
                if data.len() < 30 {
                    return None;
                }
                let (operator, expected_custom_operator) =
                    Self::xls_value_filter_operator_for_type(filter_type)?;
                let criterion = Self::parse_frt0864_numeric_filter_criterion(data, 20)?;
                if criterion.0 != expected_custom_operator {
                    return None;
                }
                Some(XlsValueFilterKind::Comparison {
                    operator,
                    value: criterion.1,
                })
            }
            24 | 25 => {
                if data.len() < 44 {
                    return None;
                }
                let count = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
                if count != 2 {
                    return None;
                }
                let first = Self::parse_frt0864_numeric_filter_criterion(data, 20)?;
                let second = Self::parse_frt0864_numeric_filter_criterion(data, 30)?;
                let flag = u32::from_le_bytes([data[40], data[41], data[42], data[43]]);
                match filter_type {
                    24 if first.0 == 0x06 && second.0 == 0x03 && flag == 1 => {
                        Some(XlsValueFilterKind::Between {
                            start: first.1,
                            end: second.1,
                            not_between: false,
                        })
                    }
                    25 if first.0 == 0x01 && second.0 == 0x04 && flag == 2 => {
                        Some(XlsValueFilterKind::Between {
                            start: first.1,
                            end: second.1,
                            not_between: true,
                        })
                    }
                    _ => None,
                }
            }
            _ => None,
        }
    }

    fn xls_value_filter_operator_for_type(filter_type: u32) -> Option<(PivotFilterOperator, u8)> {
        match filter_type {
            18 => Some((PivotFilterOperator::Equals, 0x02)),
            19 => Some((PivotFilterOperator::NotEquals, 0x05)),
            20 => Some((PivotFilterOperator::GreaterThan, 0x04)),
            21 => Some((PivotFilterOperator::GreaterThanOrEqual, 0x06)),
            22 => Some((PivotFilterOperator::LessThan, 0x01)),
            23 => Some((PivotFilterOperator::LessThanOrEqual, 0x03)),
            _ => None,
        }
    }

    fn parse_pivot_frt0864_date_filter_kind(
        filter_type: u32,
        data: &[u8],
    ) -> Option<XlsDateFilterKind> {
        match filter_type {
            26 | 27 | 28 | 62 | 63 | 64 => {
                if data.len() < 30 {
                    return None;
                }
                let (operator, expected_custom_operator, expected_date_filter_type) =
                    Self::xls_date_filter_operator_for_type(filter_type)?;
                let actual_date_filter_type =
                    u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
                let count = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
                if actual_date_filter_type != expected_date_filter_type || count != 1 {
                    return None;
                }
                let criterion = Self::parse_frt0864_numeric_filter_criterion(data, 20)?;
                if criterion.0 != expected_custom_operator {
                    return None;
                }
                Some(XlsDateFilterKind::Comparison {
                    operator,
                    value: criterion.1,
                })
            }
            29 | 65 => {
                if data.len() < 44 {
                    return None;
                }
                let expected_filter_type = if filter_type == 29 { 7 } else { 43 };
                let actual_filter_type =
                    u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
                let count = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
                if actual_filter_type != expected_filter_type || count != 2 {
                    return None;
                }
                let first = Self::parse_frt0864_numeric_filter_criterion(data, 20)?;
                let second = Self::parse_frt0864_numeric_filter_criterion(data, 30)?;
                let flag = u32::from_le_bytes([data[40], data[41], data[42], data[43]]);
                match filter_type {
                    29 if first.0 == 0x06 && second.0 == 0x03 && flag == 1 => {
                        Some(XlsDateFilterKind::Between {
                            start: first.1,
                            end: second.1,
                            not_between: false,
                        })
                    }
                    65 if first.0 == 0x01 && second.0 == 0x04 && flag == 2 => {
                        Some(XlsDateFilterKind::Between {
                            start: first.1,
                            end: second.1,
                            not_between: true,
                        })
                    }
                    _ => None,
                }
            }
            30..=61 => {
                if data.len() < 16 {
                    return None;
                }
                let cft = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
                if filter_type != cft + 22 {
                    return None;
                }
                Some(XlsDateFilterKind::Period(Self::xls_date_period_for_cft(
                    cft,
                )?))
            }
            _ => None,
        }
    }

    fn is_xls_date_filter_type(filter_type: u32) -> bool {
        matches!(filter_type, 26 | 27 | 28 | 29 | 62 | 63 | 64 | 65)
            || (30..=61).contains(&filter_type)
    }

    fn xls_date_filter_operator_for_type(
        filter_type: u32,
    ) -> Option<(PivotFilterOperator, u8, u32)> {
        match filter_type {
            26 => Some((PivotFilterOperator::Equals, 0x02, 4)),
            62 => Some((PivotFilterOperator::NotEquals, 0x05, 40)),
            27 => Some((PivotFilterOperator::LessThan, 0x01, 5)),
            63 => Some((PivotFilterOperator::LessThanOrEqual, 0x03, 41)),
            28 => Some((PivotFilterOperator::GreaterThan, 0x04, 6)),
            64 => Some((PivotFilterOperator::GreaterThanOrEqual, 0x06, 42)),
            _ => None,
        }
    }

    fn xls_date_period_for_cft(cft: u32) -> Option<PivotDatePeriod> {
        Some(match cft {
            0x08 => PivotDatePeriod::Tomorrow,
            0x09 => PivotDatePeriod::Today,
            0x0A => PivotDatePeriod::Yesterday,
            0x0B => PivotDatePeriod::NextWeek,
            0x0C => PivotDatePeriod::ThisWeek,
            0x0D => PivotDatePeriod::LastWeek,
            0x0E => PivotDatePeriod::NextMonth,
            0x0F => PivotDatePeriod::ThisMonth,
            0x10 => PivotDatePeriod::LastMonth,
            0x11 => PivotDatePeriod::NextQuarter,
            0x12 => PivotDatePeriod::ThisQuarter,
            0x13 => PivotDatePeriod::LastQuarter,
            0x14 => PivotDatePeriod::NextYear,
            0x15 => PivotDatePeriod::ThisYear,
            0x16 => PivotDatePeriod::LastYear,
            0x17 => PivotDatePeriod::YearToDate,
            0x18 => PivotDatePeriod::Quarter(1),
            0x19 => PivotDatePeriod::Quarter(2),
            0x1A => PivotDatePeriod::Quarter(3),
            0x1B => PivotDatePeriod::Quarter(4),
            0x1C => PivotDatePeriod::Month(1),
            0x1D => PivotDatePeriod::Month(2),
            0x1E => PivotDatePeriod::Month(3),
            0x1F => PivotDatePeriod::Month(4),
            0x20 => PivotDatePeriod::Month(5),
            0x21 => PivotDatePeriod::Month(6),
            0x22 => PivotDatePeriod::Month(7),
            0x23 => PivotDatePeriod::Month(8),
            0x24 => PivotDatePeriod::Month(9),
            0x25 => PivotDatePeriod::Month(10),
            0x26 => PivotDatePeriod::Month(11),
            0x27 => PivotDatePeriod::Month(12),
            _ => return None,
        })
    }

    fn parse_frt0864_numeric_filter_criterion(data: &[u8], offset: usize) -> Option<(u8, f64)> {
        if data.len() < offset.saturating_add(10) || data[offset] != 0x04 {
            return None;
        }
        let mut value = [0u8; 8];
        value.copy_from_slice(&data[offset + 2..offset + 10]);
        let value = f64::from_le_bytes(value);
        value.is_finite().then_some((data[offset + 1], value))
    }

    fn parse_pivot_xnum(data: &[u8]) -> f64 {
        let mut bytes = [0u8; 8];
        bytes[0..2].copy_from_slice(&data[6..8]);
        bytes[2..4].copy_from_slice(&data[4..6]);
        bytes[4..6].copy_from_slice(&data[2..4]);
        bytes[6..8].copy_from_slice(&data[0..2]);
        f64::from_le_bytes(bytes)
    }

    fn parse_sxfdb_name(data: &[u8]) -> XlsResult<String> {
        if data.len() < 17 {
            return Ok(String::new());
        }
        let mut offset = 14usize;
        read_unicode_string(data, &mut offset)
    }

    fn parse_sxfdb_field_index(data: &[u8], offset: usize) -> Option<usize> {
        if data.len() < offset.saturating_add(2) {
            return None;
        }
        let value = i16::from_le_bytes([data[offset], data[offset + 1]]);
        (value >= 0).then_some(value as usize)
    }

    fn parse_sxstring(data: &[u8]) -> XlsResult<PivotValue> {
        let mut offset = 0usize;
        let text = read_unicode_string(data, &mut offset)?;
        if text.is_empty() {
            Ok(PivotValue::Blank)
        } else {
            Ok(PivotValue::String(text))
        }
    }

    fn parse_sxdtr(data: &[u8], date_system: DateSystem) -> Option<f64> {
        if data.len() < 8 {
            return None;
        }
        let year = u16::from_le_bytes([data[0], data[1]]) as i32;
        let month = u16::from_le_bytes([data[2], data[3]]) as u32;
        let day = u32::from(data[4]);
        let hour = u32::from(data[5]);
        let minute = u32::from(data[6]);
        let second = u32::from(data[7]);
        if !(1900..=9999).contains(&year)
            || !(1..=12).contains(&month)
            || day > 31
            || (day == 0 && (year != 1900 || month != 1))
            || hour > 23
            || minute > 59
            || second > 59
        {
            return None;
        }
        Some(
            date_to_serial(year, month, day, date_system)
                + f64::from(hour * 3600 + minute * 60 + second) / 86_400.0,
        )
    }

    fn parse_sxrng(data: &[u8]) -> Option<PendingPivotRangeGroup> {
        if data.len() < 2 {
            return None;
        }
        Some(PendingPivotRangeGroup {
            flags: u16::from_le_bytes([data[0], data[1]]),
            numbers: Vec::with_capacity(3),
        })
    }

    fn parse_sxfmla(data: &[u8]) -> Option<XlsPivotFormulaPending> {
        if data.len() < 4 {
            return None;
        }
        let token_len = u16::from_le_bytes([data[0], data[1]]) as usize;
        let expected_sx_names = u16::from_le_bytes([data[2], data[3]]) as usize;
        let token_end = 4usize.checked_add(token_len)?;
        if token_end > data.len() {
            return None;
        }
        Some(XlsPivotFormulaPending {
            ptgs: data[4..token_end].to_vec(),
            expected_sx_names,
            sx_name_field_indexes: Vec::with_capacity(expected_sx_names),
        })
    }

    fn parse_sxname(data: &[u8]) -> Option<usize> {
        if data.len() < 8 {
            return None;
        }
        let field_index = i16::from_le_bytes([data[2], data[3]]);
        (field_index >= 0).then_some(field_index as usize)
    }

    fn parse_sxname_item_pair_count(data: &[u8]) -> Option<usize> {
        if data.len() < 8 {
            return None;
        }
        let field_index = i16::from_le_bytes([data[2], data[3]]);
        let pair_count = u16::from_le_bytes([data[6], data[7]]) as usize;
        (field_index < 0 && pair_count > 0).then_some(pair_count)
    }

    fn parse_sxpair(data: &[u8]) -> Option<(usize, usize)> {
        if data.len() < 4 {
            return None;
        }
        Some((
            u16::from_le_bytes([data[0], data[1]]) as usize,
            u16::from_le_bytes([data[2], data[3]]) as usize,
        ))
    }

    fn calculated_item_formula_from_pending(
        pending: &XlsPivotCalculatedItemPending,
        fields: &[XlsPivotCacheField],
    ) -> Option<XlsPivotCalculatedItemFormula> {
        let (field_index, _) = *pending.item_refs.first()?;
        let mut item_names = Vec::with_capacity(pending.item_refs.len());
        for (ref_field_index, item_index) in &pending.item_refs {
            if *ref_field_index != field_index {
                return None;
            }
            let item = fields
                .get(*ref_field_index)?
                .shared_items
                .get(*item_index)?;
            item_names.push(Self::pivot_value_formula_reference(item));
        }
        let sx_name_indexes = (0..item_names.len()).collect::<Vec<_>>();
        let formula = Self::decompile_pivot_formula_with(
            &pending.ptgs,
            &sx_name_indexes,
            &item_names,
            Self::pivot_formula_item_reference,
        )?;
        Some(XlsPivotCalculatedItemFormula {
            field_index,
            formula,
        })
    }

    fn pivot_value_formula_reference(value: &PivotValue) -> String {
        match value {
            PivotValue::String(text) => text.clone(),
            PivotValue::Number(value) => {
                if value.fract() == 0.0 {
                    format!("{value:.0}")
                } else {
                    value.to_string()
                }
            }
            PivotValue::Boolean(value) => {
                if *value {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            PivotValue::Error(error) => Self::format_formula_error(error.code()),
            PivotValue::Blank => String::new(),
        }
    }

    fn attach_pending_pivot_formula(
        current_field: &mut Option<XlsPivotCacheField>,
        pending_formula: &mut Option<XlsPivotFormulaPending>,
        fields: &[XlsPivotCacheField],
    ) {
        let ready = pending_formula.as_ref().is_some_and(|pending| {
            pending.sx_name_field_indexes.len() >= pending.expected_sx_names
        });
        if !ready {
            return;
        }

        let Some(pending) = pending_formula.take() else {
            return;
        };
        let Some(field) = current_field.as_mut() else {
            return;
        };

        let mut field_names = fields
            .iter()
            .map(|field| field.name.clone())
            .collect::<Vec<_>>();
        field_names.push(field.name.clone());
        if let Some(formula) = Self::decompile_pivot_formula(
            &pending.ptgs,
            &pending.sx_name_field_indexes,
            &field_names,
        ) {
            field.formula = Some(formula);
        }
    }

    fn attach_pending_pivot_grouping(
        current_field: &mut Option<XlsPivotCacheField>,
        pending_group_range: &mut Option<PendingPivotRangeGroup>,
        fields: &[XlsPivotCacheField],
    ) {
        let ready = pending_group_range
            .as_ref()
            .is_some_and(|pending| pending.numbers.len() >= 3);
        if !ready {
            return;
        }
        let Some(pending) = pending_group_range.take() else {
            return;
        };
        let Some(field) = current_field else {
            return;
        };
        if let Some(unit) = xls_date_group_unit_from_flags(pending.flags) {
            let field_name = field
                .group_base_field
                .and_then(|index| fields.get(index))
                .map(|field| field.name.clone())
                .unwrap_or_else(|| field.name.clone());
            field.grouping = Some(PivotGrouping::Date {
                field: PivotFieldRef::new(field_name),
                units: vec![unit],
            });
            return;
        }
        if pending.flags & 0x001C != 0 {
            return;
        }
        let start_value = pending.numbers[0];
        let end_value = pending.numbers[1];
        let interval = pending.numbers[2];
        if !interval.is_finite() || interval <= 0.0 {
            return;
        }
        let start = if pending.flags & 0x0001 != 0 {
            None
        } else {
            start_value.is_finite().then_some(start_value)
        };
        let end = if pending.flags & 0x0002 != 0 {
            None
        } else {
            end_value.is_finite().then_some(end_value)
        };
        field.grouping = Some(PivotGrouping::Number {
            field: PivotFieldRef::new(field.name.clone()),
            start,
            end,
            interval,
        });
    }

    fn decompile_pivot_formula(
        ptgs: &[u8],
        sx_name_field_indexes: &[usize],
        field_names: &[String],
    ) -> Option<String> {
        Self::decompile_pivot_formula_with(
            ptgs,
            sx_name_field_indexes,
            field_names,
            Self::pivot_formula_field_reference,
        )
    }

    fn decompile_pivot_formula_with(
        ptgs: &[u8],
        sx_name_field_indexes: &[usize],
        field_names: &[String],
        name_formatter: fn(&str) -> String,
    ) -> Option<String> {
        let mut hooks = XlsPivotFormulaHooks {
            indexes: sx_name_field_indexes,
            names: field_names,
            formatter: name_formatter,
        };
        decompile_biff_pivot_formula(ptgs, &mut hooks).ok()
    }

    fn pivot_formula_field_reference(name: &str) -> String {
        if Self::is_simple_pivot_formula_name(name) {
            name.to_string()
        } else {
            format!("[{}]", name.replace(']', "]]"))
        }
    }

    fn pivot_formula_item_reference(name: &str) -> String {
        if Self::is_simple_pivot_formula_item_name(name) {
            name.to_string()
        } else {
            Self::pivot_formula_field_reference(name)
        }
    }

    fn is_simple_pivot_formula_item_name(name: &str) -> bool {
        let mut chars = name.chars();
        let Some(first) = chars.next() else {
            return false;
        };
        if !(first.is_ascii_alphabetic() || first == '_') {
            return false;
        }
        chars.all(|ch| ch.is_ascii_alphanumeric() || ch == '_' || ch == '.')
    }

    fn is_simple_pivot_formula_name(name: &str) -> bool {
        let mut chars = name.chars();
        let Some(first) = chars.next() else {
            return false;
        };
        if !(first.is_ascii_alphabetic() || first == '_') {
            return false;
        }
        if !chars.all(|ch| ch.is_ascii_alphanumeric() || ch == '_' || ch == '.') {
            return false;
        }
        CellAddress::parse(name).is_err()
    }

    fn format_formula_error(code: u8) -> String {
        match code {
            0x00 => "#NULL!",
            0x07 => "#DIV/0!",
            0x0F => "#VALUE!",
            0x17 => "#REF!",
            0x1D => "#NAME?",
            0x24 => "#NUM!",
            0x2A => "#N/A",
            _ => "#VALUE!",
        }
        .to_string()
    }

    fn parse_dconref(data: &[u8]) -> Option<PivotSource> {
        if data.len() < 10 {
            return None;
        }
        let start_row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let end_row = u16::from_le_bytes([data[2], data[3]]) as u32;
        let start_col = u16::from(data[4]);
        let end_col = u16::from(data[5]);
        let encoded_len = u16::from_le_bytes([data[6], data[7]]) as usize;
        if encoded_len == 0 {
            return None;
        }
        let flags = data[8];
        let source_marker = if flags & 0x01 == 0 {
            u16::from(*data.get(9)?)
        } else {
            let marker = data.get(9..11)?;
            u16::from_le_bytes([marker[0], marker[1]])
        };
        let sheet_len = encoded_len.saturating_sub(1);
        let source_name = if flags & 0x01 == 0 {
            let start = 10usize;
            let end = start.saturating_add(sheet_len);
            if end > data.len() {
                return None;
            }
            String::from_utf8_lossy(&data[start..end]).into_owned()
        } else {
            let start = 11usize;
            let byte_len = sheet_len.saturating_mul(2);
            let end = start.saturating_add(byte_len);
            if end > data.len() {
                return None;
            }
            let units = data[start..end]
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>();
            String::from_utf16_lossy(&units)
        };
        let range = CellRange::from_indices(start_row, start_col, end_row, end_col);

        if source_marker == 0x01 {
            let (target, sheet_name) = Self::parse_external_dconref_source(&source_name)
                .unwrap_or_else(|| (source_name.clone(), String::new()));
            return Some(PivotSource::Consolidation {
                ranges: vec![PivotSourceRange::new(sheet_name, range)
                    .with_external_relationship_target(target)],
            });
        }

        Some(PivotSource::range_on_sheet(source_name, range))
    }

    fn parse_external_dconref_source(source: &str) -> Option<(String, String)> {
        let open = source.find('[')?;
        let close = source[open + 1..].find(']')? + open + 1;
        let target = source[open + 1..close].to_string();
        let sheet = source[close + 1..].to_string();
        if target.is_empty() || sheet.is_empty() {
            return None;
        }
        Some((target, sheet))
    }

    fn parse_sxtbpg(data: &[u8]) -> Vec<u16> {
        data.chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect()
    }

    fn parse_dconname(data: &[u8]) -> Option<PivotSource> {
        let mut offset = 0usize;
        let name = read_unicode_string(data, &mut offset).ok()?;
        (!name.is_empty()).then(|| PivotSource::table(name))
    }

    fn read_xlunicode_no_cch(
        data: &[u8],
        offset: &mut usize,
        char_count: u16,
    ) -> XlsResult<String> {
        let flags = read_u8(data, offset)?;
        read_character_data(data, offset, char_count, flags)
    }

    fn push_axis_field(fields: &mut Vec<PivotField>, field: PivotField) {
        if fields
            .last()
            .is_some_and(|last| last.field.name.eq_ignore_ascii_case(&field.field.name))
        {
            return;
        }
        fields.push(field);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn wrapped_page_fields_use_rendered_row_count_for_target() {
        let mut layout = duke_sheets_core::PivotLayout::default();
        layout.page_wrap = 2;
        layout.page_over_then_down = true;

        let target = pivot_target_from_body(CellAddress::new(3, 5), &layout, 2);

        assert_eq!(target, CellAddress::new(1, 5));
    }
}
