use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{BufReader, Read, Seek};
use std::sync::Arc;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::{resolve_rel_path, SheetRel};
use crate::biff12::RecordIter;
use crate::biff12::{parser, records};
use crate::error::{XlsbError, XlsbResult};
use duke_sheets_core::style::NumberFormat;
use duke_sheets_core::{
    CellAddress, CellError, CellRange, PivotAggregate, PivotCacheInfo, PivotCacheSourceKind,
    PivotCalculatedField, PivotCalculatedItem, PivotDateGroupUnit, PivotDatePeriod, PivotField,
    PivotFieldRef, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayoutKind,
    PivotManualGroup, PivotMeasure, PivotRefreshStatus, PivotShowAs, PivotSort, PivotSource,
    PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable, PivotValue, PivotValuesAxis,
    WorkbookConnection, WorkbookConnectionKind,
};
use duke_sheets_formula::decompile::{
    decompile_pivot_formula as decompile_biff_pivot_formula, PivotFormulaHooks,
    PivotVariableArgCount,
};
use ssfmt::{date_serial::date_to_serial, DateSystem};

#[derive(Debug, Clone)]
struct PivotCacheDefinition {
    cache_id: u32,
    source: PivotSource,
    source_kind: PivotCacheSourceKind,
    fields: Vec<PivotCacheField>,
    calculated_items: Vec<PivotCalculatedItem>,
    groupings: Vec<PivotGrouping>,
    record_count: Option<u64>,
    refresh_on_load: bool,
    background_query: bool,
    missing_items_limit: Option<u32>,
}

#[derive(Debug, Clone)]
struct PivotCacheField {
    name: String,
    formula: Option<String>,
    formula_tokens: Option<Vec<u8>>,
    pname_field_indexes: Vec<usize>,
    grouping: Option<PivotGrouping>,
    group_parent: Option<usize>,
    group_base: Option<usize>,
    shared_items: Vec<PivotValue>,
    calculated_item_indexes: HashSet<usize>,
    group_items: Vec<PivotValue>,
    discrete_item_indexes: Vec<u32>,
    cache_record_value_mode: CacheRecordValueMode,
}

#[derive(Debug, Clone, Default)]
struct PendingCalculatedItem {
    tokens: Vec<u8>,
    target_field: Option<usize>,
    target_item: Option<usize>,
    pname_item_refs: Vec<(usize, usize)>,
}

#[derive(Debug, Clone, Default)]
struct PivotCacheRelationships {
    external_targets: HashMap<String, String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CacheRecordValueMode {
    Unknown,
    SharedItemIndex,
    Number,
    DateTime,
}

#[cfg(test)]
#[derive(Debug, Default)]
struct PivotCacheParseCounts {
    definitions: usize,
    records: usize,
}

pub(crate) struct PivotReadContext<'a> {
    date_system: DateSystem,
    connections: &'a HashMap<u32, WorkbookConnection>,
    caches: HashMap<String, Arc<PivotCacheDefinition>>,
    #[cfg(test)]
    parse_counts: PivotCacheParseCounts,
}

impl<'a> PivotReadContext<'a> {
    pub(crate) fn new(
        date_system: DateSystem,
        connections: &'a HashMap<u32, WorkbookConnection>,
    ) -> Self {
        Self {
            date_system,
            connections,
            caches: HashMap::new(),
            #[cfg(test)]
            parse_counts: PivotCacheParseCounts::default(),
        }
    }

    fn read_cache<R: Read + Seek>(
        &mut self,
        archive: &mut zip::ZipArchive<R>,
        path: &str,
    ) -> XlsbResult<Option<Arc<PivotCacheDefinition>>> {
        let key = canonical_package_path(path);
        if let Some(cache) = self.caches.get(&key) {
            return Ok(Some(cache.clone()));
        }

        let cache_id = trailing_number(path).unwrap_or(0);
        let cache = read_pivot_cache_definition(
            archive,
            cache_id,
            path,
            self.date_system,
            self.connections,
            #[cfg(test)]
            &mut self.parse_counts,
        )?
        .map(Arc::new);
        if let Some(cache) = &cache {
            self.caches.insert(key, cache.clone());
        }
        Ok(cache)
    }

    #[cfg(test)]
    pub(crate) fn cache_definition_parse_count(&self) -> usize {
        self.parse_counts.definitions
    }

    #[cfg(test)]
    pub(crate) fn cache_records_parse_count(&self) -> usize {
        self.parse_counts.records
    }
}

pub(crate) fn read_pivot_tables_for_sheet<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    sheet_path: &str,
    sheet_rels: &HashMap<String, SheetRel>,
    num_fmts: &HashMap<u32, String>,
    context: &mut PivotReadContext<'_>,
) -> XlsbResult<Vec<PivotTable>> {
    let mut pivot_paths = sheet_rels
        .values()
        .filter(|rel| rel.rel_type.ends_with("/pivotTable"))
        .map(|rel| resolve_rel_path(sheet_path, &rel.target))
        .collect::<Vec<_>>();
    pivot_paths.sort();

    let mut pivots = Vec::new();
    for pivot_path in pivot_paths {
        if let Some(pivot) = read_pivot_table(archive, &pivot_path, num_fmts, context)? {
            pivots.push(pivot);
        }
    }
    Ok(pivots)
}

fn read_pivot_table<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
    num_fmts: &HashMap<u32, String>,
    context: &mut PivotReadContext<'_>,
) -> XlsbResult<Option<PivotTable>> {
    let Some(cache_path) = related_part_path(archive, path, "/pivotCacheDefinition")? else {
        return Ok(None);
    };
    let Some(cache) = context.read_cache(archive, &cache_path)? else {
        return Ok(None);
    };

    let file = match archive.by_name(path) {
        Ok(file) => file,
        Err(_) => return Ok(None),
    };
    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(1024);

    let mut name = String::new();
    let mut target = CellAddress::new(0, 0);
    let mut rendered_range = None;
    let mut rows = Vec::new();
    let mut columns = Vec::new();
    let mut row_field_indexes = Vec::new();
    let mut column_field_indexes = Vec::new();
    let mut page_fields = Vec::new();
    let mut measures = Vec::new();
    let mut filters = Vec::new();
    let mut field_options = Vec::new();
    let mut current_sxvd_index = None;
    let mut current_data_field = None;
    let mut current_pivot_filter = None;
    let mut current_label_filter_criteria = Vec::new();
    let mut current_value_filter_criteria = Vec::new();
    let mut pending_top_n_filters = Vec::new();
    let mut pending_label_filters = Vec::new();
    let mut pending_value_filters = Vec::new();
    let mut pending_date_filters = Vec::new();
    let mut in_auto_sort_data_field_ref = false;
    let mut layout = duke_sheets_core::PivotLayout::default();
    let mut style = PivotStyle::default();
    let mut preserve_formatting = true;

    loop {
        let record = iter.next_record(&mut buf);
        let (typ, len) = match record {
            Ok(record) => record,
            Err(XlsbError::Parse(message)) if message.contains("unexpected end") => break,
            Err(err) => return Err(err),
        };
        let payload = &buf[..len];
        match typ {
            records::BRT_BEGIN_SXVIEW => {
                parse_begin_sxview(payload, &mut name, &mut layout, &mut preserve_formatting)?;
            }
            records::BRT_SX_LOCATION => {
                parse_sx_location(payload, &mut target, &mut rendered_range);
            }
            records::BRT_BEGIN_SXVD => {
                field_options.push(parse_pivot_field_options(payload)?);
                current_sxvd_index = Some(field_options.len() - 1);
            }
            records::BRT_BEGIN_SXVI if current_sxvd_index.is_some() => {
                if let Some(index) = current_sxvd_index {
                    if let Some(options) = field_options.get_mut(index) {
                        apply_pivot_field_item_options(options, payload);
                    }
                }
            }
            records::BRT_BEGIN_PIVOT_AREA_REF
                if current_sxvd_index.is_some()
                    && payload.len() >= 4
                    && parser::read_i32(payload, 0) == -2 =>
            {
                in_auto_sort_data_field_ref = true;
            }
            records::BRT_PIVOT_AREA_REF_ITEM
                if in_auto_sort_data_field_ref && payload.len() >= 4 =>
            {
                if let Some(index) = current_sxvd_index {
                    if let Some(options) = field_options.get_mut(index) {
                        options.sort_by_measure_index = Some(parser::read_u32(payload, 0) as usize);
                    }
                }
            }
            records::BRT_END_PIVOT_AREA_REF => {
                in_auto_sort_data_field_ref = false;
            }
            records::BRT_END_SXVD => {
                current_sxvd_index = None;
                in_auto_sort_data_field_ref = false;
            }
            records::BRT_BEGIN_ISXVD_RWS => {
                row_field_indexes = parse_axis_fields(
                    payload,
                    PivotValuesAxis::Rows,
                    &cache,
                    &field_options,
                    &mut rows,
                    &mut layout,
                );
            }
            records::BRT_BEGIN_ISXVD_COLS => {
                column_field_indexes = parse_axis_fields(
                    payload,
                    PivotValuesAxis::Columns,
                    &cache,
                    &field_options,
                    &mut columns,
                    &mut layout,
                );
            }
            records::BRT_BEGIN_SXPI => {
                parse_page_field(
                    payload,
                    &cache,
                    &field_options,
                    &mut page_fields,
                    &mut filters,
                );
            }
            records::BRT_BEGIN_SXDI => {
                if let Some(parsed) = parse_data_field(payload, &cache, num_fmts)? {
                    let index = measures.len();
                    current_data_field = Some(CurrentDataField {
                        index,
                        base_field_index: parsed.base_field_index,
                        base_item_index: parsed.base_item_index,
                    });
                    measures.push(parsed.measure);
                } else {
                    current_data_field = None;
                }
            }
            records::BRT_SXDI14 => {
                if let Some(current) = current_data_field {
                    if let Some(measure) = measures.get_mut(current.index) {
                        apply_data_field_14(
                            payload,
                            measure,
                            current.base_field_index,
                            current.base_item_index,
                            &cache,
                        )?;
                    }
                }
            }
            records::BRT_END_SXDI => {
                current_data_field = None;
            }
            records::BRT_SX_VIEW_STYLE => {
                style = parse_pivot_style(payload)?;
            }
            records::BRT_BEGIN_SX_FILTER => {
                current_pivot_filter = parse_begin_sx_filter(payload);
                current_label_filter_criteria.clear();
                current_value_filter_criteria.clear();
                if let Some(current) = &current_pivot_filter {
                    if let Some(filter) = parse_pivot_begin_label_filter(current, payload) {
                        pending_label_filters.push(filter);
                    }
                }
            }
            records::BRT_TOP10_FILTER => {
                if let Some(current) = current_pivot_filter {
                    if let Some(filter) = parse_pivot_top10_filter(current, payload) {
                        pending_top_n_filters.push(filter);
                    }
                }
            }
            records::BRT_CUSTOM_FILTER => {
                if let Some(current) = &current_pivot_filter {
                    if let Some(filter) = parse_pivot_label_filter(current, payload) {
                        pending_label_filters.push(filter);
                    } else if let Some(criteria) =
                        parse_pivot_label_filter_criteria(current, payload)
                    {
                        current_label_filter_criteria.push(criteria);
                    } else if let Some(criteria) =
                        parse_pivot_value_filter_criteria(current, payload)
                    {
                        current_value_filter_criteria.push(criteria);
                    }
                }
            }
            records::BRT_DYNAMIC_FILTER => {
                if let Some(current) = &current_pivot_filter {
                    if let Some(filter) = parse_pivot_date_period_filter(current, payload) {
                        pending_date_filters.push(filter);
                    }
                }
            }
            records::BRT_END_SX_FILTER => {
                if let Some(current) = &current_pivot_filter {
                    if let Some(filter) =
                        parse_pivot_label_between_filter(current, &current_label_filter_criteria)
                    {
                        pending_label_filters.push(filter);
                    }
                    if let Some(filter) =
                        parse_pivot_value_filter(current, &current_value_filter_criteria)
                    {
                        pending_value_filters.push(filter);
                    }
                    if let Some(filter) =
                        parse_pivot_date_filter(current, &current_value_filter_criteria)
                    {
                        pending_date_filters.push(filter);
                    }
                }
                current_pivot_filter = None;
                current_label_filter_criteria.clear();
                current_value_filter_criteria.clear();
            }
            records::BRT_END_SXVIEW => break,
            _ => {}
        }
    }

    if name.trim().is_empty() {
        name = format!("PivotTable{}", cache.cache_id);
    }
    apply_sort_by_measure_options(
        &mut rows,
        &mut columns,
        &mut page_fields,
        &cache,
        &field_options,
        &measures,
    );
    push_hidden_axis_field_item_filters(&mut filters, &cache, &field_options, &row_field_indexes);
    push_hidden_axis_field_item_filters(
        &mut filters,
        &cache,
        &field_options,
        &column_field_indexes,
    );
    let unplaced_field_indexes = (0..field_options.len())
        .filter(|&field_index| {
            let field_name = semantic_cache_field_name(&cache, field_index);
            !rows
                .iter()
                .chain(&columns)
                .chain(&page_fields)
                .any(|field| field.field.name.eq_ignore_ascii_case(&field_name))
        })
        .collect::<Vec<_>>();
    push_hidden_axis_field_item_filters(
        &mut filters,
        &cache,
        &field_options,
        &unplaced_field_indexes,
    );
    push_top_n_filters(&mut filters, &cache, &measures, &pending_top_n_filters);
    push_label_filters(&mut filters, &cache, &pending_label_filters);
    push_value_filters(&mut filters, &cache, &measures, &pending_value_filters);
    push_date_filters(&mut filters, &cache, &pending_date_filters);

    let mut pivot = PivotTable::new(0, name, cache.source.clone(), target);
    pivot.rows = rows;
    pivot.columns = columns;
    pivot.page_fields = page_fields;
    pivot.measures = measures;
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
    pivot.calculated_items = cache.calculated_items.clone();
    pivot.groupings = cache.groupings.clone();
    pivot.filters = filters;
    pivot.layout = layout;
    pivot.style = style;
    pivot.refresh_policy.refresh_on_open = cache.refresh_on_load;
    pivot.refresh_policy.preserve_formatting = preserve_formatting;
    pivot.refresh_policy.background_query = cache.background_query;
    pivot.refresh_policy.missing_items_limit = cache.missing_items_limit;
    pivot.rendered_range = rendered_range;
    pivot.set_cache_info(Some(PivotCacheInfo {
        cache_id: cache.cache_id,
        source_kind: cache.source_kind,
        record_count: cache.record_count,
        refreshed_version: None,
        refresh_status: PivotRefreshStatus::NotRefreshed,
    }));
    Ok(Some(pivot))
}

fn read_pivot_cache_definition<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    cache_id: u32,
    path: &str,
    date_system: DateSystem,
    connections: &HashMap<u32, WorkbookConnection>,
    #[cfg(test)] parse_counts: &mut PivotCacheParseCounts,
) -> XlsbResult<Option<PivotCacheDefinition>> {
    let relationships = read_pivot_cache_definition_relationships(archive, path)?;
    let file = match archive.by_name(path) {
        Ok(file) => file,
        Err(_) => return Ok(None),
    };
    #[cfg(test)]
    {
        parse_counts.definitions += 1;
    }
    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(1024);

    let mut source = None;
    let mut source_kind = PivotCacheSourceKind::Unknown;
    let mut fields = Vec::new();
    let mut current_field: Option<PivotCacheField> = None;
    let mut calculated_items = Vec::new();
    let mut current_calculated_item: Option<PendingCalculatedItem> = None;
    let mut record_count = None;
    let mut refresh_on_load = false;
    let mut background_query = false;
    let mut missing_items_limit = None;
    let mut in_shared_items = false;
    let mut in_group_items = false;
    let mut in_group_discrete = false;
    let mut consolidation_pages: Vec<Vec<String>> = Vec::new();
    let mut current_consolidation_page: Option<Vec<String>> = None;
    let mut consolidation_ranges: Vec<PivotSourceRange> = Vec::new();

    loop {
        let record = iter.next_record(&mut buf);
        let (typ, len) = match record {
            Ok(record) => record,
            Err(XlsbError::Parse(message)) if message.contains("unexpected end") => break,
            Err(err) => return Err(err),
        };
        let payload = &buf[..len];
        match typ {
            records::BRT_BEGIN_PIVOT_CACHE_DEF => {
                refresh_on_load = payload.get(3).is_some_and(|flags| flags & 0x04 != 0);
                background_query = payload.get(3).is_some_and(|flags| flags & 0x20 != 0);
                if payload.len() >= 8 {
                    let limit = parser::read_i32(payload, 4);
                    if limit >= 0 {
                        missing_items_limit = Some(limit as u32);
                    }
                }
                if payload.len() >= 21 {
                    record_count = Some(parser::read_u32(payload, 17) as u64);
                }
            }
            records::BRT_BEGIN_PCD_SOURCE => {
                if let Some((kind, parsed)) = parse_cache_source_header(payload, connections) {
                    source_kind = kind;
                    source = parsed;
                }
            }
            records::BRT_BEGIN_PCDS_SHEET => {
                if matches!(
                    source_kind,
                    PivotCacheSourceKind::Unknown | PivotCacheSourceKind::Worksheet
                ) {
                    if let Some(parsed) = parse_cache_sheet_source(payload)? {
                        source = Some(parsed);
                        source_kind = PivotCacheSourceKind::Worksheet;
                    }
                }
            }
            records::BRT_BEGIN_PCDSC_PAGE => {
                current_consolidation_page = Some(Vec::new());
            }
            records::BRT_BEGIN_PCDSC_PITEM => {
                if let Some(page) = &mut current_consolidation_page {
                    let (label, _) = parser::wide_str(payload, 0)?;
                    page.push(label);
                }
            }
            records::BRT_END_PCDSC_PAGE => {
                if let Some(page) = current_consolidation_page.take() {
                    consolidation_pages.push(page);
                }
            }
            records::BRT_BEGIN_PCDSC_SET => {
                if let Some(range) =
                    parse_consolidation_source_set(payload, &consolidation_pages, &relationships)?
                {
                    consolidation_ranges.push(range);
                    source_kind = PivotCacheSourceKind::Consolidation;
                }
            }
            records::BRT_BEGIN_PCD_FIELD => {
                current_field = Some(parse_cache_field(payload)?);
            }
            records::BRT_BEGIN_PCD_SHARED_ITEMS => {
                in_shared_items = true;
                if let Some(field) = &mut current_field {
                    field.cache_record_value_mode = parse_cache_record_value_mode(payload);
                }
            }
            records::BRT_END_PCD_SHARED_ITEMS => {
                in_shared_items = false;
            }
            records::BRT_BEGIN_PCDFG_ITEMS => {
                in_group_items = true;
            }
            records::BRT_END_PCDFG_ITEMS => {
                in_group_items = false;
            }
            records::BRT_BEGIN_PCDFG_DISCRETE => {
                in_group_discrete = true;
                if let Some(field) = &mut current_field {
                    field.discrete_item_indexes = parse_pivot_group_discrete(payload);
                }
            }
            records::BRT_END_PCDFG_DISCRETE => {
                in_group_discrete = false;
            }
            records::BRT_PCDI_INDEX => {
                if in_group_discrete && payload.len() >= 4 {
                    if let Some(field) = &mut current_field {
                        field
                            .discrete_item_indexes
                            .push(parser::read_u32(payload, 0));
                    }
                }
            }
            records::BRT_PCDI_MISSING
            | records::BRT_PCDI_BOOLEAN
            | records::BRT_PCDI_ERROR
            | records::BRT_PCDI_DATETIME
            | records::BRT_PCDI_NUMBER
            | records::BRT_PCDI_STRING
            | records::BRT_PCDIA_MISSING
            | records::BRT_PCDIA_BOOLEAN
            | records::BRT_PCDIA_ERROR
            | records::BRT_PCDIA_DATETIME
            | records::BRT_PCDIA_NUMBER
            | records::BRT_PCDIA_STRING => {
                if in_shared_items {
                    if let Some(field) = &mut current_field {
                        let index = field.shared_items.len();
                        field
                            .shared_items
                            .push(parse_shared_item(typ, payload, date_system)?);
                        if pcdia_formula_item(typ, payload) {
                            field.calculated_item_indexes.insert(index);
                        }
                    }
                } else if in_group_items {
                    if let Some(field) = &mut current_field {
                        field
                            .group_items
                            .push(parse_shared_item(typ, payload, date_system)?);
                    }
                }
            }
            records::BRT_BEGIN_PCDF_GROUP => {
                if let Some(field) = &mut current_field {
                    parse_pivot_field_group(payload, field);
                }
            }
            records::BRT_BEGIN_PCDFG_RANGE => {
                if let Some(field) = &mut current_field {
                    field.grouping = parse_pivot_group_range(&field.name, payload);
                }
            }
            records::BRT_BEGIN_PCD_CALC_ITEM => {
                current_calculated_item = parse_pcd_calculated_item(payload);
            }
            records::BRT_BEGIN_PR_FILTER => {
                if let Some(item) = &mut current_calculated_item {
                    if payload.len() >= 4 {
                        item.target_field = Some(parser::read_u32(payload, 0) as usize);
                    }
                }
            }
            records::BRT_BEGIN_PRF_ITEM => {
                if let Some(item) = &mut current_calculated_item {
                    if payload.len() >= 4 {
                        item.target_item = Some(parser::read_u32(payload, 0) as usize);
                    }
                }
            }
            records::BRT_BEGIN_PNPAIR => {
                if let Some(item) = &mut current_calculated_item {
                    if let Some(pair) = parse_pnpair(payload) {
                        item.pname_item_refs.push(pair);
                    }
                }
            }
            records::BRT_BEGIN_PNAME => {
                if let Some(field) = &mut current_field {
                    if let Some(field_index) = parse_pname(payload) {
                        field.pname_field_indexes.push(field_index);
                    }
                }
            }
            records::BRT_END_PCD_CALC_ITEM => {
                if let Some(item) = current_calculated_item.take() {
                    if let Some(item) = calculated_item_from_pending(item, &fields) {
                        calculated_items.push(item);
                    }
                }
            }
            records::BRT_END_PCD_FIELD => {
                if let Some(mut field) = current_field.take() {
                    attach_pivot_formula(&mut field, &fields);
                    fields.push(field);
                }
            }
            records::BRT_END_PIVOT_CACHE_DEF => break,
            _ => {}
        }
    }
    drop(iter);

    if let Some(records_path) = related_part_path(archive, path, "/pivotCacheRecords")? {
        read_pivot_cache_records(
            archive,
            &records_path,
            date_system,
            &mut fields,
            &mut record_count,
            #[cfg(test)]
            parse_counts,
        )?;
    }

    if matches!(source_kind, PivotCacheSourceKind::Consolidation)
        && !consolidation_ranges.is_empty()
    {
        source = Some(PivotSource::Consolidation {
            ranges: consolidation_ranges,
        });
    }

    let Some(source) = source else {
        return Ok(None);
    };
    let groupings = semantic_groupings_from_cache_fields(&fields);

    Ok(Some(PivotCacheDefinition {
        cache_id,
        source,
        source_kind,
        fields,
        calculated_items,
        groupings,
        record_count,
        refresh_on_load,
        background_query,
        missing_items_limit,
    }))
}

fn read_pivot_cache_records<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
    date_system: DateSystem,
    fields: &mut [PivotCacheField],
    record_count: &mut Option<u64>,
    #[cfg(test)] parse_counts: &mut PivotCacheParseCounts,
) -> XlsbResult<()> {
    let need_value_enrichment = fields.iter().any(cache_record_field_needs_enrichment);
    if !need_value_enrichment && record_count.is_some() {
        return Ok(());
    }

    let file = match archive.by_name(path) {
        Ok(file) => file,
        Err(_) => return Ok(()),
    };
    #[cfg(test)]
    {
        parse_counts.records += 1;
    }
    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(1024);
    let mut actual_count = 0u64;
    let record_field_indexes = if need_value_enrichment {
        fields
            .iter()
            .enumerate()
            .filter_map(|(index, field)| cache_record_field_has_value(field).then_some(index))
            .collect::<Vec<_>>()
    } else {
        Vec::new()
    };
    let mut seen_values = if need_value_enrichment {
        fields
            .iter()
            .map(|field| field.shared_items.iter().cloned().collect::<HashSet<_>>())
            .collect::<Vec<_>>()
    } else {
        Vec::new()
    };

    loop {
        let record = iter.next_record(&mut buf);
        let (typ, len) = match record {
            Ok(record) => record,
            Err(XlsbError::Parse(message)) if message.contains("unexpected end") => break,
            Err(err) => return Err(err),
        };
        let payload = &buf[..len];
        match typ {
            records::BRT_BEGIN_PCD_RECORDS if payload.len() >= 4 => {
                *record_count = Some(parser::read_u32(payload, 0) as u64);
            }
            records::BRT_PCD_RECORD => {
                actual_count += 1;
                if need_value_enrichment {
                    parse_cache_record_row(
                        payload,
                        date_system,
                        fields,
                        &record_field_indexes,
                        &mut seen_values,
                    )?;
                }
            }
            records::BRT_END_PCD_RECORDS => break,
            _ => {}
        }
    }

    if record_count.is_none() && actual_count > 0 {
        *record_count = Some(actual_count);
    }
    Ok(())
}

fn cache_record_field_needs_enrichment(field: &PivotCacheField) -> bool {
    field.shared_items.is_empty()
        && matches!(
            field.cache_record_value_mode,
            CacheRecordValueMode::Number | CacheRecordValueMode::DateTime
        )
}

fn parse_cache_record_row(
    payload: &[u8],
    date_system: DateSystem,
    fields: &mut [PivotCacheField],
    record_field_indexes: &[usize],
    seen_values: &mut [HashSet<PivotValue>],
) -> XlsbResult<()> {
    let mut offset = 0usize;

    for field_index in record_field_indexes.iter().copied() {
        let mode = cache_record_value_mode(&fields[field_index]);
        match mode {
            CacheRecordValueMode::SharedItemIndex => {
                if offset + 4 > payload.len() {
                    break;
                }
                offset += 4;
            }
            CacheRecordValueMode::Number => {
                if offset + 8 > payload.len() {
                    break;
                }
                let value = PivotValue::Number(parser::read_f64(payload, offset));
                intern_cache_record_value(fields, seen_values, field_index, value);
                offset += 8;
            }
            CacheRecordValueMode::DateTime => {
                if offset + 8 > payload.len() {
                    break;
                }
                let value = parse_shared_item(
                    records::BRT_PCDI_DATETIME,
                    &payload[offset..offset + 8],
                    date_system,
                )?;
                intern_cache_record_value(fields, seen_values, field_index, value);
                offset += 8;
            }
            CacheRecordValueMode::Unknown => break,
        }
    }

    Ok(())
}

fn cache_record_field_has_value(field: &PivotCacheField) -> bool {
    if field.formula.is_some() || field.formula_tokens.is_some() {
        return false;
    }
    !matches!(
        field.grouping,
        Some(PivotGrouping::Date { .. }) if field.group_base.is_some()
    )
}

fn cache_record_value_mode(field: &PivotCacheField) -> CacheRecordValueMode {
    match field.cache_record_value_mode {
        CacheRecordValueMode::SharedItemIndex => CacheRecordValueMode::SharedItemIndex,
        CacheRecordValueMode::Unknown if !field.shared_items.is_empty() => {
            CacheRecordValueMode::SharedItemIndex
        }
        mode => mode,
    }
}

fn intern_cache_record_value(
    fields: &mut [PivotCacheField],
    seen_values: &mut [HashSet<PivotValue>],
    field_index: usize,
    value: PivotValue,
) {
    let Some(seen) = seen_values.get_mut(field_index) else {
        return;
    };
    if seen.insert(value.clone()) {
        if let Some(field) = fields.get_mut(field_index) {
            field.shared_items.push(value);
        }
    }
}

fn parse_begin_sxview(
    payload: &[u8],
    name: &mut String,
    layout: &mut duke_sheets_core::PivotLayout,
    preserve_formatting: &mut bool,
) -> XlsbResult<()> {
    if payload.len() < 32 {
        return Ok(());
    }
    apply_sxview_layout_flags(payload, layout, preserve_formatting);
    *preserve_formatting = payload[4] & 0x80 != 0;
    layout.values_axis = match payload[12] {
        0x01 => PivotValuesAxis::Rows,
        0x02 => PivotValuesAxis::Columns,
        _ => layout.values_axis,
    };
    layout.page_wrap = payload[13] as u32;
    let values_position = parser::read_i32(payload, 16);
    if values_position >= 0 {
        layout.values_axis_position = Some(values_position as u32);
    }

    let mut offset = 32;
    let (view_name, consumed) = parser::wide_str(payload, offset)?;
    *name = view_name;
    offset += consumed;
    if sxview_flag(payload, 19) && offset < payload.len() {
        let (data_caption, consumed) = parser::wide_str(payload, offset)?;
        offset += consumed;
        if !data_caption.is_empty() {
            layout.data_caption = data_caption;
        }
    }
    if sxview_flag(payload, 20) && offset < payload.len() {
        let (caption, consumed) = parser::wide_str(payload, offset)?;
        offset += consumed;
        layout.grand_total_caption = Some(caption);
    }
    if !sxview_flag(payload, 38) && offset < payload.len() {
        let (caption, consumed) = parser::wide_str(payload, offset)?;
        offset += consumed;
        layout.error_caption = Some(caption);
    }
    if !sxview_flag(payload, 39) && offset < payload.len() {
        let (caption, _) = parser::wide_str(payload, offset)?;
        layout.missing_caption = Some(caption);
    }
    Ok(())
}

fn apply_sxview_layout_flags(
    payload: &[u8],
    layout: &mut duke_sheets_core::PivotLayout,
    preserve_formatting: &mut bool,
) {
    let primary = payload[1];
    layout.show_items = primary & (1 << 0) != 0;
    layout.edit_data = primary & (1 << 1) != 0;
    layout.disable_field_list = primary & (1 << 2) != 0;
    layout.show_calculated_members = primary & (1 << 4) == 0;
    layout.visual_totals = primary & (1 << 5) == 0;
    layout.show_multiple_label = primary & (1 << 6) != 0;

    let flags = parser::read_u16(payload, 2);
    layout.show_data_drop_down = flags & (1 << 0) == 0;
    layout.show_expand_collapse = flags & (1 << 4) == 0;
    layout.print_drill_indicators = flags & (1 << 5) != 0;
    layout.show_member_property_tips = flags & (1 << 6) != 0;
    layout.show_data_tips = flags & (1 << 7) == 0;
    layout.indent = ((flags >> 8) & 0x7F) as u32;
    layout.show_field_headers = flags & (1 << 15) == 0;

    layout.show_empty_rows = sxview_flag(payload, 2);
    layout.show_empty_columns = sxview_flag(payload, 3);
    layout.enable_wizard = sxview_flag(payload, 4);
    layout.enable_drill = sxview_flag(payload, 5);
    layout.enable_field_properties = sxview_flag(payload, 6);
    *preserve_formatting = sxview_flag(payload, 7);
    layout.show_error = sxview_flag(payload, 9);
    layout.show_missing = sxview_flag(payload, 10);
    layout.page_over_then_down = sxview_flag(payload, 11);
    layout.show_row_grand_totals = sxview_flag(payload, 13);
    layout.show_column_grand_totals = sxview_flag(payload, 14);
    layout.field_print_titles = sxview_flag(payload, 15);
    layout.item_print_titles = sxview_flag(payload, 17);
    layout.merge_item_labels = sxview_flag(payload, 18);
    layout.kind = if sxview_flag(payload, 32) || sxview_flag(payload, 35) {
        PivotLayoutKind::Compact
    } else if sxview_flag(payload, 33) || sxview_flag(payload, 34) {
        PivotLayoutKind::Outline
    } else {
        PivotLayoutKind::Tabular
    };
}

fn sxview_flag(payload: &[u8], bit: usize) -> bool {
    let byte = 4 + bit / 8;
    payload
        .get(byte)
        .is_some_and(|value| value & (1u8 << (bit % 8)) != 0)
}

fn parse_sx_location(
    payload: &[u8],
    target: &mut CellAddress,
    rendered_range: &mut Option<CellRange>,
) {
    if payload.len() < 16 {
        return;
    }
    let start_row = parser::read_u32(payload, 0);
    let end_row = parser::read_u32(payload, 4);
    let start_col = parser::read_u32(payload, 8).min(u16::MAX as u32) as u16;
    let end_col = parser::read_u32(payload, 12).min(u16::MAX as u32) as u16;
    let page_rows = if payload.len() >= 32 {
        parser::read_u32(payload, 28)
    } else {
        0
    };
    let target_row = if page_rows > 0 {
        start_row.saturating_sub(page_rows.saturating_add(1))
    } else {
        start_row
    };
    *target = CellAddress::new(target_row, start_col);
    *rendered_range = Some(CellRange::from_indices(
        start_row, start_col, end_row, end_col,
    ));
}

fn parse_axis_fields(
    payload: &[u8],
    axis: PivotValuesAxis,
    cache: &PivotCacheDefinition,
    field_options: &[PivotFieldOptions],
    fields: &mut Vec<PivotField>,
    layout: &mut duke_sheets_core::PivotLayout,
) -> Vec<usize> {
    let mut field_indexes = Vec::new();
    if payload.len() < 4 {
        return field_indexes;
    }
    let count = parser::read_u32(payload, 0) as usize;
    for offset in (4..payload.len()).step_by(4).take(count) {
        if offset + 4 > payload.len() {
            break;
        }
        let index = parser::read_i32(payload, offset);
        if index == -2 {
            layout.values_axis = axis;
            layout
                .values_axis_position
                .get_or_insert(fields.len() as u32);
        } else if index >= 0 {
            if let Some(field) = pivot_axis_field(
                cache,
                index as usize,
                field_options.get(index as usize).cloned(),
            ) {
                field_indexes.push(index as usize);
                push_axis_field(fields, field);
            }
        }
    }
    field_indexes
}

fn parse_page_field(
    payload: &[u8],
    cache: &PivotCacheDefinition,
    field_options: &[PivotFieldOptions],
    page_fields: &mut Vec<PivotField>,
    filters: &mut Vec<PivotFilter>,
) {
    if payload.len() < 8 {
        return;
    }
    let field_index = parser::read_u32(payload, 0) as usize;
    let selected_item = parser::read_u32(payload, 4);
    let Some(field) = cache.fields.get(field_index) else {
        return;
    };
    if let Some(field) =
        pivot_axis_field(cache, field_index, field_options.get(field_index).cloned())
    {
        push_axis_field(page_fields, field);
    }

    let filter_items = pivot_cache_filter_items(field);
    let filter_field_name = page_filter_field_name(cache, field_index);
    if let Some(item) = filter_items.get(selected_item as usize) {
        filters.push(PivotFilter::FieldItems {
            field: duke_sheets_core::PivotFieldRef::new(filter_field_name),
            allowed_items: vec![item.clone()],
        });
    } else if selected_item == 0x0010_00FE {
        if let Some(options) = field_options.get(field_index) {
            let hidden_items = &options.hidden_items;
            if hidden_items.is_empty() {
                return;
            }
            let allowed_items = filter_items
                .iter()
                .enumerate()
                .filter(|(index, _)| !hidden_items.contains(index))
                .map(|(_, item)| item.clone())
                .collect::<Vec<_>>();
            if !allowed_items.is_empty() && allowed_items.len() < filter_items.len() {
                filters.push(PivotFilter::FieldItems {
                    field: duke_sheets_core::PivotFieldRef::new(filter_field_name),
                    allowed_items,
                });
            }
        }
    }
}

fn page_filter_field_name(cache: &PivotCacheDefinition, field_index: usize) -> String {
    let Some(field) = cache.fields.get(field_index) else {
        return String::new();
    };
    if field.group_base.is_some() && matches!(field.grouping, Some(PivotGrouping::Date { .. })) {
        return field.name.clone();
    }
    semantic_cache_field_name(cache, field_index)
}

fn push_hidden_axis_field_item_filters(
    filters: &mut Vec<PivotFilter>,
    cache: &PivotCacheDefinition,
    field_options: &[PivotFieldOptions],
    field_indexes: &[usize],
) {
    for &field_index in field_indexes {
        let Some(field) = cache.fields.get(field_index) else {
            continue;
        };
        let Some(options) = field_options.get(field_index) else {
            continue;
        };
        let filter_items = pivot_cache_filter_items(field);
        if options.hidden_items.is_empty() || filter_items.is_empty() {
            continue;
        }

        let field_name = semantic_cache_field_name(cache, field_index);
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
        let allowed_items = filter_items
            .iter()
            .enumerate()
            .filter(|(index, _)| !hidden_items.contains(index))
            .map(|(_, item)| item.clone())
            .collect::<Vec<_>>();
        if !allowed_items.is_empty() && allowed_items.len() < filter_items.len() {
            filters.push(PivotFilter::FieldItems {
                field: duke_sheets_core::PivotFieldRef::new(field_name),
                allowed_items,
            });
        }
    }
}

fn pivot_cache_filter_items(field: &PivotCacheField) -> &[PivotValue] {
    if !field.group_items.is_empty() && (field.grouping.is_some() || field.shared_items.is_empty())
    {
        &field.group_items
    } else {
        &field.shared_items
    }
}

#[derive(Debug, Clone)]
struct PivotFieldOptions {
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
    hidden_items: Vec<usize>,
}

#[derive(Debug, Clone, Copy)]
struct CurrentPivotFilter {
    field_index: usize,
    measure_index: Option<usize>,
    filter_type: u32,
}

#[derive(Debug, Clone, Copy)]
struct PendingTopNFilter {
    field_index: usize,
    measure_index: usize,
    n: u32,
    top: bool,
    percent: bool,
}

#[derive(Debug, Clone)]
struct PendingLabelFilter {
    field_index: usize,
    kind: PendingLabelFilterKind,
}

#[derive(Debug, Clone)]
enum PendingLabelFilterKind {
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

#[derive(Debug, Clone)]
struct LabelFilterCriteria {
    operator_code: u8,
    value: String,
}

#[derive(Debug, Clone, Copy)]
struct PendingValueFilter {
    field_index: usize,
    measure_index: usize,
    kind: PendingValueFilterKind,
}

#[derive(Debug, Clone, Copy)]
struct PendingDateFilter {
    field_index: usize,
    kind: PendingDateFilterKind,
}

#[derive(Debug, Clone, Copy)]
enum PendingValueFilterKind {
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
enum PendingDateFilterKind {
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
struct ValueFilterCriteria {
    operator_code: u8,
    value: f64,
}

impl Default for PivotFieldOptions {
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
            hidden_items: Vec::new(),
        }
    }
}

fn parse_pivot_field_options(payload: &[u8]) -> XlsbResult<PivotFieldOptions> {
    if payload.len() < 20 {
        return Ok(PivotFieldOptions::default());
    }

    let subtotal_flags = parser::read_u16(payload, 1);
    let field_flags = payload[3];
    let behavior_flags = parser::read_u32(payload, 8);
    let auto_show_count = parser::read_u32(payload, 12);
    let mut offset = 20usize;
    let caption = if field_flags & 0x20 != 0 {
        let (value, consumed) = parser::wide_str(payload, offset)?;
        offset = offset.saturating_add(consumed);
        Some(value)
    } else {
        None
    };
    let subtotal_caption = if field_flags & 0x40 != 0 {
        let (value, _) = parser::wide_str(payload, offset)?;
        Some(value)
    } else {
        None
    };

    let subtotals = parse_pivot_subtotals(subtotal_flags);
    let subtotal = subtotals.first().copied().unwrap_or_else(|| {
        if subtotal_flags & 0x0001 != 0 {
            PivotSubtotal::Automatic
        } else {
            PivotSubtotal::None
        }
    });

    Ok(PivotFieldOptions {
        caption,
        sort: parse_pivot_sort(behavior_flags),
        sort_by_measure_index: None,
        subtotal,
        subtotal_caption,
        subtotals,
        show_empty_items: behavior_flags & (1 << 5) != 0,
        show_drop_downs: field_flags & 0x02 == 0,
        subtotal_top: behavior_flags & (1 << 8) != 0,
        insert_blank_row: behavior_flags & (1 << 7) != 0,
        insert_page_break: behavior_flags & (1 << 11) != 0,
        include_new_items_in_filter: behavior_flags & (1 << 18) == 0,
        item_page_count: if behavior_flags & (1 << 17) == 0 && auto_show_count != 0 {
            auto_show_count
        } else {
            10
        },
        hidden_items: Vec::new(),
    })
}

fn apply_pivot_field_item_options(options: &mut PivotFieldOptions, payload: &[u8]) {
    if payload.len() < 7 {
        return;
    }
    let flags = parser::read_u16(payload, 1);
    let item_index = parser::read_i32(payload, 3);
    if flags & 0x0001 != 0 && item_index >= 0 {
        options.hidden_items.push(item_index as usize);
    }
}

fn parse_pivot_sort(behavior_flags: u32) -> PivotSort {
    if behavior_flags & (1 << 12) == 0 {
        PivotSort::None
    } else if behavior_flags & (1 << 13) != 0 {
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

fn pivot_axis_field(
    cache: &PivotCacheDefinition,
    field_index: usize,
    options: Option<PivotFieldOptions>,
) -> Option<PivotField> {
    cache.fields.get(field_index)?;
    let options = options.unwrap_or_default();
    let mut field = PivotField::new(semantic_cache_field_name(cache, field_index));
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
    field.include_new_items_in_filter = options.include_new_items_in_filter;
    field.item_page_count = options.item_page_count;
    Some(field)
}

fn apply_sort_by_measure_options(
    rows: &mut [PivotField],
    columns: &mut [PivotField],
    page_fields: &mut [PivotField],
    cache: &PivotCacheDefinition,
    field_options: &[PivotFieldOptions],
    measures: &[PivotMeasure],
) {
    for (field_index, options) in field_options.iter().enumerate() {
        let Some(measure_index) = options.sort_by_measure_index else {
            continue;
        };
        let Some(measure) = measures.get(measure_index).cloned() else {
            continue;
        };
        let field_name = semantic_cache_field_name(cache, field_index);
        for field in rows
            .iter_mut()
            .chain(columns.iter_mut())
            .chain(page_fields.iter_mut())
            .filter(|field| field.field.name.eq_ignore_ascii_case(&field_name))
        {
            field.sort_by_measure = Some(measure.clone());
        }
    }
}

fn parse_begin_sx_filter(payload: &[u8]) -> Option<CurrentPivotFilter> {
    if payload.len() < 24 {
        return None;
    }
    let filter_type = parser::read_u32(payload, 8);
    if !matches!(filter_type, 1 | 2 | 4..=25) && !is_xlsb_date_filter_type(filter_type) {
        return None;
    }
    let measure_index = parser::read_u32(payload, 20);
    if matches!(filter_type, 1 | 2) && measure_index == u32::MAX {
        return None;
    }
    Some(CurrentPivotFilter {
        field_index: parser::read_u32(payload, 0) as usize,
        measure_index: (measure_index != u32::MAX).then_some(measure_index as usize),
        filter_type,
    })
}

fn parse_pivot_top10_filter(
    current: CurrentPivotFilter,
    payload: &[u8],
) -> Option<PendingTopNFilter> {
    if payload.len() < 17 {
        return None;
    }
    let value = parser::read_f64(payload, 1);
    if !value.is_finite() || value <= 0.0 || value > u32::MAX as f64 {
        return None;
    }
    let rounded = value.round();
    if (value - rounded).abs() > f64::EPSILON {
        return None;
    }
    let flags = payload[0];
    Some(PendingTopNFilter {
        field_index: current.field_index,
        measure_index: current.measure_index?,
        n: rounded as u32,
        top: flags & 0x01 != 0,
        percent: flags & 0x02 != 0 || current.filter_type == 2,
    })
}

fn parse_pivot_label_filter(
    current: &CurrentPivotFilter,
    payload: &[u8],
) -> Option<PendingLabelFilter> {
    if !matches!(current.filter_type, 5..=15) || payload.len() < 10 {
        return None;
    }
    let (raw_value, _) = parser::wide_str(payload, 10).ok()?;
    let expected_operator_code = match current.filter_type {
        14 => 1,
        5 | 7 | 9 | 11 => 5,
        6 | 8 | 10 => 2,
        15 => 3,
        12 => 4,
        13 => 6,
        _ => return None,
    };
    let expected_discriminator = match current.filter_type {
        5 | 12 | 13 | 14 | 15 => 1,
        6 | 7 | 8 | 9 | 10 | 11 => 0,
        _ => return None,
    };
    if payload[1] != expected_operator_code || payload[2] != expected_discriminator {
        return None;
    }
    let (operator, value) = match current.filter_type {
        5 => (PivotFilterOperator::NotEquals, raw_value),
        6 => (
            PivotFilterOperator::BeginsWith,
            raw_value.strip_suffix('*')?.to_string(),
        ),
        7 => (
            PivotFilterOperator::DoesNotBeginWith,
            raw_value.strip_suffix('*')?.to_string(),
        ),
        8 => (
            PivotFilterOperator::EndsWith,
            raw_value.strip_prefix('*')?.to_string(),
        ),
        9 => (
            PivotFilterOperator::DoesNotEndWith,
            raw_value.strip_prefix('*')?.to_string(),
        ),
        10 => (
            PivotFilterOperator::Contains,
            raw_value
                .strip_prefix('*')
                .and_then(|value| value.strip_suffix('*'))?
                .to_string(),
        ),
        11 => (
            PivotFilterOperator::DoesNotContain,
            raw_value
                .strip_prefix('*')
                .and_then(|value| value.strip_suffix('*'))?
                .to_string(),
        ),
        12 => (PivotFilterOperator::GreaterThan, raw_value),
        13 => (PivotFilterOperator::GreaterThanOrEqual, raw_value),
        14 => (PivotFilterOperator::LessThan, raw_value),
        15 => (PivotFilterOperator::LessThanOrEqual, raw_value),
        _ => return None,
    };
    if value.is_empty() {
        return None;
    }
    Some(PendingLabelFilter {
        field_index: current.field_index,
        kind: PendingLabelFilterKind::Comparison { operator, value },
    })
}

fn parse_pivot_begin_label_filter(
    current: &CurrentPivotFilter,
    payload: &[u8],
) -> Option<PendingLabelFilter> {
    if current.filter_type != 4 || payload.len() < 30 {
        return None;
    }
    let (value, _) = parser::wide_str(payload, 30).ok()?;
    if value.is_empty() {
        return None;
    }
    Some(PendingLabelFilter {
        field_index: current.field_index,
        kind: PendingLabelFilterKind::Comparison {
            operator: PivotFilterOperator::Equals,
            value,
        },
    })
}

fn parse_pivot_label_filter_criteria(
    current: &CurrentPivotFilter,
    payload: &[u8],
) -> Option<LabelFilterCriteria> {
    if !matches!(current.filter_type, 16 | 17) || payload.len() < 10 {
        return None;
    }
    if payload[0] != 6 || payload[2] != 1 {
        return None;
    }
    let (value, _) = parser::wide_str(payload, 10).ok()?;
    if value.is_empty() {
        return None;
    }
    Some(LabelFilterCriteria {
        operator_code: payload[1],
        value,
    })
}

fn parse_pivot_label_between_filter(
    current: &CurrentPivotFilter,
    criteria: &[LabelFilterCriteria],
) -> Option<PendingLabelFilter> {
    if !matches!(current.filter_type, 16 | 17) || criteria.len() < 2 {
        return None;
    }
    let expected = if current.filter_type == 16 {
        (6, 3, false)
    } else {
        (1, 4, true)
    };
    if criteria[0].operator_code != expected.0 || criteria[1].operator_code != expected.1 {
        return None;
    }
    Some(PendingLabelFilter {
        field_index: current.field_index,
        kind: PendingLabelFilterKind::Between {
            start: criteria[0].value.clone(),
            end: criteria[1].value.clone(),
            not_between: expected.2,
        },
    })
}

fn parse_pivot_value_filter_criteria(
    current: &CurrentPivotFilter,
    payload: &[u8],
) -> Option<ValueFilterCriteria> {
    if !(matches!(current.filter_type, 18..=25)
        || is_xlsb_custom_date_filter_type(current.filter_type))
        || payload.len() < 10
    {
        return None;
    }
    if payload[0] != 4 {
        return None;
    }
    let value = parser::read_f64(payload, 2);
    if !value.is_finite() {
        return None;
    }
    Some(ValueFilterCriteria {
        operator_code: payload[1],
        value,
    })
}

fn parse_pivot_date_filter(
    current: &CurrentPivotFilter,
    criteria: &[ValueFilterCriteria],
) -> Option<PendingDateFilter> {
    let kind = match current.filter_type {
        26 | 27 | 28 | 62 | 63 | 64 => {
            let criteria = criteria.first()?;
            let (expected_operator_code, operator) =
                xlsb_date_filter_operator_for_type(current.filter_type)?;
            if criteria.operator_code != expected_operator_code {
                return None;
            }
            PendingDateFilterKind::Comparison {
                operator,
                value: criteria.value,
            }
        }
        29 | 65 => {
            if criteria.len() < 2 {
                return None;
            }
            let expected = if current.filter_type == 29 {
                (6, 3, false)
            } else {
                (1, 4, true)
            };
            if criteria[0].operator_code != expected.0 || criteria[1].operator_code != expected.1 {
                return None;
            }
            PendingDateFilterKind::Between {
                start: criteria[0].value,
                end: criteria[1].value,
                not_between: expected.2,
            }
        }
        _ => return None,
    };
    Some(PendingDateFilter {
        field_index: current.field_index,
        kind,
    })
}

fn parse_pivot_date_period_filter(
    current: &CurrentPivotFilter,
    payload: &[u8],
) -> Option<PendingDateFilter> {
    if payload.len() < 4 {
        return None;
    }
    let cft = parser::read_u32(payload, 0);
    let period = xlsb_date_period_for_cft(cft)?;
    if current.filter_type != cft + 22 {
        return None;
    }
    Some(PendingDateFilter {
        field_index: current.field_index,
        kind: PendingDateFilterKind::Period(period),
    })
}

fn is_xlsb_date_filter_type(filter_type: u32) -> bool {
    is_xlsb_custom_date_filter_type(filter_type) || (30..=61).contains(&filter_type)
}

fn is_xlsb_custom_date_filter_type(filter_type: u32) -> bool {
    matches!(filter_type, 26 | 27 | 28 | 29 | 62 | 63 | 64 | 65)
}

fn xlsb_date_filter_operator_for_type(filter_type: u32) -> Option<(u8, PivotFilterOperator)> {
    Some(match filter_type {
        26 => (2, PivotFilterOperator::Equals),
        62 => (5, PivotFilterOperator::NotEquals),
        27 => (1, PivotFilterOperator::LessThan),
        63 => (3, PivotFilterOperator::LessThanOrEqual),
        28 => (4, PivotFilterOperator::GreaterThan),
        64 => (6, PivotFilterOperator::GreaterThanOrEqual),
        _ => return None,
    })
}

fn xlsb_date_period_for_cft(cft: u32) -> Option<PivotDatePeriod> {
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

fn parse_pivot_value_filter(
    current: &CurrentPivotFilter,
    criteria: &[ValueFilterCriteria],
) -> Option<PendingValueFilter> {
    let measure_index = current.measure_index?;
    let kind = match current.filter_type {
        18 | 19 | 20 | 21 | 22 | 23 => {
            let criteria = criteria.first()?;
            let (expected_operator_code, operator) = match current.filter_type {
                18 => (2, PivotFilterOperator::Equals),
                19 => (5, PivotFilterOperator::NotEquals),
                20 => (4, PivotFilterOperator::GreaterThan),
                21 => (6, PivotFilterOperator::GreaterThanOrEqual),
                22 => (1, PivotFilterOperator::LessThan),
                23 => (3, PivotFilterOperator::LessThanOrEqual),
                _ => return None,
            };
            if criteria.operator_code != expected_operator_code {
                return None;
            }
            PendingValueFilterKind::Comparison {
                operator,
                value: criteria.value,
            }
        }
        24 | 25 => {
            if criteria.len() < 2 {
                return None;
            }
            let expected = if current.filter_type == 24 {
                (6, 3, false)
            } else {
                (1, 4, true)
            };
            if criteria[0].operator_code != expected.0 || criteria[1].operator_code != expected.1 {
                return None;
            }
            PendingValueFilterKind::Between {
                start: criteria[0].value,
                end: criteria[1].value,
                not_between: expected.2,
            }
        }
        _ => return None,
    };
    Some(PendingValueFilter {
        field_index: current.field_index,
        measure_index,
        kind,
    })
}

fn push_top_n_filters(
    filters: &mut Vec<PivotFilter>,
    cache: &PivotCacheDefinition,
    measures: &[PivotMeasure],
    pending: &[PendingTopNFilter],
) {
    for filter in pending {
        let Some(measure) = measures.get(filter.measure_index).cloned() else {
            continue;
        };
        if cache.fields.get(filter.field_index).is_none() {
            continue;
        }
        filters.push(PivotFilter::TopN {
            field: PivotFieldRef::new(semantic_cache_field_name(cache, filter.field_index)),
            measure,
            n: filter.n,
            top: filter.top,
            percent: filter.percent,
        });
    }
}

fn push_label_filters(
    filters: &mut Vec<PivotFilter>,
    cache: &PivotCacheDefinition,
    pending: &[PendingLabelFilter],
) {
    for filter in pending {
        if cache.fields.get(filter.field_index).is_none() {
            continue;
        }
        let field = PivotFieldRef::new(semantic_cache_field_name(cache, filter.field_index));
        match &filter.kind {
            PendingLabelFilterKind::Comparison { operator, value } => {
                filters.push(PivotFilter::Label {
                    field,
                    operator: *operator,
                    value: value.clone(),
                });
            }
            PendingLabelFilterKind::Between {
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

fn push_value_filters(
    filters: &mut Vec<PivotFilter>,
    cache: &PivotCacheDefinition,
    measures: &[PivotMeasure],
    pending: &[PendingValueFilter],
) {
    for filter in pending {
        if cache.fields.get(filter.field_index).is_none() {
            continue;
        }
        let Some(measure) = measures.get(filter.measure_index).cloned() else {
            continue;
        };
        let field = PivotFieldRef::new(semantic_cache_field_name(cache, filter.field_index));
        match filter.kind {
            PendingValueFilterKind::Comparison { operator, value } => {
                filters.push(PivotFilter::Value {
                    field,
                    measure,
                    operator,
                    value,
                });
            }
            PendingValueFilterKind::Between {
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
    cache: &PivotCacheDefinition,
    pending: &[PendingDateFilter],
) {
    for filter in pending {
        if cache.fields.get(filter.field_index).is_none() {
            continue;
        }
        let field = PivotFieldRef::new(semantic_cache_field_name(cache, filter.field_index));
        match filter.kind {
            PendingDateFilterKind::Comparison { operator, value } => {
                filters.push(PivotFilter::Date {
                    field,
                    operator,
                    value,
                });
            }
            PendingDateFilterKind::Between {
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
            PendingDateFilterKind::Period(period) => {
                filters.push(PivotFilter::DatePeriod { field, period });
            }
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct CurrentDataField {
    index: usize,
    base_field_index: usize,
    base_item_index: usize,
}

#[derive(Debug, Clone)]
struct ParsedDataField {
    measure: PivotMeasure,
    base_field_index: usize,
    base_item_index: usize,
}

fn parse_data_field(
    payload: &[u8],
    cache: &PivotCacheDefinition,
    num_fmts: &HashMap<u32, String>,
) -> XlsbResult<Option<ParsedDataField>> {
    if payload.len() < 25 {
        return Ok(None);
    }
    let field_index = parser::read_u32(payload, 0) as usize;
    let Some(field) = cache.fields.get(field_index) else {
        return Ok(None);
    };
    let aggregate = parse_aggregate(parser::read_u32(payload, 4));
    let show_as = parser::read_u32(payload, 8);
    let base_field = parser::read_u32(payload, 12) as usize;
    let base_item = parser::read_u32(payload, 16) as usize;
    let num_format = parser::read_u32(payload, 20);
    let caption = if payload[24] != 0 {
        parser::wide_str(payload, 25)?.0
    } else {
        String::new()
    };
    let mut measure = PivotMeasure::new(field.name.clone(), aggregate);
    if !caption.is_empty() {
        measure.name = Some(caption);
    }
    measure.show_as = parse_data_field_show_as(show_as, base_field, base_item, cache)
        .unwrap_or(PivotShowAs::Normal);
    measure.number_format = pivot_number_format_code(num_format, num_fmts);
    Ok(Some(ParsedDataField {
        measure,
        base_field_index: base_field,
        base_item_index: base_item,
    }))
}

fn apply_data_field_14(
    payload: &[u8],
    measure: &mut PivotMeasure,
    base_field_index: usize,
    base_item_index: usize,
    cache: &PivotCacheDefinition,
) -> XlsbResult<()> {
    if payload.len() < 13 {
        return Ok(());
    }
    let show_as = parser::read_u32(payload, 4);
    if let Some(show_as) =
        parse_data_field_show_as_14(show_as, base_field_index, base_item_index, cache)
    {
        measure.show_as = show_as;
    }
    Ok(())
}

fn parse_data_field_show_as(
    code: u32,
    base_field_index: usize,
    base_item_index: usize,
    cache: &PivotCacheDefinition,
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

fn parse_data_field_show_as_14(
    code: u32,
    base_field_index: usize,
    _base_item_index: usize,
    cache: &PivotCacheDefinition,
) -> Option<PivotShowAs> {
    let base_field = || {
        cache
            .fields
            .get(base_field_index)
            .map(|field| PivotFieldRef::new(field.name.clone()))
    };
    Some(match code {
        9 => PivotShowAs::PercentOfParentRowTotal,
        10 => PivotShowAs::PercentOfParentColumnTotal,
        11 => PivotShowAs::PercentOfParentTotal {
            base_field: base_field()?,
        },
        13 => PivotShowAs::RankAscending {
            base_field: base_field()?,
        },
        14 => PivotShowAs::RankDescending {
            base_field: base_field()?,
        },
        _ => return None,
    })
}

fn pivot_number_format_code(num_fmt_id: u32, num_fmts: &HashMap<u32, String>) -> Option<String> {
    if num_fmt_id == NumberFormat::ID_GENERAL {
        return None;
    }
    if let Some(code) = num_fmts.get(&num_fmt_id) {
        return Some(code.clone());
    }
    let format = NumberFormat::BuiltIn(num_fmt_id);
    let code = format.format_string();
    (code != "General").then(|| code.to_string())
}

fn parse_pivot_style(payload: &[u8]) -> XlsbResult<PivotStyle> {
    if payload.len() < 6 {
        return Ok(PivotStyle::default());
    }
    let flags = parser::read_u16(payload, 0);
    let (name, _) = parser::wide_str(payload, 2)?;
    let mut style = PivotStyle::default();
    style.show_last_column = (flags & 0x02) != 0;
    style.show_row_stripes = (flags & 0x04) != 0;
    style.show_column_stripes = (flags & 0x08) != 0;
    style.show_row_headers = (flags & 0x10) != 0;
    style.show_column_headers = (flags & 0x20) != 0;
    if !name.is_empty() {
        style.name = Some(name);
    }
    Ok(style)
}

fn parse_cache_source_header(
    payload: &[u8],
    connections: &HashMap<u32, WorkbookConnection>,
) -> Option<(PivotCacheSourceKind, Option<PivotSource>)> {
    if payload.len() < 8 {
        return None;
    }
    let source_type = parser::read_u32(payload, 0);
    let connection_id = parser::read_u32(payload, 4);
    Some(match source_type {
        0 => (PivotCacheSourceKind::Worksheet, None),
        1 => (
            external_source_kind_for_connection(connection_id, connections),
            Some(external_source_for_connection(connection_id, connections)),
        ),
        2 => (
            PivotCacheSourceKind::Consolidation,
            Some(PivotSource::Consolidation { ranges: Vec::new() }),
        ),
        3 => (
            PivotCacheSourceKind::Scenario,
            Some(PivotSource::Scenario {
                name: String::new(),
            }),
        ),
        _ => (PivotCacheSourceKind::Unknown, None),
    })
}

fn external_source_kind_for_connection(
    connection_id: u32,
    connections: &HashMap<u32, WorkbookConnection>,
) -> PivotCacheSourceKind {
    match connections
        .get(&connection_id)
        .map(|connection| &connection.kind)
    {
        Some(WorkbookConnectionKind::Olap { .. }) => PivotCacheSourceKind::Olap,
        _ => PivotCacheSourceKind::External,
    }
}

fn external_source_for_connection(
    connection_id: u32,
    connections: &HashMap<u32, WorkbookConnection>,
) -> PivotSource {
    let Some(connection) = connections.get(&connection_id) else {
        return PivotSource::External {
            connection_name: connection_id.to_string(),
            command_text: None,
        };
    };
    match &connection.kind {
        WorkbookConnectionKind::Olap {
            command,
            command_type,
            ..
        } => PivotSource::Olap {
            connection_name: connection.name.clone(),
            cube: (*command_type == Some(1))
                .then(|| command.clone())
                .flatten(),
            command_text: command.clone(),
        },
        WorkbookConnectionKind::Database { command, .. } => PivotSource::External {
            connection_name: connection.name.clone(),
            command_text: command.clone(),
        },
        _ => PivotSource::External {
            connection_name: connection.name.clone(),
            command_text: None,
        },
    }
}

fn parse_cache_sheet_source(payload: &[u8]) -> XlsbResult<Option<PivotSource>> {
    if payload.len() < 3 {
        return Ok(None);
    }
    let source_is_name = payload[0] & 0x01 != 0;
    let load_external_rel = payload[2] & 0x01 != 0;
    let load_sheet = payload[2] & 0x02 != 0;
    if load_external_rel {
        return Ok(None);
    }

    let mut offset = 3;
    let sheet_name = if load_sheet {
        let (sheet_name, consumed) = parser::wide_str(payload, offset)?;
        offset += consumed;
        sheet_name
    } else {
        String::new()
    };

    if source_is_name {
        let (name, _) = parser::wide_str(payload, offset)?;
        return Ok((!name.is_empty()).then(|| PivotSource::table(name)));
    }

    let range_offset = offset;
    if range_offset + 16 > payload.len() {
        return Ok(None);
    }
    let start_row = parser::read_u32(payload, range_offset);
    let end_row = parser::read_u32(payload, range_offset + 4);
    let start_col = parser::read_u32(payload, range_offset + 8).min(u16::MAX as u32) as u16;
    let end_col = parser::read_u32(payload, range_offset + 12).min(u16::MAX as u32) as u16;
    let range = CellRange::from_indices(start_row, start_col, end_row, end_col);
    if sheet_name.is_empty() {
        Ok(Some(PivotSource::range(range)))
    } else {
        Ok(Some(PivotSource::range_on_sheet(sheet_name, range)))
    }
}

fn parse_consolidation_source_set(
    payload: &[u8],
    pages: &[Vec<String>],
    relationships: &PivotCacheRelationships,
) -> XlsbResult<Option<PivotSourceRange>> {
    if payload.len() < 19 {
        return Ok(None);
    }

    let item_indexes = [
        parser::read_u32(payload, 0),
        parser::read_u32(payload, 4),
        parser::read_u32(payload, 8),
        parser::read_u32(payload, 12),
    ];
    let source_is_name = payload[16] != 0;
    let load_external_rel = payload[18] & 0x01 != 0;
    let load_sheet = payload[18] & 0x02 != 0;

    let mut offset = 19;
    let sheet = if load_sheet {
        let (sheet, consumed) = parser::wide_str(payload, offset)?;
        offset += consumed;
        Some(sheet)
    } else {
        None
    };
    let external_relationship_id = if load_external_rel {
        let (relationship_id, consumed) = parser::wide_str(payload, offset)?;
        offset += consumed;
        Some(relationship_id)
    } else {
        None
    };
    let (name, range) = if source_is_name {
        let (name, _) = parser::wide_str(payload, offset)?;
        (Some(name), None)
    } else {
        (
            None,
            parse_unchecked_rfx(payload.get(offset..offset + 16).unwrap_or(&[])),
        )
    };
    if name.is_none() && range.is_none() && external_relationship_id.is_none() {
        return Ok(None);
    }

    let external_relationship_target = external_relationship_id
        .as_ref()
        .and_then(|id| relationships.external_targets.get(id).cloned());

    Ok(Some(PivotSourceRange {
        sheet,
        range,
        name,
        external_relationship_id,
        external_relationship_target,
        page_items: consolidation_page_items(&item_indexes, pages),
    }))
}

fn parse_unchecked_rfx(payload: &[u8]) -> Option<CellRange> {
    if payload.len() < 16 {
        return None;
    }

    let first_row = parser::read_i32(payload, 0);
    let last_row = parser::read_i32(payload, 4);
    let first_col = parser::read_i32(payload, 8);
    let last_col = parser::read_i32(payload, 12);
    if first_row < 0
        || last_row < first_row
        || first_col < 0
        || last_col < first_col
        || first_row > 1_048_575
        || last_row > 1_048_575
        || first_col > 16_383
        || last_col > 16_383
    {
        return None;
    }

    Some(CellRange::from_indices(
        first_row as u32,
        first_col as u16,
        last_row as u32,
        last_col as u16,
    ))
}

fn consolidation_page_items(item_indexes: &[u32; 4], pages: &[Vec<String>]) -> Vec<String> {
    item_indexes
        .iter()
        .copied()
        .zip(pages.iter())
        .filter_map(|(index, page)| {
            (index != u32::MAX)
                .then(|| page.get(index as usize).cloned())
                .flatten()
        })
        .collect()
}

fn parse_cache_field(payload: &[u8]) -> XlsbResult<PivotCacheField> {
    if payload.len() < 24 {
        return Ok(PivotCacheField {
            name: String::new(),
            formula: None,
            formula_tokens: None,
            pname_field_indexes: Vec::new(),
            grouping: None,
            group_parent: None,
            group_base: None,
            shared_items: Vec::new(),
            calculated_item_indexes: HashSet::new(),
            group_items: Vec::new(),
            discrete_item_indexes: Vec::new(),
            cache_record_value_mode: CacheRecordValueMode::Unknown,
        });
    }
    let flags = parser::read_u16(payload, 0);
    let (name, consumed) = parser::wide_str(payload, 20)?;
    let formula_offset = 20 + consumed;
    let formula_tokens = if flags & 0x0100 != 0 {
        parse_pivot_parsed_formula(payload, formula_offset)
    } else {
        None
    };
    Ok(PivotCacheField {
        name,
        formula: None,
        formula_tokens,
        pname_field_indexes: Vec::new(),
        grouping: None,
        group_parent: None,
        group_base: None,
        shared_items: Vec::new(),
        calculated_item_indexes: HashSet::new(),
        group_items: Vec::new(),
        discrete_item_indexes: Vec::new(),
        cache_record_value_mode: CacheRecordValueMode::Unknown,
    })
}

fn parse_cache_record_value_mode(payload: &[u8]) -> CacheRecordValueMode {
    if payload.len() < 6 {
        return CacheRecordValueMode::Unknown;
    }
    let flags = parser::read_u16(payload, 0);
    let stored_count = parser::read_u32(payload, 2);
    if stored_count > 0 {
        return CacheRecordValueMode::SharedItemIndex;
    }
    if flags & 0x0004 != 0 {
        return CacheRecordValueMode::DateTime;
    }
    if flags & 0x0040 != 0 {
        return CacheRecordValueMode::Number;
    }
    CacheRecordValueMode::Unknown
}

fn parse_pivot_group_discrete(payload: &[u8]) -> Vec<u32> {
    if payload.len() < 4 {
        return Vec::new();
    }
    let count = parser::read_u32(payload, 0) as usize;
    (0..count)
        .filter_map(|index| {
            let offset = 4 + index.saturating_mul(4);
            (offset + 4 <= payload.len()).then(|| parser::read_u32(payload, offset))
        })
        .collect()
}

fn parse_pivot_field_group(payload: &[u8], field: &mut PivotCacheField) {
    if payload.len() < 8 {
        return;
    }
    let parent = parser::read_i32(payload, 0);
    let base = parser::read_i32(payload, 4);
    if parent >= 0 {
        field.group_parent = Some(parent as usize);
    }
    if base >= 0 {
        field.group_base = Some(base as usize);
    }
}

fn semantic_groupings_from_cache_fields(fields: &[PivotCacheField]) -> Vec<PivotGrouping> {
    let mut groupings = Vec::new();
    let mut date_units_by_base: BTreeMap<usize, Vec<PivotDateGroupUnit>> = BTreeMap::new();

    for (index, field) in fields.iter().enumerate() {
        if let Some(grouping) = manual_grouping_from_cache_field(fields, field) {
            groupings.push(grouping);
            continue;
        }

        match &field.grouping {
            Some(PivotGrouping::Date { units, .. }) if field.group_base.is_some() => {
                if let (Some(base), Some(unit)) = (field.group_base, units.first().copied()) {
                    date_units_by_base.entry(base).or_default().push(unit);
                }
            }
            Some(PivotGrouping::Date { units, .. }) => {
                if let Some(unit) = units.first().copied() {
                    date_units_by_base.entry(index).or_default().push(unit);
                }
            }
            Some(grouping) => groupings.push(grouping.clone()),
            None => {}
        }
    }

    for (base, units) in date_units_by_base {
        if let Some(field) = fields.get(base) {
            groupings.push(PivotGrouping::Date {
                field: PivotFieldRef::new(field.name.clone()),
                units,
            });
        }
    }

    groupings
}

fn manual_grouping_from_cache_field(
    fields: &[PivotCacheField],
    field: &PivotCacheField,
) -> Option<PivotGrouping> {
    if field.discrete_item_indexes.is_empty() || field.group_items.is_empty() {
        return None;
    }
    let base_field = field
        .group_base
        .and_then(|base| fields.get(base))
        .unwrap_or(field);

    let mut members_by_group: BTreeMap<u32, Vec<PivotValue>> = BTreeMap::new();
    for (base_index, group_index) in field.discrete_item_indexes.iter().copied().enumerate() {
        let Some(member) = base_field.shared_items.get(base_index).cloned() else {
            continue;
        };
        members_by_group
            .entry(group_index)
            .or_default()
            .push(member);
    }

    let groups = members_by_group
        .into_iter()
        .filter_map(|(group_index, members)| {
            let group_value = field.group_items.get(group_index as usize)?;
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

fn parse_pivot_group_range(field_name: &str, payload: &[u8]) -> Option<PivotGrouping> {
    if payload.len() < 26 {
        return None;
    }
    let group_by = payload[0];
    let flags = payload[1];
    if group_by != 0x00 {
        let unit = xlsb_date_group_unit(group_by)?;
        return Some(PivotGrouping::Date {
            field: PivotFieldRef::new(field_name.to_string()),
            units: vec![unit],
        });
    }
    if flags & 0x04 != 0 {
        return None;
    }
    let start_value = parser::read_f64(payload, 2);
    let end_value = parser::read_f64(payload, 10);
    let interval = parser::read_f64(payload, 18);
    if !interval.is_finite() || interval <= 0.0 {
        return None;
    }
    let start = if flags & 0x01 != 0 {
        None
    } else {
        start_value.is_finite().then_some(start_value)
    };
    let end = if flags & 0x02 != 0 {
        None
    } else {
        end_value.is_finite().then_some(end_value)
    };
    Some(PivotGrouping::Number {
        field: PivotFieldRef::new(field_name.to_string()),
        start,
        end,
        interval,
    })
}

fn xlsb_date_group_unit(group_by: u8) -> Option<PivotDateGroupUnit> {
    Some(match group_by {
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

fn parse_pivot_parsed_formula(payload: &[u8], offset: usize) -> Option<Vec<u8>> {
    if offset + 4 > payload.len() {
        return None;
    }
    let token_len = parser::read_u32(payload, offset) as usize;
    let tokens_start = offset + 4;
    let tokens_end = tokens_start.checked_add(token_len)?;
    if tokens_end > payload.len() {
        return None;
    }
    Some(payload[tokens_start..tokens_end].to_vec())
}

fn parse_pcd_calculated_item(payload: &[u8]) -> Option<PendingCalculatedItem> {
    Some(PendingCalculatedItem {
        tokens: parse_pivot_parsed_formula(payload, 4)?,
        ..PendingCalculatedItem::default()
    })
}

fn parse_pname(payload: &[u8]) -> Option<usize> {
    if payload.len() < 6 {
        return None;
    }
    Some(parser::read_u32(payload, 0) as usize)
}

fn parse_pnpair(payload: &[u8]) -> Option<(usize, usize)> {
    if payload.len() < 9 {
        return None;
    }
    Some((
        parser::read_u32(payload, 1) as usize,
        parser::read_i32(payload, 5).try_into().ok()?,
    ))
}

fn calculated_item_from_pending(
    pending: PendingCalculatedItem,
    fields: &[PivotCacheField],
) -> Option<PivotCalculatedItem> {
    let field_index = pending.target_field?;
    let item_index = pending.target_item?;
    let field = fields.get(field_index)?;
    let item = field.shared_items.get(item_index)?.clone();
    let formula = decompile_pivot_item_formula(&pending.tokens, &pending.pname_item_refs, fields)?;
    Some(PivotCalculatedItem::new(field.name.clone(), item, formula))
}

fn attach_pivot_formula(field: &mut PivotCacheField, previous_fields: &[PivotCacheField]) {
    let Some(tokens) = field.formula_tokens.take() else {
        return;
    };
    let mut field_names = previous_fields
        .iter()
        .map(|field| field.name.clone())
        .collect::<Vec<_>>();
    field_names.push(field.name.clone());
    field.formula = decompile_pivot_formula(&tokens, &field.pname_field_indexes, &field_names);
}

fn decompile_pivot_formula(
    tokens: &[u8],
    pname_field_indexes: &[usize],
    field_names: &[String],
) -> Option<String> {
    decompile_pivot_formula_with(
        tokens,
        pname_field_indexes,
        field_names,
        pivot_formula_field_reference,
    )
}

fn decompile_pivot_formula_with(
    tokens: &[u8],
    pname_field_indexes: &[usize],
    field_names: &[String],
    name_formatter: fn(&str) -> String,
) -> Option<String> {
    let mut hooks = XlsbPivotFormulaHooks {
        indexes: pname_field_indexes,
        names: field_names,
        formatter: name_formatter,
    };
    decompile_biff_pivot_formula(tokens, &mut hooks).ok()
}

struct XlsbPivotFormulaHooks<'a> {
    indexes: &'a [usize],
    names: &'a [String],
    formatter: fn(&str) -> String,
}

impl PivotFormulaHooks for XlsbPivotFormulaHooks<'_> {
    fn resolve_name(&mut self, index: u32) -> Option<String> {
        let field_index = *self.indexes.get(index as usize)?;
        self.names
            .get(field_index)
            .map(|name| (self.formatter)(name))
    }

    fn variable_arg_count(&self) -> PivotVariableArgCount {
        PivotVariableArgCount::Biff12
    }
}

fn decompile_pivot_item_formula(
    tokens: &[u8],
    pname_item_refs: &[(usize, usize)],
    fields: &[PivotCacheField],
) -> Option<String> {
    let names = pname_item_refs
        .iter()
        .map(|(field_index, item_index)| {
            fields
                .get(*field_index)?
                .shared_items
                .get(*item_index)
                .map(pivot_value_formula_reference)
        })
        .collect::<Option<Vec<_>>>()?;
    let indexes = (0..names.len()).collect::<Vec<_>>();
    decompile_pivot_formula_with(tokens, &indexes, &names, pivot_formula_item_reference)
}

fn pivot_formula_field_reference(name: &str) -> String {
    if is_simple_pivot_formula_name(name) {
        name.to_string()
    } else {
        format!("[{}]", name.replace(']', "]]"))
    }
}

fn pivot_formula_item_reference(name: &str) -> String {
    if is_simple_pivot_formula_item_name(name) {
        name.to_string()
    } else {
        pivot_formula_field_reference(name)
    }
}

fn pivot_value_formula_reference(value: &PivotValue) -> String {
    match value {
        PivotValue::String(value) => value.clone(),
        PivotValue::Number(value) => format_formula_number(*value),
        PivotValue::Boolean(value) => {
            if *value {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        PivotValue::Error(error) => format_formula_error(error.code()),
        PivotValue::Blank => "\"\"".into(),
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

fn format_formula_number(value: f64) -> String {
    if value.is_finite() && value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
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

fn parse_shared_item(
    record_type: u16,
    payload: &[u8],
    date_system: DateSystem,
) -> XlsbResult<PivotValue> {
    Ok(match record_type {
        records::BRT_PCDI_MISSING | records::BRT_PCDIA_MISSING => PivotValue::Blank,
        records::BRT_PCDI_BOOLEAN | records::BRT_PCDIA_BOOLEAN => {
            PivotValue::Boolean(payload.first().copied().unwrap_or(0) != 0)
        }
        records::BRT_PCDI_ERROR | records::BRT_PCDIA_ERROR => {
            PivotValue::Error(parse_cell_error(payload.first().copied().unwrap_or(0x0F)))
        }
        records::BRT_PCDI_NUMBER | records::BRT_PCDIA_NUMBER if payload.len() >= 8 => {
            PivotValue::Number(parser::read_f64(payload, 0))
        }
        records::BRT_PCDI_DATETIME | records::BRT_PCDIA_DATETIME if payload.len() >= 8 => {
            let year = parser::read_u16(payload, 0) as i32;
            let month = parser::read_u16(payload, 2) as u32;
            let day = payload[4] as u32;
            let hour = payload[5] as f64;
            let minute = payload[6] as f64;
            let second = payload[7] as f64;
            let serial = date_to_serial(year, month, day, date_system)
                + ((hour * 3600.0) + (minute * 60.0) + second) / 86_400.0;
            PivotValue::Number(serial)
        }
        records::BRT_PCDI_STRING | records::BRT_PCDIA_STRING => {
            let (value, _) = parser::wide_str(payload, 0)?;
            PivotValue::String(value)
        }
        _ => PivotValue::Blank,
    })
}

fn pcdia_formula_item(record_type: u16, payload: &[u8]) -> bool {
    let Some(flags_offset) = pcdia_addl_info_offset(record_type, payload) else {
        return false;
    };
    flags_offset + 2 <= payload.len() && parser::read_u16(payload, flags_offset) & 0x0002 != 0
}

fn pcdia_addl_info_offset(record_type: u16, payload: &[u8]) -> Option<usize> {
    match record_type {
        records::BRT_PCDIA_MISSING => Some(0),
        records::BRT_PCDIA_BOOLEAN | records::BRT_PCDIA_ERROR => Some(1),
        records::BRT_PCDIA_NUMBER | records::BRT_PCDIA_DATETIME => Some(8),
        records::BRT_PCDIA_STRING => parser::wide_str(payload, 0).ok().map(|(_, len)| len),
        _ => None,
    }
}

fn parse_aggregate(code: u32) -> PivotAggregate {
    match code {
        0x01 => PivotAggregate::Count,
        0x02 => PivotAggregate::Average,
        0x03 => PivotAggregate::Max,
        0x04 => PivotAggregate::Min,
        0x05 => PivotAggregate::Product,
        0x06 => PivotAggregate::CountNumbers,
        0x07 => PivotAggregate::StdDev,
        0x08 => PivotAggregate::StdDevP,
        0x09 => PivotAggregate::Var,
        0x0A => PivotAggregate::VarP,
        _ => PivotAggregate::Sum,
    }
}

fn parse_cell_error(code: u8) -> CellError {
    match code {
        0x00 => CellError::Null,
        0x07 => CellError::Div0,
        0x0F => CellError::Value,
        0x17 => CellError::Ref,
        0x1D => CellError::Name,
        0x24 => CellError::Num,
        0x2A => CellError::Na,
        0x2B => CellError::GettingData,
        _ => CellError::Value,
    }
}

fn related_part_path<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    part_path: &str,
    relationship_suffix: &str,
) -> XlsbResult<Option<String>> {
    let rels_path = rels_path_for_part(part_path);
    let file = match archive.by_name(&rels_path) {
        Ok(file) => file,
        Err(_) => return Ok(None),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut target = String::new();
                let mut rel_type = String::new();
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Target" => target = String::from_utf8_lossy(&attr.value).into_owned(),
                        b"Type" => rel_type = String::from_utf8_lossy(&attr.value).into_owned(),
                        _ => {}
                    }
                }
                if rel_type.ends_with(relationship_suffix) && !target.is_empty() {
                    return Ok(Some(resolve_rel_path(part_path, &target)));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsbError::Parse(format!("XML error in .rels: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn read_pivot_cache_definition_relationships<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    cache_definition_path: &str,
) -> XlsbResult<PivotCacheRelationships> {
    let rels_path = rels_path_for_part(cache_definition_path);
    let file = match archive.by_name(&rels_path) {
        Ok(file) => file,
        Err(_) => return Ok(PivotCacheRelationships::default()),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut relationships = PivotCacheRelationships::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                let mut rel_type = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => {
                            id = attr.unescape_value().ok().map(|value| value.to_string());
                        }
                        b"Target" => {
                            target = attr.unescape_value().ok().map(|value| value.to_string());
                        }
                        b"Type" => {
                            rel_type = attr.unescape_value().ok().map(|value| value.to_string());
                        }
                        _ => {}
                    }
                }
                if let (Some(id), Some(target), Some(rel_type)) = (id, target, rel_type) {
                    if rel_type.ends_with("/externalLinkPath") && !target.is_empty() {
                        relationships.external_targets.insert(id, target);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsbError::Parse(format!("XML error in .rels: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(relationships)
}

fn rels_path_for_part(part_path: &str) -> String {
    let (base_dir, file_name) = part_path.rsplit_once('/').unwrap_or(("", part_path));
    if base_dir.is_empty() {
        format!("_rels/{file_name}.rels")
    } else {
        format!("{base_dir}/_rels/{file_name}.rels")
    }
}

fn canonical_package_path(path: &str) -> String {
    let normalized = path.replace('\\', "/");
    let mut parts = Vec::new();
    for part in normalized.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    parts.join("/").to_ascii_lowercase()
}

fn trailing_number(path: &str) -> Option<u32> {
    let stem = path.rsplit('/').next()?.split('.').next()?;
    let digits = stem
        .chars()
        .rev()
        .take_while(|ch| ch.is_ascii_digit())
        .collect::<String>();
    if digits.is_empty() {
        return None;
    }
    digits.chars().rev().collect::<String>().parse().ok()
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

fn semantic_cache_field_name(cache: &PivotCacheDefinition, field_index: usize) -> String {
    let Some(field) = cache.fields.get(field_index) else {
        return String::new();
    };
    field
        .group_base
        .and_then(|base| cache.fields.get(base))
        .map(|field| field.name.clone())
        .unwrap_or_else(|| field.name.clone())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn test_cache_field(name: &str, mode: CacheRecordValueMode) -> PivotCacheField {
        PivotCacheField {
            name: name.to_string(),
            formula: None,
            formula_tokens: None,
            pname_field_indexes: Vec::new(),
            grouping: None,
            group_parent: None,
            group_base: None,
            shared_items: Vec::new(),
            calculated_item_indexes: HashSet::new(),
            group_items: Vec::new(),
            discrete_item_indexes: Vec::new(),
            cache_record_value_mode: mode,
        }
    }

    #[test]
    fn parse_cache_source_header_keeps_non_range_source_kinds() {
        let empty_connections = HashMap::new();
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        assert_eq!(
            parse_cache_source_header(&payload, &empty_connections),
            Some((PivotCacheSourceKind::Worksheet, None))
        );

        let mut payload = Vec::new();
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&7u32.to_le_bytes());
        match parse_cache_source_header(&payload, &empty_connections) {
            Some((
                PivotCacheSourceKind::External,
                Some(PivotSource::External {
                    connection_name,
                    command_text: None,
                }),
            )) => assert_eq!(connection_name, "7"),
            other => panic!("unexpected external source: {other:?}"),
        }

        let mut payload = Vec::new();
        payload.extend_from_slice(&2u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        assert!(matches!(
            parse_cache_source_header(&payload, &empty_connections),
            Some((
                PivotCacheSourceKind::Consolidation,
                Some(PivotSource::Consolidation { .. })
            ))
        ));

        let mut payload = Vec::new();
        payload.extend_from_slice(&3u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        assert!(matches!(
            parse_cache_source_header(&payload, &empty_connections),
            Some((
                PivotCacheSourceKind::Scenario,
                Some(PivotSource::Scenario { .. })
            ))
        ));
    }

    #[test]
    fn parse_cache_source_header_resolves_connection_metadata() {
        let connections = HashMap::from([(
            7,
            WorkbookConnection::database(7, "SalesConnection", "Provider=Test;")
                .with_command("select * from Sales")
                .with_command_type(2),
        )]);
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&7u32.to_le_bytes());

        match parse_cache_source_header(&payload, &connections) {
            Some((
                PivotCacheSourceKind::External,
                Some(PivotSource::External {
                    connection_name,
                    command_text: Some(command_text),
                }),
            )) => {
                assert_eq!(connection_name, "SalesConnection");
                assert_eq!(command_text, "select * from Sales");
            }
            other => panic!("unexpected external source: {other:?}"),
        }
    }

    #[test]
    fn parse_cache_source_header_resolves_olap_connection_metadata() {
        let mut connection = WorkbookConnection::olap(10, "CubeSales");
        connection.kind = WorkbookConnectionKind::Olap {
            connection: Some("Provider=MSOLAP;Data Source=olapserver;".to_string()),
            command: Some("SalesCube".to_string()),
            command_type: Some(1),
            local: false,
            local_connection: None,
            local_refresh: true,
            send_locale: true,
            row_drill_count: Some(1000),
        };
        let connections = HashMap::from([(10, connection)]);
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&10u32.to_le_bytes());

        match parse_cache_source_header(&payload, &connections) {
            Some((
                PivotCacheSourceKind::Olap,
                Some(PivotSource::Olap {
                    connection_name,
                    cube: Some(cube),
                    command_text: Some(command_text),
                }),
            )) => {
                assert_eq!(connection_name, "CubeSales");
                assert_eq!(cube, "SalesCube");
                assert_eq!(command_text, "SalesCube");
            }
            other => panic!("unexpected OLAP source: {other:?}"),
        }
    }

    #[test]
    fn parse_consolidation_source_set_range_with_page_items() {
        let pages = vec![
            vec!["North".to_string(), "South".to_string()],
            vec!["FY24".to_string()],
        ];
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
        payload.push(0); // fName: range is specified by rfx.
        payload.push(0); // fBuiltIn.
        payload.push(0x02); // fLoadSheet.
        payload.extend_from_slice(&crate::biff12::encode_wide_str("Data"));
        payload.extend_from_slice(&1i32.to_le_bytes());
        payload.extend_from_slice(&4i32.to_le_bytes());
        payload.extend_from_slice(&1i32.to_le_bytes());
        payload.extend_from_slice(&2i32.to_le_bytes());

        let parsed =
            parse_consolidation_source_set(&payload, &pages, &PivotCacheRelationships::default())
                .expect("parse consolidation set")
                .expect("range source");

        assert_eq!(parsed.sheet.as_deref(), Some("Data"));
        assert_eq!(parsed.range, Some(CellRange::from_indices(1, 1, 4, 2)));
        assert_eq!(
            parsed.page_items,
            vec!["South".to_string(), "FY24".to_string()]
        );
    }

    #[test]
    fn parse_consolidation_source_set_preserves_named_source() {
        let mut payload = Vec::new();
        for _ in 0..4 {
            payload.extend_from_slice(&u32::MAX.to_le_bytes());
        }
        payload.push(1); // fName: source is a defined name, not an rfx range.
        payload.push(0);
        payload.push(0);
        payload.extend_from_slice(&crate::biff12::encode_wide_str("NamedSource"));

        let parsed =
            parse_consolidation_source_set(&payload, &[], &PivotCacheRelationships::default())
                .expect("parse named consolidation set")
                .expect("named source");
        assert_eq!(parsed.name.as_deref(), Some("NamedSource"));
        assert_eq!(parsed.sheet, None);
        assert_eq!(parsed.range, None);
    }

    #[test]
    fn parse_consolidation_source_set_preserves_external_relationship() {
        let mut payload = Vec::new();
        for _ in 0..4 {
            payload.extend_from_slice(&u32::MAX.to_le_bytes());
        }
        payload.push(0); // fName: range is specified by rfx.
        payload.push(0);
        payload.push(0x03); // fLoadRelId + fLoadSheet.
        payload.extend_from_slice(&crate::biff12::encode_wide_str("ExternalData"));
        payload.extend_from_slice(&crate::biff12::encode_wide_str("rIdExternal"));
        payload.extend_from_slice(&0i32.to_le_bytes());
        payload.extend_from_slice(&2i32.to_le_bytes());
        payload.extend_from_slice(&0i32.to_le_bytes());
        payload.extend_from_slice(&1i32.to_le_bytes());

        let mut relationships = PivotCacheRelationships::default();
        relationships.external_targets.insert(
            "rIdExternal".to_string(),
            "file:///C:/data/source.xlsx".to_string(),
        );

        let parsed = parse_consolidation_source_set(&payload, &[], &relationships)
            .expect("parse external consolidation set")
            .expect("external source");
        assert_eq!(parsed.sheet.as_deref(), Some("ExternalData"));
        assert_eq!(parsed.range, Some(CellRange::from_indices(0, 0, 2, 1)));
        assert_eq!(
            parsed.external_relationship_id.as_deref(),
            Some("rIdExternal")
        );
        assert_eq!(
            parsed.external_relationship_target.as_deref(),
            Some("file:///C:/data/source.xlsx")
        );
    }

    #[test]
    fn parse_cache_record_rows_interns_metadata_only_numeric_items() {
        let mut fields = vec![
            test_cache_field("Bucket", CacheRecordValueMode::Number),
            test_cache_field("Revenue", CacheRecordValueMode::Number),
        ];
        let record_field_indexes = vec![0, 1];
        let mut seen_values = fields
            .iter()
            .map(|field| field.shared_items.iter().cloned().collect::<HashSet<_>>())
            .collect::<Vec<_>>();

        let mut row = Vec::new();
        row.extend_from_slice(&1.0f64.to_le_bytes());
        row.extend_from_slice(&10.0f64.to_le_bytes());
        parse_cache_record_row(
            &row,
            DateSystem::Date1900,
            &mut fields,
            &record_field_indexes,
            &mut seen_values,
        )
        .expect("parse first row");

        let mut row = Vec::new();
        row.extend_from_slice(&2.0f64.to_le_bytes());
        row.extend_from_slice(&20.0f64.to_le_bytes());
        parse_cache_record_row(
            &row,
            DateSystem::Date1900,
            &mut fields,
            &record_field_indexes,
            &mut seen_values,
        )
        .expect("parse second row");

        let mut row = Vec::new();
        row.extend_from_slice(&1.0f64.to_le_bytes());
        row.extend_from_slice(&10.0f64.to_le_bytes());
        parse_cache_record_row(
            &row,
            DateSystem::Date1900,
            &mut fields,
            &record_field_indexes,
            &mut seen_values,
        )
        .expect("parse duplicate row");

        assert_eq!(
            fields[0].shared_items,
            vec![PivotValue::Number(1.0), PivotValue::Number(2.0)]
        );
        assert_eq!(
            fields[1].shared_items,
            vec![PivotValue::Number(10.0), PivotValue::Number(20.0)]
        );
    }

    #[test]
    fn parse_pivot_datetime_shared_item_uses_workbook_date_system() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1904u16.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.push(2);
        payload.push(12);
        payload.push(0);
        payload.push(0);

        let value = parse_shared_item(records::BRT_PCDI_DATETIME, &payload, DateSystem::Date1904)
            .expect("parse date1904 datetime");
        assert_eq!(value, PivotValue::Number(1.5));

        let value = parse_shared_item(records::BRT_PCDI_DATETIME, &payload, DateSystem::Date1900)
            .expect("parse date1900 datetime");
        match value {
            PivotValue::Number(serial) => assert!(serial > 1400.0, "{serial}"),
            other => panic!("unexpected pivot value: {other:?}"),
        }
    }
}
