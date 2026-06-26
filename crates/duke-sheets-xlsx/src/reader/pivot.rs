use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{BufRead, BufReader, Cursor, Read, Seek};

use chrono::{Datelike, NaiveDate};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use quick_xml::Writer;
use ssfmt::{date_serial::date_to_serial, DateSystem};

use super::archive_by_name;
use crate::error::{XlsxError, XlsxResult};
use duke_sheets_core::style::NumberFormat;
use duke_sheets_core::{
    CellAddress, CellError, CellRange, PivotAggregate, PivotCacheInfo, PivotCacheSourceKind,
    PivotCalculatedField, PivotCalculatedItem, PivotDateGroupUnit, PivotDatePeriod, PivotExtension,
    PivotField, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayout, PivotLayoutKind,
    PivotManualGroup, PivotMeasure, PivotRefreshStatus, PivotShowAs, PivotSort, PivotSource,
    PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable, PivotValue, PivotValuesAxis,
    WorkbookConnection, WorkbookConnectionKind,
};

#[derive(Debug, Clone)]
pub(super) struct PivotCacheDefinition {
    pub(super) source: PivotSource,
    pub(super) source_kind: PivotCacheSourceKind,
    pub(super) fields: Vec<PivotCacheField>,
    pub(super) calculated_fields: Vec<PivotCalculatedField>,
    pub(super) calculated_items: Vec<PivotCalculatedItem>,
    pub(super) groupings: Vec<PivotGrouping>,
    pub(super) record_count: Option<u64>,
    pub(super) refreshed_version: Option<String>,
    pub(super) refresh_on_load: bool,
    pub(super) background_query: bool,
    pub(super) missing_items_limit: Option<u32>,
}

#[derive(Debug, Clone)]
pub(super) struct PivotCacheField {
    pub(super) name: String,
    formula: Option<String>,
    pub(super) shared_items: Vec<PivotValue>,
    grouping: Option<PivotGrouping>,
    group_base: Option<usize>,
    group_parent: Option<usize>,
    discrete_item_indexes: Vec<u32>,
    group_items: Vec<PivotValue>,
}

#[derive(Debug, Clone)]
struct CurrentCalculatedItem {
    field_index: Option<usize>,
    formula: Option<String>,
    item_index: Option<u32>,
    reference_field_index: Option<usize>,
}

#[derive(Debug, Clone)]
struct CurrentPivotFieldItems {
    field_index: usize,
    hidden_items: Vec<u32>,
    collapsed_items: Vec<u32>,
}

pub(super) fn read_pivot_cache_definition<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    _cache_id: u32,
    path: &str,
    connections: &HashMap<u32, WorkbookConnection>,
) -> XlsxResult<Option<PivotCacheDefinition>> {
    let file = match archive_by_name(archive, path) {
        Ok(file) => file,
        Err(_) => return Ok(None),
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut source: Option<PivotSource> = None;
    let mut source_kind = PivotCacheSourceKind::Unknown;
    let mut connection_name = None;
    let mut connection_id = None;
    let mut fields = Vec::new();
    let mut calculated_items = Vec::new();
    let mut record_count = None;
    let mut refreshed_version = None;
    let mut refresh_on_load = false;
    let mut background_query = false;
    let mut missing_items_limit = None;
    let mut consolidation_ranges = Vec::new();
    let mut consolidation_pages = Vec::new();
    let mut current_consolidation_page: Option<Vec<String>> = None;
    let mut current_field: Option<PivotCacheField> = None;
    let mut current_calculated_item: Option<CurrentCalculatedItem> = None;
    let mut in_shared_items = false;
    let mut in_discrete_pr = false;
    let mut in_group_items = false;
    let mut in_consolidation = false;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"pivotCacheDefinition" => {
                    record_count = attr_u64(&e, b"recordCount");
                    refreshed_version = attr_string(&e, b"refreshedVersion");
                    refresh_on_load = attr_bool(&e, b"refreshOnLoad").unwrap_or(false);
                    background_query = attr_bool(&e, b"backgroundQuery").unwrap_or(false);
                    missing_items_limit = attr_u32(&e, b"missingItemsLimit");
                }
                b"cacheSource" => {
                    source_kind = parse_cache_source_kind(&e);
                    connection_name = attr_string(&e, b"connectionId");
                    connection_id = attr_u32(&e, b"connectionId");
                }
                b"worksheetSource" => {
                    source = parse_cache_worksheet_source(&e, source_kind)?;
                    if !matches!(source_kind, PivotCacheSourceKind::Scenario) {
                        source_kind = PivotCacheSourceKind::Worksheet;
                    }
                }
                b"consolidation" => in_consolidation = true,
                b"page" if in_consolidation => {
                    current_consolidation_page = Some(Vec::new());
                }
                b"rangeSet" if in_consolidation => {
                    if let Some(range) = parse_consolidation_range_set(&e, &consolidation_pages)? {
                        consolidation_ranges.push(range);
                    }
                }
                b"cacheField" => {
                    current_field = Some(PivotCacheField {
                        name: attr_string(&e, b"name").unwrap_or_default(),
                        formula: attr_string(&e, b"formula"),
                        shared_items: Vec::new(),
                        grouping: None,
                        group_base: None,
                        group_parent: None,
                        discrete_item_indexes: Vec::new(),
                        group_items: Vec::new(),
                    });
                }
                b"sharedItems" => in_shared_items = true,
                b"discretePr" => in_discrete_pr = true,
                b"groupItems" => in_group_items = true,
                b"calculatedItem" => {
                    current_calculated_item = Some(CurrentCalculatedItem {
                        field_index: attr_u32(&e, b"field").map(|value| value as usize),
                        formula: attr_string(&e, b"formula"),
                        item_index: None,
                        reference_field_index: None,
                    });
                }
                b"reference" => {
                    if let Some(item) = &mut current_calculated_item {
                        item.reference_field_index =
                            attr_u32(&e, b"field").map(|value| value as usize);
                    }
                }
                b"fieldGroup" => {
                    if let Some(field) = &mut current_field {
                        parse_field_group_attrs(field, &e);
                    }
                }
                b"rangePr" => {
                    if let Some(field) = &mut current_field {
                        field.grouping = parse_range_grouping(&field.name, &e);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"pivotCacheDefinition" => {
                    record_count = attr_u64(&e, b"recordCount");
                    refreshed_version = attr_string(&e, b"refreshedVersion");
                    refresh_on_load = attr_bool(&e, b"refreshOnLoad").unwrap_or(false);
                    background_query = attr_bool(&e, b"backgroundQuery").unwrap_or(false);
                    missing_items_limit = attr_u32(&e, b"missingItemsLimit");
                }
                b"cacheSource" => {
                    source_kind = parse_cache_source_kind(&e);
                    connection_name = attr_string(&e, b"connectionId");
                    connection_id = attr_u32(&e, b"connectionId");
                }
                b"worksheetSource" => {
                    source = parse_cache_worksheet_source(&e, source_kind)?;
                    if !matches!(source_kind, PivotCacheSourceKind::Scenario) {
                        source_kind = PivotCacheSourceKind::Worksheet;
                    }
                }
                b"rangeSet" if in_consolidation => {
                    if let Some(range) = parse_consolidation_range_set(&e, &consolidation_pages)? {
                        consolidation_ranges.push(range);
                    }
                }
                b"pageItem" if in_consolidation => {
                    if let Some(page) = &mut current_consolidation_page {
                        if let Some(name) = attr_string(&e, b"name") {
                            page.push(name);
                        }
                    }
                }
                b"cacheField" => fields.push(PivotCacheField {
                    name: attr_string(&e, b"name").unwrap_or_default(),
                    formula: attr_string(&e, b"formula"),
                    shared_items: Vec::new(),
                    grouping: None,
                    group_base: None,
                    group_parent: None,
                    discrete_item_indexes: Vec::new(),
                    group_items: Vec::new(),
                }),
                b"fieldGroup" => {
                    if let Some(field) = &mut current_field {
                        parse_field_group_attrs(field, &e);
                    }
                }
                b"rangePr" => {
                    if let Some(field) = &mut current_field {
                        field.grouping = parse_range_grouping(&field.name, &e);
                    }
                }
                b"m" | b"n" | b"b" | b"s" | b"e" if in_shared_items => {
                    if let Some(field) = &mut current_field {
                        field.shared_items.push(parse_shared_item(&e)?);
                    }
                }
                b"x" if in_discrete_pr => {
                    if let Some(field) = &mut current_field {
                        if let Some(index) = attr_u32(&e, b"v") {
                            field.discrete_item_indexes.push(index);
                        }
                    }
                }
                b"m" | b"n" | b"b" | b"s" | b"e" if in_group_items => {
                    if let Some(field) = &mut current_field {
                        field.group_items.push(parse_shared_item(&e)?);
                    }
                }
                b"x" if current_calculated_item.is_some() => {
                    if let Some(item) = &mut current_calculated_item {
                        if let Some(index) = attr_u32(&e, b"v") {
                            let reference_matches_field = match item.reference_field_index {
                                Some(field) => Some(field) == item.field_index,
                                None => true,
                            };
                            if item.item_index.is_none() || reference_matches_field {
                                item.item_index = Some(index);
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"cacheField" => {
                    if let Some(field) = current_field.take() {
                        fields.push(field);
                    }
                }
                b"page" if in_consolidation => {
                    consolidation_pages.push(current_consolidation_page.take().unwrap_or_default());
                }
                b"consolidation" => in_consolidation = false,
                b"sharedItems" => in_shared_items = false,
                b"discretePr" => in_discrete_pr = false,
                b"groupItems" => in_group_items = false,
                b"reference" => {
                    if let Some(item) = &mut current_calculated_item {
                        item.reference_field_index = None;
                    }
                }
                b"calculatedItem" => {
                    if let Some(item) = current_calculated_item.take() {
                        if let Some(item) = pivot_calculated_item_from_context(item, &fields) {
                            calculated_items.push(item);
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    if source.is_none()
        && matches!(source_kind, PivotCacheSourceKind::Consolidation)
        && !consolidation_ranges.is_empty()
    {
        source = Some(PivotSource::Consolidation {
            ranges: consolidation_ranges,
        });
    }
    let source = source.unwrap_or_else(|| {
        placeholder_source_for_kind(source_kind, connection_id, connection_name, connections)
    });
    let groupings = semantic_groupings_from_cache_fields(&fields);
    let calculated_fields = fields
        .iter()
        .filter_map(|field| {
            Some(PivotCalculatedField::new(
                field.name.clone(),
                field.formula.clone()?,
            ))
        })
        .collect();

    Ok(Some(PivotCacheDefinition {
        source,
        source_kind,
        fields,
        calculated_fields,
        calculated_items,
        groupings,
        record_count,
        refreshed_version,
        refresh_on_load,
        background_query,
        missing_items_limit,
    }))
}

fn pivot_calculated_item_from_context(
    item: CurrentCalculatedItem,
    fields: &[PivotCacheField],
) -> Option<PivotCalculatedItem> {
    let field_index = item.field_index?;
    let item_index = item.item_index?;
    let field = fields.get(field_index)?;
    let value = field.shared_items.get(item_index as usize)?;
    Some(PivotCalculatedItem::new(
        field.name.clone(),
        value.clone(),
        item.formula?,
    ))
}

fn parse_field_group_attrs(field: &mut PivotCacheField, e: &BytesStart<'_>) {
    field.group_base = attr_u32(e, b"base").map(|value| value as usize);
    field.group_parent = attr_u32(e, b"par").map(|value| value as usize);
}

fn semantic_groupings_from_cache_fields(fields: &[PivotCacheField]) -> Vec<PivotGrouping> {
    let mut groupings = Vec::new();
    let mut date_units_by_base: BTreeMap<usize, Vec<PivotDateGroupUnit>> = BTreeMap::new();

    for field in fields {
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
            Some(grouping) => groupings.push(grouping.clone()),
            None => {}
        }
    }

    for (base, units) in date_units_by_base {
        if let Some(field) = fields.get(base) {
            groupings.push(PivotGrouping::Date {
                field: field.name.clone().into(),
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
    let base = field.group_base?;
    let base_field = fields.get(base)?;
    if field.discrete_item_indexes.is_empty() || field.group_items.is_empty() {
        return None;
    }

    let mut members_by_group: BTreeMap<u32, Vec<PivotValue>> = BTreeMap::new();
    for (base_index, group_index) in field.discrete_item_indexes.iter().copied().enumerate() {
        let member = base_field.shared_items.get(base_index)?.clone();
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
        field: base_field.name.clone().into(),
        groups,
    })
}

fn parse_cache_source_kind(e: &BytesStart<'_>) -> PivotCacheSourceKind {
    match attr_string(e, b"type").as_deref() {
        Some("worksheet") => PivotCacheSourceKind::Worksheet,
        Some("external") => PivotCacheSourceKind::External,
        Some("consolidation") => PivotCacheSourceKind::Consolidation,
        Some("scenario") => PivotCacheSourceKind::Scenario,
        Some("olap") => PivotCacheSourceKind::Olap,
        _ => PivotCacheSourceKind::Unknown,
    }
}

fn placeholder_source_for_kind(
    kind: PivotCacheSourceKind,
    connection_id: Option<u32>,
    connection_name: Option<String>,
    connections: &HashMap<u32, WorkbookConnection>,
) -> PivotSource {
    let connection = connection_id.and_then(|id| connections.get(&id));
    let external_name = connection
        .map(|connection| connection.name.clone())
        .or(connection_name.clone())
        .unwrap_or_default();
    let command_text = connection.and_then(database_connection_command);
    match kind {
        PivotCacheSourceKind::Consolidation => PivotSource::Consolidation { ranges: Vec::new() },
        PivotCacheSourceKind::Scenario => PivotSource::Scenario {
            name: String::new(),
        },
        PivotCacheSourceKind::Olap => PivotSource::Olap {
            connection_name: external_name,
            cube: None,
            command_text,
        },
        PivotCacheSourceKind::External
        | PivotCacheSourceKind::Worksheet
        | PivotCacheSourceKind::Unknown => PivotSource::External {
            connection_name: external_name,
            command_text,
        },
    }
}

fn database_connection_command(connection: &WorkbookConnection) -> Option<String> {
    match &connection.kind {
        WorkbookConnectionKind::Database { command, .. } => command.clone(),
        _ => None,
    }
}

fn parse_cache_worksheet_source(
    e: &BytesStart<'_>,
    source_kind: PivotCacheSourceKind,
) -> XlsxResult<Option<PivotSource>> {
    if matches!(source_kind, PivotCacheSourceKind::Scenario) {
        return Ok(Some(PivotSource::Scenario {
            name: attr_string(e, b"name").unwrap_or_default(),
        }));
    }

    parse_worksheet_source(e)
}

fn parse_worksheet_source(e: &BytesStart<'_>) -> XlsxResult<Option<PivotSource>> {
    if let Some(name) = attr_string(e, b"name") {
        return Ok(Some(PivotSource::table(name)));
    }

    let Some(range_ref) = attr_string(e, b"ref") else {
        return Ok(None);
    };
    let range = CellRange::parse(&range_ref).map_err(|_| {
        XlsxError::InvalidFormat(format!("bad pivot worksheetSource ref: {range_ref}"))
    })?;

    Ok(Some(match attr_string(e, b"sheet") {
        Some(sheet) => PivotSource::range_on_sheet(sheet, range),
        None => PivotSource::range(range),
    }))
}

fn parse_consolidation_range_set(
    e: &BytesStart<'_>,
    pages: &[Vec<String>],
) -> XlsxResult<Option<PivotSourceRange>> {
    let Some(range_ref) = attr_string(e, b"ref") else {
        return Ok(None);
    };
    let range = CellRange::parse(&range_ref).map_err(|_| {
        XlsxError::InvalidFormat(format!("bad pivot consolidation rangeSet ref: {range_ref}"))
    })?;
    let mut source_range =
        PivotSourceRange::new(attr_string(e, b"sheet").unwrap_or_default(), range);
    source_range.name = attr_string(e, b"name");
    source_range.page_items = consolidation_range_page_items(e, pages);
    Ok(Some(source_range))
}

fn consolidation_range_page_items(e: &BytesStart<'_>, pages: &[Vec<String>]) -> Vec<String> {
    [
        b"i1".as_slice(),
        b"i2".as_slice(),
        b"i3".as_slice(),
        b"i4".as_slice(),
    ]
    .iter()
    .enumerate()
    .filter_map(|(page_index, attr)| {
        let item_index = attr_u32(e, attr)? as usize;
        pages
            .get(page_index)
            .and_then(|page| page.get(item_index))
            .cloned()
    })
    .collect()
}

fn parse_shared_item(e: &BytesStart<'_>) -> XlsxResult<PivotValue> {
    Ok(match e.name().local_name().as_ref() {
        b"m" => PivotValue::Blank,
        b"n" => PivotValue::Number(
            attr_string(e, b"v")
                .and_then(|value| value.parse().ok())
                .unwrap_or(0.0),
        ),
        b"b" => PivotValue::Boolean(attr_bool(e, b"v").unwrap_or(false)),
        b"s" => PivotValue::String(attr_string(e, b"v").unwrap_or_default()),
        b"e" => {
            let value = attr_string(e, b"v").unwrap_or_else(|| "#N/A".to_string());
            PivotValue::Error(CellError::parse(&value).unwrap_or(CellError::Na))
        }
        _ => PivotValue::Blank,
    })
}

fn parse_range_grouping(field_name: &str, e: &BytesStart<'_>) -> Option<PivotGrouping> {
    let group_by = attr_string(e, b"groupBy").unwrap_or_else(|| "range".to_string());
    if group_by.eq_ignore_ascii_case("range") {
        let start = match attr_bool(e, b"autoStart") {
            Some(true) => None,
            Some(false) | None => attr_f64(e, b"startNum"),
        };
        let end = match attr_bool(e, b"autoEnd") {
            Some(true) => None,
            Some(false) | None => attr_f64(e, b"endNum"),
        };
        let interval = attr_f64(e, b"groupInterval").unwrap_or(1.0);
        return Some(PivotGrouping::Number {
            field: field_name.to_string().into(),
            start,
            end,
            interval,
        });
    }

    let unit = parse_date_group_by(&group_by)?;
    Some(PivotGrouping::Date {
        field: field_name.to_string().into(),
        units: vec![unit],
    })
}

fn parse_date_group_by(value: &str) -> Option<PivotDateGroupUnit> {
    Some(match value {
        "seconds" => PivotDateGroupUnit::Seconds,
        "minutes" => PivotDateGroupUnit::Minutes,
        "hours" => PivotDateGroupUnit::Hours,
        "days" => PivotDateGroupUnit::Days,
        "months" => PivotDateGroupUnit::Months,
        "quarters" => PivotDateGroupUnit::Quarters,
        "years" => PivotDateGroupUnit::Years,
        _ => return None,
    })
}

pub(super) fn read_pivot_table<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
    caches: &HashMap<u32, PivotCacheDefinition>,
    num_fmts: &HashMap<u32, String>,
) -> XlsxResult<Option<PivotTable>> {
    let file = match archive_by_name(archive, path) {
        Ok(file) => file,
        Err(_) => return Ok(None),
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut name = String::new();
    let mut cache_id = 0u32;
    let mut target = CellAddress::new(0, 0);
    let mut rendered_range = None;
    let mut rows = Vec::new();
    let mut columns = Vec::new();
    let mut page_fields = Vec::new();
    let mut measures = Vec::new();
    let mut page_filters = Vec::new();
    let mut advanced_filters = Vec::new();
    let mut style = PivotStyle::default();
    let mut layout = PivotLayout::default();
    let mut preserve_formatting = true;
    let mut axis_context: Option<AxisContext> = None;
    let mut pivot_field_index = 0usize;
    let mut current_pivot_field: Option<CurrentPivotFieldItems> = None;
    let mut current_data_field: Option<CurrentDataField> = None;
    let mut current_pivot_filter: Option<CurrentPivotFilter> = None;
    let mut current_pivot_filter_depth = 0usize;
    let mut in_pivot_filters = false;
    let mut options_by_field: HashMap<usize, PivotFieldOptions> = HashMap::new();
    let mut hidden_items_by_field: HashMap<usize, Vec<u32>> = HashMap::new();
    let mut element_stack = Vec::new();
    let mut extensions = Vec::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let local_name = e.name().local_name();
                let local = local_name.as_ref();
                if local == b"ext" && is_pivot_table_extension_parent(&element_stack) {
                    let extension =
                        read_pivot_extension(&mut xml_reader, e.into_owned(), &mut buf)?;
                    extensions.push(extension);
                    buf.clear();
                    continue;
                }

                element_stack.push(pivot_xml_element(local));
                match local {
                    b"pivotTableDefinition" => {
                        parse_pivot_table_attrs(
                            &e,
                            &mut name,
                            &mut cache_id,
                            &mut preserve_formatting,
                            &mut layout,
                        );
                    }
                    b"location" => parse_location(&e, &mut target, &mut rendered_range)?,
                    b"pivotField" => {
                        options_by_field.insert(pivot_field_index, parse_pivot_field_options(&e));
                        current_pivot_field = Some(CurrentPivotFieldItems {
                            field_index: pivot_field_index,
                            hidden_items: Vec::new(),
                            collapsed_items: Vec::new(),
                        });
                        pivot_field_index += 1;
                    }
                    b"dataField" if attr_u32(&e, b"fld").is_some() => {
                        let Some(cache) = caches.get(&cache_id) else {
                            buf.clear();
                            continue;
                        };
                        if let Some(measure) = parse_data_field(&e, cache, num_fmts) {
                            let base_field_index = attr_i32(&e, b"baseField")
                                .and_then(|value| usize::try_from(value).ok());
                            measures.push(measure);
                            current_data_field = Some(CurrentDataField {
                                measure_index: measures.len() - 1,
                                base_field_index,
                            });
                        }
                    }
                    b"rowFields" => axis_context = Some(AxisContext::Rows),
                    b"colFields" => axis_context = Some(AxisContext::Columns),
                    b"pageFields" => axis_context = Some(AxisContext::Page),
                    b"dataFields" => axis_context = Some(AxisContext::Data),
                    b"filters" if current_pivot_filter.is_none() => in_pivot_filters = true,
                    b"filter" if in_pivot_filters && current_pivot_filter.is_none() => {
                        current_pivot_filter = parse_current_pivot_filter(&e);
                        current_pivot_filter_depth = usize::from(current_pivot_filter.is_some());
                    }
                    b"customFilter" => {
                        if let Some(filter) = &mut current_pivot_filter {
                            current_pivot_filter_depth += 1;
                            parse_pivot_custom_filter_attrs(filter, &e);
                        }
                    }
                    b"top10" => {
                        if let Some(filter) = &mut current_pivot_filter {
                            current_pivot_filter_depth += 1;
                            parse_pivot_top_filter_attrs(filter, &e);
                        }
                    }
                    _ if current_pivot_filter.is_some() => {
                        current_pivot_filter_depth += 1;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let local_name = e.name().local_name();
                let local = local_name.as_ref();
                if local == b"ext" && is_pivot_table_extension_parent(&element_stack) {
                    extensions.push(empty_pivot_extension(&e)?);
                    buf.clear();
                    continue;
                }

                match local {
                    b"pivotTableDefinition" => {
                        parse_pivot_table_attrs(
                            &e,
                            &mut name,
                            &mut cache_id,
                            &mut preserve_formatting,
                            &mut layout,
                        );
                    }
                    b"location" => parse_location(&e, &mut target, &mut rendered_range)?,
                    b"pivotField" => {
                        options_by_field.insert(pivot_field_index, parse_pivot_field_options(&e));
                        pivot_field_index += 1;
                    }
                    b"item" => {
                        if let Some(field_items) = &mut current_pivot_field {
                            if let Some(index) = attr_u32(&e, b"x") {
                                if attr_bool(&e, b"h").unwrap_or(false) {
                                    field_items.hidden_items.push(index);
                                }
                                if !attr_bool(&e, b"sd").unwrap_or(true) {
                                    field_items.collapsed_items.push(index);
                                }
                            }
                        }
                    }
                    b"field" => {
                        let Some(cache) = caches.get(&cache_id) else {
                            buf.clear();
                            continue;
                        };
                        if attr_i32(&e, b"x") == Some(-2) {
                            match axis_context {
                                Some(AxisContext::Rows) => {
                                    layout.values_axis = PivotValuesAxis::Rows;
                                    layout.values_axis_position.get_or_insert(rows.len() as u32);
                                }
                                Some(AxisContext::Columns) => {
                                    layout.values_axis = PivotValuesAxis::Columns;
                                    layout
                                        .values_axis_position
                                        .get_or_insert(columns.len() as u32);
                                }
                                _ => {}
                            }
                        } else if let Some(field_index) =
                            attr_i32(&e, b"x").and_then(|v| usize::try_from(v).ok())
                        {
                            if let Some(field) = pivot_axis_field(
                                cache,
                                field_index,
                                options_by_field.get(&field_index).cloned(),
                            ) {
                                match axis_context {
                                    Some(AxisContext::Rows) => push_axis_field(&mut rows, field),
                                    Some(AxisContext::Columns) => {
                                        push_axis_field(&mut columns, field)
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"pageField" => {
                        let Some(cache) = caches.get(&cache_id) else {
                            buf.clear();
                            continue;
                        };
                        if let Some(field_index) =
                            attr_i32(&e, b"fld").and_then(|v| usize::try_from(v).ok())
                        {
                            if let Some(field) = pivot_axis_field(
                                cache,
                                field_index,
                                options_by_field.get(&field_index).cloned(),
                            ) {
                                push_axis_field(&mut page_fields, field);
                                if let Some(item_index) = attr_u32(&e, b"item") {
                                    if let Some(cache_field) = cache.fields.get(field_index) {
                                        if let Some(item) =
                                            cache_field.shared_items.get(item_index as usize)
                                        {
                                            page_filters.push(PivotFilter::FieldItems {
                                                field: duke_sheets_core::PivotFieldRef::new(
                                                    semantic_cache_field_name(cache, field_index)
                                                        .unwrap_or_else(|| {
                                                            cache_field.name.clone()
                                                        }),
                                                ),
                                                allowed_items: vec![item.clone()],
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"dataField" => {
                        let Some(cache) = caches.get(&cache_id) else {
                            buf.clear();
                            continue;
                        };
                        if let Some(pivot_show_as) = attr_string(&e, b"pivotShowAs") {
                            if let Some(current) = current_data_field {
                                let source_field_index = attr_u32(&e, b"sourceField")
                                    .map(|value| value as usize)
                                    .or(current.base_field_index);
                                if let Some(show_as) = parse_x14_show_as(
                                    &pivot_show_as,
                                    cache,
                                    source_field_index,
                                ) {
                                    if let Some(measure) = measures.get_mut(current.measure_index) {
                                        measure.show_as = show_as;
                                    }
                                }
                            }
                        } else if let Some(measure) = parse_data_field(&e, cache, num_fmts) {
                            measures.push(measure);
                        }
                    }
                    b"filter" if in_pivot_filters && current_pivot_filter.is_none() => {
                        let Some(cache) = caches.get(&cache_id) else {
                            buf.clear();
                            continue;
                        };
                        if let Some(filter) = parse_current_pivot_filter(&e)
                            .and_then(|filter| pivot_filter_from_context(filter, cache, &measures))
                        {
                            advanced_filters.push(filter);
                        }
                    }
                    b"customFilter" => {
                        if let Some(filter) = &mut current_pivot_filter {
                            parse_pivot_custom_filter_attrs(filter, &e);
                        }
                    }
                    b"top10" => {
                        if let Some(filter) = &mut current_pivot_filter {
                            parse_pivot_top_filter_attrs(filter, &e);
                        }
                    }
                    b"pivotTableStyleInfo" => style = parse_pivot_style(&e),
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let local_name = e.name().local_name();
                let local = local_name.as_ref();
                match local {
                    b"pivotField" => {
                        if let Some(field_items) = current_pivot_field.take() {
                            if !field_items.hidden_items.is_empty() {
                                hidden_items_by_field
                                    .insert(field_items.field_index, field_items.hidden_items);
                            }
                            if !field_items.collapsed_items.is_empty() {
                                options_by_field
                                    .entry(field_items.field_index)
                                    .or_default()
                                    .collapsed_item_indexes = field_items.collapsed_items;
                            }
                        }
                    }
                    b"rowFields" | b"colFields" | b"pageFields" | b"dataFields" => {
                        axis_context = None
                    }
                    b"dataField" => {
                        if e.name().as_ref() == b"dataField" {
                            current_data_field = None;
                        }
                    }
                    b"filter"
                        if current_pivot_filter.is_some() && current_pivot_filter_depth == 1 =>
                    {
                        let Some(cache) = caches.get(&cache_id) else {
                            current_pivot_filter = None;
                            current_pivot_filter_depth = 0;
                            element_stack.pop();
                            buf.clear();
                            continue;
                        };
                        if let Some(filter) = current_pivot_filter
                            .take()
                            .and_then(|filter| pivot_filter_from_context(filter, cache, &measures))
                        {
                            advanced_filters.push(filter);
                        }
                        current_pivot_filter_depth = 0;
                    }
                    b"filters" if current_pivot_filter.is_none() => in_pivot_filters = false,
                    _ if current_pivot_filter.is_some() => {
                        current_pivot_filter_depth = current_pivot_filter_depth.saturating_sub(1);
                    }
                    _ => {}
                }
                element_stack.pop();
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    let Some(cache) = caches.get(&cache_id) else {
        return Ok(None);
    };

    let mut pivot = PivotTable::new(0, name, cache.source.clone(), target);
    pivot.rows = rows;
    pivot.columns = columns;
    pivot.page_fields = page_fields;
    pivot.measures = measures;
    pivot.calculated_fields = cache.calculated_fields.clone();
    pivot.layout = layout;
    pivot.refresh_policy.refresh_on_open = cache.refresh_on_load;
    pivot.refresh_policy.preserve_formatting = preserve_formatting;
    pivot.refresh_policy.background_query = cache.background_query;
    pivot.refresh_policy.missing_items_limit = cache.missing_items_limit;
    pivot.style = style;
    pivot.rendered_range = rendered_range;
    pivot.groupings = cache.groupings.clone();
    pivot.extensions = extensions;
    pivot.set_cache_info(Some(PivotCacheInfo {
        cache_id,
        source_kind: cache.source_kind,
        record_count: cache.record_count,
        refreshed_version: cache.refreshed_version.clone(),
        refresh_status: PivotRefreshStatus::NotRefreshed,
    }));

    let hidden_item_filters = hidden_items_by_field
        .into_iter()
        .filter_map(|(field_index, hidden)| {
            let field = cache.fields.get(field_index)?;
            if page_filters.iter().any(|filter| {
                matches!(
                    filter,
                    PivotFilter::FieldItems { field: filter_field, .. }
                        if filter_field.name.eq_ignore_ascii_case(&field.name)
                )
            }) {
                return None;
            }

            let hidden = hidden.into_iter().collect::<HashSet<_>>();
            let allowed_items = field
                .shared_items
                .iter()
                .enumerate()
                .filter(|(index, _)| !hidden.contains(&(*index as u32)))
                .map(|(_, value)| value.clone())
                .collect::<Vec<_>>();
            Some(PivotFilter::FieldItems {
                field: duke_sheets_core::PivotFieldRef::new(field.name.clone()),
                allowed_items,
            })
        })
        .collect::<Vec<_>>();
    pivot.filters = page_filters;
    pivot.filters.extend(hidden_item_filters);
    pivot.filters.extend(advanced_filters);
    pivot.calculated_items = cache.calculated_items.clone();

    Ok(Some(pivot))
}

#[derive(Debug, Clone, Copy)]
enum AxisContext {
    Rows,
    Columns,
    Page,
    Data,
}

#[derive(Debug, Clone, Copy)]
struct CurrentDataField {
    measure_index: usize,
    base_field_index: Option<usize>,
}

#[derive(Debug, Clone)]
struct CurrentPivotFilter {
    field_index: usize,
    filter_type: String,
    measure_index: Option<usize>,
    string_value1: Option<String>,
    string_value2: Option<String>,
    custom_values: Vec<String>,
    top: Option<bool>,
    percent: Option<bool>,
    top_n: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LabelPivotFilterType {
    Comparison(PivotFilterOperator),
    Between { not_between: bool },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ValuePivotFilterType {
    Comparison(PivotFilterOperator),
    Between { not_between: bool },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DatePivotFilterType {
    Comparison(PivotFilterOperator),
    Between { not_between: bool },
    Period(PivotDatePeriod),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PivotXmlElement {
    PivotTableDefinition,
    ExtList,
    Other,
}

fn pivot_xml_element(local_name: &[u8]) -> PivotXmlElement {
    match local_name {
        b"pivotTableDefinition" => PivotXmlElement::PivotTableDefinition,
        b"extLst" => PivotXmlElement::ExtList,
        _ => PivotXmlElement::Other,
    }
}

fn is_pivot_table_extension_parent(stack: &[PivotXmlElement]) -> bool {
    matches!(
        stack,
        [
            PivotXmlElement::PivotTableDefinition,
            PivotXmlElement::ExtList
        ]
    )
}

fn read_pivot_extension<B: BufRead>(
    xml_reader: &mut Reader<B>,
    start: BytesStart<'static>,
    buf: &mut Vec<u8>,
) -> XlsxResult<PivotExtension> {
    let uri = attr_string(&start, b"uri").unwrap_or_default();
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer.write_event(Event::Start(start))?;

    let mut depth = 1usize;
    loop {
        buf.clear();
        let event = xml_reader.read_event_into(buf)?;
        match event {
            Event::Start(e) => {
                depth += 1;
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::End(e) => {
                writer.write_event(Event::End(e.into_owned()))?;
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    break;
                }
            }
            Event::Empty(e) => writer.write_event(Event::Empty(e.into_owned()))?,
            Event::Text(e) => writer.write_event(Event::Text(e.into_owned()))?,
            Event::CData(e) => writer.write_event(Event::CData(e.into_owned()))?,
            Event::Comment(e) => writer.write_event(Event::Comment(e.into_owned()))?,
            Event::Decl(e) => writer.write_event(Event::Decl(e.into_owned()))?,
            Event::PI(e) => writer.write_event(Event::PI(e.into_owned()))?,
            Event::DocType(e) => writer.write_event(Event::DocType(e.into_owned()))?,
            Event::Eof => {
                return Err(XlsxError::InvalidFormat(
                    "unexpected EOF while reading pivot extension".into(),
                ))
            }
        }
    }

    Ok(PivotExtension {
        uri,
        payload: writer.into_inner().into_inner(),
    })
}

fn empty_pivot_extension(ext: &BytesStart<'_>) -> XlsxResult<PivotExtension> {
    let uri = attr_string(ext, b"uri").unwrap_or_default();
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer.write_event(Event::Empty(ext.clone().into_owned()))?;
    Ok(PivotExtension {
        uri,
        payload: writer.into_inner().into_inner(),
    })
}

fn parse_pivot_table_attrs(
    e: &BytesStart<'_>,
    name: &mut String,
    cache_id: &mut u32,
    preserve_formatting: &mut bool,
    layout: &mut PivotLayout,
) {
    if let Some(value) = attr_string(e, b"name") {
        *name = value;
    }
    if let Some(value) = attr_string(e, b"dataCaption") {
        layout.data_caption = value;
    }
    if let Some(value) = attr_bool(e, b"dataOnRows") {
        layout.values_axis = if value {
            PivotValuesAxis::Rows
        } else {
            PivotValuesAxis::Columns
        };
    }
    if let Some(value) = attr_u32(e, b"dataPosition") {
        layout.values_axis_position = Some(value);
    }
    if let Some(value) = attr_string(e, b"grandTotalCaption") {
        layout.grand_total_caption = Some(value);
    }
    if let Some(value) = attr_string(e, b"errorCaption") {
        layout.error_caption = Some(value);
    }
    if let Some(value) = attr_bool(e, b"showError") {
        layout.show_error = value;
    }
    if let Some(value) = attr_string(e, b"missingCaption") {
        layout.missing_caption = Some(value);
    }
    if let Some(value) = attr_bool(e, b"showMissing") {
        layout.show_missing = value;
    }
    if let Some(value) = attr_u32(e, b"cacheId") {
        *cache_id = value;
    }
    if let Some(value) = attr_bool(e, b"asteriskTotals") {
        layout.asterisk_totals = value;
    }
    if let Some(value) = attr_bool(e, b"showItems") {
        layout.show_items = value;
    }
    if let Some(value) = attr_bool(e, b"editData") {
        layout.edit_data = value;
    }
    if let Some(value) = attr_bool(e, b"disableFieldList") {
        layout.disable_field_list = value;
    }
    if let Some(value) = attr_bool(e, b"showCalcMbrs") {
        layout.show_calculated_members = value;
    }
    if let Some(value) = attr_bool(e, b"visualTotals") {
        layout.visual_totals = value;
    }
    if let Some(value) = attr_bool(e, b"showMultipleLabel") {
        layout.show_multiple_label = value;
    }
    if let Some(value) = attr_bool(e, b"showDataDropDown") {
        layout.show_data_drop_down = value;
    }
    if let Some(value) = attr_bool(e, b"rowGrandTotals") {
        layout.show_row_grand_totals = value;
    }
    if let Some(value) = attr_bool(e, b"colGrandTotals") {
        layout.show_column_grand_totals = value;
    }
    if let Some(value) = attr_bool(e, b"preserveFormatting") {
        *preserve_formatting = value;
    }
    if let Some(value) = attr_bool(e, b"showHeaders") {
        layout.show_field_headers = value;
    }
    if let Some(value) = attr_bool(e, b"showDrill") {
        layout.show_expand_collapse = value;
    }
    if let Some(value) = attr_bool(e, b"printDrill") {
        layout.print_drill_indicators = value;
    }
    if let Some(value) = attr_bool(e, b"showMemberPropertyTips") {
        layout.show_member_property_tips = value;
    }
    if let Some(value) = attr_bool(e, b"showDataTips") {
        layout.show_data_tips = value;
    }
    if let Some(value) = attr_bool(e, b"enableWizard") {
        layout.enable_wizard = value;
    }
    if let Some(value) = attr_bool(e, b"enableDrill") {
        layout.enable_drill = value;
    }
    if let Some(value) = attr_bool(e, b"enableFieldProperties") {
        layout.enable_field_properties = value;
    }
    if let Some(value) = attr_bool(e, b"itemPrintTitles") {
        layout.item_print_titles = value;
    }
    if let Some(value) = attr_bool(e, b"fieldPrintTitles") {
        layout.field_print_titles = value;
    }
    if let Some(value) = attr_u32(e, b"pageWrap") {
        layout.page_wrap = value;
    }
    if let Some(value) = attr_bool(e, b"pageOverThenDown") {
        layout.page_over_then_down = value;
    }
    if let Some(value) = attr_bool(e, b"subtotalHiddenItems") {
        layout.subtotal_hidden_items = value;
    }
    if let Some(value) = attr_bool(e, b"mergeItem") {
        layout.merge_item_labels = value;
    }
    if let Some(value) = attr_bool(e, b"showDropZones") {
        layout.show_drop_zones = value;
    }
    if let Some(value) = attr_u32(e, b"indent") {
        layout.indent = value;
    }
    if let Some(value) = attr_bool(e, b"showEmptyRow") {
        layout.show_empty_rows = value;
    }
    if let Some(value) = attr_bool(e, b"showEmptyCol") {
        layout.show_empty_columns = value;
    }

    let compact = attr_bool(e, b"compact").unwrap_or(true);
    let outline = attr_bool(e, b"outline").unwrap_or(false);
    layout.kind = if outline {
        PivotLayoutKind::Outline
    } else if compact {
        PivotLayoutKind::Compact
    } else {
        PivotLayoutKind::Tabular
    };
}

fn parse_location(
    e: &BytesStart<'_>,
    target: &mut CellAddress,
    rendered_range: &mut Option<CellRange>,
) -> XlsxResult<()> {
    let Some(ref_text) = attr_string(e, b"ref") else {
        return Ok(());
    };
    let range = CellRange::parse(&ref_text)
        .map_err(|_| XlsxError::InvalidFormat(format!("bad pivot location ref: {ref_text}")))?;
    *target = range.start;
    *rendered_range = Some(range);
    Ok(())
}

fn parse_pivot_style(e: &BytesStart<'_>) -> PivotStyle {
    let mut style = PivotStyle::default();
    style.name = attr_string(e, b"name");
    if let Some(value) = attr_bool(e, b"showRowHeaders") {
        style.show_row_headers = value;
    }
    if let Some(value) = attr_bool(e, b"showColHeaders") {
        style.show_column_headers = value;
    }
    if let Some(value) = attr_bool(e, b"showRowStripes") {
        style.show_row_stripes = value;
    }
    if let Some(value) = attr_bool(e, b"showColStripes") {
        style.show_column_stripes = value;
    }
    if let Some(value) = attr_bool(e, b"showLastColumn") {
        style.show_last_column = value;
    }
    style
}

fn parse_pivot_sort(e: &BytesStart<'_>) -> PivotSort {
    match attr_string(e, b"sortType").as_deref() {
        Some("ascending") => PivotSort::Ascending,
        Some("descending") => PivotSort::Descending,
        Some("manual") | None => PivotSort::None,
        Some(_) => PivotSort::None,
    }
}

#[derive(Debug, Clone)]
struct PivotFieldOptions {
    sort: PivotSort,
    subtotal: PivotSubtotal,
    subtotals: Vec<PivotSubtotal>,
    collapsed_item_indexes: Vec<u32>,
    show_empty_items: bool,
    show_drop_downs: bool,
    subtotal_top: bool,
    insert_blank_row: bool,
    insert_page_break: bool,
    include_new_items_in_filter: bool,
    item_page_count: u32,
}

impl Default for PivotFieldOptions {
    fn default() -> Self {
        Self {
            sort: PivotSort::None,
            subtotal: PivotSubtotal::Automatic,
            subtotals: Vec::new(),
            collapsed_item_indexes: Vec::new(),
            show_empty_items: false,
            show_drop_downs: true,
            subtotal_top: true,
            insert_blank_row: false,
            insert_page_break: false,
            include_new_items_in_filter: false,
            item_page_count: 10,
        }
    }
}

fn parse_pivot_field_options(e: &BytesStart<'_>) -> PivotFieldOptions {
    PivotFieldOptions {
        sort: parse_pivot_sort(e),
        subtotal: parse_pivot_subtotal(e),
        subtotals: parse_pivot_subtotals(e),
        collapsed_item_indexes: Vec::new(),
        show_empty_items: attr_bool(e, b"showAll").unwrap_or(false),
        show_drop_downs: attr_bool(e, b"showDropDowns").unwrap_or(true),
        subtotal_top: attr_bool(e, b"subtotalTop").unwrap_or(true),
        insert_blank_row: attr_bool(e, b"insertBlankRow").unwrap_or(false),
        insert_page_break: attr_bool(e, b"insertPageBreak").unwrap_or(false),
        include_new_items_in_filter: attr_bool(e, b"includeNewItemsInFilter").unwrap_or(false),
        item_page_count: attr_u32(e, b"itemPageCount").unwrap_or(10),
    }
}

fn parse_pivot_subtotal(e: &BytesStart<'_>) -> PivotSubtotal {
    if let Some(subtotal) = parse_pivot_subtotals(e).into_iter().next() {
        return subtotal;
    }

    if attr_bool(e, b"defaultSubtotal").unwrap_or(true) {
        PivotSubtotal::Automatic
    } else {
        PivotSubtotal::None
    }
}

fn parse_pivot_subtotals(e: &BytesStart<'_>) -> Vec<PivotSubtotal> {
    let mut subtotals = Vec::new();
    if attr_bool(e, b"sumSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::Sum);
    }
    if attr_bool(e, b"countASubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::Count);
    }
    if attr_bool(e, b"countSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::CountNumbers);
    }
    if attr_bool(e, b"avgSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::Average);
    }
    if attr_bool(e, b"minSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::Min);
    }
    if attr_bool(e, b"maxSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::Max);
    }
    if attr_bool(e, b"productSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::Product);
    }
    if attr_bool(e, b"stdDevSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::StdDev);
    }
    if attr_bool(e, b"stdDevPSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::StdDevP);
    }
    if attr_bool(e, b"varSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::Var);
    }
    if attr_bool(e, b"varPSubtotal").unwrap_or(false) {
        subtotals.push(PivotSubtotal::VarP);
    }
    subtotals
}

fn pivot_axis_field(
    cache: &PivotCacheDefinition,
    field_index: usize,
    options: Option<PivotFieldOptions>,
) -> Option<PivotField> {
    let field_name = semantic_cache_field_name(cache, field_index)?;
    let options = options.unwrap_or_default();
    let mut pivot_field = PivotField::new(field_name);
    pivot_field.sort = options.sort;
    pivot_field.subtotal = options.subtotal;
    pivot_field.subtotals = options.subtotals;
    pivot_field.show_empty_items = options.show_empty_items;
    pivot_field.show_drop_downs = options.show_drop_downs;
    pivot_field.subtotal_top = options.subtotal_top;
    pivot_field.insert_blank_row = options.insert_blank_row;
    pivot_field.insert_page_break = options.insert_page_break;
    pivot_field.include_new_items_in_filter = options.include_new_items_in_filter;
    pivot_field.item_page_count = options.item_page_count;
    pivot_field.collapsed_items =
        pivot_field_item_values(cache, field_index, &options.collapsed_item_indexes);
    Some(pivot_field)
}

fn pivot_field_item_values(
    cache: &PivotCacheDefinition,
    field_index: usize,
    item_indexes: &[u32],
) -> Vec<PivotValue> {
    let Some(field) = cache.fields.get(field_index) else {
        return Vec::new();
    };
    item_indexes
        .iter()
        .filter_map(|index| pivot_field_item_value(field, *index).cloned())
        .collect()
}

fn pivot_field_item_value(field: &PivotCacheField, index: u32) -> Option<&PivotValue> {
    if field.group_base.is_some() {
        field.group_items.get(index as usize)
    } else {
        field.shared_items.get(index as usize)
    }
}

fn semantic_cache_field_name(cache: &PivotCacheDefinition, field_index: usize) -> Option<String> {
    let field = cache.fields.get(field_index)?;
    field
        .group_base
        .and_then(|base| cache.fields.get(base))
        .map(|field| field.name.clone())
        .or_else(|| Some(field.name.clone()))
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

fn parse_aggregate(value: &str) -> Option<PivotAggregate> {
    Some(match value {
        "average" => PivotAggregate::Average,
        "count" => PivotAggregate::Count,
        "countNums" => PivotAggregate::CountNumbers,
        "max" => PivotAggregate::Max,
        "min" => PivotAggregate::Min,
        "product" => PivotAggregate::Product,
        "stdDev" => PivotAggregate::StdDev,
        "stdDevp" => PivotAggregate::StdDevP,
        "sum" => PivotAggregate::Sum,
        "var" => PivotAggregate::Var,
        "varp" => PivotAggregate::VarP,
        _ => return None,
    })
}

fn parse_data_field(
    e: &BytesStart<'_>,
    cache: &PivotCacheDefinition,
    num_fmts: &HashMap<u32, String>,
) -> Option<PivotMeasure> {
    let field_index = attr_u32(e, b"fld").map(|value| value as usize)?;
    let field = cache.fields.get(field_index)?;
    let aggregate = attr_string(e, b"subtotal")
        .as_deref()
        .and_then(parse_aggregate)
        .unwrap_or(PivotAggregate::Sum);
    let mut measure = PivotMeasure::new(field.name.clone(), aggregate);
    measure.name = attr_string(e, b"name");
    measure.show_as = attr_string(e, b"showDataAs")
        .as_deref()
        .and_then(|value| {
            parse_show_as(
                value,
                cache,
                attr_i32(e, b"baseField").and_then(|value| usize::try_from(value).ok()),
                attr_u32(e, b"baseItem"),
            )
        })
        .unwrap_or(PivotShowAs::Normal);
    measure.number_format = attr_u32(e, b"numFmtId")
        .and_then(|num_fmt_id| pivot_number_format_code(num_fmt_id, num_fmts));
    Some(measure)
}

fn pivot_number_format_code(num_fmt_id: u32, num_fmts: &HashMap<u32, String>) -> Option<String> {
    if num_fmt_id == NumberFormat::ID_GENERAL {
        return None;
    }

    num_fmts.get(&num_fmt_id).cloned().or_else(|| {
        let format = NumberFormat::BuiltIn(num_fmt_id);
        let code = format.format_string();
        (code != "General").then(|| code.to_string())
    })
}

fn parse_current_pivot_filter(e: &BytesStart<'_>) -> Option<CurrentPivotFilter> {
    Some(CurrentPivotFilter {
        field_index: attr_u32(e, b"fld")? as usize,
        filter_type: attr_string(e, b"type")?,
        measure_index: attr_u32(e, b"iMeasureFld").map(|value| value as usize),
        string_value1: attr_string(e, b"stringValue1"),
        string_value2: attr_string(e, b"stringValue2"),
        custom_values: Vec::new(),
        top: None,
        percent: None,
        top_n: None,
    })
}

fn parse_pivot_custom_filter_attrs(filter: &mut CurrentPivotFilter, e: &BytesStart<'_>) {
    if let Some(value) = attr_string(e, b"val") {
        filter.custom_values.push(value);
    }
}

fn parse_pivot_top_filter_attrs(filter: &mut CurrentPivotFilter, e: &BytesStart<'_>) {
    if let Some(value) = attr_bool(e, b"top") {
        filter.top = Some(value);
    }
    if let Some(value) = attr_bool(e, b"percent") {
        filter.percent = Some(value);
    }
    if let Some(value) = attr_u32(e, b"val") {
        filter.top_n = Some(value);
    }
}

fn pivot_filter_from_context(
    filter: CurrentPivotFilter,
    cache: &PivotCacheDefinition,
    measures: &[PivotMeasure],
) -> Option<PivotFilter> {
    let field = cache.fields.get(filter.field_index)?;
    if let Some(filter_type) = parse_label_filter_type(&filter.filter_type) {
        let field = duke_sheets_core::PivotFieldRef::new(field.name.clone());
        return match filter_type {
            LabelPivotFilterType::Comparison(operator) => {
                let value = pivot_filter_operand1(&filter).unwrap_or_default();
                Some(PivotFilter::Label {
                    field,
                    operator,
                    value,
                })
            }
            LabelPivotFilterType::Between { not_between } => {
                let start = pivot_filter_operand1(&filter)?;
                let end = pivot_filter_operand2(&filter)?;
                Some(PivotFilter::LabelBetween {
                    field,
                    start,
                    end,
                    not_between,
                })
            }
        };
    }

    if let Some(filter_type) = parse_value_filter_type(&filter.filter_type) {
        let measure = measure_for_pivot_filter(filter.measure_index, measures)?.clone();
        let field = duke_sheets_core::PivotFieldRef::new(field.name.clone());
        return match filter_type {
            ValuePivotFilterType::Comparison(operator) => {
                let value =
                    pivot_filter_operand1(&filter).and_then(|value| value.parse::<f64>().ok())?;
                Some(PivotFilter::Value {
                    field,
                    measure,
                    operator,
                    value,
                })
            }
            ValuePivotFilterType::Between { not_between } => {
                let start =
                    pivot_filter_operand1(&filter).and_then(|value| value.parse::<f64>().ok())?;
                let end =
                    pivot_filter_operand2(&filter).and_then(|value| value.parse::<f64>().ok())?;
                Some(PivotFilter::ValueBetween {
                    field,
                    measure,
                    start,
                    end,
                    not_between,
                })
            }
        };
    }

    if let Some(filter_type) = parse_date_filter_type(&filter.filter_type) {
        let field = duke_sheets_core::PivotFieldRef::new(field.name.clone());
        return match filter_type {
            DatePivotFilterType::Comparison(operator) => {
                let value = pivot_filter_operand1(&filter)
                    .and_then(|value| parse_pivot_filter_date_value(&value))?;
                Some(PivotFilter::Date {
                    field,
                    operator,
                    value,
                })
            }
            DatePivotFilterType::Between { not_between } => {
                let start = pivot_filter_operand1(&filter)
                    .and_then(|value| parse_pivot_filter_date_value(&value))?;
                let end = pivot_filter_operand2(&filter)
                    .and_then(|value| parse_pivot_filter_date_value(&value))?;
                Some(PivotFilter::DateBetween {
                    field,
                    start,
                    end,
                    not_between,
                })
            }
            DatePivotFilterType::Period(period) => Some(PivotFilter::DatePeriod { field, period }),
        };
    }

    if let Some((type_top, type_percent)) = parse_top_n_filter_type(&filter.filter_type) {
        let measure = measure_for_pivot_filter(filter.measure_index, measures)?.clone();
        return Some(PivotFilter::TopN {
            field: duke_sheets_core::PivotFieldRef::new(field.name.clone()),
            measure,
            n: filter.top_n?,
            top: filter.top.unwrap_or(type_top),
            percent: filter.percent.unwrap_or(type_percent),
        });
    }

    Some(PivotFilter::Unsupported {
        kind: filter.filter_type,
        detail: Some(format!("field={}", field.name)),
    })
}

fn pivot_filter_operand1(filter: &CurrentPivotFilter) -> Option<String> {
    filter
        .string_value1
        .clone()
        .or_else(|| filter.custom_values.first().cloned())
}

fn pivot_filter_operand2(filter: &CurrentPivotFilter) -> Option<String> {
    filter
        .string_value2
        .clone()
        .or_else(|| filter.custom_values.get(1).cloned())
}

fn measure_for_pivot_filter(
    measure_index: Option<usize>,
    measures: &[PivotMeasure],
) -> Option<&PivotMeasure> {
    match measure_index {
        Some(index) => measures.get(index),
        None if measures.len() == 1 => measures.first(),
        None => None,
    }
}

fn parse_label_filter_type(value: &str) -> Option<LabelPivotFilterType> {
    Some(match value {
        "captionEqual" => LabelPivotFilterType::Comparison(PivotFilterOperator::Equals),
        "captionNotEqual" => LabelPivotFilterType::Comparison(PivotFilterOperator::NotEquals),
        "captionLessThan" => LabelPivotFilterType::Comparison(PivotFilterOperator::LessThan),
        "captionLessThanOrEqual" => {
            LabelPivotFilterType::Comparison(PivotFilterOperator::LessThanOrEqual)
        }
        "captionGreaterThan" => LabelPivotFilterType::Comparison(PivotFilterOperator::GreaterThan),
        "captionGreaterThanOrEqual" => {
            LabelPivotFilterType::Comparison(PivotFilterOperator::GreaterThanOrEqual)
        }
        "captionBeginsWith" => LabelPivotFilterType::Comparison(PivotFilterOperator::BeginsWith),
        "captionNotBeginsWith" => {
            LabelPivotFilterType::Comparison(PivotFilterOperator::DoesNotBeginWith)
        }
        "captionEndsWith" => LabelPivotFilterType::Comparison(PivotFilterOperator::EndsWith),
        "captionNotEndsWith" => {
            LabelPivotFilterType::Comparison(PivotFilterOperator::DoesNotEndWith)
        }
        "captionContains" => LabelPivotFilterType::Comparison(PivotFilterOperator::Contains),
        "captionNotContains" => {
            LabelPivotFilterType::Comparison(PivotFilterOperator::DoesNotContain)
        }
        "captionBetween" => LabelPivotFilterType::Between { not_between: false },
        "captionNotBetween" => LabelPivotFilterType::Between { not_between: true },
        _ => return None,
    })
}

fn parse_value_filter_type(value: &str) -> Option<ValuePivotFilterType> {
    Some(match value {
        "valueEqual" => ValuePivotFilterType::Comparison(PivotFilterOperator::Equals),
        "valueNotEqual" => ValuePivotFilterType::Comparison(PivotFilterOperator::NotEquals),
        "valueLessThan" => ValuePivotFilterType::Comparison(PivotFilterOperator::LessThan),
        "valueLessThanOrEqual" => {
            ValuePivotFilterType::Comparison(PivotFilterOperator::LessThanOrEqual)
        }
        "valueGreaterThan" => ValuePivotFilterType::Comparison(PivotFilterOperator::GreaterThan),
        "valueGreaterThanOrEqual" => {
            ValuePivotFilterType::Comparison(PivotFilterOperator::GreaterThanOrEqual)
        }
        "valueBetween" => ValuePivotFilterType::Between { not_between: false },
        "valueNotBetween" => ValuePivotFilterType::Between { not_between: true },
        _ => return None,
    })
}

fn parse_date_filter_type(value: &str) -> Option<DatePivotFilterType> {
    Some(match value {
        "dateEqual" => DatePivotFilterType::Comparison(PivotFilterOperator::Equals),
        "dateNotEqual" => DatePivotFilterType::Comparison(PivotFilterOperator::NotEquals),
        "dateOlderThan" => DatePivotFilterType::Comparison(PivotFilterOperator::LessThan),
        "dateOlderThanOrEqual" => {
            DatePivotFilterType::Comparison(PivotFilterOperator::LessThanOrEqual)
        }
        "dateNewerThan" => DatePivotFilterType::Comparison(PivotFilterOperator::GreaterThan),
        "dateNewerThanOrEqual" => {
            DatePivotFilterType::Comparison(PivotFilterOperator::GreaterThanOrEqual)
        }
        "dateBetween" => DatePivotFilterType::Between { not_between: false },
        "dateNotBetween" => DatePivotFilterType::Between { not_between: true },
        "tomorrow" => DatePivotFilterType::Period(PivotDatePeriod::Tomorrow),
        "today" => DatePivotFilterType::Period(PivotDatePeriod::Today),
        "yesterday" => DatePivotFilterType::Period(PivotDatePeriod::Yesterday),
        "nextWeek" => DatePivotFilterType::Period(PivotDatePeriod::NextWeek),
        "thisWeek" => DatePivotFilterType::Period(PivotDatePeriod::ThisWeek),
        "lastWeek" => DatePivotFilterType::Period(PivotDatePeriod::LastWeek),
        "nextMonth" => DatePivotFilterType::Period(PivotDatePeriod::NextMonth),
        "thisMonth" => DatePivotFilterType::Period(PivotDatePeriod::ThisMonth),
        "lastMonth" => DatePivotFilterType::Period(PivotDatePeriod::LastMonth),
        "nextQuarter" => DatePivotFilterType::Period(PivotDatePeriod::NextQuarter),
        "thisQuarter" => DatePivotFilterType::Period(PivotDatePeriod::ThisQuarter),
        "lastQuarter" => DatePivotFilterType::Period(PivotDatePeriod::LastQuarter),
        "nextYear" => DatePivotFilterType::Period(PivotDatePeriod::NextYear),
        "thisYear" => DatePivotFilterType::Period(PivotDatePeriod::ThisYear),
        "lastYear" => DatePivotFilterType::Period(PivotDatePeriod::LastYear),
        "yearToDate" => DatePivotFilterType::Period(PivotDatePeriod::YearToDate),
        "Q1" => DatePivotFilterType::Period(PivotDatePeriod::Quarter(1)),
        "Q2" => DatePivotFilterType::Period(PivotDatePeriod::Quarter(2)),
        "Q3" => DatePivotFilterType::Period(PivotDatePeriod::Quarter(3)),
        "Q4" => DatePivotFilterType::Period(PivotDatePeriod::Quarter(4)),
        "M1" => DatePivotFilterType::Period(PivotDatePeriod::Month(1)),
        "M2" => DatePivotFilterType::Period(PivotDatePeriod::Month(2)),
        "M3" => DatePivotFilterType::Period(PivotDatePeriod::Month(3)),
        "M4" => DatePivotFilterType::Period(PivotDatePeriod::Month(4)),
        "M5" => DatePivotFilterType::Period(PivotDatePeriod::Month(5)),
        "M6" => DatePivotFilterType::Period(PivotDatePeriod::Month(6)),
        "M7" => DatePivotFilterType::Period(PivotDatePeriod::Month(7)),
        "M8" => DatePivotFilterType::Period(PivotDatePeriod::Month(8)),
        "M9" => DatePivotFilterType::Period(PivotDatePeriod::Month(9)),
        "M10" => DatePivotFilterType::Period(PivotDatePeriod::Month(10)),
        "M11" => DatePivotFilterType::Period(PivotDatePeriod::Month(11)),
        "M12" => DatePivotFilterType::Period(PivotDatePeriod::Month(12)),
        _ => return None,
    })
}

fn parse_pivot_filter_date_value(value: &str) -> Option<f64> {
    if let Ok(serial) = value.parse::<f64>() {
        return serial.is_finite().then_some(serial);
    }

    let date_part = value.get(..10).unwrap_or(value);
    let date = NaiveDate::parse_from_str(date_part, "%Y-%m-%d").ok()?;
    Some(date_to_serial(
        date.year(),
        date.month(),
        date.day(),
        DateSystem::Date1900,
    ))
}

fn parse_top_n_filter_type(value: &str) -> Option<(bool, bool)> {
    Some(match value {
        "topCount" => (true, false),
        "topPercent" => (true, true),
        "bottomCount" => (false, false),
        "bottomPercent" => (false, true),
        _ => return None,
    })
}

fn parse_show_as(
    value: &str,
    cache: &PivotCacheDefinition,
    base_field_index: Option<usize>,
    base_item_index: Option<u32>,
) -> Option<PivotShowAs> {
    Some(match value {
        "normal" => PivotShowAs::Normal,
        "percentOfTotal" => PivotShowAs::PercentOfGrandTotal,
        "percentOfRow" => PivotShowAs::PercentOfRowTotal,
        "percentOfCol" => PivotShowAs::PercentOfColumnTotal,
        "index" => PivotShowAs::Index,
        "runTotal" => PivotShowAs::RunningTotal {
            base_field: duke_sheets_core::PivotFieldRef::new(
                cache.fields.get(base_field_index?)?.name.clone(),
            ),
        },
        "difference" => PivotShowAs::DifferenceFrom {
            base_field: duke_sheets_core::PivotFieldRef::new(
                cache.fields.get(base_field_index?)?.name.clone(),
            ),
            base_item: cache
                .fields
                .get(base_field_index?)?
                .shared_items
                .get(base_item_index? as usize)?
                .clone(),
        },
        "percentDiff" => PivotShowAs::PercentDifferenceFrom {
            base_field: duke_sheets_core::PivotFieldRef::new(
                cache.fields.get(base_field_index?)?.name.clone(),
            ),
            base_item: cache
                .fields
                .get(base_field_index?)?
                .shared_items
                .get(base_item_index? as usize)?
                .clone(),
        },
        "rankAscending" => PivotShowAs::RankAscending {
            base_field: duke_sheets_core::PivotFieldRef::new(
                cache.fields.get(base_field_index?)?.name.clone(),
            ),
        },
        "rankDescending" => PivotShowAs::RankDescending {
            base_field: duke_sheets_core::PivotFieldRef::new(
                cache.fields.get(base_field_index?)?.name.clone(),
            ),
        },
        _ => return None,
    })
}

fn parse_x14_show_as(
    value: &str,
    cache: &PivotCacheDefinition,
    source_field_index: Option<usize>,
) -> Option<PivotShowAs> {
    match value {
        "percentOfParent" => Some(PivotShowAs::PercentOfParentTotal {
            base_field: duke_sheets_core::PivotFieldRef::new(
                cache.fields.get(source_field_index?)?.name.clone(),
            ),
        }),
        "percentOfParentRow" => Some(PivotShowAs::PercentOfParentRowTotal),
        "percentOfParentCol" => Some(PivotShowAs::PercentOfParentColumnTotal),
        "rankAscending" | "rankDescending" => parse_show_as(value, cache, source_field_index, None),
        _ => None,
    }
}

fn attr_string(e: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    e.attributes().flatten().find_map(|attr| {
        if attr.key.local_name().as_ref() == name {
            attr.unescape_value().ok().map(|value| value.to_string())
        } else {
            None
        }
    })
}

fn attr_bool(e: &BytesStart<'_>, name: &[u8]) -> Option<bool> {
    attr_string(e, name).map(|value| value == "1" || value.eq_ignore_ascii_case("true"))
}

fn attr_u32(e: &BytesStart<'_>, name: &[u8]) -> Option<u32> {
    attr_string(e, name).and_then(|value| value.parse().ok())
}

fn attr_i32(e: &BytesStart<'_>, name: &[u8]) -> Option<i32> {
    attr_string(e, name).and_then(|value| value.parse().ok())
}

fn attr_u64(e: &BytesStart<'_>, name: &[u8]) -> Option<u64> {
    attr_string(e, name).and_then(|value| value.parse().ok())
}

fn attr_f64(e: &BytesStart<'_>, name: &[u8]) -> Option<f64> {
    attr_string(e, name).and_then(|value| value.parse().ok())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cache_source_with_type(value: &'static str) -> BytesStart<'static> {
        let mut element = BytesStart::new("cacheSource");
        element.push_attribute(("type", value));
        element
    }

    fn minimal_cache_with_field(field_name: &str) -> PivotCacheDefinition {
        PivotCacheDefinition {
            source: PivotSource::range(CellRange::parse("A1:A2").unwrap()),
            source_kind: PivotCacheSourceKind::Worksheet,
            fields: vec![PivotCacheField {
                name: field_name.to_string(),
                formula: None,
                shared_items: Vec::new(),
                grouping: None,
                group_base: None,
                group_parent: None,
                discrete_item_indexes: Vec::new(),
                group_items: Vec::new(),
            }],
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            groupings: Vec::new(),
            record_count: None,
            refreshed_version: None,
            refresh_on_load: false,
            background_query: false,
            missing_items_limit: None,
        }
    }

    #[test]
    fn parses_cache_source_type_for_diagnostics() {
        assert_eq!(
            parse_cache_source_kind(&cache_source_with_type("worksheet")),
            PivotCacheSourceKind::Worksheet
        );
        assert_eq!(
            parse_cache_source_kind(&cache_source_with_type("external")),
            PivotCacheSourceKind::External
        );
        assert_eq!(
            parse_cache_source_kind(&cache_source_with_type("consolidation")),
            PivotCacheSourceKind::Consolidation
        );
        assert_eq!(
            parse_cache_source_kind(&cache_source_with_type("scenario")),
            PivotCacheSourceKind::Scenario
        );
        assert_eq!(
            parse_cache_source_kind(&cache_source_with_type("olap")),
            PivotCacheSourceKind::Olap
        );
        assert_eq!(
            parse_cache_source_kind(&cache_source_with_type("mystery")),
            PivotCacheSourceKind::Unknown
        );
    }

    #[test]
    fn placeholder_sources_keep_non_refreshable_kind_shape() {
        let empty_connections = HashMap::new();
        match placeholder_source_for_kind(
            PivotCacheSourceKind::External,
            Some(7),
            Some("7".to_string()),
            &empty_connections,
        ) {
            PivotSource::External {
                connection_name, ..
            } => assert_eq!(connection_name, "7"),
            other => panic!("unexpected source: {other:?}"),
        }
        let connections = HashMap::from([(
            7,
            WorkbookConnection::database(7, "SalesConnection", "Provider=Test;")
                .with_command("select Region, Revenue from Sales"),
        )]);
        match placeholder_source_for_kind(
            PivotCacheSourceKind::External,
            Some(7),
            Some("7".to_string()),
            &connections,
        ) {
            PivotSource::External {
                connection_name,
                command_text,
            } => {
                assert_eq!(connection_name, "SalesConnection");
                assert_eq!(
                    command_text.as_deref(),
                    Some("select Region, Revenue from Sales")
                );
            }
            other => panic!("unexpected source: {other:?}"),
        }
        assert!(matches!(
            placeholder_source_for_kind(
                PivotCacheSourceKind::Consolidation,
                None,
                None,
                &empty_connections
            ),
            PivotSource::Consolidation { .. }
        ));
        assert!(matches!(
            placeholder_source_for_kind(
                PivotCacheSourceKind::Scenario,
                None,
                None,
                &empty_connections
            ),
            PivotSource::Scenario { .. }
        ));
        assert!(matches!(
            placeholder_source_for_kind(
                PivotCacheSourceKind::Olap,
                Some(3),
                Some("3".to_string()),
                &empty_connections
            ),
            PivotSource::Olap { .. }
        ));
    }

    #[test]
    fn date_pivot_filter_types_are_parsed() {
        let cache = minimal_cache_with_field("Order Date");
        let filter = CurrentPivotFilter {
            field_index: 0,
            filter_type: "dateBetween".to_string(),
            measure_index: None,
            string_value1: Some("2024-01-01".to_string()),
            string_value2: Some("2024-01-31T00:00:00".to_string()),
            custom_values: Vec::new(),
            top: None,
            percent: None,
            top_n: None,
        };

        let parsed = pivot_filter_from_context(filter, &cache, &[]).expect("date filter");
        match parsed {
            PivotFilter::DateBetween {
                field,
                start,
                end,
                not_between,
            } => {
                assert_eq!(field.name, "Order Date");
                assert_eq!(start, date_to_serial(2024, 1, 1, DateSystem::Date1900));
                assert_eq!(end, date_to_serial(2024, 1, 31, DateSystem::Date1900));
                assert!(!not_between);
            }
            other => panic!("unexpected filter: {other:?}"),
        }

        let serial = date_to_serial(2024, 2, 1, DateSystem::Date1900);
        let filter = CurrentPivotFilter {
            field_index: 0,
            filter_type: "dateNewerThanOrEqual".to_string(),
            measure_index: None,
            string_value1: None,
            string_value2: None,
            custom_values: vec![serial.to_string()],
            top: None,
            percent: None,
            top_n: None,
        };

        let parsed = pivot_filter_from_context(filter, &cache, &[]).expect("date filter");
        match parsed {
            PivotFilter::Date {
                field,
                operator,
                value,
            } => {
                assert_eq!(field.name, "Order Date");
                assert_eq!(operator, PivotFilterOperator::GreaterThanOrEqual);
                assert_eq!(value, serial);
            }
            other => panic!("unexpected filter: {other:?}"),
        }

        let filter = CurrentPivotFilter {
            field_index: 0,
            filter_type: "thisQuarter".to_string(),
            measure_index: None,
            string_value1: None,
            string_value2: None,
            custom_values: Vec::new(),
            top: None,
            percent: None,
            top_n: None,
        };

        let parsed = pivot_filter_from_context(filter, &cache, &[]).expect("date period filter");
        match parsed {
            PivotFilter::DatePeriod { field, period } => {
                assert_eq!(field.name, "Order Date");
                assert_eq!(period, PivotDatePeriod::ThisQuarter);
            }
            other => panic!("unexpected filter: {other:?}"),
        }
    }

    #[test]
    fn range_pivot_filter_types_are_parsed() {
        let cache = minimal_cache_with_field("Region");
        let filter = CurrentPivotFilter {
            field_index: 0,
            filter_type: "captionBetween".to_string(),
            measure_index: None,
            string_value1: Some("East".to_string()),
            string_value2: Some("North".to_string()),
            custom_values: Vec::new(),
            top: None,
            percent: None,
            top_n: None,
        };

        let parsed = pivot_filter_from_context(filter, &cache, &[]).expect("caption range filter");
        match parsed {
            PivotFilter::LabelBetween {
                field,
                start,
                end,
                not_between,
            } => {
                assert_eq!(field.name, "Region");
                assert_eq!(start, "East");
                assert_eq!(end, "North");
                assert!(!not_between);
            }
            other => panic!("unexpected filter: {other:?}"),
        }

        let measures = vec![PivotMeasure::new("Revenue", PivotAggregate::Sum)];
        let filter = CurrentPivotFilter {
            field_index: 0,
            filter_type: "valueNotBetween".to_string(),
            measure_index: Some(0),
            string_value1: None,
            string_value2: None,
            custom_values: vec!["10".to_string(), "30".to_string()],
            top: None,
            percent: None,
            top_n: None,
        };

        let parsed = pivot_filter_from_context(filter, &cache, &measures).expect("value range");
        match parsed {
            PivotFilter::ValueBetween {
                field,
                measure,
                start,
                end,
                not_between,
            } => {
                assert_eq!(field.name, "Region");
                assert_eq!(measure.field.name, "Revenue");
                assert_eq!(start, 10.0);
                assert_eq!(end, 30.0);
                assert!(not_between);
            }
            other => panic!("unexpected filter: {other:?}"),
        }
    }

    #[test]
    fn unsupported_pivot_filter_types_are_preserved() {
        let cache = minimal_cache_with_field("Region");
        let filter = CurrentPivotFilter {
            field_index: 0,
            filter_type: "captionAboveAverage".to_string(),
            measure_index: None,
            string_value1: None,
            string_value2: None,
            custom_values: Vec::new(),
            top: None,
            percent: None,
            top_n: None,
        };

        let parsed = pivot_filter_from_context(filter, &cache, &[]).expect("preserved filter");
        match parsed {
            PivotFilter::Unsupported { kind, detail } => {
                assert_eq!(kind, "captionAboveAverage");
                assert_eq!(detail.as_deref(), Some("field=Region"));
            }
            other => panic!("unexpected filter: {other:?}"),
        }
    }
}
