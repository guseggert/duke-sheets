use std::collections::{HashMap, HashSet};
use std::io::{BufReader, Read, Seek};

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

use super::archive_by_name;
use crate::error::{XlsxError, XlsxResult};
use duke_sheets_core::{
    CellAddress, CellError, CellRange, PivotAggregate, PivotCacheInfo, PivotCacheSourceKind,
    PivotCalculatedField, PivotDateGroupUnit, PivotField, PivotFilter, PivotGrouping,
    PivotLayoutKind, PivotMeasure, PivotRefreshStatus, PivotShowAs, PivotSource, PivotStyle,
    PivotTable, PivotValue,
};

#[derive(Debug, Clone)]
pub(super) struct PivotCacheDefinition {
    pub(super) source: PivotSource,
    pub(super) fields: Vec<PivotCacheField>,
    pub(super) calculated_fields: Vec<PivotCalculatedField>,
    pub(super) groupings: Vec<PivotGrouping>,
    pub(super) record_count: Option<u64>,
    pub(super) refreshed_version: Option<String>,
}

#[derive(Debug, Clone)]
pub(super) struct PivotCacheField {
    pub(super) name: String,
    formula: Option<String>,
    pub(super) shared_items: Vec<PivotValue>,
    grouping: Option<PivotGrouping>,
}

pub(super) fn read_pivot_cache_definition<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    _cache_id: u32,
    path: &str,
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
    let mut fields = Vec::new();
    let mut record_count = None;
    let mut refreshed_version = None;
    let mut current_field: Option<PivotCacheField> = None;
    let mut in_shared_items = false;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"pivotCacheDefinition" => {
                    record_count = attr_u64(&e, b"recordCount");
                    refreshed_version = attr_string(&e, b"refreshedVersion");
                }
                b"worksheetSource" => source = parse_worksheet_source(&e)?,
                b"cacheField" => {
                    current_field = Some(PivotCacheField {
                        name: attr_string(&e, b"name").unwrap_or_default(),
                        formula: attr_string(&e, b"formula"),
                        shared_items: Vec::new(),
                        grouping: None,
                    });
                }
                b"sharedItems" => in_shared_items = true,
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
                }
                b"worksheetSource" => source = parse_worksheet_source(&e)?,
                b"cacheField" => fields.push(PivotCacheField {
                    name: attr_string(&e, b"name").unwrap_or_default(),
                    formula: attr_string(&e, b"formula"),
                    shared_items: Vec::new(),
                    grouping: None,
                }),
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
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"cacheField" => {
                    if let Some(field) = current_field.take() {
                        fields.push(field);
                    }
                }
                b"sharedItems" => in_shared_items = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    let source = source.unwrap_or_else(|| PivotSource::External {
        connection_name: String::new(),
        command_text: None,
    });
    let groupings = fields
        .iter()
        .filter_map(|field| field.grouping.clone())
        .collect();
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
        fields,
        calculated_fields,
        groupings,
        record_count,
        refreshed_version,
    }))
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
    let mut style = PivotStyle::default();
    let mut row_grand_totals = true;
    let mut col_grand_totals = true;
    let mut show_headers = true;
    let mut layout_kind = PivotLayoutKind::Compact;
    let mut axis_context: Option<AxisContext> = None;
    let mut pivot_field_index = 0usize;
    let mut current_pivot_field: Option<(usize, Vec<u32>)> = None;
    let mut current_data_field: Option<CurrentDataField> = None;
    let mut hidden_items_by_field: HashMap<usize, Vec<u32>> = HashMap::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"pivotTableDefinition" => {
                    parse_pivot_table_attrs(
                        &e,
                        &mut name,
                        &mut cache_id,
                        &mut row_grand_totals,
                        &mut col_grand_totals,
                        &mut show_headers,
                        &mut layout_kind,
                    );
                }
                b"location" => parse_location(&e, &mut target, &mut rendered_range)?,
                b"pivotField" => {
                    current_pivot_field = Some((pivot_field_index, Vec::new()));
                    pivot_field_index += 1;
                }
                b"dataField" if attr_u32(&e, b"fld").is_some() => {
                    let Some(cache) = caches.get(&cache_id) else {
                        buf.clear();
                        continue;
                    };
                    if let Some(measure) = parse_data_field(&e, cache) {
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
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"pivotTableDefinition" => {
                    parse_pivot_table_attrs(
                        &e,
                        &mut name,
                        &mut cache_id,
                        &mut row_grand_totals,
                        &mut col_grand_totals,
                        &mut show_headers,
                        &mut layout_kind,
                    );
                }
                b"location" => parse_location(&e, &mut target, &mut rendered_range)?,
                b"pivotField" => {
                    pivot_field_index += 1;
                }
                b"item" => {
                    if let Some((_, hidden)) = &mut current_pivot_field {
                        if attr_bool(&e, b"h").unwrap_or(false) {
                            if let Some(index) = attr_u32(&e, b"x") {
                                hidden.push(index);
                            }
                        }
                    }
                }
                b"field" => {
                    let Some(cache) = caches.get(&cache_id) else {
                        buf.clear();
                        continue;
                    };
                    if let Some(field_index) =
                        attr_i32(&e, b"x").and_then(|v| usize::try_from(v).ok())
                    {
                        if let Some(field) = cache.fields.get(field_index) {
                            match axis_context {
                                Some(AxisContext::Rows) => {
                                    rows.push(PivotField::new(field.name.clone()))
                                }
                                Some(AxisContext::Columns) => {
                                    columns.push(PivotField::new(field.name.clone()))
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
                        if let Some(field) = cache.fields.get(field_index) {
                            page_fields.push(PivotField::new(field.name.clone()));
                            if let Some(item_index) = attr_u32(&e, b"item") {
                                if let Some(item) = field.shared_items.get(item_index as usize) {
                                    page_filters.push(PivotFilter::FieldItems {
                                        field: duke_sheets_core::PivotFieldRef::new(
                                            field.name.clone(),
                                        ),
                                        allowed_items: vec![item.clone()],
                                    });
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
                            if let Some(show_as) =
                                parse_x14_show_as(&pivot_show_as, cache, current.base_field_index)
                            {
                                if let Some(measure) = measures.get_mut(current.measure_index) {
                                    measure.show_as = show_as;
                                }
                            }
                        }
                    } else if let Some(measure) = parse_data_field(&e, cache) {
                        measures.push(measure);
                    }
                }
                b"pivotTableStyleInfo" => style = parse_pivot_style(&e),
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"pivotField" => {
                    if let Some((field_index, hidden_items)) = current_pivot_field.take() {
                        if !hidden_items.is_empty() {
                            hidden_items_by_field.insert(field_index, hidden_items);
                        }
                    }
                }
                b"rowFields" | b"colFields" | b"pageFields" | b"dataFields" => axis_context = None,
                b"dataField" => {
                    if e.name().as_ref() == b"dataField" {
                        current_data_field = None;
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

    let Some(cache) = caches.get(&cache_id) else {
        return Ok(None);
    };

    let mut pivot = PivotTable::new(0, name, cache.source.clone(), target);
    pivot.rows = rows;
    pivot.columns = columns;
    pivot.page_fields = page_fields;
    pivot.measures = measures;
    pivot.calculated_fields = cache.calculated_fields.clone();
    pivot.layout.show_row_grand_totals = row_grand_totals;
    pivot.layout.show_column_grand_totals = col_grand_totals;
    pivot.layout.show_field_headers = show_headers;
    pivot.layout.kind = layout_kind;
    pivot.style = style;
    pivot.rendered_range = rendered_range;
    pivot.groupings = cache.groupings.clone();
    pivot.cache_info = Some(PivotCacheInfo {
        cache_id,
        source_kind: cache_source_kind(&cache.source),
        record_count: cache.record_count,
        refreshed_version: cache.refreshed_version.clone(),
        refresh_status: PivotRefreshStatus::NotRefreshed,
    });

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

fn parse_pivot_table_attrs(
    e: &BytesStart<'_>,
    name: &mut String,
    cache_id: &mut u32,
    row_grand_totals: &mut bool,
    col_grand_totals: &mut bool,
    show_headers: &mut bool,
    layout_kind: &mut PivotLayoutKind,
) {
    if let Some(value) = attr_string(e, b"name") {
        *name = value;
    }
    if let Some(value) = attr_u32(e, b"cacheId") {
        *cache_id = value;
    }
    if let Some(value) = attr_bool(e, b"rowGrandTotals") {
        *row_grand_totals = value;
    }
    if let Some(value) = attr_bool(e, b"colGrandTotals") {
        *col_grand_totals = value;
    }
    if let Some(value) = attr_bool(e, b"showHeaders") {
        *show_headers = value;
    }

    let compact = attr_bool(e, b"compact").unwrap_or(true);
    let outline = attr_bool(e, b"outline").unwrap_or(false);
    *layout_kind = if outline {
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
    style
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

fn parse_data_field(e: &BytesStart<'_>, cache: &PivotCacheDefinition) -> Option<PivotMeasure> {
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
    Some(measure)
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
    base_field_index: Option<usize>,
) -> Option<PivotShowAs> {
    match value {
        "rankAscending" | "rankDescending" => parse_show_as(value, cache, base_field_index, None),
        _ => None,
    }
}

fn cache_source_kind(source: &PivotSource) -> PivotCacheSourceKind {
    match source {
        PivotSource::WorksheetRange { .. } | PivotSource::Table { .. } => {
            PivotCacheSourceKind::Worksheet
        }
        PivotSource::External { .. } => PivotCacheSourceKind::External,
        PivotSource::Consolidation { .. } => PivotCacheSourceKind::Consolidation,
        PivotSource::Scenario { .. } => PivotCacheSourceKind::Scenario,
        PivotSource::Olap { .. } => PivotCacheSourceKind::Olap,
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
