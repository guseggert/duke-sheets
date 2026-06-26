use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::{resolve_rel_path, SheetRel};
use crate::biff12::parser;
use crate::biff12::records;
use crate::biff12::RecordIter;
use crate::error::{XlsbError, XlsbResult};
use duke_sheets_core::{
    CellAddress, CellError, CellRange, PivotAggregate, PivotCacheInfo, PivotCacheSourceKind,
    PivotField, PivotFilter, PivotMeasure, PivotRefreshStatus, PivotSource, PivotStyle, PivotTable,
    PivotValue, PivotValuesAxis,
};

#[derive(Debug, Clone)]
struct PivotCacheDefinition {
    cache_id: u32,
    source: PivotSource,
    source_kind: PivotCacheSourceKind,
    fields: Vec<PivotCacheField>,
    record_count: Option<u64>,
    refresh_on_load: bool,
}

#[derive(Debug, Clone)]
struct PivotCacheField {
    name: String,
    shared_items: Vec<PivotValue>,
}

pub(crate) fn read_pivot_tables_for_sheet<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    sheet_path: &str,
    sheet_rels: &HashMap<String, SheetRel>,
) -> XlsbResult<Vec<PivotTable>> {
    let mut pivot_paths = sheet_rels
        .values()
        .filter(|rel| rel.rel_type.ends_with("/pivotTable"))
        .map(|rel| resolve_rel_path(sheet_path, &rel.target))
        .collect::<Vec<_>>();
    pivot_paths.sort();

    let mut pivots = Vec::new();
    for pivot_path in pivot_paths {
        if let Some(pivot) = read_pivot_table(archive, &pivot_path)? {
            pivots.push(pivot);
        }
    }
    Ok(pivots)
}

fn read_pivot_table<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> XlsbResult<Option<PivotTable>> {
    let Some(cache_path) = related_part_path(archive, path, "/pivotCacheDefinition")? else {
        return Ok(None);
    };
    let cache_id = trailing_number(&cache_path).unwrap_or(0);
    let Some(cache) = read_pivot_cache_definition(archive, cache_id, &cache_path)? else {
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
    let mut page_fields = Vec::new();
    let mut measures = Vec::new();
    let mut filters = Vec::new();
    let mut layout = duke_sheets_core::PivotLayout::default();
    let mut style = PivotStyle::default();

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
                parse_begin_sxview(payload, &mut name, &mut layout)?;
            }
            records::BRT_SX_LOCATION => {
                parse_sx_location(payload, &mut target, &mut rendered_range);
            }
            records::BRT_BEGIN_ISXVD_RWS => {
                parse_axis_fields(
                    payload,
                    PivotValuesAxis::Rows,
                    &cache,
                    &mut rows,
                    &mut layout,
                );
            }
            records::BRT_BEGIN_ISXVD_COLS => {
                parse_axis_fields(
                    payload,
                    PivotValuesAxis::Columns,
                    &cache,
                    &mut columns,
                    &mut layout,
                );
            }
            records::BRT_BEGIN_SXPI => {
                parse_page_field(payload, &cache, &mut page_fields, &mut filters);
            }
            records::BRT_BEGIN_SXDI => {
                if let Some(measure) = parse_data_field(payload, &cache)? {
                    measures.push(measure);
                }
            }
            records::BRT_SX_VIEW_STYLE => {
                style = parse_pivot_style(payload)?;
            }
            records::BRT_END_SXVIEW => break,
            _ => {}
        }
    }

    if name.trim().is_empty() {
        name = format!("PivotTable{}", cache.cache_id);
    }

    let mut pivot = PivotTable::new(0, name, cache.source.clone(), target);
    pivot.rows = rows;
    pivot.columns = columns;
    pivot.page_fields = page_fields;
    pivot.measures = measures;
    pivot.filters = filters;
    pivot.layout = layout;
    pivot.style = style;
    pivot.refresh_policy.refresh_on_open = cache.refresh_on_load;
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
) -> XlsbResult<Option<PivotCacheDefinition>> {
    let file = match archive.by_name(path) {
        Ok(file) => file,
        Err(_) => return Ok(None),
    };
    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(1024);

    let mut source = None;
    let mut source_kind = PivotCacheSourceKind::Unknown;
    let mut fields = Vec::new();
    let mut current_field: Option<PivotCacheField> = None;
    let mut record_count = None;
    let mut refresh_on_load = false;

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
                if payload.len() >= 21 {
                    record_count = Some(parser::read_u32(payload, 17) as u64);
                }
            }
            records::BRT_BEGIN_PCDS_SHEET => {
                if let Some(parsed) = parse_cache_sheet_source(payload)? {
                    source = Some(parsed);
                    source_kind = PivotCacheSourceKind::Worksheet;
                }
            }
            records::BRT_BEGIN_PCD_FIELD => {
                current_field = Some(PivotCacheField {
                    name: parse_cache_field_name(payload)?,
                    shared_items: Vec::new(),
                });
            }
            records::BRT_PCDI_MISSING
            | records::BRT_PCDI_BOOLEAN
            | records::BRT_PCDI_ERROR
            | records::BRT_PCDI_NUMBER
            | records::BRT_PCDI_STRING => {
                if let Some(field) = &mut current_field {
                    field.shared_items.push(parse_shared_item(typ, payload)?);
                }
            }
            records::BRT_END_PCD_FIELD => {
                if let Some(field) = current_field.take() {
                    fields.push(field);
                }
            }
            records::BRT_END_PIVOT_CACHE_DEF => break,
            _ => {}
        }
    }

    let Some(source) = source else {
        return Ok(None);
    };

    Ok(Some(PivotCacheDefinition {
        cache_id,
        source,
        source_kind,
        fields,
        record_count,
        refresh_on_load,
    }))
}

fn parse_begin_sxview(
    payload: &[u8],
    name: &mut String,
    layout: &mut duke_sheets_core::PivotLayout,
) -> XlsbResult<()> {
    if payload.len() < 32 {
        return Ok(());
    }
    layout.values_axis = match payload[12] {
        0x01 => PivotValuesAxis::Rows,
        0x02 => PivotValuesAxis::Columns,
        _ => layout.values_axis,
    };
    let values_position = parser::read_i32(payload, 16);
    if values_position >= 0 {
        layout.values_axis_position = Some(values_position as u32);
    }

    let (view_name, consumed) = parser::wide_str(payload, 32)?;
    *name = view_name;
    let data_caption_offset = 32 + consumed;
    if data_caption_offset < payload.len() {
        let (data_caption, _) = parser::wide_str(payload, data_caption_offset)?;
        if !data_caption.is_empty() {
            layout.data_caption = data_caption;
        }
    }
    Ok(())
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
    fields: &mut Vec<PivotField>,
    layout: &mut duke_sheets_core::PivotLayout,
) {
    if payload.len() < 4 {
        return;
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
            if let Some(field) = cache
                .fields
                .get(index as usize)
                .map(|field| PivotField::new(field.name.clone()))
            {
                push_axis_field(fields, field);
            }
        }
    }
}

fn parse_page_field(
    payload: &[u8],
    cache: &PivotCacheDefinition,
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
    push_axis_field(page_fields, PivotField::new(field.name.clone()));

    if let Some(item) = field.shared_items.get(selected_item as usize) {
        filters.push(PivotFilter::FieldItems {
            field: duke_sheets_core::PivotFieldRef::new(field.name.clone()),
            allowed_items: vec![item.clone()],
        });
    }
}

fn parse_data_field(
    payload: &[u8],
    cache: &PivotCacheDefinition,
) -> XlsbResult<Option<PivotMeasure>> {
    if payload.len() < 25 {
        return Ok(None);
    }
    let field_index = parser::read_u32(payload, 0) as usize;
    let Some(field) = cache.fields.get(field_index) else {
        return Ok(None);
    };
    let aggregate = parse_aggregate(payload[24]);
    let (caption, _) = parser::wide_str(payload, 25)?;
    let mut measure = PivotMeasure::new(field.name.clone(), aggregate);
    if !caption.is_empty() {
        measure.name = Some(caption);
    }
    Ok(Some(measure))
}

fn parse_pivot_style(payload: &[u8]) -> XlsbResult<PivotStyle> {
    if payload.len() < 6 {
        return Ok(PivotStyle::default());
    }
    let (name, _) = parser::wide_str(payload, 2)?;
    let mut style = PivotStyle::default();
    if !name.is_empty() {
        style.name = Some(name);
    }
    Ok(style)
}

fn parse_cache_sheet_source(payload: &[u8]) -> XlsbResult<Option<PivotSource>> {
    if payload.len() < 23 {
        return Ok(None);
    }
    let (sheet_name, consumed) = parser::wide_str(payload, 3)?;
    let range_offset = 3 + consumed;
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

fn parse_cache_field_name(payload: &[u8]) -> XlsbResult<String> {
    if payload.len() < 24 {
        return Ok(String::new());
    }
    parser::wide_str(payload, 20).map(|(name, _)| name)
}

fn parse_shared_item(record_type: u16, payload: &[u8]) -> XlsbResult<PivotValue> {
    Ok(match record_type {
        records::BRT_PCDI_MISSING => PivotValue::Blank,
        records::BRT_PCDI_BOOLEAN => {
            PivotValue::Boolean(payload.first().copied().unwrap_or(0) != 0)
        }
        records::BRT_PCDI_ERROR => {
            PivotValue::Error(parse_cell_error(payload.first().copied().unwrap_or(0x0F)))
        }
        records::BRT_PCDI_NUMBER if payload.len() >= 8 => {
            PivotValue::Number(parser::read_f64(payload, 0))
        }
        records::BRT_PCDI_STRING => {
            let (value, _) = parser::wide_str(payload, 0)?;
            PivotValue::String(value)
        }
        _ => PivotValue::Blank,
    })
}

fn parse_aggregate(code: u8) -> PivotAggregate {
    match code {
        0x02 => PivotAggregate::Count,
        0x03 => PivotAggregate::CountNumbers,
        0x04 => PivotAggregate::Average,
        0x05 => PivotAggregate::Max,
        0x06 => PivotAggregate::Min,
        0x07 => PivotAggregate::Product,
        0x08 => PivotAggregate::StdDev,
        0x09 => PivotAggregate::StdDevP,
        0x0A => PivotAggregate::Var,
        0x0B => PivotAggregate::VarP,
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

fn rels_path_for_part(part_path: &str) -> String {
    let (base_dir, file_name) = part_path.rsplit_once('/').unwrap_or(("", part_path));
    if base_dir.is_empty() {
        format!("_rels/{file_name}.rels")
    } else {
        format!("{base_dir}/_rels/{file_name}.rels")
    }
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
