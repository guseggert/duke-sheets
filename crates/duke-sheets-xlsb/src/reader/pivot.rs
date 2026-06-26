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
    PivotCalculatedField, PivotDateGroupUnit, PivotField, PivotFieldRef, PivotFilter,
    PivotGrouping, PivotMeasure, PivotRefreshStatus, PivotSource, PivotStyle, PivotTable,
    PivotValue, PivotValuesAxis,
};
use ssfmt::{date_serial::date_to_serial, DateSystem};

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
    formula: Option<String>,
    formula_tokens: Option<Vec<u8>>,
    pname_field_indexes: Vec<usize>,
    grouping: Option<PivotGrouping>,
    shared_items: Vec<PivotValue>,
}

#[derive(Debug, Clone)]
struct PivotFormulaText {
    text: String,
    precedence: u8,
}

impl PivotFormulaText {
    fn atom(text: String) -> Self {
        Self {
            text,
            precedence: 8,
        }
    }
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
    pivot.groupings = cache
        .fields
        .iter()
        .filter_map(|field| field.grouping.clone())
        .collect();
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
    let mut in_shared_items = false;

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
                current_field = Some(parse_cache_field(payload)?);
            }
            records::BRT_BEGIN_PCD_SHARED_ITEMS => {
                in_shared_items = true;
            }
            records::BRT_END_PCD_SHARED_ITEMS => {
                in_shared_items = false;
            }
            records::BRT_PCDI_MISSING
            | records::BRT_PCDI_BOOLEAN
            | records::BRT_PCDI_ERROR
            | records::BRT_PCDI_DATETIME
            | records::BRT_PCDI_NUMBER
            | records::BRT_PCDI_STRING => {
                if in_shared_items {
                    if let Some(field) = &mut current_field {
                        field.shared_items.push(parse_shared_item(typ, payload)?);
                    }
                }
            }
            records::BRT_BEGIN_PCDFG_RANGE => {
                if let Some(field) = &mut current_field {
                    field.grouping = parse_pivot_group_range(&field.name, payload);
                }
            }
            records::BRT_BEGIN_PNAME => {
                if let Some(field) = &mut current_field {
                    if let Some(field_index) = parse_pname(payload) {
                        field.pname_field_indexes.push(field_index);
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

fn parse_cache_field(payload: &[u8]) -> XlsbResult<PivotCacheField> {
    if payload.len() < 24 {
        return Ok(PivotCacheField {
            name: String::new(),
            formula: None,
            formula_tokens: None,
            pname_field_indexes: Vec::new(),
            grouping: None,
            shared_items: Vec::new(),
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
        shared_items: Vec::new(),
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

fn parse_pname(payload: &[u8]) -> Option<usize> {
    if payload.len() < 6 {
        return None;
    }
    Some(parser::read_u32(payload, 0) as usize)
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
    let mut stack: Vec<PivotFormulaText> = Vec::new();
    let mut pos = 0usize;
    while pos < tokens.len() {
        let token = tokens[pos];
        pos += 1;
        match token {
            0x03 => push_pivot_binary(&mut stack, "+", 3, false)?,
            0x04 => push_pivot_binary(&mut stack, "-", 3, true)?,
            0x05 => push_pivot_binary(&mut stack, "*", 4, false)?,
            0x06 => push_pivot_binary(&mut stack, "/", 4, true)?,
            0x07 => push_pivot_binary(&mut stack, "^", 5, true)?,
            0x08 => push_pivot_binary(&mut stack, "&", 2, false)?,
            0x09 => push_pivot_binary(&mut stack, "<", 1, true)?,
            0x0A => push_pivot_binary(&mut stack, "<=", 1, true)?,
            0x0B => push_pivot_binary(&mut stack, "=", 1, true)?,
            0x0C => push_pivot_binary(&mut stack, ">=", 1, true)?,
            0x0D => push_pivot_binary(&mut stack, ">", 1, true)?,
            0x0E => push_pivot_binary(&mut stack, "<>", 1, true)?,
            0x12 => push_pivot_prefix(&mut stack, "+")?,
            0x13 => push_pivot_prefix(&mut stack, "-")?,
            0x14 => push_pivot_percent(&mut stack)?,
            0x15 => push_pivot_paren(&mut stack)?,
            0x17 => {
                let value = read_pivot_short_string(tokens, &mut pos)?;
                stack.push(PivotFormulaText::atom(format!(
                    "\"{}\"",
                    value.replace('"', "\"\"")
                )));
            }
            0x18 => {
                if pos + 5 > tokens.len() {
                    return None;
                }
                let subtype = tokens[pos];
                pos += 1;
                if subtype != 0x1D {
                    return None;
                }
                let pname_index = u32::from_le_bytes([
                    tokens[pos],
                    tokens[pos + 1],
                    tokens[pos + 2],
                    tokens[pos + 3],
                ]) as usize;
                pos += 4;
                let field_index = *pname_field_indexes.get(pname_index)?;
                let field_name = field_names.get(field_index)?;
                stack.push(PivotFormulaText::atom(pivot_formula_field_reference(
                    field_name,
                )));
            }
            0x1C => {
                if pos >= tokens.len() {
                    return None;
                }
                stack.push(PivotFormulaText::atom(format_formula_error(tokens[pos])));
                pos += 1;
            }
            0x1D => {
                if pos >= tokens.len() {
                    return None;
                }
                stack.push(PivotFormulaText::atom(if tokens[pos] != 0 {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }));
                pos += 1;
            }
            0x1E => {
                if pos + 2 > tokens.len() {
                    return None;
                }
                let value = u16::from_le_bytes([tokens[pos], tokens[pos + 1]]);
                pos += 2;
                stack.push(PivotFormulaText::atom(value.to_string()));
            }
            0x1F => {
                if pos + 8 > tokens.len() {
                    return None;
                }
                let value = f64::from_le_bytes(tokens[pos..pos + 8].try_into().ok()?);
                pos += 8;
                stack.push(PivotFormulaText::atom(format_formula_number(value)));
            }
            _ => return None,
        }
    }

    (stack.len() == 1).then(|| stack.pop().unwrap().text)
}

fn read_pivot_short_string(data: &[u8], offset: &mut usize) -> Option<String> {
    if *offset + 2 > data.len() {
        return None;
    }
    let char_count = data[*offset] as usize;
    *offset += 1;
    let flags = data[*offset];
    *offset += 1;
    if flags & 0x01 != 0 {
        let byte_len = char_count.checked_mul(2)?;
        if *offset + byte_len > data.len() {
            return None;
        }
        let units = data[*offset..*offset + byte_len]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        *offset += byte_len;
        Some(String::from_utf16_lossy(&units))
    } else {
        if *offset + char_count > data.len() {
            return None;
        }
        let text = data[*offset..*offset + char_count]
            .iter()
            .map(|&byte| byte as char)
            .collect::<String>();
        *offset += char_count;
        Some(text)
    }
}

fn push_pivot_binary(
    stack: &mut Vec<PivotFormulaText>,
    op: &str,
    precedence: u8,
    parenthesize_equal_right: bool,
) -> Option<()> {
    let right = stack.pop()?;
    let left = stack.pop()?;
    let left_text = if left.precedence < precedence {
        format!("({})", left.text)
    } else {
        left.text
    };
    let right_text = if right.precedence < precedence
        || (parenthesize_equal_right && right.precedence == precedence)
    {
        format!("({})", right.text)
    } else {
        right.text
    };
    stack.push(PivotFormulaText {
        text: format!("{left_text}{op}{right_text}"),
        precedence,
    });
    Some(())
}

fn push_pivot_prefix(stack: &mut Vec<PivotFormulaText>, op: &str) -> Option<()> {
    let operand = stack.pop()?;
    let text = if operand.precedence < 6 {
        format!("{op}({})", operand.text)
    } else {
        format!("{op}{}", operand.text)
    };
    stack.push(PivotFormulaText {
        text,
        precedence: 6,
    });
    Some(())
}

fn push_pivot_percent(stack: &mut Vec<PivotFormulaText>) -> Option<()> {
    let operand = stack.pop()?;
    let text = if operand.precedence < 7 {
        format!("({})%", operand.text)
    } else {
        format!("{}%", operand.text)
    };
    stack.push(PivotFormulaText {
        text,
        precedence: 7,
    });
    Some(())
}

fn push_pivot_paren(stack: &mut Vec<PivotFormulaText>) -> Option<()> {
    let operand = stack.pop()?;
    stack.push(PivotFormulaText {
        text: format!("({})", operand.text),
        precedence: 8,
    });
    Some(())
}

fn pivot_formula_field_reference(name: &str) -> String {
    if is_simple_pivot_formula_name(name) {
        name.to_string()
    } else {
        format!("[{}]", name.replace(']', "]]"))
    }
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
        records::BRT_PCDI_DATETIME if payload.len() >= 8 => {
            let year = parser::read_u16(payload, 0) as i32;
            let month = parser::read_u16(payload, 2) as u32;
            let day = payload[4] as u32;
            let hour = payload[5] as f64;
            let minute = payload[6] as f64;
            let second = payload[7] as f64;
            let serial = date_to_serial(year, month, day, DateSystem::Date1900)
                + ((hour * 3600.0) + (minute * 60.0) + second) / 86_400.0;
            PivotValue::Number(serial)
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
