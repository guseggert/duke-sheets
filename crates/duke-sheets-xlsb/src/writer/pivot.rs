use std::collections::{HashMap, HashSet};
use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, ptg, records, RecordWriter};
use crate::error::{XlsbError, XlsbResult};
use duke_sheets_core::{
    CellError, CellRange, CellValue, PivotAggregate, PivotDateGroupUnit, PivotField, PivotFilter,
    PivotGrouping, PivotManualGroup, PivotTable, PivotValue, PivotValuesAxis, Workbook,
};
use duke_sheets_formula::ast::{BinaryOperator, UnaryOperator};
use duke_sheets_formula::FormulaExpr;
use duke_sheets_pivot::{
    FormatPivotCache, FormatPivotCacheField, FormatPivotPlan, FormatPivotSource, FormatPivotTable,
};
use ssfmt::{
    date_serial::{serial_to_date, serial_to_time},
    DateSystem,
};

pub(crate) const RT_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
pub(crate) const RT_PIVOT_CACHE_DEFINITION: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
pub(crate) const RT_PIVOT_CACHE_RECORDS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
pub(crate) const RT_BINARY_INDEX: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex";

pub(crate) const CT_PIVOT_TABLE: &str = "application/vnd.ms-excel.pivotTable";
pub(crate) const CT_PIVOT_CACHE_DEFINITION: &str = "application/vnd.ms-excel.pivotCacheDefinition";
pub(crate) const CT_PIVOT_CACHE_RECORDS: &str = "application/vnd.ms-excel.pivotCacheRecords";
pub(crate) const CT_BINARY_INDEX: &str = "application/vnd.ms-excel.binIndexWs";

#[derive(Debug)]
struct XlsbPivotGroupingInfo<'a> {
    grouping: &'a PivotGrouping,
    date_unit: Option<PivotDateGroupUnit>,
    source_items: Vec<PivotValue>,
    source_item_ids: Vec<u32>,
    group_items: Vec<PivotValue>,
    base_item_group_ids: Vec<u32>,
    group_item_ids: Vec<u32>,
    source_item_kind: PivotCacheItemKind,
}

#[derive(Debug)]
struct XlsbPivotCacheLayout {
    fields: Vec<XlsbPivotCacheFieldLayout>,
    source_to_layout: Vec<usize>,
    manual_derived_by_base: HashMap<usize, usize>,
    date_derived_by_base: HashMap<usize, Vec<usize>>,
}

#[derive(Debug)]
enum XlsbPivotCacheFieldLayout {
    Source {
        source_index: usize,
        manual_parent: Option<usize>,
    },
    ManualDerived {
        base_source_index: usize,
        grouping_info_index: usize,
        name: String,
    },
    DateUnitDerived {
        base_source_index: usize,
        grouping_info_index: usize,
        name: String,
    },
}

pub(crate) fn sheet_has_pivots(plan: &FormatPivotPlan, sheet_index: usize) -> bool {
    plan.tables
        .iter()
        .any(|part| part.sheet_index == sheet_index)
}

pub(crate) fn sheet_pivot_parts(
    plan: &FormatPivotPlan,
    sheet_index: usize,
) -> impl Iterator<Item = &FormatPivotTable> {
    plan.tables
        .iter()
        .filter(move |part| part.sheet_index == sheet_index)
}

pub(crate) fn write_pivot_parts<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    workbook: &Workbook,
    plan: &FormatPivotPlan,
) -> XlsbResult<()> {
    for cache in &plan.caches {
        write_pivot_cache_definition_part(zip, options, workbook, plan, cache)?;
        write_pivot_cache_records_part(zip, options, workbook, plan, cache)?;
        write_pivot_cache_definition_rels(zip, options, cache.cache_num)?;
    }

    for part in &plan.tables {
        let cache = cache_for_table(plan, part)?;
        write_pivot_table_part(zip, options, workbook, plan, part, cache)?;
        write_pivot_table_rels(zip, options, part)?;
    }

    for sheet_index in 0..workbook.sheet_count() {
        if sheet_has_pivots(plan, sheet_index) {
            write_binary_index_part(zip, options, sheet_index)?;
        }
    }

    Ok(())
}

fn write_pivot_cache_definition_part<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    let path = format!("xl/pivotCache/pivotCacheDefinition{}.bin", cache.cache_num);
    zip.start_file(path, *options)?;

    let usage = cache_field_usage(workbook, plan, cache)?;
    let groupings = groupings_for_cache(workbook, plan, cache)?;
    validate_xlsb_pivot_groupings(cache, groupings)?;
    let date_system = workbook_date_system(workbook.settings().date_1904);
    let grouping_infos = xlsb_pivot_grouping_infos(workbook, cache, groupings, date_system)?;
    let layout = build_xlsb_pivot_cache_layout(cache, &grouping_infos)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    write_ac_block(&mut rw, 10, true)?;
    write_begin_pivot_cache_def(&mut rw, cache)?;
    write_pivot_cache_source(&mut rw, cache)?;

    rw.write_record(
        records::BRT_BEGIN_PCD_FIELDS,
        &(layout.fields.len() as u32).to_le_bytes(),
    )?;
    for (layout_index, field_layout) in layout.fields.iter().enumerate() {
        match field_layout {
            XlsbPivotCacheFieldLayout::Source {
                source_index,
                manual_parent,
            } => {
                let field = &cache.fields[*source_index];
                let grouping = grouping_for_field(groupings, &field.name);
                let grouping_info = grouping_info_for_field(&grouping_infos, &field.name);
                let item_kind = grouping_info
                    .map(|info| info.source_item_kind)
                    .unwrap_or(PivotCacheItemKind::Normal);
                let shared_items = grouping_info
                    .map(|info| info.source_items.as_slice())
                    .unwrap_or(&field.shared_items);
                let pnames = write_pcd_field(&mut rw, field, cache)?;
                write_pcd_shared_items(
                    &mut rw,
                    shared_items,
                    usage.store_items[*source_index],
                    item_kind,
                    date_system,
                )?;
                if let Some(parent_index) = manual_parent {
                    write_pcd_field_base_group(&mut rw, *parent_index)?;
                } else if let Some(parent_index) = layout
                    .date_derived_by_base
                    .get(source_index)
                    .and_then(|derived| derived.first())
                {
                    write_pcd_field_base_group(&mut rw, *parent_index)?;
                } else if let Some(grouping) =
                    grouping.filter(|grouping| !is_multi_unit_date_grouping(Some(*grouping)))
                {
                    write_pcd_field_group(
                        &mut rw,
                        None,
                        Some(layout_index),
                        grouping,
                        grouping_info,
                        date_system,
                        true,
                    )?;
                }
                write_pnames(&mut rw, &pnames)?;
            }
            XlsbPivotCacheFieldLayout::ManualDerived {
                base_source_index,
                grouping_info_index,
                name,
            } => {
                let info = &grouping_infos[*grouping_info_index];
                let base_layout_index = layout.source_to_layout[*base_source_index];
                let pnames = write_pcd_field_header(&mut rw, name, false, None, cache)?;
                write_pcd_field_group(
                    &mut rw,
                    None,
                    Some(base_layout_index),
                    info.grouping,
                    Some(info),
                    date_system,
                    true,
                )?;
                write_pnames(&mut rw, &pnames)?;
            }
            XlsbPivotCacheFieldLayout::DateUnitDerived {
                base_source_index,
                grouping_info_index,
                name,
            } => {
                let info = &grouping_infos[*grouping_info_index];
                let base_layout_index = layout.source_to_layout[*base_source_index];
                let pnames = write_pcd_field_header(&mut rw, name, false, None, cache)?;
                write_pcd_field_group(
                    &mut rw,
                    None,
                    Some(base_layout_index),
                    info.grouping,
                    Some(info),
                    date_system,
                    true,
                )?;
                write_pnames(&mut rw, &pnames)?;
            }
        }
        rw.write_record(records::BRT_END_PCD_FIELD, &[])?;
    }
    rw.write_record(records::BRT_END_PCD_FIELDS, &[])?;

    write_cache_definition_ext(&mut rw)?;
    rw.write_record(records::BRT_END_PIVOT_CACHE_DEF, &[])?;

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}

fn write_pivot_cache_records_part<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    let path = format!("xl/pivotCache/pivotCacheRecords{}.bin", cache.cache_num);
    zip.start_file(path, *options)?;

    let usage = cache_field_usage(workbook, plan, cache)?;
    let groupings = groupings_for_cache(workbook, plan, cache)?;
    validate_xlsb_pivot_groupings(cache, groupings)?;
    let date_system = workbook_date_system(workbook.settings().date_1904);
    let grouping_infos = xlsb_pivot_grouping_infos(workbook, cache, groupings, date_system)?;
    let layout = build_xlsb_pivot_cache_layout(cache, &grouping_infos)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);
    rw.write_record(
        records::BRT_BEGIN_PCD_RECORDS,
        &(cache.row_count as u32).to_le_bytes(),
    )?;

    if cache.save_data {
        for row in 0..cache.row_count {
            let mut payload = Vec::new();
            for (field_index, field) in cache.fields.iter().enumerate() {
                if matches!(
                    layout.fields.get(field_index),
                    Some(XlsbPivotCacheFieldLayout::DateUnitDerived { .. })
                ) {
                    continue;
                }
                let grouping_info = grouping_info_for_field(&grouping_infos, &field.name);
                let item_id = grouping_info
                    .and_then(|info| info.source_item_ids.get(row).copied())
                    .or_else(|| field.item_ids.get(row).copied())
                    .unwrap_or(0);
                let value = if let Some(info) = grouping_info {
                    info.source_items
                        .get(item_id as usize)
                        .unwrap_or(&PivotValue::Blank)
                } else {
                    field
                        .shared_items
                        .get(item_id as usize)
                        .unwrap_or(&PivotValue::Blank)
                };
                let store_items = usage.store_items[field_index];
                if store_items {
                    payload.extend_from_slice(&item_id.to_le_bytes());
                } else {
                    write_inline_record_value(&mut payload, value);
                }
            }
            rw.write_record(records::BRT_PCD_RECORD, &payload)?;
        }
    }

    rw.write_record(records::BRT_END_PCD_RECORDS, &[])?;
    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}

fn write_pivot_table_part<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    part: &FormatPivotTable,
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    let worksheet = workbook
        .worksheet(part.sheet_index)
        .ok_or_else(|| XlsbError::InvalidFormat("pivot table sheet not found".into()))?;
    let pivot = worksheet
        .pivot_tables()
        .get(part.pivot_index)
        .ok_or_else(|| XlsbError::InvalidFormat("pivot table not found".into()))?;
    let usage = cache_field_usage(workbook, plan, cache)?;
    let groupings = groupings_for_cache(workbook, plan, cache)?;
    let date_system = workbook_date_system(workbook.settings().date_1904);
    let grouping_infos = xlsb_pivot_grouping_infos(workbook, cache, groupings, date_system)?;
    let layout = build_xlsb_pivot_cache_layout(cache, &grouping_infos)?;

    let path = format!("xl/pivotTables/pivotTable{}.bin", part.table_num);
    zip.start_file(path, *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    write_ac_block(&mut rw, 7, false)?;
    write_begin_sx_view(&mut rw, pivot)?;
    write_sx_location(&mut rw, pivot, cache, &layout, &grouping_infos)?;
    rw.write_record(records::BRT_END_SX_LOCATION, &[])?;
    write_sx_fields(&mut rw, pivot, cache, &usage, &layout, &grouping_infos)?;
    let values_on_rows = values_field_on_axis(pivot, PivotValuesAxis::Rows);
    let values_on_columns = values_field_on_axis(pivot, PivotValuesAxis::Columns);
    write_axis_fields(
        &mut rw,
        records::BRT_BEGIN_ISXVD_RWS,
        records::BRT_END_ISXVD_RWS,
        cache,
        &layout,
        &pivot.rows,
        values_on_rows,
        pivot.layout.values_axis_position,
    )?;
    write_axis_items(
        &mut rw,
        records::BRT_BEGIN_SX_ROW_ITEMS,
        records::BRT_END_SX_ROW_ITEMS,
        cache,
        &layout,
        &pivot.rows,
        values_on_rows,
        pivot.layout.values_axis_position,
        pivot.measures.len(),
        &grouping_infos,
    )?;
    write_axis_fields(
        &mut rw,
        records::BRT_BEGIN_ISXVD_COLS,
        records::BRT_END_ISXVD_COLS,
        cache,
        &layout,
        &pivot.columns,
        values_on_columns,
        pivot.layout.values_axis_position,
    )?;
    write_axis_items(
        &mut rw,
        records::BRT_BEGIN_SX_COL_ITEMS,
        records::BRT_END_SX_COL_ITEMS,
        cache,
        &layout,
        &pivot.columns,
        values_on_columns,
        pivot.layout.values_axis_position,
        pivot.measures.len(),
        &grouping_infos,
    )?;
    write_page_fields(&mut rw, pivot, cache)?;
    write_data_fields(&mut rw, pivot, cache)?;
    write_sx_style(&mut rw, pivot)?;
    rw.write_record(records::BRT_END_SXVIEW, &[])?;

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}

fn write_pivot_table_rels<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    part: &FormatPivotTable,
) -> XlsbResult<()> {
    let path = format!("xl/pivotTables/_rels/pivotTable{}.bin.rels", part.table_num);
    zip.start_file(path, *options)?;
    let target = format!("../pivotCache/pivotCacheDefinition{}.bin", part.cache_num);
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
         <Relationship Id=\"rId1\" Type=\"{}\" Target=\"{}\"/>\
         </Relationships>",
        RT_PIVOT_CACHE_DEFINITION, target
    );
    zip.write_all(xml.as_bytes())?;
    Ok(())
}

fn write_pivot_cache_definition_rels<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    cache_num: usize,
) -> XlsbResult<()> {
    let path = format!(
        "xl/pivotCache/_rels/pivotCacheDefinition{}.bin.rels",
        cache_num
    );
    zip.start_file(path, *options)?;
    let target = format!("pivotCacheRecords{}.bin", cache_num);
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
         <Relationship Id=\"rId1\" Type=\"{}\" Target=\"{}\"/>\
         </Relationships>",
        RT_PIVOT_CACHE_RECORDS, target
    );
    zip.write_all(xml.as_bytes())?;
    Ok(())
}

fn groupings_for_cache<'a>(
    workbook: &'a Workbook,
    plan: &'a FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsbResult<&'a [PivotGrouping]> {
    let Some(part) = plan
        .tables
        .iter()
        .find(|part| part.cache_num == cache.cache_num)
    else {
        return Ok(&[]);
    };
    let worksheet = workbook
        .worksheet(part.sheet_index)
        .ok_or_else(|| XlsbError::InvalidFormat("pivot table sheet not found".into()))?;
    let pivot = worksheet
        .pivot_tables()
        .get(part.pivot_index)
        .ok_or_else(|| XlsbError::InvalidFormat("pivot table not found".into()))?;
    Ok(&pivot.groupings)
}

fn validate_xlsb_pivot_groupings(
    cache: &FormatPivotCache,
    groupings: &[PivotGrouping],
) -> XlsbResult<()> {
    let mut grouped_fields = HashSet::new();
    for grouping in groupings {
        let field_name = grouping_field_name(grouping);
        if cache.field_index(field_name).is_none() {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot grouping references unknown cache field: {field_name}"
            )));
        }
        if !grouped_fields.insert(field_name.to_lowercase()) {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot cache has more than one grouping for field {field_name}"
            )));
        }

        match grouping {
            PivotGrouping::Number {
                start,
                end,
                interval,
                ..
            } => {
                if !interval.is_finite() || *interval <= 0.0 {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot grouping for field {field_name} has an invalid interval"
                    )));
                }
                if start.is_some_and(|value| !value.is_finite())
                    || end.is_some_and(|value| !value.is_finite())
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot grouping for field {field_name} has a non-finite bound"
                    )));
                }
                if let (Some(start), Some(end)) = (*start, *end) {
                    if end < start {
                        return Err(XlsbError::InvalidFormat(format!(
                            "XLSB pivot grouping for field {field_name} has end before start"
                        )));
                    }
                }
            }
            PivotGrouping::Date { units, .. } => {
                if units.is_empty() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot date grouping has no units: {field_name}"
                    )));
                }
                let mut seen_units = HashSet::new();
                for unit in units {
                    if !seen_units.insert(*unit) {
                        return Err(XlsbError::InvalidFormat(format!(
                            "XLSB pivot date grouping has a duplicate unit for field {field_name}"
                        )));
                    }
                }
            }
            PivotGrouping::Manual { groups, .. } => {
                validate_xlsb_manual_grouping(field_name, groups)?;
            }
        }
    }
    Ok(())
}

fn validate_xlsb_manual_grouping(field_name: &str, groups: &[PivotManualGroup]) -> XlsbResult<()> {
    if groups.is_empty() {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot manual grouping for field {field_name} has no groups"
        )));
    }

    let mut group_names = HashSet::new();
    let mut members = HashSet::new();
    for group in groups {
        if group.name.trim().is_empty() {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot manual grouping for field {field_name} has a blank group name"
            )));
        }
        if group.members.is_empty() {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot manual group {} has no members",
                group.name
            )));
        }
        if !group_names.insert(group.name.to_lowercase()) {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot manual grouping for field {field_name} has duplicate group name {}",
                group.name
            )));
        }
        for member in &group.members {
            if !members.insert(member.clone()) {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot manual grouping for field {field_name} assigns item {member} to more than one group"
                )));
            }
        }
    }
    Ok(())
}

fn grouping_for_field<'a>(
    groupings: &'a [PivotGrouping],
    field_name: &str,
) -> Option<&'a PivotGrouping> {
    groupings
        .iter()
        .find(|grouping| grouping_field_name(grouping).eq_ignore_ascii_case(field_name))
}

fn grouping_field_name(grouping: &PivotGrouping) -> &str {
    match grouping {
        PivotGrouping::Number { field, .. }
        | PivotGrouping::Date { field, .. }
        | PivotGrouping::Manual { field, .. } => &field.name,
    }
}

fn xlsb_pivot_grouping_infos<'a>(
    workbook: &Workbook,
    cache: &FormatPivotCache,
    groupings: &'a [PivotGrouping],
    date_system: DateSystem,
) -> XlsbResult<Vec<XlsbPivotGroupingInfo<'a>>> {
    let mut infos = Vec::new();
    for grouping in groupings {
        match grouping {
            PivotGrouping::Date { units, .. } => {
                for unit in units {
                    infos.push(xlsb_pivot_grouping_info(
                        workbook,
                        cache,
                        grouping,
                        Some(*unit),
                        date_system,
                    )?);
                }
            }
            PivotGrouping::Number { .. } | PivotGrouping::Manual { .. } => {
                infos.push(xlsb_pivot_grouping_info(
                    workbook,
                    cache,
                    grouping,
                    None,
                    date_system,
                )?);
            }
        }
    }
    Ok(infos)
}

fn xlsb_pivot_grouping_info<'a>(
    workbook: &Workbook,
    cache: &FormatPivotCache,
    grouping: &'a PivotGrouping,
    date_unit: Option<PivotDateGroupUnit>,
    date_system: DateSystem,
) -> XlsbResult<XlsbPivotGroupingInfo<'a>> {
    let field_name = grouping_field_name(grouping);
    let field_index = cache.field_index(field_name).ok_or_else(|| {
        XlsbError::InvalidFormat(format!(
            "XLSB pivot grouping references unknown cache field: {field_name}"
        ))
    })?;
    let (sheet_index, range, source_col) = grouping_source_range(cache, field_index, field_name)?;
    let worksheet = workbook
        .worksheet(sheet_index)
        .ok_or_else(|| XlsbError::InvalidFormat("pivot source worksheet not found".into()))?;

    if let PivotGrouping::Manual { groups, .. } = grouping {
        let (source_items, source_item_ids) =
            manual_group_source_items(worksheet, range, source_col, field_name)?;
        let (group_items, base_item_group_ids) =
            manual_group_items_and_ids(field_name, &source_items, groups)?;
        let group_item_ids = source_item_ids
            .iter()
            .map(|item_id| {
                base_item_group_ids
                    .get(*item_id as usize)
                    .copied()
                    .ok_or_else(|| {
                        XlsbError::InvalidFormat(format!(
                            "XLSB pivot manual grouping for field {field_name} has an out-of-range source item index"
                        ))
                    })
            })
            .collect::<XlsbResult<Vec<_>>>()?;
        return Ok(XlsbPivotGroupingInfo {
            grouping,
            date_unit,
            source_items,
            source_item_ids,
            group_items,
            base_item_group_ids,
            group_item_ids,
            source_item_kind: PivotCacheItemKind::Normal,
        });
    }

    let mut source_lookup: HashMap<u64, u32> = HashMap::new();
    let mut source_items = Vec::new();
    let mut source_item_ids = Vec::new();
    let mut source_values = Vec::new();

    for row in range.start.row.saturating_add(1)..=range.end.row {
        let CellValue::Number(value) = worksheet.get_value_at(row, source_col) else {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot grouping for field {field_name} requires numeric source values"
            )));
        };
        if !value.is_finite() {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot grouping for field {field_name} has non-finite source value: {value}"
            )));
        }
        if matches!(grouping, PivotGrouping::Date { .. })
            && !valid_pcdi_datetime(value, date_system)
        {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot date grouping for field {field_name} has invalid date serial: {value}"
            )));
        }

        let key = value.to_bits();
        let item_id = if let Some(item_id) = source_lookup.get(&key) {
            *item_id
        } else {
            let item_id = checked_u32(source_items.len(), "pivot grouped source item index")?;
            source_lookup.insert(key, item_id);
            source_items.push(PivotValue::Number(value));
            item_id
        };
        source_item_ids.push(item_id);
        source_values.push(value);
    }

    let (group_items, group_item_ids, source_item_kind) = match grouping {
        PivotGrouping::Number {
            start,
            end,
            interval,
            ..
        } => {
            let (items, ids) =
                numeric_group_items_and_ids(&source_values, *start, *end, *interval, field_name)?;
            (items, ids, PivotCacheItemKind::Normal)
        }
        PivotGrouping::Date { units, .. } => {
            let unit = date_unit
                .or_else(|| units.first().copied())
                .ok_or_else(|| {
                    XlsbError::InvalidFormat(format!(
                        "XLSB pivot date grouping has no unit: {}",
                        grouping_field_name(grouping)
                    ))
                })?;
            let (items, ids) =
                date_group_items_and_ids(&source_values, unit, date_system, field_name)?;
            (items, ids, PivotCacheItemKind::DateTime)
        }
        PivotGrouping::Manual { .. } => unreachable!("manual pivot groupings return early"),
    };

    Ok(XlsbPivotGroupingInfo {
        grouping,
        date_unit,
        source_items,
        source_item_ids,
        group_items,
        base_item_group_ids: Vec::new(),
        group_item_ids,
        source_item_kind,
    })
}

fn manual_group_items_and_ids(
    field_name: &str,
    source_items: &[PivotValue],
    groups: &[PivotManualGroup],
) -> XlsbResult<(Vec<PivotValue>, Vec<u32>)> {
    let mut member_to_group = HashMap::new();
    for group in groups {
        for member in &group.members {
            if !source_items.iter().any(|item| item == member) {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot manual group {} references item not found in field {field_name}: {member}",
                    group.name
                )));
            }
            member_to_group.insert(member.clone(), group.name.clone());
        }
    }

    let mut group_items = Vec::new();
    let mut ungrouped_item_indexes = HashMap::new();
    for item in source_items {
        if member_to_group.contains_key(item) {
            continue;
        }
        let index = checked_u32(group_items.len(), "pivot manual ungrouped item index")?;
        ungrouped_item_indexes.insert(item.clone(), index);
        group_items.push(item.clone());
    }

    let mut group_name_indexes = HashMap::new();
    for group in groups {
        let index = checked_u32(group_items.len(), "pivot manual group item index")?;
        group_name_indexes.insert(group.name.clone(), index);
        group_items.push(PivotValue::String(group.name.clone()));
    }

    let base_item_group_ids = source_items
        .iter()
        .map(|item| {
            if let Some(group_name) = member_to_group.get(item) {
                group_name_indexes.get(group_name).copied().ok_or_else(|| {
                    XlsbError::InvalidFormat(format!(
                        "XLSB pivot manual grouping for field {field_name} could not map group {group_name}"
                    ))
                })
            } else {
                ungrouped_item_indexes.get(item).copied().ok_or_else(|| {
                    XlsbError::InvalidFormat(format!(
                        "XLSB pivot manual grouping for field {field_name} could not map ungrouped item {item}"
                    ))
                })
            }
        })
        .collect::<XlsbResult<Vec<_>>>()?;

    Ok((group_items, base_item_group_ids))
}

fn manual_group_source_items(
    worksheet: &duke_sheets_core::worksheet::Worksheet,
    range: CellRange,
    source_col: u16,
    field_name: &str,
) -> XlsbResult<(Vec<PivotValue>, Vec<u32>)> {
    let mut source_items = Vec::new();
    let mut source_item_ids = Vec::new();
    let mut lookup = HashMap::new();
    for row in range.start.row.saturating_add(1)..=range.end.row {
        let cell_value = worksheet.get_value_at(row, source_col);
        let value = PivotValue::from_cell_value(&cell_value);
        let item_id = if let Some(item_id) = lookup.get(&value) {
            *item_id
        } else {
            let item_id = checked_u32(source_items.len(), "pivot manual source item index")?;
            lookup.insert(value.clone(), item_id);
            source_items.push(value);
            item_id
        };
        source_item_ids.push(item_id);
    }
    if source_items.is_empty() {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot manual grouping for field {field_name} has no source items"
        )));
    }
    Ok((source_items, source_item_ids))
}

fn grouping_source_range(
    cache: &FormatPivotCache,
    field_index: usize,
    field_name: &str,
) -> XlsbResult<(usize, CellRange, u16)> {
    let FormatPivotSource::Worksheet {
        sheet_index, range, ..
    } = &cache.source
    else {
        return Err(XlsbError::InvalidFormat(
            "XLSB pivot grouping requires worksheet-range source data".into(),
        ));
    };
    let source_col = u32::from(range.start.col)
        .checked_add(field_index as u32)
        .ok_or_else(|| {
            XlsbError::InvalidFormat("XLSB pivot grouping field index overflow".into())
        })?;
    if source_col > u16::MAX as u32 || source_col > u32::from(range.end.col) {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot grouping for field {field_name} must reference a source field"
        )));
    }
    Ok((*sheet_index, *range, source_col as u16))
}

fn date_group_items_and_ids(
    source_values: &[f64],
    unit: PivotDateGroupUnit,
    date_system: DateSystem,
    field_name: &str,
) -> XlsbResult<(Vec<PivotValue>, Vec<u32>)> {
    let mut group_items = Vec::new();
    let mut group_lookup = HashMap::new();
    let mut group_item_ids = Vec::with_capacity(source_values.len());

    for value in source_values {
        let group_value = date_group_item_value(*value, unit, date_system)?;
        let key = group_value.to_bits();
        let item_id = if let Some(item_id) = group_lookup.get(&key) {
            *item_id
        } else {
            let item_id = checked_u32(group_items.len(), "pivot date group item index")?;
            group_lookup.insert(key, item_id);
            group_items.push(PivotValue::Number(group_value));
            item_id
        };
        group_item_ids.push(item_id);
    }

    if group_items.is_empty() {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot date grouping for field {field_name} has no source dates"
        )));
    }
    Ok((group_items, group_item_ids))
}

fn numeric_group_items_and_ids(
    source_values: &[f64],
    start: Option<f64>,
    end: Option<f64>,
    interval: f64,
    field_name: &str,
) -> XlsbResult<(Vec<PivotValue>, Vec<u32>)> {
    let (source_min, source_max) = numeric_min_max(source_values).ok_or_else(|| {
        XlsbError::InvalidFormat(format!(
            "XLSB pivot numeric grouping for field {field_name} has no source values"
        ))
    })?;
    let effective_start = start.unwrap_or(source_min);
    let effective_end = end.unwrap_or(source_max);
    if effective_end < effective_start {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot numeric grouping for field {field_name} has end before start"
        )));
    }

    let mut group_items = Vec::new();
    let has_underflow = start.is_some();
    let has_overflow = end.is_some();
    if has_underflow {
        group_items.push(PivotValue::String(format!(
            "<{}",
            format_group_number(effective_start)
        )));
    }

    let all_integer_bins = is_integral_number(effective_start)
        && is_integral_number(effective_end)
        && is_integral_number(interval);
    let mut current = effective_start;
    let mut bin_count = 0usize;
    while current <= effective_end {
        if bin_count >= 1_048_576 {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot numeric grouping for field {field_name} has too many bins"
            )));
        }
        let next = current + interval;
        if !next.is_finite() || next <= current {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot numeric grouping for field {field_name} has an invalid interval"
            )));
        }
        let upper = if next >= effective_end {
            effective_end
        } else if all_integer_bins {
            next - 1.0
        } else {
            next
        };
        group_items.push(PivotValue::String(format!(
            "{}-{}",
            format_group_number(current),
            format_group_number(upper)
        )));
        bin_count += 1;
        if next >= effective_end {
            break;
        }
        current = next;
    }

    if has_overflow {
        group_items.push(PivotValue::String(format!(
            ">{}",
            format_group_number(effective_end)
        )));
    }

    let mut group_item_ids = Vec::with_capacity(source_values.len());
    let first_bin_index = if has_underflow { 1usize } else { 0usize };
    let last_index = group_items.len().saturating_sub(1);
    let last_bin_index = if has_overflow {
        last_index.saturating_sub(1)
    } else {
        last_index
    };
    for value in source_values {
        let item_id = if has_underflow && *value < effective_start {
            0
        } else if has_overflow && *value > effective_end {
            last_index
        } else {
            let offset = ((*value - effective_start) / interval).floor();
            let offset = if offset.is_finite() && offset >= 0.0 {
                offset as usize
            } else {
                0
            };
            first_bin_index.saturating_add(offset).min(last_bin_index)
        };
        group_item_ids.push(checked_u32(item_id, "pivot numeric group item index")?);
    }

    Ok((group_items, group_item_ids))
}

fn numeric_min_max(values: &[f64]) -> Option<(f64, f64)> {
    let mut min = None::<f64>;
    let mut max = None::<f64>;
    for value in values {
        min = Some(min.map_or(*value, |current| current.min(*value)));
        max = Some(max.map_or(*value, |current| current.max(*value)));
    }
    min.zip(max)
}

fn is_integral_number(value: f64) -> bool {
    value.is_finite() && value.fract() == 0.0
}

fn format_group_number(value: f64) -> String {
    if is_integral_number(value) && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        return format!("{}", value as i64);
    }
    value.to_string()
}

fn grouping_info_for_field<'a, 'b>(
    infos: &'a [XlsbPivotGroupingInfo<'b>],
    field_name: &str,
) -> Option<&'a XlsbPivotGroupingInfo<'b>> {
    infos
        .iter()
        .find(|info| grouping_field_name(info.grouping).eq_ignore_ascii_case(field_name))
}

fn grouping_info_for_field_and_unit<'a, 'b>(
    infos: &'a [XlsbPivotGroupingInfo<'b>],
    field_name: &str,
    unit: PivotDateGroupUnit,
) -> Option<(usize, &'a XlsbPivotGroupingInfo<'b>)> {
    infos.iter().enumerate().find(|(_, info)| {
        grouping_field_name(info.grouping).eq_ignore_ascii_case(field_name)
            && info.date_unit == Some(unit)
    })
}

fn is_multi_unit_date_grouping(grouping: Option<&PivotGrouping>) -> bool {
    matches!(grouping, Some(PivotGrouping::Date { units, .. }) if units.len() > 1)
}

fn build_xlsb_pivot_cache_layout(
    cache: &FormatPivotCache,
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> XlsbResult<XlsbPivotCacheLayout> {
    let mut fields = cache
        .fields
        .iter()
        .enumerate()
        .map(|(source_index, _)| XlsbPivotCacheFieldLayout::Source {
            source_index,
            manual_parent: None,
        })
        .collect::<Vec<_>>();
    let source_to_layout = (0..cache.fields.len()).collect::<Vec<_>>();
    let mut manual_derived_by_base = HashMap::new();
    let mut date_derived_by_base: HashMap<usize, Vec<usize>> = HashMap::new();
    let mut names = cache
        .fields
        .iter()
        .map(|field| field.name.clone())
        .collect::<Vec<_>>();

    for (grouping_info_index, info) in grouping_infos.iter().enumerate() {
        if !matches!(info.grouping, PivotGrouping::Manual { .. }) {
            continue;
        }
        let field_name = grouping_field_name(info.grouping);
        let base_source_index = cache.field_index(field_name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "XLSB pivot manual grouping references unknown cache field: {field_name}"
            ))
        })?;
        let derived_index = fields.len();
        if let Some(XlsbPivotCacheFieldLayout::Source { manual_parent, .. }) =
            fields.get_mut(source_to_layout[base_source_index])
        {
            *manual_parent = Some(derived_index);
        }
        manual_derived_by_base.insert(base_source_index, derived_index);
        let name = unique_manual_grouped_header(&names, field_name);
        names.push(name.clone());
        fields.push(XlsbPivotCacheFieldLayout::ManualDerived {
            base_source_index,
            grouping_info_index,
            name,
        });
    }

    let mut date_grouped_fields = HashSet::new();
    let mut claimed_date_fields = HashSet::new();
    for grouping_info in grouping_infos {
        let PivotGrouping::Date { units, .. } = grouping_info.grouping else {
            continue;
        };
        if units.len() <= 1 {
            continue;
        }

        let field_name = grouping_field_name(grouping_info.grouping);
        if !date_grouped_fields.insert(field_name.to_lowercase()) {
            continue;
        }
        let base_source_index = cache.field_index(field_name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "XLSB pivot date grouping references unknown cache field: {field_name}"
            ))
        })?;
        for unit in units {
            let Some((grouping_info_index, _)) =
                grouping_info_for_field_and_unit(grouping_infos, field_name, *unit)
            else {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot date grouping info missing for field {field_name}"
                )));
            };
            let layout_index = find_date_grouped_field_index(
                cache,
                base_source_index,
                field_name,
                *unit,
                &claimed_date_fields,
            )
            .ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "XLSB pivot date grouping field missing for {field_name} {}",
                    date_group_unit_caption(*unit)
                ))
            })?;
            claimed_date_fields.insert(layout_index);
            let name = cache.fields[layout_index].name.clone();
            fields[layout_index] = XlsbPivotCacheFieldLayout::DateUnitDerived {
                base_source_index,
                grouping_info_index,
                name,
            };
            date_derived_by_base
                .entry(base_source_index)
                .or_default()
                .push(layout_index);
        }
    }

    Ok(XlsbPivotCacheLayout {
        fields,
        source_to_layout,
        manual_derived_by_base,
        date_derived_by_base,
    })
}

fn unique_manual_grouped_header(existing_names: &[String], field_name: &str) -> String {
    for suffix in 2usize.. {
        let candidate = format!("{field_name}{suffix}");
        if existing_names
            .iter()
            .all(|name| !name.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
    }
    unreachable!("unbounded suffix search should always return")
}

fn find_date_grouped_field_index(
    cache: &FormatPivotCache,
    base_source_index: usize,
    field_name: &str,
    unit: PivotDateGroupUnit,
    claimed: &HashSet<usize>,
) -> Option<usize> {
    let base = date_grouped_header(field_name, unit);
    (1usize..).find_map(|suffix| {
        let candidate = if suffix == 1 {
            base.clone()
        } else {
            format!("{base} {suffix}")
        };
        cache.fields.iter().enumerate().find_map(|(index, field)| {
            (index != base_source_index
                && !claimed.contains(&index)
                && field.name.eq_ignore_ascii_case(&candidate))
            .then_some(index)
        })
    })
}

fn date_grouped_header(field_name: &str, unit: PivotDateGroupUnit) -> String {
    format!("{field_name} ({})", date_group_unit_caption(unit))
}

fn date_group_unit_caption(unit: PivotDateGroupUnit) -> &'static str {
    match unit {
        PivotDateGroupUnit::Seconds => "Seconds",
        PivotDateGroupUnit::Minutes => "Minutes",
        PivotDateGroupUnit::Hours => "Hours",
        PivotDateGroupUnit::Days => "Days",
        PivotDateGroupUnit::Months => "Months",
        PivotDateGroupUnit::Quarters => "Quarters",
        PivotDateGroupUnit::Years => "Years",
    }
}

fn write_binary_index_part<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    sheet_index: usize,
) -> XlsbResult<()> {
    let path = format!("xl/worksheets/binaryIndex{}.bin", sheet_index + 1);
    zip.start_file(path, *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    rw.write_record(
        0x002A,
        &[
            0x00, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x1A, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ],
    )?;
    rw.write_record(
        0x0028,
        &[
            0x07, 0x00, 0x00, 0x00, 0x9F, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x01, 0x00, 0x01, 0x00, 0x2D, 0x00, 0x00, 0x00, 0x73, 0x00, 0x00, 0x00, 0xB9, 0x00,
            0x00, 0x00,
        ],
    )?;
    rw.write_record(0x0115, &[])?;

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}

fn write_ac_block<W: Write>(
    rw: &mut RecordWriter<W>,
    feature: u32,
    cache_shape: bool,
) -> std::io::Result<()> {
    rw.write_record(
        records::BRT_BEGIN_AC_BLOCKS,
        &[0x01, 0x00, 0x00, 0x10, 0x00, 0x80],
    )?;
    let mut payload = Vec::new();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&feature.to_le_bytes());
    if cache_shape {
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
        payload.extend_from_slice(&1u32.to_le_bytes());
    } else {
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
    }
    rw.write_record(0x0C00, &payload)?;
    rw.write_record(records::BRT_END_AC_BLOCKS, &[])
}

fn write_begin_pivot_cache_def<W: Write>(
    rw: &mut RecordWriter<W>,
    cache: &FormatPivotCache,
) -> std::io::Result<()> {
    let mut payload = Vec::new();
    payload.push(8);
    payload.push(3);
    payload.push(8);
    let mut flags = 0x11u8;
    if cache.refresh_on_load {
        flags |= 0x04;
    }
    payload.push(flags);
    let ghost_limit = cache
        .missing_items_limit
        .map(|limit| limit as i32)
        .unwrap_or(-1);
    payload.extend_from_slice(&ghost_limit.to_le_bytes());
    payload.extend_from_slice(&0f64.to_le_bytes());
    payload.push(0x02);
    payload.extend_from_slice(&(cache.row_count as u32).to_le_bytes());
    payload.extend_from_slice(&encode_wide_str("rId1"));
    payload.extend_from_slice(&0u32.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_PIVOT_CACHE_DEF, &payload)
}

fn write_pivot_cache_source<W: Write>(
    rw: &mut RecordWriter<W>,
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    let FormatPivotSource::Worksheet {
        sheet_name, range, ..
    } = &cache.source
    else {
        return Err(XlsbError::InvalidFormat(
            "XLSB pivot cache writing currently requires worksheet or table sources".into(),
        ));
    };

    rw.write_record(records::BRT_BEGIN_PCD_SOURCE, &[0; 8])?;
    let mut payload = Vec::new();
    payload.extend_from_slice(&[0x00, 0x00, 0x02]);
    payload.extend_from_slice(&encode_wide_str(sheet_name));
    write_unchecked_rfx(&mut payload, *range);
    rw.write_record(records::BRT_BEGIN_PCDS_SHEET, &payload)?;
    rw.write_record(records::BRT_END_PCDS_SHEET, &[])?;
    rw.write_record(records::BRT_END_PCD_SOURCE, &[])?;
    Ok(())
}

fn write_pcd_field<W: Write>(
    rw: &mut RecordWriter<W>,
    field: &FormatPivotCacheField,
    cache: &FormatPivotCache,
) -> XlsbResult<Vec<usize>> {
    write_pcd_field_header(
        rw,
        &field.name,
        field.database_field,
        field.formula.as_deref(),
        cache,
    )
}

fn write_pcd_field_header<W: Write>(
    rw: &mut RecordWriter<W>,
    name: &str,
    database_field: bool,
    formula: Option<&str>,
    cache: &FormatPivotCache,
) -> XlsbResult<Vec<usize>> {
    let mut payload = Vec::new();
    let mut flags: u16 = if database_field { 0x0004 } else { 0x0000 };
    if formula.is_some() {
        flags |= 0x0100;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(name));
    let pnames = if let Some(formula) = formula {
        write_pivot_parsed_formula(&mut payload, formula, cache)?
    } else {
        Vec::new()
    };
    rw.write_record(records::BRT_BEGIN_PCD_FIELD, &payload)?;
    Ok(pnames)
}

fn write_pivot_parsed_formula(
    payload: &mut Vec<u8>,
    formula: &str,
    cache: &FormatPivotCache,
) -> XlsbResult<Vec<usize>> {
    let trimmed = formula.trim();
    let parse_input = if trimmed.starts_with('=') {
        trimmed.to_string()
    } else {
        format!("={trimmed}")
    };
    let expr = duke_sheets_formula::parse_formula(&parse_input).map_err(|err| {
        XlsbError::InvalidFormat(format!(
            "XLSB pivot calculated field formula could not be parsed: {err}"
        ))
    })?;

    let mut rgce = Vec::new();
    let mut pnames = Vec::new();
    compile_pivot_formula_expr(&expr, &mut rgce, &mut pnames, cache).map_err(|_| {
        XlsbError::InvalidFormat(format!(
            "XLSB pivot calculated field formula uses unsupported syntax: {formula}"
        ))
    })?;

    payload.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
    payload.extend_from_slice(&rgce);
    payload.extend_from_slice(&0u32.to_le_bytes());
    Ok(pnames)
}

fn write_pnames<W: Write>(rw: &mut RecordWriter<W>, field_indexes: &[usize]) -> XlsbResult<()> {
    if field_indexes.is_empty() {
        return Ok(());
    }
    rw.write_record(
        records::BRT_BEGIN_PNAMES,
        &(field_indexes.len() as u32).to_le_bytes(),
    )?;
    for field_index in field_indexes {
        let field_index = checked_u32(*field_index, "pivot calculated formula source field index")?;
        let mut payload = Vec::with_capacity(6);
        payload.extend_from_slice(&field_index.to_le_bytes());
        payload.push(0xFF);
        payload.push(0x00);
        rw.write_record(records::BRT_BEGIN_PNAME, &payload)?;
        rw.write_record(records::BRT_END_PNAME, &[])?;
    }
    rw.write_record(records::BRT_END_PNAMES, &[])?;
    Ok(())
}

fn compile_pivot_formula_expr(
    expr: &FormulaExpr,
    out: &mut Vec<u8>,
    pnames: &mut Vec<usize>,
    cache: &FormatPivotCache,
) -> Result<(), ()> {
    match expr {
        FormulaExpr::Number(n) => emit_pivot_number(*n, out),
        FormulaExpr::String(s) => emit_pivot_string(s, out),
        FormulaExpr::Boolean(b) => {
            out.push(ptg::PTG_BOOL);
            out.push(if *b { 1 } else { 0 });
            Ok(())
        }
        FormulaExpr::Error(e) => {
            out.push(ptg::PTG_ERR);
            out.push(error_code(*e));
            Ok(())
        }
        FormulaExpr::NameRef(name) => emit_pivot_sxname(name, out, pnames, cache),
        FormulaExpr::StructuredRef(reference) => {
            let Some(column) = reference.column.as_deref() else {
                return Err(());
            };
            emit_pivot_sxname(column, out, pnames, cache)
        }
        FormulaExpr::BinaryOp { op, left, right } => {
            compile_pivot_formula_expr(left, out, pnames, cache)?;
            compile_pivot_formula_expr(right, out, pnames, cache)?;
            out.push(match op {
                BinaryOperator::Add => ptg::PTG_ADD,
                BinaryOperator::Subtract => ptg::PTG_SUB,
                BinaryOperator::Multiply => ptg::PTG_MUL,
                BinaryOperator::Divide => ptg::PTG_DIV,
                BinaryOperator::Power => ptg::PTG_POWER,
                BinaryOperator::Concat => ptg::PTG_CONCAT,
                BinaryOperator::LessThan => ptg::PTG_LT,
                BinaryOperator::LessEqual => ptg::PTG_LE,
                BinaryOperator::Equal => ptg::PTG_EQ,
                BinaryOperator::GreaterEqual => ptg::PTG_GE,
                BinaryOperator::GreaterThan => ptg::PTG_GT,
                BinaryOperator::NotEqual => ptg::PTG_NE,
                BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect => {
                    return Err(());
                }
            });
            Ok(())
        }
        FormulaExpr::UnaryOp { op, operand } => {
            compile_pivot_formula_expr(operand, out, pnames, cache)?;
            out.push(match op {
                UnaryOperator::Plus => ptg::PTG_UPLUS,
                UnaryOperator::Negate => ptg::PTG_UMINUS,
                UnaryOperator::Percent => ptg::PTG_PERCENT,
                UnaryOperator::Paren => ptg::PTG_PAREN,
                UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange => {
                    return Err(());
                }
            });
            Ok(())
        }
        FormulaExpr::CellRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::Function { .. }
        | FormulaExpr::ExternalFunction { .. }
        | FormulaExpr::Array(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => Err(()),
    }
}

fn emit_pivot_number(value: f64, out: &mut Vec<u8>) -> Result<(), ()> {
    if value.is_finite() && value.fract() == 0.0 && (0.0..=u16::MAX as f64).contains(&value) {
        out.push(ptg::PTG_INT);
        out.extend_from_slice(&(value as u16).to_le_bytes());
    } else {
        out.push(ptg::PTG_NUM);
        out.extend_from_slice(&value.to_le_bytes());
    }
    Ok(())
}

fn emit_pivot_string(value: &str, out: &mut Vec<u8>) -> Result<(), ()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > u8::MAX as usize {
        return Err(());
    }
    out.push(ptg::PTG_STR);
    out.push(units.len() as u8);
    let high_byte = units.iter().any(|&unit| unit > 0xFF);
    out.push(if high_byte { 0x01 } else { 0x00 });
    if high_byte {
        for unit in units {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    } else {
        for unit in units {
            out.push(unit as u8);
        }
    }
    Ok(())
}

fn emit_pivot_sxname(
    field_name: &str,
    out: &mut Vec<u8>,
    pnames: &mut Vec<usize>,
    cache: &FormatPivotCache,
) -> Result<(), ()> {
    let field_index = cache.field_index(field_name).ok_or(())?;
    let pname_index =
        checked_u32(pnames.len(), "pivot calculated formula PName index").map_err(|_| ())?;
    pnames.push(field_index);
    out.extend_from_slice(&[0x18, 0x1D]);
    out.extend_from_slice(&pname_index.to_le_bytes());
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PivotCacheItemKind {
    Normal,
    DateTime,
}

fn write_pcd_shared_items<W: Write>(
    rw: &mut RecordWriter<W>,
    values: &[PivotValue],
    store_items: bool,
    item_kind: PivotCacheItemKind,
    date_system: DateSystem,
) -> XlsbResult<()> {
    let stats = SharedItemStats::from_values(values);
    let mut payload = Vec::new();
    payload.extend_from_slice(&stats.flags(item_kind).to_le_bytes());
    let stored_count = if store_items { values.len() as u32 } else { 0 };
    payload.extend_from_slice(&stored_count.to_le_bytes());
    if stats.has_number {
        payload.extend_from_slice(&stats.min_number.unwrap_or(0.0).to_le_bytes());
        payload.extend_from_slice(&stats.max_number.unwrap_or(0.0).to_le_bytes());
    }
    rw.write_record(records::BRT_BEGIN_PCD_SHARED_ITEMS, &payload)?;
    if store_items {
        for item in values {
            write_pcdi_item(rw, item, item_kind, date_system)?;
        }
    }
    rw.write_record(records::BRT_END_PCD_SHARED_ITEMS, &[])?;
    Ok(())
}

fn write_pcd_field_group<W: Write>(
    rw: &mut RecordWriter<W>,
    parent_index: Option<usize>,
    base_index: Option<usize>,
    grouping: &PivotGrouping,
    grouping_info: Option<&XlsbPivotGroupingInfo<'_>>,
    date_system: DateSystem,
    include_group_items: bool,
) -> XlsbResult<()> {
    let parent_index = checked_optional_i32(parent_index, "pivot grouped parent field index")?;
    let base_index = checked_optional_i32(base_index, "pivot grouped base field index")?;
    let mut group_payload = Vec::with_capacity(8);
    group_payload.extend_from_slice(&parent_index.to_le_bytes());
    group_payload.extend_from_slice(&base_index.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_PCDF_GROUP, &group_payload)?;

    match grouping {
        PivotGrouping::Number {
            start,
            end,
            interval,
            ..
        } => {
            let mut range_payload = Vec::with_capacity(26);
            range_payload.push(0x00);
            let mut flags = 0u8;
            if start.is_none() {
                flags |= 0x01;
            }
            if end.is_none() {
                flags |= 0x02;
            }
            range_payload.push(flags);
            range_payload.extend_from_slice(&start.unwrap_or(0.0).to_le_bytes());
            range_payload.extend_from_slice(&end.unwrap_or(0.0).to_le_bytes());
            range_payload.extend_from_slice(&interval.to_le_bytes());
            rw.write_record(records::BRT_BEGIN_PCDFG_RANGE, &range_payload)?;
            rw.write_record(records::BRT_END_PCDFG_RANGE, &[])?;
        }
        PivotGrouping::Date { units, .. } => {
            let mut range_payload = Vec::with_capacity(26);
            let info = grouping_info.ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "XLSB pivot date grouping info missing for field {}",
                    grouping_field_name(grouping)
                ))
            })?;
            let unit = info
                .date_unit
                .or_else(|| units.first().copied())
                .ok_or_else(|| {
                    XlsbError::InvalidFormat(format!(
                        "XLSB pivot date grouping has no unit: {}",
                        grouping_field_name(grouping)
                    ))
                })?;
            let (start, end) = source_item_min_max(&info.source_items).ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "XLSB pivot date grouping for field {} has no source dates",
                    grouping_field_name(grouping)
                ))
            })?;
            range_payload.push(xlsb_date_group_by(unit));
            range_payload.push(0x07);
            range_payload.extend_from_slice(&start.to_le_bytes());
            range_payload.extend_from_slice(&end.to_le_bytes());
            range_payload.extend_from_slice(&1.0f64.to_le_bytes());
            rw.write_record(records::BRT_BEGIN_PCDFG_RANGE, &range_payload)?;
            rw.write_record(records::BRT_END_PCDFG_RANGE, &[])?;
        }
        PivotGrouping::Manual { .. } => {
            let info = grouping_info.ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "XLSB pivot manual grouping info missing for field {}",
                    grouping_field_name(grouping)
                ))
            })?;
            let discrete_count = checked_u32(
                info.base_item_group_ids.len(),
                "pivot manual discrete item count",
            )?;
            rw.write_record(
                records::BRT_BEGIN_PCDFG_DISCRETE,
                &discrete_count.to_le_bytes(),
            )?;
            for item_id in &info.base_item_group_ids {
                rw.write_record(records::BRT_PCDI_INDEX, &item_id.to_le_bytes())?;
            }
            rw.write_record(records::BRT_END_PCDFG_DISCRETE, &[])?;
        }
    }

    if include_group_items {
        let group_items = pivot_group_items(grouping, grouping_info)?;
        rw.write_record(
            records::BRT_BEGIN_PCDFG_ITEMS,
            &(group_items.len() as u32).to_le_bytes(),
        )?;
        for item in &group_items {
            write_pcdi_item(rw, item, PivotCacheItemKind::Normal, date_system)?;
        }
        rw.write_record(records::BRT_END_PCDFG_ITEMS, &[])?;
    }
    rw.write_record(records::BRT_END_PCDF_GROUP, &[])?;
    Ok(())
}

fn write_pcd_field_base_group<W: Write>(
    rw: &mut RecordWriter<W>,
    parent_index: usize,
) -> XlsbResult<()> {
    let parent_index = checked_i32(parent_index, "pivot manual parent field index")?;
    let mut group_payload = Vec::with_capacity(8);
    group_payload.extend_from_slice(&parent_index.to_le_bytes());
    group_payload.extend_from_slice(&(-1i32).to_le_bytes());
    rw.write_record(records::BRT_BEGIN_PCDF_GROUP, &group_payload)?;
    rw.write_record(records::BRT_END_PCDF_GROUP, &[])?;
    Ok(())
}

fn checked_optional_i32(value: Option<usize>, what: &str) -> XlsbResult<i32> {
    value.map_or(Ok(-1), |value| checked_i32(value, what))
}

fn xlsb_date_group_by(unit: PivotDateGroupUnit) -> u8 {
    match unit {
        PivotDateGroupUnit::Seconds => 0x01,
        PivotDateGroupUnit::Minutes => 0x02,
        PivotDateGroupUnit::Hours => 0x03,
        PivotDateGroupUnit::Days => 0x04,
        PivotDateGroupUnit::Months => 0x05,
        PivotDateGroupUnit::Quarters => 0x06,
        PivotDateGroupUnit::Years => 0x07,
    }
}

fn source_item_min_max(items: &[PivotValue]) -> Option<(f64, f64)> {
    let mut min = None::<f64>;
    let mut max = None::<f64>;
    for item in items {
        let PivotValue::Number(value) = item else {
            continue;
        };
        min = Some(min.map_or(*value, |current| current.min(*value)));
        max = Some(max.map_or(*value, |current| current.max(*value)));
    }
    min.zip(max)
}

fn pivot_group_items(
    grouping: &PivotGrouping,
    grouping_info: Option<&XlsbPivotGroupingInfo<'_>>,
) -> XlsbResult<Vec<PivotValue>> {
    match grouping {
        PivotGrouping::Number { .. }
        | PivotGrouping::Date { .. }
        | PivotGrouping::Manual { .. } => grouping_info
            .map(|info| info.group_items.clone())
            .ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "XLSB pivot grouping info missing for field {}",
                    grouping_field_name(grouping)
                ))
            }),
    }
}

fn date_group_item_value(
    serial: f64,
    unit: PivotDateGroupUnit,
    date_system: DateSystem,
) -> XlsbResult<f64> {
    let Some((year, month, day)) = serial_to_date(serial, date_system) else {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot date grouping has invalid serial: {serial}"
        )));
    };
    let (hour, minute, second) = serial_to_time(serial);
    Ok(match unit {
        PivotDateGroupUnit::Years => year as f64,
        PivotDateGroupUnit::Quarters => ((month - 1) / 3 + 1) as f64,
        PivotDateGroupUnit::Months => month as f64,
        PivotDateGroupUnit::Days => day as f64,
        PivotDateGroupUnit::Hours => hour as f64,
        PivotDateGroupUnit::Minutes => minute as f64,
        PivotDateGroupUnit::Seconds => second as f64,
    })
}

fn write_pcdi_item<W: Write>(
    rw: &mut RecordWriter<W>,
    value: &PivotValue,
    item_kind: PivotCacheItemKind,
    date_system: DateSystem,
) -> XlsbResult<()> {
    match (item_kind, value) {
        (PivotCacheItemKind::DateTime, PivotValue::Number(serial)) => {
            write_pcdi_datetime(rw, *serial, date_system)?;
        }
        (PivotCacheItemKind::DateTime, _) => {
            return Err(XlsbError::InvalidFormat(
                "XLSB pivot datetime cache item requires a numeric serial".into(),
            ));
        }
        (_, PivotValue::Blank) => rw.write_record(records::BRT_PCDI_MISSING, &[])?,
        (_, PivotValue::Boolean(value)) => {
            rw.write_record(records::BRT_PCDI_BOOLEAN, &[if *value { 1 } else { 0 }])?
        }
        (_, PivotValue::Number(value)) => {
            rw.write_record(records::BRT_PCDI_NUMBER, &value.to_le_bytes())?
        }
        (_, PivotValue::String(value)) => {
            rw.write_record(records::BRT_PCDI_STRING, &encode_wide_str(value))?
        }
        (_, PivotValue::Error(value)) => {
            rw.write_record(records::BRT_PCDI_ERROR, &[error_code(*value)])?
        }
    }
    Ok(())
}

fn write_pcdi_datetime<W: Write>(
    rw: &mut RecordWriter<W>,
    serial: f64,
    date_system: DateSystem,
) -> XlsbResult<()> {
    let Some((year, month, day)) = serial_to_date(serial, date_system) else {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot datetime cache item has invalid serial: {serial}"
        )));
    };
    if !(1900..=9999).contains(&year) {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot datetime cache item year is out of range: {year}"
        )));
    }
    let (hour, minute, second) = serial_to_time(serial);
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(&(year as u16).to_le_bytes());
    payload.extend_from_slice(&(month as u16).to_le_bytes());
    payload.push(day as u8);
    payload.push(hour as u8);
    payload.push(minute as u8);
    payload.push(second as u8);
    rw.write_record(records::BRT_PCDI_DATETIME, &payload)?;
    Ok(())
}

fn valid_pcdi_datetime(serial: f64, date_system: DateSystem) -> bool {
    serial.is_finite()
        && serial_to_date(serial, date_system).is_some_and(|(year, month, day)| {
            (1900..=9999).contains(&year) && (1..=12).contains(&month) && day <= 31
        })
}

fn workbook_date_system(date_1904: bool) -> DateSystem {
    if date_1904 {
        DateSystem::Date1904
    } else {
        DateSystem::Date1900
    }
}

fn write_inline_record_value(out: &mut Vec<u8>, value: &PivotValue) {
    match value {
        PivotValue::Number(value) => out.extend_from_slice(&value.to_le_bytes()),
        PivotValue::Boolean(value) => out.push(if *value { 1 } else { 0 }),
        PivotValue::Error(value) => out.push(error_code(*value)),
        PivotValue::Blank => {}
        PivotValue::String(_) => out.extend_from_slice(&0u32.to_le_bytes()),
    }
}

fn write_cache_definition_ext<W: Write>(rw: &mut RecordWriter<W>) -> std::io::Result<()> {
    rw.write_record(records::BRT_BEGIN_FRT, &0x0000_0E02u32.to_le_bytes())?;
    rw.write_record(0x042A, &[0; 9])?;
    rw.write_record(0x042B, &[])?;
    rw.write_record(records::BRT_END_FRT, &[])
}

fn write_begin_sx_view<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
) -> std::io::Result<()> {
    let mut payload = Vec::new();
    let mut header = [
        0x00, 0x41, 0x40, 0x01, 0xF0, 0x64, 0x09, 0x00, 0xD9, 0x00, 0x00, 0x00, 0x02, 0x00, 0x08,
        0x03, 0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00,
    ];
    header[12] = values_axis_code(pivot.layout.values_axis);
    let values_position = pivot
        .layout
        .values_axis_position
        .map(|position| position.min(i32::MAX as u32) as i32)
        .unwrap_or(-1);
    header[16..20].copy_from_slice(&values_position.to_le_bytes());
    payload.extend_from_slice(&header);
    payload.extend_from_slice(&encode_wide_str(&pivot.name));
    let data_caption = if pivot.layout.data_caption.trim().is_empty() {
        "Values"
    } else {
        pivot.layout.data_caption.as_str()
    };
    payload.extend_from_slice(&encode_wide_str(data_caption));
    rw.write_record(records::BRT_BEGIN_SXVIEW, &payload)
}

fn write_sx_location<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> std::io::Result<()> {
    let range = pivot
        .rendered_range
        .unwrap_or_else(|| estimated_pivot_range(pivot, cache, layout, grouping_infos));
    let first_data_col = expanded_axis_field_count(cache, layout, &pivot.rows).max(1) as u32;
    let (page_rows, page_cols) = page_field_area_size(pivot);
    let mut payload = Vec::new();
    payload.extend_from_slice(&range.start.row.to_le_bytes());
    payload.extend_from_slice(&range.end.row.to_le_bytes());
    payload.extend_from_slice(&(range.start.col as u32).to_le_bytes());
    payload.extend_from_slice(&(range.end.col as u32).to_le_bytes());
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&first_data_col.to_le_bytes());
    payload.extend_from_slice(&page_rows.to_le_bytes());
    payload.extend_from_slice(&page_cols.to_le_bytes());
    rw.write_record(records::BRT_SX_LOCATION, &payload)
}

fn write_sx_fields<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    usage: &CacheFieldUsage,
    layout: &XlsbPivotCacheLayout,
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> XlsbResult<()> {
    rw.write_record(
        records::BRT_BEGIN_SXVDS,
        &(layout.fields.len() as u32).to_le_bytes(),
    )?;
    for field_layout in &layout.fields {
        match field_layout {
            XlsbPivotCacheFieldLayout::Source { source_index, .. } => {
                let field = &cache.fields[*source_index];
                let axis = if layout.date_derived_by_base.contains_key(source_index) {
                    0
                } else {
                    pivot_field_axis(pivot, &field.name)
                };
                write_sx_field(rw, axis)?;
                if usage.store_items[*source_index] {
                    let item_count = grouping_info_for_field(grouping_infos, &field.name)
                        .map(|info| info.source_items.len())
                        .unwrap_or(field.shared_items.len());
                    write_sx_field_items(rw, item_count)?;
                }
            }
            XlsbPivotCacheFieldLayout::ManualDerived {
                base_source_index,
                grouping_info_index,
                ..
            }
            | XlsbPivotCacheFieldLayout::DateUnitDerived {
                base_source_index,
                grouping_info_index,
                ..
            } => {
                let base_field = &cache.fields[*base_source_index];
                write_sx_field(rw, pivot_field_axis(pivot, &base_field.name))?;
                write_sx_field_items(rw, grouping_infos[*grouping_info_index].group_items.len())?;
            }
        }
        rw.write_record(records::BRT_END_SXVD, &[])?;
    }
    rw.write_record(records::BRT_END_SXVDS, &[])?;
    Ok(())
}

fn write_sx_field<W: Write>(rw: &mut RecordWriter<W>, axis: u8) -> std::io::Result<()> {
    let mut payload: [u8; 20] = if axis & 0x07 != 0 {
        [
            0x00, 0x01, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00, 0x5F, 0xB1, 0x04, 0x00, 0x0A, 0x00,
            0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
        ]
    } else {
        [
            0x00, 0x01, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00, 0x7F, 0x81, 0x04, 0x00, 0x0A, 0x00,
            0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
        ]
    };
    payload[0] = axis & 0x0F;
    rw.write_record(records::BRT_BEGIN_SXVD, &payload)
}

fn write_sx_field_items<W: Write>(
    rw: &mut RecordWriter<W>,
    item_count: usize,
) -> std::io::Result<()> {
    let count = item_count as u32 + 1;
    rw.write_record(records::BRT_BEGIN_SXVIS, &count.to_le_bytes())?;
    for item_index in 0..item_count as u32 {
        let mut payload = Vec::with_capacity(7);
        payload.extend_from_slice(&[0, 0, 0]);
        payload.extend_from_slice(&item_index.to_le_bytes());
        rw.write_record(records::BRT_BEGIN_SXVI, &payload)?;
        rw.write_record(records::BRT_END_SXVI, &[])?;
    }
    let mut default_payload = Vec::with_capacity(7);
    default_payload.extend_from_slice(&[1, 0, 0]);
    default_payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_SXVI, &default_payload)?;
    rw.write_record(records::BRT_END_SXVI, &[])?;
    rw.write_record(records::BRT_END_SXVIS, &[])
}

fn write_axis_fields<W: Write>(
    rw: &mut RecordWriter<W>,
    begin_record: u16,
    end_record: u16,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
    include_values_field: bool,
    values_position: Option<u32>,
) -> XlsbResult<()> {
    if fields.is_empty() && !include_values_field {
        return Ok(());
    }

    let mut indexes = expanded_axis_field_indexes(cache, layout, fields)?;
    if include_values_field {
        let position = values_position
            .map(|position| position as usize)
            .unwrap_or(indexes.len())
            .min(indexes.len());
        indexes.insert(position, -2);
    }

    let mut payload = Vec::new();
    payload.extend_from_slice(&(indexes.len() as u32).to_le_bytes());
    for index in indexes {
        payload.extend_from_slice(&index.to_le_bytes());
    }
    rw.write_record(begin_record, &payload)?;
    rw.write_record(end_record, &[])?;
    Ok(())
}

fn write_axis_items<W: Write>(
    rw: &mut RecordWriter<W>,
    begin_record: u16,
    end_record: u16,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
    include_values_field: bool,
    values_position: Option<u32>,
    measure_count: usize,
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> XlsbResult<()> {
    let line_tuples = axis_line_tuples(
        cache,
        layout,
        fields,
        include_values_field,
        values_position,
        measure_count,
        grouping_infos,
    )?;
    let grand_tuples = axis_grand_total_tuples(
        expanded_axis_field_count(cache, layout, fields),
        include_values_field,
        values_position,
        measure_count,
    );
    let count = line_tuples.len() + grand_tuples.len();
    rw.write_record(begin_record, &(count as u32).to_le_bytes())?;
    for (tuple, data_item) in &line_tuples {
        write_sxli(rw, tuple, false, *data_item)?;
    }
    for (tuple, data_item) in &grand_tuples {
        write_sxli(rw, tuple, true, *data_item)?;
    }
    rw.write_record(end_record, &[])?;
    Ok(())
}

fn axis_line_tuples(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
    include_values_field: bool,
    values_position: Option<u32>,
    measure_count: usize,
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> XlsbResult<Vec<(Vec<u32>, Option<u32>)>> {
    if fields.is_empty() {
        if include_values_field {
            return Ok(tuples_with_data_items(
                vec![Vec::new()],
                values_position,
                measure_count,
            ));
        }
        return Ok(vec![(Vec::new(), None)]);
    }

    let tuples = axis_item_tuples(cache, layout, fields, grouping_infos)?;
    if include_values_field {
        Ok(tuples_with_data_items(
            tuples,
            values_position,
            measure_count,
        ))
    } else {
        Ok(tuples.into_iter().map(|tuple| (tuple, None)).collect())
    }
}

fn axis_grand_total_tuples(
    field_count: usize,
    include_values_field: bool,
    values_position: Option<u32>,
    measure_count: usize,
) -> Vec<(Vec<u32>, Option<u32>)> {
    if field_count == 0 {
        return Vec::new();
    }

    let grand_tuple = vec![0; field_count];
    if include_values_field {
        tuples_with_data_items(vec![grand_tuple], values_position, measure_count)
    } else {
        vec![(grand_tuple, None)]
    }
}

fn tuples_with_data_items(
    tuples: Vec<Vec<u32>>,
    values_position: Option<u32>,
    measure_count: usize,
) -> Vec<(Vec<u32>, Option<u32>)> {
    let measure_count = measure_count.max(1);
    let mut expanded = Vec::with_capacity(tuples.len() * measure_count);
    for tuple in tuples {
        for data_item in 0..measure_count as u32 {
            let mut tuple = tuple.clone();
            let position = values_position
                .map(|position| position as usize)
                .unwrap_or(tuple.len())
                .min(tuple.len());
            tuple.insert(position, data_item);
            expanded.push((tuple, Some(data_item)));
        }
    }
    expanded
}

fn write_page_fields<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    if pivot.page_fields.is_empty() {
        return Ok(());
    }

    rw.write_record(
        records::BRT_BEGIN_SXPIS,
        &(pivot.page_fields.len() as u32).to_le_bytes(),
    )?;
    for field in &pivot.page_fields {
        let index = cache.field_index(&field.field.name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "pivot references unknown page field {}",
                field.field.name
            ))
        })?;
        let selected_item =
            selected_page_item_index(pivot, &field.field.name, &cache.fields[index])
                .unwrap_or(0x0010_00FE);
        let mut payload = Vec::new();
        payload.extend_from_slice(&(index as u32).to_le_bytes());
        payload.extend_from_slice(&selected_item.to_le_bytes());
        payload.extend_from_slice(&(-1i32).to_le_bytes());
        payload.push(0);
        rw.write_record(records::BRT_BEGIN_SXPI, &payload)?;
        rw.write_record(records::BRT_END_SXPI, &[])?;
    }
    rw.write_record(records::BRT_END_SXPIS, &[])?;
    Ok(())
}

fn write_sxli<W: Write>(
    rw: &mut RecordWriter<W>,
    item_indexes: &[u32],
    grand_total: bool,
    data_item: Option<u32>,
) -> std::io::Result<()> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&(if grand_total { 13u16 } else { 0u16 }).to_le_bytes());
    payload.extend_from_slice(&(item_indexes.len() as u32).to_le_bytes());
    payload.extend_from_slice(&data_item.unwrap_or(0).to_le_bytes());
    rw.write_record(records::BRT_BEGIN_SXLI, &payload)?;
    for item_index in item_indexes {
        rw.write_record(records::BRT_SXLI_ITEM, &item_index.to_le_bytes())?;
        rw.write_record(records::BRT_END_SXLI_ITEM, &[])?;
    }
    rw.write_record(records::BRT_END_SXLI, &[])
}

fn write_data_fields<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    rw.write_record(
        records::BRT_BEGIN_SXDIS,
        &(pivot.measures.len() as u32).to_le_bytes(),
    )?;
    for measure in &pivot.measures {
        let field_index = cache.field_index(&measure.field.name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "pivot references unknown measure field {}",
                measure.field.name
            ))
        })?;
        let mut payload = Vec::new();
        payload.extend_from_slice(&(field_index as u32).to_le_bytes());
        payload.extend_from_slice(&[0; 20]);
        payload.push(aggregate_code(measure.aggregate));
        payload.extend_from_slice(&encode_wide_str(&measure.caption()));
        rw.write_record(records::BRT_BEGIN_SXDI, &payload)?;
        rw.write_record(records::BRT_END_SXDI, &[])?;
    }
    rw.write_record(records::BRT_END_SXDIS, &[])?;
    Ok(())
}

fn write_sx_style<W: Write>(rw: &mut RecordWriter<W>, pivot: &PivotTable) -> std::io::Result<()> {
    let style_name = pivot
        .style
        .name
        .as_deref()
        .filter(|name| !name.trim().is_empty())
        .unwrap_or("PivotStyleMedium9");
    let mut payload = Vec::new();
    payload.extend_from_slice(&pivot_style_flags(&pivot.style).to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(style_name));
    rw.write_record(records::BRT_SX_VIEW_STYLE, &payload)
}

fn pivot_style_flags(style: &duke_sheets_core::PivotStyle) -> u16 {
    let mut flags = 0u16;
    if style.show_last_column {
        flags |= 0x02;
    }
    if style.show_row_stripes {
        flags |= 0x04;
    }
    if style.show_column_stripes {
        flags |= 0x08;
    }
    if style.show_row_headers {
        flags |= 0x10;
    }
    if style.show_column_headers {
        flags |= 0x20;
    }
    flags
}

fn estimated_pivot_range(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> CellRange {
    let (page_rows, _) = page_field_area_size(pivot);
    let body_start_row = pivot.target.row
        + if page_rows == 0 {
            0
        } else {
            page_rows.saturating_add(1)
        };
    let row_item_count = axis_item_count(cache, layout, &pivot.rows, grouping_infos).max(1);
    let row_header_count = pivot.columns.len() as u32 + 1;
    let row_count = row_header_count + row_item_count as u32 + 1;

    let col_item_count = axis_item_count(cache, layout, &pivot.columns, grouping_infos).max(1);
    let measure_count = pivot.measures.len().max(1);
    let value_col_count = col_item_count * measure_count;
    let row_field_count = expanded_axis_field_count(cache, layout, &pivot.rows).max(1);
    let col_count = row_field_count as u16 + value_col_count as u16;
    CellRange::from_indices(
        body_start_row,
        pivot.target.col,
        body_start_row + row_count.saturating_sub(1),
        pivot.target.col + col_count.saturating_sub(1),
    )
}

fn page_field_area_size(pivot: &PivotTable) -> (u32, u32) {
    let count = pivot.page_fields.len();
    if count == 0 {
        return (0, 0);
    }

    let wrap = pivot.layout.page_wrap as usize;
    let row_count = if wrap == 0 {
        count
    } else if pivot.layout.page_over_then_down {
        (count + wrap - 1) / wrap
    } else {
        wrap.min(count)
    };
    let col_count = if wrap == 0 {
        1
    } else if pivot.layout.page_over_then_down {
        wrap.min(count)
    } else {
        (count + row_count - 1) / row_count
    };
    (row_count as u32, col_count as u32)
}

fn axis_item_count(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> usize {
    if fields.is_empty() {
        return 1;
    }
    axis_item_tuples(cache, layout, fields, grouping_infos)
        .map(|tuples| tuples.len())
        .unwrap_or(1)
}

fn axis_item_tuples(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> XlsbResult<Vec<Vec<u32>>> {
    if fields.is_empty() {
        return Ok(Vec::new());
    }
    let indexes = expanded_axis_source_indexes(cache, layout, fields)?;

    let mut seen = HashSet::new();
    let mut tuples = Vec::new();
    for row in 0..cache.row_count {
        let tuple = indexes
            .iter()
            .map(|index| match index {
                ExpandedAxisSourceIndex::ManualDerived {
                    grouping_info_index,
                    ..
                }
                | ExpandedAxisSourceIndex::DateUnitDerived {
                    grouping_info_index,
                    ..
                } => grouping_infos[*grouping_info_index]
                    .group_item_ids
                    .get(row)
                    .copied()
                    .unwrap_or(0),
                ExpandedAxisSourceIndex::Source { source_index } => {
                    let field = &cache.fields[*source_index];
                    grouping_info_for_field(grouping_infos, &field.name)
                        .and_then(|info| info.source_item_ids.get(row).copied())
                        .or_else(|| field.item_ids.get(row).copied())
                        .unwrap_or(0)
                }
            })
            .collect::<Vec<_>>();
        if seen.insert(tuple.clone()) {
            tuples.push(tuple);
        }
    }
    Ok(tuples)
}

#[derive(Debug, Clone, Copy)]
enum ExpandedAxisSourceIndex {
    ManualDerived {
        layout_index: usize,
        grouping_info_index: usize,
    },
    DateUnitDerived {
        layout_index: usize,
        grouping_info_index: usize,
    },
    Source {
        source_index: usize,
    },
}

fn expanded_axis_field_indexes(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
) -> XlsbResult<Vec<i32>> {
    expanded_axis_source_indexes(cache, layout, fields)?
        .into_iter()
        .map(|index| {
            let layout_index = match index {
                ExpandedAxisSourceIndex::ManualDerived { layout_index, .. }
                | ExpandedAxisSourceIndex::DateUnitDerived { layout_index, .. } => layout_index,
                ExpandedAxisSourceIndex::Source { source_index } => {
                    layout.source_to_layout[source_index]
                }
            };
            checked_i32(layout_index, "pivot axis field index")
        })
        .collect()
}

fn expanded_axis_source_indexes(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
) -> XlsbResult<Vec<ExpandedAxisSourceIndex>> {
    let mut indexes = Vec::new();
    for field in fields {
        let source_index = cache.field_index(&field.field.name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "pivot references unknown axis field {}",
                field.field.name
            ))
        })?;
        if let Some(derived_index) = layout.manual_derived_by_base.get(&source_index) {
            let Some(XlsbPivotCacheFieldLayout::ManualDerived {
                grouping_info_index,
                ..
            }) = layout.fields.get(*derived_index)
            else {
                return Err(XlsbError::InvalidFormat(
                    "pivot manual derived axis field has invalid layout".into(),
                ));
            };
            indexes.push(ExpandedAxisSourceIndex::ManualDerived {
                layout_index: *derived_index,
                grouping_info_index: *grouping_info_index,
            });
            indexes.push(ExpandedAxisSourceIndex::Source { source_index });
            continue;
        }
        if let Some(derived_indexes) = layout.date_derived_by_base.get(&source_index) {
            for derived_index in derived_indexes {
                let Some(XlsbPivotCacheFieldLayout::DateUnitDerived {
                    grouping_info_index,
                    ..
                }) = layout.fields.get(*derived_index)
                else {
                    return Err(XlsbError::InvalidFormat(
                        "pivot date derived axis field has invalid layout".into(),
                    ));
                };
                indexes.push(ExpandedAxisSourceIndex::DateUnitDerived {
                    layout_index: *derived_index,
                    grouping_info_index: *grouping_info_index,
                });
            }
            continue;
        }
        indexes.push(ExpandedAxisSourceIndex::Source { source_index });
    }
    Ok(indexes)
}

fn expanded_axis_field_count(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
) -> usize {
    expanded_axis_source_indexes(cache, layout, fields)
        .map(|indexes| indexes.len())
        .unwrap_or(fields.len())
}

fn cache_for_table<'a>(
    plan: &'a FormatPivotPlan,
    part: &FormatPivotTable,
) -> XlsbResult<&'a FormatPivotCache> {
    plan.caches
        .iter()
        .find(|cache| cache.cache_num == part.cache_num)
        .ok_or_else(|| XlsbError::InvalidFormat("pivot cache part not found".into()))
}

#[derive(Debug, Clone)]
struct CacheFieldUsage {
    store_items: Vec<bool>,
}

fn cache_field_usage(
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsbResult<CacheFieldUsage> {
    let mut axis_names = HashSet::new();
    for part in plan
        .tables
        .iter()
        .filter(|part| part.cache_num == cache.cache_num)
    {
        let worksheet = workbook
            .worksheet(part.sheet_index)
            .ok_or_else(|| XlsbError::InvalidFormat("pivot table sheet not found".into()))?;
        let pivot = worksheet
            .pivot_tables()
            .get(part.pivot_index)
            .ok_or_else(|| XlsbError::InvalidFormat("pivot table not found".into()))?;
        for field in pivot
            .rows
            .iter()
            .chain(pivot.columns.iter())
            .chain(pivot.page_fields.iter())
        {
            axis_names.insert(field.field.name.to_lowercase());
        }
        for filter in &pivot.filters {
            if let Some(name) = filter_field_name(filter) {
                axis_names.insert(name.to_lowercase());
            }
        }
        for grouping in &pivot.groupings {
            axis_names.insert(grouping_field_name(grouping).to_lowercase());
        }
    }

    let store_items = cache
        .fields
        .iter()
        .map(|field| {
            let used_on_axis = axis_names.contains(&field.name.to_lowercase());
            used_on_axis
                || field
                    .shared_items
                    .iter()
                    .any(|value| !matches!(value, PivotValue::Number(_)))
        })
        .collect();
    Ok(CacheFieldUsage { store_items })
}

fn selected_page_item_index(
    pivot: &PivotTable,
    field_name: &str,
    field: &FormatPivotCacheField,
) -> Option<u32> {
    let PivotFilter::FieldItems { allowed_items, .. } = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems {
                field: filter_field,
                ..
            } if filter_field.name.eq_ignore_ascii_case(field_name)
        )
    })?
    else {
        return None;
    };

    let [item] = allowed_items.as_slice() else {
        return None;
    };
    field
        .shared_items
        .iter()
        .position(|candidate| candidate == item)
        .map(|index| index as u32)
}

fn pivot_field_axis(pivot: &PivotTable, field_name: &str) -> u8 {
    const SX_AXIS_ROW: u8 = 0x01;
    const SX_AXIS_COL: u8 = 0x02;
    const SX_AXIS_PAGE: u8 = 0x04;
    const SX_AXIS_DATA: u8 = 0x08;

    let mut axis = if pivot
        .rows
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        SX_AXIS_ROW
    } else if pivot
        .columns
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        SX_AXIS_COL
    } else if pivot
        .page_fields
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        SX_AXIS_PAGE
    } else {
        0
    };

    if pivot
        .measures
        .iter()
        .any(|measure| measure.field.name.eq_ignore_ascii_case(field_name))
    {
        axis |= SX_AXIS_DATA;
    }

    axis
}

fn values_field_on_axis(pivot: &PivotTable, axis: PivotValuesAxis) -> bool {
    pivot.measures.len() > 1
        && pivot.layout.values_axis == axis
        && (matches!(axis, PivotValuesAxis::Rows) || pivot.layout.values_axis_position.is_some())
}

fn values_axis_code(axis: PivotValuesAxis) -> u8 {
    match axis {
        PivotValuesAxis::Rows => 0x01,
        PivotValuesAxis::Columns => 0x02,
    }
}

fn filter_field_name(filter: &PivotFilter) -> Option<&str> {
    match filter {
        PivotFilter::FieldItems { field, .. }
        | PivotFilter::Label { field, .. }
        | PivotFilter::LabelBetween { field, .. }
        | PivotFilter::Date { field, .. }
        | PivotFilter::DateBetween { field, .. }
        | PivotFilter::DatePeriod { field, .. }
        | PivotFilter::Value { field, .. }
        | PivotFilter::ValueBetween { field, .. }
        | PivotFilter::TopN { field, .. } => Some(field.name.as_str()),
        PivotFilter::Unsupported { .. } => None,
    }
}

#[derive(Debug, Default)]
struct SharedItemStats {
    has_blank: bool,
    has_bool: bool,
    has_number: bool,
    has_string: bool,
    has_error: bool,
    min_number: Option<f64>,
    max_number: Option<f64>,
    all_numbers_integer: bool,
}

impl SharedItemStats {
    fn from_values(values: &[PivotValue]) -> Self {
        let mut stats = Self {
            all_numbers_integer: true,
            ..Self::default()
        };
        for value in values {
            match value {
                PivotValue::Blank => stats.has_blank = true,
                PivotValue::Boolean(_) => stats.has_bool = true,
                PivotValue::Number(value) => {
                    stats.has_number = true;
                    stats.all_numbers_integer &= value.fract() == 0.0;
                    stats.min_number = Some(stats.min_number.map_or(*value, |min| min.min(*value)));
                    stats.max_number = Some(stats.max_number.map_or(*value, |max| max.max(*value)));
                }
                PivotValue::String(_) => stats.has_string = true,
                PivotValue::Error(_) => stats.has_error = true,
            }
        }
        stats
    }

    fn flags(&self, item_kind: PivotCacheItemKind) -> u16 {
        let mut flags = 0x0400u16;
        if item_kind == PivotCacheItemKind::DateTime {
            flags = 0;
            if self.has_number {
                flags |= 0x0004 | 0x0100;
            }
            return flags;
        }
        if self.has_number {
            flags |= 0x0002 | 0x0040 | 0x0100;
            if self.all_numbers_integer {
                flags |= 0x0080;
            }
        }
        if self.has_blank {
            flags |= 0x0001 | 0x0002 | 0x0010;
        }
        if self.has_bool || self.has_error || self.has_string {
            flags |= 0x0001 | 0x0002 | 0x0008;
        }
        let kind_count = [
            self.has_blank,
            self.has_bool,
            self.has_number,
            self.has_string,
            self.has_error,
        ]
        .into_iter()
        .filter(|has| *has)
        .count();
        if kind_count > 1 {
            flags |= 0x0020;
        }
        flags
    }
}

fn aggregate_code(aggregate: PivotAggregate) -> u8 {
    match aggregate {
        PivotAggregate::Sum => 0x01,
        PivotAggregate::Count => 0x02,
        PivotAggregate::CountNumbers => 0x03,
        PivotAggregate::Average => 0x04,
        PivotAggregate::Max => 0x05,
        PivotAggregate::Min => 0x06,
        PivotAggregate::Product => 0x07,
        PivotAggregate::StdDev => 0x08,
        PivotAggregate::StdDevP => 0x09,
        PivotAggregate::Var => 0x0A,
        PivotAggregate::VarP => 0x0B,
    }
}

fn checked_u32(value: usize, label: &str) -> XlsbResult<u32> {
    if value > u32::MAX as usize {
        return Err(XlsbError::InvalidFormat(format!(
            "{label} exceeds XLSB limits"
        )));
    }
    Ok(value as u32)
}

fn checked_i32(value: usize, label: &str) -> XlsbResult<i32> {
    if value > i32::MAX as usize {
        return Err(XlsbError::InvalidFormat(format!(
            "{label} exceeds XLSB limits"
        )));
    }
    Ok(value as i32)
}

fn write_unchecked_rfx(out: &mut Vec<u8>, range: CellRange) {
    out.extend_from_slice(&range.start.row.to_le_bytes());
    out.extend_from_slice(&range.end.row.to_le_bytes());
    out.extend_from_slice(&(range.start.col as u32).to_le_bytes());
    out.extend_from_slice(&(range.end.col as u32).to_le_bytes());
}

fn error_code(error: CellError) -> u8 {
    match error {
        CellError::Null => 0x00,
        CellError::Div0 => 0x07,
        CellError::Value => 0x0F,
        CellError::Ref => 0x17,
        CellError::Name => 0x1D,
        CellError::Num => 0x24,
        CellError::Na => 0x2A,
        CellError::GettingData => 0x2B,
        CellError::Spill | CellError::Calc => 0x0F,
    }
}
