use std::collections::{HashMap, HashSet};
use std::io::{Seek, Write};

use chrono::Datelike;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, ptg, records, RecordWriter};
use crate::error::{XlsbError, XlsbResult};
use crate::writer::styles::StyleMapping;
use duke_sheets_core::style::NumberFormat;
use duke_sheets_core::{
    CellError, CellRange, CellValue, PivotAggregate, PivotDateGroupUnit, PivotDatePeriod,
    PivotField, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayoutKind, PivotManualGroup,
    PivotMeasure, PivotShowAs, PivotSort, PivotSourceRange, PivotSubtotal, PivotTable, PivotValue,
    PivotValuesAxis, Workbook, WorkbookConnection, WorkbookConnectionKind,
};
use duke_sheets_formula::ast::{BinaryOperator, CellReference, UnaryOperator};
use duke_sheets_formula::decompile::function_table::{
    function_index, function_is_biff8_addin, function_is_fixed_arity,
};
use duke_sheets_formula::FormulaExpr;
use duke_sheets_pivot::{
    FormatPivotCache, FormatPivotCacheField, FormatPivotPlan, FormatPivotSource, FormatPivotTable,
};
use ssfmt::{
    date_serial::{date_to_serial, serial_to_date, serial_to_time},
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
pub(crate) const RT_EXTERNAL_LINK_PATH: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

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
struct XlsbPivotAxisTuples {
    rows: Vec<Vec<u32>>,
    columns: Vec<Vec<u32>>,
}

enum XlsbVisibleRowIter<'a> {
    All(std::ops::Range<usize>),
    Filtered(std::iter::Copied<std::slice::Iter<'a, usize>>),
}

impl Iterator for XlsbVisibleRowIter<'_> {
    type Item = usize;

    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::All(rows) => rows.next(),
            Self::Filtered(rows) => rows.next(),
        }
    }
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

pub(crate) fn cache_has_records(cache: &FormatPivotCache) -> bool {
    !matches!(cache.source, FormatPivotSource::Olap { .. })
}

pub(crate) fn write_pivot_parts<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    styles: &StyleMapping,
) -> XlsbResult<()> {
    for cache in &plan.caches {
        write_pivot_cache_definition_part(zip, options, workbook, plan, cache)?;
        if cache_has_records(cache) {
            write_pivot_cache_records_part(zip, options, workbook, plan, cache)?;
        }
        write_pivot_cache_definition_rels(zip, options, cache)?;
    }

    for part in &plan.tables {
        let cache = cache_for_table(plan, part)?;
        write_pivot_table_part(zip, options, workbook, plan, part, cache, styles)?;
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
    write_pivot_cache_source(&mut rw, workbook, cache)?;

    if cache_has_records(cache) {
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
                        .unwrap_or(usage.item_kinds[*source_index]);
                    let shared_items = grouping_info
                        .map(|info| info.source_items.as_slice())
                        .unwrap_or(&field.shared_items);
                    let calculated_item_indexes =
                        calculated_item_indexes_for_field(cache, *source_index, shared_items);
                    let pnames = write_pcd_field(&mut rw, field, cache)?;
                    if field.formula.is_none() {
                        write_pcd_shared_items(
                            &mut rw,
                            shared_items,
                            usage.store_items[*source_index],
                            item_kind,
                            date_system,
                            &calculated_item_indexes,
                        )?;
                    }
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
        write_pcd_calculated_items(&mut rw, cache)?;
    } else {
        let pivot = pivot_for_cache(workbook, plan, cache)?.ok_or_else(|| {
            XlsbError::InvalidFormat("OLAP pivot cache has no pivot table".into())
        })?;
        write_olap_pcd_fields(&mut rw, pivot, cache, &layout)?;
        write_olap_cache_hierarchies(&mut rw, pivot, cache, &layout)?;
        write_olap_dimensions(&mut rw, pivot, cache, &layout)?;
    }
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
                if field.formula.is_some() {
                    continue;
                }
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
    styles: &StyleMapping,
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
    let effective_columns = effective_column_fields(pivot, cache);
    let axis_tuples = build_xlsb_axis_tuples(
        part,
        pivot,
        cache,
        &layout,
        &grouping_infos,
        &effective_columns,
    )?;
    validate_xlsb_pivot_axis_field_options(pivot)?;
    validate_xlsb_pivot_filters(pivot, cache, &layout)?;

    let path = format!("xl/pivotTables/pivotTable{}.bin", part.table_num);
    zip.start_file(path, *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    write_ac_block(&mut rw, 7, false)?;
    write_begin_sx_view(&mut rw, pivot)?;
    write_sx_location(&mut rw, pivot, cache, &layout, &axis_tuples)?;
    rw.write_record(records::BRT_END_SX_LOCATION, &[])?;
    write_sx_fields(&mut rw, pivot, cache, &usage, &layout, &grouping_infos)?;
    if !cache_has_records(cache) {
        write_olap_pivot_hierarchies(&mut rw, pivot, cache, &layout)?;
    }
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
    if !cache_has_records(cache) {
        write_olap_axis_hierarchies(
            &mut rw,
            records::BRT_BEGIN_ISXTH_RWS,
            records::BRT_END_ISXTH_RWS,
            cache,
            &layout,
            &pivot.rows,
            values_on_rows,
            pivot.layout.values_axis_position,
        )?;
    }
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
        &axis_tuples.rows,
    )?;
    write_axis_fields(
        &mut rw,
        records::BRT_BEGIN_ISXVD_COLS,
        records::BRT_END_ISXVD_COLS,
        cache,
        &layout,
        &effective_columns,
        values_on_columns,
        pivot.layout.values_axis_position,
    )?;
    if !cache_has_records(cache) {
        write_olap_axis_hierarchies(
            &mut rw,
            records::BRT_BEGIN_ISXTH_COLS,
            records::BRT_END_ISXTH_COLS,
            cache,
            &layout,
            &effective_columns,
            values_on_columns,
            pivot.layout.values_axis_position,
        )?;
    }
    write_axis_items(
        &mut rw,
        records::BRT_BEGIN_SX_COL_ITEMS,
        records::BRT_END_SX_COL_ITEMS,
        cache,
        &layout,
        &effective_columns,
        values_on_columns,
        pivot.layout.values_axis_position,
        pivot.measures.len(),
        &axis_tuples.columns,
    )?;
    write_page_fields(&mut rw, pivot, cache, &layout, &grouping_infos)?;
    write_data_fields(&mut rw, pivot, cache, styles)?;
    write_sx_style(&mut rw, pivot)?;
    write_pivot_filters(&mut rw, pivot, cache, &layout, date_system)?;
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
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    let path = format!(
        "xl/pivotCache/_rels/pivotCacheDefinition{}.bin.rels",
        cache.cache_num
    );
    zip.start_file(path, *options)?;
    let mut xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        .to_string();
    if cache_has_records(cache) {
        let target = format!("pivotCacheRecords{}.bin", cache.cache_num);
        xml.push_str(&format!(
            "<Relationship Id=\"rId1\" Type=\"{}\" Target=\"{}\"/>",
            RT_PIVOT_CACHE_RECORDS,
            xml_attr_escape(&target)
        ));
    }
    for (id, target) in consolidation_external_relationships(&cache.source)? {
        xml.push_str(&format!(
            "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\" TargetMode=\"External\"/>",
            xml_attr_escape(&id),
            RT_EXTERNAL_LINK_PATH,
            xml_attr_escape(&target)
        ));
    }
    xml.push_str("</Relationships>");
    zip.write_all(xml.as_bytes())?;
    Ok(())
}

fn consolidation_external_relationships(
    source: &FormatPivotSource,
) -> XlsbResult<Vec<(String, String)>> {
    let FormatPivotSource::Consolidation { ranges } = source else {
        return Ok(Vec::new());
    };

    let mut relationships = Vec::new();
    for range in ranges {
        if range.external_relationship_id.is_some() && range.external_relationship_target.is_none()
        {
            return Err(XlsbError::InvalidFormat(
                "XLSB external consolidation references require a relationship target".into(),
            ));
        }
        let Some(target) = &range.external_relationship_target else {
            continue;
        };
        if target.trim().is_empty() {
            return Err(XlsbError::InvalidFormat(
                "XLSB external consolidation relationship target cannot be blank".into(),
            ));
        }
        let id = consolidation_external_relationship_id(range, relationships.len() + 1);
        relationships.push((id, target.clone()));
    }
    Ok(relationships)
}

fn groupings_for_cache<'a>(
    workbook: &'a Workbook,
    plan: &'a FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsbResult<&'a [PivotGrouping]> {
    Ok(pivot_for_cache(workbook, plan, cache)?
        .map(|pivot| pivot.groupings.as_slice())
        .unwrap_or(&[]))
}

fn pivot_for_cache<'a>(
    workbook: &'a Workbook,
    plan: &FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsbResult<Option<&'a PivotTable>> {
    let Some(part) = plan
        .tables
        .iter()
        .find(|part| part.cache_num == cache.cache_num)
    else {
        return Ok(None);
    };
    let worksheet = workbook
        .worksheet(part.sheet_index)
        .ok_or_else(|| XlsbError::InvalidFormat("pivot table sheet not found".into()))?;
    let pivot = worksheet
        .pivot_tables()
        .get(part.pivot_index)
        .ok_or_else(|| XlsbError::InvalidFormat("pivot table not found".into()))?;
    Ok(Some(pivot))
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
) -> XlsbResult<()> {
    if cache.background_query && !matches!(cache.source, FormatPivotSource::External { .. }) {
        return Err(XlsbError::InvalidFormat(
            "XLSB pivot background query is only valid for external pivot cache sources".into(),
        ));
    }
    let mut payload = Vec::new();
    payload.push(8);
    payload.push(3);
    payload.push(8);
    let has_records = cache_has_records(cache);
    let mut flags = if has_records { 0x11u8 } else { 0x10u8 };
    if cache.refresh_on_load {
        flags |= 0x04;
    }
    if cache.background_query {
        flags |= 0x20;
    }
    payload.push(flags);
    let ghost_limit = cache
        .missing_items_limit
        .map(|limit| limit as i32)
        .unwrap_or(-1);
    payload.extend_from_slice(&ghost_limit.to_le_bytes());
    payload.extend_from_slice(&0f64.to_le_bytes());
    payload.push(if has_records { 0x02 } else { 0x00 });
    payload
        .extend_from_slice(&(if has_records { cache.row_count } else { 0 } as u32).to_le_bytes());
    if has_records {
        payload.extend_from_slice(&encode_wide_str("rId1"));
    }
    payload.extend_from_slice(&0u32.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_PIVOT_CACHE_DEF, &payload)?;
    Ok(())
}

fn write_pivot_cache_source<W: Write>(
    rw: &mut RecordWriter<W>,
    workbook: &Workbook,
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    match &cache.source {
        FormatPivotSource::Worksheet {
            sheet_name,
            range,
            table_name,
            ..
        } => write_worksheet_pivot_cache_source(rw, sheet_name, *range, table_name.as_deref()),
        FormatPivotSource::Consolidation { ranges } => {
            write_consolidation_pivot_cache_source(rw, ranges)
        }
        FormatPivotSource::External {
            connection_name, ..
        } => write_external_pivot_cache_source(rw, workbook, connection_name),
        FormatPivotSource::Olap {
            connection_name,
            cube,
            command_text,
        } => write_olap_pivot_cache_source(
            rw,
            workbook,
            connection_name,
            cube.as_deref(),
            command_text.as_deref(),
        ),
        FormatPivotSource::Scenario { name } => write_scenario_pivot_cache_source(rw, name),
    }
}

fn write_external_pivot_cache_source<W: Write>(
    rw: &mut RecordWriter<W>,
    workbook: &Workbook,
    connection_name: &str,
) -> XlsbResult<()> {
    let connection = find_workbook_connection(workbook, connection_name).ok_or_else(|| {
        XlsbError::InvalidFormat(format!(
            "XLSB external pivot source references unknown connection: {connection_name}"
        ))
    })?;
    if matches!(&connection.kind, WorkbookConnectionKind::Olap { .. }) {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB external pivot source requires a non-OLAP connection: {connection_name}"
        )));
    }
    write_connection_pivot_cache_source(rw, connection.id)
}

fn write_olap_pivot_cache_source<W: Write>(
    rw: &mut RecordWriter<W>,
    workbook: &Workbook,
    connection_name: &str,
    cube: Option<&str>,
    command_text: Option<&str>,
) -> XlsbResult<()> {
    let connection = find_workbook_connection(workbook, connection_name).ok_or_else(|| {
        XlsbError::InvalidFormat(format!(
            "XLSB OLAP pivot source references unknown connection: {connection_name}"
        ))
    })?;
    let WorkbookConnectionKind::Olap {
        command,
        command_type,
        ..
    } = &connection.kind
    else {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB OLAP pivot source requires an OLAP connection: {connection_name}"
        )));
    };

    if let (Some(cube), Some(command)) = (cube, command.as_deref()) {
        if *command_type == Some(1) && command != cube {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB OLAP pivot source cube does not match connection command: {connection_name}"
            )));
        }
    }
    if let (Some(command_text), Some(command)) = (command_text, command.as_deref()) {
        if command != command_text {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB OLAP pivot source command text does not match connection command: {connection_name}"
            )));
        }
    }

    write_connection_pivot_cache_source(rw, connection.id)
}

fn find_workbook_connection<'a>(
    workbook: &'a Workbook,
    connection_name: &str,
) -> Option<&'a WorkbookConnection> {
    connection_name
        .parse::<u32>()
        .ok()
        .and_then(|id| workbook.data_connection_by_id(id))
        .or_else(|| workbook.data_connection_by_name(connection_name))
}

fn write_connection_pivot_cache_source<W: Write>(
    rw: &mut RecordWriter<W>,
    connection_id: u32,
) -> XlsbResult<()> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&connection_id.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_PCD_SOURCE, &payload)?;
    rw.write_record(records::BRT_END_PCD_SOURCE, &[])?;
    Ok(())
}

fn write_scenario_pivot_cache_source<W: Write>(
    rw: &mut RecordWriter<W>,
    name: &str,
) -> XlsbResult<()> {
    if !name.is_empty() {
        return Err(XlsbError::InvalidFormat(
            "XLSB named scenario pivot source authoring requires Scenario Manager records and is not implemented yet".into(),
        ));
    }
    let mut payload = Vec::new();
    payload.extend_from_slice(&3u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_PCD_SOURCE, &payload)?;
    rw.write_record(records::BRT_END_PCD_SOURCE, &[])?;
    Ok(())
}

fn write_worksheet_pivot_cache_source<W: Write>(
    rw: &mut RecordWriter<W>,
    sheet_name: &str,
    range: CellRange,
    table_name: Option<&str>,
) -> XlsbResult<()> {
    rw.write_record(records::BRT_BEGIN_PCD_SOURCE, &[0; 8])?;
    let mut payload = Vec::new();
    if let Some(table_name) = table_name {
        // MS-XLSB 2.4.167: fName=1 means the source is specified by
        // namedRange. Excel uses the same path for table-backed pivots.
        payload.extend_from_slice(&[0x01, 0x00, 0x00]);
        payload.extend_from_slice(&encode_wide_str(table_name));
    } else {
        payload.extend_from_slice(&[0x00, 0x00, 0x02]);
        payload.extend_from_slice(&encode_wide_str(sheet_name));
        write_unchecked_rfx(&mut payload, range);
    }
    rw.write_record(records::BRT_BEGIN_PCDS_SHEET, &payload)?;
    rw.write_record(records::BRT_END_PCDS_SHEET, &[])?;
    rw.write_record(records::BRT_END_PCD_SOURCE, &[])?;
    Ok(())
}

fn write_consolidation_pivot_cache_source<W: Write>(
    rw: &mut RecordWriter<W>,
    ranges: &[PivotSourceRange],
) -> XlsbResult<()> {
    if ranges.is_empty() {
        return Err(XlsbError::InvalidFormat(
            "XLSB consolidation pivot sources require at least one range".into(),
        ));
    }
    let pages = consolidation_pages(ranges)?;

    let mut source_payload = Vec::new();
    source_payload.extend_from_slice(&2u32.to_le_bytes());
    source_payload.extend_from_slice(&0u32.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_PCD_SOURCE, &source_payload)?;
    rw.write_record(records::BRT_BEGIN_PCDS_CONSOL, &0u16.to_le_bytes())?;

    if !pages.is_empty() {
        rw.write_record(
            records::BRT_BEGIN_PCDSC_PAGES,
            &checked_u32(pages.len(), "XLSB consolidation page count")?.to_le_bytes(),
        )?;
        for page in &pages {
            rw.write_record(
                records::BRT_BEGIN_PCDSC_PAGE,
                &checked_u32(page.len(), "XLSB consolidation page item count")?.to_le_bytes(),
            )?;
            for item in page {
                rw.write_record(records::BRT_BEGIN_PCDSC_PITEM, &encode_wide_str(item))?;
                rw.write_record(records::BRT_END_PCDSC_PITEM, &[])?;
            }
            rw.write_record(records::BRT_END_PCDSC_PAGE, &[])?;
        }
        rw.write_record(records::BRT_END_PCDSC_PAGES, &[])?;
    }

    rw.write_record(
        records::BRT_BEGIN_PCDSC_SETS,
        &checked_u32(ranges.len(), "XLSB consolidation source-set count")?.to_le_bytes(),
    )?;
    let mut external_index = 0usize;
    for range in ranges {
        if range.external_relationship_id.is_some() && range.external_relationship_target.is_none()
        {
            return Err(XlsbError::InvalidFormat(
                "XLSB external consolidation references require a relationship target".into(),
            ));
        }
        let external_relationship_id = if range.external_relationship_target.is_some() {
            external_index += 1;
            Some(consolidation_external_relationship_id(
                range,
                external_index,
            ))
        } else {
            None
        };
        let payload =
            consolidation_source_set_payload(range, &pages, external_relationship_id.as_deref())?;
        rw.write_record(records::BRT_BEGIN_PCDSC_SET, &payload)?;
        rw.write_record(records::BRT_END_PCDSC_SET, &[])?;
    }
    rw.write_record(records::BRT_END_PCDSC_SETS, &[])?;
    rw.write_record(records::BRT_END_PCDS_CONSOL, &[])?;
    rw.write_record(records::BRT_END_PCD_SOURCE, &[])?;
    Ok(())
}

fn consolidation_source_set_payload(
    range: &PivotSourceRange,
    pages: &[Vec<String>],
    external_relationship_id: Option<&str>,
) -> XlsbResult<Vec<u8>> {
    if range.range.is_none() && range.name.is_none() {
        return Err(XlsbError::InvalidFormat(
            "XLSB consolidation source set requires a range or name".into(),
        ));
    }

    let mut item_indexes = [u32::MAX; 4];
    for (index, item) in range.page_items.iter().enumerate() {
        let page = pages.get(index).ok_or_else(|| {
            XlsbError::InvalidFormat(
                "XLSB consolidation page item has no matching page field".into(),
            )
        })?;
        let Some(item_index) = page.iter().position(|candidate| candidate == item) else {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB consolidation page item is not declared: {item}"
            )));
        };
        item_indexes[index] = checked_u32(item_index, "XLSB consolidation page item index")?;
    }

    let source_is_name = range.range.is_none();
    let mut payload = Vec::new();
    for item_index in item_indexes {
        payload.extend_from_slice(&item_index.to_le_bytes());
    }
    payload.push(u8::from(source_is_name));
    payload.push(0);

    let mut flags = 0u8;
    if external_relationship_id.is_some() {
        flags |= 0x01;
    }
    if range.sheet.is_some() {
        flags |= 0x02;
    }
    payload.push(flags);

    if let Some(sheet) = &range.sheet {
        payload.extend_from_slice(&encode_wide_str(sheet));
    }
    if let Some(external_relationship_id) = external_relationship_id {
        payload.extend_from_slice(&encode_wide_str(external_relationship_id));
    }
    if source_is_name {
        let Some(name) = &range.name else {
            return Err(XlsbError::InvalidFormat(
                "XLSB named consolidation source requires a name".into(),
            ));
        };
        payload.extend_from_slice(&encode_wide_str(name));
    } else {
        let Some(source_range) = range.range else {
            return Err(XlsbError::InvalidFormat(
                "XLSB consolidation source set requires a range".into(),
            ));
        };
        write_unchecked_rfx(&mut payload, source_range);
    }
    Ok(payload)
}

fn consolidation_external_relationship_id(
    range: &PivotSourceRange,
    external_index: usize,
) -> String {
    range
        .external_relationship_id
        .clone()
        .unwrap_or_else(|| format!("rIdExternal{external_index}"))
}

fn consolidation_pages(ranges: &[PivotSourceRange]) -> XlsbResult<Vec<Vec<String>>> {
    let page_count = ranges
        .iter()
        .map(|range| range.page_items.len())
        .max()
        .unwrap_or(0);
    if page_count > 4 {
        return Err(XlsbError::InvalidFormat(
            "XLSB consolidation pivot sources support at most four page fields".into(),
        ));
    }

    let mut pages = vec![Vec::<String>::new(); page_count];
    for range in ranges {
        for (index, item) in range.page_items.iter().enumerate() {
            if item.trim().is_empty() {
                return Err(XlsbError::InvalidFormat(
                    "XLSB consolidation page item names cannot be blank".into(),
                ));
            }
            if !pages[index].iter().any(|candidate| candidate == item) {
                pages[index].push(item.clone());
            }
        }
    }
    Ok(pages)
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
    let num_format_id = if matches!(cache.source, FormatPivotSource::Consolidation { .. }) {
        0
    } else {
        u32::MAX
    };
    payload.extend_from_slice(&num_format_id.to_le_bytes());
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

fn write_olap_pcd_fields<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
) -> XlsbResult<()> {
    validate_olap_metadata_cache(pivot, cache, layout)?;
    rw.write_record(
        records::BRT_BEGIN_PCD_FIELDS,
        &(layout.fields.len() as u32).to_le_bytes(),
    )?;
    for (layout_index, field_layout) in layout.fields.iter().enumerate() {
        let source_index = olap_source_index(field_layout)?;
        let field = &cache.fields[source_index];
        write_olap_pcd_field_header(rw, &field.name, layout_index)?;
        rw.write_record(records::BRT_END_PCD_FIELD, &[])?;
    }
    rw.write_record(records::BRT_END_PCD_FIELDS, &[])?;
    Ok(())
}

fn validate_olap_metadata_cache(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
) -> XlsbResult<()> {
    if !cache.calculated_items.is_empty()
        || cache.fields.iter().any(|field| field.formula.is_some())
    {
        return Err(XlsbError::InvalidFormat(
            "XLSB OLAP pivot source authoring does not support calculated fields or calculated items"
                .into(),
        ));
    }
    if pivot
        .groupings
        .iter()
        .any(|grouping| cache.field_index(grouping_field_name(grouping)).is_some())
    {
        return Err(XlsbError::InvalidFormat(
            "XLSB OLAP pivot source authoring does not support OLAP grouping metadata yet".into(),
        ));
    }
    for field_layout in &layout.fields {
        olap_source_index(field_layout)?;
    }
    Ok(())
}

fn write_olap_pcd_field_header<W: Write>(
    rw: &mut RecordWriter<W>,
    name: &str,
    hierarchy_index: usize,
) -> XlsbResult<()> {
    let hierarchy_index = checked_u32(hierarchy_index, "OLAP cache hierarchy index")?;
    let mut payload = Vec::new();
    payload.extend_from_slice(&0x0004u16.to_le_bytes());
    payload.extend_from_slice(&u32::MAX.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&hierarchy_index.to_le_bytes());
    payload.extend_from_slice(&0x0000_7FFFu32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(name));
    rw.write_record(records::BRT_BEGIN_PCD_FIELD, &payload)?;
    Ok(())
}

fn write_olap_cache_hierarchies<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
) -> XlsbResult<()> {
    rw.write_record(
        records::BRT_BEGIN_PCD_HIERARCHIES,
        &(layout.fields.len() as u32).to_le_bytes(),
    )?;
    for (layout_index, field_layout) in layout.fields.iter().enumerate() {
        let source_index = olap_source_index(field_layout)?;
        let field = &cache.fields[source_index];
        let is_measure = pivot_measure_field(pivot, &field.name);
        write_olap_cache_hierarchy(rw, &field.name, is_measure)?;

        let field_index = checked_u32(layout_index, "OLAP cache hierarchy field index")?;
        let mut usage = Vec::new();
        usage.extend_from_slice(&1u32.to_le_bytes());
        usage.extend_from_slice(&field_index.to_le_bytes());
        rw.write_record(records::BRT_BEGIN_PCDH_FIELDS_USAGE, &usage)?;
        rw.write_record(records::BRT_END_PCDH_FIELDS_USAGE, &[])?;
        rw.write_record(records::BRT_END_PCD_HIERARCHY, &[])?;
    }
    rw.write_record(records::BRT_END_PCD_HIERARCHIES, &[])?;
    Ok(())
}

fn write_olap_cache_hierarchy<W: Write>(
    rw: &mut RecordWriter<W>,
    field_name: &str,
    is_measure: bool,
) -> XlsbResult<()> {
    let unique = if is_measure {
        olap_measure_unique_name(field_name)
    } else {
        olap_unique_name(field_name)
    };
    let caption = field_name.trim();
    let mut flags: u16 = 0x0010;
    if is_measure {
        flags |= 0x0001;
    } else {
        flags |= 0x0004;
    }

    let mut payload = Vec::new();
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&0i32.to_le_bytes());
    if is_measure {
        payload.push(0);
    } else {
        payload.push(0x01);
    }
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(&unique));
    payload.extend_from_slice(&encode_wide_str(caption));
    if !is_measure {
        payload.extend_from_slice(&encode_wide_str(&unique));
    }
    rw.write_record(records::BRT_BEGIN_PCD_HIERARCHY, &payload)?;
    Ok(())
}

fn write_olap_dimensions<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
) -> XlsbResult<()> {
    let mut dims = Vec::new();
    let mut has_measure_dim = false;
    for field_layout in &layout.fields {
        let source_index = olap_source_index(field_layout)?;
        let field = &cache.fields[source_index];
        if pivot_measure_field(pivot, &field.name) {
            has_measure_dim = true;
        } else {
            let unique = olap_unique_name(&field.name);
            dims.push((false, field.name.clone(), unique, field.name.clone()));
        }
    }
    if has_measure_dim {
        dims.push((
            true,
            "Measures".to_string(),
            "[Measures]".to_string(),
            "Measures".to_string(),
        ));
    }

    rw.write_record(records::BRT_BEGIN_DIMS, &(dims.len() as u32).to_le_bytes())?;
    for (is_measure, name, unique, display) in dims {
        let mut payload = Vec::new();
        payload.push(if is_measure { 0x01 } else { 0x00 });
        payload.extend_from_slice(&encode_wide_str(&name));
        payload.extend_from_slice(&encode_wide_str(&unique));
        payload.extend_from_slice(&encode_wide_str(&display));
        rw.write_record(records::BRT_BEGIN_DIM, &payload)?;
        rw.write_record(records::BRT_END_DIM, &[])?;
    }
    rw.write_record(records::BRT_END_DIMS, &[])?;
    Ok(())
}

fn olap_source_index(field_layout: &XlsbPivotCacheFieldLayout) -> XlsbResult<usize> {
    match field_layout {
        XlsbPivotCacheFieldLayout::Source { source_index, .. } => Ok(*source_index),
        XlsbPivotCacheFieldLayout::ManualDerived { .. }
        | XlsbPivotCacheFieldLayout::DateUnitDerived { .. } => Err(XlsbError::InvalidFormat(
            "XLSB OLAP pivot source authoring does not support derived cache fields".into(),
        )),
    }
}

fn pivot_measure_field(pivot: &PivotTable, field_name: &str) -> bool {
    pivot
        .measures
        .iter()
        .any(|measure| measure.field.name.eq_ignore_ascii_case(field_name))
}

fn olap_unique_name(name: &str) -> String {
    let trimmed = name.trim();
    if trimmed.starts_with('[') {
        trimmed.to_string()
    } else {
        format!("[{}]", trimmed.replace(']', "]]"))
    }
}

fn olap_measure_unique_name(name: &str) -> String {
    format!("[Measures].{}", olap_unique_name(name))
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

fn write_pcd_calculated_items<W: Write>(
    rw: &mut RecordWriter<W>,
    cache: &FormatPivotCache,
) -> XlsbResult<()> {
    if cache.calculated_items.is_empty() {
        return Ok(());
    }

    rw.write_record(
        records::BRT_BEGIN_PCD_CALC_ITEMS,
        &checked_u32(cache.calculated_items.len(), "pivot calculated item count")?.to_le_bytes(),
    )?;
    for item in &cache.calculated_items {
        let field_index = cache.field_index(&item.field.name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "XLSB pivot calculated item references unknown field: {}",
                item.field.name
            ))
        })?;
        let item_index = cache.fields[field_index]
            .shared_items
            .iter()
            .position(|value| value == &item.item)
            .ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "XLSB pivot calculated item {} was not registered in field {}",
                    item.item, item.field.name
                ))
            })?;
        let (rgce, pnames) =
            compile_pivot_calculated_item_formula(&item.formula, cache, field_index)?;

        let mut payload = Vec::new();
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
        payload.extend_from_slice(
            &checked_u32(rgce.len(), "pivot calculated item formula token length")?.to_le_bytes(),
        );
        payload.extend_from_slice(&rgce);
        payload.extend_from_slice(&0u32.to_le_bytes());
        rw.write_record(records::BRT_BEGIN_PCD_CALC_ITEM, &payload)?;
        write_pivot_rule_for_calculated_item(rw, field_index, item_index)?;
        write_calculated_item_pnames(rw, &pnames)?;
        rw.write_record(records::BRT_END_PCD_CALC_ITEM, &[])?;
    }
    rw.write_record(records::BRT_END_PCD_CALC_ITEMS, &[])?;
    Ok(())
}

fn compile_pivot_calculated_item_formula(
    formula: &str,
    cache: &FormatPivotCache,
    field_index: usize,
) -> XlsbResult<(Vec<u8>, Vec<(usize, usize)>)> {
    let trimmed = formula.trim();
    let parse_input = if trimmed.starts_with('=') {
        trimmed.to_string()
    } else {
        format!("={trimmed}")
    };
    let expr = duke_sheets_formula::parse_formula(&parse_input).map_err(|err| {
        XlsbError::InvalidFormat(format!(
            "XLSB pivot calculated item formula could not be parsed: {err}"
        ))
    })?;

    let mut rgce = Vec::new();
    let mut pnames = Vec::new();
    compile_pivot_item_formula_expr(&expr, &mut rgce, &mut pnames, cache, field_index).map_err(
        |_| {
            XlsbError::InvalidFormat(format!(
                "XLSB pivot calculated item formula uses unsupported syntax: {formula}"
            ))
        },
    )?;
    Ok((rgce, pnames))
}

fn write_pivot_rule_for_calculated_item<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: usize,
    item_index: usize,
) -> XlsbResult<()> {
    rw.write_record(
        records::BRT_BEGIN_P_RULE,
        &[0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x11, 0x00, 0x00],
    )?;
    rw.write_record(records::BRT_BEGIN_PR_FILTERS, &1u32.to_le_bytes())?;
    let mut filter = Vec::with_capacity(11);
    filter.extend_from_slice(
        &checked_u32(field_index, "pivot calculated item target field index")?.to_le_bytes(),
    );
    filter.extend_from_slice(&1u32.to_le_bytes());
    filter.extend_from_slice(&[0x01, 0x00, 0x01]);
    rw.write_record(records::BRT_BEGIN_PR_FILTER, &filter)?;
    rw.write_record(
        records::BRT_BEGIN_PRF_ITEM,
        &checked_u32(item_index, "pivot calculated item target item index")?.to_le_bytes(),
    )?;
    rw.write_record(records::BRT_END_PRF_ITEM, &[])?;
    rw.write_record(records::BRT_END_PR_FILTER, &[])?;
    rw.write_record(records::BRT_END_PR_FILTERS, &[])?;
    rw.write_record(records::BRT_END_P_RULE, &[])?;
    Ok(())
}

fn write_calculated_item_pnames<W: Write>(
    rw: &mut RecordWriter<W>,
    pnames: &[(usize, usize)],
) -> XlsbResult<()> {
    if pnames.is_empty() {
        return Ok(());
    }
    rw.write_record(
        records::BRT_BEGIN_PNAMES,
        &checked_u32(pnames.len(), "pivot calculated item PName count")?.to_le_bytes(),
    )?;
    for (field_index, item_index) in pnames {
        let mut payload = Vec::with_capacity(6);
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
        payload.push(0x01);
        payload.push(0x00);
        rw.write_record(records::BRT_BEGIN_PNAME, &payload)?;
        rw.write_record(records::BRT_BEGIN_PNPAIRS, &1u32.to_le_bytes())?;
        let mut pair = Vec::with_capacity(9);
        pair.push(0x00);
        pair.extend_from_slice(
            &checked_u32(*field_index, "pivot calculated item formula field index")?.to_le_bytes(),
        );
        pair.extend_from_slice(
            &checked_i32(*item_index, "pivot calculated item formula item index")?.to_le_bytes(),
        );
        rw.write_record(records::BRT_BEGIN_PNPAIR, &pair)?;
        rw.write_record(records::BRT_END_PNPAIR, &[])?;
        rw.write_record(records::BRT_END_PNPAIRS, &[])?;
        rw.write_record(records::BRT_END_PNAME, &[])?;
    }
    rw.write_record(records::BRT_END_PNAMES, &[])?;
    Ok(())
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
        payload.push(0x01);
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
        FormulaExpr::Function { name, args } => {
            compile_pivot_function_expr(name, args, out, |arg, out| {
                compile_pivot_formula_expr(arg, out, pnames, cache)
            })
        }
        FormulaExpr::CellRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::ExternalFunction { .. }
        | FormulaExpr::Array(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => Err(()),
    }
}

fn compile_pivot_item_formula_expr(
    expr: &FormulaExpr,
    out: &mut Vec<u8>,
    pnames: &mut Vec<(usize, usize)>,
    cache: &FormatPivotCache,
    field_index: usize,
) -> Result<(), ()> {
    match expr {
        FormulaExpr::Number(n) => emit_pivot_number(*n, out),
        FormulaExpr::String(s) => {
            if try_emit_pivot_item_sxname(s, out, pnames, cache, field_index)? {
                Ok(())
            } else {
                emit_pivot_string(s, out)
            }
        }
        FormulaExpr::Boolean(value) => {
            out.push(ptg::PTG_BOOL);
            out.push(if *value { 1 } else { 0 });
            Ok(())
        }
        FormulaExpr::Error(error) => {
            out.push(ptg::PTG_ERR);
            out.push(error.code());
            Ok(())
        }
        FormulaExpr::NameRef(name) => emit_pivot_item_sxname(name, out, pnames, cache, field_index),
        FormulaExpr::BinaryOp { op, left, right } => {
            compile_pivot_item_formula_expr(left, out, pnames, cache, field_index)?;
            compile_pivot_item_formula_expr(right, out, pnames, cache, field_index)?;
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
            compile_pivot_item_formula_expr(operand, out, pnames, cache, field_index)?;
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
        FormulaExpr::Function { name, args } => {
            compile_pivot_function_expr(name, args, out, |arg, out| {
                compile_pivot_item_formula_expr(arg, out, pnames, cache, field_index)
            })
        }
        FormulaExpr::CellRef(reference) => {
            let Some(name) = pivot_calculated_item_cell_ref_name(reference) else {
                return Err(());
            };
            emit_pivot_item_sxname(&name, out, pnames, cache, field_index)
        }
        FormulaExpr::StructuredRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::ExternalFunction { .. }
        | FormulaExpr::Array(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => Err(()),
    }
}

fn compile_pivot_function_expr<F>(
    name: &str,
    args: &[FormulaExpr],
    out: &mut Vec<u8>,
    mut compile_arg: F,
) -> Result<(), ()>
where
    F: FnMut(&FormulaExpr, &mut Vec<u8>) -> Result<(), ()>,
{
    let Some(func_idx) = function_index(name) else {
        return Err(());
    };
    if function_is_biff8_addin(func_idx) || args.len() > u8::MAX as usize {
        return Err(());
    }
    for arg in args {
        if matches!(arg, FormulaExpr::Empty) {
            return Err(());
        }
        compile_arg(arg, out)?;
    }
    if function_is_fixed_arity(func_idx, args.len()) {
        out.push(ptg::v_class(ptg::PTG_FUNC));
        out.extend_from_slice(&func_idx.to_le_bytes());
    } else {
        out.push(ptg::v_class(ptg::PTG_FUNC_VAR));
        out.push(args.len() as u8);
        out.extend_from_slice(&func_idx.to_le_bytes());
    }
    Ok(())
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

fn emit_pivot_item_sxname(
    item_name: &str,
    out: &mut Vec<u8>,
    pnames: &mut Vec<(usize, usize)>,
    cache: &FormatPivotCache,
    field_index: usize,
) -> Result<(), ()> {
    if try_emit_pivot_item_sxname(item_name, out, pnames, cache, field_index)? {
        Ok(())
    } else {
        Err(())
    }
}

fn try_emit_pivot_item_sxname(
    item_name: &str,
    out: &mut Vec<u8>,
    pnames: &mut Vec<(usize, usize)>,
    cache: &FormatPivotCache,
    field_index: usize,
) -> Result<bool, ()> {
    let item_value = PivotValue::String(item_name.to_string());
    let Some(item_index) = cache.fields[field_index]
        .shared_items
        .iter()
        .position(|value| value == &item_value)
    else {
        return Ok(false);
    };
    let pname_index =
        checked_u32(pnames.len(), "pivot calculated item formula PName index").map_err(|_| ())?;
    pnames.push((field_index, item_index));
    out.extend_from_slice(&[0x18, 0x1D]);
    out.extend_from_slice(&pname_index.to_le_bytes());
    Ok(true)
}

fn pivot_calculated_item_cell_ref_name(reference: &CellReference) -> Option<String> {
    reference
        .sheet
        .is_none()
        .then(|| reference.address.to_a1_string())
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
    calculated_item_indexes: &HashSet<usize>,
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
        for (index, item) in values.iter().enumerate() {
            write_pcdi_item(
                rw,
                item,
                item_kind,
                date_system,
                calculated_item_indexes.contains(&index),
            )?;
        }
    }
    rw.write_record(records::BRT_END_PCD_SHARED_ITEMS, &[])?;
    Ok(())
}

fn calculated_item_indexes_for_field(
    cache: &FormatPivotCache,
    field_index: usize,
    shared_items: &[PivotValue],
) -> HashSet<usize> {
    let field_name = &cache.fields[field_index].name;
    cache
        .calculated_items
        .iter()
        .filter(|item| item.field.name.eq_ignore_ascii_case(field_name))
        .filter_map(|item| shared_items.iter().position(|value| value == &item.item))
        .collect()
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
            write_pcdi_item(rw, item, PivotCacheItemKind::Normal, date_system, false)?;
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
    calculated: bool,
) -> XlsbResult<()> {
    if calculated {
        return write_pcdia_item(rw, value, item_kind, date_system);
    }

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

fn write_pcdia_item<W: Write>(
    rw: &mut RecordWriter<W>,
    value: &PivotValue,
    item_kind: PivotCacheItemKind,
    date_system: DateSystem,
) -> XlsbResult<()> {
    let mut payload = Vec::new();
    let record_type = match (item_kind, value) {
        (PivotCacheItemKind::DateTime, PivotValue::Number(serial)) => {
            append_pcdi_datetime_payload(&mut payload, *serial, date_system)?;
            records::BRT_PCDIA_DATETIME
        }
        (PivotCacheItemKind::DateTime, _) => {
            return Err(XlsbError::InvalidFormat(
                "XLSB pivot calculated datetime cache item requires a numeric serial".into(),
            ));
        }
        (_, PivotValue::Blank) => records::BRT_PCDIA_MISSING,
        (_, PivotValue::Boolean(value)) => {
            payload.push(if *value { 1 } else { 0 });
            records::BRT_PCDIA_BOOLEAN
        }
        (_, PivotValue::Number(value)) => {
            payload.extend_from_slice(&value.to_le_bytes());
            records::BRT_PCDIA_NUMBER
        }
        (_, PivotValue::String(value)) => {
            payload.extend_from_slice(&encode_wide_str(value));
            records::BRT_PCDIA_STRING
        }
        (_, PivotValue::Error(value)) => {
            payload.push(error_code(*value));
            records::BRT_PCDIA_ERROR
        }
    };
    payload.extend_from_slice(&0x0002u16.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    rw.write_record(record_type, &payload)?;
    Ok(())
}

fn write_pcdi_datetime<W: Write>(
    rw: &mut RecordWriter<W>,
    serial: f64,
    date_system: DateSystem,
) -> XlsbResult<()> {
    let mut payload = Vec::with_capacity(8);
    append_pcdi_datetime_payload(&mut payload, serial, date_system)?;
    rw.write_record(records::BRT_PCDI_DATETIME, &payload)?;
    Ok(())
}

fn append_pcdi_datetime_payload(
    payload: &mut Vec<u8>,
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
    payload.extend_from_slice(&(year as u16).to_le_bytes());
    payload.extend_from_slice(&(month as u16).to_le_bytes());
    payload.push(day as u8);
    payload.push(hour as u8);
    payload.push(minute as u8);
    payload.push(second as u8);
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
    write_sx_view_layout_flags(&mut header, pivot);
    header[12] = values_axis_code(pivot.layout.values_axis);
    header[13] = pivot.layout.page_wrap.min(u8::MAX as u32) as u8;
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
    if let Some(caption) = &pivot.layout.grand_total_caption {
        payload.extend_from_slice(&encode_wide_str(caption));
    }
    if pivot.layout.show_error || pivot.layout.error_caption.is_some() {
        payload.extend_from_slice(&encode_wide_str(
            pivot.layout.error_caption.as_deref().unwrap_or(""),
        ));
    }
    if let Some(caption) = &pivot.layout.missing_caption {
        payload.extend_from_slice(&encode_wide_str(caption));
    }
    rw.write_record(records::BRT_BEGIN_SXVIEW, &payload)
}

fn write_sx_view_layout_flags(header: &mut [u8; 32], pivot: &PivotTable) {
    let layout = &pivot.layout;

    header[1] = 0;
    set_u8_bit(&mut header[1], 0, layout.show_items);
    set_u8_bit(&mut header[1], 1, layout.edit_data);
    set_u8_bit(&mut header[1], 2, layout.disable_field_list);
    set_u8_bit(&mut header[1], 4, !layout.show_calculated_members);
    set_u8_bit(&mut header[1], 5, !layout.visual_totals);
    set_u8_bit(&mut header[1], 6, layout.show_multiple_label);

    let mut flags = 0u16;
    set_u16_bit(&mut flags, 0, !layout.show_data_drop_down);
    set_u16_bit(&mut flags, 4, !layout.show_expand_collapse);
    set_u16_bit(&mut flags, 5, layout.print_drill_indicators);
    set_u16_bit(&mut flags, 6, layout.show_member_property_tips);
    set_u16_bit(&mut flags, 7, !layout.show_data_tips);
    flags |= ((layout.indent.min(127) as u16) & 0x7F) << 8;
    set_u16_bit(&mut flags, 15, !layout.show_field_headers);
    header[2..4].copy_from_slice(&flags.to_le_bytes());

    set_sxview_flag(header, 2, layout.show_empty_rows);
    set_sxview_flag(header, 3, layout.show_empty_columns);
    set_sxview_flag(header, 4, layout.enable_wizard);
    set_sxview_flag(header, 5, layout.enable_drill);
    set_sxview_flag(header, 6, layout.enable_field_properties);
    set_sxview_flag(header, 7, pivot.refresh_policy.preserve_formatting);
    set_sxview_flag(header, 9, layout.show_error);
    set_sxview_flag(header, 10, layout.show_missing);
    set_sxview_flag(header, 11, layout.page_over_then_down);
    set_sxview_flag(header, 13, layout.show_row_grand_totals);
    set_sxview_flag(header, 14, layout.show_column_grand_totals);
    set_sxview_flag(header, 15, layout.field_print_titles);
    set_sxview_flag(header, 17, layout.item_print_titles);
    set_sxview_flag(header, 18, layout.merge_item_labels);
    set_sxview_flag(header, 19, true);
    set_sxview_flag(header, 20, layout.grand_total_caption.is_some());
    set_sxview_flag(header, 32, matches!(layout.kind, PivotLayoutKind::Compact));
    set_sxview_flag(header, 33, matches!(layout.kind, PivotLayoutKind::Outline));
    set_sxview_flag(header, 34, matches!(layout.kind, PivotLayoutKind::Outline));
    set_sxview_flag(header, 35, matches!(layout.kind, PivotLayoutKind::Compact));
    set_sxview_flag(
        header,
        38,
        !(layout.show_error || layout.error_caption.is_some()),
    );
    set_sxview_flag(header, 39, layout.missing_caption.is_none());
}

fn set_u8_bit(value: &mut u8, bit: u8, enabled: bool) {
    if enabled {
        *value |= 1 << bit;
    } else {
        *value &= !(1 << bit);
    }
}

fn set_u16_bit(value: &mut u16, bit: u8, enabled: bool) {
    if enabled {
        *value |= 1 << bit;
    } else {
        *value &= !(1 << bit);
    }
}

fn set_sxview_flag(header: &mut [u8; 32], bit: usize, enabled: bool) {
    let byte = 4 + bit / 8;
    let mask = 1u8 << (bit % 8);
    if enabled {
        header[byte] |= mask;
    } else {
        header[byte] &= !mask;
    }
}

fn write_sx_location<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    axis_tuples: &XlsbPivotAxisTuples,
) -> std::io::Result<()> {
    let range = pivot
        .rendered_range
        .unwrap_or_else(|| estimated_pivot_range(pivot, cache, layout, axis_tuples));
    let first_data_col = expanded_axis_field_count(cache, layout, &pivot.rows).max(1) as u32;
    let page_field_count = effective_page_field_count(cache, layout, pivot);
    let (page_rows, page_cols) = page_field_area_size(pivot, page_field_count);
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
                let has_date_derived = layout.date_derived_by_base.contains_key(source_index);
                let has_manual_derived = layout.manual_derived_by_base.contains_key(source_index);
                let semantic_axis = pivot_field_axis_for_source_index(pivot, cache, *source_index);
                let hide_manual_source = has_manual_derived && semantic_axis == 0x04;
                let axis = if has_date_derived || hide_manual_source {
                    0
                } else {
                    semantic_axis
                };
                let axis_field = (!has_date_derived)
                    .then(|| pivot_axis_field_for_source_index(pivot, cache, *source_index))
                    .flatten();
                let behavior_axis = if hide_manual_source {
                    semantic_axis
                } else {
                    axis
                };
                write_sx_field(
                    rw,
                    axis,
                    behavior_axis,
                    axis_field,
                    pivot_has_advanced_filter_for_source_index(pivot, cache, *source_index),
                )?;
                if !has_date_derived && usage.store_items[*source_index] && axis != 0x08 {
                    let grouped_items = grouping_info_for_field(grouping_infos, &field.name);
                    let view_items = if has_manual_derived {
                        grouped_items
                            .map(|info| info.source_items.as_slice())
                            .unwrap_or(&field.shared_items)
                    } else {
                        grouped_items
                            .map(|info| info.group_items.as_slice())
                            .unwrap_or(&field.shared_items)
                    };
                    let hidden_items = if has_manual_derived {
                        HashSet::new()
                    } else {
                        field_filter_hidden_item_indexes_for_source_index(
                            pivot,
                            cache,
                            *source_index,
                            view_items,
                        )?
                    };
                    let calculated_items =
                        calculated_item_indexes_for_field(cache, *source_index, view_items);
                    write_sx_field_items(
                        rw,
                        view_items.len(),
                        axis_field,
                        &hidden_items,
                        &calculated_items,
                    )?;
                }
                if let Some(measure_index) = sxvd_sort_measure_index(pivot, axis_field)? {
                    write_sxvd_auto_sort_scope(rw, measure_index)?;
                }
            }
            XlsbPivotCacheFieldLayout::ManualDerived {
                base_source_index,
                grouping_info_index,
                ..
            } => {
                let axis_field =
                    pivot_axis_field_for_source_index(pivot, cache, *base_source_index);
                let hidden_items = field_filter_hidden_item_indexes_for_source_index(
                    pivot,
                    cache,
                    *base_source_index,
                    &grouping_infos[*grouping_info_index].group_items,
                )?;
                let axis = pivot_field_axis_for_source_index(pivot, cache, *base_source_index);
                write_sx_field(
                    rw,
                    axis,
                    axis,
                    axis_field,
                    pivot_has_advanced_filter_for_source_index(pivot, cache, *base_source_index),
                )?;
                write_sx_field_items(
                    rw,
                    grouping_infos[*grouping_info_index].group_items.len(),
                    axis_field,
                    &hidden_items,
                    &HashSet::new(),
                )?;
                if let Some(measure_index) = sxvd_sort_measure_index(pivot, axis_field)? {
                    write_sxvd_auto_sort_scope(rw, measure_index)?;
                }
            }
            XlsbPivotCacheFieldLayout::DateUnitDerived {
                base_source_index,
                grouping_info_index,
                name,
            } => {
                let base_field = &cache.fields[*base_source_index];
                let axis_field =
                    pivot_axis_field_for_source_index(pivot, cache, *base_source_index);
                let hidden_items = field_filter_hidden_item_indexes_for_page_field(
                    pivot,
                    name,
                    &grouping_infos[*grouping_info_index].group_items,
                    &base_field.name,
                )?;
                let axis = pivot_field_axis_for_source_index(pivot, cache, *base_source_index);
                write_sx_field(
                    rw,
                    axis,
                    axis,
                    axis_field,
                    pivot_has_advanced_filter_for_source_index(pivot, cache, *base_source_index),
                )?;
                write_sx_field_items(
                    rw,
                    grouping_infos[*grouping_info_index].group_items.len(),
                    axis_field,
                    &hidden_items,
                    &HashSet::new(),
                )?;
                if let Some(measure_index) = sxvd_sort_measure_index(pivot, axis_field)? {
                    write_sxvd_auto_sort_scope(rw, measure_index)?;
                }
            }
        }
        rw.write_record(records::BRT_END_SXVD, &[])?;
    }
    rw.write_record(records::BRT_END_SXVDS, &[])?;
    Ok(())
}

fn write_olap_pivot_hierarchies<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
) -> XlsbResult<()> {
    rw.write_record(
        records::BRT_BEGIN_SXTHS,
        &(layout.fields.len() as u32).to_le_bytes(),
    )?;
    for field_layout in &layout.fields {
        let source_index = olap_source_index(field_layout)?;
        let field = &cache.fields[source_index];
        let is_measure = pivot_measure_field(pivot, &field.name);
        let mut flags = 0x0200u32;
        if is_measure {
            flags |= 0x0080 | 0x0100;
        } else {
            flags |= 0x0010 | 0x0020 | 0x0040 | 0x0080;
        }
        let mut payload = Vec::new();
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        rw.write_record(records::BRT_BEGIN_SXTH, &payload)?;
        rw.write_record(records::BRT_END_SXTH, &[])?;
    }
    rw.write_record(records::BRT_END_SXTHS, &[])?;
    Ok(())
}

fn write_sx_field<W: Write>(
    rw: &mut RecordWriter<W>,
    axis: u8,
    behavior_axis: u8,
    axis_field: Option<&PivotField>,
    has_advanced_filter: bool,
) -> std::io::Result<()> {
    let mut payload = Vec::with_capacity(64);
    payload.push(axis & 0x0F);
    let subtotal_flags = if axis == 0x08 {
        0
    } else {
        sxvd_subtotal_flags(axis_field)
    };
    payload.extend_from_slice(&subtotal_flags.to_le_bytes());

    let mut field_flags = 0x10u8;
    if axis_field.is_some_and(|field| !field.show_drop_downs) {
        field_flags |= 0x02;
    }
    if axis_field
        .and_then(|field| field.caption.as_ref())
        .is_some()
    {
        field_flags |= 0x20;
    }
    if axis_field
        .and_then(|field| field.subtotal_caption.as_ref())
        .is_some()
    {
        field_flags |= 0x40;
    }
    payload.push(field_flags);

    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(
        &sxvd_behavior_flags(behavior_axis, axis_field, has_advanced_filter).to_le_bytes(),
    );
    payload.extend_from_slice(&sxvd_item_page_count(axis_field, has_advanced_filter).to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());

    if let Some(caption) = axis_field.and_then(|field| field.caption.as_ref()) {
        payload.extend_from_slice(&encode_wide_str(caption));
    }
    if let Some(caption) = axis_field.and_then(|field| field.subtotal_caption.as_ref()) {
        payload.extend_from_slice(&encode_wide_str(caption));
    }

    rw.write_record(records::BRT_BEGIN_SXVD, &payload)
}

fn validate_xlsb_pivot_axis_field_options(pivot: &PivotTable) -> XlsbResult<()> {
    for field in pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
    {
        if field.item_page_count == 0 {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot field {} uses item page count 0; BIFF12 SXVD requires a positive item-page count",
                field.field.name
            )));
        }
        if field.item_page_count != 10 && pivot_has_advanced_filter(pivot, &field.field.name) {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot field {} changes item page count while using an advanced filter; BIFF12 stores both in the SXVD AutoShow count slot",
                field.field.name
            )));
        }
    }
    Ok(())
}

fn validate_xlsb_pivot_filters(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
) -> XlsbResult<()> {
    for filter in &pivot.filters {
        match filter {
            PivotFilter::FieldItems {
                field,
                allowed_items,
            } if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name)
                .is_some() =>
            {
                if xlsb_field_has_date_derived_axis(cache, layout, &field.name)
                    && pivot_axis_contains_field(&pivot.page_fields, &field.name)
                    && !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters date-grouped page field {}, which this BIFF12 writer slice does not encode yet",
                        pivot.name, field.name
                    )));
                }
                if allowed_items.is_empty() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot field {} requires at least one selected item",
                        field.name
                    )));
                }
            }
            PivotFilter::FieldItems { field, .. } => {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot table {} filters field {} outside the row, column, or page axes, which this BIFF12 writer slice does not encode yet",
                    pivot.name, field.name
                )));
            }
            PivotFilter::TopN {
                field, measure, n, ..
            } => {
                if *n == 0 {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} has a top-N filter on field {} with a zero threshold",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} applies a top-N filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name).is_none()
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters field {} that is not present in the native pivot field layout",
                        pivot.name, field.name
                    )));
                }
                xlsb_pivot_measure_index_for_filter(pivot, measure)?;
            }
            PivotFilter::Label {
                field,
                operator,
                value,
            } if xlsb_supported_label_filter_operator(*operator) => {
                if value.is_empty() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} has an empty label filter on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} applies a label filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name).is_none()
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters field {} that is not present in the native pivot field layout",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::Label { field, .. } => {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot table {} uses a label filter operator for field {} that this BIFF12 writer slice does not encode yet",
                    pivot.name, field.name
                )));
            }
            PivotFilter::LabelBetween {
                field, start, end, ..
            } => {
                if start.is_empty() || end.is_empty() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} has an empty label range filter bound on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} applies a label range filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name).is_none()
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters field {} that is not present in the native pivot field layout",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::Value {
                field,
                measure,
                operator,
                value,
            } if xlsb_supported_value_filter_operator(*operator) => {
                if !value.is_finite() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} has a non-finite value filter threshold on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} applies a value filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name).is_none()
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters field {} that is not present in the native pivot field layout",
                        pivot.name, field.name
                    )));
                }
                xlsb_pivot_measure_index_for_filter(pivot, measure)?;
            }
            PivotFilter::Value { field, .. } => {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot table {} uses a value filter operator for field {} that this BIFF12 writer slice does not encode yet",
                    pivot.name, field.name
                )));
            }
            PivotFilter::ValueBetween {
                field,
                measure,
                start,
                end,
                ..
            } => {
                if !start.is_finite() || !end.is_finite() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} has a non-finite value range filter threshold on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} applies a value range filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name).is_none()
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters field {} that is not present in the native pivot field layout",
                        pivot.name, field.name
                    )));
                }
                xlsb_pivot_measure_index_for_filter(pivot, measure)?;
            }
            PivotFilter::Date {
                field,
                operator,
                value,
            } if xlsb_supported_date_filter_operator(*operator) => {
                if !value.is_finite() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} has a non-finite date filter operand on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} applies a date filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name).is_none()
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters field {} that is not present in the native pivot field layout",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::Date { field, .. } => {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot table {} uses a date filter operator for field {} that this BIFF12 writer slice does not encode yet",
                    pivot.name, field.name
                )));
            }
            PivotFilter::DateBetween {
                field, start, end, ..
            } => {
                if !start.is_finite() || !end.is_finite() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} has a non-finite date range filter operand on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} applies a date range filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name).is_none()
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters field {} that is not present in the native pivot field layout",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::DatePeriod { field, period } => {
                if xlsb_date_period_filter_codes(*period).is_none() {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} uses a date-period filter for field {} that this BIFF12 writer slice does not encode yet",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} applies a date-period filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                if xlsb_layout_field_index_for_semantic_field(cache, layout, &field.name).is_none()
                {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot table {} filters field {} that is not present in the native pivot field layout",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::Unsupported { kind, .. } => {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot table {} contains unsupported preserved filter {kind}, which this BIFF12 writer slice does not encode yet",
                    pivot.name
                )));
            }
        }
    }
    Ok(())
}

fn xlsb_field_has_date_derived_axis(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    field_name: &str,
) -> bool {
    let Some(source_index) = cache.field_index(field_name) else {
        return false;
    };
    layout.date_derived_by_base.contains_key(&source_index)
}

fn pivot_axis_contains_field(fields: &[PivotField], field_name: &str) -> bool {
    fields
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
}

fn xlsb_layout_field_index_for_semantic_field(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    field_name: &str,
) -> Option<usize> {
    if let Some(source_index) = cache.field_index(field_name) {
        return layout.source_to_layout.get(source_index).copied();
    }

    layout
        .fields
        .iter()
        .enumerate()
        .find_map(|(layout_index, field_layout)| {
            let name = match field_layout {
                XlsbPivotCacheFieldLayout::Source { source_index, .. } => {
                    cache.fields.get(*source_index)?.name.as_str()
                }
                XlsbPivotCacheFieldLayout::ManualDerived { name, .. }
                | XlsbPivotCacheFieldLayout::DateUnitDerived { name, .. } => name.as_str(),
            };
            name.eq_ignore_ascii_case(field_name)
                .then_some(layout_index)
        })
}

fn xlsb_pivot_measure_index_for_filter(
    pivot: &PivotTable,
    target: &PivotMeasure,
) -> XlsbResult<u32> {
    pivot
        .measures
        .iter()
        .position(|measure| pivot_measure_matches_sort_target(measure, target))
        .map(|index| index as u32)
        .ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "XLSB pivot table {} filters by an unknown measure {}",
                pivot.name, target.field.name
            ))
        })
}

fn write_sx_field_items<W: Write>(
    rw: &mut RecordWriter<W>,
    item_count: usize,
    axis_field: Option<&PivotField>,
    hidden_items: &HashSet<u32>,
    calculated_items: &HashSet<usize>,
) -> std::io::Result<()> {
    let subtotal_item_count = sxvd_subtotal_item_count(axis_field);
    let count = item_count as u32 + subtotal_item_count as u32;
    rw.write_record(records::BRT_BEGIN_SXVIS, &count.to_le_bytes())?;
    for item_index in 0..item_count as u32 {
        let mut payload = Vec::with_capacity(7);
        payload.push(0);
        let mut flags = u16::from(hidden_items.contains(&item_index));
        if calculated_items.contains(&(item_index as usize)) {
            flags |= 0x0008;
        }
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&item_index.to_le_bytes());
        rw.write_record(records::BRT_BEGIN_SXVI, &payload)?;
        rw.write_record(records::BRT_END_SXVI, &[])?;
    }
    write_sxvd_subtotal_items(rw, axis_field)?;
    rw.write_record(records::BRT_END_SXVIS, &[])
}

fn pivot_axis_field_for_source_index<'a>(
    pivot: &'a PivotTable,
    cache: &FormatPivotCache,
    source_index: usize,
) -> Option<&'a PivotField> {
    pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .find(|field| {
            cache.field_index(&field.field.name) == Some(source_index)
                || field.field.name.eq_ignore_ascii_case(
                    cache
                        .fields
                        .get(source_index)
                        .map(|field| field.name.as_str())
                        .unwrap_or_default(),
                )
        })
}

fn sxvd_sort_measure_index(
    pivot: &PivotTable,
    axis_field: Option<&PivotField>,
) -> XlsbResult<Option<u32>> {
    let Some(axis_field) = axis_field else {
        return Ok(None);
    };
    if matches!(axis_field.sort, PivotSort::None) {
        return Ok(None);
    }
    let Some(sort_measure) = axis_field.sort_by_measure.as_ref() else {
        return Ok(None);
    };

    pivot
        .measures
        .iter()
        .position(|measure| pivot_measure_matches_sort_target(measure, sort_measure))
        .map(|index| index as u32)
        .map(Some)
        .ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "XLSB pivot table {} sorts field {} by an unknown measure",
                pivot.name, axis_field.field.name
            ))
        })
}

fn pivot_measure_matches_sort_target(measure: &PivotMeasure, target: &PivotMeasure) -> bool {
    measure.field.name.eq_ignore_ascii_case(&target.field.name)
        && measure.aggregate == target.aggregate
        && target
            .name
            .as_ref()
            .is_none_or(|name| measure.name.as_ref() == Some(name))
}

fn write_sxvd_auto_sort_scope<W: Write>(
    rw: &mut RecordWriter<W>,
    measure_index: u32,
) -> std::io::Result<()> {
    rw.write_record(records::BRT_BEGIN_SXVD14, &[])?;
    rw.write_record(
        records::BRT_BEGIN_PIVOT_AREA,
        &[0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00, 0x00, 0x00],
    )?;
    rw.write_record(records::BRT_BEGIN_PIVOT_AREA_REFS, &1u32.to_le_bytes())?;

    let mut reference = Vec::with_capacity(11);
    reference.extend_from_slice(&(-2i32).to_le_bytes());
    reference.extend_from_slice(&1u32.to_le_bytes());
    reference.extend_from_slice(&[1, 0, 0]);
    rw.write_record(records::BRT_BEGIN_PIVOT_AREA_REF, &reference)?;
    rw.write_record(
        records::BRT_PIVOT_AREA_REF_ITEM,
        &measure_index.to_le_bytes(),
    )?;
    rw.write_record(records::BRT_END_PIVOT_AREA_REF_ITEM, &[])?;
    rw.write_record(records::BRT_END_PIVOT_AREA_REF, &[])?;

    rw.write_record(records::BRT_END_PIVOT_AREA_REFS, &[])?;
    rw.write_record(records::BRT_END_PIVOT_AREA, &[])?;
    rw.write_record(records::BRT_END_SXVD14, &[])?;
    Ok(())
}

fn sxvd_subtotal_flags(axis_field: Option<&PivotField>) -> u16 {
    let Some(field) = axis_field else {
        return 0x0001;
    };

    sxvd_subtotal_items(field)
        .into_iter()
        .fold(0u16, |flags, subtotal| flags | sxvd_subtotal_flag(subtotal))
}

fn sxvd_subtotal_flag(subtotal: PivotSubtotal) -> u16 {
    match subtotal {
        PivotSubtotal::Automatic => 0x0001,
        PivotSubtotal::Sum => 0x0002,
        PivotSubtotal::Count => 0x0004,
        PivotSubtotal::Average => 0x0008,
        PivotSubtotal::Max => 0x0010,
        PivotSubtotal::Min => 0x0020,
        PivotSubtotal::Product => 0x0040,
        PivotSubtotal::CountNumbers => 0x0080,
        PivotSubtotal::StdDev => 0x0100,
        PivotSubtotal::StdDevP => 0x0200,
        PivotSubtotal::Var => 0x0400,
        PivotSubtotal::VarP => 0x0800,
        PivotSubtotal::None => 0x0000,
    }
}

fn sxvd_behavior_flags(
    axis: u8,
    axis_field: Option<&PivotField>,
    has_advanced_filter: bool,
) -> u32 {
    if axis == 0x08 {
        return 0x0004_A158;
    }

    let mut flags = if axis & 0x07 != 0 {
        0x0004_B15F
    } else {
        0x0004_817F
    };

    let Some(field) = axis_field else {
        if axis & 0x07 != 0 {
            flags &= !(1 << 12);
        }
        return flags;
    };

    set_flag(&mut flags, 5, field.show_empty_items);
    set_flag(&mut flags, 7, field.insert_blank_row);
    set_flag(&mut flags, 8, field.subtotal_top);
    set_flag(&mut flags, 11, field.insert_page_break);
    set_flag(&mut flags, 18, !field.include_new_items_in_filter);

    match field.sort {
        PivotSort::None => {
            set_flag(&mut flags, 12, false);
            set_flag(&mut flags, 13, true);
        }
        PivotSort::Ascending => {
            set_flag(&mut flags, 12, true);
            set_flag(&mut flags, 13, true);
        }
        PivotSort::Descending => {
            set_flag(&mut flags, 12, true);
            set_flag(&mut flags, 13, false);
        }
    }

    set_flag(&mut flags, 17, has_advanced_filter);

    flags
}

fn sxvd_item_page_count(axis_field: Option<&PivotField>, has_advanced_filter: bool) -> u32 {
    if has_advanced_filter {
        10
    } else {
        axis_field.map_or(10, |field| field.item_page_count)
    }
}

fn pivot_has_advanced_filter(pivot: &PivotTable, field_name: &str) -> bool {
    pivot.filters.iter().any(|filter| {
        matches!(
            filter,
            PivotFilter::TopN { field, .. } if field.name.eq_ignore_ascii_case(field_name)
        ) || matches!(
            filter,
            PivotFilter::Label {
                field,
                operator,
                ..
            } if field.name.eq_ignore_ascii_case(field_name)
                && xlsb_supported_label_filter_operator(*operator)
        ) || matches!(
            filter,
            PivotFilter::LabelBetween { field, .. }
                if field.name.eq_ignore_ascii_case(field_name)
        ) || matches!(
            filter,
            PivotFilter::Value {
                field,
                operator,
                ..
            } if field.name.eq_ignore_ascii_case(field_name)
                && xlsb_supported_value_filter_operator(*operator)
        ) || matches!(
            filter,
            PivotFilter::ValueBetween { field, .. }
                if field.name.eq_ignore_ascii_case(field_name)
        ) || matches!(
            filter,
            PivotFilter::Date {
                field,
                operator,
                ..
            } if field.name.eq_ignore_ascii_case(field_name)
                && xlsb_supported_date_filter_operator(*operator)
        ) || matches!(
            filter,
            PivotFilter::DateBetween { field, .. } | PivotFilter::DatePeriod { field, .. }
                if field.name.eq_ignore_ascii_case(field_name)
        )
    })
}

fn pivot_has_advanced_filter_for_source_index(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    source_index: usize,
) -> bool {
    pivot.filters.iter().any(|filter| {
        let matches_source = filter_field_name(filter)
            .is_some_and(|name| cache.field_index(name) == Some(source_index));
        matches_source
            && (matches!(filter, PivotFilter::TopN { .. })
                || matches!(
                    filter,
                    PivotFilter::Label { operator, .. }
                        if xlsb_supported_label_filter_operator(*operator)
                )
                || matches!(filter, PivotFilter::LabelBetween { .. })
                || matches!(
                    filter,
                    PivotFilter::Value { operator, .. }
                        if xlsb_supported_value_filter_operator(*operator)
                )
                || matches!(filter, PivotFilter::ValueBetween { .. })
                || matches!(
                    filter,
                    PivotFilter::Date { operator, .. }
                        if xlsb_supported_date_filter_operator(*operator)
                )
                || matches!(
                    filter,
                    PivotFilter::DateBetween { .. } | PivotFilter::DatePeriod { .. }
                ))
    })
}

fn xlsb_supported_label_filter_operator(operator: PivotFilterOperator) -> bool {
    matches!(
        operator,
        PivotFilterOperator::Equals
            | PivotFilterOperator::NotEquals
            | PivotFilterOperator::LessThan
            | PivotFilterOperator::LessThanOrEqual
            | PivotFilterOperator::GreaterThan
            | PivotFilterOperator::GreaterThanOrEqual
            | PivotFilterOperator::BeginsWith
            | PivotFilterOperator::DoesNotBeginWith
            | PivotFilterOperator::EndsWith
            | PivotFilterOperator::DoesNotEndWith
            | PivotFilterOperator::Contains
            | PivotFilterOperator::DoesNotContain
    )
}

fn xlsb_supported_value_filter_operator(operator: PivotFilterOperator) -> bool {
    matches!(
        operator,
        PivotFilterOperator::Equals
            | PivotFilterOperator::NotEquals
            | PivotFilterOperator::LessThan
            | PivotFilterOperator::LessThanOrEqual
            | PivotFilterOperator::GreaterThan
            | PivotFilterOperator::GreaterThanOrEqual
    )
}

fn xlsb_supported_date_filter_operator(operator: PivotFilterOperator) -> bool {
    xlsb_date_filter_type_and_operator(operator).is_some()
}

fn set_flag(flags: &mut u32, bit: u8, value: bool) {
    let mask = 1u32 << bit;
    if value {
        *flags |= mask;
    } else {
        *flags &= !mask;
    }
}

fn sxvd_subtotal_item_count(axis_field: Option<&PivotField>) -> usize {
    axis_field
        .map(sxvd_subtotal_items)
        .unwrap_or_else(|| vec![PivotSubtotal::Automatic])
        .len()
}

fn write_sxvd_subtotal_items<W: Write>(
    rw: &mut RecordWriter<W>,
    axis_field: Option<&PivotField>,
) -> std::io::Result<()> {
    let Some(field) = axis_field else {
        return write_sxvd_subtotal_item(rw, PivotSubtotal::Automatic);
    };

    for subtotal in sxvd_subtotal_items(field) {
        write_sxvd_subtotal_item(rw, subtotal)?;
    }
    Ok(())
}

fn sxvd_subtotal_items(field: &PivotField) -> Vec<PivotSubtotal> {
    if field.subtotals.is_empty() {
        return match field.subtotal {
            PivotSubtotal::None => Vec::new(),
            subtotal => vec![subtotal],
        };
    }

    let custom = field
        .subtotals
        .iter()
        .copied()
        .filter(|subtotal| subtotal.is_custom_function())
        .collect::<Vec<_>>();
    if !custom.is_empty() {
        return custom;
    }

    if field
        .subtotals
        .iter()
        .any(|subtotal| matches!(subtotal, PivotSubtotal::Automatic))
    {
        vec![PivotSubtotal::Automatic]
    } else {
        Vec::new()
    }
}

fn write_sxvd_subtotal_item<W: Write>(
    rw: &mut RecordWriter<W>,
    subtotal: PivotSubtotal,
) -> std::io::Result<()> {
    let mut payload = Vec::with_capacity(7);
    payload.push(sxvd_subtotal_item_type(subtotal));
    payload.extend_from_slice(&[0, 0]);
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    rw.write_record(records::BRT_BEGIN_SXVI, &payload)?;
    rw.write_record(records::BRT_END_SXVI, &[])
}

fn sxvd_subtotal_item_type(subtotal: PivotSubtotal) -> u8 {
    match subtotal {
        PivotSubtotal::Automatic => 0x01,
        PivotSubtotal::Sum => 0x02,
        PivotSubtotal::Count => 0x03,
        PivotSubtotal::Average => 0x04,
        PivotSubtotal::Max => 0x05,
        PivotSubtotal::Min => 0x06,
        PivotSubtotal::Product => 0x07,
        PivotSubtotal::CountNumbers => 0x08,
        PivotSubtotal::StdDev => 0x09,
        PivotSubtotal::StdDevP => 0x0A,
        PivotSubtotal::Var => 0x0B,
        PivotSubtotal::VarP => 0x0C,
        PivotSubtotal::None => 0x00,
    }
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

fn write_olap_axis_hierarchies<W: Write>(
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

    let mut indexes = olap_axis_hierarchy_indexes(cache, layout, fields)?;
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

fn olap_axis_hierarchy_indexes(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
) -> XlsbResult<Vec<i32>> {
    let mut indexes = Vec::with_capacity(fields.len());
    for field in fields {
        let source_index = cache.field_index(&field.field.name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "XLSB OLAP pivot axis references unknown cache field: {}",
                field.field.name
            ))
        })?;
        let hierarchy_index = *layout.source_to_layout.get(source_index).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "XLSB OLAP pivot axis references unmapped cache field: {}",
                field.field.name
            ))
        })?;
        indexes.push(checked_i32(
            hierarchy_index,
            "OLAP pivot hierarchy axis index",
        )?);
    }
    Ok(indexes)
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
    base_tuples: &[Vec<u32>],
) -> XlsbResult<()> {
    let line_tuples = axis_line_tuples(
        fields,
        include_values_field,
        values_position,
        measure_count,
        base_tuples,
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
    fields: &[PivotField],
    include_values_field: bool,
    values_position: Option<u32>,
    measure_count: usize,
    base_tuples: &[Vec<u32>],
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

    if include_values_field {
        Ok(tuples_with_data_items(
            base_tuples.to_vec(),
            values_position,
            measure_count,
        ))
    } else {
        Ok(base_tuples
            .iter()
            .cloned()
            .map(|tuple| (tuple, None))
            .collect())
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
    layout: &XlsbPivotCacheLayout,
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
) -> XlsbResult<()> {
    let page_field_selections =
        expanded_page_field_selection_items(pivot, cache, layout, grouping_infos)?;
    if page_field_selections.is_empty() {
        return Ok(());
    }
    rw.write_record(
        records::BRT_BEGIN_SXPIS,
        &(page_field_selections.len() as u32).to_le_bytes(),
    )?;
    for selection in page_field_selections {
        let selected_item = selected_page_item_index(
            pivot,
            cache,
            &selection.field_name,
            selection.source_index,
            selection.items,
        )?
        .unwrap_or(0x0010_00FE);
        let mut payload = Vec::new();
        payload.extend_from_slice(&(selection.layout_index as u32).to_le_bytes());
        payload.extend_from_slice(&selected_item.to_le_bytes());
        payload.extend_from_slice(&(-1i32).to_le_bytes());
        payload.push(0);
        rw.write_record(records::BRT_BEGIN_SXPI, &payload)?;
        rw.write_record(records::BRT_END_SXPI, &[])?;
    }
    rw.write_record(records::BRT_END_SXPIS, &[])?;
    Ok(())
}

struct PageFieldSelection<'a> {
    field_name: String,
    source_index: Option<usize>,
    layout_index: usize,
    items: &'a [PivotValue],
}

fn expanded_page_field_selection_items<'a>(
    pivot: &PivotTable,
    cache: &'a FormatPivotCache,
    layout: &'a XlsbPivotCacheLayout,
    grouping_infos: &'a [XlsbPivotGroupingInfo<'_>],
) -> XlsbResult<Vec<PageFieldSelection<'a>>> {
    let mut selections = Vec::new();
    for field in &pivot.page_fields {
        selections.extend(page_field_selection_items(
            cache,
            layout,
            grouping_infos,
            &field.field.name,
        )?);
    }
    for source_index in synthetic_consolidation_page_source_indexes(pivot, cache) {
        let field = &cache.fields[source_index];
        selections.push(PageFieldSelection {
            field_name: field.name.clone(),
            source_index: Some(source_index),
            layout_index: layout.source_to_layout[source_index],
            items: &field.shared_items,
        });
    }
    Ok(selections)
}

fn page_field_selection_items<'a>(
    cache: &'a FormatPivotCache,
    layout: &'a XlsbPivotCacheLayout,
    grouping_infos: &'a [XlsbPivotGroupingInfo<'_>],
    field_name: &str,
) -> XlsbResult<Vec<PageFieldSelection<'a>>> {
    let source_index = cache.field_index(field_name).ok_or_else(|| {
        XlsbError::InvalidFormat(format!("pivot references unknown page field {field_name}"))
    })?;
    if let Some(derived_index) = layout.manual_derived_by_base.get(&source_index) {
        let Some(XlsbPivotCacheFieldLayout::ManualDerived {
            grouping_info_index,
            ..
        }) = layout.fields.get(*derived_index)
        else {
            return Err(XlsbError::InvalidFormat(
                "pivot manual derived page field has invalid layout".into(),
            ));
        };
        return Ok(vec![PageFieldSelection {
            field_name: field_name.to_string(),
            source_index: Some(source_index),
            layout_index: *derived_index,
            items: &grouping_infos[*grouping_info_index].group_items,
        }]);
    }
    if let Some(derived_indexes) = layout.date_derived_by_base.get(&source_index) {
        let mut selections = Vec::with_capacity(derived_indexes.len());
        for derived_index in derived_indexes {
            let Some(XlsbPivotCacheFieldLayout::DateUnitDerived {
                grouping_info_index,
                name,
                ..
            }) = layout.fields.get(*derived_index)
            else {
                return Err(XlsbError::InvalidFormat(
                    "pivot date derived page field has invalid layout".into(),
                ));
            };
            selections.push(PageFieldSelection {
                field_name: name.clone(),
                source_index: Some(source_index),
                layout_index: *derived_index,
                items: &grouping_infos[*grouping_info_index].group_items,
            });
        }
        return Ok(selections);
    }
    if let Some(grouping_info) = grouping_info_for_field(grouping_infos, field_name) {
        if !is_multi_unit_date_grouping(Some(grouping_info.grouping)) {
            return Ok(vec![PageFieldSelection {
                field_name: field_name.to_string(),
                source_index: Some(source_index),
                layout_index: layout.source_to_layout[source_index],
                items: &grouping_info.group_items,
            }]);
        }
    }

    Ok(vec![PageFieldSelection {
        field_name: field_name.to_string(),
        source_index: Some(source_index),
        layout_index: layout.source_to_layout[source_index],
        items: &cache.fields[source_index].shared_items,
    }])
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
    styles: &StyleMapping,
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
        let caption = measure.caption();
        let data_field = xlsb_data_field_options(measure, cache, styles)?;
        let mut payload = Vec::new();
        payload.extend_from_slice(&(field_index as u32).to_le_bytes());
        payload.extend_from_slice(&aggregate_code(measure.aggregate).to_le_bytes());
        payload.extend_from_slice(&data_field.show_as.to_le_bytes());
        payload.extend_from_slice(&data_field.base_field.to_le_bytes());
        payload.extend_from_slice(&data_field.base_item.to_le_bytes());
        payload.extend_from_slice(&data_field.num_format.to_le_bytes());
        payload.push(u8::from(!caption.is_empty()));
        if !caption.is_empty() {
            payload.extend_from_slice(&encode_wide_str(&caption));
        }
        rw.write_record(records::BRT_BEGIN_SXDI, &payload)?;
        if let Some(show_as_ext) = data_field.show_as_ext {
            write_data_field_ext(rw, show_as_ext)?;
        }
        rw.write_record(records::BRT_END_SXDI, &[])?;
    }
    rw.write_record(records::BRT_END_SXDIS, &[])?;
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct XlsbDataFieldOptions {
    show_as: u32,
    show_as_ext: Option<u32>,
    base_field: u32,
    base_item: u32,
    num_format: u32,
}

fn xlsb_data_field_options(
    measure: &PivotMeasure,
    cache: &FormatPivotCache,
    styles: &StyleMapping,
) -> XlsbResult<XlsbDataFieldOptions> {
    let (show_as, show_as_ext, base_field, base_item) = match &measure.show_as {
        PivotShowAs::Normal => (0, None, None, None),
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        } => (1, None, Some(base_field), Some(base_item)),
        PivotShowAs::PercentDifferenceFrom {
            base_field,
            base_item,
        } => (3, None, Some(base_field), Some(base_item)),
        PivotShowAs::RunningTotal { base_field } => (4, None, Some(base_field), None),
        PivotShowAs::PercentOfRowTotal => (5, None, None, None),
        PivotShowAs::PercentOfColumnTotal => (6, None, None, None),
        PivotShowAs::PercentOfGrandTotal => (7, None, None, None),
        PivotShowAs::Index => (8, None, None, None),
        PivotShowAs::PercentOfParentRowTotal => (0, Some(9), None, None),
        PivotShowAs::PercentOfParentColumnTotal => (0, Some(10), None, None),
        PivotShowAs::PercentOfParentTotal { base_field } => (0, Some(11), Some(base_field), None),
        PivotShowAs::RankAscending { base_field } => (0, Some(13), Some(base_field), None),
        PivotShowAs::RankDescending { base_field } => (0, Some(14), Some(base_field), None),
    };
    let base_field_index = if let Some(base_field) = base_field {
        cache.field_index(&base_field.name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "pivot show-as base field not found: {}",
                base_field.name
            ))
        })?
    } else {
        0
    };
    let base_item_index = if let Some(base_item) = base_item {
        let field = cache.fields.get(base_field_index).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "pivot show-as base field index out of range: {base_field_index}"
            ))
        })?;
        field
            .shared_items
            .iter()
            .position(|candidate| candidate == base_item)
            .ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "pivot show-as base item not found in field {}: {base_item}",
                    field.name
                ))
            })?
    } else {
        0
    };

    Ok(XlsbDataFieldOptions {
        show_as,
        show_as_ext,
        base_field: checked_u32(base_field_index, "pivot show-as base field index")?,
        base_item: checked_u32(base_item_index, "pivot show-as base item index")?,
        num_format: pivot_measure_number_format_id(measure, styles)?,
    })
}

fn write_data_field_ext<W: Write>(
    rw: &mut RecordWriter<W>,
    show_as_ext: u32,
) -> std::io::Result<()> {
    rw.write_record(records::BRT_BEGIN_FRT, &0x0000_0E02u32.to_le_bytes())?;
    let mut payload = Vec::with_capacity(13);
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&show_as_ext.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.push(0);
    rw.write_record(records::BRT_SXDI14, &payload)?;
    rw.write_record(records::BRT_END_FRT, &[])
}

fn pivot_measure_number_format_id(
    measure: &PivotMeasure,
    styles: &StyleMapping,
) -> XlsbResult<u32> {
    let Some(number_format) = measure.number_format.as_deref() else {
        return Ok(0);
    };
    if let Some(id) = builtin_number_format_id(number_format) {
        return Ok(id);
    }
    styles
        .custom_numfmt_id(number_format)
        .map(u32::from)
        .ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "XLSB pivot measure number format was not registered: {number_format}"
            ))
        })
}

fn builtin_number_format_id(format_code: &str) -> Option<u32> {
    if format_code.eq_ignore_ascii_case("General") {
        return Some(NumberFormat::ID_GENERAL);
    }
    (1..=49).find(|id| NumberFormat::BuiltIn(*id).format_string() == format_code)
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

fn write_pivot_filters<W: Write>(
    rw: &mut RecordWriter<W>,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    date_system: DateSystem,
) -> XlsbResult<()> {
    let writable_filters = pivot
        .filters
        .iter()
        .filter_map(|filter| match filter {
            PivotFilter::TopN {
                field,
                measure,
                n,
                top,
                percent,
            } => Some(XlsbWritablePivotFilter::TopN {
                field_name: field.name.as_str(),
                measure,
                n: *n,
                top: *top,
                percent: *percent,
            }),
            PivotFilter::Label {
                field,
                operator,
                value,
            } if xlsb_supported_label_filter_operator(*operator) => {
                Some(XlsbWritablePivotFilter::Label {
                    field_name: field.name.as_str(),
                    operator: *operator,
                    value: value.as_str(),
                })
            }
            PivotFilter::LabelBetween {
                field,
                start,
                end,
                not_between,
            } => Some(XlsbWritablePivotFilter::LabelBetween {
                field_name: field.name.as_str(),
                start: start.as_str(),
                end: end.as_str(),
                not_between: *not_between,
            }),
            PivotFilter::Value {
                field,
                measure,
                operator,
                value,
            } if xlsb_supported_value_filter_operator(*operator) => {
                Some(XlsbWritablePivotFilter::Value {
                    field_name: field.name.as_str(),
                    measure,
                    operator: *operator,
                    value: *value,
                })
            }
            PivotFilter::ValueBetween {
                field,
                measure,
                start,
                end,
                not_between,
            } => Some(XlsbWritablePivotFilter::ValueBetween {
                field_name: field.name.as_str(),
                measure,
                start: *start,
                end: *end,
                not_between: *not_between,
            }),
            PivotFilter::Date {
                field,
                operator,
                value,
            } if xlsb_supported_date_filter_operator(*operator) => {
                Some(XlsbWritablePivotFilter::Date {
                    field_name: field.name.as_str(),
                    operator: *operator,
                    value: *value,
                })
            }
            PivotFilter::DateBetween {
                field,
                start,
                end,
                not_between,
            } => Some(XlsbWritablePivotFilter::DateBetween {
                field_name: field.name.as_str(),
                start: *start,
                end: *end,
                not_between: *not_between,
            }),
            PivotFilter::DatePeriod { field, period } => {
                Some(XlsbWritablePivotFilter::DatePeriod {
                    field_name: field.name.as_str(),
                    period: *period,
                })
            }
            _ => None,
        })
        .collect::<Vec<_>>();
    if writable_filters.is_empty() {
        return Ok(());
    }

    rw.write_record(
        records::BRT_BEGIN_SX_FILTERS,
        &(writable_filters.len() as u32).to_le_bytes(),
    )?;
    for (index, filter) in writable_filters.into_iter().enumerate() {
        let field_name = filter.field_name();
        let field_index = xlsb_layout_field_index_for_semantic_field(cache, layout, field_name)
            .ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "XLSB pivot table {} filters unknown field {}",
                    pivot.name, field_name
                ))
            })? as u32;
        match filter {
            XlsbWritablePivotFilter::TopN {
                measure,
                n,
                top,
                percent,
                ..
            } => {
                let measure_index = xlsb_pivot_measure_index_for_filter(pivot, measure)?;
                write_pivot_top_n_filter(
                    rw,
                    field_index,
                    measure_index,
                    (index + 1) as u32,
                    n,
                    top,
                    percent,
                )?;
            }
            XlsbWritablePivotFilter::Label {
                operator, value, ..
            } => {
                write_pivot_label_filter(rw, field_index, (index + 1) as u32, operator, value)?;
            }
            XlsbWritablePivotFilter::LabelBetween {
                start,
                end,
                not_between,
                ..
            } => {
                write_pivot_label_between_filter(
                    rw,
                    field_index,
                    (index + 1) as u32,
                    start,
                    end,
                    not_between,
                )?;
            }
            XlsbWritablePivotFilter::Value {
                measure,
                operator,
                value,
                ..
            } => {
                let measure_index = xlsb_pivot_measure_index_for_filter(pivot, measure)?;
                write_pivot_value_filter(
                    rw,
                    field_index,
                    measure_index,
                    (index + 1) as u32,
                    operator,
                    value,
                )?;
            }
            XlsbWritablePivotFilter::ValueBetween {
                measure,
                start,
                end,
                not_between,
                ..
            } => {
                let measure_index = xlsb_pivot_measure_index_for_filter(pivot, measure)?;
                write_pivot_value_between_filter(
                    rw,
                    field_index,
                    measure_index,
                    (index + 1) as u32,
                    start,
                    end,
                    not_between,
                )?;
            }
            XlsbWritablePivotFilter::Date {
                operator, value, ..
            } => {
                write_pivot_date_filter(rw, field_index, (index + 1) as u32, operator, value)?;
            }
            XlsbWritablePivotFilter::DateBetween {
                start,
                end,
                not_between,
                ..
            } => {
                write_pivot_date_between_filter(
                    rw,
                    field_index,
                    (index + 1) as u32,
                    start,
                    end,
                    not_between,
                )?;
            }
            XlsbWritablePivotFilter::DatePeriod { period, .. } => {
                write_pivot_date_period_filter(
                    rw,
                    field_index,
                    (index + 1) as u32,
                    period,
                    date_system,
                )?;
            }
        }
    }
    rw.write_record(records::BRT_END_SX_FILTERS, &[])?;
    Ok(())
}

enum XlsbWritablePivotFilter<'a> {
    TopN {
        field_name: &'a str,
        measure: &'a PivotMeasure,
        n: u32,
        top: bool,
        percent: bool,
    },
    Label {
        field_name: &'a str,
        operator: PivotFilterOperator,
        value: &'a str,
    },
    LabelBetween {
        field_name: &'a str,
        start: &'a str,
        end: &'a str,
        not_between: bool,
    },
    Value {
        field_name: &'a str,
        measure: &'a PivotMeasure,
        operator: PivotFilterOperator,
        value: f64,
    },
    ValueBetween {
        field_name: &'a str,
        measure: &'a PivotMeasure,
        start: f64,
        end: f64,
        not_between: bool,
    },
    Date {
        field_name: &'a str,
        operator: PivotFilterOperator,
        value: f64,
    },
    DateBetween {
        field_name: &'a str,
        start: f64,
        end: f64,
        not_between: bool,
    },
    DatePeriod {
        field_name: &'a str,
        period: PivotDatePeriod,
    },
}

impl XlsbWritablePivotFilter<'_> {
    fn field_name(&self) -> &str {
        match self {
            Self::TopN { field_name, .. }
            | Self::Label { field_name, .. }
            | Self::LabelBetween { field_name, .. }
            | Self::Value { field_name, .. }
            | Self::ValueBetween { field_name, .. }
            | Self::Date { field_name, .. }
            | Self::DateBetween { field_name, .. }
            | Self::DatePeriod { field_name, .. } => field_name,
        }
    }
}

fn write_pivot_top_n_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    measure_index: u32,
    filter_id: u32,
    n: u32,
    top: bool,
    percent: bool,
) -> std::io::Result<()> {
    let mut payload = Vec::with_capacity(30);
    payload.extend_from_slice(&field_index.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&(if percent { 2u32 } else { 1u32 }).to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_id.to_le_bytes());
    payload.extend_from_slice(&measure_index.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_SX_FILTER, &payload)?;

    write_ac_block(rw, 0, false)?;
    rw.write_record(records::BRT_BEGIN_A_FILTER, &[0; 16])?;
    rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &[0; 6])?;

    let mut top_payload = Vec::with_capacity(17);
    let mut flags = 0x04u8;
    if top {
        flags |= 0x01;
    }
    if percent {
        flags |= 0x02;
    }
    let value = n as f64;
    top_payload.push(flags);
    top_payload.extend_from_slice(&value.to_le_bytes());
    top_payload.extend_from_slice(&value.to_le_bytes());
    rw.write_record(records::BRT_TOP10_FILTER, &top_payload)?;

    rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    rw.write_record(records::BRT_END_A_FILTER, &[])?;
    rw.write_record(records::BRT_END_SX_FILTER, &[])?;
    Ok(())
}

fn write_pivot_label_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    filter_id: u32,
    operator: PivotFilterOperator,
    value: &str,
) -> std::io::Result<()> {
    let filter_type = match operator {
        PivotFilterOperator::Equals => 4u32,
        PivotFilterOperator::NotEquals => 5u32,
        PivotFilterOperator::BeginsWith => 6u32,
        PivotFilterOperator::DoesNotBeginWith => 7u32,
        PivotFilterOperator::EndsWith => 8u32,
        PivotFilterOperator::DoesNotEndWith => 9u32,
        PivotFilterOperator::Contains => 10u32,
        PivotFilterOperator::DoesNotContain => 11u32,
        PivotFilterOperator::GreaterThan => 12u32,
        PivotFilterOperator::GreaterThanOrEqual => 13u32,
        PivotFilterOperator::LessThan => 14u32,
        PivotFilterOperator::LessThanOrEqual => 15u32,
    };
    let mut payload = Vec::with_capacity(30 + value.len() * 2);
    payload.extend_from_slice(&field_index.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_type.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_id.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&4u16.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(value));
    rw.write_record(records::BRT_BEGIN_SX_FILTER, &payload)?;

    write_ac_block(rw, 0, false)?;
    rw.write_record(records::BRT_BEGIN_A_FILTER, &[0; 16])?;
    rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &[0; 6])?;
    if let Some(custom) = label_filter_custom_filter(operator, value) {
        rw.write_record(records::BRT_BEGIN_CUSTOM_FILTERS, &0i32.to_le_bytes())?;

        let mut custom_payload = Vec::with_capacity(10 + custom.value.len() * 2);
        custom_payload.push(6u8);
        custom_payload.push(custom.operator_code);
        custom_payload.push(custom.discriminator);
        custom_payload.push(0u8);
        custom_payload.extend_from_slice(&[0u8; 6]);
        custom_payload.extend_from_slice(&encode_wide_str(custom.value.as_str()));
        rw.write_record(records::BRT_CUSTOM_FILTER, &custom_payload)?;

        rw.write_record(records::BRT_END_CUSTOM_FILTERS, &[])?;
    }
    rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    rw.write_record(records::BRT_END_A_FILTER, &[])?;
    rw.write_record(records::BRT_END_SX_FILTER, &[])?;
    Ok(())
}

fn write_pivot_label_between_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    filter_id: u32,
    start: &str,
    end: &str,
    not_between: bool,
) -> std::io::Result<()> {
    let filter_type = if not_between { 17u32 } else { 16u32 };
    let mut payload = Vec::with_capacity(30 + start.len() * 2);
    payload.extend_from_slice(&field_index.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_type.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_id.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&4u16.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(start));
    rw.write_record(records::BRT_BEGIN_SX_FILTER, &payload)?;

    write_ac_block(rw, 0, false)?;
    rw.write_record(records::BRT_BEGIN_A_FILTER, &[0; 16])?;
    rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &[0; 6])?;
    rw.write_record(
        records::BRT_BEGIN_CUSTOM_FILTERS,
        &(if not_between { 0i32 } else { 1i32 }).to_le_bytes(),
    )?;
    if not_between {
        write_xlsb_string_custom_filter(rw, 1, start)?;
        write_xlsb_string_custom_filter(rw, 4, end)?;
    } else {
        write_xlsb_string_custom_filter(rw, 6, start)?;
        write_xlsb_string_custom_filter(rw, 3, end)?;
    }
    rw.write_record(records::BRT_END_CUSTOM_FILTERS, &[])?;
    rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    rw.write_record(records::BRT_END_A_FILTER, &[])?;
    rw.write_record(records::BRT_END_SX_FILTER, &[])?;
    Ok(())
}

fn write_pivot_value_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    measure_index: u32,
    filter_id: u32,
    operator: PivotFilterOperator,
    value: f64,
) -> std::io::Result<()> {
    let (filter_type, custom_operator) = value_filter_type_and_operator(operator);
    write_pivot_value_filter_header(rw, field_index, measure_index, filter_id, filter_type)?;
    write_ac_block(rw, 0, false)?;
    rw.write_record(records::BRT_BEGIN_A_FILTER, &[0; 16])?;
    rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &[0; 6])?;
    rw.write_record(records::BRT_BEGIN_CUSTOM_FILTERS, &0i32.to_le_bytes())?;
    write_xlsb_numeric_custom_filter(rw, custom_operator, value)?;
    rw.write_record(records::BRT_END_CUSTOM_FILTERS, &[])?;
    rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    rw.write_record(records::BRT_END_A_FILTER, &[])?;
    rw.write_record(records::BRT_END_SX_FILTER, &[])?;
    Ok(())
}

fn write_pivot_value_between_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    measure_index: u32,
    filter_id: u32,
    start: f64,
    end: f64,
    not_between: bool,
) -> std::io::Result<()> {
    let filter_type = if not_between { 25 } else { 24 };
    write_pivot_value_filter_header(rw, field_index, measure_index, filter_id, filter_type)?;
    write_ac_block(rw, 0, false)?;
    rw.write_record(records::BRT_BEGIN_A_FILTER, &[0; 16])?;
    rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &[0; 6])?;
    rw.write_record(
        records::BRT_BEGIN_CUSTOM_FILTERS,
        &(if not_between { 0i32 } else { 1i32 }).to_le_bytes(),
    )?;
    if not_between {
        write_xlsb_numeric_custom_filter(rw, 1, start)?;
        write_xlsb_numeric_custom_filter(rw, 4, end)?;
    } else {
        write_xlsb_numeric_custom_filter(rw, 6, start)?;
        write_xlsb_numeric_custom_filter(rw, 3, end)?;
    }
    rw.write_record(records::BRT_END_CUSTOM_FILTERS, &[])?;
    rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    rw.write_record(records::BRT_END_A_FILTER, &[])?;
    rw.write_record(records::BRT_END_SX_FILTER, &[])?;
    Ok(())
}

fn write_pivot_date_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    filter_id: u32,
    operator: PivotFilterOperator,
    value: f64,
) -> std::io::Result<()> {
    let (filter_type, custom_operator) =
        xlsb_date_filter_type_and_operator(operator).expect("validated date filter operator");
    write_pivot_date_filter_header(rw, field_index, filter_id, filter_type)?;
    write_ac_block(rw, 0, false)?;
    rw.write_record(records::BRT_BEGIN_A_FILTER, &[0; 16])?;
    rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &[0; 6])?;
    rw.write_record(records::BRT_BEGIN_CUSTOM_FILTERS, &0i32.to_le_bytes())?;
    write_xlsb_numeric_custom_filter(rw, custom_operator, value)?;
    rw.write_record(records::BRT_END_CUSTOM_FILTERS, &[])?;
    rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    rw.write_record(records::BRT_END_A_FILTER, &[])?;
    rw.write_record(records::BRT_END_SX_FILTER, &[])?;
    Ok(())
}

fn write_pivot_date_between_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    filter_id: u32,
    start: f64,
    end: f64,
    not_between: bool,
) -> std::io::Result<()> {
    let filter_type = if not_between { 65 } else { 29 };
    write_pivot_date_filter_header(rw, field_index, filter_id, filter_type)?;
    write_ac_block(rw, 0, false)?;
    rw.write_record(records::BRT_BEGIN_A_FILTER, &[0; 16])?;
    rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &[0; 6])?;
    rw.write_record(
        records::BRT_BEGIN_CUSTOM_FILTERS,
        &(if not_between { 0i32 } else { 1i32 }).to_le_bytes(),
    )?;
    if not_between {
        write_xlsb_numeric_custom_filter(rw, 1, start)?;
        write_xlsb_numeric_custom_filter(rw, 4, end)?;
    } else {
        write_xlsb_numeric_custom_filter(rw, 6, start)?;
        write_xlsb_numeric_custom_filter(rw, 3, end)?;
    }
    rw.write_record(records::BRT_END_CUSTOM_FILTERS, &[])?;
    rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    rw.write_record(records::BRT_END_A_FILTER, &[])?;
    rw.write_record(records::BRT_END_SX_FILTER, &[])?;
    Ok(())
}

fn write_pivot_date_period_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    filter_id: u32,
    period: PivotDatePeriod,
    date_system: DateSystem,
) -> std::io::Result<()> {
    let (filter_type, cft) =
        xlsb_date_period_filter_codes(period).expect("validated date period filter");
    write_pivot_date_filter_header(rw, field_index, filter_id, filter_type)?;
    write_ac_block(rw, 0, false)?;
    rw.write_record(records::BRT_BEGIN_A_FILTER, &[0; 16])?;
    rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &[0; 6])?;

    let (value, max_value) =
        pivot_date_period_filter_bounds(period, date_system).unwrap_or((0.0, 0.0));
    let mut payload = Vec::with_capacity(21);
    payload.extend_from_slice(&cft.to_le_bytes());
    payload.push(if value != 0.0 || max_value != 0.0 {
        1
    } else {
        0
    });
    payload.extend_from_slice(&value.to_le_bytes());
    payload.extend_from_slice(&max_value.to_le_bytes());
    rw.write_record(records::BRT_DYNAMIC_FILTER, &payload)?;

    rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    rw.write_record(records::BRT_END_A_FILTER, &[])?;
    rw.write_record(records::BRT_END_SX_FILTER, &[])?;
    Ok(())
}

fn write_pivot_date_filter_header<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    filter_id: u32,
    filter_type: u32,
) -> std::io::Result<()> {
    let mut payload = Vec::with_capacity(30);
    payload.extend_from_slice(&field_index.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_type.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_id.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_SX_FILTER, &payload)
}

fn write_pivot_value_filter_header<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: u32,
    measure_index: u32,
    filter_id: u32,
    filter_type: u32,
) -> std::io::Result<()> {
    let mut payload = Vec::with_capacity(30);
    payload.extend_from_slice(&field_index.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_type.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&filter_id.to_le_bytes());
    payload.extend_from_slice(&measure_index.to_le_bytes());
    payload.extend_from_slice(&(-1i32).to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_SX_FILTER, &payload)
}

fn value_filter_type_and_operator(operator: PivotFilterOperator) -> (u32, u8) {
    match operator {
        PivotFilterOperator::Equals => (18, 2),
        PivotFilterOperator::NotEquals => (19, 5),
        PivotFilterOperator::GreaterThan => (20, 4),
        PivotFilterOperator::GreaterThanOrEqual => (21, 6),
        PivotFilterOperator::LessThan => (22, 1),
        PivotFilterOperator::LessThanOrEqual => (23, 3),
        _ => unreachable!("unsupported XLSB value filter operator"),
    }
}

fn xlsb_date_filter_type_and_operator(operator: PivotFilterOperator) -> Option<(u32, u8)> {
    Some(match operator {
        PivotFilterOperator::Equals => (26, 2),
        PivotFilterOperator::NotEquals => (62, 5),
        PivotFilterOperator::LessThan => (27, 1),
        PivotFilterOperator::LessThanOrEqual => (63, 3),
        PivotFilterOperator::GreaterThan => (28, 4),
        PivotFilterOperator::GreaterThanOrEqual => (64, 6),
        PivotFilterOperator::BeginsWith
        | PivotFilterOperator::DoesNotBeginWith
        | PivotFilterOperator::EndsWith
        | PivotFilterOperator::DoesNotEndWith
        | PivotFilterOperator::Contains
        | PivotFilterOperator::DoesNotContain => return None,
    })
}

fn xlsb_date_period_filter_codes(period: PivotDatePeriod) -> Option<(u32, u32)> {
    let cft = match period {
        PivotDatePeriod::Tomorrow => 0x08,
        PivotDatePeriod::Today => 0x09,
        PivotDatePeriod::Yesterday => 0x0A,
        PivotDatePeriod::NextWeek => 0x0B,
        PivotDatePeriod::ThisWeek => 0x0C,
        PivotDatePeriod::LastWeek => 0x0D,
        PivotDatePeriod::NextMonth => 0x0E,
        PivotDatePeriod::ThisMonth => 0x0F,
        PivotDatePeriod::LastMonth => 0x10,
        PivotDatePeriod::NextQuarter => 0x11,
        PivotDatePeriod::ThisQuarter => 0x12,
        PivotDatePeriod::LastQuarter => 0x13,
        PivotDatePeriod::NextYear => 0x14,
        PivotDatePeriod::ThisYear => 0x15,
        PivotDatePeriod::LastYear => 0x16,
        PivotDatePeriod::YearToDate => 0x17,
        PivotDatePeriod::Quarter(1) => 0x18,
        PivotDatePeriod::Quarter(2) => 0x19,
        PivotDatePeriod::Quarter(3) => 0x1A,
        PivotDatePeriod::Quarter(4) => 0x1B,
        PivotDatePeriod::Month(1) => 0x1C,
        PivotDatePeriod::Month(2) => 0x1D,
        PivotDatePeriod::Month(3) => 0x1E,
        PivotDatePeriod::Month(4) => 0x1F,
        PivotDatePeriod::Month(5) => 0x20,
        PivotDatePeriod::Month(6) => 0x21,
        PivotDatePeriod::Month(7) => 0x22,
        PivotDatePeriod::Month(8) => 0x23,
        PivotDatePeriod::Month(9) => 0x24,
        PivotDatePeriod::Month(10) => 0x25,
        PivotDatePeriod::Month(11) => 0x26,
        PivotDatePeriod::Month(12) => 0x27,
        PivotDatePeriod::Month(_) | PivotDatePeriod::Quarter(_) => return None,
    };
    Some((cft + 22, cft))
}

fn pivot_date_period_filter_bounds(
    period: PivotDatePeriod,
    date_system: DateSystem,
) -> Option<(f64, f64)> {
    let today = chrono::Local::now().date_naive();
    let year = today.year();
    let month = today.month();
    let day = today.day();
    match period {
        PivotDatePeriod::Tomorrow => {
            let date = today.checked_add_signed(chrono::Duration::days(1))?;
            Some(exclusive_day_range(
                date.year(),
                date.month(),
                date.day(),
                date_system,
            ))
        }
        PivotDatePeriod::Today => Some(exclusive_day_range(year, month, day, date_system)),
        PivotDatePeriod::Yesterday => {
            let date = today.checked_sub_signed(chrono::Duration::days(1))?;
            Some(exclusive_day_range(
                date.year(),
                date.month(),
                date.day(),
                date_system,
            ))
        }
        PivotDatePeriod::NextWeek => {
            let date = today.checked_add_signed(chrono::Duration::days(7))?;
            week_filter_bounds(date.year(), date.month(), date.day(), date_system)
        }
        PivotDatePeriod::ThisWeek => week_filter_bounds(year, month, day, date_system),
        PivotDatePeriod::LastWeek => {
            let date = today.checked_sub_signed(chrono::Duration::days(7))?;
            week_filter_bounds(date.year(), date.month(), date.day(), date_system)
        }
        PivotDatePeriod::NextMonth => {
            let (year, month) = shift_month(year, month, 1)?;
            Some(exclusive_month_range(year, month, date_system))
        }
        PivotDatePeriod::ThisMonth => Some(exclusive_month_range(year, month, date_system)),
        PivotDatePeriod::LastMonth => {
            let (year, month) = shift_month(year, month, -1)?;
            Some(exclusive_month_range(year, month, date_system))
        }
        PivotDatePeriod::NextQuarter => {
            let (start_year, start_month) = quarter_start_for_shift(year, month, 1)?;
            Some(exclusive_month_span(
                start_year,
                start_month,
                3,
                date_system,
            )?)
        }
        PivotDatePeriod::ThisQuarter => {
            let start_month = ((month - 1) / 3) * 3 + 1;
            Some(exclusive_month_span(year, start_month, 3, date_system)?)
        }
        PivotDatePeriod::LastQuarter => {
            let (start_year, start_month) = quarter_start_for_shift(year, month, -1)?;
            Some(exclusive_month_span(
                start_year,
                start_month,
                3,
                date_system,
            )?)
        }
        PivotDatePeriod::NextYear => Some(exclusive_year_range(year + 1, date_system)),
        PivotDatePeriod::ThisYear => Some(exclusive_year_range(year, date_system)),
        PivotDatePeriod::LastYear => Some(exclusive_year_range(year - 1, date_system)),
        PivotDatePeriod::YearToDate => Some((
            date_to_serial(year, 1, 1, date_system),
            date_to_serial(year, month, day, date_system) + 1.0,
        )),
        PivotDatePeriod::Month(_) | PivotDatePeriod::Quarter(_) => None,
    }
}

fn exclusive_day_range(year: i32, month: u32, day: u32, date_system: DateSystem) -> (f64, f64) {
    let start = date_to_serial(year, month, day, date_system);
    (start, start + 1.0)
}

fn week_filter_bounds(
    year: i32,
    month: u32,
    day: u32,
    date_system: DateSystem,
) -> Option<(f64, f64)> {
    let date = chrono::NaiveDate::from_ymd_opt(year, month, day)?;
    let start = date.checked_sub_signed(chrono::Duration::days(
        date.weekday().num_days_from_monday() as i64,
    ))?;
    let end = start.checked_add_signed(chrono::Duration::days(7))?;
    Some((
        date_to_serial(start.year(), start.month(), start.day(), date_system),
        date_to_serial(end.year(), end.month(), end.day(), date_system),
    ))
}

fn exclusive_month_range(year: i32, month: u32, date_system: DateSystem) -> (f64, f64) {
    let (end_year, end_month) = shift_month(year, month, 1).unwrap_or((year + 1, 1));
    (
        date_to_serial(year, month, 1, date_system),
        date_to_serial(end_year, end_month, 1, date_system),
    )
}

fn exclusive_month_span(
    year: i32,
    month: u32,
    months: i32,
    date_system: DateSystem,
) -> Option<(f64, f64)> {
    let (end_year, end_month) = shift_month(year, month, months)?;
    Some((
        date_to_serial(year, month, 1, date_system),
        date_to_serial(end_year, end_month, 1, date_system),
    ))
}

fn exclusive_year_range(year: i32, date_system: DateSystem) -> (f64, f64) {
    (
        date_to_serial(year, 1, 1, date_system),
        date_to_serial(year + 1, 1, 1, date_system),
    )
}

fn shift_month(year: i32, month: u32, delta: i32) -> Option<(i32, u32)> {
    if !(1..=12).contains(&month) {
        return None;
    }
    let zero_based = year.checked_mul(12)? + month as i32 - 1 + delta;
    let shifted_year = zero_based.div_euclid(12);
    let shifted_month = zero_based.rem_euclid(12) as u32 + 1;
    Some((shifted_year, shifted_month))
}

fn quarter_start_for_shift(year: i32, month: u32, delta: i32) -> Option<(i32, u32)> {
    let start_month = ((month - 1) / 3) * 3 + 1;
    shift_month(year, start_month, delta * 3)
}

fn write_xlsb_numeric_custom_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    operator_code: u8,
    value: f64,
) -> std::io::Result<()> {
    let mut payload = Vec::with_capacity(10);
    payload.push(4u8);
    payload.push(operator_code);
    payload.extend_from_slice(&value.to_le_bytes());
    rw.write_record(records::BRT_CUSTOM_FILTER, &payload)
}

fn write_xlsb_string_custom_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    operator_code: u8,
    value: &str,
) -> std::io::Result<()> {
    let mut payload = Vec::with_capacity(10 + value.len() * 2);
    payload.push(6u8);
    payload.push(operator_code);
    payload.push(1u8);
    payload.push(0u8);
    payload.extend_from_slice(&[0u8; 6]);
    payload.extend_from_slice(&encode_wide_str(value));
    rw.write_record(records::BRT_CUSTOM_FILTER, &payload)
}

struct XlsbLabelCustomFilter {
    operator_code: u8,
    discriminator: u8,
    value: String,
}

fn label_filter_custom_filter(
    operator: PivotFilterOperator,
    value: &str,
) -> Option<XlsbLabelCustomFilter> {
    let (operator_code, discriminator, value) = match operator {
        PivotFilterOperator::BeginsWith => (2, 0, format!("{value}*")),
        PivotFilterOperator::EndsWith => (2, 0, format!("*{value}")),
        PivotFilterOperator::Contains => (2, 0, format!("*{value}*")),
        PivotFilterOperator::NotEquals => (5, 1, value.to_string()),
        PivotFilterOperator::DoesNotBeginWith => (5, 0, format!("{value}*")),
        PivotFilterOperator::DoesNotEndWith => (5, 0, format!("*{value}")),
        PivotFilterOperator::DoesNotContain => (5, 0, format!("*{value}*")),
        PivotFilterOperator::LessThan => (1, 1, value.to_string()),
        PivotFilterOperator::LessThanOrEqual => (3, 1, value.to_string()),
        PivotFilterOperator::GreaterThan => (4, 1, value.to_string()),
        PivotFilterOperator::GreaterThanOrEqual => (6, 1, value.to_string()),
        _ => return None,
    };
    Some(XlsbLabelCustomFilter {
        operator_code,
        discriminator,
        value,
    })
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
    axis_tuples: &XlsbPivotAxisTuples,
) -> CellRange {
    let page_field_count = effective_page_field_count(cache, layout, pivot);
    let (page_rows, _) = page_field_area_size(pivot, page_field_count);
    let body_start_row = pivot.target.row
        + if page_rows == 0 {
            0
        } else {
            page_rows.saturating_add(1)
        };
    let row_item_count = axis_item_count(&pivot.rows, &axis_tuples.rows).max(1);
    let effective_columns = effective_column_fields(pivot, cache);
    let row_header_count = effective_columns.len() as u32 + 1;
    let row_count = row_header_count + row_item_count as u32 + 1;

    let col_item_count = axis_item_count(&effective_columns, &axis_tuples.columns).max(1);
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

fn page_field_area_size(pivot: &PivotTable, count: usize) -> (u32, u32) {
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

fn axis_item_count(fields: &[PivotField], tuples: &[Vec<u32>]) -> usize {
    if fields.is_empty() {
        return 1;
    }
    tuples.len()
}

fn build_xlsb_axis_tuples(
    part: &FormatPivotTable,
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    grouping_infos: &[XlsbPivotGroupingInfo<'_>],
    effective_columns: &[PivotField],
) -> XlsbResult<XlsbPivotAxisTuples> {
    let rows = if pivot.rows.is_empty() {
        Vec::new()
    } else if let Some(tuples) = &part.axis_tuples.rows {
        tuples.clone()
    } else {
        axis_item_tuples(part, cache, layout, &pivot.rows, grouping_infos)?
    };
    let columns = if effective_columns.is_empty() {
        Vec::new()
    } else if let Some(tuples) = &part.axis_tuples.columns {
        tuples.clone()
    } else {
        axis_item_tuples(part, cache, layout, effective_columns, grouping_infos)?
    };
    Ok(XlsbPivotAxisTuples { rows, columns })
}

fn effective_column_fields(pivot: &PivotTable, cache: &FormatPivotCache) -> Vec<PivotField> {
    let mut fields = pivot.columns.clone();
    if let Some(source_index) = synthetic_consolidation_column_source_index(pivot, cache) {
        let mut field = PivotField::new(cache.fields[source_index].name.clone());
        field.sort = PivotSort::None;
        fields.push(field);
    }
    fields
}

fn axis_item_tuples(
    part: &FormatPivotTable,
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
    for row in xlsb_visible_row_indexes(part, cache.row_count) {
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

fn xlsb_visible_row_indexes(part: &FormatPivotTable, row_count: usize) -> XlsbVisibleRowIter<'_> {
    match part.visible_rows.as_deref() {
        Some(rows) => XlsbVisibleRowIter::Filtered(rows.iter().copied()),
        None => XlsbVisibleRowIter::All(0..row_count),
    }
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

fn expanded_page_field_count(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    fields: &[PivotField],
) -> usize {
    fields
        .iter()
        .map(|field| {
            cache
                .field_index(&field.field.name)
                .and_then(|source_index| {
                    layout
                        .date_derived_by_base
                        .get(&source_index)
                        .map(|indexes| indexes.len())
                        .or_else(|| {
                            layout
                                .manual_derived_by_base
                                .contains_key(&source_index)
                                .then_some(1)
                        })
                })
                .unwrap_or(1)
        })
        .sum()
}

fn effective_page_field_count(
    cache: &FormatPivotCache,
    layout: &XlsbPivotCacheLayout,
    pivot: &PivotTable,
) -> usize {
    expanded_page_field_count(cache, layout, &pivot.page_fields)
        + synthetic_consolidation_page_source_indexes(pivot, cache).len()
}

fn synthetic_consolidation_page_source_indexes(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
) -> Vec<usize> {
    consolidation_page_source_indexes(cache)
        .into_iter()
        .filter(|index| !pivot_axis_contains_source_index(&pivot.page_fields, cache, *index))
        .collect()
}

fn synthetic_consolidation_column_source_index(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
) -> Option<usize> {
    if !matches!(cache.source, FormatPivotSource::Consolidation { .. }) {
        return None;
    }
    let source_index = cache.field_index("Column")?;
    if pivot_axis_contains_source_index(&pivot.rows, cache, source_index)
        || pivot_axis_contains_source_index(&pivot.columns, cache, source_index)
        || pivot_axis_contains_source_index(&pivot.page_fields, cache, source_index)
    {
        return None;
    }
    Some(source_index)
}

fn consolidation_page_source_indexes(cache: &FormatPivotCache) -> Vec<usize> {
    let FormatPivotSource::Consolidation { ranges } = &cache.source else {
        return Vec::new();
    };
    let page_count = ranges
        .iter()
        .map(|range| range.page_items.len())
        .max()
        .unwrap_or(0);
    (0..page_count)
        .filter_map(|index| {
            let name = format!("Page{}", index + 1);
            cache.field_index(&name)
        })
        .collect()
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
    item_kinds: Vec<PivotCacheItemKind>,
}

fn cache_field_usage(
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsbResult<CacheFieldUsage> {
    let mut axis_names = HashSet::new();
    let mut date_filter_names = HashSet::new();
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
            insert_cache_axis_name(&mut axis_names, cache, &field.field.name);
        }
        for filter in &pivot.filters {
            if let Some(name) = filter_field_name(filter) {
                insert_cache_axis_name(&mut axis_names, cache, name);
            }
            match filter {
                PivotFilter::Date { field, .. }
                | PivotFilter::DateBetween { field, .. }
                | PivotFilter::DatePeriod { field, .. } => {
                    insert_cache_axis_name(&mut date_filter_names, cache, &field.name);
                }
                _ => {}
            }
        }
        for grouping in &pivot.groupings {
            insert_cache_axis_name(&mut axis_names, cache, grouping_field_name(grouping));
        }
    }

    let date_system = workbook_date_system(workbook.settings().date_1904);
    let mut store_items = Vec::with_capacity(cache.fields.len());
    let mut item_kinds = Vec::with_capacity(cache.fields.len());
    for field in &cache.fields {
        let field_key = field.name.to_lowercase();
        let item_kind = if date_filter_names.contains(&field_key) {
            for value in &field.shared_items {
                let PivotValue::Number(serial) = value else {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot date filter for field {} requires numeric date source values",
                        field.name
                    )));
                };
                if !valid_pcdi_datetime(*serial, date_system) {
                    return Err(XlsbError::InvalidFormat(format!(
                        "XLSB pivot date filter for field {} has invalid date serial: {serial}",
                        field.name
                    )));
                }
            }
            PivotCacheItemKind::DateTime
        } else {
            PivotCacheItemKind::Normal
        };
        let used_on_axis = axis_names.contains(&field_key)
            || matches!(cache.source, FormatPivotSource::Consolidation { .. });
        store_items.push(
            used_on_axis
                || field
                    .shared_items
                    .iter()
                    .any(|value| !matches!(value, PivotValue::Number(_))),
        );
        item_kinds.push(item_kind);
    }
    Ok(CacheFieldUsage {
        store_items,
        item_kinds,
    })
}

fn insert_cache_axis_name(names: &mut HashSet<String>, cache: &FormatPivotCache, field_name: &str) {
    if let Some(index) = cache.field_index(field_name) {
        names.insert(cache.fields[index].name.to_lowercase());
    } else {
        names.insert(field_name.to_lowercase());
    }
}

fn selected_page_item_index(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    field_name: &str,
    source_index: Option<usize>,
    shared_items: &[PivotValue],
) -> XlsbResult<Option<u32>> {
    let Some(filter) = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems {
                field: filter_field,
                ..
            } if pivot_filter_field_matches(cache, &filter_field.name, field_name, source_index)
        )
    }) else {
        return Ok(None);
    };
    let PivotFilter::FieldItems { allowed_items, .. } = filter else {
        return Ok(None);
    };

    if allowed_items.is_empty() {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot page field {field_name} requires at least one selected item"
        )));
    }

    if allowed_items.len() > 1 {
        field_filter_hidden_item_indexes_for_optional_source_index(
            pivot,
            cache,
            field_name,
            source_index,
            shared_items,
        )?;
        return Ok(Some(0x0010_00FE));
    }

    let item = &allowed_items[0];
    let Some(index) = shared_items.iter().position(|candidate| candidate == item) else {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot page field {field_name} selected item is not present in the cache"
        )));
    };
    Ok(Some(index as u32))
}

fn field_filter_hidden_item_indexes_for_source_index(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    source_index: usize,
    shared_items: &[PivotValue],
) -> XlsbResult<HashSet<u32>> {
    let field_name = cache
        .fields
        .get(source_index)
        .map(|field| field.name.as_str())
        .unwrap_or("");
    field_filter_hidden_item_indexes_for_optional_source_index(
        pivot,
        cache,
        field_name,
        Some(source_index),
        shared_items,
    )
}

fn field_filter_hidden_item_indexes_for_optional_source_index(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    field_name: &str,
    source_index: Option<usize>,
    shared_items: &[PivotValue],
) -> XlsbResult<HashSet<u32>> {
    field_filter_hidden_item_indexes_for_page_source_index(
        pivot,
        cache,
        field_name,
        source_index,
        shared_items,
        field_name,
        source_index,
    )
}

fn field_filter_hidden_item_indexes_for_page_field(
    pivot: &PivotTable,
    field_name: &str,
    shared_items: &[PivotValue],
    page_field_name: &str,
) -> XlsbResult<HashSet<u32>> {
    let Some(PivotFilter::FieldItems { allowed_items, .. }) = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems {
                field: filter_field,
                ..
            } if filter_field.name.eq_ignore_ascii_case(field_name)
        )
    }) else {
        return Ok(HashSet::new());
    };

    if pivot_axis_contains_field(&pivot.page_fields, page_field_name) && allowed_items.len() <= 1 {
        return Ok(HashSet::new());
    }

    let mut allowed_indexes = HashSet::with_capacity(allowed_items.len());
    for item in allowed_items {
        let Some(index) = shared_items.iter().position(|candidate| candidate == item) else {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot field {field_name} selected item is not present in the cache"
            )));
        };
        allowed_indexes.insert(index as u32);
    }

    Ok((0..shared_items.len() as u32)
        .filter(|index| !allowed_indexes.contains(index))
        .collect())
}

fn field_filter_hidden_item_indexes_for_page_source_index(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    field_name: &str,
    source_index: Option<usize>,
    shared_items: &[PivotValue],
    page_field_name: &str,
    page_source_index: Option<usize>,
) -> XlsbResult<HashSet<u32>> {
    let Some(PivotFilter::FieldItems { allowed_items, .. }) = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems {
                field: filter_field,
                ..
            } if pivot_filter_field_matches(cache, &filter_field.name, field_name, source_index)
        )
    }) else {
        return Ok(HashSet::new());
    };

    let page_axis_contains = page_source_index.map_or_else(
        || pivot_axis_contains_field(&pivot.page_fields, page_field_name),
        |index| pivot_axis_contains_source_index(&pivot.page_fields, cache, index),
    );
    if page_axis_contains && allowed_items.len() <= 1 {
        return Ok(HashSet::new());
    }

    let mut allowed_indexes = HashSet::with_capacity(allowed_items.len());
    for item in allowed_items {
        let Some(index) = shared_items.iter().position(|candidate| candidate == item) else {
            return Err(XlsbError::InvalidFormat(format!(
                "XLSB pivot field {field_name} selected item is not present in the cache"
            )));
        };
        allowed_indexes.insert(index as u32);
    }

    Ok((0..shared_items.len() as u32)
        .filter(|index| !allowed_indexes.contains(index))
        .collect())
}

fn pivot_filter_field_matches(
    cache: &FormatPivotCache,
    candidate_name: &str,
    field_name: &str,
    source_index: Option<usize>,
) -> bool {
    candidate_name.eq_ignore_ascii_case(field_name)
        || source_index.is_some_and(|index| cache.field_index(candidate_name) == Some(index))
}

fn pivot_field_axis_for_source_index(
    pivot: &PivotTable,
    cache: &FormatPivotCache,
    source_index: usize,
) -> u8 {
    const SX_AXIS_ROW: u8 = 0x01;
    const SX_AXIS_COL: u8 = 0x02;
    const SX_AXIS_PAGE: u8 = 0x04;
    const SX_AXIS_DATA: u8 = 0x08;

    let mut axis = if pivot_axis_contains_source_index(&pivot.rows, cache, source_index) {
        SX_AXIS_ROW
    } else if pivot_axis_contains_source_index(&pivot.columns, cache, source_index) {
        SX_AXIS_COL
    } else if synthetic_consolidation_column_source_index(pivot, cache) == Some(source_index) {
        SX_AXIS_COL
    } else if pivot_axis_contains_source_index(&pivot.page_fields, cache, source_index) {
        SX_AXIS_PAGE
    } else if synthetic_consolidation_page_source_indexes(pivot, cache).contains(&source_index) {
        SX_AXIS_PAGE
    } else {
        0
    };

    if pivot
        .measures
        .iter()
        .any(|measure| cache.field_index(&measure.field.name) == Some(source_index))
    {
        axis |= SX_AXIS_DATA;
    }

    axis
}

fn pivot_axis_contains_source_index(
    fields: &[PivotField],
    cache: &FormatPivotCache,
    source_index: usize,
) -> bool {
    fields
        .iter()
        .any(|field| cache.field_index(&field.field.name) == Some(source_index))
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

fn aggregate_code(aggregate: PivotAggregate) -> u32 {
    match aggregate {
        PivotAggregate::Sum => 0x00,
        PivotAggregate::Count => 0x01,
        PivotAggregate::Average => 0x02,
        PivotAggregate::Max => 0x03,
        PivotAggregate::Min => 0x04,
        PivotAggregate::Product => 0x05,
        PivotAggregate::CountNumbers => 0x06,
        PivotAggregate::StdDev => 0x07,
        PivotAggregate::StdDevP => 0x08,
        PivotAggregate::Var => 0x09,
        PivotAggregate::VarP => 0x0A,
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

fn xml_attr_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
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
