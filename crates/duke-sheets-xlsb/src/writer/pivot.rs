use std::collections::HashSet;
use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, records, RecordWriter};
use crate::error::{XlsbError, XlsbResult};
use duke_sheets_core::{
    CellError, CellRange, PivotAggregate, PivotField, PivotFilter, PivotTable, PivotValue, Workbook,
};
use duke_sheets_pivot::{
    FormatPivotCache, FormatPivotCacheField, FormatPivotPlan, FormatPivotSource, FormatPivotTable,
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
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    write_ac_block(&mut rw, 10, true)?;
    write_begin_pivot_cache_def(&mut rw, cache)?;
    write_pivot_cache_source(&mut rw, cache)?;

    rw.write_record(
        records::BRT_BEGIN_PCD_FIELDS,
        &(cache.fields.len() as u32).to_le_bytes(),
    )?;
    for (field, store_items) in cache.fields.iter().zip(&usage.store_items) {
        write_pcd_field(&mut rw, field)?;
        write_pcd_shared_items(&mut rw, field, *store_items)?;
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
                let item_id = field.item_ids.get(row).copied().unwrap_or(0);
                let value = field
                    .shared_items
                    .get(item_id as usize)
                    .unwrap_or(&PivotValue::Blank);
                if usage.store_items[field_index] {
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

    let path = format!("xl/pivotTables/pivotTable{}.bin", part.table_num);
    zip.start_file(path, *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    write_ac_block(&mut rw, 7, false)?;
    write_begin_sx_view(&mut rw, pivot)?;
    write_sx_location(&mut rw, pivot, cache)?;
    rw.write_record(records::BRT_END_SX_LOCATION, &[])?;
    write_sx_fields(&mut rw, pivot, cache, &usage)?;
    write_axis_fields(
        &mut rw,
        records::BRT_BEGIN_ISXVD_RWS,
        records::BRT_END_ISXVD_RWS,
        cache,
        &pivot.rows,
    )?;
    write_axis_items(
        &mut rw,
        records::BRT_BEGIN_SX_ROW_ITEMS,
        records::BRT_END_SX_ROW_ITEMS,
        cache,
        &pivot.rows,
    )?;
    write_axis_fields(
        &mut rw,
        records::BRT_BEGIN_ISXVD_COLS,
        records::BRT_END_ISXVD_COLS,
        cache,
        &pivot.columns,
    )?;
    write_axis_items(
        &mut rw,
        records::BRT_BEGIN_SX_COL_ITEMS,
        records::BRT_END_SX_COL_ITEMS,
        cache,
        &pivot.columns,
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
) -> std::io::Result<()> {
    let mut payload = Vec::new();
    let flags: u16 = if field.database_field { 0x0004 } else { 0x0000 };
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(&field.name));
    rw.write_record(records::BRT_BEGIN_PCD_FIELD, &payload)
}

fn write_pcd_shared_items<W: Write>(
    rw: &mut RecordWriter<W>,
    field: &FormatPivotCacheField,
    store_items: bool,
) -> std::io::Result<()> {
    let stats = SharedItemStats::from_values(&field.shared_items);
    let mut payload = Vec::new();
    payload.extend_from_slice(&stats.flags().to_le_bytes());
    let stored_count = if store_items {
        field.shared_items.len() as u32
    } else {
        0
    };
    payload.extend_from_slice(&stored_count.to_le_bytes());
    if stats.has_number {
        payload.extend_from_slice(&stats.min_number.unwrap_or(0.0).to_le_bytes());
        payload.extend_from_slice(&stats.max_number.unwrap_or(0.0).to_le_bytes());
    }
    rw.write_record(records::BRT_BEGIN_PCD_SHARED_ITEMS, &payload)?;
    if store_items {
        for item in &field.shared_items {
            write_pcdi_item(rw, item)?;
        }
    }
    rw.write_record(records::BRT_END_PCD_SHARED_ITEMS, &[])
}

fn write_pcdi_item<W: Write>(rw: &mut RecordWriter<W>, value: &PivotValue) -> std::io::Result<()> {
    match value {
        PivotValue::Blank => rw.write_record(records::BRT_PCDI_MISSING, &[]),
        PivotValue::Boolean(value) => {
            rw.write_record(records::BRT_PCDI_BOOLEAN, &[if *value { 1 } else { 0 }])
        }
        PivotValue::Number(value) => {
            rw.write_record(records::BRT_PCDI_NUMBER, &value.to_le_bytes())
        }
        PivotValue::String(value) => {
            rw.write_record(records::BRT_PCDI_STRING, &encode_wide_str(value))
        }
        PivotValue::Error(value) => rw.write_record(records::BRT_PCDI_ERROR, &[error_code(*value)]),
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
    payload.extend_from_slice(&[
        0x00, 0x41, 0x40, 0x01, 0xF0, 0x64, 0x09, 0x00, 0xD9, 0x00, 0x00, 0x00, 0x02, 0x00, 0x08,
        0x03, 0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00,
    ]);
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
) -> std::io::Result<()> {
    let range = pivot
        .rendered_range
        .unwrap_or_else(|| estimated_pivot_range(pivot, cache));
    let first_data_col = pivot.rows.len().max(1) as u32;
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
) -> XlsbResult<()> {
    rw.write_record(
        records::BRT_BEGIN_SXVDS,
        &(cache.fields.len() as u32).to_le_bytes(),
    )?;
    for (index, field) in cache.fields.iter().enumerate() {
        write_sx_field(rw, pivot_field_axis(pivot, &field.name))?;
        if usage.store_items[index] {
            write_sx_field_items(rw, field)?;
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
    field: &FormatPivotCacheField,
) -> std::io::Result<()> {
    let count = field.shared_items.len() as u32 + 1;
    rw.write_record(records::BRT_BEGIN_SXVIS, &count.to_le_bytes())?;
    for item_index in 0..field.shared_items.len() as u32 {
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
    fields: &[PivotField],
) -> XlsbResult<()> {
    if fields.is_empty() {
        return Ok(());
    }

    let mut payload = Vec::new();
    payload.extend_from_slice(&(fields.len() as u32).to_le_bytes());
    for field in fields {
        let index = cache.field_index(&field.field.name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "pivot references unknown field {}",
                field.field.name
            ))
        })?;
        payload.extend_from_slice(&(index as u32).to_le_bytes());
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
    fields: &[PivotField],
) -> XlsbResult<()> {
    let tuples = axis_item_tuples(cache, fields)?;
    let count = if fields.is_empty() {
        1
    } else {
        tuples.len() + 1
    };
    rw.write_record(begin_record, &(count as u32).to_le_bytes())?;
    if fields.is_empty() {
        write_sxli(rw, &[], false)?;
    } else {
        for tuple in &tuples {
            write_sxli(rw, tuple, false)?;
        }
        write_sxli(rw, &[0], true)?;
    }
    rw.write_record(end_record, &[])?;
    Ok(())
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
) -> std::io::Result<()> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&(if grand_total { 13u16 } else { 0u16 }).to_le_bytes());
    payload.extend_from_slice(&(item_indexes.len() as u32).to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
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
    payload.extend_from_slice(&0x0030u16.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(style_name));
    rw.write_record(records::BRT_SX_VIEW_STYLE, &payload)
}

fn estimated_pivot_range(pivot: &PivotTable, cache: &FormatPivotCache) -> CellRange {
    let (page_rows, _) = page_field_area_size(pivot);
    let body_start_row = pivot.target.row
        + if page_rows == 0 {
            0
        } else {
            page_rows.saturating_add(1)
        };
    let row_item_count = axis_item_count(cache, &pivot.rows).max(1);
    let row_header_count = pivot.columns.len() as u32 + 1;
    let row_count = row_header_count + row_item_count as u32 + 1;

    let col_item_count = axis_item_count(cache, &pivot.columns).max(1);
    let measure_count = pivot.measures.len().max(1);
    let value_col_count = col_item_count * measure_count;
    let col_count = pivot.rows.len().max(1) as u16 + value_col_count as u16;
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

fn axis_item_count(cache: &FormatPivotCache, fields: &[PivotField]) -> usize {
    if fields.is_empty() {
        return 1;
    }
    axis_item_tuples(cache, fields)
        .map(|tuples| tuples.len())
        .unwrap_or(1)
}

fn axis_item_tuples(cache: &FormatPivotCache, fields: &[PivotField]) -> XlsbResult<Vec<Vec<u32>>> {
    if fields.is_empty() {
        return Ok(Vec::new());
    }
    let indexes = fields
        .iter()
        .map(|field| {
            cache.field_index(&field.field.name).ok_or_else(|| {
                XlsbError::InvalidFormat(format!(
                    "pivot references unknown axis field {}",
                    field.field.name
                ))
            })
        })
        .collect::<XlsbResult<Vec<_>>>()?;

    let mut seen = HashSet::new();
    let mut tuples = Vec::new();
    for row in 0..cache.row_count {
        let tuple = indexes
            .iter()
            .map(|index| cache.fields[*index].item_ids.get(row).copied().unwrap_or(0))
            .collect::<Vec<_>>();
        if seen.insert(tuple.clone()) {
            tuples.push(tuple);
        }
    }
    Ok(tuples)
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

    fn flags(&self) -> u16 {
        let mut flags = 0x0400u16;
        if self.has_number {
            flags |= 0x0001 | 0x0002 | 0x0040 | 0x0100;
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
