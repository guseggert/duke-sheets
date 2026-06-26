use std::collections::HashSet;
use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, ptg, records, RecordWriter};
use crate::error::{XlsbError, XlsbResult};
use duke_sheets_core::{
    CellError, CellRange, PivotAggregate, PivotField, PivotFilter, PivotGrouping, PivotTable,
    PivotValue, PivotValuesAxis, Workbook,
};
use duke_sheets_formula::ast::{BinaryOperator, UnaryOperator};
use duke_sheets_formula::FormulaExpr;
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
    let groupings = groupings_for_cache(workbook, plan, cache)?;
    validate_xlsb_pivot_groupings(cache, groupings)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    write_ac_block(&mut rw, 10, true)?;
    write_begin_pivot_cache_def(&mut rw, cache)?;
    write_pivot_cache_source(&mut rw, cache)?;

    rw.write_record(
        records::BRT_BEGIN_PCD_FIELDS,
        &(cache.fields.len() as u32).to_le_bytes(),
    )?;
    for (field_index, (field, store_items)) in
        cache.fields.iter().zip(&usage.store_items).enumerate()
    {
        let pnames = write_pcd_field(&mut rw, field, cache)?;
        write_pcd_shared_items(&mut rw, field, *store_items)?;
        if let Some(grouping) = grouping_for_field(groupings, &field.name) {
            write_pcd_field_group(&mut rw, field_index, field, grouping)?;
        }
        write_pnames(&mut rw, &pnames)?;
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
    let values_on_rows = values_field_on_axis(pivot, PivotValuesAxis::Rows);
    let values_on_columns = values_field_on_axis(pivot, PivotValuesAxis::Columns);
    write_axis_fields(
        &mut rw,
        records::BRT_BEGIN_ISXVD_RWS,
        records::BRT_END_ISXVD_RWS,
        cache,
        &pivot.rows,
        values_on_rows,
        pivot.layout.values_axis_position,
    )?;
    write_axis_items(
        &mut rw,
        records::BRT_BEGIN_SX_ROW_ITEMS,
        records::BRT_END_SX_ROW_ITEMS,
        cache,
        &pivot.rows,
        values_on_rows,
        pivot.layout.values_axis_position,
        pivot.measures.len(),
    )?;
    write_axis_fields(
        &mut rw,
        records::BRT_BEGIN_ISXVD_COLS,
        records::BRT_END_ISXVD_COLS,
        cache,
        &pivot.columns,
        values_on_columns,
        pivot.layout.values_axis_position,
    )?;
    write_axis_items(
        &mut rw,
        records::BRT_BEGIN_SX_COL_ITEMS,
        records::BRT_END_SX_COL_ITEMS,
        cache,
        &pivot.columns,
        values_on_columns,
        pivot.layout.values_axis_position,
        pivot.measures.len(),
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
            }
            PivotGrouping::Date { .. } | PivotGrouping::Manual { .. } => {
                return Err(XlsbError::InvalidFormat(format!(
                    "XLSB pivot grouping currently supports numeric range grouping only: {field_name}"
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
    let mut payload = Vec::new();
    let mut flags: u16 = if field.database_field { 0x0004 } else { 0x0000 };
    if field.formula.is_some() {
        flags |= 0x0100;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(&field.name));
    let pnames = if let Some(formula) = field.formula.as_deref() {
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

fn write_pcd_field_group<W: Write>(
    rw: &mut RecordWriter<W>,
    field_index: usize,
    field: &FormatPivotCacheField,
    grouping: &PivotGrouping,
) -> XlsbResult<()> {
    let PivotGrouping::Number {
        start,
        end,
        interval,
        ..
    } = grouping
    else {
        return Err(XlsbError::InvalidFormat(format!(
            "XLSB pivot grouping currently supports numeric range grouping only: {}",
            grouping_field_name(grouping)
        )));
    };

    let field_index = checked_i32(field_index, "pivot grouped cache field index")?;
    let mut group_payload = Vec::with_capacity(8);
    group_payload.extend_from_slice(&(-1i32).to_le_bytes());
    group_payload.extend_from_slice(&field_index.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_PCDF_GROUP, &group_payload)?;

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

    rw.write_record(
        records::BRT_BEGIN_PCDFG_ITEMS,
        &(field.shared_items.len() as u32).to_le_bytes(),
    )?;
    for item in &field.shared_items {
        write_pcdi_item(rw, item)?;
    }
    rw.write_record(records::BRT_END_PCDFG_ITEMS, &[])?;
    rw.write_record(records::BRT_END_PCDF_GROUP, &[])?;
    Ok(())
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
    include_values_field: bool,
    values_position: Option<u32>,
) -> XlsbResult<()> {
    if fields.is_empty() && !include_values_field {
        return Ok(());
    }

    let mut indexes = Vec::new();
    for field in fields {
        let index = cache.field_index(&field.field.name).ok_or_else(|| {
            XlsbError::InvalidFormat(format!(
                "pivot references unknown field {}",
                field.field.name
            ))
        })?;
        indexes.push(index as i32);
    }
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
    fields: &[PivotField],
    include_values_field: bool,
    values_position: Option<u32>,
    measure_count: usize,
) -> XlsbResult<()> {
    let line_tuples = axis_line_tuples(
        cache,
        fields,
        include_values_field,
        values_position,
        measure_count,
    )?;
    let grand_tuples = axis_grand_total_tuples(
        fields.len(),
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
    fields: &[PivotField],
    include_values_field: bool,
    values_position: Option<u32>,
    measure_count: usize,
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

    let tuples = axis_item_tuples(cache, fields)?;
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
