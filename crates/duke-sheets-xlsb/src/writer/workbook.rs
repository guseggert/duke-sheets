use std::collections::HashMap;
use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;


use crate::biff12::compiler::{compile_formula, CompileContext};
use crate::biff12::{encode_wide_str, records, RecordWriter};
use crate::error::XlsbResult;
use duke_sheets_core::named_range::NameScope;
use duke_sheets_core::worksheet::SheetVisibility;
use duke_sheets_core::{CellAddress, Workbook};

pub(crate) fn write_workbook<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    workbook: &Workbook,
    has_formulas: bool,
    xlfn_names: &HashMap<String, u32>,
) -> XlsbResult<()> {
    zip.start_file("xl/workbook.bin", *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    rw.write_record(0x0083, &[])?; // BrtBeginBook

    write_file_version(&mut rw)?;

    let mut wb_prop_flags: u8 = 0x20;
    if workbook.settings().date_1904 {
        wb_prop_flags |= 0x01;
    }
    let mut wb_prop = Vec::new();
    wb_prop.extend_from_slice(&[wb_prop_flags, 0x00, 0x01, 0x00]);
    wb_prop.extend_from_slice(&0u32.to_le_bytes());
    wb_prop.extend_from_slice(&encode_wide_str(""));
    rw.write_record(records::BRT_WB_PROP, &wb_prop)?;

    write_book_views(&mut rw, workbook.active_sheet())?;

    rw.write_record(0x008F, &[])?; // BrtBeginBundleShs

    for i in 0..workbook.sheet_count() {
        let ws = workbook.worksheet(i).unwrap();
        let rel_id = format!("rId{}", i + 1);

        let mut payload = Vec::new();
        let visibility: u32 = match ws.visibility() {
            SheetVisibility::Visible => 0,
            SheetVisibility::Hidden => 1,
            SheetVisibility::VeryHidden => 2,
        };
        let sheet_id = (i + 1) as u32;
        payload.extend_from_slice(&visibility.to_le_bytes());
        payload.extend_from_slice(&sheet_id.to_le_bytes());
        payload.extend_from_slice(&encode_wide_str(&rel_id));
        payload.extend_from_slice(&encode_wide_str(ws.name()));
        rw.write_record(records::BRT_BUNDLE_SH, &payload)?;
    }

    rw.write_record(records::BRT_END_BUNDLE_SHS, &[])?;

    let has_user_names = !workbook.named_ranges().is_empty();
    let has_print_settings = (0..workbook.sheet_count()).any(|i| {
        let ws = workbook.worksheet(i).unwrap();
        let ps = ws.page_setup();
        ps.print_area.is_some() || ps.repeat_rows.is_some() || ps.repeat_cols.is_some()
    });

    if has_formulas || has_user_names || has_print_settings {
        write_extern_sheet(&mut rw, workbook.sheet_count())?;
    }

    if !xlfn_names.is_empty() {
        write_xlfn_name_records(&mut rw, xlfn_names)?;
    }

    if has_user_names {
        write_user_name_records(&mut rw, workbook, xlfn_names)?;
    }

    if has_print_settings {
        write_print_name_records(&mut rw, workbook, xlfn_names)?;
    }

    rw.write_record(0x0084, &[])?; // BrtEndBook

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}

fn write_file_version<W: Write>(rw: &mut RecordWriter<W>) -> std::io::Result<()> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str("xl"));
    payload.extend_from_slice(&encode_wide_str("7"));
    payload.extend_from_slice(&encode_wide_str("7"));
    payload.extend_from_slice(&encode_wide_str("29628"));
    rw.write_record(records::BRT_FILE_VERSION, &payload)
}

fn write_book_views<W: Write>(
    rw: &mut RecordWriter<W>,
    active_sheet: usize,
) -> std::io::Result<()> {
    rw.write_record(records::BRT_BEGIN_BOOK_VIEWS, &[])?;
    let mut bv = vec![0u8; 29];
    bv[0..4].copy_from_slice(&(-120i32).to_le_bytes());
    bv[4..8].copy_from_slice(&(-120i32).to_le_bytes());
    bv[8..12].copy_from_slice(&15600u32.to_le_bytes());
    bv[12..16].copy_from_slice(&11040u32.to_le_bytes());
    bv[16..20].copy_from_slice(&0x0258u32.to_le_bytes());
    bv[20..24].copy_from_slice(&(active_sheet as u32).to_le_bytes());
    bv[24..28].copy_from_slice(&(active_sheet as u32).to_le_bytes());
    rw.write_record(0x009E, &bv)?;
    rw.write_record(records::BRT_END_BOOK_VIEWS, &[])
}

fn write_xlfn_name_records<W: Write>(
    rw: &mut RecordWriter<W>,
    xlfn_names: &HashMap<String, u32>,
) -> std::io::Result<()> {
    let mut sorted: Vec<(&String, &u32)> = xlfn_names.iter().collect();
    sorted.sort_by_key(|(_, idx)| **idx);

    for (name, _) in sorted {
        let prefixed = format!("_xlfn.{name}");
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u32.to_le_bytes()); // flags
        payload.push(0u8); // chKey
        payload.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes()); // itab: workbook scope
        payload.extend_from_slice(&encode_wide_str(&prefixed));
        payload.extend_from_slice(&0u32.to_le_bytes()); // cce (rgce size = 0)
        payload.extend_from_slice(&0u32.to_le_bytes()); // cb (rgcb size = 0)
        for _ in 0..5 {
            payload.extend_from_slice(&encode_wide_str(""));
        }
        rw.write_record(records::BRT_NAME, &payload)?;
    }
    Ok(())
}

fn write_user_name_records<W: Write>(
    rw: &mut RecordWriter<W>,
    workbook: &Workbook,
    xlfn_names: &HashMap<String, u32>,
) -> std::io::Result<()> {
    let sheet_names: Vec<String> = (0..workbook.sheet_count())
        .map(|i| workbook.worksheet(i).unwrap().name().to_string())
        .collect();
    let compile_ctx = CompileContext {
        sheet_names,
        xlfn_names: xlfn_names.clone(),
    };

    for nr in workbook.named_ranges().iter() {
        let compiled = match compile_formula(&nr.refers_to, &compile_ctx) {
            Ok(c) => c,
            Err(e) => {
                log::warn!("skipping named range '{}': {e}", nr.name);
                continue;
            }
        };

        let mut payload = Vec::new();
        let mut flags = 0u32;
        if nr.hidden {
            flags |= 0x01;
        }
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.push(0u8);
        let itab: u32 = match nr.scope {
            NameScope::Workbook => 0xFFFFFFFF,
            NameScope::Sheet(idx) => idx as u32,
        };
        payload.extend_from_slice(&itab.to_le_bytes());
        payload.extend_from_slice(&encode_wide_str(&nr.name));

        payload.extend_from_slice(&(compiled.rgce.len() as u32).to_le_bytes());
        payload.extend_from_slice(&compiled.rgce);
        payload.extend_from_slice(&(compiled.rgcb.len() as u32).to_le_bytes());
        payload.extend_from_slice(&compiled.rgcb);

        for _ in 0..5 {
            payload.extend_from_slice(&encode_wide_str(""));
        }

        rw.write_record(records::BRT_NAME, &payload)?;
    }
    Ok(())
}

fn write_print_name_records<W: Write>(
    rw: &mut RecordWriter<W>,
    workbook: &Workbook,
    xlfn_names: &HashMap<String, u32>,
) -> std::io::Result<()> {
    let sheet_names: Vec<String> = (0..workbook.sheet_count())
        .map(|i| workbook.worksheet(i).unwrap().name().to_string())
        .collect();
    let compile_ctx = CompileContext {
        sheet_names,
        xlfn_names: xlfn_names.clone(),
    };

    for i in 0..workbook.sheet_count() {
        let ws = workbook.worksheet(i).unwrap();
        let ps = ws.page_setup();

        if let Some(ref range) = ps.print_area {
            let formula = format_print_area_formula(ws.name(), range);
            write_builtin_name_record(rw, 0x06, i as u32, &formula, &compile_ctx)?;
        }

        if ps.repeat_rows.is_some() || ps.repeat_cols.is_some() {
            if let Some(formula) = format_print_titles_formula(ws.name(), ps.repeat_rows, None) {
                write_builtin_name_record(rw, 0x07, i as u32, &formula, &compile_ctx)?;
            }
            if let Some(formula) = format_print_titles_formula(ws.name(), None, ps.repeat_cols) {
                write_builtin_name_record(rw, 0x07, i as u32, &formula, &compile_ctx)?;
            }
        }
    }
    Ok(())
}

fn builtin_name_string(index: u8) -> &'static str {
    match index {
        0x06 => "Print_Area",
        0x07 => "Print_Titles",
        _ => "_xlnm.Unknown",
    }
}

fn write_builtin_name_record<W: Write>(
    rw: &mut RecordWriter<W>,
    builtin_index: u8,
    sheet_idx: u32,
    formula: &str,
    compile_ctx: &CompileContext,
) -> std::io::Result<()> {
    let compiled = match compile_formula(formula, compile_ctx) {
        Ok(c) => c,
        Err(e) => {
            log::warn!(
                "print setting formula compile failed for '{}': {}",
                formula,
                e
            );
            return Ok(());
        }
    };

    let mut payload = Vec::new();
    let flags = 0x21u32; // hidden + builtin
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.push(0u8);
    payload.extend_from_slice(&sheet_idx.to_le_bytes());

    payload.extend_from_slice(&encode_wide_str(builtin_name_string(builtin_index)));

    payload.extend_from_slice(&(compiled.rgce.len() as u32).to_le_bytes());
    payload.extend_from_slice(&compiled.rgce);
    payload.extend_from_slice(&(compiled.rgcb.len() as u32).to_le_bytes());
    payload.extend_from_slice(&compiled.rgcb);

    for _ in 0..5 {
        payload.extend_from_slice(&encode_wide_str(""));
    }

    rw.write_record(records::BRT_NAME, &payload)
}

fn quote_sheet_name(name: &str) -> String {
    let needs_quoting = name.contains(' ')
        || name.contains('\'')
        || name.contains('!')
        || name.contains('[')
        || name.contains(']');
    if needs_quoting {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

fn format_print_area_formula(sheet_name: &str, range: &duke_sheets_core::CellRange) -> String {
    let quoted = quote_sheet_name(sheet_name);
    let start_col = CellAddress::column_to_letters(range.start.col);
    let end_col = CellAddress::column_to_letters(range.end.col);
    format!(
        "{}!${}${}:${}${}",
        quoted,
        start_col,
        range.start.row + 1,
        end_col,
        range.end.row + 1,
    )
}

fn format_print_titles_formula(
    sheet_name: &str,
    repeat_rows: Option<(u32, u32)>,
    repeat_cols: Option<(u16, u16)>,
) -> Option<String> {
    let quoted = quote_sheet_name(sheet_name);
    let mut parts = Vec::new();

    if let Some((r1, r2)) = repeat_rows {
        parts.push(format!("{}!$A${}:$XFD${}", quoted, r1 + 1, r2 + 1,));
    }

    if let Some((c1, c2)) = repeat_cols {
        let start = CellAddress::column_to_letters(c1);
        let end = CellAddress::column_to_letters(c2);
        parts.push(format!("{}!${}$1:${}$1048576", quoted, start, end,));
    }

    if parts.is_empty() {
        None
    } else {
        Some(parts.join(","))
    }
}

fn write_extern_sheet<W: Write>(
    rw: &mut RecordWriter<W>,
    sheet_count: usize,
) -> std::io::Result<()> {
    rw.write_record(0x0161, &[])?; // BrtBeginExternals
    rw.write_record(0x0165, &[])?; // BrtSupSelf

    let count = sheet_count as u32;
    let mut payload = Vec::with_capacity(4 + sheet_count * 12);
    payload.extend_from_slice(&count.to_le_bytes());
    for i in 0..sheet_count {
        let idx = i as u32;
        payload.extend_from_slice(&0u32.to_le_bytes()); // supBookIdx = 0 (self-ref)
        payload.extend_from_slice(&idx.to_le_bytes()); // firstSheet
        payload.extend_from_slice(&idx.to_le_bytes()); // lastSheet
    }
    rw.write_record(records::BRT_EXTERN_SHEET, &payload)?;

    rw.write_record(0x0162, &[])?; // BrtEndExternals
    Ok(())
}
