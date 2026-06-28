use std::collections::BTreeMap;
use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::compiler::{self, CompileContext};
use crate::biff12::{encode_nullable_wide_str, encode_wide_str, records, RecordWriter};
use crate::error::XlsbResult;
use duke_sheets_core::conditional_format::{
    CfOperator, CfRuleType, CfValueType, IconSetStyle, TimePeriod,
};
use duke_sheets_core::style::Color;
use duke_sheets_core::validation::{
    DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
};
use duke_sheets_core::worksheet::PageOrientation;
use duke_sheets_core::{CellAddress, CellError, CellValue, Worksheet};

use super::shared_strings::SstMap;
use super::styles::StyleMapping;

pub(crate) struct SheetWriteResult {
    pub sheet_rels: Vec<SheetRel>,
    pub has_comments: bool,
}

pub(crate) struct SheetRel {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

pub(crate) fn write_worksheet<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    index: usize,
    ws: &Worksheet,
    sst: &SstMap,
    style_mapping: &StyleMapping,
    compile_ctx: &CompileContext,
    has_drawing: bool,
    dxf_mapping: &super::styles::DxfMapping,
    table_count: usize,
) -> XlsbResult<SheetWriteResult> {
    let path = format!("xl/worksheets/sheet{}.bin", index + 1);
    zip.start_file(&path, *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    rw.write_record(records::BRT_BEGIN_SHEET, &[])?;

    let ws_prop = build_ws_prop(ws);
    rw.write_record(records::BRT_WS_PROP, &ws_prop)?;

    write_ws_dim(&mut rw, ws)?;
    write_sheet_views(&mut rw, ws)?;
    write_col_infos(&mut rw, ws)?;

    write_ws_fmt_info(&mut rw, ws)?;

    rw.write_record(records::BRT_BEGIN_SHEET_DATA, &[])?;

    let rows = collect_rows(ws, index, style_mapping);
    let custom_heights = ws.custom_row_heights();
    let hidden_rows = ws.hidden_rows();

    let mut all_rows: BTreeMap<u32, Option<&Vec<CellInfo>>> = BTreeMap::new();
    for &row in rows.keys() {
        all_rows.insert(row, Some(rows.get(&row).unwrap()));
    }
    for &row in custom_heights.keys() {
        all_rows.entry(row).or_insert(None);
    }
    for (&row, &hidden) in hidden_rows.iter() {
        if hidden {
            all_rows.entry(row).or_insert(None);
        }
    }
    for (&row, &level) in ws.row_outline_levels().iter() {
        if level > 0 {
            all_rows.entry(row).or_insert(None);
        }
    }

    for (&row, cells_opt) in &all_rows {
        let col_range = cells_opt.and_then(|cells| {
            if cells.is_empty() {
                return None;
            }
            let min_col = cells.iter().map(|c| c.col as u32).min().unwrap();
            let max_col = cells.iter().map(|c| c.col as u32).max().unwrap();
            Some((min_col, max_col))
        });
        write_row_hdr(&mut rw, row, ws, col_range)?;
        if let Some(cells) = cells_opt {
            for cell_info in *cells {
                write_cell(&mut rw, cell_info, sst, compile_ctx)?;
            }
        }
    }

    rw.write_record(records::BRT_END_SHEET_DATA, &[])?;

    let mut sheet_rels = Vec::new();
    let mut rid_counter = 1u32;

    write_merge_cells(&mut rw, ws)?;
    write_hyperlinks(&mut rw, ws, &mut sheet_rels, &mut rid_counter)?;
    write_auto_filter(&mut rw, ws, dxf_mapping.color_filter_dxf_id())?;
    write_sheet_protection(&mut rw, ws)?;
    write_page_setup(&mut rw, ws)?;
    write_print_options(&mut rw, ws)?;
    write_margins(&mut rw, ws)?;
    write_header_footer(&mut rw, ws)?;
    write_page_breaks(&mut rw, ws)?;
    write_data_validations(&mut rw, ws, compile_ctx)?;
    write_conditional_formats(&mut rw, ws, index, dxf_mapping, compile_ctx)?;

    let has_comments = ws.comments().next().is_some();

    if has_comments {
        let vml_rid_num = rid_counter + 1;
        let vml_rid = format!("rId{}", vml_rid_num);
        let encoded_vml_rid = crate::biff12::encode_wide_str(&vml_rid);
        rw.write_record(records::BRT_LEGACY_DRAWING, &encoded_vml_rid)?;
    }

    if has_drawing {
        let mut drawing_rid_num = rid_counter;
        if has_comments {
            drawing_rid_num += 2;
        }
        let rid = format!("rId{}", drawing_rid_num);
        let encoded_rid = crate::biff12::encode_wide_str(&rid);
        rw.write_record(records::BRT_DRAWING, &encoded_rid)?;
    }

    if table_count > 0 {
        let table_rid_start = rid_counter;
        let count_bytes = (table_count as u32).to_le_bytes();
        rw.write_record(records::BRT_BEGIN_LIST_PARTS, &count_bytes)?;
        for t in 0..table_count {
            let rid = format!("rId{}", table_rid_start + t as u32);
            rw.write_record(
                records::BRT_LIST_PART,
                &crate::biff12::encode_wide_str(&rid),
            )?;
        }
        rw.write_record(records::BRT_END_LIST_PARTS, &[])?;
    }

    rw.write_record(0x0082, &[])?;

    drop(rw);
    zip.write_all(&buf)?;

    Ok(SheetWriteResult {
        sheet_rels,
        has_comments,
    })
}

/// Build BrtWsProp payload.
///
/// Layout per [MS-XLSB] 2.4.875:
///   Bytes 0-2:  flags (17 bits A-Q) + reserved4 (6 bits), packed into 3 bytes
///   Bytes 3-10: brtcolorTab (BrtColor, 8 bytes)
///   Bytes 11-14: rwSync (u32, 0xFFFFFFFF = none)
///   Bytes 15-18: colSync (u32, 0xFFFFFFFF = none)
///   Bytes 19+:  strName (XLWideString, empty = 4 zero bytes)
fn build_ws_prop(ws: &Worksheet) -> Vec<u8> {
    let mut payload = vec![0u8; 23];
    payload[0] = 0xc9;
    payload[1] = 0x04;
    payload[2] = 0x02;

    let color_bytes = match ws.tab_color() {
        Some(color) => encode_brt_color_ws(&color),
        None => [0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00],
    };
    payload[3..11].copy_from_slice(&color_bytes);

    payload[11..15].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
    payload[15..19].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
    payload
}

/// Encode a Color as an 8-byte BrtColor for use in BrtWsProp.
///
/// Layout: xColorType(u8) + index(u8) + tint(i16) + R(u8) + G(u8) + B(u8) + A(u8)
fn encode_brt_color_ws(color: &duke_sheets_core::style::Color) -> [u8; 8] {
    use duke_sheets_core::style::Color;
    let mut buf = [0u8; 8];
    match color {
        Color::Auto => buf[0] = 0,
        Color::Indexed(idx) => {
            buf[0] = 1;
            buf[1] = *idx;
        }
        Color::Rgb { r, g, b } => {
            buf[0] = 2;
            buf[4] = *r;
            buf[5] = *g;
            buf[6] = *b;
            buf[7] = 0xFF;
        }
        Color::Argb { a, r, g, b } => {
            buf[0] = 2;
            buf[4] = *r;
            buf[5] = *g;
            buf[6] = *b;
            buf[7] = *a;
        }
        Color::Theme { index, tint } => {
            buf[0] = 3;
            buf[1] = *index;
            let tint_i16 = if *tint == 0 {
                0i16
            } else {
                ((*tint as f64 / 100.0) * 32767.0).round() as i16
            };
            buf[2..4].copy_from_slice(&tint_i16.to_le_bytes());
        }
    }
    buf
}

/// Emit BrtWsFmtInfo per [MS-XLSB] §2.4.873:
///   dxGCol u32, cchDefColWidth u16, miyDefRwHeight u16,
///   flags u16 (fUnsynced, fDyZero, fExAsc, fExDesc, reserved 12),
///   iOutLevelRw u8, iOutLevelCol u8.
///
/// miyDefRwHeight is in twips (1 point = 20 twips). cchDefColWidth is
/// in "characters of the maximum digit width of the Normal style font"
/// — Excel's column-width unit. We round our f64 column width to u16
/// (Excel's XLSB stores no fraction here even though XLSX does).
///
/// fUnsynced (bit 0) tells Excel "the default row height was set
/// explicitly". Without it Excel resets miyDefRwHeight to 300 (15pt)
/// on save.
fn write_ws_fmt_info<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let mut payload = Vec::with_capacity(12);
    payload.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes()); // dxGCol = auto
    let cch_def_col_width = ws.default_column_width().round().clamp(0.0, 255.0) as u16;
    payload.extend_from_slice(&cch_def_col_width.to_le_bytes());
    let row_height_pt = ws.default_row_height();
    let miy_def_rw_height = (row_height_pt * 20.0).round().clamp(0.0, u16::MAX as f64) as u16;
    payload.extend_from_slice(&miy_def_rw_height.to_le_bytes());
    let mut flags: u16 = 0;
    if (row_height_pt - 15.0).abs() > 0.001 {
        flags |= 0x0001; // fUnsynced
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.push(0); // iOutLevelRw
    payload.push(0); // iOutLevelCol
    rw.write_record(records::BRT_WS_FMT_INFO, &payload)
}

fn write_ws_dim<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let mut payload = vec![0u8; 16];
    if let Some(range) = ws.used_range() {
        payload[0..4].copy_from_slice(&range.start.row.to_le_bytes());
        payload[4..8].copy_from_slice(&range.end.row.to_le_bytes());
        payload[8..12].copy_from_slice(&(range.start.col as u32).to_le_bytes());
        payload[12..16].copy_from_slice(&(range.end.col as u32).to_le_bytes());
    }
    rw.write_record(records::BRT_WS_DIM, &payload)
}

fn write_sheet_views<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let freeze = ws.freeze_panes();

    rw.write_record(records::BRT_BEGIN_VIEWS, &[])?;

    let mut sv = vec![0u8; 30];
    let mut flags: u32 = 0x03DC;
    if freeze.is_some() {
        flags |= 0x08;
    }
    if ws.is_selected() {
        flags |= 0x02;
    }
    sv[0..4].copy_from_slice(&flags.to_le_bytes());
    sv[14] = 0x40; // icvHdr
    let zoom = ws.zoom_scale().unwrap_or(100);
    sv[18..20].copy_from_slice(&zoom.to_le_bytes()); // wScale
    rw.write_record(records::BRT_BEGIN_SHEET_VIEW, &sv)?;

    if let Some(fp) = freeze {
        write_freeze_pane(rw, fp.row, fp.col)?;
    } else if let Some(sp) = ws.split_panes() {
        write_split_pane(rw, sp)?;
    }

    let (ac_row, ac_col) = ws.selection_active_cell().unwrap_or((0, 0));
    let mut sel = vec![0u8; 36];
    sel[0..4].copy_from_slice(&3u32.to_le_bytes());
    sel[4..8].copy_from_slice(&ac_row.to_le_bytes());
    sel[8..12].copy_from_slice(&(ac_col as u32).to_le_bytes());
    sel[16..20].copy_from_slice(&1u32.to_le_bytes());
    sel[20..24].copy_from_slice(&ac_row.to_le_bytes());
    sel[24..28].copy_from_slice(&ac_row.to_le_bytes());
    sel[28..32].copy_from_slice(&(ac_col as u32).to_le_bytes());
    sel[32..36].copy_from_slice(&(ac_col as u32).to_le_bytes());
    rw.write_record(records::BRT_SEL, &sel)?;

    rw.write_record(records::BRT_END_SHEET_VIEW, &[])?;
    rw.write_record(records::BRT_END_VIEWS, &[])
}

fn write_freeze_pane<W: Write>(
    rw: &mut RecordWriter<W>,
    row: u32,
    col: u16,
) -> std::io::Result<()> {
    let mut payload = vec![0u8; 29];
    payload[0..8].copy_from_slice(&(col as f64).to_le_bytes());
    payload[8..16].copy_from_slice(&(row as f64).to_le_bytes());
    payload[16..20].copy_from_slice(&row.to_le_bytes());
    payload[20..24].copy_from_slice(&(col as u32).to_le_bytes());
    payload[24..28].copy_from_slice(&3u32.to_le_bytes());
    payload[28] = 1; // frozen
    rw.write_record(records::BRT_PANE, &payload)
}

fn write_split_pane<W: Write>(
    rw: &mut RecordWriter<W>,
    sp: &duke_sheets_core::worksheet::SplitPanes,
) -> std::io::Result<()> {
    let mut payload = vec![0u8; 29];
    payload[0..8].copy_from_slice(&sp.x_split.to_le_bytes());
    payload[8..16].copy_from_slice(&sp.y_split.to_le_bytes());
    let (tl_row, tl_col) = sp.top_left.unwrap_or((0, 0));
    payload[16..20].copy_from_slice(&tl_row.to_le_bytes());
    payload[20..24].copy_from_slice(&(tl_col as u32).to_le_bytes());
    let active_pane: u32 = match sp.active_pane.as_deref() {
        Some("topRight") => 1,
        Some("bottomLeft") => 2,
        Some("bottomRight") => 3,
        _ => 0,
    };
    payload[24..28].copy_from_slice(&active_pane.to_le_bytes());
    payload[28] = 0; // split (not frozen)
    rw.write_record(records::BRT_PANE, &payload)
}

fn write_col_infos<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let custom_widths = ws.custom_column_widths();
    let hidden_cols = ws.hidden_columns();
    let outline_levels = ws.column_outline_levels();

    if custom_widths.is_empty() && hidden_cols.is_empty() && outline_levels.is_empty() {
        return Ok(());
    }

    let mut all_cols: BTreeMap<u16, (Option<f64>, bool, u8)> = BTreeMap::new();
    for (&col, &width) in custom_widths {
        all_cols.entry(col).or_insert((None, false, 0)).0 = Some(width);
    }
    for (&col, &hidden) in hidden_cols {
        if hidden {
            all_cols.entry(col).or_insert((None, false, 0)).1 = true;
        }
    }
    for (&col, &level) in outline_levels {
        if level > 0 {
            all_cols.entry(col).or_insert((None, false, 0)).2 = level;
        }
    }

    rw.write_record(records::BRT_BEGIN_COL_INFOS, &[])?;

    for (&col, &(width_opt, hidden, outline_level)) in &all_cols {
        let mut payload = vec![0u8; 18];
        let col32 = col as u32;
        payload[0..4].copy_from_slice(&col32.to_le_bytes());
        payload[4..8].copy_from_slice(&col32.to_le_bytes());

        let coldx = if let Some(w) = width_opt {
            (w * 256.0).round() as u32
        } else {
            (8.43_f64 * 256.0).round() as u32
        };
        payload[8..12].copy_from_slice(&coldx.to_le_bytes());
        payload[12..16].copy_from_slice(&0u32.to_le_bytes());

        let mut flags: u16 = 0;
        if hidden {
            flags |= 0x01;
        }
        if width_opt.is_some() {
            flags |= 0x02;
        }
        flags |= ((outline_level & 0x07) as u16) << 8;
        payload[16..18].copy_from_slice(&flags.to_le_bytes());

        rw.write_record(records::BRT_COL_INFO, &payload)?;
    }

    rw.write_record(records::BRT_END_COL_INFOS, &[])
}

struct CellInfo {
    col: u16,
    style_ref: u32,
    value: CellValue,
    formula_text: Option<String>,
}

fn collect_rows(
    ws: &Worksheet,
    sheet_index: usize,
    style_mapping: &StyleMapping,
) -> BTreeMap<u32, Vec<CellInfo>> {
    let mut rows: BTreeMap<u32, Vec<CellInfo>> = BTreeMap::new();
    for (row, col, cell) in ws.iter_cells() {
        let formula_text = ws.formula_data_at(row, col).map(|fd| fd.text.clone());
        if cell.value.is_empty() && cell.style_index == 0 && formula_text.is_none() {
            continue;
        }
        let style_ref = style_mapping.xf_index(sheet_index, cell.style_index);
        rows.entry(row).or_default().push(CellInfo {
            col,
            style_ref,
            value: cell.value.clone(),
            formula_text,
        });
    }
    for (row, col, _) in ws.formula_cells() {
        let cells = rows.entry(row).or_default();
        if cells.iter().any(|c| c.col == col) {
            continue;
        }
        let formula_text = ws.formula_data_at(row, col).map(|fd| fd.text.clone());
        cells.push(CellInfo {
            col,
            style_ref: 0,
            value: CellValue::Empty,
            formula_text,
        });
    }
    for cells in rows.values_mut() {
        cells.sort_by_key(|c| c.col);
    }
    rows
}

fn write_row_hdr<W: Write>(
    rw: &mut RecordWriter<W>,
    row: u32,
    ws: &Worksheet,
    col_range: Option<(u32, u32)>,
) -> std::io::Result<()> {
    let custom_heights = ws.custom_row_heights();
    let has_custom_height = custom_heights.contains_key(&row);
    let hidden = ws.is_row_hidden(row);

    let miy_rw: u16 = if has_custom_height {
        (ws.row_height(row) * 20.0).round() as u16
    } else {
        300
    };

    // grbitRw byte 1 bits: 4=fDyZero(hidden), 5=fUnsynced(custom height)
    let mut flags: u8 = 0;
    if hidden {
        flags |= 0x10;
    }
    if has_custom_height {
        flags |= 0x20;
    }

    let ccolspan: u32 = if col_range.is_some() { 1 } else { 0 };
    let outline_level = ws.row_outline_level(row);

    let mut payload = Vec::with_capacity(17 + 8 * ccolspan as usize);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&miy_rw.to_le_bytes());
    payload.push(outline_level & 0x07);
    payload.push(flags);
    payload.push(0);
    payload.extend_from_slice(&ccolspan.to_le_bytes());
    if let Some((col_min, col_max)) = col_range {
        payload.extend_from_slice(&col_min.to_le_bytes());
        payload.extend_from_slice(&col_max.to_le_bytes());
    }

    rw.write_record(records::BRT_ROW_HDR, &payload)
}

fn cell_prefix(col: u16, style_ref: u32) -> [u8; 8] {
    let mut prefix = [0u8; 8];
    prefix[0..4].copy_from_slice(&(col as u32).to_le_bytes());
    prefix[4] = (style_ref & 0xFF) as u8;
    prefix[5] = ((style_ref >> 8) & 0xFF) as u8;
    prefix[6] = ((style_ref >> 16) & 0xFF) as u8;
    prefix
}

fn write_cell<W: Write>(
    rw: &mut RecordWriter<W>,
    cell: &CellInfo,
    sst: &SstMap,
    compile_ctx: &CompileContext,
) -> std::io::Result<()> {
    let prefix = cell_prefix(cell.col, cell.style_ref);

    if let Some(ref formula_text) = cell.formula_text {
        return write_formula_cell(rw, &prefix, &cell.value, formula_text, compile_ctx);
    }

    match &cell.value {
        CellValue::Number(n) => {
            let mut payload = Vec::with_capacity(16);
            payload.extend_from_slice(&prefix);
            payload.extend_from_slice(&n.to_le_bytes());
            rw.write_record(records::BRT_CELL_REAL, &payload)
        }
        CellValue::String(s) => {
            let s_str = s.as_str();
            if let Some(idx) = sst.get_plain(s_str) {
                let mut payload = Vec::with_capacity(12);
                payload.extend_from_slice(&prefix);
                payload.extend_from_slice(&idx.to_le_bytes());
                rw.write_record(records::BRT_CELL_ISST, &payload)
            } else {
                let mut payload = Vec::with_capacity(8 + 4 + s_str.len() * 2);
                payload.extend_from_slice(&prefix);
                payload.extend_from_slice(&encode_wide_str(s_str));
                rw.write_record(records::BRT_CELL_ST, &payload)
            }
        }
        CellValue::Boolean(b) => {
            let mut payload = Vec::with_capacity(9);
            payload.extend_from_slice(&prefix);
            payload.push(if *b { 1 } else { 0 });
            rw.write_record(records::BRT_CELL_BOOL, &payload)
        }
        CellValue::Error(e) => {
            let mut payload = Vec::with_capacity(9);
            payload.extend_from_slice(&prefix);
            payload.push(error_code(e));
            rw.write_record(records::BRT_CELL_ERROR, &payload)
        }
        CellValue::Empty => {
            if cell.style_ref != 0 {
                rw.write_record(records::BRT_CELL_BLANK, &prefix)
            } else {
                Ok(())
            }
        }
        CellValue::RichText(runs) => {
            if let Some(idx) = sst.get_rich(runs) {
                let mut payload = Vec::with_capacity(12);
                payload.extend_from_slice(&prefix);
                payload.extend_from_slice(&idx.to_le_bytes());
                rw.write_record(records::BRT_CELL_ISST, &payload)
            } else {
                let plain: String = runs.iter().map(|r| r.text.as_str()).collect();
                let mut payload = Vec::with_capacity(8 + 4 + plain.len() * 2);
                payload.extend_from_slice(&prefix);
                payload.extend_from_slice(&encode_wide_str(&plain));
                rw.write_record(records::BRT_CELL_ST, &payload)
            }
        }
        CellValue::SpillTarget { .. } => Ok(()),
    }
}

fn write_formula_cell<W: Write>(
    rw: &mut RecordWriter<W>,
    prefix: &[u8; 8],
    value: &CellValue,
    formula_text: &str,
    compile_ctx: &CompileContext,
) -> std::io::Result<()> {
    let compiled = match compiler::compile_formula(formula_text, compile_ctx) {
        Ok(c) => c,
        Err(e) => {
            log::warn!("formula compilation failed for '{}': {}", formula_text, e);
            compiler::CompiledFormula {
                rgce: vec![],
                rgcb: vec![],
            }
        }
    };

    let formula_tail = {
        let grbit = 0u16;
        let cce = compiled.rgce.len() as u32;
        let cb = compiled.rgcb.len() as u32;
        let mut tail = Vec::with_capacity(2 + 4 + compiled.rgce.len() + 4 + compiled.rgcb.len());
        tail.extend_from_slice(&grbit.to_le_bytes());
        tail.extend_from_slice(&cce.to_le_bytes());
        tail.extend_from_slice(&compiled.rgce);
        tail.extend_from_slice(&cb.to_le_bytes());
        tail.extend_from_slice(&compiled.rgcb);
        tail
    };

    match value {
        CellValue::Number(n) => {
            let mut payload = Vec::with_capacity(8 + 8 + formula_tail.len());
            payload.extend_from_slice(prefix);
            payload.extend_from_slice(&n.to_le_bytes());
            payload.extend_from_slice(&formula_tail);
            rw.write_record(records::BRT_FMLA_NUM, &payload)
        }
        CellValue::Boolean(b) => {
            let mut payload = Vec::with_capacity(8 + 1 + formula_tail.len());
            payload.extend_from_slice(prefix);
            payload.push(if *b { 1 } else { 0 });
            payload.extend_from_slice(&formula_tail);
            rw.write_record(records::BRT_FMLA_BOOL, &payload)
        }
        CellValue::Error(e) => {
            let mut payload = Vec::with_capacity(8 + 1 + formula_tail.len());
            payload.extend_from_slice(prefix);
            payload.push(error_code(e));
            payload.extend_from_slice(&formula_tail);
            rw.write_record(records::BRT_FMLA_ERROR, &payload)
        }
        CellValue::String(s) => {
            let s_str = s.as_str();
            let mut payload = Vec::with_capacity(8 + 4 + s_str.len() * 2 + formula_tail.len());
            payload.extend_from_slice(prefix);
            payload.extend_from_slice(&encode_wide_str(s_str));
            payload.extend_from_slice(&formula_tail);
            rw.write_record(records::BRT_FMLA_STRING, &payload)
        }
        _ => {
            let mut payload = Vec::with_capacity(8 + 8 + formula_tail.len());
            payload.extend_from_slice(prefix);
            payload.extend_from_slice(&0.0f64.to_le_bytes());
            payload.extend_from_slice(&formula_tail);
            rw.write_record(records::BRT_FMLA_NUM, &payload)
        }
    }
}

fn write_merge_cells<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let regions = ws.merged_regions();
    if regions.is_empty() {
        return Ok(());
    }

    let count = regions.len() as u32;
    rw.write_record(records::BRT_BEGIN_MERGE_CELLS, &count.to_le_bytes())?;

    for region in regions {
        let mut payload = vec![0u8; 16];
        payload[0..4].copy_from_slice(&region.start.row.to_le_bytes());
        payload[4..8].copy_from_slice(&region.end.row.to_le_bytes());
        payload[8..12].copy_from_slice(&(region.start.col as u32).to_le_bytes());
        payload[12..16].copy_from_slice(&(region.end.col as u32).to_le_bytes());
        rw.write_record(records::BRT_MERGE_CELL, &payload)?;
    }

    rw.write_record(records::BRT_END_MERGE_CELLS, &[])
}

fn write_hyperlinks<W: Write>(
    rw: &mut RecordWriter<W>,
    ws: &Worksheet,
    sheet_rels: &mut Vec<SheetRel>,
    rid_counter: &mut u32,
) -> std::io::Result<()> {
    let hyperlinks = ws.hyperlinks();
    if hyperlinks.is_empty() {
        return Ok(());
    }

    let mut sorted: Vec<(&CellAddress, &duke_sheets_core::Hyperlink)> = hyperlinks.iter().collect();
    sorted.sort_by_key(|(addr, _)| (addr.row, addr.col));

    for (addr, link) in sorted {
        let is_external = !link.target.is_empty() && !link.target.starts_with('#');

        let rel_id = if is_external {
            let rid = format!("rId{}", *rid_counter);
            *rid_counter += 1;
            sheet_rels.push(SheetRel {
                id: rid.clone(),
                rel_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                        .to_string(),
                target: link.target.clone(),
                target_mode: Some("External".to_string()),
            });
            rid
        } else {
            String::new()
        };

        let location = if link.target.starts_with('#') {
            link.target[1..].to_string()
        } else {
            link.location.clone().unwrap_or_default()
        };

        let tooltip = link.tooltip.clone().unwrap_or_default();
        let display = link.display.clone().unwrap_or_default();

        let mut payload = Vec::new();
        payload.extend_from_slice(&addr.row.to_le_bytes());
        payload.extend_from_slice(&addr.row.to_le_bytes());
        payload.extend_from_slice(&(addr.col as u32).to_le_bytes());
        payload.extend_from_slice(&(addr.col as u32).to_le_bytes());
        payload.extend_from_slice(&encode_wide_str(&rel_id));
        payload.extend_from_slice(&encode_wide_str(&location));
        payload.extend_from_slice(&encode_wide_str(&tooltip));
        payload.extend_from_slice(&encode_wide_str(&display));

        rw.write_record(records::BRT_H_LINK, &payload)?;
    }

    Ok(())
}

fn write_auto_filter<W: Write>(
    rw: &mut RecordWriter<W>,
    ws: &Worksheet,
    color_filter_dxf_id: Option<u32>,
) -> std::io::Result<()> {
    use duke_sheets_core::auto_filter::{ColumnFilter, FilterOperator};

    let af = match ws.auto_filter() {
        Some(af) => af,
        None => return Ok(()),
    };

    let range = &af.range;
    let mut payload = vec![0u8; 16];
    payload[0..4].copy_from_slice(&range.start.row.to_le_bytes());
    payload[4..8].copy_from_slice(&range.end.row.to_le_bytes());
    payload[8..12].copy_from_slice(&(range.start.col as u32).to_le_bytes());
    payload[12..16].copy_from_slice(&(range.end.col as u32).to_le_bytes());
    rw.write_record(records::BRT_BEGIN_A_FILTER, &payload)?;

    for fc in &af.filter_columns {
        // BrtBeginFilterColumn per [MS-XLSB] §2.4.84:
        //   dwCol u32 + flags u16 (fHideArrow bit 0, fNoBtn bit 1).
        let mut col_payload = Vec::with_capacity(6);
        col_payload.extend_from_slice(&(fc.col_id as u32).to_le_bytes());
        let mut fc_flags: u16 = 0;
        if fc.hidden_button {
            fc_flags |= 1; // fHideArrow
        }
        if !fc.show_button {
            fc_flags |= 1 << 1; // fNoBtn
        }
        col_payload.extend_from_slice(&fc_flags.to_le_bytes());
        rw.write_record(records::BRT_BEGIN_FILTER_COLUMN, &col_payload)?;

        match &fc.filter {
            ColumnFilter::Top10(t) => {
                // BrtTop10Filter per §2.4.852: 1 byte flags + 8 bytes
                // xNumValue + 8 bytes xNumFilter. flags = fTop bit 0,
                // fPercent bit 1, fApplied bit 2, reserved 5 bits.
                let mut flags: u8 = 0;
                if t.top {
                    flags |= 1;
                }
                if t.percent {
                    flags |= 1 << 1;
                }
                // fApplied: when set, xNumFilter MUST be an actual
                // value from the filtered range. Only claim it when we
                // carry a computed filter value.
                if t.filter_val.is_some() {
                    flags |= 1 << 2;
                }
                let mut top_payload = Vec::with_capacity(17);
                top_payload.push(flags);
                top_payload.extend_from_slice(&t.val.to_le_bytes());
                top_payload.extend_from_slice(&t.filter_val.unwrap_or(t.val).to_le_bytes());
                rw.write_record(records::BRT_TOP10_FILTER, &top_payload)?;
            }
            ColumnFilter::Values(v) => {
                // BrtBeginFilters per §2.4.85: u32 fBlank + u32 unused.
                let mut hdr = Vec::with_capacity(8);
                hdr.extend_from_slice(&(v.blank as u32).to_le_bytes());
                hdr.extend_from_slice(&0u32.to_le_bytes());
                rw.write_record(records::BRT_BEGIN_FILTERS, &hdr)?;
                for value in &v.values {
                    let mut buf = Vec::new();
                    buf.extend_from_slice(&encode_wide_str(value));
                    rw.write_record(records::BRT_FILTER, &buf)?;
                }
                rw.write_record(records::BRT_END_FILTERS, &[])?;
            }
            ColumnFilter::Custom(cf) => {
                // BrtBeginCustomFilters carries an i32 AND/OR flag:
                // 0 = AND, nonzero = OR (LO CustomFilter::importRecord).
                let join: i32 = if cf.and { 0 } else { 1 };
                rw.write_record(records::BRT_BEGIN_CUSTOM_FILTERS, &join.to_le_bytes())?;
                for cond in cf.conditions.iter().take(2) {
                    // BrtCustomFilter per §2.4.348 (cross-checked
                    // against LO FilterCriterionModel::readBiffData):
                    //   vts u8 + operator u8 + 8 bytes xNumOrError +
                    //   rgch string, present when vts = 6 (string).
                    // We always store conditions as text strings, so
                    // vts = 6 and the 8 value bytes are zero.
                    let op_byte: u8 = match cond.operator {
                        FilterOperator::LessThan => 1,
                        FilterOperator::Equal => 2,
                        FilterOperator::LessThanOrEqual => 3,
                        FilterOperator::GreaterThan => 4,
                        FilterOperator::NotEqual => 5,
                        FilterOperator::GreaterThanOrEqual => 6,
                    };
                    let mut buf = Vec::with_capacity(20);
                    buf.push(6u8); // vts: vtString
                    buf.push(op_byte);
                    buf.extend_from_slice(&[0u8; 8]); // xNumOrError
                    buf.extend_from_slice(&encode_wide_str(&cond.value));
                    rw.write_record(records::BRT_CUSTOM_FILTER, &buf)?;
                }
                rw.write_record(records::BRT_END_CUSTOM_FILTERS, &[])?;
            }
            ColumnFilter::Dynamic(d) => {
                // BrtDynamicFilter per [MS-XLSB] §2.4.362 (id 0x00AB):
                //   cft u32 (filter type), flags u8 (fApplied bit 0
                //   + 7 reserved), xNumValue f64, xNumValueMax f64.
                //   21 bytes total. Excel emits flags=0 even when the
                //   filter is actively applied, so we match.
                let mut buf = Vec::with_capacity(21);
                buf.extend_from_slice(&dynamic_filter_cft(d.filter_type).to_le_bytes());
                buf.push(0);
                buf.extend_from_slice(&d.val.unwrap_or(0.0).to_le_bytes());
                buf.extend_from_slice(&d.max_val.unwrap_or(0.0).to_le_bytes());
                rw.write_record(records::BRT_DYNAMIC_FILTER, &buf)?;
            }
            ColumnFilter::Color(c) => {
                // BrtColorFilter per [MS-XLSB] §2.4.339 (id 0x00A8):
                //   dxfid u32 + fCellColor u32. 8 bytes.
                // The dxfid must reference a real BrtDXF entry in
                // styles.bin — Excel refuses to open the file when it
                // dangles — so use the writer-synthesized entry, not
                // the model's index into some other file's dxf table.
                let dxfid = color_filter_dxf_id.unwrap_or(0);
                let mut buf = Vec::with_capacity(8);
                buf.extend_from_slice(&dxfid.to_le_bytes());
                buf.extend_from_slice(&(c.cell_color as u32).to_le_bytes());
                rw.write_record(records::BRT_COLOR_FILTER, &buf)?;
            }
        }

        rw.write_record(records::BRT_END_FILTER_COLUMN, &[])?;
    }

    rw.write_record(records::BRT_END_A_FILTER, &[])
}

fn write_margins<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let ps = ws.page_setup();
    let mut payload = vec![0u8; 48];
    payload[0..8].copy_from_slice(&ps.left_margin.to_le_bytes());
    payload[8..16].copy_from_slice(&ps.right_margin.to_le_bytes());
    payload[16..24].copy_from_slice(&ps.top_margin.to_le_bytes());
    payload[24..32].copy_from_slice(&ps.bottom_margin.to_le_bytes());
    payload[32..40].copy_from_slice(&ps.header_margin.to_le_bytes());
    payload[40..48].copy_from_slice(&ps.footer_margin.to_le_bytes());
    rw.write_record(records::BRT_MARGINS, &payload)
}

fn write_page_setup<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let ps = ws.page_setup();

    let is_default = ps.paper_size == 1
        && ps.scale == 100
        && ps.orientation == PageOrientation::Portrait
        && ps.fit_to_width.is_none()
        && ps.fit_to_height.is_none();

    // Real Excel omits BrtPageSetup entirely for default values, so
    // do the same. Excel synthesises sensible defaults on read.
    if is_default {
        return Ok(());
    }

    // BrtPageSetup payload per [MS-XLSB] §2.4.722:
    //   iPaperSize, iScale, iRes, iVRes, iCopies, iPageStart,
    //   iFitWidth, iFitHeight  (8 × u32 = 32 bytes)
    //   flags (u16):
    //     bit 0 fLeftToRight, bit 1 fLandscape, bit 2 reserved1,
    //     bit 3 fNoColor, bit 4 fDraft, bit 5 fNotes,
    //     bit 6 fNoOrient, bit 7 fUsePage, bit 8 fEndNotes,
    //     bits 9-10 iErrors (2 bits), bits 11-15 reserved2 (5 bits)
    //   szRelID (XLNullableWideString, NULL = u32 0xFFFFFFFF)
    let mut payload = Vec::with_capacity(38);
    payload.extend_from_slice(&(ps.paper_size as u32).to_le_bytes());
    payload.extend_from_slice(&(ps.scale as u32).to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes()); // iRes
    payload.extend_from_slice(&0u32.to_le_bytes()); // iVRes
    payload.extend_from_slice(&1u32.to_le_bytes()); // iCopies
    payload.extend_from_slice(&1u32.to_le_bytes()); // iPageStart
    payload.extend_from_slice(&(ps.fit_to_width.unwrap_or(0) as u32).to_le_bytes());
    payload.extend_from_slice(&(ps.fit_to_height.unwrap_or(0) as u32).to_le_bytes());

    let mut flags: u16 = 0;
    if ps.orientation == PageOrientation::Landscape {
        flags |= 1 << 1;
    }
    payload.extend_from_slice(&flags.to_le_bytes());

    // szRelID NULL marker.
    payload.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes());

    rw.write_record(records::BRT_PAGE_SETUP, &payload)
}

fn write_print_options<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let ps = ws.page_setup();
    if !ps.print_headings && !ps.print_gridlines {
        return Ok(());
    }
    let mut grbit: u16 = 0;
    if ps.print_headings {
        grbit |= 0x04;
    }
    if ps.print_gridlines {
        grbit |= 0x08;
    }
    rw.write_record(records::BRT_PRINT_OPTIONS, &grbit.to_le_bytes())
}

fn write_header_footer<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let ps = ws.page_setup();

    let has_content = ps.odd_header.is_some()
        || ps.odd_footer.is_some()
        || ps.even_header.is_some()
        || ps.even_footer.is_some()
        || ps.first_header.is_some()
        || ps.first_footer.is_some()
        || ps.different_odd_even
        || ps.different_first;

    if !has_content {
        return Ok(());
    }

    // BrtBeginHeaderFooter per [MS-XLSB] §2.4.91:
    //   2 bytes flags (4 named bits + 12 reserved)
    //   6 HeaderFooterString fields (XLNullableWideString):
    //     stHeader, stFooter, stHeaderEven, stFooterEven,
    //     stHeaderFirst, stFooterFirst.
    // The block is closed by BrtEndHeaderFooter.
    let mut flags: u16 = 0;
    if ps.different_odd_even {
        flags |= 0x01;
    }
    if ps.different_first {
        flags |= 0x02;
    }
    if ps.scale_with_doc {
        flags |= 0x04;
    }
    if ps.align_with_margins {
        flags |= 0x08;
    }

    let mut payload = Vec::new();
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&encode_nullable_wide_str(ps.odd_header.as_deref()));
    payload.extend_from_slice(&encode_nullable_wide_str(ps.odd_footer.as_deref()));
    payload.extend_from_slice(&encode_nullable_wide_str(ps.even_header.as_deref()));
    payload.extend_from_slice(&encode_nullable_wide_str(ps.even_footer.as_deref()));
    payload.extend_from_slice(&encode_nullable_wide_str(ps.first_header.as_deref()));
    payload.extend_from_slice(&encode_nullable_wide_str(ps.first_footer.as_deref()));

    rw.write_record(records::BRT_HEADER_FOOTER, &payload)?;
    rw.write_record(records::BRT_END_HEADER_FOOTER, &[])
}

fn write_sheet_protection<W: Write>(
    rw: &mut RecordWriter<W>,
    ws: &Worksheet,
) -> std::io::Result<()> {
    let prot = match ws.protection() {
        Some(p) => p,
        None => return Ok(()),
    };

    let mut payload = Vec::with_capacity(66);
    let password = prot.password_hash.unwrap_or(0);
    payload.extend_from_slice(&password.to_le_bytes()); // u16
    for &val in &[
        prot.protected,
        true, // fObjects
        true, // fScenarios
        prot.format_cells,
        prot.format_columns,
        prot.format_rows,
        prot.insert_columns,
        prot.insert_rows,
        prot.insert_hyperlinks,
        prot.delete_columns,
        prot.delete_rows,
        prot.select_locked_cells,
        prot.sort,
        prot.auto_filter,
        prot.pivot_tables,
        prot.select_unlocked_cells,
    ] {
        payload.extend_from_slice(&(val as u32).to_le_bytes());
    }
    rw.write_record(records::BRT_SHEET_PROTECTION, &payload)
}

fn write_page_breaks<W: Write>(rw: &mut RecordWriter<W>, ws: &Worksheet) -> std::io::Result<()> {
    let row_breaks = ws.row_breaks();
    if !row_breaks.is_empty() {
        let man_count = row_breaks.iter().filter(|b| b.man).count() as u32;
        let mut header = Vec::with_capacity(8);
        header.extend_from_slice(&(row_breaks.len() as u32).to_le_bytes());
        header.extend_from_slice(&man_count.to_le_bytes());
        rw.write_record(records::BRT_BEGIN_RW_BRK, &header)?;
        for brk in row_breaks {
            let mut payload = Vec::with_capacity(20);
            payload.extend_from_slice(&brk.id.to_le_bytes());
            payload.extend_from_slice(&brk.min.to_le_bytes());
            payload.extend_from_slice(&brk.max.to_le_bytes());
            payload.extend_from_slice(&(brk.man as u32).to_le_bytes());
            payload.extend_from_slice(&(brk.pt as u32).to_le_bytes());
            rw.write_record(records::BRT_BRK, &payload)?;
        }
        rw.write_record(records::BRT_END_RW_BRK, &[])?;
    }

    let col_breaks = ws.col_breaks();
    if !col_breaks.is_empty() {
        let man_count = col_breaks.iter().filter(|b| b.man).count() as u32;
        let mut header = Vec::with_capacity(8);
        header.extend_from_slice(&(col_breaks.len() as u32).to_le_bytes());
        header.extend_from_slice(&man_count.to_le_bytes());
        rw.write_record(records::BRT_BEGIN_COL_BRK, &header)?;
        for brk in col_breaks {
            let mut payload = Vec::with_capacity(20);
            payload.extend_from_slice(&brk.id.to_le_bytes());
            payload.extend_from_slice(&brk.min.to_le_bytes());
            payload.extend_from_slice(&brk.max.to_le_bytes());
            payload.extend_from_slice(&(brk.man as u32).to_le_bytes());
            payload.extend_from_slice(&(brk.pt as u32).to_le_bytes());
            rw.write_record(records::BRT_BRK, &payload)?;
        }
        rw.write_record(records::BRT_END_COL_BRK, &[])?;
    }

    Ok(())
}

fn write_data_validations<W: Write>(
    rw: &mut RecordWriter<W>,
    ws: &Worksheet,
    compile_ctx: &CompileContext,
) -> std::io::Result<()> {
    let validations = ws.data_validations();
    if validations.is_empty() {
        return Ok(());
    }

    // BrtBeginDVals carries an 18-byte DVals payload per [MS-XLSB]
    // §2.5.36: fWnClosed bit + 15 reserved bits, xLeft u32, yTop
    // u32, unused3 u32, idvMac u32 (count of BrtDVal records).
    let mut begin_payload = Vec::with_capacity(18);
    begin_payload.extend_from_slice(&0u16.to_le_bytes()); // fWnClosed=0 + reserved
    begin_payload.extend_from_slice(&0u32.to_le_bytes()); // xLeft
    begin_payload.extend_from_slice(&0u32.to_le_bytes()); // yTop
    begin_payload.extend_from_slice(&0u32.to_le_bytes()); // unused3
    begin_payload.extend_from_slice(&(validations.len() as u32).to_le_bytes()); // idvMac
    rw.write_record(records::BRT_BEGIN_DVAL, &begin_payload)?;

    for dv in validations {
        let payload = build_dval_payload(dv, compile_ctx);
        rw.write_record(records::BRT_DVAL, &payload)?;
    }

    rw.write_record(records::BRT_END_DVAL, &[])
}

/// Build a BrtDVal record payload per [MS-XLSB] §2.4.356.
///
/// Layout (no fixed-size header beyond the bit-packed u32):
///   - 4 bytes: bit-packed header
///       bits 0-3:   valType (4)
///       bits 4-6:   errStyle (3)
///       bit  7:     unused (1)
///       bit  8:     fAllowBlank (1)
///       bit  9:     fSuppressCombo (1)
///       bits 10-17: mdImeMode (8)
///       bit  18:    fShowInputMsg (1)
///       bit  19:    fShowErrorMsg (1)
///       bits 20-23: typOperator (4)
///       bits 24-31: reserved (8) — MUST be 0
///   - sqrfx: UncheckedSqRfX (cFx u32 + cFx × UncheckedRfX 16 bytes each)
///   - DValStrings: 4 XLNullableWideStrings
///       (strErrorTitle, strError, strPromptTitle, strPrompt)
///   - formula1: DVParsedFormula (cce u32 + rgce + cb u32 + rgcb)
///   - formula2: DVParsedFormula (cce u32 + rgce + cb u32 + rgcb)
fn build_dval_payload(dv: &DataValidation, compile_ctx: &CompileContext) -> Vec<u8> {
    let mut payload = Vec::new();

    let val_type: u32 = match &dv.validation_type {
        ValidationType::None => 0,
        ValidationType::Whole { .. } => 1,
        ValidationType::Decimal { .. } => 2,
        ValidationType::List { .. } => 3,
        ValidationType::Date { .. } => 4,
        ValidationType::Time { .. } => 5,
        ValidationType::TextLength { .. } => 6,
        ValidationType::Custom { .. } => 7,
    };
    let err_style: u32 = match dv.error_style {
        ValidationErrorStyle::Stop => 0,
        ValidationErrorStyle::Warning => 1,
        ValidationErrorStyle::Information => 2,
    };
    let typ_operator: u32 = match &dv.validation_type {
        ValidationType::Whole { operator, .. }
        | ValidationType::Decimal { operator, .. }
        | ValidationType::Date { operator, .. }
        | ValidationType::Time { operator, .. }
        | ValidationType::TextLength { operator, .. } => validation_op_code(operator),
        _ => 0u32,
    };

    let suppress_combo = !dv.show_dropdown;

    let mut header: u32 = 0;
    header |= val_type & 0xF;
    header |= (err_style & 0x7) << 4;
    // fStrLookup (bit 7): formula1 is a literal string list that the
    // application splits into dropdown entries.
    if matches!(&dv.validation_type, ValidationType::List { source } if list_source_is_literal(source))
    {
        header |= 1 << 7;
    }
    if dv.allow_blank {
        header |= 1 << 8;
    }
    if suppress_combo {
        header |= 1 << 9;
    }
    if dv.show_input_message {
        header |= 1 << 18;
    }
    if dv.show_error_alert {
        header |= 1 << 19;
    }
    header |= (typ_operator & 0xF) << 20;
    payload.extend_from_slice(&header.to_le_bytes());

    // sqrfx: UncheckedSqRfX
    let count = dv.ranges.len() as u32;
    payload.extend_from_slice(&count.to_le_bytes());
    for range in &dv.ranges {
        payload.extend_from_slice(&range.start.row.to_le_bytes());
        payload.extend_from_slice(&range.end.row.to_le_bytes());
        payload.extend_from_slice(&(range.start.col as u32).to_le_bytes());
        payload.extend_from_slice(&(range.end.col as u32).to_le_bytes());
    }

    // DValStrings: 4 XLNullableWideStrings
    payload.extend_from_slice(&encode_nullable_wide_str(dv.error_title.as_deref()));
    payload.extend_from_slice(&encode_nullable_wide_str(dv.error_message.as_deref()));
    payload.extend_from_slice(&encode_nullable_wide_str(dv.input_title.as_deref()));
    payload.extend_from_slice(&encode_nullable_wide_str(dv.input_message.as_deref()));

    // formula1, formula2 (DVParsedFormula)
    let (formula1_text, formula2_text) = dv_formula_texts(&dv.validation_type);
    write_dv_parsed_formula(&mut payload, formula1_text.as_deref(), compile_ctx);
    write_dv_parsed_formula(&mut payload, formula2_text.as_deref(), compile_ctx);

    payload
}

/// Extracts the formula1 and formula2 text from a validation type.
///
/// For List, formula1 is the source string wrapped in quotes (so it
/// compiles to a tStr token). For numeric/date/time/textLength types
/// formula1=value1 and formula2=value2 (when present). For Custom
/// formula1 is the formula. None has no formulas.
fn dv_formula_texts(vt: &ValidationType) -> (Option<String>, Option<String>) {
    match vt {
        ValidationType::None => (None, None),
        ValidationType::Whole { value1, value2, .. }
        | ValidationType::Decimal { value1, value2, .. }
        | ValidationType::Date { value1, value2, .. }
        | ValidationType::Time { value1, value2, .. }
        | ValidationType::TextLength { value1, value2, .. } => {
            (Some(value1.clone()), value2.clone())
        }
        ValidationType::List { source } => {
            // Literal sources ("Red,Green,Blue") are stored as a quoted
            // string that compiles to a single tStr token (paired with
            // the fStrLookup header bit); range/formula sources compile
            // to reference tokens. Mirrors the XLSX writer's heuristic.
            let text = if let Some(stripped) = source.strip_prefix('=') {
                stripped.to_string()
            } else if list_source_is_literal(source) {
                format!("\"{}\"", source.replace('"', "\"\""))
            } else {
                source.clone()
            };
            (Some(text), None)
        }
        ValidationType::Custom { formula } => (Some(formula.clone()), None),
    }
}

/// Whether a List validation source is a literal value list rather than
/// a range reference or formula. Mirrors the XLSX writer's heuristic.
fn list_source_is_literal(source: &str) -> bool {
    !(source.starts_with('=')
        || source.contains('!')
        || source
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '$' || c == ':'))
}

/// DVParsedFormula MUST NOT contain PtgArray, PtgIsect, or PtgUnion
/// (among others) per [MS-XLSB]; Excel repairs files that carry them.
/// Scan the AST for those constructs before compiling.
fn dv_formula_allowed(expr: &duke_sheets_formula::FormulaExpr) -> bool {
    use duke_sheets_formula::ast::BinaryOperator;
    use duke_sheets_formula::FormulaExpr;
    match expr {
        FormulaExpr::Array(_) => false,
        FormulaExpr::BinaryOp { op, left, right } => {
            !matches!(op, BinaryOperator::Union | BinaryOperator::Intersect)
                && dv_formula_allowed(left)
                && dv_formula_allowed(right)
        }
        FormulaExpr::UnaryOp { operand, .. } => dv_formula_allowed(operand),
        FormulaExpr::Function { args, .. } => args.iter().all(dv_formula_allowed),
        _ => true,
    }
}

/// Encode a DVParsedFormula: cce(u32) + rgce + cb(u32) + rgcb.
///
/// Empty / missing formula is encoded with cce=0 and cb=0 (8 bytes).
fn write_dv_parsed_formula(payload: &mut Vec<u8>, text: Option<&str>, ctx: &CompileContext) {
    let compiled = match text {
        Some(t) if !t.is_empty() => {
            let normalized = if t.starts_with('=') {
                t.to_string()
            } else {
                format!("={t}")
            };
            match duke_sheets_formula::parse_formula(&normalized) {
                Ok(expr) if !dv_formula_allowed(&expr) => {
                    log::warn!(
                        "DV formula '{t}' uses tokens forbidden in DVParsedFormula; dropping"
                    );
                    None
                }
                _ => match compiler::compile_formula(t, ctx) {
                    Ok(c) => Some(c),
                    Err(e) => {
                        log::warn!("DV formula compilation failed for '{t}': {e}");
                        None
                    }
                },
            }
        }
        _ => None,
    };

    let (rgce, rgcb): (&[u8], &[u8]) = match &compiled {
        Some(c) => (&c.rgce, &c.rgcb),
        None => (&[], &[]),
    };
    payload.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
    payload.extend_from_slice(rgce);
    payload.extend_from_slice(&(rgcb.len() as u32).to_le_bytes());
    payload.extend_from_slice(rgcb);
}

fn write_conditional_formats<W: Write>(
    rw: &mut RecordWriter<W>,
    ws: &Worksheet,
    sheet_index: usize,
    dxf_mapping: &super::styles::DxfMapping,
    compile_ctx: &CompileContext,
) -> std::io::Result<()> {
    let rules = ws.conditional_formats();
    if rules.is_empty() {
        return Ok(());
    }

    let mut cond_fmt_payload = Vec::new();
    let ccf = rules.len() as u32;
    cond_fmt_payload.extend_from_slice(&ccf.to_le_bytes());
    cond_fmt_payload.extend_from_slice(&0u32.to_le_bytes()); // fPivot
    let mut all_ranges: Vec<&duke_sheets_core::CellRange> = Vec::new();
    for rule in rules {
        for range in &rule.ranges {
            if !all_ranges.iter().any(|r| {
                r.start.row == range.start.row
                    && r.end.row == range.end.row
                    && r.start.col == range.start.col
                    && r.end.col == range.end.col
            }) {
                all_ranges.push(range);
            }
        }
    }
    cond_fmt_payload.extend_from_slice(&(all_ranges.len() as u32).to_le_bytes());
    for range in &all_ranges {
        cond_fmt_payload.extend_from_slice(&range.start.row.to_le_bytes());
        cond_fmt_payload.extend_from_slice(&range.end.row.to_le_bytes());
        cond_fmt_payload.extend_from_slice(&(range.start.col as u32).to_le_bytes());
        cond_fmt_payload.extend_from_slice(&(range.end.col as u32).to_le_bytes());
    }
    rw.write_record(records::BRT_BEGIN_COND_FMT, &cond_fmt_payload)?;

    for (rule_idx, rule) in rules.iter().enumerate() {
        let (i_type, i_template, i_param, flags_extra): (u32, u32, u32, u16) = match &rule.rule_type
        {
            CfRuleType::CellIs { operator, .. } => {
                (1, 0, cf_op_code(operator), 0) // CF_TEMPLATE_EXPR
            }
            CfRuleType::Expression { .. } => (2, 2, 0, 0), // CF_TEMPLATE_FMLA
            CfRuleType::ColorScale { .. } => (3, 3, 0, 0), // CF_TEMPLATE_GRADIENT
            CfRuleType::DataBar { .. } => (4, 5, 0, 0),    // CF_TEMPLATE_DATABAR
            CfRuleType::IconSet { .. } => (5, 6, 0, 0),    // CF_TEMPLATE_MULTISTATE
            CfRuleType::Top10 {
                rank,
                percent,
                bottom,
            } => {
                let mut f: u16 = 0;
                if *bottom {
                    f |= 0x08; // fBottom = bit3
                }
                if *percent {
                    f |= 0x10; // fPercent = bit4
                }
                (6, 7, *rank, f) // CF_TEMPLATE_FILTER
            }
            CfRuleType::UniqueValues => (7, 11, 0, 0), // CF_TEMPLATE_UNIQUEVALUES
            CfRuleType::DuplicateValues => (8, 12, 0, 0), // CF_TEMPLATE_DUPLICATEVALUES
            CfRuleType::ContainsText { .. } => (9, 8, 0, 0), // CF_TEMPLATE_CONTAINSTEXT
            CfRuleType::ContainsBlanks => (10, 13, 0, 0), // CF_TEMPLATE_CONTAINSBLANKS
            CfRuleType::NotContainsBlanks => (11, 14, 0, 0), // CF_TEMPLATE_CONTAINSNOBLANKS
            CfRuleType::ContainsErrors => (12, 15, 0, 0), // CF_TEMPLATE_CONTAINSERRORS
            CfRuleType::NotContainsErrors => (13, 16, 0, 0), // CF_TEMPLATE_CONTAINSNOERRORS
            CfRuleType::AboveAverage { above, .. } => {
                let mut f: u16 = 0;
                if *above {
                    f |= 0x04; // fAbove = bit2
                }
                let tmpl = if *above { 25 } else { 26 };
                (14, tmpl, 0, f)
            }
            CfRuleType::BeginsWith { .. } => (15, 8, 0, 0),
            CfRuleType::EndsWith { .. } => (16, 8, 0, 0),
            CfRuleType::TimePeriod { period } => (17, 9, time_period_param(*period), 0),
        };

        let dxf_id: u32 = match &rule.rule_type {
            CfRuleType::ColorScale { .. }
            | CfRuleType::DataBar { .. }
            | CfRuleType::IconSet { .. } => 0xFFFFFFFF,
            _ => dxf_mapping
                .dxf_id_for_rule(sheet_index, rule_idx)
                .or(rule.dxf_id)
                .unwrap_or(0xFFFFFFFF),
        };
        let priority = rule.priority as u32;

        let mut flags: u16 = flags_extra;
        if rule.stop_if_true {
            flags |= 0x02; // fStopTrue = bit1
        }

        let fmla1 = compile_cf_formula(&rule.rule_type, 1, compile_ctx);
        let fmla2 = compile_cf_formula(&rule.rule_type, 2, compile_ctx);

        let cb_fmla1 = fmla1.as_ref().map_or(0u32, |f| f.rgce.len() as u32);
        let cb_fmla2 = fmla2.as_ref().map_or(0u32, |f| f.rgce.len() as u32);
        let cb_fmla3 = 0u32;

        let mut payload = Vec::new();
        payload.extend_from_slice(&i_type.to_le_bytes()); // iType
        payload.extend_from_slice(&i_template.to_le_bytes()); // iTemplate
        payload.extend_from_slice(&dxf_id.to_le_bytes()); // dxfId
        payload.extend_from_slice(&priority.to_le_bytes()); // iPri
        payload.extend_from_slice(&i_param.to_le_bytes()); // iParam
        payload.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        payload.extend_from_slice(&0u32.to_le_bytes()); // reserved2
        payload.extend_from_slice(&flags.to_le_bytes()); // flags (u16)
        payload.extend_from_slice(&cb_fmla1.to_le_bytes()); // cbFmla1
        payload.extend_from_slice(&cb_fmla2.to_le_bytes()); // cbFmla2
        payload.extend_from_slice(&cb_fmla3.to_le_bytes()); // cbFmla3

        let text = match &rule.rule_type {
            CfRuleType::ContainsText { text } => Some(text.as_str()),
            CfRuleType::BeginsWith { text } => Some(text.as_str()),
            CfRuleType::EndsWith { text } => Some(text.as_str()),
            _ => None,
        };
        match text {
            Some(t) => payload.extend_from_slice(&encode_wide_str(t)), // XLWideString
            None => payload.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes()), // XLNullableWideString null
        }

        if let Some(ref compiled) = fmla1 {
            payload.extend_from_slice(&(compiled.rgce.len() as u32).to_le_bytes()); // cce
            payload.extend_from_slice(&compiled.rgce); // rgce
            payload.extend_from_slice(&(compiled.rgcb.len() as u32).to_le_bytes()); // cb
            payload.extend_from_slice(&compiled.rgcb); // rgcb
        }
        if let Some(ref compiled) = fmla2 {
            payload.extend_from_slice(&(compiled.rgce.len() as u32).to_le_bytes()); // cce
            payload.extend_from_slice(&compiled.rgce); // rgce
            payload.extend_from_slice(&(compiled.rgcb.len() as u32).to_le_bytes()); // cb
            payload.extend_from_slice(&compiled.rgcb); // rgcb
        }

        rw.write_record(records::BRT_BEGIN_CF_RULE, &payload)?;

        match &rule.rule_type {
            CfRuleType::ColorScale { colors } => {
                rw.write_record(records::BRT_BEGIN_COLOR_SCALE, &[])?;
                for cv in colors {
                    write_cfvo_record(rw, cv.value_type, cfvo_value_to_f64(&cv.value))?;
                }
                for cv in colors {
                    write_cf_color_record(rw, &cv.color)?;
                }
                rw.write_record(records::BRT_END_COLOR_SCALE, &[])?;
            }
            CfRuleType::DataBar {
                min_value,
                max_value,
                color,
                show_value,
                ..
            } => {
                let mut db_payload = Vec::new();
                write_cf_data_bar_payload(
                    &mut db_payload,
                    min_value,
                    max_value,
                    color,
                    *show_value,
                );
                rw.write_record(records::BRT_BEGIN_DATA_BAR, &db_payload)?;
                rw.write_record(records::BRT_END_DATA_BAR, &[])?;
            }
            CfRuleType::IconSet {
                icon_style,
                values,
                reverse,
                show_value,
            } => {
                let mut is_payload = Vec::new();
                write_cf_icon_set_payload(
                    &mut is_payload,
                    *icon_style,
                    values,
                    *reverse,
                    *show_value,
                );
                rw.write_record(records::BRT_BEGIN_ICON_SET, &is_payload)?;
                rw.write_record(records::BRT_END_ICON_SET, &[])?;
            }
            _ => {}
        }

        rw.write_record(records::BRT_END_CF_RULE, &[])?;
    }

    rw.write_record(records::BRT_END_COND_FMT, &[])
}

fn compile_cf_formula(
    rule_type: &CfRuleType,
    index: u8,
    ctx: &CompileContext,
) -> Option<compiler::CompiledFormula> {
    let text = match (rule_type, index) {
        (CfRuleType::CellIs { formula1, .. }, 1) => Some(formula1.as_str()),
        (CfRuleType::CellIs { formula2, .. }, 2) => formula2.as_deref(),
        (CfRuleType::Expression { formula }, 1) => Some(formula.as_str()),
        _ => None,
    };
    let text = text?;
    if text.is_empty() {
        return None;
    }
    match compiler::compile_formula(text, ctx) {
        Ok(compiled) => Some(compiled),
        Err(e) => {
            log::warn!("CF formula compile failed: {e}");
            None
        }
    }
}

fn validation_op_code(op: &ValidationOperator) -> u32 {
    match op {
        ValidationOperator::Between => 0,
        ValidationOperator::NotBetween => 1,
        ValidationOperator::Equal => 2,
        ValidationOperator::NotEqual => 3,
        ValidationOperator::GreaterThan => 4,
        ValidationOperator::LessThan => 5,
        ValidationOperator::GreaterThanOrEqual => 6,
        ValidationOperator::LessThanOrEqual => 7,
    }
}

/// Map our DynamicFilterType enum to the BIFF12 cft codes per
/// [MS-XLSB] §2.4.362. Returns 0 (CFTNIL) for `Null`.
fn dynamic_filter_cft(t: duke_sheets_core::auto_filter::DynamicFilterType) -> u32 {
    use duke_sheets_core::auto_filter::DynamicFilterType as D;
    match t {
        D::Null => 0x00,
        D::AboveAverage => 0x01,
        D::BelowAverage => 0x02,
        D::Tomorrow => 0x08,
        D::Today => 0x09,
        D::Yesterday => 0x0A,
        D::NextWeek => 0x0B,
        D::ThisWeek => 0x0C,
        D::LastWeek => 0x0D,
        D::NextMonth => 0x0E,
        D::ThisMonth => 0x0F,
        D::LastMonth => 0x10,
        D::NextQuarter => 0x11,
        D::ThisQuarter => 0x12,
        D::LastQuarter => 0x13,
        D::NextYear => 0x14,
        D::ThisYear => 0x15,
        D::LastYear => 0x16,
        D::YearToDate => 0x17,
        D::Q1 => 0x18,
        D::Q2 => 0x19,
        D::Q3 => 0x1A,
        D::Q4 => 0x1B,
        D::M1 => 0x1C,
        D::M2 => 0x1D,
        D::M3 => 0x1E,
        D::M4 => 0x1F,
        D::M5 => 0x20,
        D::M6 => 0x21,
        D::M7 => 0x22,
        D::M8 => 0x23,
        D::M9 => 0x24,
        D::M10 => 0x25,
        D::M11 => 0x26,
        D::M12 => 0x27,
    }
}

fn cf_op_code(op: &CfOperator) -> u32 {
    match op {
        CfOperator::Between => 1,
        CfOperator::NotBetween => 2,
        CfOperator::Equal => 3,
        CfOperator::NotEqual => 4,
        CfOperator::GreaterThan => 5,
        CfOperator::LessThan => 6,
        CfOperator::GreaterThanOrEqual => 7,
        CfOperator::LessThanOrEqual => 8,
    }
}

fn time_period_param(period: TimePeriod) -> u32 {
    match period {
        TimePeriod::Today => 0,
        TimePeriod::Yesterday => 1,
        TimePeriod::Tomorrow => 2,
        TimePeriod::Last7Days => 3,
        TimePeriod::ThisWeek => 4,
        TimePeriod::LastWeek => 5,
        TimePeriod::NextWeek => 6,
        TimePeriod::ThisMonth => 7,
        TimePeriod::LastMonth => 8,
        TimePeriod::NextMonth => 9,
    }
}

fn cfvo_type_to_u32(vt: CfValueType) -> u32 {
    match vt {
        CfValueType::Min | CfValueType::AutoMin => 2,
        CfValueType::Max | CfValueType::AutoMax => 3,
        CfValueType::Num => 4,
        CfValueType::Percent => 5,
        CfValueType::Formula => 6,
        CfValueType::Percentile => 7,
    }
}

fn cfvo_value_to_f64(value: &Option<String>) -> f64 {
    value
        .as_deref()
        .and_then(|s| s.parse::<f64>().ok())
        .unwrap_or(0.0)
}

fn encode_cf_color(color: &Color) -> [u8; 8] {
    let mut buf = [0u8; 8];
    match color {
        Color::Rgb { r, g, b } => {
            buf[0] = 5;
            buf[1] = 0xFF;
            buf[4] = *r;
            buf[5] = *g;
            buf[6] = *b;
            buf[7] = 0xFF;
        }
        Color::Argb { a, r, g, b } => {
            buf[0] = 5;
            buf[1] = 0xFF;
            buf[4] = *r;
            buf[5] = *g;
            buf[6] = *b;
            buf[7] = *a;
        }
        Color::Theme { index, tint } => {
            buf[0] = 3;
            buf[1] = *index;
            let tint_i16 = if *tint == 0 {
                0i16
            } else {
                ((*tint as f64 / 100.0) * 32767.0).round() as i16
            };
            buf[2..4].copy_from_slice(&tint_i16.to_le_bytes());
        }
        Color::Indexed(idx) => {
            buf[0] = 1;
            buf[1] = *idx;
        }
        Color::Auto => {
            buf[0] = 1;
            buf[1] = 64;
        }
    }
    buf
}

fn write_cfvo_record<W: Write>(
    rw: &mut RecordWriter<W>,
    vt: CfValueType,
    value: f64,
) -> std::io::Result<()> {
    let mut payload = [0u8; 24];
    payload[0..4].copy_from_slice(&cfvo_type_to_u32(vt).to_le_bytes());
    // cce=0, cb=0 (no formula)
    payload[12..20].copy_from_slice(&value.to_le_bytes());
    // pad=0
    rw.write_record(records::BRT_CFVO, &payload)
}

fn write_cf_color_record<W: Write>(rw: &mut RecordWriter<W>, color: &Color) -> std::io::Result<()> {
    rw.write_record(records::BRT_CF_COLOR, &encode_cf_color(color))
}

fn write_cf_data_bar_payload(
    payload: &mut Vec<u8>,
    min_value: &duke_sheets_core::conditional_format::CfValue,
    max_value: &duke_sheets_core::conditional_format::CfValue,
    color: &Color,
    show_value: bool,
) {
    payload.extend_from_slice(&cfvo_type_to_u32(min_value.value_type).to_le_bytes());
    payload.extend_from_slice(&cfvo_value_to_f64(&min_value.value).to_le_bytes());
    payload.extend_from_slice(&cfvo_type_to_u32(max_value.value_type).to_le_bytes());
    payload.extend_from_slice(&cfvo_value_to_f64(&max_value.value).to_le_bytes());
    payload.extend_from_slice(&encode_cf_color(color));
    let flags: u8 = if show_value { 0x01 } else { 0x00 };
    payload.push(flags);
}

fn icon_set_to_u32(style: IconSetStyle) -> u32 {
    match style {
        IconSetStyle::Arrows3 => 0,
        IconSetStyle::Arrows3Gray => 1,
        IconSetStyle::Flags3 => 2,
        IconSetStyle::TrafficLights3 => 3,
        IconSetStyle::TrafficLights3Black => 4,
        IconSetStyle::Signs3 => 5,
        IconSetStyle::Symbols3 => 6,
        IconSetStyle::Symbols3Circled => 7,
        IconSetStyle::Arrows4 => 8,
        IconSetStyle::Arrows4Gray => 9,
        IconSetStyle::RedToBlack4 => 10,
        IconSetStyle::Rating4 => 11,
        IconSetStyle::TrafficLights4 => 12,
        IconSetStyle::Arrows5 => 13,
        IconSetStyle::Arrows5Gray => 14,
        IconSetStyle::Rating5 => 15,
        IconSetStyle::Quarters5 => 16,
        IconSetStyle::Stars3 => 17,
        IconSetStyle::Triangles3 => 18,
        IconSetStyle::Boxes5 => 19,
    }
}

fn write_cf_icon_set_payload(
    payload: &mut Vec<u8>,
    icon_style: IconSetStyle,
    values: &[duke_sheets_core::conditional_format::CfValue],
    reverse: bool,
    show_value: bool,
) {
    payload.extend_from_slice(&icon_set_to_u32(icon_style).to_le_bytes());
    payload.push(values.len() as u8);
    for v in values {
        payload.extend_from_slice(&cfvo_type_to_u32(v.value_type).to_le_bytes());
        payload.extend_from_slice(&cfvo_value_to_f64(&v.value).to_le_bytes());
    }
    let mut flags: u8 = 0;
    if reverse {
        flags |= 0x01;
    }
    if show_value {
        flags |= 0x02;
    }
    payload.push(flags);
}

fn error_code(e: &CellError) -> u8 {
    match e {
        CellError::Null => 0x00,
        CellError::Div0 => 0x07,
        CellError::Value => 0x0F,
        CellError::Ref => 0x17,
        CellError::Name => 0x1D,
        CellError::Num => 0x24,
        CellError::Na => 0x2A,
        CellError::GettingData => 0x2B,
        _ => 0x0F,
    }
}
