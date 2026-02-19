//! XLS (BIFF8) reader.
//!
//! Opens a Compound File Binary (CFB/OLE2) container, reads the `Workbook`
//! stream, parses BIFF8 records, and populates a `duke_sheets_core::Workbook`.

use std::io::{Cursor, Read, Seek};
use std::path::Path;

use duke_sheets_core::cell::SharedString;
use duke_sheets_core::worksheet::SheetProtection;
use duke_sheets_core::{CellError, CellValue, Style, Workbook};

use crate::biff::formula::{ExternSheetEntry, FormulaContext, NameRecord, SupBook, BUILTIN_NAMES};
use crate::biff::parser::{read_f64, read_rk, read_u16, read_u32};
use crate::biff::records;
use crate::biff::strings::{
    parse_sst_continued, read_character_data, read_short_string, read_unicode_string,
};
use crate::biff::{self, BiffRecord};
use crate::error::{XlsError, XlsResult};
use crate::styles::{self, StyleContext};

/// XLS file reader.
pub struct XlsReader;

/// Metadata for a sheet parsed from the BOUNDSHEET record.
#[derive(Debug)]
struct SheetInfo {
    /// Absolute byte offset of the sheet's BOF in the Workbook stream.
    #[allow(dead_code)]
    offset: u32,
    /// Sheet visibility: 0 = visible, 1 = hidden, 2 = very hidden.
    visibility: u8,
    /// Sheet type: 0 = worksheet, 2 = chart, 6 = macro/VBA.
    sheet_type: u8,
    /// Sheet name.
    name: String,
}

impl XlsReader {
    /// Read an XLS file from a filesystem path.
    pub fn read_file<P: AsRef<Path>>(path: P) -> XlsResult<Workbook> {
        let file = std::fs::File::open(path.as_ref())?;
        Self::read(file)
    }

    /// Read an XLS file from any `Read + Seek` source.
    pub fn read<R: Read + Seek>(reader: R) -> XlsResult<Workbook> {
        // Open CFB container
        let mut cfb = cfb::CompoundFile::open(reader)?;

        // Read the "Workbook" stream (some files use "Book" for BIFF5)
        let stream_path = if cfb.exists("/Workbook") {
            "/Workbook"
        } else if cfb.exists("/Book") {
            "/Book"
        } else {
            return Err(XlsError::InvalidFormat(
                "no Workbook or Book stream found in CFB".into(),
            ));
        };

        let mut stream_data = Vec::new();
        {
            let mut stream = cfb.open_stream(stream_path)?;
            stream.read_to_end(&mut stream_data)?;
        }

        // Parse all BIFF records from the stream
        let mut cursor = Cursor::new(&stream_data);
        let all_records = biff::read_all_records(&mut cursor)?;

        // Phase 1: Parse workbook globals
        let mut sst: Vec<String> = Vec::new();
        let mut sheets: Vec<SheetInfo> = Vec::new();
        let mut date_mode_1904 = false;
        let mut in_globals = false;
        let mut style_ctx = StyleContext::new();
        let mut active_sheet_idx: u16 = 0;
        let mut workbook_protected = false;
        let mut workbook_password_hash: Option<u16> = None;
        let mut supbooks: Vec<SupBook> = Vec::new();
        let mut extern_sheet: Vec<ExternSheetEntry> = Vec::new();
        let mut names: Vec<NameRecord> = Vec::new();

        // Find where globals end by iterating until we see an EOF
        // after the first BOF (globals BOF).
        let mut globals_end_idx = 0;

        for (idx, rec) in all_records.iter().enumerate() {
            match rec.record_type {
                records::BOF => {
                    let (version, dt) = biff::parse_bof(&rec.data)?;
                    if dt == records::BOF_WORKBOOK_GLOBALS {
                        if version != records::BIFF8_VERSION {
                            return Err(XlsError::UnsupportedVersion(format!(
                                "expected BIFF8 (0x0600), got 0x{version:04X}"
                            )));
                        }
                        in_globals = true;
                    }
                }
                records::EOF if in_globals => {
                    globals_end_idx = idx;
                    break;
                }
                records::SST if in_globals => {
                    sst = parse_sst_continued(&rec.data, &rec.continue_offsets)?;
                }
                records::BOUNDSHEET if in_globals => {
                    let info = Self::parse_boundsheet(&rec.data)?;
                    sheets.push(info);
                }
                records::DATEMODE if in_globals => {
                    if rec.data.len() >= 2 {
                        let mode = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        date_mode_1904 = mode == 1;
                    }
                }
                // ── Style records ────────────────────────────────────
                records::FONT if in_globals => {
                    if let Ok(font) = styles::parse_font(&rec.data) {
                        style_ctx.fonts.push(font);
                    }
                }
                records::FORMAT if in_globals => {
                    if let Ok((id, s)) = styles::parse_format(&rec.data) {
                        style_ctx.formats.insert(id, s);
                    }
                }
                records::XF if in_globals => {
                    if let Ok(xf) = styles::parse_xf(&rec.data) {
                        style_ctx.xfs.push(xf);
                    }
                }
                records::PALETTE if in_globals => {
                    let _ = styles::apply_palette(&rec.data, &mut style_ctx.palette);
                }
                records::WINDOW1 if in_globals => {
                    // WINDOW1: bytes 10-11 = active/selected sheet index (u16)
                    if rec.data.len() >= 12 {
                        active_sheet_idx = u16::from_le_bytes([rec.data[10], rec.data[11]]);
                    }
                }
                records::PROTECT if in_globals => {
                    if rec.data.len() >= 2 {
                        let val = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        workbook_protected = val == 1;
                    }
                }
                records::PASSWORD if in_globals => {
                    if rec.data.len() >= 2 {
                        let hash = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        if hash != 0 {
                            workbook_password_hash = Some(hash);
                        }
                    }
                }
                // ── Formula context records ──────────────────────────
                records::SUPBOOK if in_globals => {
                    if let Ok(sb) = Self::parse_supbook(&rec.data) {
                        supbooks.push(sb);
                    }
                }
                records::EXTERNSHEET if in_globals => {
                    if let Ok(entries) = Self::parse_externsheet(&rec.data) {
                        extern_sheet = entries;
                    }
                }
                records::NAME if in_globals => {
                    if let Ok(nr) = Self::parse_name(&rec.data) {
                        names.push(nr);
                    }
                }
                _ => {}
            }
        }

        if globals_end_idx == 0 && !in_globals {
            return Err(XlsError::InvalidFormat(
                "no workbook globals BOF found".into(),
            ));
        }

        // Build the resolved style table (one Style per XF record)
        let style_table = style_ctx.build_style_table();

        // Build the workbook
        let mut workbook = Workbook::empty();
        workbook.settings_mut().date_1904 = date_mode_1904;

        // Build formula context from globals data
        let sheet_names: Vec<String> = sheets.iter().map(|s| s.name.clone()).collect();
        let formula_ctx = FormulaContext {
            sheet_names,
            extern_sheet,
            supbooks,
            names,
        };

        // Phase 2: Parse each worksheet substream
        // The records after globals_end_idx contain per-sheet substreams
        // (BOF..EOF pairs). We match them to SheetInfo entries in order.
        let remaining_records = &all_records[globals_end_idx + 1..];
        let sheet_record_groups = Self::split_sheet_records(remaining_records)?;

        let mut wb_sheet_idx = 0usize; // Index into the workbook's sheets
        for (biff_idx, info) in sheets.iter().enumerate() {
            // Only handle worksheets (type 0), skip charts/macros
            if info.sheet_type != 0 {
                continue;
            }

            // Add the sheet to the workbook
            workbook
                .add_worksheet_with_name(&info.name)
                .map_err(|e| XlsError::Core(e))?;

            let ws = workbook.worksheet_mut(wb_sheet_idx).unwrap();

            // Apply sheet visibility (0 = visible, 1 = hidden, 2 = very hidden)
            if info.visibility != 0 {
                ws.set_visible(false);
            }

            // Get this sheet's records (indexed by BIFF order, not wb order)
            if let Some(sheet_records) = sheet_record_groups.get(biff_idx) {
                Self::parse_sheet_records(sheet_records, ws, &sst, &style_table, &formula_ctx)?;
            }

            wb_sheet_idx += 1;
        }

        // Apply active sheet index (WINDOW1)
        let active = active_sheet_idx as usize;
        if active < workbook.sheet_count() {
            let _ = workbook.set_active_sheet(active);
        }

        // Store workbook-level protection info (currently unused, but parsed)
        let _ = (workbook_protected, workbook_password_hash);

        Ok(workbook)
    }

    /// Parse a BOUNDSHEET record body.
    fn parse_boundsheet(data: &[u8]) -> XlsResult<SheetInfo> {
        let mut offset = 0;
        let abs_offset = read_u32(data, &mut offset)?;
        let visibility = data.get(offset).copied().unwrap_or(0);
        offset += 1;
        let sheet_type = data.get(offset).copied().unwrap_or(0);
        offset += 1;
        let name = read_short_string(data, &mut offset)?;

        Ok(SheetInfo {
            offset: abs_offset,
            visibility,
            sheet_type,
            name,
        })
    }

    /// Split remaining records into per-sheet groups (each BOF..EOF pair is one sheet).
    fn split_sheet_records(records: &[BiffRecord]) -> XlsResult<Vec<Vec<&BiffRecord>>> {
        let mut groups: Vec<Vec<&BiffRecord>> = Vec::new();
        let mut current: Option<Vec<&BiffRecord>> = None;
        let mut depth = 0;

        for rec in records {
            match rec.record_type {
                records::BOF => {
                    if depth == 0 {
                        current = Some(Vec::new());
                    }
                    depth += 1;
                    // Don't include the BOF itself in the records we process
                }
                records::EOF => {
                    depth -= 1;
                    if depth == 0 {
                        if let Some(group) = current.take() {
                            groups.push(group);
                        }
                    }
                }
                _ => {
                    if let Some(ref mut group) = current {
                        group.push(rec);
                    }
                }
            }
        }

        Ok(groups)
    }

    /// Parse cell records from a sheet's record group.
    fn parse_sheet_records(
        records: &[&BiffRecord],
        ws: &mut duke_sheets_core::Worksheet,
        sst: &[String],
        styles: &[Style],
        formula_ctx: &FormulaContext,
    ) -> XlsResult<()> {
        // We need to track the last FORMULA record to associate a STRING record
        let mut pending_formula_cell: Option<(u32, u16)> = None;
        let mut sheet_protected = false;
        let mut sheet_password_hash: Option<u16> = None;

        for rec in records {
            match rec.record_type {
                records::LABELSST => {
                    Self::parse_labelsst(&rec.data, ws, sst, styles)?;
                    pending_formula_cell = None;
                }
                records::LABEL => {
                    Self::parse_label(&rec.data, ws, styles)?;
                    pending_formula_cell = None;
                }
                records::NUMBER => {
                    Self::parse_number(&rec.data, ws, styles)?;
                    pending_formula_cell = None;
                }
                records::RK => {
                    Self::parse_rk(&rec.data, ws, styles)?;
                    pending_formula_cell = None;
                }
                records::MULRK => {
                    Self::parse_mulrk(&rec.data, ws, styles)?;
                    pending_formula_cell = None;
                }
                records::BLANK => {
                    Self::parse_blank(&rec.data, ws, styles)?;
                    pending_formula_cell = None;
                }
                records::MULBLANK => {
                    Self::parse_mulblank(&rec.data, ws, styles)?;
                    pending_formula_cell = None;
                }
                records::BOOLERR => {
                    Self::parse_boolerr(&rec.data, ws, styles)?;
                    pending_formula_cell = None;
                }
                records::FORMULA => {
                    pending_formula_cell = Self::parse_formula(&rec.data, ws, styles, formula_ctx)?;
                }
                records::STRING => {
                    // Cached string value for the preceding FORMULA
                    if let Some((row, col)) = pending_formula_cell.take() {
                        Self::parse_formula_string(&rec.data, ws, row, col)?;
                    }
                }
                records::MERGECELLS => {
                    Self::parse_mergecells(&rec.data, ws)?;
                }
                records::ROW => {
                    Self::parse_row(&rec.data, ws)?;
                }
                records::COLINFO => {
                    Self::parse_colinfo(&rec.data, ws)?;
                }
                records::PROTECT => {
                    if rec.data.len() >= 2 {
                        let val = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        sheet_protected = val == 1;
                    }
                }
                records::PASSWORD => {
                    if rec.data.len() >= 2 {
                        let hash = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        if hash != 0 {
                            sheet_password_hash = Some(hash);
                        }
                    }
                }
                _ => {
                    // Skip unknown/unhandled records
                }
            }
        }

        // Apply sheet protection if the PROTECT record was present
        if sheet_protected {
            ws.set_protection(Some(SheetProtection {
                protected: true,
                password_hash: sheet_password_hash,
                ..Default::default()
            }));
        }

        Ok(())
    }

    // ── Style application helper ─────────────────────────────────────────

    /// Apply a style from the XF table to a cell.
    #[inline]
    fn apply_style(
        ws: &mut duke_sheets_core::Worksheet,
        row: u32,
        col: u16,
        xf_idx: u16,
        styles: &[Style],
    ) -> XlsResult<()> {
        let idx = xf_idx as usize;
        if idx != 0 && idx < styles.len() {
            let style = &styles[idx];
            // Only apply if the style differs from the default
            if *style != Style::default() {
                ws.set_cell_style_at(row, col, style)?;
            }
        }
        Ok(())
    }

    // ── Cell record parsers ──────────────────────────────────────────────

    /// LABELSST: row(2) + col(2) + xf(2) + sst_index(4)
    fn parse_labelsst(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        sst: &[String],
        styles: &[Style],
    ) -> XlsResult<()> {
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let xf_idx = read_u16(data, &mut off)?;
        let sst_idx = read_u32(data, &mut off)? as usize;

        if let Some(s) = sst.get(sst_idx) {
            ws.set_cell_value_at(row, col, CellValue::String(SharedString::new(s)))?;
        }
        Self::apply_style(ws, row, col, xf_idx, styles)?;
        Ok(())
    }

    /// LABEL: row(2) + col(2) + xf(2) + unicode_string
    fn parse_label(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
    ) -> XlsResult<()> {
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let xf_idx = read_u16(data, &mut off)?;
        let text = read_unicode_string(data, &mut off)?;

        ws.set_cell_value_at(row, col, CellValue::String(SharedString::new(&text)))?;
        Self::apply_style(ws, row, col, xf_idx, styles)?;
        Ok(())
    }

    /// NUMBER: row(2) + col(2) + xf(2) + f64(8)
    fn parse_number(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
    ) -> XlsResult<()> {
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let xf_idx = read_u16(data, &mut off)?;
        let value = read_f64(data, &mut off)?;

        ws.set_cell_value_at(row, col, CellValue::Number(value))?;
        Self::apply_style(ws, row, col, xf_idx, styles)?;
        Ok(())
    }

    /// RK: row(2) + col(2) + xf(2) + rk(4)
    fn parse_rk(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
    ) -> XlsResult<()> {
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let xf_idx = read_u16(data, &mut off)?;
        let value = read_rk(data, &mut off)?;

        ws.set_cell_value_at(row, col, CellValue::Number(value))?;
        Self::apply_style(ws, row, col, xf_idx, styles)?;
        Ok(())
    }

    /// MULRK: row(2) + first_col(2) + [xf(2) + rk(4)]* + last_col(2)
    fn parse_mulrk(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
    ) -> XlsResult<()> {
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let first_col = read_u16(data, &mut off)?;

        // last_col is the last 2 bytes of the record
        if data.len() < 6 {
            return Err(XlsError::Parse("MULRK record too short".into()));
        }
        let last_col = u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]]);
        let rk_data_end = data.len() - 2; // exclude the trailing last_col field

        let mut col = first_col;
        while off + 6 <= rk_data_end && col <= last_col {
            let xf_idx = read_u16(data, &mut off)?;
            let value = read_rk(data, &mut off)?;
            ws.set_cell_value_at(row, col, CellValue::Number(value))?;
            Self::apply_style(ws, row, col, xf_idx, styles)?;
            col += 1;
        }

        Ok(())
    }

    /// BLANK: row(2) + col(2) + xf(2)
    /// An empty cell that carries formatting.
    fn parse_blank(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
    ) -> XlsResult<()> {
        if data.len() < 6 {
            return Ok(());
        }
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let xf_idx = read_u16(data, &mut off)?;
        Self::apply_style(ws, row, col, xf_idx, styles)?;
        Ok(())
    }

    /// MULBLANK: row(2) + first_col(2) + [xf(2)]* + last_col(2)
    /// Multiple blank cells with formatting.
    fn parse_mulblank(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
    ) -> XlsResult<()> {
        if data.len() < 6 {
            return Ok(());
        }
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let first_col = read_u16(data, &mut off)?;
        let last_col = u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]]);
        let xf_data_end = data.len() - 2;

        let mut col = first_col;
        while off + 2 <= xf_data_end && col <= last_col {
            let xf_idx = read_u16(data, &mut off)?;
            Self::apply_style(ws, row, col, xf_idx, styles)?;
            col += 1;
        }
        Ok(())
    }

    /// BOOLERR: row(2) + col(2) + xf(2) + value(1) + is_error(1)
    fn parse_boolerr(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
    ) -> XlsResult<()> {
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let xf_idx = read_u16(data, &mut off)?;
        let val = data.get(off).copied().unwrap_or(0);
        off += 1;
        let is_error = data.get(off).copied().unwrap_or(0);

        let cell_value = if is_error != 0 {
            let err = match val {
                0x00 => CellError::Null,
                0x07 => CellError::Div0,
                0x0F => CellError::Value,
                0x17 => CellError::Ref,
                0x1D => CellError::Name,
                0x24 => CellError::Num,
                0x2A => CellError::Na,
                _ => CellError::Value,
            };
            CellValue::Error(err)
        } else {
            CellValue::Boolean(val != 0)
        };

        ws.set_cell_value_at(row, col, cell_value)?;
        Self::apply_style(ws, row, col, xf_idx, styles)?;
        Ok(())
    }

    /// FORMULA: row(2) + col(2) + xf(2) + result(8) + options(2) + reserved(4)
    ///        + cce(2) + formula_tokens(cce bytes)
    ///
    /// Returns the (row, col) if the cached result is a string (meaning a
    /// STRING record should follow).
    fn parse_formula(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
        formula_ctx: &FormulaContext,
    ) -> XlsResult<Option<(u32, u16)>> {
        if data.len() < 20 {
            return Err(XlsError::Parse("FORMULA record too short".into()));
        }

        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let xf_idx = read_u16(data, &mut off)?;

        // 8-byte result field
        let result_bytes = &data[off..off + 8];
        off += 8;

        let _options = read_u16(data, &mut off)?;
        let _reserved = read_u32(data, &mut off)?;
        // off is now 20

        // Decompile formula token bytes into text
        let formula_text = if off + 2 <= data.len() {
            let cce = read_u16(data, &mut off)? as usize;
            if cce > 0 && off + cce <= data.len() {
                let token_bytes = &data[off..off + cce];
                let text = crate::biff::formula::decompile(token_bytes, formula_ctx);
                if text.is_empty() {
                    String::new()
                } else {
                    format!("={}", text)
                }
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        // Check if result is a special type (bytes 6-7 == 0xFFFF)
        let mut return_pending = false;
        if result_bytes[6] == 0xFF && result_bytes[7] == 0xFF {
            let result_type = result_bytes[0];
            match result_type {
                0x00 => {
                    // String — the actual string follows in a STRING record.
                    ws.set_cell_value_at(
                        row,
                        col,
                        CellValue::Formula {
                            text: formula_text,
                            cached_value: None,
                            array_result: None,
                        },
                    )?;
                    return_pending = true;
                }
                0x01 => {
                    let bool_val = result_bytes[2] != 0;
                    ws.set_cell_value_at(
                        row,
                        col,
                        CellValue::Formula {
                            text: formula_text,
                            cached_value: Some(Box::new(CellValue::Boolean(bool_val))),
                            array_result: None,
                        },
                    )?;
                }
                0x02 => {
                    let err = match result_bytes[2] {
                        0x00 => CellError::Null,
                        0x07 => CellError::Div0,
                        0x0F => CellError::Value,
                        0x17 => CellError::Ref,
                        0x1D => CellError::Name,
                        0x24 => CellError::Num,
                        0x2A => CellError::Na,
                        _ => CellError::Value,
                    };
                    ws.set_cell_value_at(
                        row,
                        col,
                        CellValue::Formula {
                            text: formula_text,
                            cached_value: Some(Box::new(CellValue::Error(err))),
                            array_result: None,
                        },
                    )?;
                }
                _ => {
                    // Empty or unknown cached result
                    ws.set_cell_value_at(
                        row,
                        col,
                        CellValue::Formula {
                            text: formula_text,
                            cached_value: None,
                            array_result: None,
                        },
                    )?;
                }
            }
        } else {
            // IEEE 754 double
            let value = f64::from_le_bytes(result_bytes.try_into().unwrap());
            ws.set_cell_value_at(
                row,
                col,
                CellValue::Formula {
                    text: formula_text,
                    cached_value: Some(Box::new(CellValue::Number(value))),
                    array_result: None,
                },
            )?;
        }

        Self::apply_style(ws, row, col, xf_idx, styles)?;

        if return_pending {
            Ok(Some((row, col)))
        } else {
            Ok(None)
        }
    }

    /// STRING record: cached string value for a preceding FORMULA.
    ///
    /// Updates the cached_value of the preceding FORMULA cell while preserving
    /// the decompiled formula text.
    fn parse_formula_string(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        row: u32,
        col: u16,
    ) -> XlsResult<()> {
        let mut off = 0;
        let string_val = read_unicode_string(data, &mut off)?;

        // Preserve the formula text that was already decompiled from the
        // FORMULA record's token bytes.
        let existing_text = match ws.get_value_at(row, col) {
            CellValue::Formula { text, .. } => text.to_string(),
            _ => String::new(),
        };

        ws.set_cell_value_at(
            row,
            col,
            CellValue::Formula {
                text: existing_text,
                cached_value: Some(Box::new(CellValue::String(SharedString::new(&string_val)))),
                array_result: None,
            },
        )?;
        Ok(())
    }

    // ── Structural record parsers ────────────────────────────────────────

    /// MERGECELLS: count(2) + [first_row(2) + last_row(2) + first_col(2) + last_col(2)]*
    fn parse_mergecells(data: &[u8], ws: &mut duke_sheets_core::Worksheet) -> XlsResult<()> {
        let mut off = 0;
        let count = read_u16(data, &mut off)? as usize;

        for _ in 0..count {
            if off + 8 > data.len() {
                break;
            }
            let first_row = read_u16(data, &mut off)? as u32;
            let last_row = read_u16(data, &mut off)? as u32;
            let first_col = read_u16(data, &mut off)?;
            let last_col = read_u16(data, &mut off)?;

            let range = duke_sheets_core::CellRange::new(
                duke_sheets_core::CellAddress::new(first_row, first_col),
                duke_sheets_core::CellAddress::new(last_row, last_col),
            );
            let _ = ws.merge_cells(&range);
        }

        Ok(())
    }

    /// ROW: row_index(2) + first_col(2) + last_col_plus1(2) + height(2) + ...
    fn parse_row(data: &[u8], ws: &mut duke_sheets_core::Worksheet) -> XlsResult<()> {
        if data.len() < 8 {
            return Ok(());
        }
        let mut off = 0;
        let row_index = read_u16(data, &mut off)? as u32;
        let _first_col = read_u16(data, &mut off)?;
        let _last_col_plus1 = read_u16(data, &mut off)?;
        let raw_height = read_u16(data, &mut off)?;

        let height_twips = raw_height & 0x7FFF;
        let height_pt = height_twips as f64 / 20.0;

        if data.len() >= 16 {
            let mut opt_off = 12;
            let options = read_u32(data, &mut opt_off).unwrap_or(0);
            let hidden = (options & 0x20) != 0;
            let custom_height = (options & 0x40) != 0;

            if hidden {
                ws.set_row_hidden(row_index, true);
            }
            if custom_height && height_pt > 0.0 {
                ws.set_row_height(row_index, height_pt);
            }
        }

        Ok(())
    }

    // ── Formula context record parsers ──────────────────────────────────

    /// Parse a SUPBOOK record.
    ///
    /// Self-reference: cch == 0x0401. Add-in: ctab == 1, cch == 0x003A.
    /// External: virtPath + sheet names.
    fn parse_supbook(data: &[u8]) -> XlsResult<SupBook> {
        if data.len() < 4 {
            return Err(XlsError::Parse("SUPBOOK record too short".into()));
        }
        let mut off = 0;
        let ctab = read_u16(data, &mut off)?;
        let cch = read_u16(data, &mut off)?;

        if cch == 0x0401 {
            // Self-reference: this workbook
            return Ok(SupBook::SelfRef { sheet_count: ctab });
        }

        if ctab == 1 && cch == 0x003A {
            // Add-in functions sentinel
            return Ok(SupBook::AddIn);
        }

        // External workbook reference: read encoded path + sheet names.
        // Path is encoded as XLUnicodeStringNoCch: flags(1) + chars(cch).
        let path = if off < data.len() && cch > 0 {
            let flags = data[off];
            off += 1;
            // The first byte of the path may be 0x01 (encoded virtPath)
            // We read cch chars and strip the leading 0x01 if present.
            let raw = read_character_data(data, &mut off, cch, flags)?;
            // Strip leading encoding byte if present
            if raw.starts_with('\x01') {
                raw[1..].to_string()
            } else {
                raw
            }
        } else {
            String::new()
        };

        // Read ctab sheet name strings
        let mut sheets = Vec::with_capacity(ctab as usize);
        for _ in 0..ctab {
            if off >= data.len() {
                break;
            }
            let s = read_unicode_string(data, &mut off)?;
            sheets.push(s);
        }

        Ok(SupBook::External { path, sheets })
    }

    /// Parse an EXTERNSHEET record.
    ///
    /// Format: cXTI(u16) + cXTI × (iSupBook(u16), itabFirst(u16), itabLast(u16))
    fn parse_externsheet(data: &[u8]) -> XlsResult<Vec<ExternSheetEntry>> {
        if data.len() < 2 {
            return Err(XlsError::Parse("EXTERNSHEET record too short".into()));
        }
        let mut off = 0;
        let count = read_u16(data, &mut off)? as usize;
        let mut entries = Vec::with_capacity(count);

        for _ in 0..count {
            if off + 6 > data.len() {
                break;
            }
            let sup_book_idx = read_u16(data, &mut off)?;
            let first_sheet = read_u16(data, &mut off)?;
            let last_sheet = read_u16(data, &mut off)?;
            entries.push(ExternSheetEntry {
                sup_book_idx,
                first_sheet,
                last_sheet,
            });
        }

        Ok(entries)
    }

    /// Parse a NAME (defined name / Lbl) record.
    ///
    /// 14-byte header: flags(2) + chKey(1) + cch(1) + cce(2) + reserved(2) +
    /// itab(2) + reserved(4). Then name string (XLUnicodeStringNoCch format:
    /// flags byte + character data). Built-in names have cch == 1 and the
    /// "character" is an index into BUILTIN_NAMES.
    fn parse_name(data: &[u8]) -> XlsResult<NameRecord> {
        if data.len() < 14 {
            return Err(XlsError::Parse("NAME record too short".into()));
        }
        let mut off = 0;
        let flags = read_u16(data, &mut off)?;
        let _ch_key = data[off];
        off += 1;
        let cch = data[off] as u16;
        off += 1;
        let _cce = read_u16(data, &mut off)?;
        let _reserved1 = read_u16(data, &mut off)?;
        let itab = read_u16(data, &mut off)?;
        off += 4; // reserved 4 bytes
                  // off is now 14

        let is_builtin = (flags & 0x0020) != 0;

        // Read name string: XLUnicodeStringNoCch format = flags_byte + chars
        let name = if off < data.len() && cch > 0 {
            let name_flags = data[off];
            off += 1;

            if is_builtin && cch == 1 {
                // Built-in name: the single "character" is an index byte
                let idx = data.get(off).copied().unwrap_or(0) as usize;
                let _skip = if (name_flags & 0x01) != 0 { 2 } else { 1 };
                if idx < BUILTIN_NAMES.len() {
                    BUILTIN_NAMES[idx].to_string()
                } else {
                    format!("_builtin{}", idx)
                }
            } else {
                read_character_data(data, &mut off, cch, name_flags)?
            }
        } else {
            String::new()
        };

        Ok(NameRecord {
            name,
            sheet_idx: itab,
            is_builtin,
        })
    }

    /// COLINFO: first_col(2) + last_col(2) + width(2) + xf(2) + options(2) + reserved(2)
    fn parse_colinfo(data: &[u8], ws: &mut duke_sheets_core::Worksheet) -> XlsResult<()> {
        if data.len() < 10 {
            return Ok(());
        }
        let mut off = 0;
        let first_col = read_u16(data, &mut off)?;
        let last_col = read_u16(data, &mut off)?;
        let raw_width = read_u16(data, &mut off)?;
        let _xf = read_u16(data, &mut off)?;
        let options = read_u16(data, &mut off)?;

        let hidden = (options & 0x0001) != 0;
        let width_chars = raw_width as f64 / 256.0;

        for col in first_col..=last_col {
            if hidden {
                ws.set_column_hidden(col, true);
            }
            if width_chars > 0.0 {
                ws.set_column_width(col, width_chars);
            }
        }

        Ok(())
    }
}
