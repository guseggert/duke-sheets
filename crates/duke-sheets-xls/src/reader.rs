//! XLS (BIFF8) reader.
//!
//! Opens a Compound File Binary (CFB/OLE2) container, reads the `Workbook`
//! stream, parses BIFF8 records, and populates a `duke_sheets_core::Workbook`.

use std::io::{Cursor, Read, Seek};
use std::path::Path;

use duke_sheets_core::cell::SharedString;
use duke_sheets_core::{CellError, CellValue, Style, Workbook};

use crate::biff::parser::{read_f64, read_rk, read_u16, read_u32};
use crate::biff::records;
use crate::biff::strings::{parse_sst, read_short_string, read_unicode_string};
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
    #[allow(dead_code)]
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
                    sst = parse_sst(&rec.data)?;
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

            // Get this sheet's records (indexed by BIFF order, not wb order)
            if let Some(sheet_records) = sheet_record_groups.get(biff_idx) {
                Self::parse_sheet_records(sheet_records, ws, &sst, &style_table)?;
            }

            wb_sheet_idx += 1;
        }

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
    ) -> XlsResult<()> {
        // We need to track the last FORMULA record to associate a STRING record
        let mut pending_formula_cell: Option<(u32, u16)> = None;

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
                    pending_formula_cell = Self::parse_formula(&rec.data, ws, styles)?;
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
                _ => {
                    // Skip unknown/unhandled records
                }
            }
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

    /// FORMULA: row(2) + col(2) + xf(2) + result(8) + options(2) + reserved(4) + formula_data(...)
    ///
    /// Returns the (row, col) if the cached result is a string (meaning a
    /// STRING record should follow).
    fn parse_formula(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
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
                            text: String::new(),
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
                            text: String::new(),
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
                            text: String::new(),
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
                            text: String::new(),
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
                    text: String::new(),
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
    fn parse_formula_string(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        row: u32,
        col: u16,
    ) -> XlsResult<()> {
        let mut off = 0;
        let text = read_unicode_string(data, &mut off)?;

        ws.set_cell_value_at(
            row,
            col,
            CellValue::Formula {
                text: String::new(),
                cached_value: Some(Box::new(CellValue::String(SharedString::new(&text)))),
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
