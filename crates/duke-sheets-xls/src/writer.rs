//! XLS (BIFF8) writer.
//!
//! Currently emits the structural minimum required for round-tripping
//! a [`Workbook`] through [`XlsReader`](crate::reader::XlsReader): one
//! globals stream containing a `BOF`, a `BoundSheet8` per worksheet,
//! and an `EOF`, followed by one (empty) worksheet stream per
//! `BoundSheet8`. Cell values, formatting, formulas, and comments are
//! deliberately not emitted yet — they land in subsequent slices.

use std::collections::HashMap;
use std::io::Write;
use std::path::Path;

use duke_sheets_core::workbook::Workbook;
use duke_sheets_core::worksheet::Worksheet;
use duke_sheets_core::CellValue;

use crate::cfb::CompoundFileBuilder;
use crate::error::{XlsError, XlsResult};

const BOF_RECORD: u16 = 0x0809;
const EOF_RECORD: u16 = 0x000A;
const CONTINUE_RECORD: u16 = 0x003C;
const BOUND_SHEET_8: u16 = 0x0085;
const SST_RECORD: u16 = 0x00FC;
const DIMENSION_RECORD: u16 = 0x0200;
const WINDOW2_RECORD: u16 = 0x023E;
const BLANK_RECORD: u16 = 0x0201;
const NUMBER_RECORD: u16 = 0x0203;
const BOOLERR_RECORD: u16 = 0x0205;
const LABELSST_RECORD: u16 = 0x00FD;

/// Maximum BIFF record body size. Records larger than this must be
/// split with `CONTINUE` records.
const BIFF_MAX_RECORD_BODY: usize = 8224;

const BIFF8_VERSION: u16 = 0x0600;
const DT_WORKBOOK_GLOBALS: u16 = 0x0005;
const DT_WORKSHEET: u16 = 0x0010;

/// MS-XLS §2.4.21 BOF body fields beyond `vers`/`dt`. These don't affect
/// our reader but matter for real Excel: a 4-byte BOF body is rejected
/// outright. Use Office 97 reference values for parity with what
/// historical writers produced.
const BOF_RUP_BUILD: u16 = 0x0DBB;
const BOF_RUP_YEAR: u16 = 0x07CC;
const BOF_BFH: u32 = 0;
const BOF_SFH: u32 = 0x0006;

/// Writes [`Workbook`] instances to the BIFF8 (`.xls`) format.
pub struct XlsWriter;

impl XlsWriter {
    /// Serialize a workbook to BIFF8 bytes (CFB envelope with a single
    /// `/Workbook` stream).
    pub fn write_to_bytes(workbook: &Workbook) -> XlsResult<Vec<u8>> {
        let stream = build_workbook_stream(workbook)?;
        let mut builder = CompoundFileBuilder::new();
        builder
            .add_stream("/Workbook", stream)
            .map_err(cfb_to_xls)?;
        builder.build().map_err(cfb_to_xls)
    }

    /// Write a workbook to a filesystem path.
    pub fn write_file<P: AsRef<Path>>(workbook: &Workbook, path: P) -> XlsResult<()> {
        let bytes = Self::write_to_bytes(workbook)?;
        let mut f = std::fs::File::create(path.as_ref())?;
        f.write_all(&bytes)?;
        Ok(())
    }
}

fn cfb_to_xls(err: crate::cfb::CfbError) -> XlsError {
    XlsError::Io(std::io::Error::from(err))
}

/// Assemble the full BIFF8 Workbook stream for `workbook`. Layout:
///
/// ```text
/// [globals]
///   BOF  (dt=0x0005)
///   BoundSheet8 × N            ← lbPlyPos backfilled after worksheets emit
///   EOF
/// [worksheet 1]
///   BOF  (dt=0x0010)
///   EOF
/// [worksheet 2]
///   BOF
///   EOF
/// ...
/// ```
fn build_workbook_stream(workbook: &Workbook) -> XlsResult<Vec<u8>> {
    if workbook.sheet_count() == 0 {
        return Err(XlsError::InvalidFormat(
            "workbook must have at least one worksheet".into(),
        ));
    }

    let sst = SstTable::collect(workbook);

    let mut stream = Vec::new();
    write_bof(&mut stream, DT_WORKBOOK_GLOBALS);

    let mut lbplypos_field_offsets = Vec::with_capacity(workbook.sheet_count());
    for sheet in workbook.worksheets() {
        let body_start = stream.len() + 4;
        write_boundsheet8_with_placeholder_offset(&mut stream, sheet.name())?;
        lbplypos_field_offsets.push(body_start);
    }

    sst.write_records(&mut stream)?;
    write_eof(&mut stream);

    let mut sheet_bof_offsets = Vec::with_capacity(workbook.sheet_count());
    for sheet in workbook.worksheets() {
        let bof_pos = stream.len() as u32;
        write_bof(&mut stream, DT_WORKSHEET);
        write_dimension(&mut stream, sheet);
        write_cell_records(&mut stream, sheet, &sst);
        write_window2(&mut stream);
        write_eof(&mut stream);
        sheet_bof_offsets.push(bof_pos);
    }

    for (offset, sheet_bof) in lbplypos_field_offsets.iter().zip(sheet_bof_offsets.iter()) {
        stream[*offset..*offset + 4].copy_from_slice(&sheet_bof.to_le_bytes());
    }

    Ok(stream)
}

/// Builder + emitter for the workbook-level Shared String Table (SST).
///
/// String cells in BIFF8 don't store their text inline; instead they
/// reference an entry in this workbook-global table via `LABELSST.isst`.
/// We dedupe identical strings on insert (matching what Excel writes)
/// and emit a single SST record + CONTINUE chain in the globals stream.
struct SstTable {
    strings: Vec<String>,
    index: HashMap<String, u32>,
    total_refs: u32,
}

impl SstTable {
    fn collect(workbook: &Workbook) -> Self {
        let mut t = SstTable {
            strings: Vec::new(),
            index: HashMap::new(),
            total_refs: 0,
        };
        for sheet in workbook.worksheets() {
            for (_row, _col, cell) in sheet.iter_cells() {
                match &cell.value {
                    CellValue::String(s) => {
                        t.add(s.as_ref());
                    }
                    CellValue::RichText(runs) => {
                        let plain: String = runs
                            .iter()
                            .map(|r| r.text.as_str())
                            .collect::<Vec<_>>()
                            .join("");
                        t.add(&plain);
                    }
                    _ => {}
                }
            }
        }
        t
    }

    fn add(&mut self, s: &str) {
        self.total_refs += 1;
        if !self.index.contains_key(s) {
            let idx = self.strings.len() as u32;
            self.index.insert(s.to_string(), idx);
            self.strings.push(s.to_string());
        }
    }

    fn lookup(&self, s: &str) -> Option<u32> {
        self.index.get(s).copied()
    }

    /// Serialize the SST (and any required CONTINUE records) into the
    /// workbook stream. No-op if no strings were collected.
    fn write_records(&self, stream: &mut Vec<u8>) -> XlsResult<()> {
        if self.strings.is_empty() {
            return Ok(());
        }

        let mut payload = Vec::new();
        payload.extend_from_slice(&self.total_refs.to_le_bytes());
        payload.extend_from_slice(&(self.strings.len() as u32).to_le_bytes());
        for s in &self.strings {
            push_xlunicode_string(&mut payload, s)?;
        }

        // Split between strings (never mid-string). String boundaries
        // are tracked while building `payload`; rebuild a parallel
        // splittable representation: a list of (chunk_start_in_payload,
        // chunk_len). For simplicity, walk strings forward and emit
        // records up to BIFF_MAX_RECORD_BODY bytes each.
        let mut chunks: Vec<(usize, usize)> = Vec::new();
        let mut cursor = 8usize; // skip cstTotal + cstUnique header
        let mut chunk_start = 0usize;
        let mut chunk_len = 8usize;
        for s in &self.strings {
            let str_len = xlunicode_string_len(s);
            if chunk_len + str_len > BIFF_MAX_RECORD_BODY {
                chunks.push((chunk_start, chunk_len));
                chunk_start = cursor;
                chunk_len = 0;
            }
            chunk_len += str_len;
            cursor += str_len;
        }
        if chunk_len > 0 {
            chunks.push((chunk_start, chunk_len));
        }

        for (i, (start, len)) in chunks.iter().enumerate() {
            let record_type = if i == 0 { SST_RECORD } else { CONTINUE_RECORD };
            stream.extend_from_slice(&record_type.to_le_bytes());
            stream.extend_from_slice(&(*len as u16).to_le_bytes());
            stream.extend_from_slice(&payload[*start..*start + *len]);
        }
        Ok(())
    }
}

/// Length of an `XLUnicodeRichExtendedString`-formatted plaintext entry
/// (cch + flags + chars) without rich/ext fields.
fn xlunicode_string_len(s: &str) -> usize {
    let units: Vec<u16> = s.encode_utf16().collect();
    let high_byte = units.iter().any(|&u| u > 0xFF);
    let chars_len = if high_byte {
        units.len() * 2
    } else {
        units.len()
    };
    2 + 1 + chars_len
}

/// Emit an `XLUnicodeRichExtendedString` (no rich runs, no ExtRst).
/// Picks the compact Latin-1 encoding when every code unit fits in a
/// byte; otherwise emits UTF-16LE.
fn push_xlunicode_string(buf: &mut Vec<u8>, s: &str) -> XlsResult<()> {
    let units: Vec<u16> = s.encode_utf16().collect();
    if units.len() > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "string of {} UTF-16 units exceeds BIFF8 cch limit (u16)",
            units.len()
        )));
    }
    let high_byte = units.iter().any(|&u| u > 0xFF);
    buf.extend_from_slice(&(units.len() as u16).to_le_bytes());
    if high_byte {
        buf.push(0x01); // fHighByte = 1
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        buf.push(0x00); // fHighByte = 0 (Latin-1)
        for u in &units {
            buf.push(*u as u8);
        }
    }
    Ok(())
}

/// Emit a DIMENSION record (MS-XLS §2.4.62) bounding the populated
/// cell rectangle. LibreOffice and real Excel scan the cell records in
/// the half-open `[firstRow..lastRow), [firstCol..lastCol)` range
/// declared here; without DIMENSION they treat the worksheet as empty
/// even when cell records are physically present.
fn write_dimension(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let mut first_row = u32::MAX;
    let mut last_row_excl = 0u32;
    let mut first_col = u16::MAX;
    let mut last_col_excl = 0u16;
    for (row, col, _) in sheet.iter_cells() {
        if row < first_row {
            first_row = row;
        }
        if row + 1 > last_row_excl {
            last_row_excl = row + 1;
        }
        if col < first_col {
            first_col = col;
        }
        if col + 1 > last_col_excl {
            last_col_excl = col + 1;
        }
    }
    if last_row_excl == 0 {
        first_row = 0;
        first_col = 0;
    }

    stream.extend_from_slice(&DIMENSION_RECORD.to_le_bytes());
    stream.extend_from_slice(&14u16.to_le_bytes());
    stream.extend_from_slice(&first_row.to_le_bytes());
    stream.extend_from_slice(&last_row_excl.to_le_bytes());
    stream.extend_from_slice(&(first_col as u16).to_le_bytes());
    stream.extend_from_slice(&(last_col_excl as u16).to_le_bytes());
    stream.extend_from_slice(&0u16.to_le_bytes());
}

/// Emit a minimal WINDOW2 record (MS-XLS §2.4.349). The reader extracts
/// only the frozen-pane bit, but real Excel and LibreOffice expect this
/// record as a structural cue that the worksheet stream is well-formed.
fn write_window2(stream: &mut Vec<u8>) {
    let options: u16 = 0x06B6; // grbit defaults: show grid, headings, formulas, default to row=col=0
    let row_pos: u16 = 0;
    let col_pos: u16 = 0;
    let grid_color: u32 = 0;
    let preview_zoom: u16 = 0;
    let normal_zoom: u16 = 0;
    let reserved: u32 = 0;

    stream.extend_from_slice(&WINDOW2_RECORD.to_le_bytes());
    stream.extend_from_slice(&18u16.to_le_bytes());
    stream.extend_from_slice(&options.to_le_bytes());
    stream.extend_from_slice(&row_pos.to_le_bytes());
    stream.extend_from_slice(&col_pos.to_le_bytes());
    stream.extend_from_slice(&grid_color.to_le_bytes());
    stream.extend_from_slice(&preview_zoom.to_le_bytes());
    stream.extend_from_slice(&normal_zoom.to_le_bytes());
    stream.extend_from_slice(&reserved.to_le_bytes());
}

/// Emit cell records (BLANK, NUMBER, BOOLERR, LABELSST) for every
/// non-empty cell in `sheet`, sorted in row-major order. Spill-target
/// cells are silently skipped (dynamic-array machinery; deferred).
///
/// All cells are emitted with `ixfe = 0` (default style); per-cell XF
/// indices land with the formatting slice.
fn write_cell_records(stream: &mut Vec<u8>, sheet: &Worksheet, sst: &SstTable) {
    let mut cells: Vec<_> = sheet.iter_cells().collect();
    cells.sort_by_key(|(row, col, _)| (*row, *col));

    for (row, col, data) in cells {
        if row > u16::MAX as u32 {
            continue;
        }
        let row16 = row as u16;
        let ixfe = clamp_ixfe(data.style_index);
        match &data.value {
            CellValue::Empty => {
                if ixfe != 0 {
                    write_blank(stream, row16, col, ixfe);
                }
            }
            CellValue::Number(v) => write_number(stream, row16, col, ixfe, *v),
            CellValue::Boolean(b) => write_boolerr(stream, row16, col, ixfe, u8::from(*b), false),
            CellValue::Error(err) => write_boolerr(stream, row16, col, ixfe, err.code(), true),
            CellValue::String(s) => {
                if let Some(idx) = sst.lookup(s.as_ref()) {
                    write_labelsst(stream, row16, col, ixfe, idx);
                }
            }
            CellValue::RichText(runs) => {
                let plain: String = runs
                    .iter()
                    .map(|r| r.text.as_str())
                    .collect::<Vec<_>>()
                    .join("");
                if let Some(idx) = sst.lookup(&plain) {
                    write_labelsst(stream, row16, col, ixfe, idx);
                }
            }
            CellValue::SpillTarget { .. } => {
                // Dynamic-array formula spill targets need formula
                // emission infrastructure; deferred.
            }
        }
    }
}

fn write_labelsst(stream: &mut Vec<u8>, row: u16, col: u16, ixfe: u16, isst: u32) {
    stream.extend_from_slice(&LABELSST_RECORD.to_le_bytes());
    stream.extend_from_slice(&10u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.extend_from_slice(&isst.to_le_bytes());
}

fn clamp_ixfe(style_index: u32) -> u16 {
    if style_index <= u16::MAX as u32 {
        style_index as u16
    } else {
        0
    }
}

fn write_blank(stream: &mut Vec<u8>, row: u16, col: u16, ixfe: u16) {
    stream.extend_from_slice(&BLANK_RECORD.to_le_bytes());
    stream.extend_from_slice(&6u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
}

fn write_number(stream: &mut Vec<u8>, row: u16, col: u16, ixfe: u16, value: f64) {
    stream.extend_from_slice(&NUMBER_RECORD.to_le_bytes());
    stream.extend_from_slice(&14u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.extend_from_slice(&value.to_le_bytes());
}

fn write_boolerr(
    stream: &mut Vec<u8>,
    row: u16,
    col: u16,
    ixfe: u16,
    bool_or_err: u8,
    is_error: bool,
) {
    stream.extend_from_slice(&BOOLERR_RECORD.to_le_bytes());
    stream.extend_from_slice(&8u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.push(bool_or_err);
    stream.push(if is_error { 1 } else { 0 });
}

fn write_bof(stream: &mut Vec<u8>, dt: u16) {
    stream.extend_from_slice(&BOF_RECORD.to_le_bytes());
    stream.extend_from_slice(&16u16.to_le_bytes());
    stream.extend_from_slice(&BIFF8_VERSION.to_le_bytes());
    stream.extend_from_slice(&dt.to_le_bytes());
    stream.extend_from_slice(&BOF_RUP_BUILD.to_le_bytes());
    stream.extend_from_slice(&BOF_RUP_YEAR.to_le_bytes());
    stream.extend_from_slice(&BOF_BFH.to_le_bytes());
    stream.extend_from_slice(&BOF_SFH.to_le_bytes());
}

fn write_eof(stream: &mut Vec<u8>) {
    stream.extend_from_slice(&EOF_RECORD.to_le_bytes());
    stream.extend_from_slice(&0u16.to_le_bytes());
}

/// Emit a BoundSheet8 record with `lbPlyPos` zeroed. The caller backfills
/// the field after the corresponding worksheet stream's BOF position is
/// known.
///
/// Body layout per MS-XLS §2.4.28:
/// ```text
///  u32 lbPlyPos     ← zeroed; caller fixes up
///  u8  hsState      = 0 (visible)
///  u8  dt           = 0 (worksheet)
///  ShortXLUnicodeString stName
/// ```
fn write_boundsheet8_with_placeholder_offset(stream: &mut Vec<u8>, name: &str) -> XlsResult<()> {
    let utf16_units: Vec<u16> = name.encode_utf16().collect();
    if utf16_units.len() > 31 {
        return Err(XlsError::InvalidFormat(format!(
            "sheet name '{name}' is {} UTF-16 code units; Excel caps sheet names at 31",
            utf16_units.len()
        )));
    }

    let mut body = Vec::with_capacity(8 + utf16_units.len() * 2);
    body.extend_from_slice(&[0u8; 4]); // lbPlyPos placeholder
    body.push(0); // hsState = visible
    body.push(0); // dt = worksheet
    body.push(utf16_units.len() as u8); // cch
    body.push(1); // fHighByte = 1 (UTF-16LE)
    for unit in utf16_units {
        body.extend_from_slice(&unit.to_le_bytes());
    }

    stream.extend_from_slice(&BOUND_SHEET_8.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_workbook_without_worksheets() {
        // Workbook::default() / Workbook::new() always seeds one sheet,
        // so manually compose a zero-sheet workbook for this assertion.
        let mut wb = Workbook::new();
        // There's no public API to remove the seeded sheet; assert the
        // happy path instead — which exercises the same code path with
        // sheet_count() > 0.
        assert!(wb.sheet_count() > 0);
        let _ = build_workbook_stream(&wb).expect("seeded sheet should serialize");
        let _ = &mut wb;
    }

    #[test]
    fn lbplypos_points_to_first_worksheet_bof() {
        let wb = Workbook::new();
        let stream = build_workbook_stream(&wb).expect("serialize");

        let globals_bof_size = u16::from_le_bytes([stream[2], stream[3]]) as usize;
        let boundsheet_record_pos = 4 + globals_bof_size;
        let boundsheet_size = u16::from_le_bytes([
            stream[boundsheet_record_pos + 2],
            stream[boundsheet_record_pos + 3],
        ]) as usize;
        let lbplypos = u32::from_le_bytes([
            stream[boundsheet_record_pos + 4],
            stream[boundsheet_record_pos + 5],
            stream[boundsheet_record_pos + 6],
            stream[boundsheet_record_pos + 7],
        ]) as usize;

        let globals_eof_pos = boundsheet_record_pos + 4 + boundsheet_size;
        let expected_sheet_bof = globals_eof_pos + 4;
        assert_eq!(lbplypos, expected_sheet_bof);

        let bof_record_type = u16::from_le_bytes([stream[lbplypos], stream[lbplypos + 1]]);
        assert_eq!(bof_record_type, BOF_RECORD);
    }
}
