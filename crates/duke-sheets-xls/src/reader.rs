//! XLS (BIFF8) reader.
//!
//! Opens a Compound File Binary (CFB/OLE2) container, reads the `Workbook`
//! stream, parses BIFF8 records, and populates a `duke_sheets_core::Workbook`.

use std::io::{Cursor, Read, Seek};
use std::path::Path;

use duke_sheets_core::cell::SharedString;
use duke_sheets_core::conditional_format::{CfOperator, CfRuleType, ConditionalFormatRule};
use duke_sheets_core::rich_text::RichTextRun;
use duke_sheets_core::validation::{
    DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
};
use duke_sheets_core::worksheet::{Selection, SheetProtection};
use duke_sheets_core::{
    CellAddress, CellComment, CellError, CellRange, CellValue, Hyperlink, ProtectedRange, Style,
    Workbook, WorkbookProtection, Worksheet,
};

use crate::biff::formula::token_parser::ParsedToken;
use crate::biff::formula::{
    ExternName, ExternSheetEntry, FormulaContext, NameRecord, SupBook, BUILTIN_NAMES,
};
use crate::biff::parser::{read_f64, read_rk, read_u16, read_u32, read_u8};
use crate::biff::records;
use crate::biff::strings::{
    parse_sst_entries, read_character_data, read_short_string, read_unicode_string, SstEntry,
};
use crate::biff::{self, BiffRecord};
use crate::error::{XlsError, XlsResult};
use crate::styles::{self, StyleContext};

/// XLS file reader.
pub struct XlsReader;

/// One embedded image's raw bytes + format, extracted from a
/// `MSODRAWINGGROUP`'s `BSTORE_CONTAINER`. The 1-based position of
/// each entry in the workbook's blip-store vector is the `pib`
/// (picture blip id) that picture shapes reference via their FOPT
/// `0x0104` property.
#[derive(Debug, Clone)]
struct BlipData {
    format: duke_sheets_chart::ImageFormat,
    data: Vec<u8>,
}

/// One non-patriarch `SP_CONTAINER`'s parsed atoms. Nodes are
/// collected in document (pre-order) order — a group's own shape,
/// then its children, then the group's next sibling — matching the
/// order of the sheet's OBJ records: the Nth ClientData-bearing node
/// pairs with the Nth OBJ record.
#[derive(Debug, Clone, Default)]
struct EscherShapeNode {
    /// MSOSPT shape type from the FSP atom header instance.
    shape_type: u16,
    /// Shape ID from the FSP atom.
    spid: u32,
    /// FSP flip flags.
    flip_h: bool,
    flip_v: bool,
    /// FOPT rotation (0x0004) wire value: FixedPoint 16.16 degrees
    /// ([MS-ODRAW] 2.3.18.5). Model conversion (to 60,000ths of a
    /// degree) happens via `officeart_fixed_to_rotation`.
    rotation: Option<i32>,
    /// FOPT pib (0x0104): 1-based blip-store index.
    blip_id: Option<u32>,
    /// Basic shape fill and line properties from FOPT.
    fill_color: Option<u32>,
    fill_enabled: Option<bool>,
    line_color: Option<u32>,
    line_width: Option<u32>,
    line_dashing: Option<u32>,
    line_no_fill: Option<bool>,
    /// FOPT wzName (0x0380).
    name: Option<String>,
    /// FOPT wzDescription (0x0381).
    alt_text: Option<String>,
    /// FOPT Group Shape Boolean Properties (0x03BF) fHidden, gated
    /// on its use-bit (MS-ODRAW §2.3.4.44).
    hidden: bool,
    /// Sheet anchor (top-level shapes).
    client_anchor: Option<crate::biff::escher::OfficeArtClientAnchor>,
    /// Group-space anchor (grouped shapes).
    child_anchor: Option<crate::biff::escher::OfficeArtChildAnchor>,
    /// Group child coordinate space (group shapes only).
    fspgr: Option<crate::biff::escher::OfficeArtFspgr>,
    /// Whether the SP container carries a ClientData marker (and so
    /// pairs with an OBJ record).
    has_client_data: bool,
    /// Whether this node is a shape group (an `SpgrContainer`).
    is_group: bool,
    /// Group children in document order (groups only).
    children: Vec<EscherShapeNode>,
}

/// One NOTE record's fields, collected during the record loop and
/// resolved against comment shapes after the sheet's drawing stream
/// is assembled.
#[derive(Debug, Clone)]
struct NoteData {
    row: u32,
    col: u16,
    visible: bool,
    obj_id: u16,
    author: String,
}

fn officeart_color_to_core(value: u32) -> Option<duke_sheets_core::Color> {
    // [MS-ODRAW] 2.2.2: OfficeArtCOLORREF carries RGB in the low three
    // bytes. Scheme/system-index colors need external context.
    let flags = value >> 24;
    // fSystemRGB/fPaletteRGB still carry usable RGB channels. System,
    // scheme, and palette-index references require external context.
    if flags & 0x19 != 0 {
        return None;
    }
    Some(duke_sheets_core::Color::rgb(
        value as u8,
        (value >> 8) as u8,
        (value >> 16) as u8,
    ))
}

fn officeart_fixed_to_rotation(value: i32) -> i32 {
    let numerator = i64::from(value) * 60_000;
    let rotation = if numerator >= 0 {
        (numerator + 32_768) / 65_536
    } else {
        (numerator - 32_768) / 65_536
    };
    rotation.clamp(i64::from(i32::MIN), i64::from(i32::MAX)) as i32
}

fn officeart_dash_to_drawing(value: u32) -> String {
    match value {
        1 => "sysDash",
        2 => "sysDot",
        3 => "sysDashDot",
        4 => "sysDashDotDot",
        5 => "dot",
        6 => "dash",
        7 => "lgDash",
        8 => "dashDot",
        9 => "lgDashDot",
        10 => "lgDashDotDot",
        _ => "solid",
    }
    .to_string()
}

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

/// Result from parsing a FORMULA record.
enum FormulaResult {
    /// Normal formula parsed. Contains string-pending (row, col) if the
    /// cached result is a string (meaning a STRING record should follow).
    Done(Option<(u32, u16)>),
    /// This cell uses a shared formula whose SHAREDFMLA record hasn't been
    /// seen yet. The cell is written with empty formula text; the caller
    /// must backfill the text when the SHAREDFMLA record arrives.
    SharedPending {
        cell_row: u32,
        cell_col: u16,
        master_row: u16,
        master_col: u16,
    },
    TablePending {
        cell_row: u32,
        cell_col: u16,
        master_row: u16,
        master_col: u16,
    },
}

enum DoperValue {
    None,
    Float(f64),
    StrInfo { cch: u8 },
    Bool(bool),
    Error(u8),
    Blanks,
    NonBlanks,
}

/// Peek at a raw Workbook stream to tell whether it has a FilePass
/// record immediately after the globals BOF. Cheaper than running the
/// full BIFF parser; used to short-circuit the decryption path.
fn is_encrypted_workbook_stream(stream: &[u8]) -> bool {
    let mut cursor = 0usize;
    let mut seen_bof = false;
    while cursor + 4 <= stream.len() {
        let record_type = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let size = u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]) as usize;
        let body_end = cursor + 4 + size;
        if body_end > stream.len() {
            return false;
        }
        if record_type == 0x0809 {
            // BOF
            if seen_bof {
                // Second BOF means we've moved past globals; no FilePass.
                return false;
            }
            seen_bof = true;
        } else if record_type == 0x002F {
            return true;
        } else if seen_bof {
            // Any non-BOF non-FilePass record means the globals block
            // started without FilePass; it's plaintext.
            return false;
        }
        cursor = body_end;
    }
    false
}

/// Pick the BIFF workbook stream path inside a CFB container, or return a
/// clean error describing why no readable workbook is present.
///
/// Encrypted OOXML files (password-protected `.xlsx`/`.xlsm`) are stored as CFB
/// containers with the streams `EncryptionInfo` and `EncryptedPackage`, so
/// they're misclassified as XLS by the top-level format sniffer. Detecting
/// those streams here lets us return a specific `Encrypted` error rather than
/// a misleading "no Workbook or Book stream found in CFB".
pub(crate) fn resolve_workbook_stream<F: Fn(&str) -> bool>(exists: F) -> XlsResult<&'static str> {
    if exists("/Workbook") {
        Ok("/Workbook")
    } else if exists("/Book") {
        Ok("/Book")
    } else if exists("/EncryptedPackage") || exists("/EncryptionInfo") {
        Err(XlsError::Encrypted(
            "file is an encrypted OOXML document; decryption is not supported".into(),
        ))
    } else {
        Err(XlsError::InvalidFormat(
            "no Workbook or Book stream found in CFB".into(),
        ))
    }
}

impl XlsReader {
    /// Read an XLS file from a filesystem path.
    pub fn read_file<P: AsRef<Path>>(path: P) -> XlsResult<Workbook> {
        let file = std::fs::File::open(path.as_ref())?;
        Self::read(file)
    }

    /// Read an XLS file from a filesystem path, supplying a password for
    /// encrypted workbooks. When `password` is `None` and
    /// `try_velvet_sweatshop` is true, encrypted workbooks are
    /// transparently retried with the `VelvetSweatshop` sentinel before
    /// reporting them as encrypted.
    pub fn read_file_with_password<P: AsRef<Path>>(
        path: P,
        password: Option<&str>,
        try_velvet_sweatshop: bool,
    ) -> XlsResult<Workbook> {
        let file = std::fs::File::open(path.as_ref())?;
        Self::read_with_password(file, password, try_velvet_sweatshop)
    }

    /// Read an XLS file from any `Read + Seek` source.
    pub fn read<R: Read + Seek>(reader: R) -> XlsResult<Workbook> {
        Self::read_with_password(reader, None, false)
    }

    /// Read an XLS file from any `Read + Seek` source, supplying a
    /// password for encrypted workbooks.
    ///
    /// `try_velvet_sweatshop` enables the Excel-compatible auto-retry
    /// with the well-known sentinel password when no explicit password
    /// is supplied. Wrong passwords return [`XlsError::BadPassword`].
    pub fn read_with_password<R: Read + Seek>(
        reader: R,
        password: Option<&str>,
        try_velvet_sweatshop: bool,
    ) -> XlsResult<Workbook> {
        let cfb = crate::cfb::CompoundFile::open(reader).map_err(std::io::Error::from)?;
        let stream_path = resolve_workbook_stream(|p| cfb.exists(p))?;
        let mut stream_data = cfb.read_stream(stream_path).map_err(std::io::Error::from)?;

        if is_encrypted_workbook_stream(&stream_data) {
            let try_pw = match password {
                Some(p) => Some(p),
                None if try_velvet_sweatshop => Some("VelvetSweatshop"),
                None => None,
            };
            if let Some(pw) = try_pw {
                match duke_sheets_crypto::xls::decrypt_workbook_stream(&stream_data, pw) {
                    Ok(decrypted) => {
                        stream_data = decrypted;
                    }
                    Err(duke_sheets_crypto::CryptoError::BadPassword) if password.is_none() => {
                        return Err(XlsError::Encrypted(
                            "workbook is encrypted but no password was supplied".into(),
                        ));
                    }
                    Err(e) => return Err(e.into()),
                }
            }
        }

        // Parse all BIFF records from the stream
        let mut cursor = Cursor::new(&stream_data);
        let all_records = biff::read_all_records(&mut cursor)?;

        // Bail out immediately on encrypted workbooks rather than feeding
        // ciphertext through the BIFF record parsers.
        biff::check_not_encrypted(&all_records)?;

        // Phase 1: Parse workbook globals
        let mut sst: Vec<SstEntry> = Vec::new();
        let mut sheets: Vec<SheetInfo> = Vec::new();
        let mut date_mode_1904 = false;
        let mut in_globals = false;
        let mut style_ctx = StyleContext::new();
        let mut active_sheet_idx: u16 = 0;
        let mut workbook_protected = false;
        let mut workbook_windows_protected = false;
        let mut workbook_password_hash: Option<u16> = None;
        let mut supbooks: Vec<SupBook> = Vec::new();
        let mut extern_sheet: Vec<ExternSheetEntry> = Vec::new();
        let mut extern_names: Vec<ExternName> = Vec::new();
        let mut names: Vec<NameRecord> = Vec::new();
        // Workbook-globals blip store, populated from MSODRAWINGGROUP
        // records. Indexed 1-based by the FOPT `pib` (picture blip id)
        // property referenced from picture SP_CONTAINERs.
        let mut blip_store: Vec<Option<BlipData>> = Vec::new();
        // Excel splits a large drawing group across multiple
        // MSODRAWINGGROUP records, each holding a fragment of one
        // logical DggContainer stream, so bodies are concatenated
        // here and walked once after the globals loop.
        let mut msodrawinggroup_bytes: Vec<u8> = Vec::new();

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
                    sst = parse_sst_entries(&rec.data, &rec.continue_offsets)?;
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
                records::WINDOWPROTECT if in_globals => {
                    if rec.data.len() >= 2 {
                        let val = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        workbook_windows_protected = val == 1;
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
                records::EXTERNNAME if in_globals => {
                    // EXTERNNAME records follow the SUPBOOK they belong to, so
                    // associate each with the most recently seen SUPBOOK.
                    if !supbooks.is_empty() {
                        if let Ok(name) = Self::parse_externname(&rec.data) {
                            extern_names.push(ExternName {
                                supbook_idx: (supbooks.len() - 1) as u16,
                                name,
                            });
                        }
                    }
                }
                records::NAME if in_globals => {
                    if let Ok(nr) = Self::parse_name(&rec.data) {
                        names.push(nr);
                    }
                }
                records::MSODRAWINGGROUP if in_globals => {
                    msodrawinggroup_bytes.extend_from_slice(&rec.data);
                }
                _ => {}
            }
        }

        if !msodrawinggroup_bytes.is_empty() {
            Self::parse_msodrawinggroup(&msodrawinggroup_bytes, &mut blip_store);
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
            extern_names,
            extern_name_index_base: 1,
            base_cell: None,
        };

        let mut filter_db_ranges: std::collections::HashMap<usize, CellRange> =
            std::collections::HashMap::new();
        for name_rec in &formula_ctx.names {
            if name_rec.name == "_FilterDatabase" && !name_rec.formula_body.is_empty() {
                if let Some(range) =
                    Self::extract_filter_db_range(&name_rec.formula_body, &formula_ctx)
                {
                    if name_rec.sheet_idx != 0xFFFFFFFF {
                        filter_db_ranges.insert(name_rec.sheet_idx as usize, range);
                    }
                }
            }
        }

        // Extract Print_Area (builtin name 0x06) ranges per sheet
        let mut print_area_ranges: std::collections::HashMap<usize, CellRange> =
            std::collections::HashMap::new();
        // Extract Print_Titles (builtin name 0x07) per sheet → (repeat_rows, repeat_cols)
        let mut print_titles: std::collections::HashMap<
            usize,
            (Option<(u32, u32)>, Option<(u16, u16)>),
        > = std::collections::HashMap::new();
        for name_rec in &formula_ctx.names {
            if name_rec.formula_body.is_empty() {
                continue;
            }
            if name_rec.sheet_idx != 0xFFFFFFFF {
                let sheet = name_rec.sheet_idx as usize;
                if name_rec.name == "Print_Area" {
                    if let Some(range) =
                        Self::extract_filter_db_range(&name_rec.formula_body, &formula_ctx)
                    {
                        print_area_ranges.insert(sheet, range);
                    }
                } else if name_rec.name == "Print_Titles" {
                    let titles = Self::extract_print_titles(&name_rec.formula_body);
                    if titles.0.is_some() || titles.1.is_some() {
                        print_titles.insert(sheet, titles);
                    }
                }
            }
        }

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
            workbook.add_worksheet_with_name_unchecked(&info.name);

            let ws = workbook.worksheet_mut(wb_sheet_idx).unwrap();
            ws.set_date_1904(date_mode_1904);

            // Apply sheet visibility (0 = visible, 1 = hidden, 2 = very hidden)
            let visibility = match info.visibility {
                1 => duke_sheets_core::SheetVisibility::Hidden,
                2 => duke_sheets_core::SheetVisibility::VeryHidden,
                _ => duke_sheets_core::SheetVisibility::Visible,
            };
            ws.set_visibility(visibility);

            // Get this sheet's records (indexed by BIFF order, not wb order)
            if let Some(sheet_records) = sheet_record_groups.get(biff_idx) {
                let af_range = filter_db_ranges.get(&biff_idx);
                Self::parse_sheet_records(
                    sheet_records,
                    ws,
                    &sst,
                    &style_table,
                    &style_ctx,
                    &formula_ctx,
                    af_range,
                    &blip_store,
                )?;
            }

            // Apply Print_Area from NAME formula body
            if let Some(range) = print_area_ranges.get(&biff_idx) {
                ws.set_print_area(range.clone());
            }

            // Apply Print_Titles (repeat rows/cols) from NAME formula body
            if let Some((rows, cols)) = print_titles.get(&biff_idx) {
                if let Some((r1, r2)) = rows {
                    ws.set_repeat_rows(*r1, *r2);
                }
                if let Some((c1, c2)) = cols {
                    ws.set_repeat_cols(*c1, *c2);
                }
            }

            wb_sheet_idx += 1;
        }

        // Apply active sheet index (WINDOW1)
        let active = active_sheet_idx as usize;
        if active < workbook.sheet_count() {
            let _ = workbook.set_active_sheet(active);
        }

        if workbook_protected || workbook_windows_protected || workbook_password_hash.is_some() {
            workbook.set_workbook_protection(Some(WorkbookProtection {
                structure: workbook_protected,
                windows: workbook_windows_protected,
                password_hash: workbook_password_hash,
            }));
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
        sst: &[SstEntry],
        styles: &[Style],
        style_ctx: &StyleContext,
        formula_ctx: &FormulaContext,
        auto_filter_range: Option<&CellRange>,
        blip_store: &[Option<BlipData>],
    ) -> XlsResult<()> {
        // We need to track the last FORMULA record to associate a STRING record
        let mut pending_formula_cell: Option<(u32, u16)> = None;
        let mut sheet_protected = false;
        let mut sheet_password_hash: Option<u16> = None;
        let mut window2_frozen = false;

        // Shared formula support: stores (master_row, master_col) → token bytes
        let mut shared_formulas: std::collections::HashMap<(u16, u16), Vec<u8>> =
            std::collections::HashMap::new();
        // Array formula support: stores (top_left_row, top_left_col) → (token_data, extra_data)
        let mut array_formulas: std::collections::HashMap<(u16, u16), (Vec<u8>, Vec<u8>)> =
            std::collections::HashMap::new();
        // Data table support: stores (master_row, master_col) → (input1_ref, input2_ref)
        let mut data_tables: std::collections::HashMap<(u16, u16), (String, String)> =
            std::collections::HashMap::new();
        // Master cell that appeared before its SHAREDFMLA or ARRAY record.
        // (cell_row, cell_col, master_row, master_col)
        let mut pending_shared: Option<(u32, u16, u16, u16)> = None;
        // Cells with PTG_TBL whose TABLE record hasn't been seen yet.
        let mut pending_table_cells: Vec<(u32, u16, u16, u16)> = Vec::new();

        // Comment support: OBJ → TXO → NOTE correlation
        let mut last_obj_id: Option<u16> = None;
        let mut obj_texts: std::collections::HashMap<u16, duke_sheets_core::ControlText> =
            std::collections::HashMap::new();
        // NOTE records, resolved against comment shapes after the loop.
        let mut notes: Vec<NoteData> = Vec::new();

        // Conditional formatting: CONDFMT range header for following CF records
        let mut cf_ranges: Vec<CellRange> = Vec::new();

        let mut auto_filter_columns: Vec<duke_sheets_core::FilterColumn> = Vec::new();

        // Hyperlink tooltip: HLINKTOOLTIP records keyed by (row, col)
        let mut hlink_tooltips: std::collections::HashMap<(u32, u16), String> =
            std::collections::HashMap::new();

        // Per-sheet Escher byte stream: all MSODRAWING record bodies
        // concatenated in BIFF order. The picture SP_CONTAINERs and
        // ClientTextbox markers live here, possibly split across
        // multiple records.
        let mut escher_bytes: Vec<u8> = Vec::new();
        // Full OBJ record bodies in BIFF order. Used after the record
        // loop to link OBJ entries to their SP_CONTAINERs by position
        // (the OBJ↔shape association in BIFF8 is purely positional).
        let mut obj_bodies: Vec<Vec<u8>> = Vec::new();

        for rec in records {
            match rec.record_type {
                records::LABELSST => {
                    Self::parse_labelsst(&rec.data, ws, sst, styles, style_ctx)?;
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
                    let result = Self::parse_formula(
                        &rec.data,
                        ws,
                        styles,
                        formula_ctx,
                        &shared_formulas,
                        &array_formulas,
                        &data_tables,
                    )?;
                    match result {
                        FormulaResult::Done(string_pending) => {
                            pending_formula_cell = string_pending;
                        }
                        FormulaResult::SharedPending {
                            cell_row,
                            cell_col,
                            master_row,
                            master_col,
                        } => {
                            pending_shared = Some((cell_row, cell_col, master_row, master_col));
                            pending_formula_cell = None;
                        }
                        FormulaResult::TablePending {
                            cell_row,
                            cell_col,
                            master_row,
                            master_col,
                        } => {
                            pending_table_cells.push((cell_row, cell_col, master_row, master_col));
                            pending_formula_cell = None;
                        }
                    }
                }
                records::STRING => {
                    // Cached string value for the preceding FORMULA
                    if let Some((row, col)) = pending_formula_cell.take() {
                        Self::parse_formula_string(&rec.data, ws, row, col)?;
                    }
                }
                records::SHAREDFMLA => {
                    // Parse shared formula and store for later tExp resolution
                    if let Some((master_row, master_col, token_data)) =
                        Self::parse_sharedfmla(&rec.data)
                    {
                        shared_formulas.insert((master_row, master_col), token_data.clone());

                        // Backfill the pending master cell if it was waiting
                        if let Some((cell_row, cell_col, exp_row, exp_col)) = pending_shared.take()
                        {
                            if exp_row == master_row && exp_col == master_col {
                                Self::backfill_shared_formula(
                                    ws,
                                    cell_row,
                                    cell_col,
                                    &token_data,
                                    formula_ctx,
                                )?;
                            }
                        }
                    }
                }
                records::ARRAY => {
                    // Parse array formula and store for tExp resolution.
                    // ARRAY record: Ref8U(6) + options(2) + reserved(4) + cce(2) + rgce + rgcb
                    if let Some((master_row, master_col, token_data, extra_data)) =
                        Self::parse_array_record(&rec.data)
                    {
                        let key = (master_row, master_col);
                        array_formulas.insert(key, (token_data.clone(), extra_data.clone()));

                        // Backfill any pending cell waiting for this array formula
                        if let Some((cell_row, cell_col, exp_row, exp_col)) = pending_shared.take()
                        {
                            if exp_row == master_row && exp_col == master_col {
                                Self::backfill_array_formula(
                                    ws,
                                    cell_row,
                                    cell_col,
                                    &token_data,
                                    &extra_data,
                                    formula_ctx,
                                )?;
                            }
                        }
                    }
                }
                records::TABLE => {
                    if let Some((master_row, master_col, input1, input2)) =
                        Self::parse_table_record(&rec.data)
                    {
                        data_tables
                            .insert((master_row, master_col), (input1.clone(), input2.clone()));
                        // Backfill all pending cells for this table
                        let text = format!("=TABLE({},{})", input1, input2);
                        pending_table_cells.retain(|&(cell_row, cell_col, mr, mc)| {
                            if mr == master_row && mc == master_col {
                                Self::backfill_table_formula(ws, cell_row, cell_col, &text).ok();
                                false
                            } else {
                                true
                            }
                        });
                    }
                }
                records::MERGECELLS => Self::parse_mergecells(&rec.data, ws)?,
                records::ROW => Self::parse_row(&rec.data, ws)?,
                records::COLINFO => Self::parse_colinfo(&rec.data, ws)?,
                records::DEFCOLWIDTH => Self::parse_defcolwidth(&rec.data, ws)?,
                records::DEFAULTROWHEIGHT => Self::parse_defaultrowheight(&rec.data, ws)?,
                records::WINDOW2 => window2_frozen = Self::parse_window2(&rec.data),
                records::PANE => Self::parse_pane(&rec.data, ws, window2_frozen)?,
                records::SELECTION => Self::parse_selection(&rec.data, ws)?,
                records::SHEETLAYOUT => Self::parse_sheetlayout(&rec.data, ws),
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
                records::FEAT => Self::parse_feat_protection(&rec.data, ws),
                // ── Drawing / comments (OBJ → TXO → NOTE) ────────────
                records::MSODRAWING => {
                    escher_bytes.extend_from_slice(&rec.data);
                }
                records::OBJ => {
                    last_obj_id = Self::parse_obj_id(&rec.data);
                    obj_bodies.push(rec.data.clone());
                }
                records::TXO => {
                    if let Some(oid) = last_obj_id.take() {
                        if let Some(text) =
                            Self::parse_txo_text(&rec.data, &rec.continue_offsets, style_ctx)
                        {
                            obj_texts.insert(oid, text);
                        }
                    }
                }
                records::NOTE => {
                    if let Some(note) = Self::parse_note(&rec.data)? {
                        notes.push(note);
                    }
                }
                // ── Hyperlinks ──────────────────────────────────────
                records::HLINK => Self::parse_hlink(&rec.data, ws)?,
                records::HLINKTOOLTIP => {
                    Self::parse_hlinktooltip(&rec.data, &mut hlink_tooltips);
                }
                // ── Conditional formatting ──────────────────────────
                records::CONDFMT => cf_ranges = Self::parse_condfmt(&rec.data),
                records::CF => Self::parse_cf(&rec.data, ws, &cf_ranges, formula_ctx),
                // ── Data validation ─────────────────────────────────
                records::DVAL => {
                    // Header only - DV records follow with actual rules
                }
                records::DV => Self::parse_dv(&rec.data, ws, formula_ctx)?,
                records::AUTOFILTERINFO => {}
                records::AUTOFILTER => {
                    if let Some(fc) = Self::parse_autofilter(&rec.data) {
                        auto_filter_columns.push(fc);
                    }
                }
                records::SETUP => {
                    if rec.data.len() >= 12 {
                        let mut ps = ws.page_setup().clone();
                        let mut off = 0;
                        let paper_size = read_u16(&rec.data, &mut off).unwrap_or(1);
                        let scale = read_u16(&rec.data, &mut off).unwrap_or(100);
                        let _page_start = read_u16(&rec.data, &mut off).unwrap_or(1);
                        let fit_width = read_u16(&rec.data, &mut off).unwrap_or(0);
                        let fit_height = read_u16(&rec.data, &mut off).unwrap_or(0);
                        let grbit = read_u16(&rec.data, &mut off).unwrap_or(0);

                        ps.paper_size = paper_size.min(255) as u8;
                        ps.scale = scale.clamp(10, 400);
                        if fit_width > 0 {
                            ps.fit_to_width = Some(fit_width);
                        }
                        if fit_height > 0 {
                            ps.fit_to_height = Some(fit_height);
                        }

                        if (grbit & 0x0040) == 0 {
                            ps.orientation = if (grbit & 0x0002) != 0 {
                                duke_sheets_core::PageOrientation::Landscape
                            } else {
                                duke_sheets_core::PageOrientation::Portrait
                            };
                        }

                        if rec.data.len() >= 32 {
                            let mut hdr_off = 16usize;
                            if let Ok(hdr_margin) = read_f64(&rec.data, &mut hdr_off) {
                                ps.header_margin = hdr_margin;
                            }

                            let mut ftr_off = 24usize;
                            if let Ok(ftr_margin) = read_f64(&rec.data, &mut ftr_off) {
                                ps.footer_margin = ftr_margin;
                            }
                        }

                        ws.set_page_setup(ps);
                    }
                }
                records::HEADER => {
                    if rec.data.len() >= 3 {
                        let mut off = 0;
                        if let Ok(text) = read_unicode_string(&rec.data, &mut off) {
                            if !text.is_empty() {
                                let mut ps = ws.page_setup().clone();
                                ps.odd_header = Some(text);
                                ws.set_page_setup(ps);
                            }
                        }
                    }
                }
                records::FOOTER => {
                    if rec.data.len() >= 3 {
                        let mut off = 0;
                        if let Ok(text) = read_unicode_string(&rec.data, &mut off) {
                            if !text.is_empty() {
                                let mut ps = ws.page_setup().clone();
                                ps.odd_footer = Some(text);
                                ws.set_page_setup(ps);
                            }
                        }
                    }
                }
                records::LEFT_MARGIN => {
                    if rec.data.len() >= 8 {
                        let mut off = 0;
                        if let Ok(val) = read_f64(&rec.data, &mut off) {
                            let mut ps = ws.page_setup().clone();
                            ps.left_margin = val;
                            ws.set_page_setup(ps);
                        }
                    }
                }
                records::RIGHT_MARGIN => {
                    if rec.data.len() >= 8 {
                        let mut off = 0;
                        if let Ok(val) = read_f64(&rec.data, &mut off) {
                            let mut ps = ws.page_setup().clone();
                            ps.right_margin = val;
                            ws.set_page_setup(ps);
                        }
                    }
                }
                records::TOP_MARGIN => {
                    if rec.data.len() >= 8 {
                        let mut off = 0;
                        if let Ok(val) = read_f64(&rec.data, &mut off) {
                            let mut ps = ws.page_setup().clone();
                            ps.top_margin = val;
                            ws.set_page_setup(ps);
                        }
                    }
                }
                records::BOTTOM_MARGIN => {
                    if rec.data.len() >= 8 {
                        let mut off = 0;
                        if let Ok(val) = read_f64(&rec.data, &mut off) {
                            let mut ps = ws.page_setup().clone();
                            ps.bottom_margin = val;
                            ws.set_page_setup(ps);
                        }
                    }
                }
                records::HCENTER | records::VCENTER => {}
                records::SCL => {
                    // Zoom: numerator(u16) + denominator(u16)
                    if rec.data.len() >= 4 {
                        let num = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        let den = u16::from_le_bytes([rec.data[2], rec.data[3]]);
                        if den != 0 {
                            let zoom = ((num as u32) * 100 / (den as u32)) as u16;
                            ws.set_zoom_scale(Some(zoom.clamp(10, 400)));
                        }
                    }
                }
                records::PRINTHEADERS => {
                    if rec.data.len() >= 2 {
                        let flag = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        if flag != 0 {
                            let mut ps = ws.page_setup().clone();
                            ps.print_headings = true;
                            ws.set_page_setup(ps);
                        }
                    }
                }
                records::PRINTGRIDLINES => {
                    if rec.data.len() >= 2 {
                        let flag = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                        if flag != 0 {
                            let mut ps = ws.page_setup().clone();
                            ps.print_gridlines = true;
                            ws.set_page_setup(ps);
                        }
                    }
                }
                records::HPAGEBREAKS => Self::parse_page_breaks(&rec.data, ws, true),
                records::VPAGEBREAKS => Self::parse_page_breaks(&rec.data, ws, false),
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

        // Apply HLINKTOOLTIP records to existing hyperlinks
        for ((row, col), tooltip) in &hlink_tooltips {
            let addr = CellAddress::new(*row, *col).to_a1_string();
            if let Some(hl) = ws.hyperlink_mut(&addr) {
                hl.tooltip = Some(tooltip.clone());
            }
        }

        if let Some(range) = auto_filter_range {
            let af = duke_sheets_core::AutoFilter {
                range: range.clone(),
                filter_columns: auto_filter_columns,
            };
            ws.set_auto_filter(Some(af));
        } else if !auto_filter_columns.is_empty() {
            log::warn!("AUTOFILTER records found without _FilterDatabase name");
        }

        // Assemble the drawing list from the per-sheet Escher byte
        // stream (concatenated MSODRAWING bodies), the OBJ record
        // bodies, the TXO texts, and the NOTE records.
        Self::build_sheet_drawings(
            &escher_bytes,
            &obj_bodies,
            &obj_texts,
            &notes,
            blip_store,
            formula_ctx,
            ws,
        );

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
        sst: &[SstEntry],
        styles: &[Style],
        style_ctx: &StyleContext,
    ) -> XlsResult<()> {
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let xf_idx = read_u16(data, &mut off)?;
        let sst_idx = read_u32(data, &mut off)? as usize;

        if let Some(entry) = sst.get(sst_idx) {
            match entry {
                SstEntry::Plain(s) => {
                    ws.set_cell_value_at(row, col, CellValue::String(SharedString::new(s)))?;
                }
                SstEntry::Rich { text, runs } => {
                    let mut rich_runs = Self::sst_runs_to_rich_text(text, runs, style_ctx);
                    // If the first run has font: None, it inherits from the cell's
                    // XF font record.  Resolve the XF font so the first run carries
                    // the correct formatting (italic, color, etc.).
                    if let Some(first) = rich_runs.first_mut() {
                        if first.font.is_none() {
                            if let Some(xf) = style_ctx.xfs.get(xf_idx as usize) {
                                first.font = style_ctx.resolve_run_font(xf.font_index);
                            }
                        }
                    }
                    ws.set_cell_value_at(row, col, CellValue::rich_text(rich_runs))?;
                }
            }
        }
        Self::apply_style(ws, row, col, xf_idx, styles)?;
        Ok(())
    }

    /// Convert BIFF8 SST formatting runs into `RichTextRun` segments.
    ///
    /// Each `FormattingRun` marks where a new font begins (char_pos, font_index).
    /// We split the text at those character positions and attach the resolved
    /// `RunFont` from the workbook's font table.
    fn sst_runs_to_rich_text(
        text: &str,
        runs: &[crate::biff::strings::FormattingRun],
        style_ctx: &StyleContext,
    ) -> Vec<RichTextRun> {
        use duke_sheets_core::rich_text::RunFont;

        if runs.is_empty() {
            return vec![RichTextRun::plain(text)];
        }

        let chars: Vec<char> = text.chars().collect();
        let total_chars = chars.len();
        let mut result = Vec::new();

        // Build (start_pos, font_index) boundaries
        // If the first run doesn't start at 0, the leading text has no special font.
        let mut boundaries: Vec<(usize, Option<u16>)> = Vec::new();

        if runs[0].char_pos > 0 {
            boundaries.push((0, None)); // leading text with no run font
        }
        for run in runs {
            boundaries.push((run.char_pos as usize, Some(run.font_index)));
        }

        for (i, &(start, font_idx)) in boundaries.iter().enumerate() {
            if start >= total_chars {
                break;
            }
            let end = boundaries
                .get(i + 1)
                .map(|&(pos, _)| pos.min(total_chars))
                .unwrap_or(total_chars);
            if end <= start {
                continue;
            }
            let segment: String = chars[start..end].iter().collect();
            let font: Option<RunFont> = font_idx.and_then(|idx| style_ctx.resolve_run_font(idx));
            result.push(RichTextRun {
                text: segment,
                font,
            });
        }

        if result.is_empty() {
            vec![RichTextRun::plain(text)]
        } else {
            result
        }
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
    fn parse_formula(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        styles: &[Style],
        formula_ctx: &FormulaContext,
        shared_formulas: &std::collections::HashMap<(u16, u16), Vec<u8>>,
        array_formulas: &std::collections::HashMap<(u16, u16), (Vec<u8>, Vec<u8>)>,
        data_tables: &std::collections::HashMap<(u16, u16), (String, String)>,
    ) -> XlsResult<FormulaResult> {
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

        let options = read_u16(data, &mut off)?;
        let _reserved = read_u32(data, &mut off)?;
        // off is now 20

        let f_shared = (options & 0x0008) != 0;

        // Parse the token bytes to check for tExp (shared/array formula indicator)
        let formula_text = if off + 2 <= data.len() {
            let cce = read_u16(data, &mut off)? as usize;
            if cce > 0 && off + cce <= data.len() {
                let token_bytes = &data[off..off + cce];
                let extra_data = &data[off + cce..];

                // Check if this is a shared formula (tExp + fShared flag)
                let tokens = crate::biff::formula::token_parser::parse_tokens_with_extra(
                    token_bytes,
                    extra_data,
                );
                if f_shared {
                    if let Some(ParsedToken::Exp {
                        row: master_row,
                        col: master_col,
                    }) = tokens.first()
                    {
                        // Try to resolve the shared formula
                        if let Some(shared_tokens) =
                            shared_formulas.get(&(*master_row as u16, *master_col))
                        {
                            // Shared formula found - decompile with base cell
                            let shared_ctx = FormulaContext {
                                sheet_names: formula_ctx.sheet_names.clone(),
                                extern_sheet: formula_ctx.extern_sheet.clone(),
                                supbooks: formula_ctx.supbooks.clone(),
                                names: formula_ctx.names.clone(),
                                extern_names: formula_ctx.extern_names.clone(),
                                extern_name_index_base: formula_ctx.extern_name_index_base,
                                base_cell: Some((row, col)),
                            };
                            let text = crate::biff::formula::decompile(shared_tokens, &shared_ctx);
                            if text.is_empty() {
                                String::new()
                            } else {
                                format!("={}", text)
                            }
                        } else {
                            // SHAREDFMLA not seen yet - write cell now with
                            // empty text; caller will backfill later.
                            Self::write_formula_cell(
                                ws,
                                row,
                                col,
                                xf_idx,
                                result_bytes,
                                String::new(),
                                styles,
                            )?;
                            return Ok(FormulaResult::SharedPending {
                                cell_row: row,
                                cell_col: col,
                                master_row: *master_row as u16,
                                master_col: *master_col,
                            });
                        }
                    } else {
                        // fShared set but no tExp - decompile normally
                        let text =
                            crate::biff::formula::decompiler::decompile(&tokens, formula_ctx);
                        if text.is_empty() {
                            String::new()
                        } else {
                            format!("={}", text)
                        }
                    }
                } else if let Some(ParsedToken::Exp {
                    row: master_row,
                    col: master_col,
                }) = tokens.first()
                {
                    // tExp without fShared → array formula (CSE)
                    if let Some((arr_tokens, arr_extra)) =
                        array_formulas.get(&(*master_row as u16, *master_col))
                    {
                        let text = crate::biff::formula::decompile_with_extra(
                            arr_tokens,
                            arr_extra,
                            formula_ctx,
                        );
                        if text.is_empty() {
                            String::new()
                        } else {
                            format!("{{={}}}", text)
                        }
                    } else {
                        // ARRAY record not seen yet - write cell with empty
                        // text; caller will backfill later.
                        Self::write_formula_cell(
                            ws,
                            row,
                            col,
                            xf_idx,
                            result_bytes,
                            String::new(),
                            styles,
                        )?;
                        return Ok(FormulaResult::SharedPending {
                            cell_row: row,
                            cell_col: col,
                            master_row: *master_row as u16,
                            master_col: *master_col,
                        });
                    }
                } else if let Some(ParsedToken::Table {
                    row: master_row,
                    col: master_col,
                }) = tokens.first()
                {
                    // Data table formula (tTbl)
                    if let Some((input1, input2)) =
                        data_tables.get(&(*master_row as u16, *master_col))
                    {
                        format!("=TABLE({},{})", input1, input2)
                    } else {
                        // TABLE record not seen yet - write cell with empty text
                        Self::write_formula_cell(
                            ws,
                            row,
                            col,
                            xf_idx,
                            result_bytes,
                            String::new(),
                            styles,
                        )?;
                        return Ok(FormulaResult::TablePending {
                            cell_row: row,
                            cell_col: col,
                            master_row: *master_row as u16,
                            master_col: *master_col,
                        });
                    }
                } else {
                    // Normal formula - decompile from already-parsed tokens
                    let text = crate::biff::formula::decompiler::decompile(&tokens, formula_ctx);
                    if text.is_empty() {
                        String::new()
                    } else {
                        format!("={}", text)
                    }
                }
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        Self::write_formula_cell(ws, row, col, xf_idx, result_bytes, formula_text, styles)?;

        // Check if cached result is a string (STRING record follows)
        let string_pending =
            result_bytes[6] == 0xFF && result_bytes[7] == 0xFF && result_bytes[0] == 0x00;
        if string_pending {
            Ok(FormulaResult::Done(Some((row, col))))
        } else {
            Ok(FormulaResult::Done(None))
        }
    }

    /// Write a formula cell with its cached value and formula text.
    fn write_formula_cell(
        ws: &mut duke_sheets_core::Worksheet,
        row: u32,
        col: u16,
        xf_idx: u16,
        result_bytes: &[u8],
        formula_text: String,
        styles: &[Style],
    ) -> XlsResult<()> {
        let cached_value = if result_bytes[6] == 0xFF && result_bytes[7] == 0xFF {
            let result_type = result_bytes[0];
            match result_type {
                0x00 => CellValue::Empty,
                0x01 => CellValue::Boolean(result_bytes[2] != 0),
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
                    CellValue::Error(err)
                }
                _ => CellValue::Empty,
            }
        } else {
            CellValue::Number(f64::from_le_bytes(result_bytes.try_into().unwrap()))
        };

        ws.set_formula_with_cached_value_at(row, col, &formula_text, cached_value)?;
        Self::apply_style(ws, row, col, xf_idx, styles)?;
        Ok(())
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

        let value = CellValue::String(SharedString::new(&string_val));
        if ws.has_formula_at(row, col) {
            ws.set_formula_result(row, col, value)?;
        } else {
            ws.set_cell_value_at(row, col, value)?;
        }
        Ok(())
    }

    /// SHAREDFMLA (ShrFmla): Ref8U(6) + reserved(1) + cUse(1) + cce(2) + rgce(cce)
    ///
    /// Ref8U: rwFirst(2) + rwLast(2) + colFirst(1) + colLast(1) = 6 bytes.
    /// Returns (master_row, master_col, token_data) or None if too short.
    fn parse_sharedfmla(data: &[u8]) -> Option<(u16, u16, Vec<u8>)> {
        // Minimum: 6 (Ref8U) + 1 (reserved) + 1 (cUse) + 2 (cce) = 10
        if data.len() < 10 {
            return None;
        }
        let first_row = u16::from_le_bytes([data[0], data[1]]);
        let _last_row = u16::from_le_bytes([data[2], data[3]]);
        let first_col = data[4] as u16;
        let _last_col = data[5];
        // data[6] = reserved
        // data[7] = cUse (number of existing FORMULA records for this shared formula)
        let cce = u16::from_le_bytes([data[8], data[9]]) as usize;
        if data.len() < 10 + cce {
            return None;
        }
        let token_data = data[10..10 + cce].to_vec();
        Some((first_row, first_col, token_data))
    }

    /// TABLE record (0x0236): data table metadata.
    ///
    /// Format: rwFirst(2) + rwLast(2) + colFirst(1) + colLast(1) + flags(2)
    ///       + rwInpRw(2) + colInpRw(2) + rwInpCol(2) + colInpCol(2)
    ///
    /// Total: 16 bytes.
    /// Returns (master_row, master_col, input1_ref, input2_ref).
    fn parse_table_record(data: &[u8]) -> Option<(u16, u16, String, String)> {
        if data.len() < 16 {
            return None;
        }
        let first_row = u16::from_le_bytes([data[0], data[1]]);
        let _last_row = u16::from_le_bytes([data[2], data[3]]);
        let first_col = data[4] as u16;
        let _last_col = data[5];
        let flags = u16::from_le_bytes([data[6], data[7]]);
        let _f_always_calc = (flags & 0x0001) != 0;
        let f_rw = (flags & 0x0002) != 0;
        let f_tbl2 = (flags & 0x0004) != 0;

        let rw_inp_rw = u16::from_le_bytes([data[8], data[9]]);
        let col_inp_rw = u16::from_le_bytes([data[10], data[11]]);
        let rw_inp_col = u16::from_le_bytes([data[12], data[13]]);
        let col_inp_col = u16::from_le_bytes([data[14], data[15]]);

        let (input1, input2) = if f_tbl2 {
            // Two-variable table: row input + column input
            let r1 = duke_sheets_core::CellAddress::new(rw_inp_rw as u32, col_inp_rw).to_string();
            let r2 = duke_sheets_core::CellAddress::new(rw_inp_col as u32, col_inp_col).to_string();
            (r1, r2)
        } else if f_rw {
            // One-variable row input table
            let r1 = duke_sheets_core::CellAddress::new(rw_inp_rw as u32, col_inp_rw).to_string();
            (r1, String::new())
        } else {
            // One-variable column input table
            let r1 = duke_sheets_core::CellAddress::new(rw_inp_col as u32, col_inp_col).to_string();
            (String::new(), r1)
        };

        Some((first_row, first_col, input1, input2))
    }

    /// Backfill a cell's formula text after its SHAREDFMLA record is found.
    fn backfill_shared_formula(
        ws: &mut duke_sheets_core::Worksheet,
        cell_row: u32,
        cell_col: u16,
        shared_tokens: &[u8],
        formula_ctx: &FormulaContext,
    ) -> XlsResult<()> {
        // Build a temporary context with the base cell for offset resolution
        let shared_ctx = FormulaContext {
            sheet_names: formula_ctx.sheet_names.clone(),
            extern_sheet: formula_ctx.extern_sheet.clone(),
            supbooks: formula_ctx.supbooks.clone(),
            names: formula_ctx.names.clone(),
            extern_names: formula_ctx.extern_names.clone(),
            extern_name_index_base: formula_ctx.extern_name_index_base,
            base_cell: Some((cell_row, cell_col)),
        };
        let text = crate::biff::formula::decompile(shared_tokens, &shared_ctx);
        let formula_text = if text.is_empty() {
            String::new()
        } else {
            format!("={}", text)
        };

        if let Some(formula) = ws.formula_data_at_mut(cell_row, cell_col) {
            formula.text = formula_text;
        } else {
            let cached = ws.get_value_at(cell_row, cell_col);
            ws.set_formula_with_cached_value_at(cell_row, cell_col, &formula_text, cached)?;
        }
        Ok(())
    }

    /// Backfill a cell's formula text after its TABLE record is found.
    fn backfill_table_formula(
        ws: &mut duke_sheets_core::Worksheet,
        cell_row: u32,
        cell_col: u16,
        table_text: &str,
    ) -> XlsResult<()> {
        if let Some(formula) = ws.formula_data_at_mut(cell_row, cell_col) {
            formula.text = table_text.to_string();
        } else {
            let cached = ws.get_value_at(cell_row, cell_col);
            ws.set_formula_with_cached_value_at(cell_row, cell_col, table_text, cached)?;
        }
        Ok(())
    }

    /// Parse an ARRAY record.
    ///
    /// Format: Ref8U(6) + options(2) + reserved(4) + cce(2) + rgce(cce) + rgcb(rest)
    /// Ref8U: rwFirst(2) + rwLast(2) + colFirst(1) + colLast(1) = 6 bytes.
    /// Returns (master_row, master_col, token_data, extra_data) or None.
    fn parse_array_record(data: &[u8]) -> Option<(u16, u16, Vec<u8>, Vec<u8>)> {
        // Minimum: 6 (Ref8U) + 2 (options) + 4 (reserved) + 2 (cce) = 14
        if data.len() < 14 {
            return None;
        }
        let first_row = u16::from_le_bytes([data[0], data[1]]);
        let _last_row = u16::from_le_bytes([data[2], data[3]]);
        let first_col = data[4] as u16;
        let _last_col = data[5];
        // data[6..8] = options (fAlwaysCalc, fCalcOnLoad)
        // data[8..12] = reserved
        let cce = u16::from_le_bytes([data[12], data[13]]) as usize;
        if data.len() < 14 + cce {
            return None;
        }
        let token_data = data[14..14 + cce].to_vec();
        let extra_data = data[14 + cce..].to_vec();
        Some((first_row, first_col, token_data, extra_data))
    }

    /// Backfill a cell's formula text after its ARRAY record is found.
    fn backfill_array_formula(
        ws: &mut duke_sheets_core::Worksheet,
        cell_row: u32,
        cell_col: u16,
        token_data: &[u8],
        extra_data: &[u8],
        formula_ctx: &FormulaContext,
    ) -> XlsResult<()> {
        let text = crate::biff::formula::decompile_with_extra(token_data, extra_data, formula_ctx);
        let formula_text = if text.is_empty() {
            String::new()
        } else {
            format!("{{={}}}", text)
        };

        if let Some(formula) = ws.formula_data_at_mut(cell_row, cell_col) {
            formula.text = formula_text;
        } else {
            let cached = ws.get_value_at(cell_row, cell_col);
            ws.set_formula_with_cached_value_at(cell_row, cell_col, &formula_text, cached)?;
        }
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
            let outline_level = ((options >> 8) & 0x07) as u8;
            let collapsed = (options & 0x10) != 0;

            if hidden {
                ws.set_row_hidden(row_index, true);
            }
            if custom_height && height_pt > 0.0 {
                ws.set_row_height(row_index, height_pt);
            }
            if outline_level > 0 {
                ws.set_row_outline_level(row_index, outline_level);
            }
            if collapsed {
                ws.set_row_collapsed(row_index, true);
            }
        }

        Ok(())
    }

    // ── Formula context record parsers ──────────────────────────────────

    /// Parse a SUPBOOK record.
    ///
    /// Self-reference: cch == 0x0401. Add-in: ctab == 1, cch == 0x3A01.
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

        if ctab == 1 && cch == 0x3A01 {
            // Add-in functions sentinel (MS-XLS §2.4.273: ctab=0x0001,
            // cch=0x3A01). Followed by one EXTERNNAME per add-in function.
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

    /// Parse an EXTERNNAME (0x0023) record, returning the external name string.
    ///
    /// Body (MS-XLS §2.4.150): grbit(2) + reserved(2) + reserved(2) + the name
    /// as a ShortXLUnicodeString (cch:1, grbit:1, chars). The trailing
    /// name-definition formula is ignored — for an add-in function it is a
    /// `#REF!` placeholder (`02 00 1C 17`).
    fn parse_externname(data: &[u8]) -> XlsResult<String> {
        if data.len() < 8 {
            return Err(XlsError::Parse("EXTERNNAME record too short".into()));
        }
        let cch = data[6] as u16;
        let flags = data[7];
        let mut off = 8;
        read_character_data(data, &mut off, cch, flags)
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
        let cce = read_u16(data, &mut off)?;
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
                let skip = if (name_flags & 0x01) != 0 { 2 } else { 1 };
                off += skip;
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

        let formula_body = if cce > 0 && off + cce as usize <= data.len() {
            data[off..off + cce as usize].to_vec()
        } else {
            Vec::new()
        };

        Ok(NameRecord {
            name,
            sheet_idx: if itab == 0 {
                0xFFFFFFFF
            } else {
                (itab - 1) as u32
            },
            is_builtin,
            formula_body,
            comment: None,
        })
    }

    fn extract_filter_db_range(
        formula_body: &[u8],
        _formula_ctx: &FormulaContext,
    ) -> Option<CellRange> {
        // The _FilterDatabase name body is a single tArea3d (0x3B base) that
        // covers the filter range, so only the first token is relevant.
        if formula_body.first().map(|t| t & 0x7F) != Some(0x3B) || formula_body.len() < 11 {
            return None;
        }
        let first_row = u16::from_le_bytes([formula_body[3], formula_body[4]]) as u32;
        let last_row = u16::from_le_bytes([formula_body[5], formula_body[6]]) as u32;
        let first_col = u16::from_le_bytes([formula_body[7], formula_body[8]]) & 0x3FFF;
        let last_col = u16::from_le_bytes([formula_body[9], formula_body[10]]) & 0x3FFF;
        Some(CellRange::from_indices(
            first_row, first_col, last_row, last_col,
        ))
    }

    /// Extract repeat rows and repeat cols from a Print_Titles NAME formula body.
    ///
    /// The formula body may contain:
    /// - A single tArea3d for row titles only (full-column range)
    /// - A single tArea3d for column titles only (full-row range)
    /// - tMemFunc + two tArea3d tokens for both row AND column titles
    ///
    /// Row titles: first_col == 0 && last_col == 0xFF (or 0x3FFF), meaning entire rows.
    /// Column titles: first_row == 0 && last_row == 0xFFFF (or 0xFFFF), meaning entire columns.
    fn extract_print_titles(formula_body: &[u8]) -> (Option<(u32, u32)>, Option<(u16, u16)>) {
        let mut repeat_rows: Option<(u32, u32)> = None;
        let mut repeat_cols: Option<(u16, u16)> = None;
        let mut pos = 0;

        while pos < formula_body.len() {
            let token = formula_body[pos];
            let base = token & 0x7F;
            match base {
                // tArea3d: ixti(2) + first_row(2) + last_row(2) + first_col(2) + last_col(2)
                0x3B => {
                    if pos + 11 <= formula_body.len() {
                        let first_row =
                            u16::from_le_bytes([formula_body[pos + 3], formula_body[pos + 4]]);
                        let last_row =
                            u16::from_le_bytes([formula_body[pos + 5], formula_body[pos + 6]]);
                        let first_col =
                            u16::from_le_bytes([formula_body[pos + 7], formula_body[pos + 8]])
                                & 0x3FFF;
                        let last_col =
                            u16::from_le_bytes([formula_body[pos + 9], formula_body[pos + 10]])
                                & 0x3FFF;

                        // Detect row titles: spans all columns (0..0xFF or 0..0x3FFF)
                        // → repeat_rows = (first_row, last_row)
                        if first_col == 0 && last_col >= 0xFF {
                            repeat_rows = Some((first_row as u32, last_row as u32));
                        }
                        // Detect column titles: spans all rows (0..0xFFFF)
                        // → repeat_cols = (first_col, last_col)
                        else if first_row == 0 && last_row == 0xFFFF {
                            repeat_cols = Some((first_col, last_col));
                        }
                    }
                    pos += 11; // skip full tArea3d
                }
                // tMemFunc: size(2) - skip the size field, tokens follow inline
                0x29 => {
                    if pos + 3 <= formula_body.len() {
                        pos += 3; // token(1) + cce(2)
                    } else {
                        break;
                    }
                }
                // tList (union operator) - 1 byte, skip
                0x10 => pos += 1,
                _ => break,
            }
        }

        (repeat_rows, repeat_cols)
    }

    fn parse_autofilter(data: &[u8]) -> Option<duke_sheets_core::FilterColumn> {
        use duke_sheets_core::auto_filter::*;

        if data.len() < 24 {
            return None;
        }

        let i_entry = u16::from_le_bytes([data[0], data[1]]);
        let flags = u16::from_le_bytes([data[2], data[3]]);

        let join_or = (flags & 0x03) != 0;
        let f_top_n = (flags & 0x0010) != 0;
        let f_top = (flags & 0x0020) != 0;
        let f_percent = (flags & 0x0040) != 0;
        let w_top_n = (flags >> 7) & 0x01FF;

        if f_top_n {
            return Some(FilterColumn::new(
                i_entry as u32,
                ColumnFilter::Top10(Top10Filter {
                    top: f_top,
                    percent: f_percent,
                    val: w_top_n as f64,
                    filter_val: None,
                }),
            ));
        }

        let (vt1, op1, val1_info) = Self::parse_afdoper(&data[4..14]);
        let (vt2, op2, val2_info) = Self::parse_afdoper(&data[14..24]);

        let mut pos = 24;
        let str1 = if vt1 == 0x06 {
            if let DoperValue::StrInfo { cch } = &val1_info {
                Some(Self::read_xlstring_nocch(data, &mut pos, *cch as usize))
            } else {
                None
            }
        } else {
            None
        };

        let str2 = if vt2 == 0x06 {
            if let DoperValue::StrInfo { cch } = &val2_info {
                Some(Self::read_xlstring_nocch(data, &mut pos, *cch as usize))
            } else {
                None
            }
        } else {
            None
        };

        if vt1 == 0x0C && vt2 == 0x00 {
            return Some(FilterColumn::new(
                i_entry as u32,
                ColumnFilter::Values(ValueFilter {
                    values: vec![],
                    blank: true,
                }),
            ));
        }

        if vt1 == 0x0E && vt2 == 0x00 {
            return Some(FilterColumn::new(
                i_entry as u32,
                ColumnFilter::Custom(CustomFilters {
                    and: false,
                    conditions: vec![CustomFilterCondition {
                        operator: FilterOperator::NotEqual,
                        value: String::new(),
                    }],
                }),
            ));
        }

        let mut conditions = Vec::new();
        if vt1 != 0x00 {
            if let Some(cond) = Self::doper_to_condition(op1, &val1_info, &str1) {
                conditions.push(cond);
            }
        }
        if vt2 != 0x00 {
            if let Some(cond) = Self::doper_to_condition(op2, &val2_info, &str2) {
                conditions.push(cond);
            }
        }

        if conditions.is_empty() {
            return None;
        }

        if conditions.len() <= 2
            && conditions
                .iter()
                .all(|c| c.operator == FilterOperator::Equal)
            && join_or
        {
            return Some(FilterColumn::new(
                i_entry as u32,
                ColumnFilter::Values(ValueFilter {
                    values: conditions.into_iter().map(|c| c.value).collect(),
                    blank: false,
                }),
            ));
        }

        Some(FilterColumn::new(
            i_entry as u32,
            ColumnFilter::Custom(CustomFilters {
                and: !join_or,
                conditions,
            }),
        ))
    }

    fn parse_afdoper(data: &[u8]) -> (u8, u8, DoperValue) {
        if data.len() < 10 {
            return (0x00, 0x00, DoperValue::None);
        }

        let vt = data[0];
        let op = data[1];
        let payload = &data[2..10];

        let value = match vt {
            0x00 => DoperValue::None,
            0x02 => {
                let rk_bytes = [payload[0], payload[1], payload[2], payload[3]];
                DoperValue::Float(Self::decode_rk_value(&rk_bytes))
            }
            0x04 => DoperValue::Float(f64::from_le_bytes(payload.try_into().unwrap())),
            0x06 => DoperValue::StrInfo { cch: payload[4] },
            0x08 => {
                if payload[1] != 0 {
                    DoperValue::Error(payload[0])
                } else {
                    DoperValue::Bool(payload[0] != 0)
                }
            }
            0x0C => DoperValue::Blanks,
            0x0E => DoperValue::NonBlanks,
            _ => DoperValue::None,
        };

        (vt, op, value)
    }

    fn decode_rk_value(bytes: &[u8; 4]) -> f64 {
        let raw = u32::from_le_bytes(*bytes);
        let fx100 = (raw & 1) != 0;
        let fint = (raw & 2) != 0;

        let value = if fint {
            ((raw as i32) >> 2) as f64
        } else {
            f64::from_bits(((raw & 0xFFFF_FFFC) as u64) << 32)
        };

        if fx100 {
            value / 100.0
        } else {
            value
        }
    }

    fn read_xlstring_nocch(data: &[u8], pos: &mut usize, cch: usize) -> String {
        if *pos >= data.len() || cch == 0 {
            return String::new();
        }

        let flags = data[*pos];
        *pos += 1;
        let high_byte = (flags & 1) != 0;

        if high_byte {
            let byte_len = cch * 2;
            if *pos + byte_len > data.len() {
                return String::new();
            }
            let chars: Vec<u16> = data[*pos..*pos + byte_len]
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect();
            *pos += byte_len;
            String::from_utf16_lossy(&chars)
        } else {
            if *pos + cch > data.len() {
                return String::new();
            }
            let s = String::from_utf8_lossy(&data[*pos..*pos + cch]).into_owned();
            *pos += cch;
            s
        }
    }

    fn doper_to_condition(
        op: u8,
        value: &DoperValue,
        str_val: &Option<String>,
    ) -> Option<duke_sheets_core::auto_filter::CustomFilterCondition> {
        use duke_sheets_core::auto_filter::*;

        let operator = match op {
            0x01 => FilterOperator::LessThan,
            0x02 => FilterOperator::Equal,
            0x03 => FilterOperator::LessThanOrEqual,
            0x04 => FilterOperator::GreaterThan,
            0x05 => FilterOperator::NotEqual,
            0x06 => FilterOperator::GreaterThanOrEqual,
            _ => return None,
        };

        let val_str = match value {
            DoperValue::Float(f) => format!("{f}"),
            DoperValue::StrInfo { .. } => str_val.clone().unwrap_or_default(),
            DoperValue::Bool(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            DoperValue::Error(e) => format!("#ERR{e}"),
            _ => return None,
        };

        Some(CustomFilterCondition {
            operator,
            value: val_str,
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
        let outline_level = ((options >> 8) & 0x0007) as u8;
        let collapsed = (options & 0x1000) != 0;
        let width_chars = raw_width as f64 / 256.0;

        for col in first_col..=last_col {
            if hidden {
                ws.set_column_hidden(col, true);
            }
            if width_chars > 0.0 {
                ws.set_column_width(col, width_chars);
            }
            if outline_level > 0 {
                ws.set_column_outline_level(col, outline_level);
            }
            if collapsed {
                ws.set_column_collapsed(col, true);
            }
        }

        Ok(())
    }

    fn parse_defcolwidth(data: &[u8], ws: &mut duke_sheets_core::Worksheet) -> XlsResult<()> {
        if data.len() < 2 {
            return Ok(());
        }
        let mut off = 0;
        let width_chars = read_u16(data, &mut off)? as f64;
        if width_chars > 0.0 {
            ws.set_default_column_width(width_chars);
        }
        Ok(())
    }

    fn parse_defaultrowheight(data: &[u8], ws: &mut duke_sheets_core::Worksheet) -> XlsResult<()> {
        if data.len() < 2 {
            return Ok(());
        }

        let mut off = 0;
        let raw_height = if data.len() >= 4 {
            let _flags = read_u16(data, &mut off)?;
            read_u16(data, &mut off)?
        } else {
            read_u16(data, &mut off)?
        };

        let height_pt = (raw_height & 0x7FFF) as f64 / 20.0;
        if height_pt > 0.0 {
            ws.set_default_row_height(height_pt);
        }
        Ok(())
    }

    fn parse_window2(data: &[u8]) -> bool {
        if data.len() < 2 {
            return false;
        }
        let options = u16::from_le_bytes([data[0], data[1]]);
        (options & 0x0008) != 0
    }

    fn parse_pane(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        window2_frozen: bool,
    ) -> XlsResult<()> {
        if data.len() < 10 {
            return Ok(());
        }
        let mut off = 0;
        let x = read_u16(data, &mut off)?;
        let y = read_u16(data, &mut off)?;
        let _top_row = read_u16(data, &mut off)?;
        let _left_col = read_u16(data, &mut off)?;
        let _active_pane = read_u16(data, &mut off)?;

        if window2_frozen && (x > 0 || y > 0) {
            ws.set_freeze_panes(y as u32, x);
        }

        Ok(())
    }

    fn parse_selection(data: &[u8], ws: &mut duke_sheets_core::Worksheet) -> XlsResult<()> {
        if data.len() < 9 {
            return Ok(());
        }

        let mut off = 0;
        let _pane = data[off];
        off += 1;
        let active_row = read_u16(data, &mut off)? as u32;
        let active_col = read_u16(data, &mut off)?;
        let _active_ref = read_u16(data, &mut off)?;
        let ref_count = read_u16(data, &mut off)? as usize;

        let mut refs = Vec::new();
        for _ in 0..ref_count {
            if off + 6 > data.len() {
                break;
            }
            let r1 = read_u16(data, &mut off)? as u32;
            let r2 = read_u16(data, &mut off)? as u32;
            let c1 = read_u8(data, &mut off)? as u16;
            let c2 = read_u8(data, &mut off)? as u16;
            refs.push(duke_sheets_core::CellRange::from_indices(r1, c1, r2, c2).to_string());
        }

        let active = duke_sheets_core::CellAddress::new(active_row, active_col).to_string();
        let sqref = if refs.is_empty() {
            None
        } else {
            Some(refs.join(" "))
        };

        ws.add_selection(Selection {
            pane: None,
            active_cell: Some(active),
            sqref,
        });

        Ok(())
    }

    fn parse_sheetlayout(data: &[u8], ws: &mut duke_sheets_core::Worksheet) {
        if data.len() >= 11 {
            let r = data[8];
            let g = data[9];
            let b = data[10];
            ws.set_tab_color(Some(duke_sheets_core::Color::rgb(r, g, b)));
        }
    }

    fn parse_page_breaks(data: &[u8], ws: &mut duke_sheets_core::Worksheet, is_row: bool) {
        use duke_sheets_core::worksheet::PageBreak;

        if data.len() < 2 {
            return;
        }

        let count = u16::from_le_bytes([data[0], data[1]]) as usize;
        let mut breaks = Vec::with_capacity(count);
        let mut off = 2;

        for _ in 0..count {
            if off + 6 > data.len() {
                break;
            }

            let id = u16::from_le_bytes([data[off], data[off + 1]]) as u32;
            let min = u16::from_le_bytes([data[off + 2], data[off + 3]]) as u32;
            let max = u16::from_le_bytes([data[off + 4], data[off + 5]]) as u32;
            off += 6;

            breaks.push(PageBreak {
                id,
                min,
                max,
                man: true,
                pt: false,
            });
        }

        if is_row {
            ws.set_row_breaks(breaks);
        } else {
            ws.set_col_breaks(breaks);
        }
    }

    /// Parse the per-sheet Escher stream into a tree of shape nodes
    /// in document (pre-order) order: a group's own SP container,
    /// then its children, then the group's next sibling. The
    /// patriarch shape (the implicit per-sheet root group) is
    /// dropped, and its `SpgrContainer` is transparent — its members
    /// become the top-level node list.
    fn parse_shape_tree(escher_bytes: &[u8]) -> Vec<EscherShapeNode> {
        let mut nodes = Vec::new();
        let mut budget: usize = 1_000_000;
        Self::collect_shape_nodes(escher_bytes, 0, &mut nodes, &mut budget);
        nodes
    }

    fn collect_shape_nodes(
        body: &[u8],
        depth: usize,
        out: &mut Vec<EscherShapeNode>,
        budget: &mut usize,
    ) {
        use crate::biff::escher::{rec_type as er, OfficeArtRecordHeader, HEADER_LEN};
        // Groups nest at most a handful of levels in real files.
        const MAX_DEPTH: usize = 64;
        if depth > MAX_DEPTH {
            return;
        }
        let mut cursor = 0usize;
        while cursor.saturating_add(HEADER_LEN) <= body.len() {
            if *budget == 0 {
                return;
            }
            *budget -= 1;
            let Ok(h) = OfficeArtRecordHeader::read_from(&body[cursor..]) else {
                return;
            };
            let Some((body_start, body_end)) =
                Self::escher_record_bounds(cursor, body.len(), h.rec_len)
            else {
                return;
            };
            let inner = &body[body_start..body_end];
            match h.rec_type {
                er::SP_CONTAINER => {
                    if let Some(node) = Self::parse_sp_container_node(inner) {
                        out.push(node);
                    }
                }
                er::SPGR_CONTAINER => {
                    let mut members = Vec::new();
                    Self::collect_shape_nodes(inner, depth + 1, &mut members, budget);
                    // The container's first member is the group's own
                    // shape (the leaf SP carrying the FSPGR). The
                    // patriarch SPGR falls through to the flatten arm
                    // because its own SP was dropped above.
                    if !members.is_empty() && !members[0].is_group && members[0].fspgr.is_some() {
                        let mut own = members.remove(0);
                        own.is_group = true;
                        own.children = members;
                        out.push(own);
                    } else {
                        out.append(&mut members);
                    }
                }
                _ if h.is_container() => {
                    Self::collect_shape_nodes(inner, depth + 1, out, budget);
                }
                _ => {}
            }
            cursor = body_end;
        }
    }

    /// Parse one `SP_CONTAINER` body into a shape node. Returns
    /// `None` for the patriarch shape and for deleted tombstone
    /// shapes — neither participates in OBJ pairing.
    fn parse_sp_container_node(sp_body: &[u8]) -> Option<EscherShapeNode> {
        use crate::biff::escher::{
            fsp_flags, rec_type as er, FoptTable, FoptValue, OfficeArtChildAnchor,
            OfficeArtClientAnchor, OfficeArtFsp, OfficeArtFspgr, OfficeArtRecordHeader, HEADER_LEN,
        };

        let mut node = EscherShapeNode::default();
        let mut has_fsp = false;
        let mut patriarch = false;
        let mut deleted = false;

        let mut cursor = 0usize;
        while cursor.saturating_add(HEADER_LEN) <= sp_body.len() {
            let Ok(h) = OfficeArtRecordHeader::read_from(&sp_body[cursor..]) else {
                break;
            };
            let Some((_, body_end)) = Self::escher_record_bounds(cursor, sp_body.len(), h.rec_len)
            else {
                break;
            };
            match h.rec_type {
                er::FSP => {
                    if let Ok((fsp, st, _)) = OfficeArtFsp::read_from(&sp_body[cursor..]) {
                        has_fsp = true;
                        node.shape_type = st;
                        node.spid = fsp.spid;
                        node.flip_h = fsp.grf_persistence & fsp_flags::FLIP_H != 0;
                        node.flip_v = fsp.grf_persistence & fsp_flags::FLIP_V != 0;
                        patriarch = fsp.grf_persistence & fsp_flags::PATRIARCH != 0;
                        deleted = fsp.grf_persistence & fsp_flags::DELETED != 0;
                    }
                }
                er::FSPGR => {
                    if let Ok((fspgr, _)) = OfficeArtFspgr::read_from(&sp_body[cursor..]) {
                        node.fspgr = Some(fspgr);
                    }
                }
                er::FOPT => {
                    if let Ok((table, _)) = FoptTable::read_from(&sp_body[cursor..]) {
                        for entry in table.entries() {
                            match entry.id {
                                0x0004 => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        node.rotation = Some(v as i32);
                                    }
                                }
                                0x0104 => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        node.blip_id = Some(v);
                                    }
                                }
                                0x0181 => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        node.fill_color = Some(v);
                                    }
                                }
                                0x01BF => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        if v & 0x0010_0000 != 0 {
                                            node.fill_enabled = Some(v & 0x0000_0010 != 0);
                                        }
                                    }
                                }
                                0x01C0 => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        node.line_color = Some(v);
                                    }
                                }
                                0x01CB => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        node.line_width = Some(v);
                                    }
                                }
                                0x01CE => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        node.line_dashing = Some(v);
                                    }
                                }
                                0x01FF => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        // MS-ODRAW 2.3.8.44: fLine
                                        // (0x8) displays the outline;
                                        // the line is absent when it
                                        // is explicitly cleared.
                                        if v & 0x0008_0000 != 0 {
                                            node.line_no_fill = Some(v & 0x0000_0008 == 0);
                                        }
                                    }
                                }
                                0x0380 => {
                                    if let FoptValue::Complex(bytes) = &entry.value {
                                        node.name = Some(decode_utf16le_null_terminated(bytes));
                                    }
                                }
                                0x0381 => {
                                    if let FoptValue::Complex(bytes) = &entry.value {
                                        node.alt_text = Some(decode_utf16le_null_terminated(bytes));
                                    }
                                }
                                0x03BF => {
                                    if let FoptValue::Simple(v) = entry.value {
                                        node.hidden =
                                            crate::biff::escher::group_shape_props_hidden(v);
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
                er::CLIENT_ANCHOR => {
                    if let Ok((a, _)) = OfficeArtClientAnchor::read_from(&sp_body[cursor..]) {
                        node.client_anchor = Some(a);
                    }
                }
                er::CHILD_ANCHOR => {
                    if let Ok((a, _)) = OfficeArtChildAnchor::read_from(&sp_body[cursor..]) {
                        node.child_anchor = Some(a);
                    }
                }
                er::CLIENT_DATA => node.has_client_data = true,
                _ => {}
            }
            cursor = body_end;
        }

        (has_fsp && !patriarch && !deleted).then_some(node)
    }

    /// Number of nodes in the tree that pair with an OBJ record.
    fn client_data_count(nodes: &[EscherShapeNode]) -> usize {
        nodes
            .iter()
            .map(|node| usize::from(node.has_client_data) + Self::client_data_count(&node.children))
            .sum()
    }

    /// Iteratively walk OfficeArt records in pre-order (a container
    /// is visited before its children, children before the
    /// container's next sibling), i.e. document order. A record
    /// budget prevents adversarial streams from consuming unbounded
    /// CPU/memory; an explicit frame stack avoids process-aborting
    /// recursion over deeply nested containers. The callback returns
    /// whether to descend into a container's body.
    fn walk_escher_records<'a, F>(root: &'a [u8], mut visit: F)
    where
        F: FnMut(&crate::biff::escher::OfficeArtRecordHeader, &'a [u8]) -> bool,
    {
        use crate::biff::escher::{OfficeArtRecordHeader, HEADER_LEN};
        const MAX_RECORDS: usize = 1_000_000;

        // (container body, read cursor) frames. Depth is bounded by
        // the record budget: a frame is only pushed after a visit.
        let mut stack: Vec<(&'a [u8], usize)> = vec![(root, 0)];
        let mut seen = 0usize;
        while let Some(frame) = stack.last_mut() {
            let (body, cursor) = *frame;
            if cursor.saturating_add(HEADER_LEN) > body.len() {
                stack.pop();
                continue;
            }
            if seen >= MAX_RECORDS {
                return;
            }
            seen += 1;
            let Ok(h) = OfficeArtRecordHeader::read_from(&body[cursor..]) else {
                stack.pop();
                continue;
            };
            let Some((body_start, body_end)) =
                Self::escher_record_bounds(cursor, body.len(), h.rec_len)
            else {
                stack.pop();
                continue;
            };
            frame.1 = body_end;
            let inner = &body[body_start..body_end];
            if h.is_container() && visit(&h, inner) {
                stack.push((inner, 0));
            }
        }
    }

    fn escher_record_bounds(
        cursor: usize,
        enclosing_len: usize,
        rec_len: u32,
    ) -> Option<(usize, usize)> {
        let body_start = cursor.checked_add(crate::biff::escher::HEADER_LEN)?;
        let payload_len = usize::try_from(rec_len).ok()?;
        let body_end = body_start.checked_add(payload_len)?;
        (body_end <= enclosing_len).then_some((body_start, body_end))
    }

    /// Convert an Escher client anchor into the model's two-cell
    /// drawing anchor, reversing the writer's EMU quantisation and
    /// mapping the placement flag to an `editAs` hint:
    ///   0 → None (the default: move + resize with cells — the byte
    ///       layout cannot distinguish an explicit `twoCell` from an
    ///       absent hint, and both mean the same thing)
    ///   2 → editAs="oneCell"  (move only)
    ///   3 → editAs="absolute" (no move, no resize)
    /// OneCell and Absolute inputs collapse to TwoCell anchors on
    /// read since the byte layout is identical; the editAs hint
    /// preserves the semantic intent so a downstream XLSX writer can
    /// re-emit the appropriate variant.
    fn client_anchor_to_drawing_anchor(
        anchor: &crate::biff::escher::OfficeArtClientAnchor,
        metrics: &dyn duke_sheets_chart::DrawingMetrics,
    ) -> duke_sheets_chart::DrawingAnchor {
        let fraction_to_emu = |units: i16, extent: i64, denominator: i128| -> i64 {
            if extent <= 0 {
                return 0;
            }
            let numerator = i128::from(units) * i128::from(extent);
            let rounded = if numerator >= 0 {
                (numerator + denominator / 2) / denominator
            } else {
                (numerator - denominator / 2) / denominator
            };
            rounded.clamp(i128::from(i64::MIN), i128::from(i64::MAX)) as i64
        };
        let edit_as = match anchor.flag {
            2 => Some(duke_sheets_chart::EditAs::OneCell),
            3 => Some(duke_sheets_chart::EditAs::Absolute),
            _ => None,
        };
        duke_sheets_chart::DrawingAnchor::TwoCell {
            from: duke_sheets_chart::CellMarker {
                col: anchor.col_l,
                col_offset_emu: fraction_to_emu(
                    anchor.dx_l,
                    metrics.column_width_emu(anchor.col_l),
                    1024,
                ),
                row: anchor.row_t as u32,
                row_offset_emu: fraction_to_emu(
                    anchor.dy_t,
                    metrics.row_height_emu(u32::from(anchor.row_t)),
                    256,
                ),
            },
            to: duke_sheets_chart::CellMarker {
                col: anchor.col_r,
                col_offset_emu: fraction_to_emu(
                    anchor.dx_r,
                    metrics.column_width_emu(anchor.col_r),
                    1024,
                ),
                row: anchor.row_b as u32,
                row_offset_emu: fraction_to_emu(
                    anchor.dy_b,
                    metrics.row_height_emu(u32::from(anchor.row_b)),
                    256,
                ),
            },
            edit_as,
        }
    }

    /// Assemble the worksheet's drawing list from the Escher shape
    /// tree and the sheet's OBJ / TXO / NOTE records.
    ///
    /// The OfficeArt container order is the z-order, so the list is
    /// built in one pre-order pass across all shape kinds: pictures,
    /// comment boxes, form controls, and groups appear at their
    /// container positions. The Nth ClientData-bearing shape pairs
    /// with the Nth OBJ record; auxiliary UI shapes (autofilter /
    /// data-validation dropdowns) consume their OBJ slot without
    /// producing a drawing object, keeping later pairings intact.
    ///
    /// Permissive: when the shape count does not match the OBJ
    /// count, positional pairing is untrustworthy and the reader
    /// degrades to the kind-by-kind extraction (pictures with their
    /// own anchors, comments from NOTE records, controls with
    /// default anchors) rather than failing the sheet load.
    fn build_sheet_drawings(
        escher_bytes: &[u8],
        obj_bodies: &[Vec<u8>],
        obj_texts: &std::collections::HashMap<u16, duke_sheets_core::ControlText>,
        notes: &[NoteData],
        blip_store: &[Option<BlipData>],
        formula_ctx: &FormulaContext,
        ws: &mut duke_sheets_core::Worksheet,
    ) {
        use crate::biff::obj;

        let nodes = Self::parse_shape_tree(escher_bytes);
        let expected_objs = Self::client_data_count(&nodes);
        let aligned = expected_objs == obj_bodies.len();

        let mut note_used = vec![false; notes.len()];
        if aligned {
            let mut next_obj = 0usize;
            for node in &nodes {
                let mut hoisted = Vec::new();
                if let Some(object) = Self::drawing_from_node(
                    node,
                    obj_bodies,
                    &mut next_obj,
                    obj_texts,
                    notes,
                    &mut note_used,
                    blip_store,
                    formula_ctx,
                    &mut hoisted,
                    ws,
                ) {
                    ws.add_drawing(object);
                }
                for object in hoisted {
                    ws.add_drawing(object);
                }
            }
        } else {
            log::warn!(
                "escher shape count ({expected_objs}) does not match OBJ count ({}); \
                 drawing objects will use default pairing",
                obj_bodies.len()
            );
            // Pictures keep their own container anchors (they do not
            // need OBJ pairing).
            let mut flat = Vec::new();
            fn flatten<'a>(nodes: &'a [EscherShapeNode], out: &mut Vec<&'a EscherShapeNode>) {
                for node in nodes {
                    out.push(node);
                    flatten(&node.children, out);
                }
            }
            flatten(&nodes, &mut flat);
            for node in flat {
                if let Some(payload) = Self::image_payload_from_node(node, blip_store) {
                    let object = Self::top_level_image(node, payload, None, ws);
                    ws.add_drawing(object);
                }
            }
            // Comments straight from their NOTE records.
            for (i, note) in notes.iter().enumerate() {
                note_used[i] = true;
                Self::add_note_comment(note, obj_texts, ws);
            }
            // Controls from OBJ bodies with default anchors.
            for body in obj_bodies {
                let Ok(parsed) = obj::parse_obj(body) else {
                    continue;
                };
                if let Some(control) = Self::control_from_obj(&parsed, None, obj_texts, formula_ctx)
                {
                    let mut object = duke_sheets_core::DrawingObject::form_control(control);
                    object.meta.locked = parsed.grbit & obj::cmo_flags::LOCKED != 0;
                    object.meta.printable = parsed.grbit & obj::cmo_flags::PRINT != 0;
                    ws.add_drawing(object);
                }
            }
        }

        // NOTE records not claimed by a comment shape still surface
        // as comments (files with no drawing stream, or malformed
        // pairing).
        for (i, note) in notes.iter().enumerate() {
            if !note_used[i] {
                Self::add_note_comment(note, obj_texts, ws);
            }
        }
    }

    fn add_note_comment(
        note: &NoteData,
        obj_texts: &std::collections::HashMap<u16, duke_sheets_core::ControlText>,
        ws: &mut duke_sheets_core::Worksheet,
    ) {
        let text = obj_texts
            .get(&note.obj_id)
            .map(duke_sheets_core::ControlText::plain_text)
            .unwrap_or_default();
        // Permissive read: a NOTE pointing outside the model grid is
        // dropped rather than failing the sheet load.
        if ws
            .set_comment_at(
                note.row,
                note.col,
                CellComment::new(note.author.clone(), text),
            )
            .is_ok()
        {
            ws.set_comment_visible(note.row, note.col, note.visible);
        }
    }

    /// Build the drawing object for one top-level shape node,
    /// consuming its (and its children's) OBJ slots. Returns `None`
    /// when the node has no model representation (auxiliary UI
    /// dropdowns, unmodeled shape kinds); the OBJ slots are consumed
    /// regardless so later pairings stay intact. Comment shapes found
    /// inside groups are appended to `hoisted` as top-level objects.
    #[allow(clippy::too_many_arguments)]
    fn drawing_from_node(
        node: &EscherShapeNode,
        obj_bodies: &[Vec<u8>],
        next_obj: &mut usize,
        obj_texts: &std::collections::HashMap<u16, duke_sheets_core::ControlText>,
        notes: &[NoteData],
        note_used: &mut [bool],
        blip_store: &[Option<BlipData>],
        formula_ctx: &FormulaContext,
        hoisted: &mut Vec<duke_sheets_core::DrawingObject>,
        metrics: &dyn duke_sheets_chart::DrawingMetrics,
    ) -> Option<duke_sheets_core::DrawingObject> {
        use crate::biff::obj;

        let parsed = Self::consume_obj(node, obj_bodies, next_obj);

        if node.is_group {
            let group = Self::group_from_node(
                node,
                obj_bodies,
                next_obj,
                obj_texts,
                notes,
                note_used,
                blip_store,
                formula_ctx,
                hoisted,
                metrics,
            );
            let anchor = node
                .client_anchor
                .as_ref()
                .map(|anchor| Self::client_anchor_to_drawing_anchor(anchor, metrics))
                .unwrap_or_default();
            let mut object = duke_sheets_core::DrawingObject::group(group);
            object.anchor = anchor;
            object.meta = Self::node_meta(node, parsed.as_ref());
            return Some(object);
        }

        if let Some(payload) = Self::image_payload_from_node(node, blip_store) {
            return Some(Self::top_level_image(node, payload, parsed.as_ref(), metrics));
        }

        let parsed = parsed?;
        if parsed.ot == obj::ot::NOTE {
            let (index, note) = notes
                .iter()
                .enumerate()
                .find(|(i, note)| !note_used[*i] && note.obj_id == parsed.id)?;
            note_used[index] = true;
            let text = obj_texts
                .get(&note.obj_id)
                .map(duke_sheets_core::ControlText::plain_text)
                .unwrap_or_default();
            let mut object = duke_sheets_core::DrawingObject::comment(
                note.row,
                note.col,
                CellComment::new(note.author.clone(), text),
            );
            if let Some(anchor) = &node.client_anchor {
                object.anchor = Self::client_anchor_to_drawing_anchor(anchor, metrics);
            }
            object.meta.hidden = !note.visible;
            return Some(object);
        }

        if Self::is_shape_obj_type(parsed.ot) {
            let shape = Self::shape_from_node(node, &parsed, obj_texts)?;
            let anchor = node
                .client_anchor
                .as_ref()
                .map(|anchor| Self::client_anchor_to_drawing_anchor(anchor, metrics))
                .unwrap_or_default();
            let mut object = duke_sheets_core::DrawingObject::shape(shape).with_anchor(anchor);
            object.meta = Self::node_meta(node, Some(&parsed));
            return Some(object);
        }

        let control =
            Self::control_from_obj(&parsed, Some(node.shape_type), obj_texts, formula_ctx)?;
        let anchor = node
            .client_anchor
            .as_ref()
            .map(|anchor| Self::client_anchor_to_drawing_anchor(anchor, metrics))
            .unwrap_or_default();
        let mut object = duke_sheets_core::DrawingObject::form_control(control).with_anchor(anchor);
        object.meta = Self::node_meta(node, Some(&parsed));
        Some(object)
    }

    /// Build a [`duke_sheets_core::Group`] from a group node,
    /// consuming the children's OBJ slots. Children without a model
    /// representation are dropped from the group (their OBJ slots are
    /// still consumed); comment children are hoisted to top level.
    ///
    /// Group-space geometry note: the FSPGR rectangle and each
    /// child's `OfficeArtChildAnchor` share one unit-agnostic
    /// coordinate space — rendering scales the child rectangles from
    /// the FSPGR rect onto the group's sheet anchor. The raw values
    /// are stored in the model's `*_emu` fields unconverted; the
    /// child-to-parent scaling normalizes whatever unit they are in.
    #[allow(clippy::too_many_arguments)]
    fn group_from_node(
        node: &EscherShapeNode,
        obj_bodies: &[Vec<u8>],
        next_obj: &mut usize,
        obj_texts: &std::collections::HashMap<u16, duke_sheets_core::ControlText>,
        notes: &[NoteData],
        note_used: &mut [bool],
        blip_store: &[Option<BlipData>],
        formula_ctx: &FormulaContext,
        hoisted: &mut Vec<duke_sheets_core::DrawingObject>,
        metrics: &dyn duke_sheets_chart::DrawingMetrics,
    ) -> duke_sheets_core::Group {
        use duke_sheets_core::{DrawingKind, GroupChild, GroupTransform};

        let fspgr = node.fspgr.unwrap_or_default();
        // The group's own placement mirrors its sheet anchor in the
        // worksheet's metric-aware EMU space (or the parent child
        // space for nested groups), matching what the writer emits.
        let (x_emu, y_emu, cx_emu, cy_emu) = if let Some(anchor) = &node.client_anchor {
            let (x1, y1, x2, y2) = Self::client_anchor_rect_emu(anchor, metrics);
            (x1, y1, (x2 - x1).max(0), (y2 - y1).max(0))
        } else if let Some(child) = &node.child_anchor {
            (
                i64::from(child.x_left),
                i64::from(child.y_top),
                i64::from(child.x_right - child.x_left).max(0),
                i64::from(child.y_bottom - child.y_top).max(0),
            )
        } else {
            (0, 0, 0, 0)
        };
        let transform = GroupTransform {
            x_emu,
            y_emu,
            cx_emu,
            cy_emu,
            child_x_emu: i64::from(fspgr.x_left),
            child_y_emu: i64::from(fspgr.y_top),
            child_cx_emu: i64::from(fspgr.x_right - fspgr.x_left),
            child_cy_emu: i64::from(fspgr.y_bottom - fspgr.y_top),
            rotation: node.rotation.map(officeart_fixed_to_rotation).unwrap_or(0),
            flip_h: node.flip_h,
            flip_v: node.flip_v,
        };

        let mut children = Vec::new();
        for child in &node.children {
            if child.is_group {
                let parsed = Self::consume_obj(child, obj_bodies, next_obj);
                let inner = Self::group_from_node(
                    child,
                    obj_bodies,
                    next_obj,
                    obj_texts,
                    notes,
                    note_used,
                    blip_store,
                    formula_ctx,
                    hoisted,
                    metrics,
                );
                children.push(GroupChild {
                    meta: Self::node_meta(child, parsed.as_ref()),
                    transform: Self::child_transform(child),
                    kind: DrawingKind::Group(Box::new(inner)),
                });
                continue;
            }

            let parsed = Self::consume_obj(child, obj_bodies, next_obj);
            if let Some(mut payload) = Self::image_payload_from_node(child, blip_store) {
                // The child transform is authoritative for grouped
                // shapes; the payload keeps neutral placement fields,
                // mirroring the XLSX reader.
                payload.rotation = None;
                payload.flip_h = false;
                payload.flip_v = false;
                if let Some(anchor) = &child.child_anchor {
                    payload.width_emu = i64::from(anchor.x_right - anchor.x_left).max(0);
                    payload.height_emu = i64::from(anchor.y_bottom - anchor.y_top).max(0);
                }
                let mut meta = Self::node_meta(child, parsed.as_ref());
                meta.name = Some(
                    child
                        .name
                        .clone()
                        .unwrap_or_else(|| format!("Picture {}", child.spid)),
                );
                children.push(GroupChild {
                    meta,
                    transform: Self::child_transform(child),
                    kind: DrawingKind::Image(payload),
                });
                continue;
            }
            let Some(parsed) = parsed else { continue };
            if parsed.ot == crate::biff::obj::ot::NOTE {
                // Comments are never grouped in practice; surface one
                // as a top-level object rather than losing it.
                if let Some((index, note)) = notes
                    .iter()
                    .enumerate()
                    .find(|(i, note)| !note_used[*i] && note.obj_id == parsed.id)
                {
                    note_used[index] = true;
                    let text = obj_texts
                        .get(&note.obj_id)
                        .map(duke_sheets_core::ControlText::plain_text)
                        .unwrap_or_default();
                    let mut object = duke_sheets_core::DrawingObject::comment(
                        note.row,
                        note.col,
                        CellComment::new(note.author.clone(), text),
                    );
                    object.meta.hidden = !note.visible;
                    hoisted.push(object);
                }
                continue;
            }
            if Self::is_shape_obj_type(parsed.ot) {
                if let Some(shape) = Self::shape_from_node(child, &parsed, obj_texts) {
                    let mut transform = Self::child_transform(child);
                    transform.rotation = shape.rotation;
                    children.push(GroupChild {
                        meta: Self::node_meta(child, Some(&parsed)),
                        transform,
                        kind: DrawingKind::Shape(Box::new(shape)),
                    });
                }
                continue;
            }
            if let Some(control) =
                Self::control_from_obj(&parsed, Some(child.shape_type), obj_texts, formula_ctx)
            {
                children.push(GroupChild {
                    meta: Self::node_meta(child, Some(&parsed)),
                    transform: Self::child_transform(child),
                    kind: DrawingKind::FormControl(control),
                });
            }
        }

        duke_sheets_core::Group {
            transform,
            children,
        }
    }

    /// Convert a supported OfficeArt FSP + paired ftCmo(ot=0x001E)
    /// into the public shape model. Unknown MSOSPT values are dropped
    /// rather than being mislabeled as rectangles.
    fn is_shape_obj_type(object_type: u16) -> bool {
        use crate::biff::obj::ot;

        matches!(
            object_type,
            ot::LINE
                | ot::RECTANGLE
                | ot::OVAL
                | ot::ARC
                | ot::TEXT
                | ot::POLYGON
                | ot::OFFICE_ART
        )
    }

    fn shape_from_node(
        node: &EscherShapeNode,
        parsed: &crate::biff::obj::ParsedObj,
        obj_texts: &std::collections::HashMap<u16, duke_sheets_core::ControlText>,
    ) -> Option<duke_sheets_core::Shape> {
        use crate::biff::escher::shape_type;
        use duke_sheets_core::{Shape, ShapeFill, ShapeLine};

        let preset = match node.shape_type {
            shape_type::RECTANGLE => "rect",
            shape_type::ROUND_RECTANGLE => "roundRect",
            shape_type::ELLIPSE => "ellipse",
            shape_type::ISOSCELES_TRIANGLE => "triangle",
            shape_type::LINE => "line",
            _ => return None,
        };
        let fill = match node.fill_enabled {
            Some(false) => ShapeFill::None,
            Some(true) => node
                .fill_color
                .and_then(officeart_color_to_core)
                .map(ShapeFill::Solid)
                .unwrap_or(ShapeFill::Solid(duke_sheets_core::Color::Auto)),
            None => ShapeFill::None,
        };
        let mut shape = Shape::preset(preset);
        shape.fill = fill;
        shape.line = ShapeLine {
            color: node.line_color.and_then(officeart_color_to_core),
            width_emu: node.line_width.map(i64::from),
            dash_style: node.line_dashing.map(officeart_dash_to_drawing),
            no_fill: node.line_no_fill.unwrap_or(false),
        };
        // The writer emits Left/Top TXO flags for alignment-less
        // shape text; strip those defaults back to None (mirroring
        // the control caption path) so defaults round-trip as None.
        shape.text = obj_texts.get(&parsed.id).cloned().map(|mut text| {
            if text.horizontal_alignment == Some(duke_sheets_core::HorizontalAlignment::Left) {
                text.horizontal_alignment = None;
            }
            if text.vertical_alignment == Some(duke_sheets_core::VerticalAlignment::Top) {
                text.vertical_alignment = None;
            }
            text
        });
        shape.rotation = node.rotation.map(officeart_fixed_to_rotation).unwrap_or(0);
        shape.flip_h = node.flip_h;
        shape.flip_v = node.flip_v;
        Some(shape)
    }

    /// Consume the node's OBJ slot, if it has one. A malformed OBJ
    /// body still consumes its slot.
    fn consume_obj(
        node: &EscherShapeNode,
        obj_bodies: &[Vec<u8>],
        next_obj: &mut usize,
    ) -> Option<crate::biff::obj::ParsedObj> {
        if !node.has_client_data {
            return None;
        }
        let body = obj_bodies.get(*next_obj)?;
        *next_obj += 1;
        crate::biff::obj::parse_obj(body).ok()
    }

    fn node_meta(
        node: &EscherShapeNode,
        parsed: Option<&crate::biff::obj::ParsedObj>,
    ) -> duke_sheets_core::DrawingMeta {
        use crate::biff::obj::cmo_flags;
        duke_sheets_core::DrawingMeta {
            name: node.name.clone(),
            alt_text: node.alt_text.clone(),
            hidden: node.hidden,
            locked: parsed.is_none_or(|p| p.grbit & cmo_flags::LOCKED != 0),
            printable: parsed.is_none_or(|p| p.grbit & cmo_flags::PRINT != 0),
            ..duke_sheets_core::DrawingMeta::default()
        }
    }

    fn child_transform(node: &EscherShapeNode) -> duke_sheets_core::ChildTransform {
        let anchor = node.child_anchor.unwrap_or_default();
        duke_sheets_core::ChildTransform {
            x_emu: i64::from(anchor.x_left),
            y_emu: i64::from(anchor.y_top),
            cx_emu: i64::from(anchor.x_right - anchor.x_left).max(0),
            cy_emu: i64::from(anchor.y_bottom - anchor.y_top).max(0),
            rotation: node.rotation.map(officeart_fixed_to_rotation).unwrap_or(0),
            flip_h: node.flip_h,
            flip_v: node.flip_v,
        }
    }

    /// Build the image payload for a picture-frame node whose blip
    /// resolves, with placement fields (rotation/flips) from the
    /// shape and a zero extent for the caller to fill.
    fn image_payload_from_node(
        node: &EscherShapeNode,
        blip_store: &[Option<BlipData>],
    ) -> Option<duke_sheets_chart::EmbeddedImage> {
        use crate::biff::escher::shape_type;
        if node.shape_type != shape_type::PICTURE_FRAME {
            return None;
        }
        let idx = node.blip_id?.saturating_sub(1) as usize;
        let blip = blip_store.get(idx)?.as_ref()?;
        Some(duke_sheets_chart::EmbeddedImage {
            format: blip.format,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 0,
            height_emu: 0,
            rotation: node.rotation.map(officeart_fixed_to_rotation),
            flip_h: node.flip_h,
            flip_v: node.flip_v,
            data: blip.data.clone(),
            svg_data: None,
        })
    }

    /// Wrap an image payload into a top-level drawing object,
    /// synthesising the extent from the anchored cell range and the
    /// worksheet's row and column metrics. Excel does not store a
    /// separate absolute EMU dimension on XLS pictures.
    fn top_level_image(
        node: &EscherShapeNode,
        mut payload: duke_sheets_chart::EmbeddedImage,
        parsed: Option<&crate::biff::obj::ParsedObj>,
        metrics: &dyn duke_sheets_chart::DrawingMetrics,
    ) -> duke_sheets_core::DrawingObject {
        let anchor = node.client_anchor.unwrap_or_default();
        let (x1, y1, x2, y2) = Self::client_anchor_rect_emu(&anchor, metrics);
        payload.width_emu = (x2 - x1).max(0);
        payload.height_emu = (y2 - y1).max(0);

        let mut object = duke_sheets_core::DrawingObject::image(payload)
            .with_anchor(Self::client_anchor_to_drawing_anchor(&anchor, metrics));
        object.meta = Self::node_meta(node, parsed);
        object.meta.name = Some(
            node.name
                .clone()
                .unwrap_or_else(|| format!("Picture {}", node.spid)),
        );
        object
    }

    /// Absolute EMU rectangle of a client anchor using worksheet metrics.
    fn client_anchor_rect_emu(
        anchor: &crate::biff::escher::OfficeArtClientAnchor,
        metrics: &dyn duke_sheets_chart::DrawingMetrics,
    ) -> (i64, i64, i64, i64) {
        let duke_sheets_chart::DrawingAnchor::TwoCell { from, to, .. } =
            Self::client_anchor_to_drawing_anchor(anchor, metrics)
        else {
            unreachable!("client anchors always convert to TwoCell")
        };
        let (x1, y1) = duke_sheets_chart::marker_position_emu(&from, metrics);
        let (x2, y2) = duke_sheets_chart::marker_position_emu(&to, metrics);
        let clamp = |value: i128| value.clamp(i128::from(i64::MIN), i128::from(i64::MAX)) as i64;
        (clamp(x1), clamp(y1), clamp(x2), clamp(y2))
    }

    /// Build a [`duke_sheets_core::FormControl`] from a parsed OBJ
    /// body. Returns `None` for non-control object types, for
    /// auxiliary UI dropdowns, and (when the paired shape type is
    /// known) for OBJ/shape kind mismatches.
    fn control_from_obj(
        parsed: &crate::biff::obj::ParsedObj,
        paired_shape_type: Option<u16>,
        obj_texts: &std::collections::HashMap<u16, duke_sheets_core::ControlText>,
        formula_ctx: &FormulaContext,
    ) -> Option<duke_sheets_core::FormControl> {
        use crate::biff::escher::shape_type;
        use crate::biff::obj::{self, ot};
        use duke_sheets_core::{CheckState, FormControl, FormControlKind, ListSelection};

        let is_control = matches!(
            parsed.ot,
            ot::BUTTON
                | ot::CHECKBOX
                | ot::OPTION_BUTTON
                | ot::LABEL
                | ot::GROUP_BOX
                | ot::LIST_BOX
                | ot::DROPDOWN
                | ot::SCROLLBAR
                | ot::SPINNER
                | ot::EDIT_BOX
                | ot::DIALOG_BOX
        );
        if !is_control {
            return None;
        }
        if matches!(parsed.ot, ot::LIST_BOX | ot::DROPDOWN)
            && (parsed.lbs_malformed || parsed.lbs.is_none())
        {
            return None;
        }
        // Excel persists auxiliary UI dropdowns (one ot=0x14 OBJ per
        // autofilter column, and similar pivot/table-total dropdowns)
        // that are not user Forms controls. They are marked with
        // fUIObj in ftCmo and a non-regular lct behavior class in
        // ftLbsData; skip both signals.
        if parsed.grbit & obj::cmo_flags::UI_OBJ != 0 {
            return None;
        }
        if let Some(lbs) = &parsed.lbs {
            if lbs.use_cb && lbs.lct != 0 {
                return None;
            }
        }
        if let Some(st) = paired_shape_type {
            if st != shape_type::HOST_CONTROL {
                // OBJ says control but the paired shape isn't a host
                // control: the pairing is off for this entry, skip it.
                return None;
            }
        }

        let decompile_rgce = |rgce: &Option<Vec<u8>>| -> Option<String> {
            let rgce = rgce.as_ref()?;
            if rgce.is_empty() {
                return None;
            }
            let text = crate::biff::formula::decompile(rgce, formula_ctx);
            if text.is_empty() {
                None
            } else {
                Some(text)
            }
        };

        let caption = || {
            let mut caption = obj_texts.get(&parsed.id).cloned().unwrap_or_default();
            let (default_horizontal, default_vertical) = match parsed.ot {
                ot::BUTTON => (
                    duke_sheets_core::HorizontalAlignment::Center,
                    duke_sheets_core::VerticalAlignment::Center,
                ),
                ot::CHECKBOX | ot::OPTION_BUTTON => (
                    duke_sheets_core::HorizontalAlignment::Left,
                    duke_sheets_core::VerticalAlignment::Center,
                ),
                _ => (
                    duke_sheets_core::HorizontalAlignment::Left,
                    duke_sheets_core::VerticalAlignment::Top,
                ),
            };
            if caption.horizontal_alignment == Some(default_horizontal) {
                caption.horizontal_alignment = None;
            }
            if caption.vertical_alignment == Some(default_vertical) {
                caption.vertical_alignment = None;
            }
            caption
        };
        let state = match parsed.checked {
            Some(2) => CheckState::Mixed,
            Some(v) if v != 0 => CheckState::Checked,
            _ => CheckState::Unchecked,
        };
        let cell_link = decompile_rgce(&parsed.link_rgce);

        let kind = match parsed.ot {
            ot::BUTTON => FormControlKind::Button { caption: caption() },
            ot::CHECKBOX => FormControlKind::Checkbox {
                caption: caption(),
                state,
                cell_link,
                no_3d: parsed.cbls_no_3d,
            },
            ot::OPTION_BUTTON => FormControlKind::OptionButton {
                caption: caption(),
                // Mixed is checkbox-only (MS-XLS 2.5.141); clamp
                // out-of-spec radio states to Checked.
                state: if state == CheckState::Mixed {
                    CheckState::Checked
                } else {
                    state
                },
                cell_link,
                first_in_group: parsed.radio.map(|(_, first)| first).unwrap_or(false),
                no_3d: parsed.cbls_no_3d,
            },
            ot::LABEL => FormControlKind::Label { caption: caption() },
            ot::GROUP_BOX => FormControlKind::GroupBox {
                caption: caption(),
                no_3d: parsed.gbo_no_3d.unwrap_or(false),
            },
            ot::LIST_BOX => {
                let lbs = parsed.lbs.clone().unwrap_or_default();
                let selection = match lbs.sel_type {
                    1 => ListSelection::Multi,
                    2 => ListSelection::Extend,
                    _ => ListSelection::Single,
                };
                // iSel is one-based (0 = none); the model is
                // zero-based. bsels positions are already 0-based.
                let selected: Vec<u16> = if lbs.sel_type == 0 {
                    if lbs.sel > 0 {
                        vec![lbs.sel - 1]
                    } else {
                        Vec::new()
                    }
                } else {
                    lbs.multi_sel
                        .iter()
                        .enumerate()
                        .filter(|(_, &s)| s)
                        .map(|(idx, _)| idx as u16)
                        .collect()
                };
                FormControlKind::ListBox {
                    input_range: decompile_rgce(&Some(lbs.input_rgce)),
                    cell_link,
                    selection,
                    selected,
                    no_3d: lbs.no_3d,
                }
            }
            ot::DROPDOWN => {
                let lbs = parsed.lbs.clone().unwrap_or_default();
                FormControlKind::Dropdown {
                    input_range: decompile_rgce(&Some(lbs.input_rgce)),
                    cell_link,
                    selected: if lbs.sel > 0 { Some(lbs.sel - 1) } else { None },
                    lines: lbs.drop.as_ref().map(|d| d.lines).unwrap_or(8),
                    no_3d: lbs.no_3d,
                }
            }
            ot::SCROLLBAR | ot::SPINNER => {
                let sbs = parsed.sbs.unwrap_or_default();
                let clamp = |v: i16| v.max(0) as u16;
                if parsed.ot == ot::SCROLLBAR {
                    FormControlKind::Scrollbar {
                        value: clamp(sbs.val),
                        min: clamp(sbs.min),
                        max: clamp(sbs.max),
                        increment: clamp(sbs.inc),
                        page: clamp(sbs.page),
                        horizontal: sbs.horizontal,
                        cell_link,
                    }
                } else {
                    FormControlKind::Spinner {
                        value: clamp(sbs.val),
                        min: clamp(sbs.min),
                        max: clamp(sbs.max),
                        increment: clamp(sbs.inc),
                        cell_link,
                    }
                }
            }
            ot::EDIT_BOX | ot::DIALOG_BOX => FormControlKind::Unknown {
                object_type: if parsed.ot == ot::EDIT_BOX {
                    "EditBox".to_string()
                } else {
                    "Dialog".to_string()
                },
                legacy_object_type: Some(parsed.ot),
                caption: caption(),
                raw_properties: Vec::new(),
                raw_obj: Some(parsed.raw_body.clone()),
            },
            _ => unreachable!(),
        };

        let mut control = FormControl::new(kind);
        control.macro_name = parsed.macro_rgce.as_ref().and_then(|rgce| {
            let name = crate::biff::formula::decompile(rgce, formula_ctx);
            let name = name.strip_prefix('=').unwrap_or(&name).trim();
            (!name.is_empty()).then(|| name.to_string())
        });
        Some(control)
    }

    // ── Drawing record parsers ───────────────────────────────────────────

    /// Walk a `MSODRAWINGGROUP` record's body to extract every
    /// `OfficeArtBlip*` payload from its embedded `BSTORE_CONTAINER`,
    /// appending one [`BlipData`] entry per image into `blip_store`.
    ///
    /// Failures inside the Escher tree are swallowed and skipped —
    /// the reader is intentionally permissive so a malformed drawing
    /// group cannot prevent reading the rest of the workbook.
    fn parse_msodrawinggroup(data: &[u8], blip_store: &mut Vec<Option<BlipData>>) {
        use crate::biff::escher::rec_type as er;
        Self::walk_escher_records(data, |h, body| {
            if h.rec_type == er::BSTORE_CONTAINER {
                Self::parse_bstore_container(body, blip_store);
                false
            } else {
                true
            }
        });
    }

    /// Walk a `BSTORE_CONTAINER` body, decoding each `FBSE` child to
    /// extract its embedded blip's image bytes + format.
    fn parse_bstore_container(body: &[u8], blip_store: &mut Vec<Option<BlipData>>) {
        use crate::biff::escher::{rec_type as er, OfficeArtRecordHeader, HEADER_LEN};
        let mut cursor = 0;
        while cursor + HEADER_LEN <= body.len() {
            let Ok(h) = OfficeArtRecordHeader::read_from(&body[cursor..]) else {
                return;
            };
            let Some((entry_start, entry_end)) =
                Self::escher_record_bounds(cursor, body.len(), h.rec_len)
            else {
                return;
            };
            if h.rec_type == er::FBSE {
                // Placeholder entries (no embedded blip) must keep
                // their slot: pib references are 1-based positional
                // indices over every FBSE, and Excel emits header-only
                // entries when it rasterizes transforms on save.
                blip_store.push(Self::parse_fbse_entry(&body[entry_start..entry_end]));
            }
            cursor = entry_end;
        }
    }

    /// Parse an `OFFICEARTFBSE` body and extract its embedded blip.
    /// Returns `None` for malformed entries or formats we don't yet
    /// recognise.
    ///
    /// Body layout (MS-ODRAW §2.2.32): 36 fixed bytes (btWin32,
    /// btMacOS, rgbUid, tag, size, cRef, foDelay, usage, cbName,
    /// unused2, unused3), then optional `nameData` (cbName bytes),
    /// then the embedded blip record.
    fn parse_fbse_entry(body: &[u8]) -> Option<BlipData> {
        use crate::biff::escher::{rec_type as er, OfficeArtRecordHeader, HEADER_LEN};
        if body.len() < 36 {
            return None;
        }
        let cb_name = body[33] as usize;
        let blip_start = 36usize.checked_add(cb_name)?;
        if blip_start.checked_add(HEADER_LEN)? > body.len() {
            return None;
        }
        let h = OfficeArtRecordHeader::read_from(&body[blip_start..]).ok()?;
        let (blip_body_start, blip_body_end) =
            Self::escher_record_bounds(blip_start, body.len(), h.rec_len)?;
        let format = match h.rec_type {
            er::BLIP_PNG => duke_sheets_chart::ImageFormat::Png,
            er::BLIP_JPEG => duke_sheets_chart::ImageFormat::Jpeg,
            er::BLIP_DIB => duke_sheets_chart::ImageFormat::Bmp,
            er::BLIP_EMF => duke_sheets_chart::ImageFormat::Emf,
            er::BLIP_WMF => duke_sheets_chart::ImageFormat::Wmf,
            er::BLIP_TIFF => duke_sheets_chart::ImageFormat::Tiff,
            _ => return None,
        };
        // Blip body layout depends on format:
        //   PNG/JPEG/DIB (raster): rgbUid (16) [+ rgbUidPrimary (16)] +
        //     tag (1) + image bytes.
        //   EMF/WMF (metafile): rgbUid (16) [+ rgbUidPrimary (16)] +
        //     metafileHeader (34) + (possibly compressed) data bytes.
        // The instance's low bit selects the secondary-UID variant.
        let has_secondary_uid = (h.rec_instance & 0x0001) != 0;
        let is_metafile = matches!(h.rec_type, er::BLIP_EMF | er::BLIP_WMF);
        let suffix_len = if is_metafile { 34 } else { 1 };
        let header_inner = if has_secondary_uid {
            16 + 16 + suffix_len
        } else {
            16 + suffix_len
        };
        if header_inner > body.len() - blip_body_start {
            return None;
        }
        let data_start = blip_body_start + header_inner;
        let image_bytes = body[data_start..blip_body_end].to_vec();

        // For DIB blips, prepend a synthesised 14-byte
        // BITMAPFILEHEADER so the result is a complete BMP file the
        // downstream consumer (and Excel itself) can decode.
        let final_bytes = if matches!(format, duke_sheets_chart::ImageFormat::Bmp) {
            bmp_file_header_for_dib(&image_bytes)
                .into_iter()
                .chain(image_bytes)
                .collect()
        } else {
            image_bytes
        };

        Some(BlipData {
            format,
            data: final_bytes,
        })
    }

    // ── Comment record parsers (OBJ → TXO → NOTE) ───────────────────────

    /// Extract the object ID from an OBJ record.
    ///
    /// The OBJ record contains sub-records. The first sub-record is ftCmo
    /// (common object data): rt(2) + cb(2) + ot(2) + id(2) + flags(2) ...
    /// We only need the `id` field at offset 6.
    fn parse_obj_id(data: &[u8]) -> Option<u16> {
        if data.len() < 8 {
            return None;
        }
        let rt = u16::from_le_bytes([data[0], data[1]]);
        if rt != 0x0015 {
            // ftCmo sub-record type must be 0x0015
            return None;
        }
        Some(u16::from_le_bytes([data[6], data[7]]))
    }

    /// Extract text, formatting runs, and alignment from a TXO record.
    ///
    /// TXO header (18 bytes): options(2) + rotation(2) + reserved(6) +
    /// text_len(2) + format_run_size(2) + reserved(4).
    /// First CONTINUE: grbit(1) + text_data.
    /// Second CONTINUE: formatting runs (`ich`, `ifnt`, reserved).
    fn parse_txo_text(
        data: &[u8],
        continue_offsets: &[usize],
        style_ctx: &StyleContext,
    ) -> Option<duke_sheets_core::ControlText> {
        if data.len() < 18 {
            return None;
        }
        let flags = u16::from_le_bytes([data[0], data[1]]);
        let horizontal_alignment = match (flags >> 1) & 0x7 {
            1 => Some(duke_sheets_core::HorizontalAlignment::Left),
            2 => Some(duke_sheets_core::HorizontalAlignment::Center),
            3 => Some(duke_sheets_core::HorizontalAlignment::Right),
            4 => Some(duke_sheets_core::HorizontalAlignment::Justify),
            7 => Some(duke_sheets_core::HorizontalAlignment::Distributed),
            _ => None,
        };
        let vertical_alignment = match (flags >> 4) & 0x7 {
            1 => Some(duke_sheets_core::VerticalAlignment::Top),
            2 => Some(duke_sheets_core::VerticalAlignment::Center),
            3 => Some(duke_sheets_core::VerticalAlignment::Bottom),
            4 => Some(duke_sheets_core::VerticalAlignment::Justify),
            7 => Some(duke_sheets_core::VerticalAlignment::Distributed),
            _ => None,
        };
        let text_len = u16::from_le_bytes([data[10], data[11]]) as usize;
        if text_len == 0 {
            return Some(duke_sheets_core::ControlText {
                runs: Vec::new(),
                horizontal_alignment,
                vertical_alignment,
            });
        }

        // Text starts in the first CONTINUE block. Every subsequent
        // text continuation starts with its own encoding grbit; the
        // final `cbRuns` bytes are formatting data, not text.
        let text_start = if !continue_offsets.is_empty() {
            continue_offsets[0]
        } else if data.len() > 18 {
            18 // No CONTINUE marker - text follows header directly
        } else {
            return None;
        };

        if text_start >= data.len() {
            return None;
        }

        let cb_runs = u16::from_le_bytes([data[12], data[13]]) as usize;
        let text_end = data.len().saturating_sub(cb_runs).max(text_start);
        let mut starts: Vec<usize> = if continue_offsets.is_empty() {
            vec![text_start]
        } else {
            continue_offsets
                .iter()
                .copied()
                .filter(|&offset| offset >= text_start && offset < text_end)
                .collect()
        };
        if starts.first().copied() != Some(text_start) {
            starts.insert(0, text_start);
        }

        let mut chars = Vec::with_capacity(text_len);
        for (index, &segment_start) in starts.iter().enumerate() {
            if chars.len() >= text_len || segment_start >= text_end {
                break;
            }
            let segment_end = starts
                .get(index + 1)
                .copied()
                .unwrap_or(text_end)
                .min(text_end);
            let Some((&grbit, encoded)) = data
                .get(segment_start..segment_end)
                .and_then(|segment| segment.split_first())
            else {
                continue;
            };
            let remaining = text_len - chars.len();
            if grbit & 0x01 != 0 {
                for pair in encoded.chunks_exact(2).take(remaining) {
                    chars.push(u16::from_le_bytes([pair[0], pair[1]]));
                }
            } else {
                chars.extend(encoded.iter().take(remaining).map(|&byte| byte as u16));
            }
        }
        if chars.is_empty() {
            return None;
        }

        let run_bytes = data.get(text_end..text_end.saturating_add(cb_runs))?;
        let formatting = run_bytes
            .chunks_exact(8)
            .map(|run| {
                (
                    u16::from_le_bytes([run[0], run[1]]) as usize,
                    u16::from_le_bytes([run[2], run[3]]),
                )
            })
            .collect::<Vec<_>>();
        let plain_default =
            formatting.len() == 2 && formatting[0] == (0, 0) && formatting[1] == (text_len, 0);
        let runs = if formatting.is_empty() || plain_default {
            vec![RichTextRun::plain(String::from_utf16_lossy(&chars))]
        } else {
            let boundaries = formatting
                .iter()
                .copied()
                .filter(|(start, _)| *start < text_len)
                .collect::<Vec<_>>();
            let mut runs = Vec::new();
            if boundaries.first().is_some_and(|(start, _)| *start > 0) {
                runs.push(RichTextRun::plain(String::from_utf16_lossy(
                    &chars[..boundaries[0].0.min(chars.len())],
                )));
            }
            for (index, &(start, font_index)) in boundaries.iter().enumerate() {
                let end = boundaries
                    .get(index + 1)
                    .map(|(next, _)| *next)
                    .unwrap_or(text_len)
                    .min(chars.len());
                let start = start.min(end);
                if start == end {
                    continue;
                }
                runs.push(RichTextRun {
                    text: String::from_utf16_lossy(&chars[start..end]),
                    // ifnt 0 is the workbook default font (what the
                    // writer emits for font-less runs); resolving it
                    // would pin an explicit default onto plain runs.
                    font: (font_index != 0)
                        .then(|| style_ctx.resolve_run_font(font_index))
                        .flatten(),
                });
            }
            if runs.is_empty() {
                vec![RichTextRun::plain(String::from_utf16_lossy(&chars))]
            } else {
                runs
            }
        };
        Some(duke_sheets_core::ControlText {
            runs,
            horizontal_alignment,
            vertical_alignment,
        })
    }

    /// Parse a NOTE record's fields. The comment itself is created
    /// later, once the note can be matched with its comment shape.
    ///
    /// NOTE: row(2) + col(2) + flags(2) + objId(2) + author(XLUnicodeString).
    fn parse_note(data: &[u8]) -> XlsResult<Option<NoteData>> {
        if data.len() < 8 {
            return Ok(None);
        }
        let mut off = 0;
        let row = read_u16(data, &mut off)? as u32;
        let col = read_u16(data, &mut off)?;
        let flags = read_u16(data, &mut off)?;
        let obj_id = read_u16(data, &mut off)?;

        // Author follows as a XLUnicodeString (len(2) + flags(1) + chars)
        let author = if off < data.len() {
            read_unicode_string(data, &mut off).unwrap_or_default()
        } else {
            String::new()
        };

        Ok(Some(NoteData {
            row,
            col,
            visible: (flags & 0x0002) != 0,
            obj_id,
            author,
        }))
    }

    // ── Hyperlink record parsers ────────────────────────────────────────

    /// Parse an HLINK record.
    ///
    /// Format: Ref8U(8) + classId(16) + streamVersion(4) + flags(4) +
    /// [displayName] + [frameName] + [moniker] + [location].
    fn parse_hlink(data: &[u8], ws: &mut duke_sheets_core::Worksheet) -> XlsResult<()> {
        if data.len() < 32 {
            return Ok(());
        }
        let mut off = 0;
        let row_first = read_u16(data, &mut off)? as u32;
        let _row_last = read_u16(data, &mut off)?;
        let col_first = read_u16(data, &mut off)?;
        let _col_last = read_u16(data, &mut off)?;

        // Skip CLSID (16 bytes) and streamVersion (4 bytes)
        off += 20;

        let flags = read_u32(data, &mut off)?;
        let has_moniker = (flags & 0x01) != 0;
        let is_absolute = (flags & 0x02) != 0;
        let has_location = (flags & 0x08) != 0;
        let has_display = (flags & 0x10) != 0;
        let has_frame = (flags & 0x80) != 0;

        let display = if has_display {
            Some(Self::read_hlink_string(data, &mut off)?)
        } else {
            None
        };

        // Skip target frame name if present
        if has_frame {
            let _ = Self::read_hlink_string(data, &mut off)?;
        }

        let mut target = String::new();
        if has_moniker {
            target = Self::parse_hlink_moniker(data, &mut off, is_absolute)?;
        }

        let location = if has_location {
            Some(Self::read_hlink_string(data, &mut off)?)
        } else {
            None
        };

        // If no moniker but has location, it's an internal link
        if target.is_empty() {
            if let Some(loc) = &location {
                target = format!("#{}", loc);
            }
        }

        let hyperlink = Hyperlink {
            target,
            display,
            tooltip: None, // Set later from HLINKTOOLTIP if present
            location,
        };

        let cell_a1 = CellAddress::new(row_first, col_first).to_a1_string();
        let _ = ws.set_hyperlink(&cell_a1, hyperlink);
        Ok(())
    }

    /// Read a length-prefixed UTF-16LE string from HLINK data.
    /// Format: char_count(u32, includes null terminator) + UTF-16LE chars.
    fn read_hlink_string(data: &[u8], off: &mut usize) -> XlsResult<String> {
        if *off + 4 > data.len() {
            return Ok(String::new());
        }
        let char_count = read_u32(data, off)? as usize;
        if char_count == 0 {
            return Ok(String::new());
        }
        let byte_len = char_count * 2;
        if *off + byte_len > data.len() {
            return Ok(String::new());
        }
        let chars: Vec<u16> = data[*off..*off + byte_len]
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        *off += byte_len;
        // Strip null terminator
        let s = String::from_utf16_lossy(&chars);
        Ok(s.trim_end_matches('\0').to_string())
    }

    /// Parse the moniker portion of an HLINK record.
    fn parse_hlink_moniker(data: &[u8], off: &mut usize, _is_absolute: bool) -> XlsResult<String> {
        // URL moniker GUID 79EAC9E0-BAF9-11CE-8C82-00AA004BA90B in
        // on-disk LE-mixed format (Data1/2/3 little-endian, Data4
        // byte-ordered). Must match the writer's URL_MONIKER_CLSID
        // and the bytes Excel/LibreOffice produce.
        const URL_MONIKER: [u8; 16] = [
            0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B,
            0xA9, 0x0B,
        ];
        // File moniker GUID: 00000303-0000-0000-C000-000000000046
        const FILE_MONIKER: [u8; 16] = [
            0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x46,
        ];

        if *off + 16 > data.len() {
            return Ok(String::new());
        }
        let guid = &data[*off..*off + 16];
        *off += 16;

        if guid == URL_MONIKER {
            // URL moniker: length(4) + url(length bytes, UTF-16LE null-terminated)
            if *off + 4 > data.len() {
                return Ok(String::new());
            }
            let url_byte_len = read_u32(data, off)? as usize;
            if *off + url_byte_len > data.len() {
                return Ok(String::new());
            }
            let chars: Vec<u16> = data[*off..*off + url_byte_len]
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect();
            *off += url_byte_len;
            Ok(String::from_utf16_lossy(&chars)
                .trim_end_matches('\0')
                .to_string())
        } else if guid == FILE_MONIKER {
            // File moniker: dir_up(2) + path_len(4) + path(Latin-1) + ...
            if *off + 6 > data.len() {
                return Ok(String::new());
            }
            let dir_up = read_u16(data, off)? as usize;
            let path_len = read_u32(data, off)? as usize;
            if *off + path_len > data.len() {
                return Ok(String::new());
            }
            let short_path: String = data[*off..*off + path_len]
                .iter()
                .map(|&b| b as char)
                .collect();
            *off += path_len;

            // Skip unknown block (24 bytes)
            if *off + 24 <= data.len() {
                *off += 24;
            }

            // Try to read long filename if present
            if *off + 4 <= data.len() {
                let long_path_total = read_u32(data, off)? as usize;
                if long_path_total > 0 && *off + 6 <= data.len() {
                    *off += 2; // unknown
                    let long_path_len = read_u32(data, off)? as usize;
                    if long_path_len > 0 && *off + 2 <= data.len() {
                        *off += 2; // unknown
                        let byte_len = long_path_len;
                        if *off + byte_len <= data.len() {
                            let chars: Vec<u16> = data[*off..*off + byte_len]
                                .chunks_exact(2)
                                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                                .collect();
                            *off += byte_len;
                            let long_path = String::from_utf16_lossy(&chars)
                                .trim_end_matches('\0')
                                .to_string();
                            if !long_path.is_empty() {
                                let prefix = "../ ".repeat(dir_up);
                                return Ok(format!("{}{}", prefix.replace(" ", ""), long_path));
                            }
                        }
                    }
                }
            }

            // Fall back to short path
            let prefix = "../".repeat(dir_up);
            let trimmed = short_path.trim_end_matches('\0');
            Ok(format!("{}{}", prefix, trimmed))
        } else {
            // Unknown moniker type - skip
            Ok(String::new())
        }
    }

    /// Parse HLINKTOOLTIP (0x0800) - FRT record with tooltip text.
    fn parse_hlinktooltip(
        data: &[u8],
        tooltips: &mut std::collections::HashMap<(u32, u16), String>,
    ) {
        // FRT header: rt(2) + grbitFrt(2) + ref8(8) = 12 bytes, then UTF-16LE tooltip
        if data.len() < 14 {
            return;
        }
        let _rt = u16::from_le_bytes([data[0], data[1]]);
        let mut off = 4; // skip rt + grbitFrt
        let row_first = u16::from_le_bytes([data[off], data[off + 1]]) as u32;
        off += 2;
        let _row_last = u16::from_le_bytes([data[off], data[off + 1]]);
        off += 2;
        let col_first = u16::from_le_bytes([data[off], data[off + 1]]);
        off += 2;
        let _col_last = u16::from_le_bytes([data[off], data[off + 1]]);
        off += 2;
        // Remaining bytes are null-terminated UTF-16LE tooltip
        let remaining = &data[off..];
        let chars: Vec<u16> = remaining
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        let tooltip = String::from_utf16_lossy(&chars)
            .trim_end_matches('\0')
            .to_string();
        if !tooltip.is_empty() {
            tooltips.insert((row_first, col_first), tooltip);
        }
    }

    fn parse_feat_protection(data: &[u8], ws: &mut Worksheet) {
        if data.len() < 27 {
            return;
        }
        let rt = u16::from_le_bytes([data[0], data[1]]);
        if rt != records::FEAT {
            return;
        }
        let isf = u16::from_le_bytes([data[12], data[13]]);
        if isf != 2 {
            return;
        }
        let cref = u16::from_le_bytes([data[19], data[20]]) as usize;
        let mut off = 27usize;
        if off + cref * 8 > data.len() {
            return;
        }
        let mut ranges = Vec::with_capacity(cref);
        for _ in 0..cref {
            let row_first = u16::from_le_bytes([data[off], data[off + 1]]) as u32;
            let row_last = u16::from_le_bytes([data[off + 2], data[off + 3]]) as u32;
            let col_first = u16::from_le_bytes([data[off + 4], data[off + 5]]);
            let col_last = u16::from_le_bytes([data[off + 6], data[off + 7]]);
            off += 8;
            ranges.push(CellRange::from_indices(
                row_first, col_first, row_last, col_last,
            ));
        }
        if off + 8 > data.len() {
            return;
        }
        let flags = u32::from_le_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]);
        off += 4;
        let password = u32::from_le_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]);
        off += 4;
        let Ok(name) = read_unicode_string(data, &mut off) else {
            return;
        };
        if name.is_empty() || ranges.is_empty() {
            return;
        }

        let security_descriptor = if (flags & 0x0000_0001) != 0 && off < data.len() {
            Some(format!("hex:{}", Self::hex_encode(&data[off..])))
        } else {
            None
        };

        ws.add_protected_range(ProtectedRange {
            name,
            ranges,
            password_hash: if password != 0 {
                Some(password as u16)
            } else {
                None
            },
            security_descriptor,
        });
    }

    fn hex_encode(bytes: &[u8]) -> String {
        const HEX: &[u8; 16] = b"0123456789ABCDEF";
        let mut out = String::with_capacity(bytes.len() * 2);
        for &byte in bytes {
            out.push(HEX[(byte >> 4) as usize] as char);
            out.push(HEX[(byte & 0x0F) as usize] as char);
        }
        out
    }

    // ── Conditional formatting record parsers ────────────────────────────

    /// Parse CONDFMT (0x01B0) - conditional formatting range header.
    ///
    /// Format: cCF(2) + flags(2) + enclosing_range(8) + range_count(2) + ranges.
    fn parse_condfmt(data: &[u8]) -> Vec<CellRange> {
        if data.len() < 14 {
            return Vec::new();
        }
        let mut off = 4; // skip cCF(2) + flags(2)
                         // Skip enclosing range (8 bytes) - individual ranges are more precise
        off += 8;
        let range_count = u16::from_le_bytes([data[off], data[off + 1]]) as usize;
        off += 2;

        let mut ranges = Vec::with_capacity(range_count);
        for _ in 0..range_count {
            if off + 8 > data.len() {
                break;
            }
            let r1 = u16::from_le_bytes([data[off], data[off + 1]]) as u32;
            off += 2;
            let r2 = u16::from_le_bytes([data[off], data[off + 1]]) as u32;
            off += 2;
            let c1 = u16::from_le_bytes([data[off], data[off + 1]]);
            off += 2;
            let c2 = u16::from_le_bytes([data[off], data[off + 1]]);
            off += 2;
            ranges.push(CellRange::from_indices(r1, c1, r2, c2));
        }
        ranges
    }

    /// Parse CF (0x01B1) - conditional formatting rule.
    ///
    /// Format: ct(1) + cp(1) + cce1(2) + cce2(2) + [dxf_data] + formula1 + formula2.
    fn parse_cf(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        cf_ranges: &[CellRange],
        formula_ctx: &FormulaContext,
    ) {
        if data.len() < 6 {
            return;
        }
        let ct = data[0]; // 1=CellIs, 2=Expression
        let cp = data[1]; // comparison operator (0 for expression)
        let cce1 = u16::from_le_bytes([data[2], data[3]]) as usize;
        let cce2 = u16::from_le_bytes([data[4], data[5]]) as usize;

        // Formulas are at the END of the record
        let total = data.len();
        if total < 6 + cce1 + cce2 {
            return;
        }
        let f1_start = total - cce1 - cce2;
        let f2_start = total - cce2;

        let formula1 = if cce1 > 0 {
            let tokens = &data[f1_start..f1_start + cce1];
            let text = crate::biff::formula::decompile(tokens, formula_ctx);
            if text.is_empty() {
                None
            } else {
                Some(text)
            }
        } else {
            None
        };

        let formula2 = if cce2 > 0 {
            let tokens = &data[f2_start..f2_start + cce2];
            let text = crate::biff::formula::decompile(tokens, formula_ctx);
            if text.is_empty() {
                None
            } else {
                Some(text)
            }
        } else {
            None
        };

        let rule_type = match ct {
            0 | 1 => {
                // CellIs - map CP to CfOperator
                let operator = match cp {
                    1 => CfOperator::Between,
                    2 => CfOperator::NotBetween,
                    3 => CfOperator::Equal,
                    4 => CfOperator::NotEqual,
                    5 => CfOperator::GreaterThan,
                    6 => CfOperator::LessThan,
                    7 => CfOperator::GreaterThanOrEqual,
                    8 => CfOperator::LessThanOrEqual,
                    _ => CfOperator::Equal,
                };
                CfRuleType::CellIs {
                    operator,
                    formula1: formula1.unwrap_or_default(),
                    formula2,
                }
            }
            2 => {
                // Expression
                CfRuleType::Expression {
                    formula: formula1.unwrap_or_default(),
                }
            }
            _ => return,
        };

        let mut rule = ConditionalFormatRule::new(rule_type);
        rule.ranges = cf_ranges.to_vec();
        ws.add_conditional_format(rule);
    }

    // ── Data validation record parsers ───────────────────────────────────

    /// Parse a DV record (0x01BE) - data validation criteria.
    ///
    /// Format: flags(4) + input_title + error_title + input_msg + error_msg +
    /// cce1(2) + unused(2) + formula1 + cce2(2) + unused(2) + formula2 +
    /// range_count(2) + ranges.
    fn parse_dv(
        data: &[u8],
        ws: &mut duke_sheets_core::Worksheet,
        formula_ctx: &FormulaContext,
    ) -> XlsResult<()> {
        if data.len() < 4 {
            return Ok(());
        }
        let mut off = 0;
        let flags = read_u32(data, &mut off)?;

        let val_type = (flags & 0x0F) as u8;
        let err_style = ((flags >> 4) & 0x07) as u8;
        let is_explicit_list = (flags & 0x80) != 0;
        let allow_blank = (flags & 0x100) != 0;
        let suppress_dropdown = (flags & 0x200) != 0;
        let show_input = (flags & 0x40000) != 0;
        let show_error = (flags & 0x80000) != 0;
        let operator = ((flags >> 20) & 0x0F) as u8;

        // Read four Unicode strings: input_title, error_title, input_msg, error_msg
        let input_title = read_unicode_string(data, &mut off).unwrap_or_default();
        let error_title = read_unicode_string(data, &mut off).unwrap_or_default();
        let input_msg = read_unicode_string(data, &mut off).unwrap_or_default();
        let error_msg = read_unicode_string(data, &mut off).unwrap_or_default();

        // Formula 1: cce(2) + unused(2) + token_data(cce)
        let formula1 = if off + 4 <= data.len() {
            let cce1 = read_u16(data, &mut off)? as usize;
            let _unused = read_u16(data, &mut off)?;
            if cce1 > 0 && off + cce1 <= data.len() {
                let tokens = &data[off..off + cce1];
                off += cce1;
                let text = crate::biff::formula::decompile(tokens, formula_ctx);
                if text.is_empty() {
                    None
                } else {
                    Some(text)
                }
            } else {
                off += cce1;
                None
            }
        } else {
            None
        };

        // Formula 2: cce(2) + unused(2) + token_data(cce)
        let formula2 = if off + 4 <= data.len() {
            let cce2 = read_u16(data, &mut off)? as usize;
            let _unused = read_u16(data, &mut off)?;
            if cce2 > 0 && off + cce2 <= data.len() {
                let tokens = &data[off..off + cce2];
                off += cce2;
                let text = crate::biff::formula::decompile(tokens, formula_ctx);
                if text.is_empty() {
                    None
                } else {
                    Some(text)
                }
            } else {
                off += cce2;
                None
            }
        } else {
            None
        };

        // Ranges: range_count(2) + ranges(8 bytes each)
        let mut ranges = Vec::new();
        if off + 2 <= data.len() {
            let range_count = read_u16(data, &mut off)? as usize;
            for _ in 0..range_count {
                if off + 8 > data.len() {
                    break;
                }
                let r1 = read_u16(data, &mut off)? as u32;
                let r2 = read_u16(data, &mut off)? as u32;
                let c1 = read_u16(data, &mut off)?;
                let c2 = read_u16(data, &mut off)?;
                ranges.push(CellRange::from_indices(r1, c1, r2, c2));
            }
        }

        // Map operator byte to ValidationOperator
        let val_operator = match operator {
            0 => ValidationOperator::Between,
            1 => ValidationOperator::NotBetween,
            2 => ValidationOperator::Equal,
            3 => ValidationOperator::NotEqual,
            4 => ValidationOperator::GreaterThan,
            5 => ValidationOperator::LessThan,
            6 => ValidationOperator::GreaterThanOrEqual,
            7 => ValidationOperator::LessThanOrEqual,
            _ => ValidationOperator::Between,
        };

        // Build the validation type
        let f1 = formula1.unwrap_or_default();
        let f2 = formula2;

        let validation_type = match val_type {
            0 => ValidationType::None,
            1 => ValidationType::Whole {
                operator: val_operator,
                value1: f1,
                value2: f2,
            },
            2 => ValidationType::Decimal {
                operator: val_operator,
                value1: f1,
                value2: f2,
            },
            3 => {
                // List - formula1 may be a comma-separated string or formula
                let source = if is_explicit_list {
                    // Inline list: strip surrounding quotes if present
                    f1.trim_matches('"').to_string()
                } else {
                    f1
                };
                ValidationType::List { source }
            }
            4 => ValidationType::Date {
                operator: val_operator,
                value1: f1,
                value2: f2,
            },
            5 => ValidationType::Time {
                operator: val_operator,
                value1: f1,
                value2: f2,
            },
            6 => ValidationType::TextLength {
                operator: val_operator,
                value1: f1,
                value2: f2,
            },
            7 => ValidationType::Custom { formula: f1 },
            _ => ValidationType::None,
        };

        let error_style_val = match err_style {
            0 => ValidationErrorStyle::Stop,
            1 => ValidationErrorStyle::Warning,
            2 => ValidationErrorStyle::Information,
            _ => ValidationErrorStyle::Stop,
        };

        let validation = DataValidation {
            validation_type,
            ranges,
            allow_blank,
            show_dropdown: !suppress_dropdown,
            show_input_message: show_input,
            input_title: if input_title.is_empty() {
                None
            } else {
                Some(input_title)
            },
            input_message: if input_msg.is_empty() {
                None
            } else {
                Some(input_msg)
            },
            show_error_alert: show_error,
            error_style: error_style_val,
            error_title: if error_title.is_empty() {
                None
            } else {
                Some(error_title)
            },
            error_message: if error_msg.is_empty() {
                None
            } else {
                Some(error_msg)
            },
        };

        ws.add_data_validation(validation);
        Ok(())
    }
}

/// Synthesise the 14-byte `BITMAPFILEHEADER` for a DIB body so the
/// resulting bytes form a complete BMP file.
///
/// Layout:
///   bfType     u16 = 'BM' (0x4D42)
///   bfSize     u32 = 14 + dib.len()
///   bfReserved u32 = 0
///   bfOffBits  u32 = 14 + biSize + palette_size
///
/// `palette_size` is computed from the DIB's `biClrUsed` field
/// (offset 32 of BITMAPINFOHEADER) — or, when zero, from the BPP
/// (1/4/8-bit DIBs have an implicit 2/16/256-entry palette).
/// 16/24/32-bit DIBs have no palette.
fn bmp_file_header_for_dib(dib: &[u8]) -> [u8; 14] {
    let mut header = [0u8; 14];
    header[0] = b'B';
    header[1] = b'M';
    let total_size = 14u32 + dib.len() as u32;
    header[2..6].copy_from_slice(&total_size.to_le_bytes());
    // 6-9 reserved (zero)
    let bi_size = if dib.len() >= 4 {
        u32::from_le_bytes([dib[0], dib[1], dib[2], dib[3]])
    } else {
        40
    };
    let palette_bytes = if dib.len() >= 36 {
        let bpp = u16::from_le_bytes([dib[14], dib[15]]);
        let clr_used = u32::from_le_bytes([dib[32], dib[33], dib[34], dib[35]]);
        let entries = if clr_used != 0 {
            clr_used
        } else if bpp <= 8 {
            1u32 << bpp
        } else {
            0
        };
        entries * 4
    } else {
        0
    };
    let off_bits = 14u32 + bi_size + palette_bytes;
    header[10..14].copy_from_slice(&off_bits.to_le_bytes());
    header
}

/// Decode a UTF-16LE null-terminated byte buffer into a `String`.
/// Strips the trailing `\0` if present; tolerates odd-length input
/// by ignoring the dangling byte.
fn decode_utf16le_null_terminated(bytes: &[u8]) -> String {
    let mut units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    if units.last() == Some(&0) {
        units.pop();
    }
    String::from_utf16_lossy(&units)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Excel emits header-only FBSE placeholders (e.g. when it
    /// rasterizes picture transforms on save); pib blip ids are
    /// 1-based positions over every FBSE, so a skipped placeholder
    /// must still occupy its slot or every later picture resolves to
    /// the wrong blip (or none).
    #[test]
    fn bstore_placeholder_entries_keep_their_blip_slot() {
        // Placeholder FBSE: 36-byte body, no embedded blip record.
        let mut bstore = Vec::new();
        let fbse_placeholder_body = [0u8; 36];
        bstore.extend_from_slice(&[0x02, 0x00, 0x07, 0xF0]); // ver=2, FBSE
        bstore.extend_from_slice(&(fbse_placeholder_body.len() as u32).to_le_bytes());
        bstore.extend_from_slice(&fbse_placeholder_body);

        // Real FBSE: 36-byte header followed by an embedded PNG blip
        // (rh: ver=0, inst=0x6E0, type=0xF01E; payload = 16-byte UID +
        // tag byte + png bytes).
        let png = [0x89u8, b'P', b'N', b'G'];
        let mut blip = Vec::new();
        blip.extend_from_slice(&[0x00, 0x6E, 0x1E, 0xF0]);
        blip.extend_from_slice(&((16 + 1 + png.len()) as u32).to_le_bytes());
        blip.extend_from_slice(&[0u8; 16]);
        blip.push(0xFF);
        blip.extend_from_slice(&png);
        let mut fbse_real_body = vec![0u8; 36];
        fbse_real_body[0] = 6; // btWin32 = PNG
        fbse_real_body.extend_from_slice(&blip);
        bstore.extend_from_slice(&[0x02, 0x00, 0x07, 0xF0]);
        bstore.extend_from_slice(&(fbse_real_body.len() as u32).to_le_bytes());
        bstore.extend_from_slice(&fbse_real_body);

        let mut store: Vec<Option<BlipData>> = Vec::new();
        XlsReader::parse_bstore_container(&bstore, &mut store);
        assert_eq!(store.len(), 2, "placeholder keeps its slot");
        assert!(store[0].is_none());
        let real = store[1].as_ref().expect("second slot holds the PNG");
        assert_eq!(real.data, png);
    }

    #[test]
    fn resolve_workbook_stream_prefers_workbook() {
        let s = resolve_workbook_stream(|p| p == "/Workbook").unwrap();
        assert_eq!(s, "/Workbook");
    }

    #[test]
    fn resolve_workbook_stream_falls_back_to_book() {
        let s = resolve_workbook_stream(|p| p == "/Book").unwrap();
        assert_eq!(s, "/Book");
    }

    #[test]
    fn resolve_workbook_stream_rejects_encrypted_ooxml_package() {
        // Password-protected .xlsx files are CFB containers that hold the
        // encrypted ZIP in an EncryptedPackage stream alongside an
        // EncryptionInfo stream. They masquerade as XLS to the byte sniffer.
        let err = resolve_workbook_stream(|p| p == "/EncryptedPackage").unwrap_err();
        match err {
            XlsError::Encrypted(msg) => assert!(msg.contains("encrypted OOXML"), "msg={msg}"),
            other => panic!("expected Encrypted, got {other:?}"),
        }
    }

    #[test]
    fn resolve_workbook_stream_rejects_encrypted_ooxml_by_info_alone() {
        let err = resolve_workbook_stream(|p| p == "/EncryptionInfo").unwrap_err();
        assert!(matches!(err, XlsError::Encrypted(_)));
    }

    #[test]
    fn resolve_workbook_stream_invalid_format_when_nothing_matches() {
        let err = resolve_workbook_stream(|_| false).unwrap_err();
        match err {
            XlsError::InvalidFormat(msg) => {
                assert!(msg.contains("no Workbook or Book"), "msg={msg}")
            }
            other => panic!("expected InvalidFormat, got {other:?}"),
        }
    }

    fn rec(record_type: u16, data: Vec<u8>) -> BiffRecord {
        BiffRecord {
            record_type,
            data,
            stream_offset: 0,
            continue_offsets: Vec::new(),
        }
    }

    fn parse(records: Vec<BiffRecord>) -> duke_sheets_core::Worksheet {
        let mut ws = duke_sheets_core::Worksheet::new("Sheet1");
        let refs: Vec<&BiffRecord> = records.iter().collect();
        let formula_ctx = FormulaContext::new(vec!["Sheet1".to_string()]);
        let style_ctx = crate::styles::StyleContext::new();
        XlsReader::parse_sheet_records(
            &refs,
            &mut ws,
            &[],
            &[],
            &style_ctx,
            &formula_ctx,
            None,
            &[],
        )
        .unwrap();
        ws
    }

    #[test]
    fn test_parse_default_dimensions() {
        let mut def_col = Vec::new();
        def_col.extend_from_slice(&20u16.to_le_bytes());

        let mut def_row = Vec::new();
        def_row.extend_from_slice(&0u16.to_le_bytes());
        def_row.extend_from_slice(&400u16.to_le_bytes());

        let ws = parse(vec![
            rec(records::DEFCOLWIDTH, def_col),
            rec(records::DEFAULTROWHEIGHT, def_row),
        ]);

        assert_eq!(ws.default_column_width(), 20.0);
        assert_eq!(ws.default_row_height(), 20.0);
        assert_eq!(ws.column_width(42), 20.0);
        assert_eq!(ws.row_height(99), 20.0);
    }

    #[test]
    fn test_parse_freeze_panes_and_selection() {
        let mut window2 = Vec::new();
        window2.extend_from_slice(&0x0008u16.to_le_bytes());

        let mut pane = Vec::new();
        pane.extend_from_slice(&2u16.to_le_bytes());
        pane.extend_from_slice(&3u16.to_le_bytes());
        pane.extend_from_slice(&3u16.to_le_bytes());
        pane.extend_from_slice(&2u16.to_le_bytes());
        pane.extend_from_slice(&0u16.to_le_bytes());

        let mut selection = Vec::new();
        selection.push(0u8);
        selection.extend_from_slice(&5u16.to_le_bytes());
        selection.extend_from_slice(&4u16.to_le_bytes());
        selection.extend_from_slice(&0u16.to_le_bytes());
        selection.extend_from_slice(&1u16.to_le_bytes());
        selection.extend_from_slice(&5u16.to_le_bytes());
        selection.extend_from_slice(&6u16.to_le_bytes());
        selection.push(4u8);
        selection.push(5u8);

        let ws = parse(vec![
            rec(records::WINDOW2, window2),
            rec(records::PANE, pane),
            rec(records::SELECTION, selection),
        ]);

        assert_eq!(ws.freeze_panes().map(|p| (p.row, p.col)), Some((3, 2)));
        assert_eq!(ws.selections().len(), 1);
        let sel = &ws.selections()[0];
        assert_eq!(sel.active_cell.as_deref(), Some("E6"));
        assert_eq!(sel.sqref.as_deref(), Some("E6:F7"));
    }

    #[test]
    fn test_parse_outline_and_collapsed_from_row_colinfo() {
        let mut row = vec![0u8; 16];
        row[0..2].copy_from_slice(&4u16.to_le_bytes());
        row[6..8].copy_from_slice(&400u16.to_le_bytes());
        let row_opts: u32 = (3u32 << 8) | 0x10 | 0x20 | 0x40;
        row[12..16].copy_from_slice(&row_opts.to_le_bytes());

        let mut col = Vec::new();
        col.extend_from_slice(&2u16.to_le_bytes());
        col.extend_from_slice(&3u16.to_le_bytes());
        col.extend_from_slice(&2048u16.to_le_bytes());
        col.extend_from_slice(&0u16.to_le_bytes());
        let col_opts: u16 = 0x0001 | (2u16 << 8) | 0x1000;
        col.extend_from_slice(&col_opts.to_le_bytes());
        col.extend_from_slice(&0u16.to_le_bytes());

        let ws = parse(vec![rec(records::ROW, row), rec(records::COLINFO, col)]);

        assert!(ws.is_row_hidden(4));
        assert_eq!(ws.row_outline_level(4), 3);
        assert!(ws.is_row_collapsed(4));
        assert!((ws.row_height(4) - 20.0).abs() < 0.001);

        assert!(ws.is_column_hidden(2));
        assert!(ws.is_column_hidden(3));
        assert_eq!(ws.column_outline_level(2), 2);
        assert_eq!(ws.column_outline_level(3), 2);
        assert!(ws.is_column_collapsed(2));
        assert!(ws.is_column_collapsed(3));
    }

    #[test]
    fn test_parse_sheetlayout_tab_color() {
        let mut sheetlayout = vec![0u8; 11];
        sheetlayout[8] = 0x11;
        sheetlayout[9] = 0x22;
        sheetlayout[10] = 0x33;

        let ws = parse(vec![rec(records::SHEETLAYOUT, sheetlayout)]);

        assert_eq!(
            ws.tab_color(),
            Some(duke_sheets_core::Color::rgb(0x11, 0x22, 0x33))
        );
    }

    #[test]
    fn test_parse_setup_record() {
        let mut data = vec![0u8; 34];
        data[0..2].copy_from_slice(&9u16.to_le_bytes());
        data[2..4].copy_from_slice(&85u16.to_le_bytes());
        data[4..6].copy_from_slice(&1u16.to_le_bytes());
        data[6..8].copy_from_slice(&1u16.to_le_bytes());
        data[8..10].copy_from_slice(&2u16.to_le_bytes());
        data[10..12].copy_from_slice(&0x0002u16.to_le_bytes());
        data[16..24].copy_from_slice(&0.5f64.to_le_bytes());
        data[24..32].copy_from_slice(&0.4f64.to_le_bytes());

        let ws = parse(vec![rec(records::SETUP, data)]);
        let ps = ws.page_setup();
        assert_eq!(ps.paper_size, 9);
        assert_eq!(ps.scale, 85);
        assert_eq!(ps.fit_to_width, Some(1));
        assert_eq!(ps.fit_to_height, Some(2));
        assert_eq!(ps.orientation, duke_sheets_core::PageOrientation::Landscape);
        assert!((ps.header_margin - 0.5).abs() < 0.001);
        assert!((ps.footer_margin - 0.4).abs() < 0.001);
    }

    #[test]
    fn test_parse_margins() {
        let ws = parse(vec![
            rec(records::LEFT_MARGIN, 0.5f64.to_le_bytes().to_vec()),
            rec(records::RIGHT_MARGIN, 0.6f64.to_le_bytes().to_vec()),
            rec(records::TOP_MARGIN, 0.8f64.to_le_bytes().to_vec()),
            rec(records::BOTTOM_MARGIN, 0.9f64.to_le_bytes().to_vec()),
        ]);

        let ps = ws.page_setup();
        assert!((ps.left_margin - 0.5).abs() < 0.001);
        assert!((ps.right_margin - 0.6).abs() < 0.001);
        assert!((ps.top_margin - 0.8).abs() < 0.001);
        assert!((ps.bottom_margin - 0.9).abs() < 0.001);
    }

    #[test]
    fn test_parse_page_breaks() {
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&10u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&255u16.to_le_bytes());
        data.extend_from_slice(&25u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&255u16.to_le_bytes());

        let ws = parse(vec![rec(records::HPAGEBREAKS, data)]);
        let breaks = ws.row_breaks();
        assert_eq!(breaks.len(), 2);
        assert_eq!(breaks[0].id, 10);
        assert_eq!(breaks[1].id, 25);
    }

    #[test]
    fn test_parse_header_footer() {
        let header_text = "&LLeft&CCenter&RRight";
        let mut header_data = Vec::new();
        header_data.extend_from_slice(&(header_text.len() as u16).to_le_bytes());
        header_data.push(0x00);
        header_data.extend_from_slice(header_text.as_bytes());

        let footer_text = "&CPage &P";
        let mut footer_data = Vec::new();
        footer_data.extend_from_slice(&(footer_text.len() as u16).to_le_bytes());
        footer_data.push(0x00);
        footer_data.extend_from_slice(footer_text.as_bytes());

        let ws = parse(vec![
            rec(records::HEADER, header_data),
            rec(records::FOOTER, footer_data),
        ]);
        let ps = ws.page_setup();
        assert_eq!(ps.odd_header.as_deref(), Some(header_text));
        assert_eq!(ps.odd_footer.as_deref(), Some(footer_text));
    }

    fn rec_with_continue(
        record_type: u16,
        data: Vec<u8>,
        continue_offsets: Vec<usize>,
    ) -> BiffRecord {
        BiffRecord {
            record_type,
            data,
            stream_offset: 0,
            continue_offsets,
        }
    }

    // ── Comment tests ─────────────────────────────────────────────────

    #[test]
    fn test_parse_obj_id() {
        // ftCmo sub-record: rt=0x0015, cb=0x0012, ot=0x19(note), id=42
        let mut data = vec![0u8; 22];
        data[0..2].copy_from_slice(&0x0015u16.to_le_bytes()); // rt = ftCmo
        data[2..4].copy_from_slice(&0x0012u16.to_le_bytes()); // cb
        data[4..6].copy_from_slice(&0x0019u16.to_le_bytes()); // ot = note
        data[6..8].copy_from_slice(&42u16.to_le_bytes()); // id = 42
        assert_eq!(XlsReader::parse_obj_id(&data), Some(42));
    }

    #[test]
    fn test_parse_obj_id_wrong_subrecord() {
        let mut data = vec![0u8; 8];
        data[0..2].copy_from_slice(&0x0016u16.to_le_bytes()); // wrong rt
        assert_eq!(XlsReader::parse_obj_id(&data), None);
    }

    #[test]
    fn test_parse_txo_text_latin1() {
        // TXO header: 18 bytes, text_len at offset 10
        let mut header = vec![0u8; 18];
        header[10..12].copy_from_slice(&5u16.to_le_bytes()); // text_len = 5

        // CONTINUE data: grbit(1, compressed) + "Hello"
        let mut text_data = vec![0u8]; // grbit = 0 (Latin-1)
        text_data.extend_from_slice(b"Hello");

        let mut data = header;
        let continue_start = data.len();
        data.extend_from_slice(&text_data);

        let styles = StyleContext::new();
        let text = XlsReader::parse_txo_text(&data, &[continue_start], &styles);
        assert_eq!(
            text.map(|text| text.plain_text()),
            Some("Hello".to_string())
        );
    }

    #[test]
    fn test_parse_txo_text_utf16() {
        let mut header = vec![0u8; 18];
        header[10..12].copy_from_slice(&3u16.to_le_bytes()); // text_len = 3 chars

        // CONTINUE data: grbit(1, UTF-16) + 3 UTF-16LE chars "ABC"
        let mut text_data = vec![0x01u8]; // grbit = 1 (UTF-16LE)
        for &ch in &[b'A' as u16, b'B' as u16, b'C' as u16] {
            text_data.extend_from_slice(&ch.to_le_bytes());
        }

        let mut data = header;
        let continue_start = data.len();
        data.extend_from_slice(&text_data);

        let styles = StyleContext::new();
        let text = XlsReader::parse_txo_text(&data, &[continue_start], &styles);
        assert_eq!(text.map(|text| text.plain_text()), Some("ABC".to_string()));
    }

    #[test]
    fn test_parse_txo_text_across_encoding_switch_continues() {
        let mut data = vec![0u8; 18];
        data[10..12].copy_from_slice(&5u16.to_le_bytes()); // cchText
        data[12..14].copy_from_slice(&16u16.to_le_bytes()); // cbRuns

        let first = data.len();
        data.extend_from_slice(&[0x00, b'A', b'B', b'C']);
        let second = data.len();
        data.push(0x01);
        data.extend_from_slice(&(b'D' as u16).to_le_bytes());
        data.extend_from_slice(&(b'E' as u16).to_le_bytes());
        let runs = data.len();
        data.extend_from_slice(&[0u8; 16]);

        let styles = StyleContext::new();
        let text = XlsReader::parse_txo_text(&data, &[first, second, runs], &styles);
        assert_eq!(
            text.map(|text| text.plain_text()),
            Some("ABCDE".to_string())
        );
    }

    #[test]
    fn test_parse_note_with_comment() {
        // Build OBJ record with id=7
        let mut obj_data = vec![0u8; 22];
        obj_data[0..2].copy_from_slice(&0x0015u16.to_le_bytes());
        obj_data[2..4].copy_from_slice(&0x0012u16.to_le_bytes());
        obj_data[4..6].copy_from_slice(&0x0019u16.to_le_bytes());
        obj_data[6..8].copy_from_slice(&7u16.to_le_bytes()); // id=7

        // Build TXO with text "Review this"
        let text = b"Review this";
        let mut txo_header = vec![0u8; 18];
        txo_header[10..12].copy_from_slice(&(text.len() as u16).to_le_bytes());
        let mut txo_data = txo_header;
        let cont_off = txo_data.len();
        txo_data.push(0x00); // grbit = Latin-1
        txo_data.extend_from_slice(text);

        // Build NOTE: row=2, col=3, flags=0x0002(visible), objId=7, author="John"
        let mut note_data = Vec::new();
        note_data.extend_from_slice(&2u16.to_le_bytes()); // row
        note_data.extend_from_slice(&3u16.to_le_bytes()); // col
        note_data.extend_from_slice(&0x0002u16.to_le_bytes()); // flags (visible)
        note_data.extend_from_slice(&7u16.to_le_bytes()); // objId
                                                          // Author as XLUnicodeString: len(2) + flags(1) + chars
        let author = "John";
        note_data.extend_from_slice(&(author.len() as u16).to_le_bytes());
        note_data.push(0x00); // flags = compressed
        note_data.extend_from_slice(author.as_bytes());

        let ws = {
            let mut ws = duke_sheets_core::Worksheet::new("Sheet1");
            let recs = vec![
                rec(records::OBJ, obj_data),
                rec_with_continue(records::TXO, txo_data, vec![cont_off]),
                rec(records::NOTE, note_data),
            ];
            let refs: Vec<&BiffRecord> = recs.iter().collect();
            let formula_ctx = FormulaContext::new(vec!["Sheet1".to_string()]);
            let style_ctx = crate::styles::StyleContext::new();
            XlsReader::parse_sheet_records(
                &refs,
                &mut ws,
                &[],
                &[],
                &style_ctx,
                &formula_ctx,
                None,
                &[],
            )
            .unwrap();
            ws
        };

        let comment = ws.comment_at(2, 3).expect("comment should exist");
        assert_eq!(comment.author, "John");
        assert_eq!(comment.text, "Review this");
        assert_eq!(ws.comment_visible(2, 3), Some(true));
    }

    #[test]
    fn test_parse_note_no_text() {
        // NOTE without matching OBJ/TXO - should produce empty comment text
        let mut note_data = Vec::new();
        note_data.extend_from_slice(&0u16.to_le_bytes()); // row
        note_data.extend_from_slice(&0u16.to_le_bytes()); // col
        note_data.extend_from_slice(&0u16.to_le_bytes()); // flags
        note_data.extend_from_slice(&99u16.to_le_bytes()); // objId (no matching OBJ)
        note_data.extend_from_slice(&0u16.to_le_bytes()); // author len=0
        note_data.push(0x00); // flags

        let ws = parse(vec![rec(records::NOTE, note_data)]);
        let comment = ws.comment_at(0, 0).expect("comment should exist");
        assert_eq!(comment.text, "");
        assert_eq!(comment.author, "");
    }

    // ── Form control tests ────────────────────────────────────────────

    /// Excel-authored OBJ body of the hidden dropdown persisted for an
    /// autofilter column (ot=0x14, fUIObj set, ftLbsData lct=3).
    const AUTOFILTER_DROPDOWN_OBJ: &[u8] = &[
        0x15, 0x00, 0x12, 0x00, 0x14, 0x00, 0x01, 0x00, 0x01, 0x21, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0C, 0x00, 0x14, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x64, 0x00, 0x01, 0x00, 0x0A, 0x00, 0x00, 0x00, 0x10, 0x00, 0x01,
        0x00, 0x13, 0x00, 0xEE, 0x1F, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01, 0x03, 0x00, 0x00,
        0x02, 0x00, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    ];

    /// Minimal checkbox OBJ body: ftCmo(ot=0x0B, id=3) + ftCblsData
    /// (checked) + ftEnd.
    fn checkbox_obj_body(id: u16) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&0x0015u16.to_le_bytes()); // ftCmo
        body.extend_from_slice(&0x0012u16.to_le_bytes());
        body.extend_from_slice(&0x000Bu16.to_le_bytes()); // ot = checkbox
        body.extend_from_slice(&id.to_le_bytes());
        body.extend_from_slice(&0x0011u16.to_le_bytes()); // grbit
        body.extend_from_slice(&[0u8; 12]);
        body.extend_from_slice(&0x0012u16.to_le_bytes()); // ftCblsData
        body.extend_from_slice(&0x0008u16.to_le_bytes());
        body.extend_from_slice(&1u16.to_le_bytes()); // fChecked
        body.extend_from_slice(&[0u8; 4]); // accel + reserved
        body.extend_from_slice(&0x0002u16.to_le_bytes()); // flags (3D)
        body.extend_from_slice(&[0u8; 4]); // ftEnd
        body
    }

    #[test]
    fn autofilter_dropdown_obj_is_not_a_form_control() {
        // Excel persists auxiliary dropdown OBJs for autofilter
        // columns; they must not surface as user Dropdown controls.
        let ws = parse(vec![rec(records::OBJ, AUTOFILTER_DROPDOWN_OBJ.to_vec())]);
        assert_eq!(ws.form_control_count(), 0);
    }

    #[test]
    fn malformed_dropdown_obj_is_not_a_form_control() {
        let mut body = Vec::new();
        body.extend_from_slice(&0x0015u16.to_le_bytes()); // ftCmo
        body.extend_from_slice(&0x0012u16.to_le_bytes());
        body.extend_from_slice(&0x0014u16.to_le_bytes()); // ot = dropdown
        body.extend_from_slice(&1u16.to_le_bytes());
        body.extend_from_slice(&0x0011u16.to_le_bytes());
        body.extend_from_slice(&[0u8; 12]);
        body.extend_from_slice(&0x0013u16.to_le_bytes()); // ftLbsData
        body.extend_from_slice(&8u16.to_le_bytes());
        body.extend_from_slice(&[0x02, 0x00]); // truncated ObjFmla

        let ws = parse(vec![rec(records::OBJ, body)]);
        assert_eq!(ws.form_control_count(), 0);
    }

    #[test]
    fn control_without_escher_shape_gets_default_anchor() {
        // OBJ present but no MSODRAWING stream: the count mismatch
        // degrades to a default anchor, not a dropped control.
        let ws = parse(vec![rec(records::OBJ, checkbox_obj_body(3))]);
        assert_eq!(ws.form_control_count(), 1);
        let control = ws.form_controls().next().unwrap();
        assert_eq!(
            control.object.anchor,
            duke_sheets_chart::DrawingAnchor::default(),
            "mismatched pairing falls back to the default anchor"
        );
        match &control.payload.kind {
            duke_sheets_core::FormControlKind::Checkbox { state, .. } => {
                assert_eq!(*state, duke_sheets_core::CheckState::Checked);
            }
            other => panic!("expected Checkbox, got {other:?}"),
        }
    }

    #[test]
    fn out_of_spec_radio_mixed_state_clamps_to_checked() {
        // fChecked=2 is only legal for checkboxes; a radio carrying it
        // reads back as Checked.
        let mut body = Vec::new();
        body.extend_from_slice(&0x0015u16.to_le_bytes()); // ftCmo
        body.extend_from_slice(&0x0012u16.to_le_bytes());
        body.extend_from_slice(&0x000Cu16.to_le_bytes()); // ot = option button
        body.extend_from_slice(&4u16.to_le_bytes());
        body.extend_from_slice(&0x0011u16.to_le_bytes());
        body.extend_from_slice(&[0u8; 12]);
        body.extend_from_slice(&0x0012u16.to_le_bytes()); // ftCblsData
        body.extend_from_slice(&0x0008u16.to_le_bytes());
        body.extend_from_slice(&2u16.to_le_bytes()); // fChecked = mixed (invalid)
        body.extend_from_slice(&[0u8; 4]);
        body.extend_from_slice(&0x0002u16.to_le_bytes());
        body.extend_from_slice(&[0u8; 4]); // ftEnd

        let ws = parse(vec![rec(records::OBJ, body)]);
        assert_eq!(ws.form_control_count(), 1);
        let control = ws.form_controls().next().unwrap();
        match &control.payload.kind {
            duke_sheets_core::FormControlKind::OptionButton { state, .. } => {
                assert_eq!(*state, duke_sheets_core::CheckState::Checked);
            }
            other => panic!("expected OptionButton, got {other:?}"),
        }
    }

    #[test]
    fn escher_walk_handles_deep_nesting_and_oversized_lengths() {
        use crate::biff::escher::{rec_type, OfficeArtRecordHeader};

        let mut nested = Vec::new();
        OfficeArtRecordHeader::container(rec_type::DG_CONTAINER, 0, 0).write_to(&mut nested);
        for _ in 0..20_000 {
            let mut outer = Vec::with_capacity(nested.len() + 8);
            OfficeArtRecordHeader::container(rec_type::DG_CONTAINER, 0, nested.len() as u32)
            .write_to(&mut outer);
            outer.extend_from_slice(&nested);
            nested = outer;
        }

        let mut count = 0usize;
        XlsReader::walk_escher_records(&nested, |_, _| {
            count += 1;
            true
        });
        assert_eq!(count, 20_001);

        let mut oversized = Vec::new();
        OfficeArtRecordHeader::container(rec_type::DG_CONTAINER, 0, u32::MAX)
            .write_to(&mut oversized);
        let mut visited = false;
        XlsReader::walk_escher_records(&oversized, |_, _| {
            visited = true;
            true
        });
        assert!(!visited, "invalid record body is rejected before visiting");
    }

    #[test]
    fn escher_shape_pairing_skips_deleted_and_clientless_shapes() {
        use crate::biff::escher::{
            fsp_flags, rec_type, shape_type, write_client_data, OfficeArtFsp, OfficeArtRecordHeader,
        };

        let shape = |spid: u32, flags: u32, has_client_data: bool| {
            let mut body = Vec::new();
            OfficeArtFsp {
                spid,
                grf_persistence: flags,
            }
            .write_to(shape_type::HOST_CONTROL, &mut body);
            if has_client_data {
                write_client_data(&mut body);
            }
            let mut out = Vec::new();
            OfficeArtRecordHeader::container(rec_type::SP_CONTAINER, 0, body.len() as u32)
                .write_to(&mut out);
            out.extend_from_slice(&body);
            out
        };

        let mut bytes = shape(
            1025,
            fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT | fsp_flags::DELETED,
            true,
        );
        bytes.extend_from_slice(&shape(
            1026,
            fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT,
            false,
        ));
        bytes.extend_from_slice(&shape(
            1027,
            fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT,
            true,
        ));

        // The deleted shape is dropped from the tree entirely; the
        // clientless shape stays in the tree but does not consume an
        // OBJ pairing slot.
        let nodes = XlsReader::parse_shape_tree(&bytes);
        assert_eq!(nodes.len(), 2);
        assert_eq!(nodes[0].spid, 1026);
        assert_eq!(nodes[1].spid, 1027);
        assert_eq!(nodes[1].shape_type, shape_type::HOST_CONTROL);
        assert_eq!(XlsReader::client_data_count(&nodes), 1);
    }

    // ── Hyperlink tests ───────────────────────────────────────────────

    fn build_hlink_url(row: u16, col: u16, url: &str, display: Option<&str>) -> Vec<u8> {
        let mut data = Vec::new();
        // Ref8U: row_first, row_last, col_first, col_last
        data.extend_from_slice(&row.to_le_bytes());
        data.extend_from_slice(&row.to_le_bytes());
        data.extend_from_slice(&col.to_le_bytes());
        data.extend_from_slice(&col.to_le_bytes());

        // CLSID (16 bytes) + streamVersion (4 bytes)
        data.extend_from_slice(&[
            0x79, 0xEA, 0xC9, 0xD0, 0xBA, 0xF9, 0x11, 0xCE, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B,
            0xA9, 0x0B,
        ]);
        data.extend_from_slice(&2u32.to_le_bytes());

        // Flags
        let mut flags: u32 = 0x01 | 0x02; // hasMoniker | isAbsolute
        if display.is_some() {
            flags |= 0x10; // hasDisplayName
        }
        data.extend_from_slice(&flags.to_le_bytes());

        // Display name (if present)
        if let Some(d) = display {
            let chars: Vec<u16> = d.encode_utf16().chain(std::iter::once(0)).collect();
            data.extend_from_slice(&(chars.len() as u32).to_le_bytes());
            for ch in &chars {
                data.extend_from_slice(&ch.to_le_bytes());
            }
        }

        // URL moniker GUID 79EAC9E0-BAF9-11CE-8C82-00AA004BA90B in
        // on-disk LE-mixed format - matches the reader's URL_MONIKER
        // constant and the bytes Excel/LibreOffice emit.
        data.extend_from_slice(&[
            0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B,
            0xA9, 0x0B,
        ]);

        // URL: byte length + UTF-16LE null-terminated
        let url_chars: Vec<u16> = url.encode_utf16().chain(std::iter::once(0)).collect();
        let url_byte_len = url_chars.len() * 2;
        data.extend_from_slice(&(url_byte_len as u32).to_le_bytes());
        for ch in &url_chars {
            data.extend_from_slice(&ch.to_le_bytes());
        }

        data
    }

    #[test]
    fn test_parse_hlink_url() {
        let data = build_hlink_url(0, 0, "https://example.com", Some("Example"));
        let ws = parse(vec![rec(records::HLINK, data)]);

        let hl = ws.hyperlink("A1").expect("hyperlink should exist");
        assert_eq!(hl.target, "https://example.com");
        assert_eq!(hl.display.as_deref(), Some("Example"));
    }

    #[test]
    fn test_parse_hlink_url_no_display() {
        let data = build_hlink_url(1, 2, "https://rust-lang.org", None);
        let ws = parse(vec![rec(records::HLINK, data)]);

        let hl = ws.hyperlink("C2").expect("hyperlink should exist");
        assert_eq!(hl.target, "https://rust-lang.org");
        assert_eq!(hl.display, None);
    }

    #[test]
    fn test_parse_hlink_internal() {
        // Internal link (location only, no moniker)
        let mut data = Vec::new();
        // Ref8U
        data.extend_from_slice(&0u16.to_le_bytes()); // row
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes()); // col
        data.extend_from_slice(&0u16.to_le_bytes());
        // CLSID + streamVersion
        data.extend_from_slice(&[0u8; 16]);
        data.extend_from_slice(&2u32.to_le_bytes());
        // Flags: hasLocation only
        data.extend_from_slice(&0x08u32.to_le_bytes());
        // Location string: "Sheet2!A1"
        let loc = "Sheet2!A1";
        let chars: Vec<u16> = loc.encode_utf16().chain(std::iter::once(0)).collect();
        data.extend_from_slice(&(chars.len() as u32).to_le_bytes());
        for ch in &chars {
            data.extend_from_slice(&ch.to_le_bytes());
        }

        let ws = parse(vec![rec(records::HLINK, data)]);
        let hl = ws.hyperlink("A1").expect("hyperlink should exist");
        assert_eq!(hl.target, "#Sheet2!A1");
        assert_eq!(hl.location.as_deref(), Some("Sheet2!A1"));
    }

    // ── Conditional formatting tests ─────────────────────────────────

    fn build_condfmt(ranges: &[(u32, u32, u16, u16)]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // cCF = 1
        data.extend_from_slice(&0u16.to_le_bytes()); // flags
                                                     // Enclosing range (first range)
        let (r1, r2, c1, c2) = ranges[0];
        data.extend_from_slice(&(r1 as u16).to_le_bytes());
        data.extend_from_slice(&(r2 as u16).to_le_bytes());
        data.extend_from_slice(&c1.to_le_bytes());
        data.extend_from_slice(&c2.to_le_bytes());
        // Range count + ranges
        data.extend_from_slice(&(ranges.len() as u16).to_le_bytes());
        for &(r1, r2, c1, c2) in ranges {
            data.extend_from_slice(&(r1 as u16).to_le_bytes());
            data.extend_from_slice(&(r2 as u16).to_le_bytes());
            data.extend_from_slice(&c1.to_le_bytes());
            data.extend_from_slice(&c2.to_le_bytes());
        }
        data
    }

    fn build_cf_cellis(operator: u8, formula1: &[u8], formula2: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.push(1); // ct = CellIs
        data.push(operator); // cp
        data.extend_from_slice(&(formula1.len() as u16).to_le_bytes()); // cce1
        data.extend_from_slice(&(formula2.len() as u16).to_le_bytes()); // cce2
                                                                        // No DXF data for this test - formulas are at the end
        data.extend_from_slice(formula1);
        data.extend_from_slice(formula2);
        data
    }

    #[test]
    fn test_parse_condfmt_and_cf_cellis() {
        let condfmt_data = build_condfmt(&[(0, 9, 0, 2)]); // A1:C10
                                                           // CF: CellIs, GreaterThan (5), formula1 = integer 100
                                                           // PTG for integer: tInt (0x1E) + value(2) = [0x1E, 0x64, 0x00]
        let cf_data = build_cf_cellis(5, &[0x1E, 0x64, 0x00], &[]);

        let ws = parse(vec![
            rec(records::CONDFMT, condfmt_data),
            rec(records::CF, cf_data),
        ]);

        let cf_rules = ws.conditional_formats();
        assert_eq!(cf_rules.len(), 1);
        let rule = &cf_rules[0];
        assert_eq!(rule.ranges.len(), 1);
        match &rule.rule_type {
            CfRuleType::CellIs {
                operator, formula1, ..
            } => {
                assert_eq!(*operator, CfOperator::GreaterThan);
                assert_eq!(formula1, "100");
            }
            _ => panic!("Expected CellIs rule type"),
        }
    }

    #[test]
    fn test_parse_condfmt_expression() {
        let condfmt_data = build_condfmt(&[(0, 4, 0, 0)]); // A1:A5
                                                           // CF: Expression (ct=2), formula1 = tInt(0x1E) + 0
        let mut cf_data = Vec::new();
        cf_data.push(2); // ct = Expression
        cf_data.push(0); // cp (unused for expression)
        cf_data.extend_from_slice(&3u16.to_le_bytes()); // cce1 = 3
        cf_data.extend_from_slice(&0u16.to_le_bytes()); // cce2 = 0
        cf_data.extend_from_slice(&[0x1E, 0x01, 0x00]); // tInt(1)

        let ws = parse(vec![
            rec(records::CONDFMT, condfmt_data),
            rec(records::CF, cf_data),
        ]);

        let cf_rules = ws.conditional_formats();
        assert_eq!(cf_rules.len(), 1);
        match &cf_rules[0].rule_type {
            CfRuleType::Expression { formula } => assert_eq!(formula, "1"),
            _ => panic!("Expected Expression rule type"),
        }
    }

    // ── Data validation tests ─────────────────────────────────────────

    fn build_dv_record(
        val_type: u8,
        err_style: u8,
        operator: u8,
        allow_blank: bool,
        show_input: bool,
        show_error: bool,
        input_title: &str,
        error_title: &str,
        input_msg: &str,
        error_msg: &str,
        formula1_tokens: &[u8],
        formula2_tokens: &[u8],
        ranges: &[(u32, u32, u16, u16)],
    ) -> Vec<u8> {
        let mut data = Vec::new();

        // Build flags
        let mut flags: u32 = (val_type as u32) & 0x0F;
        flags |= ((err_style as u32) & 0x07) << 4;
        if allow_blank {
            flags |= 0x100;
        }
        if show_input {
            flags |= 0x40000;
        }
        if show_error {
            flags |= 0x80000;
        }
        flags |= ((operator as u32) & 0x0F) << 20;
        data.extend_from_slice(&flags.to_le_bytes());

        // Write unicode strings: len(u16) + flags(u8) + chars
        for s in &[input_title, error_title, input_msg, error_msg] {
            data.extend_from_slice(&(s.len() as u16).to_le_bytes());
            data.push(0x00); // compressed Latin-1
            data.extend_from_slice(s.as_bytes());
        }

        // Formula 1: cce(2) + unused(2) + tokens
        data.extend_from_slice(&(formula1_tokens.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(formula1_tokens);

        // Formula 2: cce(2) + unused(2) + tokens
        data.extend_from_slice(&(formula2_tokens.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(formula2_tokens);

        // Ranges
        data.extend_from_slice(&(ranges.len() as u16).to_le_bytes());
        for &(r1, r2, c1, c2) in ranges {
            data.extend_from_slice(&(r1 as u16).to_le_bytes());
            data.extend_from_slice(&(r2 as u16).to_le_bytes());
            data.extend_from_slice(&c1.to_le_bytes());
            data.extend_from_slice(&c2.to_le_bytes());
        }

        data
    }

    #[test]
    fn test_parse_dv_list_validation() {
        // List validation with explicit inline list
        let mut flags: u32 = 3; // type=list
        flags |= 0x80; // fStrLookup (explicit list)
        flags |= 0x100; // allow blank
        flags |= 0x40000; // show input
        flags |= 0x80000; // show error

        let mut data = Vec::new();
        data.extend_from_slice(&flags.to_le_bytes());

        // Strings: input_title, error_title, input_msg, error_msg
        for s in &["Choose", "Error", "Pick one", "Invalid"] {
            data.extend_from_slice(&(s.len() as u16).to_le_bytes());
            data.push(0x00);
            data.extend_from_slice(s.as_bytes());
        }

        // Formula 1: tStr for the list "Red,Green,Blue"
        // For simplicity, use tStr: 0x17 + len(1) + flags(1) + chars
        let list_str = "Red,Green,Blue";
        let mut f1 = vec![0x17]; // tStr
        f1.push(list_str.len() as u8);
        f1.push(0x00); // compressed
        f1.extend_from_slice(list_str.as_bytes());
        data.extend_from_slice(&(f1.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&f1);

        // Formula 2: empty
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());

        // Ranges: A1:A10
        data.extend_from_slice(&1u16.to_le_bytes()); // count
        data.extend_from_slice(&0u16.to_le_bytes()); // r1
        data.extend_from_slice(&9u16.to_le_bytes()); // r2
        data.extend_from_slice(&0u16.to_le_bytes()); // c1
        data.extend_from_slice(&0u16.to_le_bytes()); // c2

        let ws = parse(vec![rec(records::DV, data)]);
        let validations = ws.data_validations();
        assert_eq!(validations.len(), 1);
        let v = &validations[0];
        assert!(v.allow_blank);
        assert!(v.show_input_message);
        assert!(v.show_error_alert);
        assert_eq!(v.input_title.as_deref(), Some("Choose"));
        assert_eq!(v.error_title.as_deref(), Some("Error"));
        match &v.validation_type {
            ValidationType::List { source } => {
                // The decompiled formula might include quotes or not
                assert!(source.contains("Red"));
                assert!(source.contains("Green"));
                assert!(source.contains("Blue"));
            }
            _ => panic!("Expected List validation, got {:?}", v.validation_type),
        }
    }

    #[test]
    fn test_parse_dv_whole_number_between() {
        // tInt(0x1E) + value(u16): encodes an integer constant
        let f1 = vec![0x1E, 0x01, 0x00]; // tInt(1)
        let f2 = vec![0x1E, 0x64, 0x00]; // tInt(100)
        let dv_data = build_dv_record(
            1,     // whole number
            0,     // stop
            0,     // between
            true,  // allow blank
            false, // no input msg
            true,  // show error
            "",
            "Invalid",
            "",
            "Enter 1-100",
            &f1,
            &f2,
            &[(0, 9, 0, 0)], // A1:A10
        );

        let ws = parse(vec![rec(records::DV, dv_data)]);
        let validations = ws.data_validations();
        assert_eq!(validations.len(), 1);
        let v = &validations[0];
        match &v.validation_type {
            ValidationType::Whole {
                operator,
                value1,
                value2,
            } => {
                assert_eq!(*operator, ValidationOperator::Between);
                assert_eq!(value1, "1");
                assert_eq!(value2.as_deref(), Some("100"));
            }
            _ => panic!("Expected Whole validation, got {:?}", v.validation_type),
        }
        assert_eq!(v.error_title.as_deref(), Some("Invalid"));
        assert_eq!(v.error_message.as_deref(), Some("Enter 1-100"));
        assert_eq!(v.ranges.len(), 1);
    }

    #[test]
    fn test_parse_dv_custom_formula() {
        let f1 = vec![0x1E, 0x00, 0x00]; // tInt(0) as placeholder formula
        let dv_data = build_dv_record(
            7, // custom
            0, // stop
            0, // unused
            true,
            false,
            false,
            "",
            "",
            "",
            "",
            &f1,
            &[],
            &[(0, 0, 0, 0)], // A1
        );

        let ws = parse(vec![rec(records::DV, dv_data)]);
        let validations = ws.data_validations();
        assert_eq!(validations.len(), 1);
        match &validations[0].validation_type {
            ValidationType::Custom { formula } => assert_eq!(formula, "0"),
            _ => panic!("Expected Custom validation"),
        }
    }

    #[test]
    fn test_parse_autofilter_top_n() {
        let mut data = vec![0u8; 24];
        data[0] = 2;
        data[1] = 0;
        let flags: u16 = (1 << 4) | (1 << 5) | (5 << 7);
        data[2..4].copy_from_slice(&flags.to_le_bytes());

        let fc = XlsReader::parse_autofilter(&data).unwrap();
        assert_eq!(fc.col_id, 2);
        match &fc.filter {
            duke_sheets_core::auto_filter::ColumnFilter::Top10(t) => {
                assert!(t.top);
                assert!(!t.percent);
                assert_eq!(t.val as u16, 5);
            }
            other => panic!("Expected Top10, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_autofilter_custom_greater_than() {
        let mut data = vec![0u8; 24];
        data[0] = 0;
        data[1] = 0;
        data[2] = 0;
        data[3] = 0;
        data[4] = 0x04;
        data[5] = 0x04;
        data[6..14].copy_from_slice(&50.0f64.to_le_bytes());

        let fc = XlsReader::parse_autofilter(&data).unwrap();
        assert_eq!(fc.col_id, 0);
        match &fc.filter {
            duke_sheets_core::auto_filter::ColumnFilter::Custom(cf) => {
                assert_eq!(cf.conditions.len(), 1);
                assert_eq!(
                    cf.conditions[0].operator,
                    duke_sheets_core::auto_filter::FilterOperator::GreaterThan
                );
                assert!(cf.conditions[0].value.contains("50"));
            }
            other => panic!("Expected Custom, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_autofilter_string_equal() {
        let mut data = vec![0u8; 24];
        data[0] = 1;
        data[1] = 0;
        data[2] = 1;
        data[3] = 0;
        data[4] = 0x06;
        data[5] = 0x02;
        data[10] = 3;
        data[11] = 1;
        data.push(0x00);
        data.extend_from_slice(b"Red");

        let fc = XlsReader::parse_autofilter(&data).unwrap();
        assert_eq!(fc.col_id, 1);
        match &fc.filter {
            duke_sheets_core::auto_filter::ColumnFilter::Values(vf) => {
                assert_eq!(vf.values, vec!["Red"]);
            }
            other => panic!("Expected Values, got {:?}", other),
        }
    }

    #[test]
    fn test_extract_filter_db_range() {
        let mut body = vec![0x3B];
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&9u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&3u16.to_le_bytes());

        let ctx = FormulaContext {
            sheet_names: vec!["Sheet1".to_string()],
            extern_sheet: vec![],
            supbooks: vec![],
            names: vec![],
            extern_names: vec![],
            extern_name_index_base: 1,
            base_cell: None,
        };

        let range = XlsReader::extract_filter_db_range(&body, &ctx).unwrap();
        assert_eq!(range.start.row, 0);
        assert_eq!(range.end.row, 9);
        assert_eq!(range.start.col, 0);
        assert_eq!(range.end.col, 3);
    }

    #[test]
    fn test_parse_scl_zoom() {
        // SCL record: numerator(u16) + denominator(u16)
        // 75% zoom: 3/4
        let mut scl = Vec::new();
        scl.extend_from_slice(&3u16.to_le_bytes()); // numerator
        scl.extend_from_slice(&4u16.to_le_bytes()); // denominator

        let ws = parse(vec![rec(records::SCL, scl)]);
        assert_eq!(ws.zoom_scale(), Some(75));
    }

    #[test]
    fn test_parse_scl_zoom_150() {
        // 150% zoom: 3/2
        let mut scl = Vec::new();
        scl.extend_from_slice(&3u16.to_le_bytes());
        scl.extend_from_slice(&2u16.to_le_bytes());

        let ws = parse(vec![rec(records::SCL, scl)]);
        assert_eq!(ws.zoom_scale(), Some(150));
    }

    #[test]
    fn test_parse_scl_zoom_zero_denominator() {
        // Zero denominator should be ignored
        let mut scl = Vec::new();
        scl.extend_from_slice(&3u16.to_le_bytes());
        scl.extend_from_slice(&0u16.to_le_bytes());

        let ws = parse(vec![rec(records::SCL, scl)]);
        assert_eq!(ws.zoom_scale(), None);
    }

    #[test]
    fn test_parse_printheaders() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // print headings = true

        let ws = parse(vec![rec(records::PRINTHEADERS, data)]);
        assert!(ws.page_setup().print_headings);
    }

    #[test]
    fn test_parse_printheaders_off() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // print headings = false

        let ws = parse(vec![rec(records::PRINTHEADERS, data)]);
        assert!(!ws.page_setup().print_headings);
    }

    #[test]
    fn test_parse_printgridlines() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // print gridlines = true

        let ws = parse(vec![rec(records::PRINTGRIDLINES, data)]);
        assert!(ws.page_setup().print_gridlines);
    }

    #[test]
    fn test_parse_printgridlines_off() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // print gridlines = false

        let ws = parse(vec![rec(records::PRINTGRIDLINES, data)]);
        assert!(!ws.page_setup().print_gridlines);
    }

    #[test]
    fn test_extract_print_titles_rows_only() {
        // Print_Titles with repeat rows 1:3 (rows 0..2, all columns)
        // tArea3d: token(0x3B) + ixti(2) + first_row(2) + last_row(2) + first_col(2) + last_col(2)
        let mut body = vec![0x3B];
        body.extend_from_slice(&0u16.to_le_bytes()); // ixti
        body.extend_from_slice(&0u16.to_le_bytes()); // first_row = 0
        body.extend_from_slice(&2u16.to_le_bytes()); // last_row = 2
        body.extend_from_slice(&0u16.to_le_bytes()); // first_col = 0
        body.extend_from_slice(&0x00FFu16.to_le_bytes()); // last_col = 0xFF (all cols)

        let (rows, cols) = XlsReader::extract_print_titles(&body);
        assert_eq!(rows, Some((0, 2)));
        assert_eq!(cols, None);
    }

    #[test]
    fn test_extract_print_titles_cols_only() {
        // Print_Titles with repeat cols A:B (cols 0..1, all rows)
        let mut body = vec![0x3B];
        body.extend_from_slice(&0u16.to_le_bytes()); // ixti
        body.extend_from_slice(&0u16.to_le_bytes()); // first_row = 0
        body.extend_from_slice(&0xFFFFu16.to_le_bytes()); // last_row = 0xFFFF (all rows)
        body.extend_from_slice(&0u16.to_le_bytes()); // first_col = 0
        body.extend_from_slice(&1u16.to_le_bytes()); // last_col = 1

        let (rows, cols) = XlsReader::extract_print_titles(&body);
        assert_eq!(rows, None);
        assert_eq!(cols, Some((0, 1)));
    }

    #[test]
    fn test_extract_print_titles_both_rows_and_cols() {
        // Print_Titles with both rows 1:3 and cols A:B
        // tMemFunc(0x29) + cce(2) + tArea3d(rows) + tArea3d(cols) + tList(0x10)
        let mut body = Vec::new();
        body.push(0x29); // tMemFunc
        body.extend_from_slice(&22u16.to_le_bytes()); // cce = 2*11 bytes for two tArea3d

        // First tArea3d: rows 0..2, all columns
        body.push(0x3B);
        body.extend_from_slice(&0u16.to_le_bytes()); // ixti
        body.extend_from_slice(&0u16.to_le_bytes()); // first_row = 0
        body.extend_from_slice(&2u16.to_le_bytes()); // last_row = 2
        body.extend_from_slice(&0u16.to_le_bytes()); // first_col = 0
        body.extend_from_slice(&0x00FFu16.to_le_bytes()); // last_col = 0xFF

        // Second tArea3d: all rows, cols 0..1
        body.push(0x3B);
        body.extend_from_slice(&0u16.to_le_bytes()); // ixti
        body.extend_from_slice(&0u16.to_le_bytes()); // first_row = 0
        body.extend_from_slice(&0xFFFFu16.to_le_bytes()); // last_row = 0xFFFF
        body.extend_from_slice(&0u16.to_le_bytes()); // first_col = 0
        body.extend_from_slice(&1u16.to_le_bytes()); // last_col = 1

        // tList (union)
        body.push(0x10);

        let (rows, cols) = XlsReader::extract_print_titles(&body);
        assert_eq!(rows, Some((0, 2)));
        assert_eq!(cols, Some((0, 1)));
    }

    #[test]
    fn test_extract_print_titles_empty_body() {
        let body = vec![];
        let (rows, cols) = XlsReader::extract_print_titles(&body);
        assert_eq!(rows, None);
        assert_eq!(cols, None);
    }

    #[test]
    fn test_sst_runs_to_rich_text_no_runs() {
        let style_ctx = crate::styles::StyleContext::new();
        let runs = XlsReader::sst_runs_to_rich_text("Hello", &[], &style_ctx);
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Hello");
        assert!(runs[0].font.is_none());
    }

    #[test]
    fn test_sst_runs_to_rich_text_single_run_from_start() {
        use crate::biff::strings::FormattingRun;
        let mut style_ctx = crate::styles::StyleContext::new();
        // Add a bold font at index 0
        style_ctx.fonts.push(crate::styles::BiffFont {
            height_twips: 220, // 11pt
            bold: true,
            italic: false,
            underline: 0,
            strikethrough: false,
            color_index: 0x7FFF, // auto
            superscript: 0,
            name: "Calibri".to_string(),
        });

        let formatting_runs = vec![FormattingRun {
            char_pos: 0,
            font_index: 0,
        }];

        let runs = XlsReader::sst_runs_to_rich_text("Bold text", &formatting_runs, &style_ctx);
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Bold text");
        let font = runs[0].font.as_ref().unwrap();
        assert_eq!(font.bold, Some(true));
        assert_eq!(font.name, Some("Calibri".to_string()));
        assert_eq!(font.size, Some(11.0));
    }

    #[test]
    fn test_sst_runs_to_rich_text_two_runs() {
        use crate::biff::strings::FormattingRun;
        let mut style_ctx = crate::styles::StyleContext::new();
        // Font 0: normal
        style_ctx.fonts.push(crate::styles::BiffFont {
            height_twips: 220,
            bold: false,
            italic: false,
            underline: 0,
            strikethrough: false,
            color_index: 0x7FFF,
            superscript: 0,
            name: "Calibri".to_string(),
        });
        // Font 1: bold italic
        style_ctx.fonts.push(crate::styles::BiffFont {
            height_twips: 280, // 14pt
            bold: true,
            italic: true,
            underline: 0,
            strikethrough: false,
            color_index: 10, // red from palette
            superscript: 0,
            name: "Arial".to_string(),
        });

        let formatting_runs = vec![
            FormattingRun {
                char_pos: 0,
                font_index: 0,
            },
            FormattingRun {
                char_pos: 6,
                font_index: 1,
            },
        ];

        let runs = XlsReader::sst_runs_to_rich_text("Hello World", &formatting_runs, &style_ctx);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Hello ");
        // Font 0 is "normal" - no bold/italic, so RunFont will have
        // size + name set but no bold.
        let f0 = runs[0].font.as_ref().unwrap();
        assert_eq!(f0.name, Some("Calibri".to_string()));
        assert_eq!(f0.size, Some(11.0));

        assert_eq!(runs[1].text, "World");
        let f1 = runs[1].font.as_ref().unwrap();
        assert_eq!(f1.bold, Some(true));
        assert_eq!(f1.italic, Some(true));
        assert_eq!(f1.size, Some(14.0));
        assert_eq!(f1.name, Some("Arial".to_string()));
        // Red from palette index 10
        assert!(matches!(
            f1.color,
            Some(duke_sheets_core::Color::Rgb { r: 255, g: 0, b: 0 })
        ));
    }

    #[test]
    fn test_sst_runs_to_rich_text_leading_plain() {
        use crate::biff::strings::FormattingRun;
        let mut style_ctx = crate::styles::StyleContext::new();
        // Font 0: bold
        style_ctx.fonts.push(crate::styles::BiffFont {
            height_twips: 220,
            bold: true,
            italic: false,
            underline: 0,
            strikethrough: false,
            color_index: 0x7FFF,
            superscript: 0,
            name: "Calibri".to_string(),
        });

        // Run starts at char 6 ("World"), so "Hello " is plain (no run font)
        let formatting_runs = vec![FormattingRun {
            char_pos: 6,
            font_index: 0,
        }];

        let runs = XlsReader::sst_runs_to_rich_text("Hello World", &formatting_runs, &style_ctx);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Hello ");
        assert!(runs[0].font.is_none()); // leading text inherits cell style
        assert_eq!(runs[1].text, "World");
        assert_eq!(runs[1].font.as_ref().unwrap().bold, Some(true));
    }

    #[test]
    fn test_parse_table_record_two_variable() {
        // TABLE record: range A1:C3, flags=fTbl2, row input=D1, col input=E1
        let mut data = vec![0u8; 16];
        data[0..2].copy_from_slice(&0u16.to_le_bytes()); // first_row=0
        data[2..4].copy_from_slice(&2u16.to_le_bytes()); // last_row=2
        data[4] = 0; // first_col=0
        data[5] = 2; // last_col=2
        data[6..8].copy_from_slice(&0x0004u16.to_le_bytes()); // flags: fTbl2=1
        data[8..10].copy_from_slice(&0u16.to_le_bytes()); // rwInpRw=0
        data[10..12].copy_from_slice(&3u16.to_le_bytes()); // colInpRw=3 -> D1
        data[12..14].copy_from_slice(&0u16.to_le_bytes()); // rwInpCol=0
        data[14..16].copy_from_slice(&4u16.to_le_bytes()); // colInpCol=4 -> E1
        let result = XlsReader::parse_table_record(&data);
        assert!(result.is_some());
        let (mr, mc, input1, input2) = result.unwrap();
        assert_eq!(mr, 0);
        assert_eq!(mc, 0);
        assert_eq!(input1, "D1");
        assert_eq!(input2, "E1");
    }

    #[test]
    fn test_parse_table_record_row_input() {
        let mut data = vec![0u8; 16];
        data[0..2].copy_from_slice(&0u16.to_le_bytes());
        data[2..4].copy_from_slice(&4u16.to_le_bytes());
        data[4] = 0;
        data[5] = 0;
        data[6..8].copy_from_slice(&0x0002u16.to_le_bytes()); // flags: fRw=1
        data[8..10].copy_from_slice(&0u16.to_le_bytes()); // rwInpRw=0
        data[10..12].copy_from_slice(&2u16.to_le_bytes()); // colInpRw=2 -> C1
        let result = XlsReader::parse_table_record(&data);
        let (_, _, input1, input2) = result.unwrap();
        assert_eq!(input1, "C1");
        assert_eq!(input2, "");
    }

    #[test]
    fn test_parse_table_record_col_input() {
        let mut data = vec![0u8; 16];
        data[0..2].copy_from_slice(&0u16.to_le_bytes());
        data[2..4].copy_from_slice(&4u16.to_le_bytes());
        data[4] = 0;
        data[5] = 0;
        data[6..8].copy_from_slice(&0x0000u16.to_le_bytes()); // flags: fRw=0, fTbl2=0
        data[12..14].copy_from_slice(&1u16.to_le_bytes()); // rwInpCol=1
        data[14..16].copy_from_slice(&0u16.to_le_bytes()); // colInpCol=0 -> A2
        let result = XlsReader::parse_table_record(&data);
        let (_, _, input1, input2) = result.unwrap();
        assert_eq!(input1, "");
        assert_eq!(input2, "A2");
    }

    #[test]
    fn test_parse_table_record_too_short() {
        let data = vec![0u8; 10]; // less than 16
        assert!(XlsReader::parse_table_record(&data).is_none());
    }
}
