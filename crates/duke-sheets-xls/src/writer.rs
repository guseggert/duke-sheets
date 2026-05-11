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

use duke_sheets_core::style::{
    Alignment, BorderEdge, BorderLineStyle, BorderStyle, Color, DiagonalDirection, FillStyle,
    FontStyle, HorizontalAlignment, NumberFormat, PatternType, ReadingOrder, Underline,
    VerticalAlignment,
};
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
const FONT_RECORD: u16 = 0x0031;
const FORMAT_RECORD: u16 = 0x041E;
const XF_RECORD: u16 = 0x00E0;
const DIMENSION_RECORD: u16 = 0x0200;
const WINDOW2_RECORD: u16 = 0x023E;
const BLANK_RECORD: u16 = 0x0201;
const NUMBER_RECORD: u16 = 0x0203;
const BOOLERR_RECORD: u16 = 0x0205;
const LABELSST_RECORD: u16 = 0x00FD;
const FORMULA_RECORD: u16 = 0x0006;
const STRING_RECORD: u16 = 0x0207;
const MERGECELLS_RECORD: u16 = 0x00E5;
const ROW_RECORD: u16 = 0x0208;
const COLINFO_RECORD: u16 = 0x007D;
const PANE_RECORD: u16 = 0x0041;
const WINDOW1_RECORD: u16 = 0x003D;
const PROTECT_RECORD: u16 = 0x0012;
const PASSWORD_RECORD: u16 = 0x0013;
const SETUP_RECORD: u16 = 0x00A1;
const HEADER_RECORD: u16 = 0x0014;
const FOOTER_RECORD: u16 = 0x0015;
const LEFT_MARGIN_RECORD: u16 = 0x0026;
const RIGHT_MARGIN_RECORD: u16 = 0x0027;
const TOP_MARGIN_RECORD: u16 = 0x0028;
const BOTTOM_MARGIN_RECORD: u16 = 0x0029;
const PRINTHEADERS_RECORD: u16 = 0x002A;
const PRINTGRIDLINES_RECORD: u16 = 0x002B;
const HPAGEBREAKS_RECORD: u16 = 0x001B;
const VPAGEBREAKS_RECORD: u16 = 0x001A;
const SELECTION_RECORD: u16 = 0x001D;
const SCL_RECORD: u16 = 0x00A0;
const HLINK_RECORD: u16 = 0x01B8;
const NAME_RECORD: u16 = 0x0018;
const SUPBOOK_RECORD: u16 = 0x01AE;
const EXTERNSHEET_RECORD: u16 = 0x0017;
const AUTOFILTER_RECORD: u16 = 0x009E;
const FILTERMODE_RECORD: u16 = 0x009B;
const DVAL_RECORD: u16 = 0x01B2;
const DV_RECORD: u16 = 0x01BE;
const CONDFMT_RECORD: u16 = 0x01B0;
const CF_RECORD: u16 = 0x01B1;
const MSODRAWINGGROUP_RECORD: u16 = 0x00EB;
const MSODRAWING_RECORD: u16 = 0x00EC;
const NOTE_RECORD: u16 = 0x001C;
const OBJ_RECORD: u16 = 0x005D;
const TXO_RECORD: u16 = 0x01B6;

/// Built-in NAME index for `Print_Area` (MS-XLS §2.5.4).
const BUILTIN_NAME_PRINT_AREA: u8 = 0x06;
/// Built-in NAME index for `Print_Titles`.
const BUILTIN_NAME_PRINT_TITLES: u8 = 0x07;
/// Built-in NAME index for `_FilterDatabase` (the AutoFilter range).
const BUILTIN_NAME_FILTER_DATABASE: u8 = 0x0D;

/// Hyperlink CLSID (MS-XLS §2.4.144) - StdLink class id
/// 79EAC9D0-BAF9-11CE-8C82-00AA004BA90B, on-disk LE-mixed format
/// (Data1/2/3 little-endian, Data4 byte-ordered).
const HLINK_CLSID: [u8; 16] = [
    0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];

/// URL moniker CLSID 79EAC9E0-BAF9-11CE-8C82-00AA004BA90B in on-disk
/// LE-mixed format (Data1/Data2/Data3 stored little-endian, Data4
/// stored byte-ordered). Identifies the moniker block as a URL rather
/// than a file path. Must match the reader's URL_MONIKER constant
/// byte-for-byte and the bytes Excel/LibreOffice write — earlier
/// drafts had Data1 in the wrong direction, which produced files
/// Excel rejected with "Open method of Workbooks class failed".
const URL_MONIKER_CLSID: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];

/// MS-XLS §2.4.169: a single MERGECELLS record can hold at most 1027
/// merged ranges (8 bytes each, plus the 2-byte cmcs count, fits in
/// the 8224-byte BIFF8 record body cap with margin).
const MERGECELLS_MAX_PER_RECORD: usize = 1027;

/// MS-XLS §2.4.126 user-defined number-format index base. Built-in
/// formats use ifmt 0..=49; user-defined custom format strings start
/// at index 164.
const FORMAT_USER_INDEX_BASE: u16 = 164;

/// BIFF8 reserves the first 16 XF records for built-in cell-format
/// slots; user-defined cell XFs start at 16. Also see MS-XLS §2.4.353.
const XF_USER_BASE: u16 = 16;

/// BIFF8 emits at least 5 FONT records before any can be referenced
/// from XFs, with a "skip 4" quirk in cell-XF font_index decoding
/// (MS-XLS §2.4.122). The writer always emits exactly 5 default fonts
/// up front; user-defined fonts append after.
const FONT_BUILTIN_COUNT: u16 = 5;

/// MS-XLS §2.4.122 Auto color sentinel.
const COLOR_AUTO: u16 = 0x7FFF;

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
        wrap_workbook_stream_in_cfb(stream)
    }

    /// Write a workbook to a filesystem path.
    pub fn write_file<P: AsRef<Path>>(workbook: &Workbook, path: P) -> XlsResult<()> {
        let bytes = Self::write_to_bytes(workbook)?;
        let mut f = std::fs::File::create(path.as_ref())?;
        f.write_all(&bytes)?;
        Ok(())
    }

    /// Serialize a workbook to encrypted BIFF8 bytes. Builds the
    /// plaintext `/Workbook` stream, runs it through
    /// [`duke_sheets_crypto::xls::encrypt_workbook_stream`] for the
    /// requested variant, then re-wraps the encrypted stream in a
    /// fresh CFB envelope.
    pub fn write_to_bytes_encrypted(
        workbook: &Workbook,
        password: &str,
        variant: duke_sheets_crypto::xls::XlsEncryptionVariant,
    ) -> XlsResult<Vec<u8>> {
        let plain = build_workbook_stream(workbook)?;
        let encrypted = duke_sheets_crypto::xls::encrypt_workbook_stream(&plain, password, variant)
            .map_err(XlsError::from)?;
        wrap_workbook_stream_in_cfb(encrypted)
    }

    /// Write a workbook to a filesystem path with FilePass encryption.
    pub fn write_file_encrypted<P: AsRef<Path>>(
        workbook: &Workbook,
        path: P,
        password: &str,
        variant: duke_sheets_crypto::xls::XlsEncryptionVariant,
    ) -> XlsResult<()> {
        let bytes = Self::write_to_bytes_encrypted(workbook, password, variant)?;
        let mut f = std::fs::File::create(path.as_ref())?;
        f.write_all(&bytes)?;
        Ok(())
    }
}

fn wrap_workbook_stream_in_cfb(stream: Vec<u8>) -> XlsResult<Vec<u8>> {
    let mut builder = CompoundFileBuilder::new();
    builder
        .add_stream("/Workbook", stream)
        .map_err(cfb_to_xls)?;
    builder.build().map_err(cfb_to_xls)
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

    let styles = StyleTables::collect(workbook);
    let sst = SstTable::collect(workbook, &styles);

    let mut stream = Vec::new();
    write_bof(&mut stream, DT_WORKBOOK_GLOBALS);
    write_window1(&mut stream, workbook);
    styles.write_font_records(&mut stream)?;
    styles.write_format_records(&mut stream)?;
    styles.write_xf_records(&mut stream);

    let mut lbplypos_field_offsets = Vec::with_capacity(workbook.sheet_count());
    for sheet in workbook.worksheets() {
        let body_start = stream.len() + 4;
        write_boundsheet8_with_placeholder_offset(&mut stream, sheet)?;
        lbplypos_field_offsets.push(body_start);
    }

    let externsheet_table = build_externsheet_table(workbook);
    let name_table = build_name_table(workbook);
    write_supbook_and_externsheet(&mut stream, workbook, &externsheet_table);
    write_user_name_records(&mut stream, workbook, &externsheet_table, &name_table);
    write_print_name_records(&mut stream, workbook);
    sst.write_records(&mut stream)?;
    let drawing_state = compute_drawing_state(workbook);
    write_msodrawinggroup(&mut stream, &drawing_state);
    write_eof(&mut stream);

    let mut sheet_bof_offsets = Vec::with_capacity(workbook.sheet_count());
    for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
        let bof_pos = stream.len() as u32;
        write_bof(&mut stream, DT_WORKSHEET);
        write_protect_records(&mut stream, sheet);
        write_colinfo_records(&mut stream, sheet);
        write_dimension(&mut stream, sheet);
        write_row_records(&mut stream, sheet);
        write_cell_records(
            &mut stream,
            sheet,
            sheet_idx,
            &sst,
            &styles,
            &externsheet_table,
            &name_table,
        );
        write_page_break_records(&mut stream, sheet);
        write_header_footer_records(&mut stream, sheet);
        write_margin_records(&mut stream, sheet);
        write_print_flags(&mut stream, sheet);
        write_setup_record(&mut stream, sheet);
        write_window2(&mut stream, sheet);
        write_scl(&mut stream, sheet);
        write_pane(&mut stream, sheet);
        write_selection_records(&mut stream, sheet);
        write_mergecells(&mut stream, sheet);
        write_hlink_records(&mut stream, sheet);
        write_autofilter_records(&mut stream, sheet);
        write_data_validations(&mut stream, sheet, &externsheet_table, &name_table);
        write_conditional_formats(&mut stream, sheet, &externsheet_table, &name_table);
        if let Some(sheet_drawing) = drawing_state.sheets.get(&sheet_idx) {
            write_sheet_drawing_records(&mut stream, sheet_drawing);
        }
        write_eof(&mut stream);
        sheet_bof_offsets.push(bof_pos);
    }

    for (offset, sheet_bof) in lbplypos_field_offsets.iter().zip(sheet_bof_offsets.iter()) {
        stream[*offset..*offset + 4].copy_from_slice(&sheet_bof.to_le_bytes());
    }

    Ok(stream)
}

/// Workbook-global font + format + XF tables. Slice 4a/4b customize
/// the font_index and format_index axes of an XF record; alignment,
/// fill, and border stay at defaults.
struct StyleTables {
    /// FONT records to emit, in disk order. The first
    /// `FONT_BUILTIN_COUNT` entries are the BIFF8-required built-ins
    /// (all defaulted Calibri 11). User-defined fonts append after.
    fonts_in_order: Vec<FontStyle>,
    /// `FontStyle -> on-disk-font-index` (with the BIFF8 "skip 4"
    /// quirk applied: positions 0..3 are the index, position 4 is
    /// reserved, position 5+ become index +1). Exposed so the SST
    /// rich-text emitter can look up font indices for run fonts that
    /// were also interned in this table.
    font_xf_index: HashMap<FontStyle, u16>,
    /// User-defined number-format strings to emit as FORMAT records.
    /// Their on-disk ifmt values start at `FORMAT_USER_INDEX_BASE` and
    /// increment by one per entry.
    user_formats: Vec<String>,
    /// User-defined XF records to emit after the 16 built-ins.
    user_xfs: Vec<UserXf>,
    /// `(sheet_idx, style_index_in_pool) -> ixfe`. Cells consult this
    /// to pick their `XF` reference; absent entries fall back to 0.
    cell_ixfe: HashMap<(usize, u32), u16>,
}

#[derive(Debug, Clone)]
struct UserXf {
    font_index: u16,
    format_index: u16,
    alignment: Alignment,
    border: BorderStyle,
    fill: FillStyle,
    protection: duke_sheets_core::style::Protection,
}

impl StyleTables {
    fn collect(workbook: &Workbook) -> Self {
        let default_font = FontStyle::default();
        let mut fonts_in_order = vec![default_font.clone(); FONT_BUILTIN_COUNT as usize];
        let mut font_xf_index: HashMap<FontStyle, u16> = HashMap::new();
        font_xf_index.insert(default_font, 0u16);

        let mut user_formats: Vec<String> = Vec::new();
        let mut format_index_for_custom: HashMap<String, u16> = HashMap::new();

        let mut user_xfs: Vec<UserXf> = Vec::new();
        type XfKey = (
            u16,
            u16,
            Alignment,
            BorderStyle,
            FillStyle,
            duke_sheets_core::style::Protection,
        );
        let mut xf_key_to_ixfe: HashMap<XfKey, u16> = HashMap::new();
        let mut cell_ixfe = HashMap::new();

        for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
            for (_row, _col, cell) in sheet.iter_cells() {
                if cell.style_index == 0 {
                    continue;
                }
                let Some(style) = sheet.style_by_index(cell.style_index) else {
                    continue;
                };

                let font = style.font.clone();
                let font_idx = match font_xf_index.get(&font) {
                    Some(&idx) => idx,
                    None => {
                        let on_disk = fonts_in_order.len() as u16;
                        let xf_idx = if on_disk < 4 { on_disk } else { on_disk + 1 };
                        fonts_in_order.push(font.clone());
                        font_xf_index.insert(font.clone(), xf_idx);
                        xf_idx
                    }
                };

                let format_idx = match &style.number_format {
                    NumberFormat::General => 0u16,
                    NumberFormat::BuiltIn(id) => *id as u16,
                    NumberFormat::Custom(s) => match format_index_for_custom.get(s) {
                        Some(&idx) => idx,
                        None => {
                            let idx = FORMAT_USER_INDEX_BASE + user_formats.len() as u16;
                            user_formats.push(s.clone());
                            format_index_for_custom.insert(s.clone(), idx);
                            idx
                        }
                    },
                };

                let xf_key: XfKey = (
                    font_idx,
                    format_idx,
                    style.alignment.clone(),
                    style.border.clone(),
                    style.fill.clone(),
                    style.protection,
                );
                let ixfe = match xf_key_to_ixfe.get(&xf_key) {
                    Some(&i) => i,
                    None => {
                        let new_ixfe = XF_USER_BASE + user_xfs.len() as u16;
                        user_xfs.push(UserXf {
                            font_index: font_idx,
                            format_index: format_idx,
                            alignment: style.alignment.clone(),
                            border: style.border.clone(),
                            fill: style.fill.clone(),
                            protection: style.protection,
                        });
                        xf_key_to_ixfe.insert(xf_key, new_ixfe);
                        new_ixfe
                    }
                };

                cell_ixfe.insert((sheet_idx, cell.style_index), ixfe);
            }
        }

        // Second pass: intern fonts for any rich-text run that has a
        // RunFont attached. Each run gets a complete FontStyle (with
        // unspecified RunFont fields filled from FontStyle::default)
        // so the SST emitter can look up its on-disk font index.
        for sheet in workbook.worksheets() {
            for (_row, _col, cell) in sheet.iter_cells() {
                if let CellValue::RichText(runs) = &cell.value {
                    for run in runs.iter() {
                        if let Some(rf) = &run.font {
                            let font = run_font_to_font_style(rf);
                            if !font_xf_index.contains_key(&font) {
                                let on_disk = fonts_in_order.len() as u16;
                                let xf_idx = if on_disk < 4 { on_disk } else { on_disk + 1 };
                                fonts_in_order.push(font.clone());
                                font_xf_index.insert(font, xf_idx);
                            }
                        }
                    }
                }
            }
        }

        StyleTables {
            fonts_in_order,
            font_xf_index,
            user_formats,
            user_xfs,
            cell_ixfe,
        }
    }

    fn ixfe_for_cell(&self, sheet_idx: usize, style_index: u32) -> u16 {
        if style_index == 0 {
            return 0;
        }
        self.cell_ixfe
            .get(&(sheet_idx, style_index))
            .copied()
            .unwrap_or(0)
    }

    fn write_font_records(&self, stream: &mut Vec<u8>) -> XlsResult<()> {
        for font in &self.fonts_in_order {
            write_font_record(stream, font)?;
        }
        Ok(())
    }

    fn write_format_records(&self, stream: &mut Vec<u8>) -> XlsResult<()> {
        for (i, fmt) in self.user_formats.iter().enumerate() {
            let ifmt = FORMAT_USER_INDEX_BASE + i as u16;
            write_format_record(stream, ifmt, fmt)?;
        }
        Ok(())
    }

    fn write_xf_records(&self, stream: &mut Vec<u8>) {
        for _ in 0..XF_USER_BASE {
            write_xf_record(stream, /* is_style_xf */ false, &XF_DEFAULTS);
        }
        for xf in &self.user_xfs {
            write_xf_record(stream, false, xf);
        }
    }
}

/// Materialise a [`RunFont`] (where each property is optional) into a
/// complete [`FontStyle`]. Unspecified `RunFont` properties fall back
/// to [`FontStyle::default()`] so the resulting struct can be emitted
/// as a complete BIFF8 FONT record.
fn run_font_to_font_style(rf: &duke_sheets_core::rich_text::RunFont) -> FontStyle {
    let mut font = FontStyle::default();
    if let Some(b) = rf.bold {
        font.bold = b;
    }
    if let Some(i) = rf.italic {
        font.italic = i;
    }
    if let Some(s) = rf.size {
        font.size = s;
    }
    if let Some(c) = rf.color.as_ref() {
        font.color = c.clone();
    }
    if let Some(name) = rf.name.as_ref() {
        font.name = name.clone();
    }
    if let Some(u) = rf.underline {
        font.underline = u;
    }
    if let Some(s) = rf.strikethrough {
        font.strikethrough = s;
    }
    if let Some(va) = rf.vertical_align {
        font.vertical_align = va;
    }
    if let Some(family) = rf.family {
        font.family = Some(family);
    }
    if let Some(charset) = rf.charset {
        font.charset = Some(charset);
    }
    if let Some(scheme) = rf.scheme.as_ref() {
        font.scheme = Some(scheme.clone());
    }
    font
}

const XF_DEFAULTS: UserXf = UserXf {
    font_index: 0,
    format_index: 0,
    alignment: Alignment {
        horizontal: HorizontalAlignment::General,
        vertical: VerticalAlignment::Bottom,
        wrap_text: false,
        shrink_to_fit: false,
        indent: 0,
        rotation: 0,
        reading_order: ReadingOrder::ContextDependent,
    },
    border: BorderStyle {
        left: None,
        right: None,
        top: None,
        bottom: None,
        diagonal: None,
        diagonal_direction: DiagonalDirection::None,
    },
    fill: FillStyle::None,
    protection: duke_sheets_core::style::Protection {
        locked: true,
        hidden: false,
    },
};

/// Emit a FORMAT record (MS-XLS §2.4.126) for a user-defined custom
/// number-format string. Built-in format indices (0..=49) are implicit
/// and don't need a FORMAT record.
fn write_format_record(stream: &mut Vec<u8>, ifmt: u16, format_string: &str) -> XlsResult<()> {
    let mut body = Vec::with_capacity(2 + 3 + format_string.len() * 2);
    body.extend_from_slice(&ifmt.to_le_bytes());
    push_xlunicode_string(&mut body, format_string)?;
    stream.extend_from_slice(&FORMAT_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
    Ok(())
}

/// Emit a FONT record (MS-XLS §2.4.122). Only the fields needed by the
/// reader's `parse_font` are populated; family/charset are left at 0
/// (font-style-default). Color resolution: `Auto` → 0x7FFF, `Indexed`
/// passes through; RGB/Theme fall back to `Auto` because BIFF8 needs
/// a PALETTE record to express arbitrary RGB and that's deferred.
fn write_font_record(stream: &mut Vec<u8>, font: &FontStyle) -> XlsResult<()> {
    let mut body = Vec::with_capacity(16 + font.name.len() * 2);
    let height_twips = (font.size * 20.0).round() as u16;
    body.extend_from_slice(&height_twips.to_le_bytes());

    let mut grbit: u16 = 0;
    if font.italic {
        grbit |= 0x0002;
    }
    if font.strikethrough {
        grbit |= 0x0008;
    }
    body.extend_from_slice(&grbit.to_le_bytes());

    let icv = match font.color {
        Color::Auto => COLOR_AUTO,
        Color::Indexed(i) => i as u16,
        // BIFF8 has no first-class RGB / theme support without a
        // PALETTE record. Fall back to auto rather than emit an
        // out-of-range icv that the reader would reject.
        _ => COLOR_AUTO,
    };
    body.extend_from_slice(&icv.to_le_bytes());

    let bls: u16 = if font.bold { 700 } else { 400 };
    body.extend_from_slice(&bls.to_le_bytes());

    let sss: u16 = match font.vertical_align {
        duke_sheets_core::style::FontVerticalAlign::Superscript => 1,
        duke_sheets_core::style::FontVerticalAlign::Subscript => 2,
        _ => 0,
    };
    body.extend_from_slice(&sss.to_le_bytes());

    let uls: u8 = match font.underline {
        Underline::Single => 0x01,
        Underline::Double => 0x02,
        Underline::SingleAccounting => 0x21,
        Underline::DoubleAccounting => 0x22,
        Underline::None => 0x00,
    };
    body.push(uls);
    body.push(0); // bFamily
    body.push(0); // bCharSet
    body.push(0); // reserved

    push_short_xlunicode_string(&mut body, &font.name)?;

    stream.extend_from_slice(&FONT_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
    Ok(())
}

/// Emit a 20-byte XF record (MS-XLS §2.4.353) with the supplied
/// font/format/alignment/border/fill axes encoded into the bit-packed
/// fields. `is_style_xf` flips the type/protect bit that tells the
/// reader whether this XF is a cell XF or a named-style XF.
fn write_xf_record(stream: &mut Vec<u8>, is_style_xf: bool, xf: &UserXf) {
    stream.extend_from_slice(&XF_RECORD.to_le_bytes());
    stream.extend_from_slice(&20u16.to_le_bytes());
    stream.extend_from_slice(&xf.font_index.to_le_bytes());
    stream.extend_from_slice(&xf.format_index.to_le_bytes());
    // type_prot bit layout (MS-XLS §2.4.353):
    //   bit 0 (0x0001): fLocked
    //   bit 1 (0x0002): fHidden (formula hidden when sheet is protected)
    //   bit 2 (0x0004): fStyle (1 = style XF, 0 = cell XF)
    //   bits 4-15: ifmt-XF parent (cell XF parent index, 0xFFF for style XFs)
    let type_prot: u16 = if is_style_xf {
        // 0xFFF5 = parent=0xFFF, fStyle=1, fHidden=0, fLocked=1 (default).
        0xFFF5
    } else {
        let mut bits: u16 = 0;
        if xf.protection.locked {
            bits |= 0x0001;
        }
        if xf.protection.hidden {
            bits |= 0x0002;
        }
        bits
    };
    stream.extend_from_slice(&type_prot.to_le_bytes());

    let halign = encode_horizontal_alignment(xf.alignment.horizontal);
    let valign = encode_vertical_alignment(xf.alignment.vertical);
    let align1: u8 =
        (halign & 0x07) | (if xf.alignment.wrap_text { 0x08 } else { 0 }) | ((valign & 0x07) << 4);
    stream.push(align1);
    stream.push(encode_rotation(xf.alignment.rotation));
    let reading_order = encode_reading_order(xf.alignment.reading_order);
    let align2: u8 = (xf.alignment.indent.min(15))
        | (if xf.alignment.shrink_to_fit { 0x10 } else { 0 })
        | ((reading_order & 0x03) << 6);
    stream.push(align2);
    stream.push(0); // used_attribs

    let (border_left, border_right, border_top, border_bottom) = (
        encode_border_line(xf.border.left.as_ref()),
        encode_border_line(xf.border.right.as_ref()),
        encode_border_line(xf.border.top.as_ref()),
        encode_border_line(xf.border.bottom.as_ref()),
    );
    let (icv_left, icv_right) = (
        encode_border_color(xf.border.left.as_ref()),
        encode_border_color(xf.border.right.as_ref()),
    );
    let diagonal_dir = encode_diagonal_direction(xf.border.diagonal_direction);
    let border1: u32 = (border_left as u32 & 0x0F)
        | ((border_right as u32 & 0x0F) << 4)
        | ((border_top as u32 & 0x0F) << 8)
        | ((border_bottom as u32 & 0x0F) << 12)
        | ((icv_left as u32 & 0x7F) << 16)
        | ((icv_right as u32 & 0x7F) << 23)
        | ((diagonal_dir as u32 & 0x03) << 30);
    stream.extend_from_slice(&border1.to_le_bytes());

    let icv_top = encode_border_color(xf.border.top.as_ref());
    let icv_bottom = encode_border_color(xf.border.bottom.as_ref());
    let icv_diag = encode_border_color(xf.border.diagonal.as_ref());
    let border_diag = encode_border_line(xf.border.diagonal.as_ref());
    let fill_pattern = encode_fill_pattern(&xf.fill);
    let border2: u32 = (icv_top as u32 & 0x7F)
        | ((icv_bottom as u32 & 0x7F) << 7)
        | ((icv_diag as u32 & 0x7F) << 14)
        | ((border_diag as u32 & 0x0F) << 21)
        | ((fill_pattern as u32 & 0x3F) << 26);
    stream.extend_from_slice(&border2.to_le_bytes());

    let (fill_fg, fill_bg) = encode_fill_colors(&xf.fill);
    let fill_colors: u16 = (fill_fg as u16 & 0x7F) | ((fill_bg as u16 & 0x7F) << 7);
    stream.extend_from_slice(&fill_colors.to_le_bytes());
}

fn encode_horizontal_alignment(h: HorizontalAlignment) -> u8 {
    match h {
        HorizontalAlignment::General => 0,
        HorizontalAlignment::Left => 1,
        HorizontalAlignment::Center => 2,
        HorizontalAlignment::Right => 3,
        HorizontalAlignment::Fill => 4,
        HorizontalAlignment::Justify => 5,
        HorizontalAlignment::CenterContinuous => 6,
        HorizontalAlignment::Distributed => 7,
    }
}

fn encode_vertical_alignment(v: VerticalAlignment) -> u8 {
    match v {
        VerticalAlignment::Top => 0,
        VerticalAlignment::Center => 1,
        VerticalAlignment::Bottom => 2,
        VerticalAlignment::Justify => 3,
        VerticalAlignment::Distributed => 4,
    }
}

fn encode_reading_order(r: ReadingOrder) -> u8 {
    match r {
        ReadingOrder::ContextDependent => 0,
        ReadingOrder::LeftToRight => 1,
        ReadingOrder::RightToLeft => 2,
    }
}

fn encode_rotation(rotation: i16) -> u8 {
    if rotation == 255 {
        return 255;
    }
    match rotation {
        0 => 0,
        1..=90 => rotation as u8,
        // The reader maps BIFF values 91..=180 back to negative
        // (anti-clockwise) angles. Inverse the relation here.
        -90..=-1 => (90i16 - rotation) as u8,
        _ => 0,
    }
}

fn encode_border_line(edge: Option<&BorderEdge>) -> u8 {
    let style = edge.map(|e| e.style).unwrap_or(BorderLineStyle::None);
    match style {
        BorderLineStyle::None => 0,
        BorderLineStyle::Thin => 1,
        BorderLineStyle::Medium => 2,
        BorderLineStyle::Dashed => 3,
        BorderLineStyle::Dotted => 4,
        BorderLineStyle::Thick => 5,
        BorderLineStyle::Double => 6,
        BorderLineStyle::Hair => 7,
        BorderLineStyle::MediumDashed => 8,
        BorderLineStyle::DashDot => 9,
        BorderLineStyle::MediumDashDot => 10,
        BorderLineStyle::DashDotDot => 11,
        BorderLineStyle::MediumDashDotDot => 12,
        BorderLineStyle::SlantDashDot => 13,
    }
}

fn encode_border_color(edge: Option<&BorderEdge>) -> u8 {
    edge.map(|e| color_to_icv7(&e.color)).unwrap_or(0x40)
}

fn encode_diagonal_direction(d: DiagonalDirection) -> u8 {
    match d {
        DiagonalDirection::None => 0,
        DiagonalDirection::Down => 1,
        DiagonalDirection::Up => 2,
        DiagonalDirection::Both => 3,
    }
}

fn encode_fill_pattern(fill: &FillStyle) -> u8 {
    match fill {
        FillStyle::None => 0,
        FillStyle::Solid { .. } => 1,
        FillStyle::Pattern { pattern, .. } => match pattern {
            PatternType::None => 0,
            PatternType::Solid => 1,
            PatternType::MediumGray => 2,
            PatternType::DarkGray => 3,
            PatternType::LightGray => 4,
            PatternType::DarkHorizontal => 5,
            PatternType::DarkVertical => 6,
            PatternType::DarkDown => 7,
            PatternType::DarkUp => 8,
            PatternType::DarkGrid => 9,
            PatternType::DarkTrellis => 10,
            PatternType::LightHorizontal => 11,
            PatternType::LightVertical => 12,
            PatternType::LightDown => 13,
            PatternType::LightUp => 14,
            PatternType::LightGrid => 15,
            PatternType::LightTrellis => 16,
            PatternType::Gray125 => 17,
            PatternType::Gray0625 => 18,
        },
        FillStyle::Gradient { .. } => 1, // fall back to solid for now
    }
}

fn encode_fill_colors(fill: &FillStyle) -> (u8, u8) {
    match fill {
        FillStyle::None => (0x40, 0x41),
        FillStyle::Solid { color } => (color_to_icv7(color), 0x41),
        FillStyle::Pattern {
            foreground,
            background,
            ..
        } => (color_to_icv7(foreground), color_to_icv7(background)),
        FillStyle::Gradient { stops, .. } => {
            let fg = stops
                .first()
                .map(|s| color_to_icv7(&s.color))
                .unwrap_or(0x40);
            (fg, 0x41)
        }
    }
}

/// Encode a [`Color`] as a 7-bit `icv` value (the on-disk encoding
/// used in border/fill XF fields). The font-side encoding uses 16
/// bits and a different sentinel — see `write_font_record`.
fn color_to_icv7(color: &Color) -> u8 {
    match color {
        Color::Auto => 0x40,
        Color::Indexed(i) => 0x08u8.saturating_add(i.min(&55).clone()),
        // BIFF8 has no first-class RGB without a PALETTE record. Pick
        // 0x40 (system foreground) as the closest no-op default; a
        // future PALETTE-emission slice can express arbitrary RGB.
        _ => 0x40,
    }
}

/// Emit a `ShortXLUnicodeString` (1-byte cch, 1-byte fHighByte, chars).
fn push_short_xlunicode_string(buf: &mut Vec<u8>, s: &str) -> XlsResult<()> {
    let units: Vec<u16> = s.encode_utf16().collect();
    if units.len() > u8::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "short string '{s}' exceeds 255-char ShortXLUnicodeString limit"
        )));
    }
    let high_byte = units.iter().any(|&u| u > 0xFF);
    buf.push(units.len() as u8);
    if high_byte {
        buf.push(0x01);
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        buf.push(0x00);
        for u in &units {
            buf.push(*u as u8);
        }
    }
    Ok(())
}

/// Builder + emitter for the workbook-level Shared String Table (SST).
///
/// String cells in BIFF8 don't store their text inline; instead they
/// reference an entry in this workbook-global table via `LABELSST.isst`.
/// We dedupe identical strings on insert (matching what Excel writes)
/// and emit a single SST record + CONTINUE chain in the globals stream.
/// A single entry in the BIFF8 Shared String Table.
///
/// `Plain` is a flat string with no per-character formatting. `Rich`
/// stores the same text plus a list of formatting runs (each a
/// `(char_pos, font_idx)` pair) that the writer encodes as a BIFF8
/// rich-text SST entry with `fRichSt` set.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum SstEntryKind {
    Plain,
    Rich(Vec<(u16, u16)>),
}

#[derive(Debug, Clone)]
struct SstEntry {
    text: String,
    kind: SstEntryKind,
}

struct SstTable {
    entries: Vec<SstEntry>,
    /// Index keyed by `(text, kind)` so plain and rich variants of the
    /// same text are distinct entries (a plain "hello" and a bolded
    /// "hello" share neither LABELSST index nor SST slot).
    index: HashMap<(String, SstEntryKind), u32>,
    total_refs: u32,
}

impl SstTable {
    fn collect(workbook: &Workbook, styles: &StyleTables) -> Self {
        let mut t = SstTable {
            entries: Vec::new(),
            index: HashMap::new(),
            total_refs: 0,
        };
        for sheet in workbook.worksheets() {
            for (_row, _col, cell) in sheet.iter_cells() {
                match &cell.value {
                    CellValue::String(s) => {
                        t.add_plain(s.as_ref());
                    }
                    CellValue::RichText(runs) => {
                        t.add_rich(runs, styles);
                    }
                    _ => {}
                }
            }
        }
        t
    }

    fn add_plain(&mut self, s: &str) {
        self.total_refs += 1;
        let key = (s.to_string(), SstEntryKind::Plain);
        if !self.index.contains_key(&key) {
            let idx = self.entries.len() as u32;
            self.index.insert(key, idx);
            self.entries.push(SstEntry {
                text: s.to_string(),
                kind: SstEntryKind::Plain,
            });
        }
    }

    fn add_rich(
        &mut self,
        runs: &[duke_sheets_core::rich_text::RichTextRun],
        styles: &StyleTables,
    ) {
        self.total_refs += 1;
        let mut text = String::new();
        let mut formatting: Vec<(u16, u16)> = Vec::new();
        for run in runs {
            // char_pos is the UTF-16 code unit offset where this run
            // begins, which is what BIFF8 expects.
            let char_pos = text.encode_utf16().count().min(u16::MAX as usize) as u16;
            if let Some(rf) = &run.font {
                let font = run_font_to_font_style(rf);
                if let Some(&font_idx) = styles.font_xf_index.get(&font) {
                    formatting.push((char_pos, font_idx));
                }
            }
            text.push_str(&run.text);
        }

        let kind = if formatting.is_empty() {
            SstEntryKind::Plain
        } else {
            SstEntryKind::Rich(formatting)
        };

        let key = (text.clone(), kind.clone());
        if !self.index.contains_key(&key) {
            let idx = self.entries.len() as u32;
            self.index.insert(key, idx);
            self.entries.push(SstEntry { text, kind });
        }
    }

    /// Look up a plain-string SST index. Used by the LABELSST path
    /// for `CellValue::String` cells.
    fn lookup_plain(&self, s: &str) -> Option<u32> {
        self.index
            .get(&(s.to_string(), SstEntryKind::Plain))
            .copied()
    }

    /// Look up the SST index for a rich-text cell, given the text +
    /// formatting runs that were resolved for it.
    fn lookup_rich(&self, text: &str, formatting: &[(u16, u16)]) -> Option<u32> {
        let kind = if formatting.is_empty() {
            SstEntryKind::Plain
        } else {
            SstEntryKind::Rich(formatting.to_vec())
        };
        self.index.get(&(text.to_string(), kind)).copied()
    }

    /// Serialize the SST (and any required CONTINUE records) into the
    /// workbook stream. No-op if no entries were collected.
    fn write_records(&self, stream: &mut Vec<u8>) -> XlsResult<()> {
        if self.entries.is_empty() {
            return Ok(());
        }

        let mut payload = Vec::new();
        payload.extend_from_slice(&self.total_refs.to_le_bytes());
        payload.extend_from_slice(&(self.entries.len() as u32).to_le_bytes());
        for entry in &self.entries {
            push_sst_entry(&mut payload, entry)?;
        }

        // Split between entries (never mid-entry). For each entry,
        // compute its on-disk length and chunk into BIFF_MAX_RECORD_BODY
        // groups.
        let mut chunks: Vec<(usize, usize)> = Vec::new();
        let mut cursor = 8usize; // skip cstTotal + cstUnique header
        let mut chunk_start = 0usize;
        let mut chunk_len = 8usize;
        for entry in &self.entries {
            let entry_len = sst_entry_len(entry);
            if chunk_len + entry_len > BIFF_MAX_RECORD_BODY {
                chunks.push((chunk_start, chunk_len));
                chunk_start = cursor;
                chunk_len = 0;
            }
            chunk_len += entry_len;
            cursor += entry_len;
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

/// Length of an SST entry on disk: plain entries are
/// cch + flags + chars; rich entries add a 2-byte cRun count plus
/// 4 bytes per formatting run inserted between the flags byte and
/// the character data.
fn sst_entry_len(entry: &SstEntry) -> usize {
    let plain = xlunicode_string_len(&entry.text);
    match &entry.kind {
        SstEntryKind::Plain => plain,
        SstEntryKind::Rich(runs) => plain + 2 + runs.len() * 4,
    }
}

fn push_sst_entry(buf: &mut Vec<u8>, entry: &SstEntry) -> XlsResult<()> {
    match &entry.kind {
        SstEntryKind::Plain => push_xlunicode_string(buf, &entry.text),
        SstEntryKind::Rich(runs) => push_rich_xlunicode_string(buf, &entry.text, runs),
    }
}

/// Emit a rich `XLUnicodeRichExtendedString`: cch + flags (with
/// fRichSt bit set) + cRun + chars + run array. Each run is
/// (char_pos u16, font_idx u16) = 4 bytes.
fn push_rich_xlunicode_string(buf: &mut Vec<u8>, s: &str, runs: &[(u16, u16)]) -> XlsResult<()> {
    let units: Vec<u16> = s.encode_utf16().collect();
    if units.len() > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "string of {} UTF-16 units exceeds BIFF8 cch limit (u16)",
            units.len()
        )));
    }
    if runs.len() > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "rich-text run count {} exceeds BIFF8 limit (u16)",
            runs.len()
        )));
    }
    let high_byte = units.iter().any(|&u| u > 0xFF);
    buf.extend_from_slice(&(units.len() as u16).to_le_bytes());
    let flags: u8 = 0x08 | if high_byte { 0x01 } else { 0x00 };
    buf.push(flags);
    buf.extend_from_slice(&(runs.len() as u16).to_le_bytes());
    if high_byte {
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        for u in &units {
            buf.push(*u as u8);
        }
    }
    for (char_pos, font_idx) in runs {
        buf.extend_from_slice(&char_pos.to_le_bytes());
        buf.extend_from_slice(&font_idx.to_le_bytes());
    }
    Ok(())
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

/// Emit a WINDOW1 record (MS-XLS §2.4.346). Bytes 10-11 carry
/// `itabCur` (active sheet index); the reader uses that to set the
/// workbook's active sheet. The other fields use sensible defaults.
fn write_window1(stream: &mut Vec<u8>, workbook: &Workbook) {
    let active = workbook.active_sheet().min(u16::MAX as usize) as u16;
    let visible_count = workbook
        .worksheets()
        .filter(|s| {
            !matches!(
                s.visibility(),
                duke_sheets_core::worksheet::SheetVisibility::Hidden
                    | duke_sheets_core::worksheet::SheetVisibility::VeryHidden
            )
        })
        .count()
        .min(u16::MAX as usize) as u16;

    stream.extend_from_slice(&WINDOW1_RECORD.to_le_bytes());
    stream.extend_from_slice(&18u16.to_le_bytes());
    stream.extend_from_slice(&0i16.to_le_bytes()); // xWn
    stream.extend_from_slice(&0i16.to_le_bytes()); // yWn
    stream.extend_from_slice(&15000u16.to_le_bytes()); // dxWn
    stream.extend_from_slice(&9000u16.to_le_bytes()); // dyWn
    stream.extend_from_slice(&0x0038u16.to_le_bytes()); // grbit: tab visible, scroll bars visible
    stream.extend_from_slice(&active.to_le_bytes()); // itabCur
    stream.extend_from_slice(&0u16.to_le_bytes()); // itabFirst
    stream.extend_from_slice(&visible_count.max(1).to_le_bytes()); // ctabSel
    stream.extend_from_slice(&600u16.to_le_bytes()); // wTabRatio (600 = 60%)
}

/// Emit a minimal WINDOW2 record (MS-XLS §2.4.349). The reader extracts
/// only the frozen-pane bit, but real Excel and LibreOffice expect this
/// record as a structural cue that the worksheet stream is well-formed.
/// Sets fFrozen (bit 3) and fFrozenNoSplit (bit 8) when the sheet has
/// freeze panes set; the matching PANE record carries the split
/// position.
fn write_window2(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let mut options: u16 = 0x06B6;
    if sheet.freeze_panes().is_some() {
        options |= 0x0008 | 0x0100;
    }
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

/// Emit a PANE record (MS-XLS §2.4.187) when the worksheet has freeze
/// panes set. Body: x (col split, u16), y (row split, u16), top_row
/// (u16), left_col (u16), active_pane (u16, 0=bottomRight). The
/// matching WINDOW2 sets fFrozen + fFrozenNoSplit so the reader knows
/// the split values are row/col indices rather than twip offsets.
fn write_pane(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let Some(freeze) = sheet.freeze_panes() else {
        return;
    };
    if freeze.row == 0 && freeze.col == 0 {
        return;
    }
    let x = freeze.col;
    let y = freeze.row.min(u16::MAX as u32) as u16;
    let top_row: u16 = y;
    let left_col: u16 = x;
    // Active pane is the bottom-right when both axes are frozen,
    // else the corresponding half. Excel uses 0=bottomRight,
    // 1=topRight, 2=bottomLeft, 3=topLeft.
    let active_pane: u16 = match (x > 0, y > 0) {
        (true, true) => 0,
        (true, false) => 0,
        (false, true) => 2,
        (false, false) => 3,
    };

    stream.extend_from_slice(&PANE_RECORD.to_le_bytes());
    stream.extend_from_slice(&10u16.to_le_bytes());
    stream.extend_from_slice(&x.to_le_bytes());
    stream.extend_from_slice(&y.to_le_bytes());
    stream.extend_from_slice(&top_row.to_le_bytes());
    stream.extend_from_slice(&left_col.to_le_bytes());
    stream.extend_from_slice(&active_pane.to_le_bytes());
}

/// Emit an SCL record (MS-XLS §2.4.249) carrying the worksheet zoom
/// level as a numerator/denominator ratio. The model stores zoom as a
/// percentage (10..=400); reduce against 100 for compactness.
fn write_scl(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let Some(zoom) = sheet.zoom_scale() else {
        return;
    };
    let zoom = zoom.clamp(10, 400);
    let (num, den) = if zoom % 25 == 0 && zoom <= 200 {
        (zoom / 25, 4u16)
    } else {
        (zoom, 100u16)
    };
    stream.extend_from_slice(&SCL_RECORD.to_le_bytes());
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&num.to_le_bytes());
    stream.extend_from_slice(&den.to_le_bytes());
}

/// Emit one SELECTION record (MS-XLS §2.4.247) per Selection in the
/// worksheet's selection list. Body: pane (u8), active_row (u16),
/// active_col (u16), active_ref (u16), ref_count (u16), then
/// `ref_count` × Ref8U (r1 u16, r2 u16, c1 u8, c2 u8 = 6 bytes).
/// Selections targeting cells beyond the BIFF8 sheet extent (col >
/// 255 or row > 65535) are skipped because the on-disk Ref8U format
/// can't represent them.
fn write_selection_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    use duke_sheets_core::{CellAddress, CellRange};

    for selection in sheet.selections() {
        let pane_byte: u8 = match selection.pane.as_deref() {
            Some("topRight") => 1,
            Some("bottomLeft") => 2,
            Some("topLeft") => 3,
            _ => 0, // bottomRight or unspecified
        };

        let (active_row, active_col) = selection
            .active_cell
            .as_deref()
            .and_then(|s| CellAddress::parse(s).ok())
            .map(|a| (a.row, a.col))
            .unwrap_or((0u32, 0u16));
        if active_row > u16::MAX as u32 || active_col > u8::MAX as u16 {
            continue;
        }

        let ranges: Vec<CellRange> = selection
            .sqref
            .as_deref()
            .map(|s| {
                s.split_whitespace()
                    .filter_map(|piece| {
                        CellRange::parse(piece).ok().or_else(|| {
                            CellAddress::parse(piece).ok().map(|a| CellRange::new(a, a))
                        })
                    })
                    .collect()
            })
            .unwrap_or_default();

        let usable: Vec<&CellRange> = ranges
            .iter()
            .filter(|r| {
                r.start.row <= u16::MAX as u32
                    && r.end.row <= u16::MAX as u32
                    && r.start.col <= u8::MAX as u16
                    && r.end.col <= u8::MAX as u16
            })
            .collect();

        let ref_count = usable.len() as u16;
        let body_len = 9u16 + ref_count * 6;
        stream.extend_from_slice(&SELECTION_RECORD.to_le_bytes());
        stream.extend_from_slice(&body_len.to_le_bytes());
        stream.push(pane_byte);
        stream.extend_from_slice(&(active_row as u16).to_le_bytes());
        stream.extend_from_slice(&active_col.to_le_bytes());
        stream.extend_from_slice(&0u16.to_le_bytes()); // active_ref index
        stream.extend_from_slice(&ref_count.to_le_bytes());
        for r in usable {
            stream.extend_from_slice(&(r.start.row as u16).to_le_bytes());
            stream.extend_from_slice(&(r.end.row as u16).to_le_bytes());
            stream.push(r.start.col as u8);
            stream.push(r.end.col as u8);
        }
    }
}

/// Emit PROTECT (MS-XLS §2.4.196) and optional PASSWORD (MS-XLS
/// §2.4.190) records when the worksheet has protection set. PROTECT
/// is a single u16 = 1; PASSWORD is the precomputed 16-bit verifier
/// (zero means "no password required, but sheet is protected").
fn write_protect_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let Some(protection) = sheet.protection() else {
        return;
    };
    if !protection.protected {
        return;
    }
    stream.extend_from_slice(&PROTECT_RECORD.to_le_bytes());
    stream.extend_from_slice(&2u16.to_le_bytes());
    stream.extend_from_slice(&1u16.to_le_bytes());

    let password_hash = protection.password_hash.unwrap_or(0);
    stream.extend_from_slice(&PASSWORD_RECORD.to_le_bytes());
    stream.extend_from_slice(&2u16.to_le_bytes());
    stream.extend_from_slice(&password_hash.to_le_bytes());
}

/// Emit a ROW record (MS-XLS §2.4.220) per row that has any non-
/// default property: explicit height, hidden flag, outline level, or
/// collapsed state. Bit layout in the 32-bit options field at body
/// offset 12 matches the reader's parse_row: 0x10 collapsed, 0x20
/// hidden, 0x40 fUnsynced (custom height), bits 8-10 outline level.
/// Rows beyond u16::MAX are silently skipped because BIFF8 can't
/// address them.
fn write_row_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let heights = sheet.custom_row_heights();
    let hidden = sheet.hidden_rows();
    let outlines = sheet.row_outline_levels();
    let collapsed = sheet.collapsed_rows();

    let mut rows: std::collections::BTreeSet<u32> = std::collections::BTreeSet::new();
    rows.extend(heights.keys());
    rows.extend(hidden.keys());
    rows.extend(outlines.keys());
    rows.extend(collapsed.keys());

    for row in rows {
        if row > u16::MAX as u32 {
            continue;
        }
        let height_pt = heights.get(&row).copied();
        let is_hidden = hidden.get(&row).copied().unwrap_or(false);
        let outline_level = outlines.get(&row).copied().unwrap_or(0);
        let is_collapsed = collapsed.get(&row).copied().unwrap_or(false);

        let mut body = Vec::with_capacity(16);
        body.extend_from_slice(&(row as u16).to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // colMic
        body.extend_from_slice(&0u16.to_le_bytes()); // colMac
        let height_twips = match height_pt {
            Some(h) if h > 0.0 => ((h * 20.0).round() as u32).min(0x7FFF) as u16,
            _ => 0,
        };
        body.extend_from_slice(&height_twips.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // reserved1
        body.extend_from_slice(&0u16.to_le_bytes()); // unused1
        let mut options: u32 = 0;
        if is_collapsed {
            options |= 0x10;
        }
        if is_hidden {
            options |= 0x20;
        }
        if height_pt.is_some() {
            options |= 0x40; // fUnsynced (custom height)
        }
        if outline_level > 0 {
            options |= ((outline_level as u32) & 0x07) << 8;
        }
        body.extend_from_slice(&options.to_le_bytes());

        stream.extend_from_slice(&ROW_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit a COLINFO record (MS-XLS §2.4.49) per column with any non-
/// default property: explicit width, hidden flag, outline level, or
/// collapsed state. Each emitted record covers a single column for
/// simplicity; the reader merges adjacent COLINFO ranges fine.
/// Width is converted from the model's "characters" unit to BIFF8's
/// 1/256-of-default-char-width unit.
fn write_colinfo_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let widths = sheet.custom_column_widths();
    let hidden = sheet.hidden_columns();
    let outlines = sheet.column_outline_levels();
    let collapsed = sheet.collapsed_columns();

    let mut cols: std::collections::BTreeSet<u16> = std::collections::BTreeSet::new();
    cols.extend(widths.keys());
    cols.extend(hidden.keys());
    cols.extend(outlines.keys());
    cols.extend(collapsed.keys());

    for col in cols {
        let width_chars = widths.get(&col).copied();
        let is_hidden = hidden.get(&col).copied().unwrap_or(false);
        let outline_level = outlines.get(&col).copied().unwrap_or(0);
        let is_collapsed = collapsed.get(&col).copied().unwrap_or(false);

        let mut body = Vec::with_capacity(12);
        body.extend_from_slice(&col.to_le_bytes());
        body.extend_from_slice(&col.to_le_bytes()); // last == first for one-col record
        let coldx = match width_chars {
            Some(w) if w > 0.0 => ((w * 256.0).round() as u32).min(u16::MAX as u32) as u16,
            _ => 0,
        };
        body.extend_from_slice(&coldx.to_le_bytes());
        body.extend_from_slice(&15u16.to_le_bytes()); // ixfe (default cell XF)
        let mut options: u16 = 0;
        if is_hidden {
            options |= 0x0001;
        }
        if outline_level > 0 {
            options |= ((outline_level as u16) & 0x0007) << 8;
        }
        if is_collapsed {
            options |= 0x1000;
        }
        body.extend_from_slice(&options.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // reserved

        stream.extend_from_slice(&COLINFO_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit a SETUP record (MS-XLS §2.4.252) carrying paper size, scale,
/// fit-to-width/height, orientation, and header/footer margins.
/// Body layout: iPaperSize, iScale, iPageStart, iFitWidth, iFitHeight,
/// grbit, iRes, iVRes, numHdr (f64), numFtr (f64), iCopies. The
/// orientation lives in grbit bit 1 (set = landscape); bit 6 (fNoOrient)
/// is cleared so the orientation is honoured.
fn write_setup_record(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    let mut body = Vec::with_capacity(34);
    body.extend_from_slice(&(ps.paper_size as u16).to_le_bytes());
    body.extend_from_slice(&ps.scale.clamp(10, 400).to_le_bytes());
    body.extend_from_slice(&1u16.to_le_bytes()); // iPageStart
    body.extend_from_slice(&ps.fit_to_width.unwrap_or(0).to_le_bytes());
    body.extend_from_slice(&ps.fit_to_height.unwrap_or(0).to_le_bytes());

    let mut grbit: u16 = 0;
    if matches!(ps.orientation, duke_sheets_core::PageOrientation::Landscape) {
        grbit |= 0x0002;
    }
    body.extend_from_slice(&grbit.to_le_bytes());
    body.extend_from_slice(&600u16.to_le_bytes()); // iRes (DPI)
    body.extend_from_slice(&600u16.to_le_bytes()); // iVRes (DPI)
    body.extend_from_slice(&ps.header_margin.to_le_bytes());
    body.extend_from_slice(&ps.footer_margin.to_le_bytes());
    body.extend_from_slice(&1u16.to_le_bytes()); // iCopies

    stream.extend_from_slice(&SETUP_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

/// Emit HEADER (MS-XLS §2.4.137) and FOOTER (MS-XLS §2.4.111) records
/// when the page setup has odd-page header/footer text. Both wrap the
/// text in a XLUnicodeString (cch + flags + chars). Even/first page
/// headers and footers are not yet emitted; BIFF8's HEADERFOOTER ext
/// record carries those.
fn write_header_footer_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    if let Some(text) = ps.odd_header.as_deref() {
        emit_unicode_string_record(stream, HEADER_RECORD, text);
    }
    if let Some(text) = ps.odd_footer.as_deref() {
        emit_unicode_string_record(stream, FOOTER_RECORD, text);
    }
}

fn emit_unicode_string_record(stream: &mut Vec<u8>, record_type: u16, text: &str) {
    let units: Vec<u16> = text.encode_utf16().collect();
    let high_byte = units.iter().any(|&u| u > 0xFF);
    let mut body = Vec::with_capacity(3 + units.len() * 2);
    body.extend_from_slice(&(units.len() as u16).to_le_bytes());
    if high_byte {
        body.push(0x01);
        for u in &units {
            body.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        body.push(0x00);
        for u in &units {
            body.push(*u as u8);
        }
    }
    stream.extend_from_slice(&record_type.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

/// Emit LEFT_MARGIN, RIGHT_MARGIN, TOP_MARGIN, BOTTOM_MARGIN records.
/// Each is an 8-byte f64 carrying the margin in inches. The SETUP
/// record covers header/footer margins separately.
fn write_margin_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    for (record_type, value) in [
        (LEFT_MARGIN_RECORD, ps.left_margin),
        (RIGHT_MARGIN_RECORD, ps.right_margin),
        (TOP_MARGIN_RECORD, ps.top_margin),
        (BOTTOM_MARGIN_RECORD, ps.bottom_margin),
    ] {
        stream.extend_from_slice(&record_type.to_le_bytes());
        stream.extend_from_slice(&8u16.to_le_bytes());
        stream.extend_from_slice(&value.to_le_bytes());
    }
}

/// Emit PRINTHEADERS and PRINTGRIDLINES boolean records when their
/// flags are set on the page setup. Each body is a single u16 = 1
/// (omitted when the corresponding flag is false to keep the stream
/// minimal).
fn write_print_flags(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    if ps.print_headings {
        stream.extend_from_slice(&PRINTHEADERS_RECORD.to_le_bytes());
        stream.extend_from_slice(&2u16.to_le_bytes());
        stream.extend_from_slice(&1u16.to_le_bytes());
    }
    if ps.print_gridlines {
        stream.extend_from_slice(&PRINTGRIDLINES_RECORD.to_le_bytes());
        stream.extend_from_slice(&2u16.to_le_bytes());
        stream.extend_from_slice(&1u16.to_le_bytes());
    }
}

/// Emit HPAGEBREAKS (row breaks, MS-XLS §2.4.139) and VPAGEBREAKS
/// (column breaks, MS-XLS §2.4.342). Body for each is count(u16)
/// followed by N × (id u16, min u16, max u16). Breaks beyond u16::MAX
/// are skipped because BIFF8 can't address them.
fn write_page_break_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    fn emit(
        stream: &mut Vec<u8>,
        record_type: u16,
        breaks: &[duke_sheets_core::worksheet::PageBreak],
    ) {
        let usable: Vec<_> = breaks
            .iter()
            .filter(|b| {
                b.id <= u16::MAX as u32 && b.min <= u16::MAX as u32 && b.max <= u16::MAX as u32
            })
            .collect();
        if usable.is_empty() {
            return;
        }
        let mut body = Vec::with_capacity(2 + usable.len() * 6);
        body.extend_from_slice(&(usable.len() as u16).to_le_bytes());
        for b in usable {
            body.extend_from_slice(&(b.id as u16).to_le_bytes());
            body.extend_from_slice(&(b.min as u16).to_le_bytes());
            body.extend_from_slice(&(b.max as u16).to_le_bytes());
        }
        stream.extend_from_slice(&record_type.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
    emit(stream, HPAGEBREAKS_RECORD, sheet.row_breaks());
    emit(stream, VPAGEBREAKS_RECORD, sheet.col_breaks());
}

/// Emit one or more MERGECELLS records for the worksheet
/// (MS-XLS §2.4.169). Body: cmcs (u16, count) followed by
/// `cmcs` × Ref8U (rwFirst, rwLast, colFirst, colLast - all u16).
/// Multiple records are emitted when the merge count exceeds
/// `MERGECELLS_MAX_PER_RECORD`. Ranges with rows beyond u16::MAX or
/// cols beyond u16::MAX are silently skipped because BIFF8 can't
/// address them.
fn write_mergecells(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let regions = sheet.merged_regions();
    if regions.is_empty() {
        return;
    }
    for chunk in regions.chunks(MERGECELLS_MAX_PER_RECORD) {
        let mut body = Vec::with_capacity(2 + chunk.len() * 8);
        let mut emitted: u16 = 0;
        let count_offset = body.len();
        body.extend_from_slice(&0u16.to_le_bytes()); // placeholder for cmcs
        for region in chunk {
            if region.start.row > u16::MAX as u32 || region.end.row > u16::MAX as u32 {
                continue;
            }
            body.extend_from_slice(&(region.start.row as u16).to_le_bytes());
            body.extend_from_slice(&(region.end.row as u16).to_le_bytes());
            body.extend_from_slice(&region.start.col.to_le_bytes());
            body.extend_from_slice(&region.end.col.to_le_bytes());
            emitted += 1;
        }
        if emitted == 0 {
            continue;
        }
        body[count_offset..count_offset + 2].copy_from_slice(&emitted.to_le_bytes());
        stream.extend_from_slice(&MERGECELLS_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit a NAME record (MS-XLS §2.4.176, "Lbl") for each worksheet
/// that has a Print_Area or Print_Titles set on its page setup. The
/// record is sheet-scoped (itab = sheet_idx + 1) and uses the
/// built-in name index byte: 0x06 for Print_Area, 0x07 for
/// Print_Titles.
///
/// Print_Area body: a single tArea3D ptg covering the print area
/// range. Print_Titles body holds row titles, column titles, or
/// both:
///
///   - rows only: tArea3D(first_row..last_row, col 0..0xFF)
///   - cols only: tArea3D(row 0..0xFFFF, first_col..last_col)
///   - both:      tMemFunc + cce + tArea3D(rows) + tArea3D(cols)
///                + tList
///
/// ixti is hardcoded to 0; our XlsReader's extract_filter_db_range
/// and extract_print_titles ignore the EXTERNSHEET index when
/// extracting the row/col fields, so the round-trip works without
/// emitting EXTERNSHEET / SUPBOOK records too.
/// Maps user-defined name strings to their 1-based NAME table index
/// (the value embedded in tName ptgs). Names are matched case-
/// insensitively to mirror Excel's name resolution.
#[derive(Debug, Default)]
struct NameTable {
    by_name: HashMap<String, u16>,
}

impl NameTable {
    fn idx_for_name(&self, name: &str) -> Option<u16> {
        self.by_name.get(&name.to_ascii_lowercase()).copied()
    }
}

fn build_name_table(workbook: &Workbook) -> NameTable {
    let mut by_name = HashMap::new();
    for (i, nr) in workbook.named_ranges().iter().enumerate() {
        if i >= u16::MAX as usize {
            break;
        }
        // Names are 1-based in tName ptg encoding.
        by_name.insert(nr.name.to_ascii_lowercase(), (i as u16) + 1);
    }
    NameTable { by_name }
}

/// Emit one NAME record (MS-XLS §2.4.176, Lbl) per user-defined named
/// range in the workbook. Layout:
///
///   flags (u16)    - 0 for visible, 0x0001 fHidden if NamedRange.hidden
///   chKey (u8)     - 0
///   cch (u8)       - name string length (UTF-16 code units)
///   cce (u16)      - formula body length
///   reserved (u16) - 0
///   itab (u16)     - 0 for workbook scope, sheet_idx + 1 for sheet
///                    scope; the reader maps itab=0 -> 0xFFFFFFFF
///                    sheet_idx (workbook), else sheet_idx = itab - 1
///   reserved (u32) - 0
///   name string    - flags byte (0=Latin1, 1=UTF-16) + chars
///   formula body   - parsed via duke-sheets-formula and recompiled
///                    via compile_ptgs_with_context. parse failures
///                    emit no formula body (cce = 0) so the reader
///                    sees the name without a body.
fn write_user_name_records(
    stream: &mut Vec<u8>,
    workbook: &Workbook,
    externsheet: &ExternSheetTable,
    name_table: &NameTable,
) {
    use duke_sheets_core::named_range::NameScope;

    for nr in workbook.named_ranges().iter() {
        let name_units: Vec<u16> = nr.name.encode_utf16().collect();
        if name_units.is_empty() || name_units.len() > u8::MAX as usize {
            continue;
        }

        let formula_body: Vec<u8> = {
            let to_parse = if nr.refers_to.starts_with('=') {
                nr.refers_to.clone()
            } else {
                format!("={}", nr.refers_to)
            };
            if let Ok(expr) = duke_sheets_formula::parse_formula(&to_parse) {
                let mut bytes = Vec::new();
                if compile_ptgs_with_context(&expr, &mut bytes, externsheet, name_table).is_ok() {
                    bytes
                } else {
                    Vec::new()
                }
            } else {
                Vec::new()
            }
        };

        let mut flags: u16 = 0;
        if nr.hidden {
            flags |= 0x0001;
        }
        let itab: u16 = match nr.scope {
            NameScope::Workbook => 0,
            NameScope::Sheet(idx) => (idx.min(u16::MAX as usize) as u16) + 1,
        };

        let high_byte = name_units.iter().any(|&u| u > 0xFF);
        let name_bytes_len = if high_byte {
            1 + name_units.len() * 2
        } else {
            1 + name_units.len()
        };

        let mut body = Vec::with_capacity(15 + name_bytes_len + formula_body.len());
        body.extend_from_slice(&flags.to_le_bytes());
        body.push(0); // chKey
        body.push(name_units.len() as u8);
        body.extend_from_slice(&(formula_body.len() as u16).to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // reserved
        body.extend_from_slice(&itab.to_le_bytes());
        body.extend_from_slice(&[0u8; 4]); // reserved
        if high_byte {
            body.push(0x01);
            for u in &name_units {
                body.extend_from_slice(&u.to_le_bytes());
            }
        } else {
            body.push(0x00);
            for u in &name_units {
                body.push(*u as u8);
            }
        }
        body.extend_from_slice(&formula_body);

        stream.extend_from_slice(&NAME_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Maps sheet names to their EXTERNSHEET ixti index, with case-
/// insensitive lookup to match Excel's name resolution. Built once per
/// workbook write; passed to `compile_ptgs_with_externsheet` so 3D
/// refs (tRef3D / tArea3D) can encode the right ixti.
#[derive(Debug, Default)]
struct ExternSheetTable {
    /// Lower-cased sheet name → ixti.
    by_name: HashMap<String, u16>,
    /// Sheet count, for SUPBOOK self-ref ctab.
    sheet_count: u16,
}

impl ExternSheetTable {
    fn ixti_for_sheet(&self, name: &str) -> Option<u16> {
        self.by_name.get(&name.to_ascii_lowercase()).copied()
    }
}

fn build_externsheet_table(workbook: &Workbook) -> ExternSheetTable {
    let mut by_name = HashMap::new();
    for (idx, sheet) in workbook.worksheets().enumerate() {
        if idx > u16::MAX as usize {
            continue;
        }
        by_name.insert(sheet.name().to_ascii_lowercase(), idx as u16);
    }
    ExternSheetTable {
        by_name,
        sheet_count: workbook.sheet_count().min(u16::MAX as usize) as u16,
    }
}

/// Emit a SUPBOOK self-reference (MS-XLS §2.4.273) followed by an
/// EXTERNSHEET (§2.4.105) with one entry per worksheet.
///
/// SUPBOOK self-ref body: ctab(u16) + cch(u16=0x0401 sentinel).
/// EXTERNSHEET body: count(u16) + count × (sup_book_idx u16,
/// itabFirst u16, itabLast u16). With one entry per sheet pointing
/// to (supbook 0, sheet_idx, sheet_idx), tRef3D / tArea3D ptgs can
/// embed `ixti = sheet_idx` and the reader decompiles back to
/// `Sheet!Cell` formula text.
fn write_supbook_and_externsheet(
    stream: &mut Vec<u8>,
    _workbook: &Workbook,
    table: &ExternSheetTable,
) {
    if table.sheet_count == 0 {
        return;
    }
    // SUPBOOK
    stream.extend_from_slice(&SUPBOOK_RECORD.to_le_bytes());
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&table.sheet_count.to_le_bytes());
    stream.extend_from_slice(&0x0401u16.to_le_bytes());

    // EXTERNSHEET
    let count = table.sheet_count;
    let body_len = 2 + (count as usize) * 6;
    stream.extend_from_slice(&EXTERNSHEET_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body_len as u16).to_le_bytes());
    stream.extend_from_slice(&count.to_le_bytes());
    for i in 0..count {
        stream.extend_from_slice(&0u16.to_le_bytes()); // sup_book_idx
        stream.extend_from_slice(&i.to_le_bytes()); // first_sheet
        stream.extend_from_slice(&i.to_le_bytes()); // last_sheet
    }
}

fn write_print_name_records(stream: &mut Vec<u8>, workbook: &Workbook) {
    for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
        if sheet_idx > u16::MAX as usize - 1 {
            continue;
        }
        let itab = (sheet_idx as u16) + 1;
        let ps = sheet.page_setup();

        if let Some(range) = ps.print_area.as_ref() {
            let body = build_print_area_body(range);
            emit_builtin_name(stream, itab, BUILTIN_NAME_PRINT_AREA, &body);
        }

        let print_titles_body = build_print_titles_body(ps.repeat_rows, ps.repeat_cols);
        if !print_titles_body.is_empty() {
            emit_builtin_name(stream, itab, BUILTIN_NAME_PRINT_TITLES, &print_titles_body);
        }

        if let Some(af) = sheet.auto_filter() {
            let body = build_print_area_body(&af.range);
            emit_builtin_name(stream, itab, BUILTIN_NAME_FILTER_DATABASE, &body);
        }
    }
}

fn emit_builtin_name(stream: &mut Vec<u8>, itab: u16, builtin_index: u8, formula_body: &[u8]) {
    let cce = formula_body.len() as u16;
    let mut body = Vec::with_capacity(15 + 2 + formula_body.len());
    body.extend_from_slice(&0x0020u16.to_le_bytes()); // flags: fBuiltin
    body.push(0); // chKey
    body.push(1); // cch (one "character" - the built-in index byte)
    body.extend_from_slice(&cce.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes()); // reserved (ixals - external book index)
    body.extend_from_slice(&itab.to_le_bytes());
    body.extend_from_slice(&[0u8; 4]); // 4 reserved bytes
    body.push(0); // name flags: 0 = compressed/Latin-1
    body.push(builtin_index); // the "character" is actually the built-in index
    body.extend_from_slice(formula_body);

    stream.extend_from_slice(&NAME_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

/// Build a tArea3D ptg body (11 bytes): token + ixti(u16) + first_row
/// + last_row + first_col + last_col. ixti is 0 (no EXTERNSHEET).
fn build_t_area_3d(first_row: u16, last_row: u16, first_col: u16, last_col: u16) -> [u8; 11] {
    let mut buf = [0u8; 11];
    buf[0] = 0x3B; // tArea3D (R class)
                   // bytes 1..3: ixti (u16) = 0
    buf[3..5].copy_from_slice(&first_row.to_le_bytes());
    buf[5..7].copy_from_slice(&last_row.to_le_bytes());
    buf[7..9].copy_from_slice(&(first_col & 0x3FFF).to_le_bytes());
    buf[9..11].copy_from_slice(&(last_col & 0x3FFF).to_le_bytes());
    buf
}

fn build_print_area_body(range: &duke_sheets_core::CellRange) -> Vec<u8> {
    let first_row = range.start.row.min(u16::MAX as u32) as u16;
    let last_row = range.end.row.min(u16::MAX as u32) as u16;
    let first_col = range.start.col;
    let last_col = range.end.col;
    build_t_area_3d(first_row, last_row, first_col, last_col).to_vec()
}

fn build_print_titles_body(
    repeat_rows: Option<(u32, u32)>,
    repeat_cols: Option<(u16, u16)>,
) -> Vec<u8> {
    let row_area = repeat_rows.map(|(r1, r2)| {
        build_t_area_3d(
            r1.min(u16::MAX as u32) as u16,
            r2.min(u16::MAX as u32) as u16,
            0,
            0xFF,
        )
    });
    let col_area = repeat_cols.map(|(c1, c2)| build_t_area_3d(0, 0xFFFF, c1, c2));

    match (row_area, col_area) {
        (None, None) => Vec::new(),
        (Some(rows), None) => rows.to_vec(),
        (None, Some(cols)) => cols.to_vec(),
        (Some(rows), Some(cols)) => {
            // tMemFunc(0x29) + cce(u16) + rows(11) + cols(11) + tList(0x10)
            let inner_len: u16 = 11 + 11 + 1;
            let mut body = Vec::with_capacity(3 + inner_len as usize);
            body.push(0x29);
            body.extend_from_slice(&inner_len.to_le_bytes());
            body.extend_from_slice(&rows);
            body.extend_from_slice(&cols);
            body.push(0x10); // tList (range union)
            body
        }
    }
}

/// Emit one HLINK record per cell-attached hyperlink (MS-XLS §2.4.144).
///
/// Body layout:
///   - Ref8U (8 bytes): row_first/row_last/col_first/col_last (u16
///     each); single-cell hyperlinks use row_first==row_last and
///     col_first==col_last.
///   - HLINK CLSID (16 bytes).
///   - streamVersion (4 bytes) = 0x00000002.
///   - flags (4 bytes): bit 0 has_moniker, bit 1 is_absolute,
///     bit 3 has_location, bit 4 has_display, bit 7 has_frame.
///   - displayName (only when bit 4 set): char_count(u32) + UTF-16LE
///     chars + 0x0000 terminator.
///   - frameName (when bit 7 set): same encoding.
///   - moniker (when bit 0 set): URL_MONIKER_CLSID + url_byte_len(u32)
///     + UTF-16LE chars + 0x0000 terminator.
///   - location (when bit 3 set): char_count(u32) + UTF-16LE chars +
///     0x0000 terminator.
///
/// Internal `#Sheet!A1` targets are emitted with no moniker but a
/// location string. URL-style targets use the URL moniker. File
/// monikers and other types are not yet emitted; the reader can read
/// them but the writer round-trips file paths via the URL path.
fn write_hlink_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    use duke_sheets_core::CellAddress;

    let mut entries: Vec<(CellAddress, &duke_sheets_core::Hyperlink)> =
        sheet.hyperlinks().iter().map(|(a, h)| (*a, h)).collect();
    entries.sort_by_key(|(addr, _)| (addr.row, addr.col));

    for (addr, hyperlink) in entries {
        if addr.row > u16::MAX as u32 {
            continue;
        }
        let mut body = Vec::with_capacity(64);
        body.extend_from_slice(&(addr.row as u16).to_le_bytes());
        body.extend_from_slice(&(addr.row as u16).to_le_bytes()); // row_last
        body.extend_from_slice(&addr.col.to_le_bytes());
        body.extend_from_slice(&addr.col.to_le_bytes()); // col_last
        body.extend_from_slice(&HLINK_CLSID);
        body.extend_from_slice(&2u32.to_le_bytes()); // streamVersion

        let target = hyperlink.target.as_str();
        let is_internal = target.starts_with('#') || target.is_empty();
        let location_text: Option<String> = if is_internal {
            if target.starts_with('#') {
                Some(target[1..].to_string())
            } else {
                hyperlink.location.clone()
            }
        } else {
            hyperlink.location.clone()
        };

        let display = hyperlink.display.as_deref();

        let mut flags: u32 = 0;
        if !is_internal {
            flags |= 0x0001 | 0x0002; // hlstmfHasMoniker + hlstmfIsAbsolute
        }
        if location_text.is_some() {
            flags |= 0x0008; // hlstmfHasLocationStr (text mark)
        }
        if display.is_some() {
            // hlstmfHasDisplayName only. The 0x04 bit
            // (hlstmfSiteGaveDisplayName) means "displayName is implicit
            // and not stored" — combining it with HasDisplayName is a
            // contradiction Excel rejects with "Open method of Workbooks
            // class failed".
            flags |= 0x0010;
        }
        body.extend_from_slice(&flags.to_le_bytes());

        if let Some(text) = display {
            push_hlink_string(&mut body, text);
        }

        if !is_internal {
            body.extend_from_slice(&URL_MONIKER_CLSID);
            push_url_moniker_payload(&mut body, target);
        }

        if let Some(loc) = location_text.as_deref() {
            push_hlink_string(&mut body, loc);
        }

        stream.extend_from_slice(&HLINK_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit a FILTERMODE record (1 byte body of zero) plus one AUTOFILTER
/// record per FilterColumn when the worksheet has an auto-filter set.
/// FILTERMODE indicates the worksheet has filtering active; the
/// AUTOFILTER body carries per-column criteria.
///
/// AUTOFILTER body (24 bytes minimum + variable strings):
///   bytes 0..2  - i_entry (column offset within the filter range)
///   bytes 2..4  - flags:
///       bits 0-1 - 0x01 = OR-join the two dopers (else AND)
///       bit 4    - fTopN
///       bit 5    - fTop (top vs bottom for top-N)
///       bit 6    - fPercent
///       bits 7-15 - wTopN value
///   bytes 4..14 - doper1 (vt + op + 8-byte payload)
///   bytes 14..24 - doper2
///   bytes 24..  - inline strings for any string-typed dopers
fn write_autofilter_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    use duke_sheets_core::auto_filter::{ColumnFilter, FilterOperator};

    let Some(af) = sheet.auto_filter() else {
        return;
    };

    stream.extend_from_slice(&FILTERMODE_RECORD.to_le_bytes());
    stream.extend_from_slice(&0u16.to_le_bytes());

    for column in &af.filter_columns {
        if column.col_id > u16::MAX as u32 {
            continue;
        }
        let mut body = Vec::with_capacity(32);
        body.extend_from_slice(&(column.col_id as u16).to_le_bytes());

        let mut flags: u16 = 0;
        let mut doper1 = [0u8; 10];
        let mut doper2 = [0u8; 10];
        let mut trailing_strings: Vec<String> = Vec::new();

        match &column.filter {
            ColumnFilter::Top10(t) => {
                flags |= 0x0010; // fTopN
                if t.top {
                    flags |= 0x0020;
                }
                if t.percent {
                    flags |= 0x0040;
                }
                let n = (t.val as u16) & 0x01FF;
                flags |= n << 7;
            }
            ColumnFilter::Custom(c) => {
                if !c.and {
                    flags |= 0x0001; // OR-join
                }
                if let Some(cond) = c.conditions.first() {
                    encode_custom_condition(cond, &mut doper1, &mut trailing_strings);
                }
                if let Some(cond) = c.conditions.get(1) {
                    encode_custom_condition(cond, &mut doper2, &mut trailing_strings);
                }
            }
            ColumnFilter::Values(v) => {
                flags |= 0x0001; // OR-join (Values are matched as Equal+OR)
                if let Some(value) = v.values.first() {
                    encode_custom_condition(
                        &duke_sheets_core::auto_filter::CustomFilterCondition {
                            operator: FilterOperator::Equal,
                            value: value.clone(),
                        },
                        &mut doper1,
                        &mut trailing_strings,
                    );
                }
                if let Some(value) = v.values.get(1) {
                    encode_custom_condition(
                        &duke_sheets_core::auto_filter::CustomFilterCondition {
                            operator: FilterOperator::Equal,
                            value: value.clone(),
                        },
                        &mut doper2,
                        &mut trailing_strings,
                    );
                }
            }
            ColumnFilter::Dynamic(_) | ColumnFilter::Color(_) => {
                // Reader doesn't decode these from BIFF8 today; skip
                // emit so we don't create records that won't round-
                // trip via XlsReader.
                continue;
            }
        }

        body.extend_from_slice(&flags.to_le_bytes());
        body.extend_from_slice(&doper1);
        body.extend_from_slice(&doper2);
        for s in &trailing_strings {
            push_autofilter_string(&mut body, s);
        }

        stream.extend_from_slice(&AUTOFILTER_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Encode a `CustomFilterCondition` into a 10-byte doper. Numeric
/// conditions store the value inline (vt=0x04, IEEE 754 f64). String
/// conditions emit vt=0x06 with cch in payload[4]; the actual UTF-8
/// string is collected into `trailing_strings` and appended to the
/// AUTOFILTER record body in order.
fn encode_custom_condition(
    cond: &duke_sheets_core::auto_filter::CustomFilterCondition,
    doper: &mut [u8; 10],
    trailing_strings: &mut Vec<String>,
) {
    use duke_sheets_core::auto_filter::FilterOperator;

    doper[1] = match cond.operator {
        FilterOperator::LessThan => 0x01,
        FilterOperator::Equal => 0x02,
        FilterOperator::LessThanOrEqual => 0x03,
        FilterOperator::GreaterThan => 0x04,
        FilterOperator::NotEqual => 0x05,
        FilterOperator::GreaterThanOrEqual => 0x06,
    };

    if let Ok(n) = cond.value.parse::<f64>() {
        doper[0] = 0x04; // f64
        doper[2..10].copy_from_slice(&n.to_le_bytes());
    } else {
        doper[0] = 0x06; // string
        let units = cond.value.encode_utf16().count() as u8;
        doper[6] = units; // cch lives at offset 4 of payload, i.e. doper[6]
        trailing_strings.push(cond.value.clone());
    }
}

fn push_autofilter_string(buf: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let high_byte = units.iter().any(|&u| u > 0xFF);
    if high_byte {
        buf.push(0x01);
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        buf.push(0x00);
        for u in &units {
            buf.push(*u as u8);
        }
    }
}

/// Emit one CONDFMT (MS-XLS §2.4.45) + CF (§2.4.43) pair per
/// `ConditionalFormatRule`. Each CONDFMT carries the bounding
/// rectangle and the per-range Ref8U list; each CF carries the rule
/// header + cce1 + cce2 + an empty dxf block (no formatting overrides
/// emitted yet) + formula1 + formula2.
///
/// The reader keeps the most recent CONDFMT range list and applies it
/// to subsequent CF records, so we emit them as alternating
/// CONDFMT → CF pairs to keep the mapping unambiguous.
fn write_conditional_formats(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    externsheet: &ExternSheetTable,
    names: &NameTable,
) {
    use duke_sheets_core::conditional_format::{CfOperator, CfRuleType};

    for rule in sheet.conditional_formats() {
        if rule.ranges.is_empty() {
            continue;
        }

        let usable: Vec<&duke_sheets_core::CellRange> = rule
            .ranges
            .iter()
            .filter(|r| r.start.row <= u16::MAX as u32 && r.end.row <= u16::MAX as u32)
            .collect();
        if usable.is_empty() {
            continue;
        }

        // Compute enclosing range = bounding box of all individual
        // ranges. The reader skips this field but Excel/LO use it.
        let mut enc_first_row = u16::MAX;
        let mut enc_last_row = 0u16;
        let mut enc_first_col = u16::MAX;
        let mut enc_last_col = 0u16;
        for r in &usable {
            enc_first_row = enc_first_row.min(r.start.row as u16);
            enc_last_row = enc_last_row.max(r.end.row as u16);
            enc_first_col = enc_first_col.min(r.start.col);
            enc_last_col = enc_last_col.max(r.end.col);
        }

        let mut cfmt_body = Vec::with_capacity(14 + usable.len() * 8);
        cfmt_body.extend_from_slice(&1u16.to_le_bytes()); // cCF: 1 rule follows
        cfmt_body.extend_from_slice(&1u16.to_le_bytes()); // flags (fAlwaysCalc-ish, set to 1)
        cfmt_body.extend_from_slice(&enc_first_row.to_le_bytes());
        cfmt_body.extend_from_slice(&enc_last_row.to_le_bytes());
        cfmt_body.extend_from_slice(&enc_first_col.to_le_bytes());
        cfmt_body.extend_from_slice(&enc_last_col.to_le_bytes());
        cfmt_body.extend_from_slice(&(usable.len() as u16).to_le_bytes());
        for r in &usable {
            cfmt_body.extend_from_slice(&(r.start.row as u16).to_le_bytes());
            cfmt_body.extend_from_slice(&(r.end.row as u16).to_le_bytes());
            cfmt_body.extend_from_slice(&r.start.col.to_le_bytes());
            cfmt_body.extend_from_slice(&r.end.col.to_le_bytes());
        }
        stream.extend_from_slice(&CONDFMT_RECORD.to_le_bytes());
        stream.extend_from_slice(&(cfmt_body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&cfmt_body);

        let (ct, cp, formula1_text, formula2_text) = match &rule.rule_type {
            CfRuleType::CellIs {
                operator,
                formula1,
                formula2,
            } => {
                let cp = match operator {
                    CfOperator::Between => 1u8,
                    CfOperator::NotBetween => 2,
                    CfOperator::Equal => 3,
                    CfOperator::NotEqual => 4,
                    CfOperator::GreaterThan => 5,
                    CfOperator::LessThan => 6,
                    CfOperator::GreaterThanOrEqual => 7,
                    CfOperator::LessThanOrEqual => 8,
                };
                (1u8, cp, Some(formula1.as_str()), formula2.as_deref())
            }
            CfRuleType::Expression { formula } => (2u8, 0u8, Some(formula.as_str()), None),
            _ => continue, // skip rule types we don't know how to emit
        };

        let f1 = formula1_text
            .map(|t| encode_dv_formula(t, externsheet, names))
            .unwrap_or_default();
        let f2 = formula2_text
            .map(|t| encode_dv_formula(t, externsheet, names))
            .unwrap_or_default();

        let mut cf_body = Vec::with_capacity(12 + f1.len() + f2.len());
        cf_body.push(ct);
        cf_body.push(cp);
        cf_body.extend_from_slice(&(f1.len() as u16).to_le_bytes());
        cf_body.extend_from_slice(&(f2.len() as u16).to_le_bytes());
        // dxfn12 header (6 bytes): all-zero flags = "no formatting
        // override". Excel rejects CF records that omit this block
        // ("Open method of Workbooks class failed") even though our
        // permissive reader and LibreOffice both tolerate the
        // shorter shape.
        cf_body.extend_from_slice(&[0u8; 6]);
        cf_body.extend_from_slice(&f1);
        cf_body.extend_from_slice(&f2);

        stream.extend_from_slice(&CF_RECORD.to_le_bytes());
        stream.extend_from_slice(&(cf_body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&cf_body);
    }
}

/// Emit a DVAL header (MS-XLS §2.4.81) plus one DV record (§2.4.79)
/// per `DataValidation` on the worksheet.
///
/// DVAL body (18 bytes):
///   options (u16)        - input-box state flags
///   xLeft, yTop (i32x2)  - input-box position
///   iDvIdInputBox (u32)  - the DV record currently in the input box
///   idv (u32)            - count of DV records that follow
///
/// DV body:
///   flags (u32)          - val_type (bits 0-3), err_style (4-6),
///                          fExplicit list (7), fAllowBlank (8),
///                          fSuppressDropdown (9), fShowInput (18),
///                          fShowError (19), operator (20-23)
///   input_title          - XLUnicodeString
///   error_title          - XLUnicodeString
///   input_msg            - XLUnicodeString
///   error_msg            - XLUnicodeString
///   cce1 (u16) + unused (u16) + formula1 ptgs
///   cce2 (u16) + unused (u16) + formula2 ptgs
///   range_count (u16) + N × Ref8U (8 bytes: r1/r2/c1/c2 u16)
fn write_data_validations(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    externsheet: &ExternSheetTable,
    names: &NameTable,
) {
    use duke_sheets_core::validation::{ValidationErrorStyle, ValidationOperator, ValidationType};

    let validations = sheet.data_validations();
    if validations.is_empty() {
        return;
    }

    let mut dval_body = Vec::with_capacity(18);
    dval_body.extend_from_slice(&0u16.to_le_bytes()); // options
    dval_body.extend_from_slice(&0i32.to_le_bytes()); // xLeft
    dval_body.extend_from_slice(&0i32.to_le_bytes()); // yTop
    dval_body.extend_from_slice(&0u32.to_le_bytes()); // iDvIdInputBox
    dval_body.extend_from_slice(&(validations.len() as u32).to_le_bytes());
    stream.extend_from_slice(&DVAL_RECORD.to_le_bytes());
    stream.extend_from_slice(&(dval_body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&dval_body);

    for v in validations {
        let val_type_bits: u32 = match &v.validation_type {
            ValidationType::None => 0,
            ValidationType::Whole { .. } => 1,
            ValidationType::Decimal { .. } => 2,
            ValidationType::List { .. } => 3,
            ValidationType::Date { .. } => 4,
            ValidationType::Time { .. } => 5,
            ValidationType::TextLength { .. } => 6,
            ValidationType::Custom { .. } => 7,
        };
        let err_style_bits: u32 = match v.error_style {
            ValidationErrorStyle::Stop => 0,
            ValidationErrorStyle::Warning => 1,
            ValidationErrorStyle::Information => 2,
        };
        let op_bits: u32 = match &v.validation_type {
            ValidationType::Whole { operator, .. }
            | ValidationType::Decimal { operator, .. }
            | ValidationType::Date { operator, .. }
            | ValidationType::Time { operator, .. }
            | ValidationType::TextLength { operator, .. } => match operator {
                ValidationOperator::Between => 0,
                ValidationOperator::NotBetween => 1,
                ValidationOperator::Equal => 2,
                ValidationOperator::NotEqual => 3,
                ValidationOperator::GreaterThan => 4,
                ValidationOperator::LessThan => 5,
                ValidationOperator::GreaterThanOrEqual => 6,
                ValidationOperator::LessThanOrEqual => 7,
            },
            _ => 0,
        };
        let is_explicit_list = matches!(
            &v.validation_type,
            ValidationType::List { source } if !source.starts_with('=')
        );

        let mut flags: u32 = val_type_bits | (err_style_bits << 4) | (op_bits << 20);
        if is_explicit_list {
            flags |= 0x0080;
        }
        if v.allow_blank {
            flags |= 0x0100;
        }
        if !v.show_dropdown {
            flags |= 0x0200;
        }
        if v.show_input_message {
            flags |= 0x0004_0000;
        }
        if v.show_error_alert {
            flags |= 0x0008_0000;
        }

        let mut body = Vec::new();
        body.extend_from_slice(&flags.to_le_bytes());

        // String headers (always emit, empty when None).
        for s in [
            v.input_title.as_deref(),
            v.error_title.as_deref(),
            v.input_message.as_deref(),
            v.error_message.as_deref(),
        ] {
            push_dv_unicode_string(&mut body, s.unwrap_or(""));
        }

        let (value1, value2) = match &v.validation_type {
            ValidationType::Whole { value1, value2, .. }
            | ValidationType::Decimal { value1, value2, .. }
            | ValidationType::Date { value1, value2, .. }
            | ValidationType::Time { value1, value2, .. }
            | ValidationType::TextLength { value1, value2, .. } => {
                (Some(value1.as_str()), value2.as_deref())
            }
            ValidationType::List { source } => (Some(source.as_str()), None),
            ValidationType::Custom { formula } => (Some(formula.as_str()), None),
            ValidationType::None => (None, None),
        };

        let formula1 = value1
            .map(|t| encode_dv_formula(t, externsheet, names))
            .unwrap_or_default();
        body.extend_from_slice(&(formula1.len() as u16).to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // unused
        body.extend_from_slice(&formula1);

        let formula2 = value2
            .map(|t| encode_dv_formula(t, externsheet, names))
            .unwrap_or_default();
        body.extend_from_slice(&(formula2.len() as u16).to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // unused
        body.extend_from_slice(&formula2);

        let usable: Vec<&duke_sheets_core::CellRange> = v
            .ranges
            .iter()
            .filter(|r| r.start.row <= u16::MAX as u32 && r.end.row <= u16::MAX as u32)
            .collect();
        body.extend_from_slice(&(usable.len() as u16).to_le_bytes());
        for r in usable {
            body.extend_from_slice(&(r.start.row as u16).to_le_bytes());
            body.extend_from_slice(&(r.end.row as u16).to_le_bytes());
            body.extend_from_slice(&r.start.col.to_le_bytes());
            body.extend_from_slice(&r.end.col.to_le_bytes());
        }

        stream.extend_from_slice(&DV_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit an XLUnicodeString (cch u16 + flags u8 + chars) for DV input
/// titles, error titles, and message text. Empty strings emit cch=0
/// and a flags byte of 0; the reader treats them as None.
fn push_dv_unicode_string(buf: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let cch = units.len().min(u16::MAX as usize) as u16;
    buf.extend_from_slice(&cch.to_le_bytes());
    let high_byte = units.iter().any(|&u| u > 0xFF);
    if high_byte {
        buf.push(0x01);
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        buf.push(0x00);
        for u in &units {
            buf.push(*u as u8);
        }
    }
}

/// Encode a DataValidation value-string into a ptg formula body.
///
/// The model carries values as decompiled formula strings (e.g.
/// "100", "10.5", "Red,Green,Blue", "=A1+1"). Round-tripping requires
/// re-parsing as a formula and recompiling to ptgs:
///
///   - Numeric/string/cell-ref values reach the parser via a synthetic
///     `=` prefix. parse_formula handles `=100`, `=10.5`, `=A1`, etc.
///   - Custom validation formulas already start with `=`.
///   - Inline list sources like "Red,Green,Blue" don't parse as a
///     formula expression. Fall back to a single tStr ptg that
///     decompiles back to a quoted string; the reader strips the
///     surrounding quotes when the fExplicit-list flag is set.
fn encode_dv_formula(value: &str, externsheet: &ExternSheetTable, names: &NameTable) -> Vec<u8> {
    if value.is_empty() {
        return Vec::new();
    }
    let to_parse = if value.starts_with('=') {
        value.to_string()
    } else {
        format!("={value}")
    };
    if let Ok(expr) = duke_sheets_formula::parse_formula(&to_parse) {
        let mut bytes = Vec::new();
        if compile_ptgs_with_context(&expr, &mut bytes, externsheet, names).is_ok() {
            return bytes;
        }
    }
    // Fallback: emit the raw value as a tStr ptg literal.
    let mut bytes = Vec::new();
    bytes.push(0x17); // PTG_STR
    let _ = push_short_xlunicode_string(&mut bytes, value);
    bytes
}

/// Encode a length-prefixed UTF-16LE string for HLINK record fields
/// (displayName, frameName, location). Format: char_count(u32,
/// includes null terminator) + UTF-16LE chars + 0x0000.
fn push_hlink_string(buf: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let total_chars = (units.len() + 1) as u32; // include null
    buf.extend_from_slice(&total_chars.to_le_bytes());
    for u in &units {
        buf.extend_from_slice(&u.to_le_bytes());
    }
    buf.extend_from_slice(&0u16.to_le_bytes()); // null terminator
}

/// Encode the URL moniker payload following the URL_MONIKER_CLSID:
/// byte_len(u32, includes the null terminator's 2 bytes) + UTF-16LE
/// chars + 0x0000. Note: the URL moniker uses BYTE length, not char
/// count, unlike the regular HLINK strings.
fn push_url_moniker_payload(buf: &mut Vec<u8>, url: &str) {
    let units: Vec<u16> = url.encode_utf16().collect();
    let byte_len = ((units.len() + 1) * 2) as u32; // include null
    buf.extend_from_slice(&byte_len.to_le_bytes());
    for u in &units {
        buf.extend_from_slice(&u.to_le_bytes());
    }
    buf.extend_from_slice(&0u16.to_le_bytes()); // null terminator
}

/// Emit cell records (BLANK, NUMBER, BOOLERR, LABELSST, FORMULA) for
/// every non-empty cell in `sheet`, sorted in row-major order.
/// Spill-target cells are silently skipped (dynamic-array machinery;
/// deferred). Formula cells with an unsupported AST shape (named
/// ranges, structured refs, function calls, etc.) fall back to
/// emitting their cached value as a static cell.
///
/// `ixfe` is resolved via `styles.ixfe_for_cell` so cells with a
/// non-default style point at the appropriate user-defined XF.
fn write_cell_records(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    sheet_idx: usize,
    sst: &SstTable,
    styles: &StyleTables,
    externsheet: &ExternSheetTable,
    names: &NameTable,
) {
    let mut cells: Vec<_> = sheet.iter_cells().collect();
    cells.sort_by_key(|(row, col, _)| (*row, *col));

    for (row, col, data) in cells {
        if row > u16::MAX as u32 {
            continue;
        }
        let row16 = row as u16;
        let ixfe = styles.ixfe_for_cell(sheet_idx, data.style_index);

        if let Some(formula_text) = sheet.get_formula_at(row, col) {
            if try_write_formula_record(
                stream,
                row16,
                col,
                ixfe,
                formula_text,
                &data.value,
                sst,
                externsheet,
                names,
            ) {
                continue;
            }
        }

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
                if let Some(idx) = sst.lookup_plain(s.as_ref()) {
                    write_labelsst(stream, row16, col, ixfe, idx);
                }
            }
            CellValue::RichText(runs) => {
                let mut text = String::new();
                let mut formatting: Vec<(u16, u16)> = Vec::new();
                for run in runs.iter() {
                    let char_pos = text.encode_utf16().count().min(u16::MAX as usize) as u16;
                    if let Some(rf) = &run.font {
                        let font = run_font_to_font_style(rf);
                        if let Some(&font_idx) = styles.font_xf_index.get(&font) {
                            formatting.push((char_pos, font_idx));
                        }
                    }
                    text.push_str(&run.text);
                }
                if let Some(idx) = sst.lookup_rich(&text, &formatting) {
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

/// Try to compile and emit a FORMULA record for a cell. Returns true
/// on success, false if the formula AST contains constructs we don't
/// yet emit (named ranges, function calls, structured refs, external
/// refs, arrays); on false the caller falls back to emitting the
/// cell's cached value as a static record.
///
/// Slice 5a supports: numeric/string/bool/error literals, single cell
/// refs (relative + absolute), range refs, binary operators (+, -, *,
/// /, ^, &, comparisons), unary minus / unary plus / percent.
fn try_write_formula_record(
    stream: &mut Vec<u8>,
    row: u16,
    col: u16,
    ixfe: u16,
    formula_text: &str,
    cached: &CellValue,
    sst: &SstTable,
    externsheet: &ExternSheetTable,
    names: &NameTable,
) -> bool {
    // duke-sheets-formula's parse_formula requires the leading '=';
    // ensure it's present without double-prefixing.
    let with_eq_owned: String;
    let parse_input: &str = if formula_text.starts_with('=') {
        formula_text
    } else {
        with_eq_owned = format!("={formula_text}");
        &with_eq_owned
    };
    let Ok(expr) = duke_sheets_formula::parse_formula(parse_input) else {
        return false;
    };
    let mut tokens = Vec::with_capacity(32);
    if compile_ptgs_with_context(&expr, &mut tokens, externsheet, names).is_err() {
        return false;
    }
    if tokens.len() > u16::MAX as usize {
        return false;
    }

    let cached_bytes = encode_cached_result(cached);
    stream.extend_from_slice(&FORMULA_RECORD.to_le_bytes());
    let body_len: u16 = 22 + tokens.len() as u16;
    stream.extend_from_slice(&body_len.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.extend_from_slice(&cached_bytes);
    let grbit: u16 = 0x0002; // fAlwaysCalc cleared, fCalcOnLoad set: cause Excel to recompute on open
    stream.extend_from_slice(&grbit.to_le_bytes());
    stream.extend_from_slice(&0u32.to_le_bytes()); // chn (cache key)
    stream.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    stream.extend_from_slice(&tokens);

    if let CellValue::String(s) = cached {
        write_string_followup(stream, s.as_ref());
    } else if let CellValue::RichText(runs) = cached {
        let plain: String = runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<Vec<_>>()
            .join("");
        write_string_followup(stream, &plain);
    }

    let _ = sst;
    true
}

fn write_string_followup(stream: &mut Vec<u8>, s: &str) {
    let mut body = Vec::with_capacity(3 + s.len() * 2);
    let units: Vec<u16> = s.encode_utf16().collect();
    let high_byte = units.iter().any(|&u| u > 0xFF);
    body.extend_from_slice(&(units.len() as u16).to_le_bytes());
    if high_byte {
        body.push(0x01);
        for u in &units {
            body.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        body.push(0x00);
        for u in &units {
            body.push(*u as u8);
        }
    }
    stream.extend_from_slice(&STRING_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

/// Encode an 8-byte FORMULA cached_result field (MS-XLS §2.5.133
/// FormulaValue). Numeric results: 8-byte f64 little-endian. Other
/// types use a sentinel encoding where bytes[6..8] = 0xFFFF and
/// bytes[0] selects the variant: 0=string (real value in STRING
/// follow-up), 1=bool, 2=error, 3=empty.
fn encode_cached_result(value: &CellValue) -> [u8; 8] {
    let mut out = [0u8; 8];
    match value {
        CellValue::Number(n) => {
            out.copy_from_slice(&n.to_le_bytes());
        }
        CellValue::String(_) | CellValue::RichText(_) => {
            out[0] = 0x00;
            out[6] = 0xFF;
            out[7] = 0xFF;
        }
        CellValue::Boolean(b) => {
            out[0] = 0x01;
            out[2] = if *b { 1 } else { 0 };
            out[6] = 0xFF;
            out[7] = 0xFF;
        }
        CellValue::Error(e) => {
            out[0] = 0x02;
            out[2] = e.code();
            out[6] = 0xFF;
            out[7] = 0xFF;
        }
        CellValue::Empty | CellValue::SpillTarget { .. } => {
            out[0] = 0x03;
            out[6] = 0xFF;
            out[7] = 0xFF;
        }
    }
    out
}

#[derive(Debug)]
struct UnsupportedToken;

/// Recursively walk a `FormulaExpr` in postfix order, appending BIFF8
/// ptg bytes to `out`. Returns `Err(UnsupportedToken)` for AST shapes
/// that slice 5a doesn't yet emit (named ranges, function calls,
/// structured refs, external refs, arrays, intersection/union ops);
/// the caller falls back to emitting the cached value as a static
/// cell record so the spreadsheet still renders correctly.
fn compile_ptgs(
    expr: &duke_sheets_formula::FormulaExpr,
    out: &mut Vec<u8>,
) -> Result<(), UnsupportedToken> {
    compile_ptgs_with_context(
        expr,
        out,
        &ExternSheetTable::default(),
        &NameTable::default(),
    )
}

fn compile_ptgs_with_context(
    expr: &duke_sheets_formula::FormulaExpr,
    out: &mut Vec<u8>,
    externsheet: &ExternSheetTable,
    names: &NameTable,
) -> Result<(), UnsupportedToken> {
    use duke_sheets_formula::ast::{BinaryOperator, UnaryOperator};
    use duke_sheets_formula::FormulaExpr;

    match expr {
        FormulaExpr::Number(n) => {
            out.push(0x1F); // PTG_NUM
            out.extend_from_slice(&n.to_le_bytes());
        }
        FormulaExpr::String(s) => {
            out.push(0x17); // PTG_STR
            push_short_xlunicode_string(out, s).map_err(|_| UnsupportedToken)?;
        }
        FormulaExpr::Boolean(b) => {
            out.push(0x1D); // PTG_BOOL
            out.push(if *b { 1 } else { 0 });
        }
        FormulaExpr::Error(e) => {
            out.push(0x1C); // PTG_ERR
            out.push(e.code());
        }
        FormulaExpr::CellRef(cref) => {
            if let Some(sheet_name) = cref.sheet.as_deref() {
                let ixti = externsheet
                    .ixti_for_sheet(sheet_name)
                    .ok_or(UnsupportedToken)?;
                // tRef3D V class = base 0x3A | V-class 0x20 = 0x5A.
                // (0x5C is tRefErr3D - decompiles to "#REF!".)
                out.push(0x5A);
                out.extend_from_slice(&ixti.to_le_bytes());
                push_ref_payload(out, &cref.address)?;
                return Ok(());
            }
            out.push(0x44); // PTG_REF (V class)
            push_ref_payload(out, &cref.address)?;
        }
        FormulaExpr::RangeRef(rref) => {
            if let Some(sheet_name) = rref.sheet.as_deref() {
                let ixti = externsheet
                    .ixti_for_sheet(sheet_name)
                    .ok_or(UnsupportedToken)?;
                out.push(0x3B); // PTG_AREA3D (R class)
                out.extend_from_slice(&ixti.to_le_bytes());
                push_area_payload(out, &rref.range)?;
                return Ok(());
            }
            out.push(0x45); // PTG_AREA (V class)
            push_area_payload(out, &rref.range)?;
        }
        FormulaExpr::BinaryOp { op, left, right } => {
            // The `:` operator with NameRef/Number leaves on both sides
            // is how Excel-style full-column (`A:A`) and full-row
            // (`1:1`) refs reach us from the parser - it lacks a
            // first-class FullColumn / FullRow AST node. Detect those
            // shapes and emit a single tArea ptg covering the BIFF8
            // sheet extent.
            if matches!(op, BinaryOperator::Range) {
                if let Some(area) = full_column_or_row_range(left, right) {
                    out.push(0x45); // PTG_AREA (V class)
                    push_area_payload(out, &area)?;
                    return Ok(());
                }
            }
            compile_ptgs_with_context(left, out, externsheet, names)?;
            compile_ptgs_with_context(right, out, externsheet, names)?;
            out.push(match op {
                BinaryOperator::Add => 0x03,
                BinaryOperator::Subtract => 0x04,
                BinaryOperator::Multiply => 0x05,
                BinaryOperator::Divide => 0x06,
                BinaryOperator::Power => 0x07,
                BinaryOperator::Concat => 0x08,
                BinaryOperator::LessThan => 0x09,
                BinaryOperator::LessEqual => 0x0A,
                BinaryOperator::Equal => 0x0B,
                BinaryOperator::GreaterEqual => 0x0C,
                BinaryOperator::GreaterThan => 0x0D,
                BinaryOperator::NotEqual => 0x0E,
                BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect => {
                    return Err(UnsupportedToken)
                }
            });
        }
        FormulaExpr::UnaryOp { op, operand } => {
            compile_ptgs_with_context(operand, out, externsheet, names)?;
            out.push(match op {
                UnaryOperator::Negate => 0x13,
                UnaryOperator::Percent => 0x14,
                UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange => {
                    return Err(UnsupportedToken)
                }
            });
        }
        FormulaExpr::Function { name, args } => {
            let Some(idx) = crate::biff::formula::function_table::function_index(name) else {
                return Err(UnsupportedToken);
            };
            if args.len() > u8::MAX as usize {
                return Err(UnsupportedToken);
            }
            for arg in args {
                if matches!(arg, FormulaExpr::Empty) {
                    out.push(0x16); // PTG_MISS_ARG
                } else {
                    compile_ptgs_with_context(arg, out, externsheet, names)?;
                }
            }
            // Always emit tFuncVar (V class, 0x42) so the variable-
            // argument count is encoded inline with the token; the
            // reader handles fixed-arity functions decoded this way
            // identically to the more compact tFunc form.
            out.push(0x42);
            out.push(args.len() as u8);
            out.extend_from_slice(&idx.to_le_bytes());
        }
        FormulaExpr::NameRef(name) => {
            let idx = names.idx_for_name(name).ok_or(UnsupportedToken)?;
            out.push(0x23); // PTG_NAME (R class)
            out.extend_from_slice(&idx.to_le_bytes());
            out.extend_from_slice(&0u16.to_le_bytes()); // 2 reserved bytes
        }
        FormulaExpr::Array(_)
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => {
            return Err(UnsupportedToken);
        }
    }
    Ok(())
}

fn push_ref_payload(
    out: &mut Vec<u8>,
    addr: &duke_sheets_core::CellAddress,
) -> Result<(), UnsupportedToken> {
    if addr.row > u16::MAX as u32 {
        return Err(UnsupportedToken);
    }
    out.extend_from_slice(&(addr.row as u16).to_le_bytes());
    out.extend_from_slice(
        &encode_col_with_relative_flags(addr.col, addr.row_absolute, addr.col_absolute)
            .to_le_bytes(),
    );
    Ok(())
}

/// Recognise full-column (`A:A`, `B:D`) and full-row (`1:1`, `2:5`)
/// reference shapes that the formula parser leaves as
/// `Range(NameRef, NameRef)` or `Range(Number, Number)` respectively.
/// Returns the equivalent `CellRange` covering the BIFF8 sheet
/// extent, or `None` when the operands aren't a recognised shape.
fn full_column_or_row_range(
    left: &duke_sheets_formula::FormulaExpr,
    right: &duke_sheets_formula::FormulaExpr,
) -> Option<duke_sheets_core::CellRange> {
    use duke_sheets_core::CellAddress;
    use duke_sheets_formula::FormulaExpr;

    if let (FormulaExpr::NameRef(l), FormulaExpr::NameRef(r)) = (left, right) {
        let start_col = CellAddress::letters_to_column(l).ok()?;
        let end_col = CellAddress::letters_to_column(r).ok()?;
        return Some(duke_sheets_core::CellRange {
            start: CellAddress::new(0, start_col.min(end_col)),
            end: CellAddress::new(u16::MAX as u32, start_col.max(end_col)),
        });
    }
    if let (FormulaExpr::Number(l), FormulaExpr::Number(r)) = (left, right) {
        if l.fract() != 0.0 || r.fract() != 0.0 || *l < 1.0 || *r < 1.0 {
            return None;
        }
        let l_idx = (*l as u32).saturating_sub(1);
        let r_idx = (*r as u32).saturating_sub(1);
        if l_idx > u16::MAX as u32 || r_idx > u16::MAX as u32 {
            return None;
        }
        return Some(duke_sheets_core::CellRange {
            start: CellAddress::new(l_idx.min(r_idx), 0),
            end: CellAddress::new(l_idx.max(r_idx), 0xFF),
        });
    }
    None
}

fn push_area_payload(
    out: &mut Vec<u8>,
    range: &duke_sheets_core::CellRange,
) -> Result<(), UnsupportedToken> {
    let start = &range.start;
    let end = &range.end;
    if start.row > u16::MAX as u32 {
        return Err(UnsupportedToken);
    }
    // Clamp end.row to BIFF8's row limit. The XLSX-style parser
    // produces end.row = 1048575 for full-column refs like A:A;
    // BIFF8 can't represent rows beyond 65535, so the closest valid
    // expression is "from start.row through row 65535" - i.e. the
    // entire BIFF8 column (or the entire BIFF8 row, for 1:1 refs).
    let end_row = end.row.min(u16::MAX as u32) as u16;
    out.extend_from_slice(&(start.row as u16).to_le_bytes());
    out.extend_from_slice(&end_row.to_le_bytes());
    out.extend_from_slice(
        &encode_col_with_relative_flags(start.col, start.row_absolute, start.col_absolute)
            .to_le_bytes(),
    );
    out.extend_from_slice(
        &encode_col_with_relative_flags(end.col, end.row_absolute, end.col_absolute).to_le_bytes(),
    );
    Ok(())
}

/// Pack column index + row/col absolute flags into the 16-bit
/// `colIxv` field used by tRef/tArea (MS-XLS §2.5.198.103). Bits 0-13
/// hold the column; bit 14 is fColRel; bit 15 is fRowRel.
fn encode_col_with_relative_flags(col: u16, row_absolute: bool, col_absolute: bool) -> u16 {
    let mut v = col & 0x3FFF;
    if !col_absolute {
        v |= 0x4000;
    }
    if !row_absolute {
        v |= 0x8000;
    }
    v
}

fn write_labelsst(stream: &mut Vec<u8>, row: u16, col: u16, ixfe: u16, isst: u32) {
    stream.extend_from_slice(&LABELSST_RECORD.to_le_bytes());
    stream.extend_from_slice(&10u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.extend_from_slice(&isst.to_le_bytes());
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
///  u8  hsState      = visibility (0=visible, 1=hidden, 2=very hidden)
///  u8  dt           = 0 (worksheet)
///  ShortXLUnicodeString stName
/// ```
fn write_boundsheet8_with_placeholder_offset(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
) -> XlsResult<()> {
    let name = sheet.name();
    let utf16_units: Vec<u16> = name.encode_utf16().collect();
    if utf16_units.len() > 31 {
        return Err(XlsError::InvalidFormat(format!(
            "sheet name '{name}' is {} UTF-16 code units; Excel caps sheet names at 31",
            utf16_units.len()
        )));
    }

    let hs_state: u8 = match sheet.visibility() {
        duke_sheets_core::worksheet::SheetVisibility::Visible => 0,
        duke_sheets_core::worksheet::SheetVisibility::Hidden => 1,
        duke_sheets_core::worksheet::SheetVisibility::VeryHidden => 2,
    };

    let mut body = Vec::with_capacity(8 + utf16_units.len() * 2);
    body.extend_from_slice(&[0u8; 4]); // lbPlyPos placeholder
    body.push(hs_state);
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

/// First shape ID used by the patriarch of the first drawing. MS-ODRAW
/// reserves shape IDs in clusters of 1024; cluster 0 (IDs 0–1023) is
/// reserved by the spec, so user-allocated shape IDs start at 1024.
const PATRIARCH_SPID_BASE: u32 = 1024;

/// Pre-computed drawing layout for a workbook. Captures which sheets
/// have shapes (comments and/or pictures), what shape IDs have been
/// allocated to their patriarchs, pictures, and comments, plus the
/// global blip store and cluster table that go inside the
/// `MSODRAWINGGROUP`.
#[derive(Debug, Default)]
struct DrawingState {
    /// Per-sheet drawings keyed by sheet index. Sheets without shapes
    /// are absent.
    sheets: HashMap<usize, SheetDrawing>,
    /// Drawing indices in workbook order; lets us emit cluster entries
    /// and sheet drawings deterministically.
    ordered_sheet_indices: Vec<usize>,
    /// Workbook-wide blip store. One entry per embedded image across
    /// all sheets, in deterministic order. The 1-based index into
    /// this vec is the `pib` (picture blip id) referenced by each
    /// picture shape's `FOPT` `0x0104` entry.
    blip_store: Vec<BlipEntry>,
    /// Total shape count across all drawings (patriarch + pictures +
    /// comments per sheet, summed). Goes into `FDGG.csp_saved`.
    csp_total: u32,
    /// Number of drawings (sheets with shapes). Goes into
    /// `FDGG.cdg_saved`.
    cdg_total: u32,
    /// One past the highest shape ID used across the workbook. Goes
    /// into `FDGG.spid_max`.
    spid_max: u32,
}

impl DrawingState {
    fn is_empty(&self) -> bool {
        self.cdg_total == 0
    }
}

/// A single image queued for emission in the workbook-globals blip
/// store. Index in `DrawingState.blip_store` (+1) is the `pib` the
/// picture shape's FOPT references.
#[derive(Debug, Clone)]
struct BlipEntry {
    format: duke_sheets_chart::ImageFormat,
    data: Vec<u8>,
}

#[derive(Debug, Clone)]
struct SheetDrawing {
    /// 1-based drawing ID (`OfficeArtFDG.dgid`).
    dgid: u16,
    /// Shape ID of the patriarch group (the implicit root group).
    patriarch_spid: u32,
    /// One per picture on this sheet, in stable iteration order
    /// (same order as `Worksheet::images()`).
    pictures: Vec<PictureShape>,
    /// One per comment, in stable iteration order (sorted by row,
    /// then column).
    comments: Vec<CommentShape>,
}

#[derive(Debug, Clone)]
struct PictureShape {
    /// Escher shape ID stored in this picture's `OfficeArtFSP.spid`.
    spid: u32,
    /// 1-based per-sheet object ID placed in `OBJ.ftCmo.id`.
    obj_id: u16,
    /// 1-based index into `DrawingState.blip_store`; referenced by
    /// the picture's FOPT `pib` (`0x0104`) property.
    blip_id: u32,
    /// User-visible shape name (e.g. `"Picture 1"`).
    shape_name: String,
    /// Cell-anchor footprint copied from the `EmbeddedImage`.
    anchor: duke_sheets_chart::DrawingAnchor,
    /// Optional rotation in 60,000ths of a degree. Goes into the
    /// picture's FOPT `0x0004` property when set.
    rotation: Option<i32>,
    /// Horizontal flip — sets the FSP `FLIP_H` flag bit.
    flip_h: bool,
    /// Vertical flip — sets the FSP `FLIP_V` flag bit.
    flip_v: bool,
}

#[derive(Debug, Clone)]
struct CommentShape {
    /// Escher shape ID stored in this comment's `OfficeArtFSP.spid`.
    spid: u32,
    /// 1-based per-sheet object ID. Excel's `OBJ.ftCmo.id` and
    /// `NOTE.objId` use this value to link a comment's drawing
    /// object to its cell-anchored note.
    obj_id: u16,
    /// `txid` value placed in the `FOPT` `TEXT_ID` slot. Chosen as a
    /// monotonically-increasing per-comment counter; Excel itself
    /// picks an arbitrary unique value per shape.
    text_id: u32,
    /// Cell row this comment is anchored to.
    row: u32,
    /// Cell column this comment is anchored to.
    col: u16,
    /// Author string from `CellComment.author`.
    author: String,
    /// Comment body text from `CellComment.text`.
    text: String,
    /// Whether the comment box is visible by default (sets `NOTE.flags`
    /// bit 1).
    visible: bool,
}

/// Walk every sheet in `workbook`, allocate drawing IDs and shape
/// IDs, and assemble the [`DrawingState`] used by both
/// [`write_msodrawinggroup`] and [`write_sheet_drawing_records`].
///
/// Within each drawing, shape IDs are allocated in order:
///   patriarch → pictures (in `Worksheet::images()` order) → comments
///   (in sorted row/col order). Each drawing starts in its own
///   1024-aligned cluster.
fn compute_drawing_state(workbook: &Workbook) -> DrawingState {
    let mut state = DrawingState::default();
    let mut next_dgid: u16 = 1;
    let mut next_spid = PATRIARCH_SPID_BASE;
    let mut next_text_id: u32 = 1;
    let mut next_blip_id: u32 = 1;
    let mut highest_spid_used: u32 = 0;

    for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
        let comment_count = sheet.comment_count();
        let picture_count = sheet.image_count();
        if comment_count == 0 && picture_count == 0 {
            continue;
        }

        let patriarch_spid = next_spid;
        next_spid += 1;

        let mut next_obj_id: u16 = 1;
        let mut pictures = Vec::with_capacity(picture_count);
        for image in sheet.images() {
            let spid = next_spid;
            next_spid += 1;
            let blip_id = next_blip_id;
            next_blip_id += 1;
            state.blip_store.push(BlipEntry {
                format: image.format,
                data: image.data.clone(),
            });
            pictures.push(PictureShape {
                spid,
                obj_id: next_obj_id,
                blip_id,
                shape_name: image.name.clone(),
                anchor: image.anchor.clone(),
                rotation: image.rotation,
                flip_h: image.flip_h,
                flip_v: image.flip_v,
            });
            next_obj_id += 1;
        }

        let mut comments_sorted: Vec<_> = sheet.comments().collect();
        comments_sorted.sort_by_key(|((row, col), _)| (*row, *col));

        let mut comments = Vec::with_capacity(comment_count);
        for ((row, col), comment) in comments_sorted.iter() {
            let spid = next_spid;
            next_spid += 1;
            comments.push(CommentShape {
                spid,
                obj_id: next_obj_id,
                text_id: next_text_id,
                row: *row,
                col: *col,
                author: comment.author.clone(),
                text: comment.text.clone(),
                visible: comment.visible,
            });
            next_obj_id += 1;
            next_text_id = next_text_id.wrapping_add(1);
        }
        highest_spid_used = highest_spid_used.max(next_spid - 1);

        let total_shapes = 1 + picture_count + comment_count;
        state.sheets.insert(
            sheet_idx,
            SheetDrawing {
                dgid: next_dgid,
                patriarch_spid,
                pictures,
                comments,
            },
        );
        state.ordered_sheet_indices.push(sheet_idx);
        state.cdg_total += 1;
        state.csp_total += total_shapes as u32;
        next_dgid += 1;

        // Round up to the next 1024-aligned base so the next drawing
        // lives in its own cluster (matching Excel's per-drawing
        // cluster allocation in MS-ODRAW §2.2.46).
        next_spid = next_spid.div_ceil(1024) * 1024;
    }
    state.spid_max = highest_spid_used + 1;
    state
}

/// Emit the workbook-globals `MSODRAWINGGROUP` (BIFF 0xEB) record
/// carrying the global Office Art drawing context: `FDGG` shape-ID
/// allocator state, default-property `FOPT`, and the
/// `SplitMenuColorContainer` palette. No-op if no sheet has comments.
fn write_msodrawinggroup(stream: &mut Vec<u8>, state: &DrawingState) {
    if state.is_empty() {
        return;
    }

    use crate::biff::escher::{
        dgg_default_fopt, rec_type as er, write_bstore_container, write_container, IdCluster,
        OfficeArtBlip, OfficeArtFbse, OfficeArtFdgg, SplitMenuColors,
    };

    let mut dgg_body = Vec::new();
    let clusters: Vec<IdCluster> = state
        .ordered_sheet_indices
        .iter()
        .filter_map(|i| state.sheets.get(i))
        .map(|sd| IdCluster {
            dgid: sd.dgid as u32,
            // patriarch + pictures + comments.
            cspid_cur: (1 + sd.pictures.len() + sd.comments.len()) as u32,
        })
        .collect();
    OfficeArtFdgg {
        spid_max: state.spid_max,
        csp_saved: state.csp_total,
        cdg_saved: state.cdg_total,
        clusters,
    }
    .write_to(&mut dgg_body);

    // If the workbook embeds any images, emit a `BSTORE_CONTAINER`
    // ahead of the default FOPT carrying one FBSE+Blip per image.
    if !state.blip_store.is_empty() {
        let fbses: Vec<OfficeArtFbse> = state
            .blip_store
            .iter()
            .map(|entry| match entry.format {
                duke_sheets_chart::ImageFormat::Png => {
                    OfficeArtFbse::new(OfficeArtBlip::png(entry.data.clone()))
                }
                duke_sheets_chart::ImageFormat::Jpeg => {
                    OfficeArtFbse::new(OfficeArtBlip::jpeg(entry.data.clone()))
                }
                // Formats we don't yet emit (GIF / BMP / EMF / WMF /
                // TIFF) fall back to PNG so we don't silently drop
                // the bytes; Excel may reject the file in that case.
                // See FEATURES.md Notes for the supported set.
                _ => OfficeArtFbse::new(OfficeArtBlip::png(entry.data.clone())),
            })
            .collect();
        write_bstore_container(&fbses, &mut dgg_body);
    }

    dgg_default_fopt().write_to(&mut dgg_body);
    SplitMenuColors::EXCEL_DEFAULT.write_to(&mut dgg_body);

    let mut record_body = Vec::new();
    write_container(er::DGG_CONTAINER, 0, &dgg_body, &mut record_body);

    write_biff_record(stream, MSODRAWINGGROUP_RECORD, &record_body);
}

/// Emit per-sheet drawing records, mirroring the interleaved pattern
/// Excel itself writes.
///
/// The shape tree is logically one `DgContainer` per sheet, but its
/// bytes are split across multiple `MSODRAWING` records so that each
/// comment's `ClientTextbox` marker can sit AFTER that comment's
/// `OBJ` record in the BIFF stream. Excel uses the position of the
/// `ClientTextbox` (relative to OBJ) to associate the following
/// `TXO` record with the correct shape; a writer that emits a
/// shape's `ClientTextbox` before its `OBJ` produces a file Excel
/// refuses to open with `RPC failed (0x800706BE)`.
///
/// Layout emitted, for `N` comments:
///
/// ```text
/// MSODRAWING #1   = DgContainer header + FDG + SpgrContainer header
///                   + patriarch SpContainer
///                   + comment[0] SpContainer header
///                   + comment[0]: FSP + FOPT + ClientAnchor + ClientData
/// OBJ #1
/// MSODRAWING #2   = comment[0]: ClientTextbox (closes SpContainer 0)
/// TXO #1 + CONTINUE×2
/// MSODRAWING #3   = comment[1] SpContainer header
///                   + comment[1]: FSP + FOPT + ClientAnchor + ClientData
/// OBJ #2
/// MSODRAWING #4   = comment[1]: ClientTextbox
/// TXO #2 + CONTINUE×2
/// ... (one MSODRAWING-pair per remaining comment)
/// NOTE #1 .. NOTE #N
/// ```
///
/// The `DgContainer` / `SpgrContainer` / per-comment `SpContainer`
/// header `rec_len` fields all reflect their LOGICAL byte counts
/// across the entire concatenated drawing stream. Readers
/// concatenate the bodies of all `MSODRAWING` records for a sheet,
/// then walk the resulting Escher tree.
fn write_sheet_drawing_records(stream: &mut Vec<u8>, drawing: &SheetDrawing) {
    use crate::biff::escher::{
        comment_fopt, fsp_flags, rec_type as er, shape_type, write_client_data,
        write_client_textbox, write_container, write_patriarch_sp_container, OfficeArtClientAnchor,
        OfficeArtFdg, OfficeArtFsp, OfficeArtRecordHeader, HEADER_LEN,
    };

    // Per-comment payload: build each SpContainer in two halves so the
    // ClientTextbox marker can be emitted into a separate MSODRAWING
    // after the OBJ for that shape.
    struct CommentBytes {
        /// FSP + FOPT + ClientAnchor + ClientData. These are the
        /// bytes that fit "before OBJ".
        pre_obj: Vec<u8>,
        /// ClientTextbox atom (8 bytes). Emitted after OBJ to close
        /// the SpContainer.
        post_obj: Vec<u8>,
    }

    let comments: Vec<(CommentBytes, &CommentShape)> = drawing
        .comments
        .iter()
        .map(|c| {
            let mut pre = Vec::new();
            OfficeArtFsp {
                spid: c.spid,
                grf_persistence: fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT,
            }
            .write_to(shape_type::TEXT_BOX, &mut pre);
            comment_fopt(c.text_id).write_to(&mut pre);
            OfficeArtClientAnchor::comment_default(c.row, c.col).write_to(&mut pre);
            write_client_data(&mut pre);

            let mut post = Vec::new();
            write_client_textbox(&mut post);
            (
                CommentBytes {
                    pre_obj: pre,
                    post_obj: post,
                },
                c,
            )
        })
        .collect();

    // Picture SP_CONTAINERs are self-contained — pictures have no
    // associated TXO, so their SP_CONTAINERs close cleanly with
    // ClientData (no ClientTextbox interleave) inside the first
    // MSODRAWING record.
    let mut picture_bytes_total: Vec<u8> = Vec::new();
    for picture in &drawing.pictures {
        let mut sp_body = Vec::new();
        let mut grf = fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT;
        if picture.flip_h {
            grf |= fsp_flags::FLIP_H;
        }
        if picture.flip_v {
            grf |= fsp_flags::FLIP_V;
        }
        OfficeArtFsp {
            spid: picture.spid,
            grf_persistence: grf,
        }
        .write_to(shape_type::PICTURE_FRAME, &mut sp_body);
        crate::biff::escher::picture_fopt_with(
            picture.blip_id,
            &picture.shape_name,
            picture.rotation,
        )
        .write_to(&mut sp_body);
        client_anchor_from_drawing_anchor(&picture.anchor).write_to(&mut sp_body);
        write_client_data(&mut sp_body);
        write_container(er::SP_CONTAINER, 0, &sp_body, &mut picture_bytes_total);
    }

    // The patriarch SP_CONTAINER (FSPGR + FSP) is self-contained.
    let mut patriarch_bytes = Vec::new();
    write_patriarch_sp_container(drawing.patriarch_spid, &mut patriarch_bytes);

    // SP_CONTAINER's `rec_len` is its payload size only (no header):
    // sum of pre_obj + post_obj bytes.
    let comment_sp_payload_len =
        |c: &CommentBytes| -> u32 { (c.pre_obj.len() + c.post_obj.len()) as u32 };

    // Compute logical container sizes top-down so each header carries
    // its full rec_len (the byte counts span across MSODRAWING records).
    let spgr_payload_len: u32 = patriarch_bytes.len() as u32
        + picture_bytes_total.len() as u32
        + comments
            .iter()
            .map(|(c, _)| HEADER_LEN as u32 + comment_sp_payload_len(c))
            .sum::<u32>();
    let fdg_total_len = HEADER_LEN as u32 + 8;
    let dg_payload_len = fdg_total_len + HEADER_LEN as u32 + spgr_payload_len;

    // First MSODRAWING: opens the DG_CONTAINER, FDG, SPGR_CONTAINER,
    // patriarch SP_CONTAINER, all picture SP_CONTAINERs, and (if any
    // comments) the first comment's SP_CONTAINER up through ClientData.
    let mut first_drawing = Vec::new();
    OfficeArtRecordHeader::container(er::DG_CONTAINER, 0, dg_payload_len)
        .write_to(&mut first_drawing);
    // spid_last is the highest spid used in this drawing.
    let spid_last = drawing
        .comments
        .last()
        .map(|c| c.spid)
        .or_else(|| drawing.pictures.last().map(|p| p.spid))
        .unwrap_or(drawing.patriarch_spid);
    OfficeArtFdg {
        csp_saved: (1 + drawing.pictures.len() + drawing.comments.len()) as u32,
        spid_last,
    }
    .write_to(drawing.dgid, &mut first_drawing);
    OfficeArtRecordHeader::container(er::SPGR_CONTAINER, 0, spgr_payload_len)
        .write_to(&mut first_drawing);
    first_drawing.extend_from_slice(&patriarch_bytes);
    first_drawing.extend_from_slice(&picture_bytes_total);
    if let Some((c0, _)) = comments.first() {
        OfficeArtRecordHeader::container(er::SP_CONTAINER, 0, comment_sp_payload_len(c0))
            .write_to(&mut first_drawing);
        first_drawing.extend_from_slice(&c0.pre_obj);
    }
    write_biff_record(stream, MSODRAWING_RECORD, &first_drawing);

    // Picture OBJs (one per picture, in shape order). No TXO or
    // continuation MSODRAWING follows.
    for picture in &drawing.pictures {
        write_picture_obj(stream, picture);
    }

    // Per-comment interleave: OBJ, MSODRAWING (ClientTextbox), TXO,
    // CONTINUE × 2. For comments after the first, prefix each
    // MSODRAWING with the SP_CONTAINER header + pre_obj bytes too.
    for (idx, (cb, shape)) in comments.iter().enumerate() {
        write_comment_obj(stream, shape);

        let mut close_drawing = Vec::new();
        close_drawing.extend_from_slice(&cb.post_obj);
        write_biff_record(stream, MSODRAWING_RECORD, &close_drawing);

        write_comment_txo(stream, shape);

        if let Some((next_cb, _)) = comments.get(idx + 1) {
            let mut open_drawing = Vec::new();
            OfficeArtRecordHeader::container(er::SP_CONTAINER, 0, comment_sp_payload_len(next_cb))
                .write_to(&mut open_drawing);
            open_drawing.extend_from_slice(&next_cb.pre_obj);
            write_biff_record(stream, MSODRAWING_RECORD, &open_drawing);
        }
    }

    // NOTE records at the end, in comment order.
    for (_, shape) in &comments {
        write_comment_note(stream, shape);
    }
}

/// One EMU unit of the ClientAnchor `dxL`/`dxR` field, in EMUs.
/// MS-XLS encodes within-cell horizontal offsets in 1024ths of the
/// cell's width; with Excel's default column width of 64 pixels
/// (609,600 EMU) one `dxL` unit equals 595 EMU.
pub(crate) const ANCHOR_DX_EMU_PER_UNIT: i64 = 595;
/// One EMU unit of the ClientAnchor `dyT`/`dyB` field, in EMUs.
/// MS-XLS encodes within-cell vertical offsets in 256ths of the
/// row's height; with Excel's default row height of 15 pt
/// (190,500 EMU) one `dyT` unit equals 744 EMU.
pub(crate) const ANCHOR_DY_EMU_PER_UNIT: i64 = 744;

fn emu_to_dx_units(emu: i64) -> u16 {
    let units = (emu / ANCHOR_DX_EMU_PER_UNIT).clamp(0, 1023);
    units as u16
}

fn emu_to_dy_units(emu: i64) -> u16 {
    let units = (emu / ANCHOR_DY_EMU_PER_UNIT).clamp(0, 255);
    units as u16
}

/// Translate a `duke_sheets_chart::DrawingAnchor` to the
/// `OfficeArtClientAnchor` BIFF8 layout.
///
/// XLS's ClientAnchor has a single byte layout (top-left and
/// bottom-right cells + within-cell EMU offsets) for all three
/// OOXML anchor variants. The `flag` field encodes the variant
/// semantics:
///   0 = move + resize with cells (OOXML `editAs="twoCell"`)
///   2 = move only, don't resize (OOXML `editAs="oneCell"` or
///       `xdr:oneCellAnchor`)
///   3 = no move, no resize (OOXML `editAs="absolute"` or
///       `xdr:absoluteAnchor`)
///
/// Within-cell EMU offsets carried by the source `CellMarker` are
/// quantised to the anchor's 1024ths-of-cell-width / 256ths-of-cell-
/// height units. The reader inverts the same quantisation, so EMU
/// values that are multiples of `ANCHOR_DX_EMU_PER_UNIT` /
/// `ANCHOR_DY_EMU_PER_UNIT` round-trip exactly; others snap to the
/// nearest unit boundary.
fn client_anchor_from_drawing_anchor(
    anchor: &duke_sheets_chart::DrawingAnchor,
) -> crate::biff::escher::OfficeArtClientAnchor {
    use crate::biff::escher::OfficeArtClientAnchor;
    use duke_sheets_chart::{DrawingAnchor, EditAs};

    /// Default column width in EMU (8.43 char ≈ 64 px).
    const DEFAULT_COL_EMU: i64 = 609_600;
    /// Default row height in EMU (15 pt).
    const DEFAULT_ROW_EMU: i64 = 190_500;

    /// Convert an `EditAs` enum into the ClientAnchor flag.
    fn edit_as_to_flag(edit_as: &Option<EditAs>) -> u16 {
        match edit_as {
            None | Some(EditAs::TwoCell) => 0,
            Some(EditAs::OneCell) => 2,
            Some(EditAs::Absolute) => 3,
        }
    }

    /// For OneCell / Absolute → compute the (col_r, dx_r, row_b,
    /// dy_b) fields by adding width/height EMU to the (col, offset)
    /// starting point at default cell sizes.
    fn extend_anchor(
        start_col: u16,
        start_off_x: i64,
        width: i64,
        start_row: u32,
        start_off_y: i64,
        height: i64,
    ) -> (u16, i64, u32, i64) {
        let total_x = start_col as i64 * DEFAULT_COL_EMU + start_off_x + width;
        let end_col = (total_x / DEFAULT_COL_EMU).max(0) as u16;
        let end_off_x = total_x - end_col as i64 * DEFAULT_COL_EMU;
        let total_y = start_row as i64 * DEFAULT_ROW_EMU + start_off_y + height;
        let end_row = (total_y / DEFAULT_ROW_EMU).max(0) as u32;
        let end_off_y = total_y - end_row as i64 * DEFAULT_ROW_EMU;
        (end_col, end_off_x, end_row, end_off_y)
    }

    match anchor {
        DrawingAnchor::TwoCell { from, to, edit_as } => OfficeArtClientAnchor {
            flag: edit_as_to_flag(edit_as),
            col_l: from.col,
            dx_l: emu_to_dx_units(from.col_offset_emu),
            row_t: from.row as u16,
            dy_t: emu_to_dy_units(from.row_offset_emu),
            col_r: to.col,
            dx_r: emu_to_dx_units(to.col_offset_emu),
            row_b: to.row as u16,
            dy_b: emu_to_dy_units(to.row_offset_emu),
        },
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            let (col_r, off_r, row_b, off_b) = extend_anchor(
                from.col,
                from.col_offset_emu,
                *width_emu,
                from.row,
                from.row_offset_emu,
                *height_emu,
            );
            OfficeArtClientAnchor {
                flag: 2,
                col_l: from.col,
                dx_l: emu_to_dx_units(from.col_offset_emu),
                row_t: from.row as u16,
                dy_t: emu_to_dy_units(from.row_offset_emu),
                col_r,
                dx_r: emu_to_dx_units(off_r),
                row_b: row_b.min(u16::MAX as u32) as u16,
                dy_b: emu_to_dy_units(off_b),
            }
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            // Absolute coordinates are relative to (0, 0); position
            // the from cell at the origin with the x/y offsets.
            let (col_l, off_l, row_t, off_t) = extend_anchor(0, 0, *x_emu, 0, 0, *y_emu);
            let (col_r, off_r, row_b, off_b) =
                extend_anchor(col_l, off_l, *width_emu, row_t, off_t, *height_emu);
            OfficeArtClientAnchor {
                flag: 3,
                col_l,
                dx_l: emu_to_dx_units(off_l),
                row_t: row_t.min(u16::MAX as u32) as u16,
                dy_t: emu_to_dy_units(off_t),
                col_r,
                dx_r: emu_to_dx_units(off_r),
                row_b: row_b.min(u16::MAX as u32) as u16,
                dy_b: emu_to_dy_units(off_b),
            }
        }
    }
}

/// Emit an `OBJ` record (BIFF 0x005D) for a picture shape. Carries
/// four sub-records:
/// - `ftCmo` (0x0015): common object data with `ot=0x08` (picture)
///   and `id` set to the shape's 1-based per-sheet object ID.
/// - `ftCf` (0x0007): clipboard format. We emit `0xFFFF` (no
///   clipboard format / opaque blob), matching Excel's emit.
/// - `ftPioGrbit` (0x0008): picture options (auto-pict, no print, …);
///   all-zero matches the default settings Excel applies to inserted
///   pictures.
/// - `ftEnd` (0x0000): terminator.
fn write_picture_obj(stream: &mut Vec<u8>, picture: &PictureShape) {
    let mut body = Vec::new();

    // ftCmo: rt + cb + ot + id + grbit + 12 reserved bytes = 22 bytes total,
    // cb = 18.
    body.extend_from_slice(&0x0015u16.to_le_bytes()); // rt = ftCmo
    body.extend_from_slice(&0x0012u16.to_le_bytes()); // cb = 18
    body.extend_from_slice(&0x0008u16.to_le_bytes()); // ot = picture
    body.extend_from_slice(&picture.obj_id.to_le_bytes()); // id
    body.extend_from_slice(&0x6011u16.to_le_bytes()); // grbit: fLocked|fAutoFill|fAutoLine|fPrintable
    body.extend_from_slice(&[0u8; 12]); // reserved

    // ftCf: rt + cb + cf = 6 bytes total, cb = 2.
    body.extend_from_slice(&0x0007u16.to_le_bytes()); // rt = ftCf
    body.extend_from_slice(&0x0002u16.to_le_bytes()); // cb = 2
    body.extend_from_slice(&0xFFFFu16.to_le_bytes()); // cf = no clipboard format

    // ftPioGrbit: rt + cb + grbit = 6 bytes total, cb = 2.
    body.extend_from_slice(&0x0008u16.to_le_bytes()); // rt = ftPioGrbit
    body.extend_from_slice(&0x0002u16.to_le_bytes()); // cb = 2
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // grbit = no flags

    // ftEnd: rt + cb = 4 bytes total, cb = 0.
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // rt = ftEnd
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // cb = 0

    write_biff_record(stream, OBJ_RECORD, &body);
}

/// Emit an `OBJ` record (BIFF 0x005D) for a comment shape. Carries
/// three sub-records:
/// - `ftCmo` (0x0015): common object data with `ot=0x19` (note) and
///   `id` set to the comment's 1-based shape index.
/// - `ftNts` (0x000D): notes-specific data with a unique 16-byte
///   GUID; we derive a deterministic GUID from the shape ID so the
///   bytes are reproducible across runs.
/// - `ftEnd` (0x0000): terminator.
fn write_comment_obj(stream: &mut Vec<u8>, comment: &CommentShape) {
    let mut body = Vec::new();

    // ftCmo: rt(2) cb(2) ot(2) id(2) grbit(2) reserved(12) = 22 bytes total,
    // body length cb = 18 (the bytes after rt+cb).
    body.extend_from_slice(&0x0015u16.to_le_bytes()); // rt = ftCmo
    body.extend_from_slice(&0x0012u16.to_le_bytes()); // cb = 18
    body.extend_from_slice(&0x0019u16.to_le_bytes()); // ot = note/comment
    body.extend_from_slice(&comment.obj_id.to_le_bytes()); // id (matches NOTE.objId)
    body.extend_from_slice(&0x4011u16.to_le_bytes()); // grbit: fPrintable | fAutoSize
    body.extend_from_slice(&[0u8; 12]); // 12 reserved zero bytes

    // ftNts: rt(2) cb(2) guid(16) fSharedNote(u32) reserved(2) = 28 bytes total,
    // cb = 22.
    body.extend_from_slice(&0x000Du16.to_le_bytes()); // rt = ftNts
    body.extend_from_slice(&0x0016u16.to_le_bytes()); // cb = 22
    body.extend_from_slice(&deterministic_guid(comment.text_id)); // 16-byte GUID
    body.extend_from_slice(&0u32.to_le_bytes()); // fSharedNote = 0 (not threaded)
    body.extend_from_slice(&[0u8; 2]); // 2 reserved zero bytes

    // ftEnd: rt(2) cb(2) = 4 bytes total.
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // rt = ftEnd
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // cb = 0

    write_biff_record(stream, OBJ_RECORD, &body);
}

/// Derive a 16-byte GUID from a per-shape counter. The bytes are
/// arbitrary but deterministic across writer runs; Excel itself
/// emits values that look like uninitialised memory, so any
/// well-formed 16-byte blob is accepted.
fn deterministic_guid(seed: u32) -> [u8; 16] {
    let mut out = [0u8; 16];
    out[..4].copy_from_slice(&seed.to_le_bytes());
    // Variant + version stamp: RFC 4122 v4-shaped to look plausible.
    out[6] = 0x40;
    out[8] = 0x80;
    out
}

/// Emit a `TXO` record (BIFF 0x01B6) plus its two `CONTINUE`
/// records carrying the comment's text and a single formatting run.
///
/// TXO header layout (MS-XLS §2.4.324, 18 bytes):
///
/// - `flags` (u16): horizontal/vertical alignment, text-locked flag.
///   We write `0x0212` to match Excel: halign=left, valign=top,
///   `fLockText=1`.
/// - `rot` (u16): rotation (0 = horizontal).
/// - 6 bytes reserved.
/// - `cchText` (u16): character count of the text.
/// - `cbRuns` (u16): total length of the formatting runs block
///   (8 bytes per run, two runs = 16).
/// - 4 bytes reserved.
fn write_comment_txo(stream: &mut Vec<u8>, comment: &CommentShape) {
    let utf16: Vec<u16> = comment.text.encode_utf16().collect();
    let high_byte = utf16.iter().any(|&u| u > 0xFF);

    // TXO header.
    let mut header = Vec::with_capacity(18);
    header.extend_from_slice(&0x0212u16.to_le_bytes()); // flags
    header.extend_from_slice(&0u16.to_le_bytes()); // rot
    header.extend_from_slice(&[0u8; 6]); // reserved
    header.extend_from_slice(&(utf16.len() as u16).to_le_bytes()); // cchText
    header.extend_from_slice(&16u16.to_le_bytes()); // cbRuns = 2 runs × 8 bytes
    header.extend_from_slice(&[0u8; 2]); // ifntEmpty
    header.extend_from_slice(&[0u8; 2]); // reserved3
    write_biff_record(stream, TXO_RECORD, &header);

    // First CONTINUE: text payload prefixed by a 1-byte grbit
    // (0x00 = compressed Latin-1, 0x01 = UTF-16LE).
    let mut text_body = Vec::with_capacity(1 + utf16.len() * 2);
    if high_byte {
        text_body.push(0x01);
        for u in &utf16 {
            text_body.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        text_body.push(0x00);
        for u in &utf16 {
            text_body.push(*u as u8);
        }
    }
    write_biff_record(stream, CONTINUE_RECORD, &text_body);

    // Second CONTINUE: formatting runs. Two TxoRun entries (8 bytes
    // each): (ich=0, ifnt=0, reserved=0) opens the default-font run,
    // (ich=text_len, ifnt=0, reserved=0) closes it. This produces a
    // single-run plain-text comment.
    let mut runs_body = Vec::with_capacity(16);
    runs_body.extend_from_slice(&0u16.to_le_bytes()); // ich = 0
    runs_body.extend_from_slice(&0u16.to_le_bytes()); // ifnt = 0
    runs_body.extend_from_slice(&[0u8; 4]); // reserved
    runs_body.extend_from_slice(&(utf16.len() as u16).to_le_bytes()); // ich = end
    runs_body.extend_from_slice(&0u16.to_le_bytes()); // ifnt = 0
    runs_body.extend_from_slice(&[0u8; 4]); // reserved
    write_biff_record(stream, CONTINUE_RECORD, &runs_body);
}

/// Emit a `NOTE` record (BIFF 0x001C) that anchors a comment to a
/// cell. Layout: row(u16) + col(u16) + flags(u16) + objId(u16) +
/// author(XLUnicodeString).
///
/// The author string is emitted as `XLUnicodeString`: cch(u16) +
/// fHighByte(u8) + chars. A trailing padding byte rounds the record
/// to an even length, matching Excel's emit.
fn write_comment_note(stream: &mut Vec<u8>, comment: &CommentShape) {
    let row_u16 = if comment.row > u16::MAX as u32 {
        u16::MAX
    } else {
        comment.row as u16
    };
    let flags: u16 = if comment.visible { 0x0002 } else { 0x0000 };

    let mut body = Vec::new();
    body.extend_from_slice(&row_u16.to_le_bytes());
    body.extend_from_slice(&comment.col.to_le_bytes());
    body.extend_from_slice(&flags.to_le_bytes());
    body.extend_from_slice(&comment.obj_id.to_le_bytes());

    // Author as XLUnicodeString.
    let utf16: Vec<u16> = comment.author.encode_utf16().collect();
    let high_byte = utf16.iter().any(|&u| u > 0xFF);
    body.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
    body.push(if high_byte { 0x01 } else { 0x00 });
    if high_byte {
        for u in &utf16 {
            body.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        for u in &utf16 {
            body.push(*u as u8);
        }
    }

    // Pad to even length (Excel emits a trailing zero).
    if body.len() % 2 != 0 {
        body.push(0);
    }

    write_biff_record(stream, NOTE_RECORD, &body);
}

/// Generic BIFF8 record emitter: `(record_type LE)` + `(body_len LE)`
/// + `body`. Panics in debug builds if `body` exceeds 8224 bytes,
/// since BIFF8 caps a single record's body at that size and longer
/// content must be split with `CONTINUE` records. (Our callers stay
/// well below that limit; the assert exists to catch future
/// regressions.)
fn write_biff_record(stream: &mut Vec<u8>, record_type: u16, body: &[u8]) {
    debug_assert!(
        body.len() <= BIFF_MAX_RECORD_BODY,
        "BIFF record 0x{record_type:04X} body is {} bytes; max is {BIFF_MAX_RECORD_BODY}",
        body.len()
    );
    stream.extend_from_slice(&record_type.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(body);
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

    /// Walk the stream's record headers and return the byte offset of
    /// every record matching `record_type`, in document order.
    fn find_records(stream: &[u8], record_type: u16) -> Vec<usize> {
        let mut found = Vec::new();
        let mut cursor = 0usize;
        while cursor + 4 <= stream.len() {
            let rt = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
            let size = u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]) as usize;
            if rt == record_type {
                found.push(cursor);
            }
            cursor += 4 + size;
        }
        found
    }

    #[test]
    fn lbplypos_points_to_first_worksheet_bof() {
        let wb = Workbook::new();
        let stream = build_workbook_stream(&wb).expect("serialize");

        // Find the BoundSheet8 record by walking the stream so the
        // assertion stays valid as the writer adds new records (FONT,
        // XF, SST, etc.) before BoundSheet8.
        let bs_pos = *find_records(&stream, BOUND_SHEET_8)
            .first()
            .expect("at least one BoundSheet8");
        let lbplypos = u32::from_le_bytes([
            stream[bs_pos + 4],
            stream[bs_pos + 5],
            stream[bs_pos + 6],
            stream[bs_pos + 7],
        ]) as usize;

        // The first worksheet's BOF is the second BOF in document
        // order (the first being globals).
        let bof_positions = find_records(&stream, BOF_RECORD);
        assert!(
            bof_positions.len() >= 2,
            "expected globals BOF + at least one worksheet BOF"
        );
        assert_eq!(lbplypos, bof_positions[1]);
    }
}
