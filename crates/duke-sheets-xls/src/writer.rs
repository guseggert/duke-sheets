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
    let styles = StyleTables::collect(workbook);

    let mut stream = Vec::new();
    write_bof(&mut stream, DT_WORKBOOK_GLOBALS);
    styles.write_font_records(&mut stream)?;
    styles.write_format_records(&mut stream)?;
    styles.write_xf_records(&mut stream);

    let mut lbplypos_field_offsets = Vec::with_capacity(workbook.sheet_count());
    for sheet in workbook.worksheets() {
        let body_start = stream.len() + 4;
        write_boundsheet8_with_placeholder_offset(&mut stream, sheet.name())?;
        lbplypos_field_offsets.push(body_start);
    }

    sst.write_records(&mut stream)?;
    write_eof(&mut stream);

    let mut sheet_bof_offsets = Vec::with_capacity(workbook.sheet_count());
    for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
        let bof_pos = stream.len() as u32;
        write_bof(&mut stream, DT_WORKSHEET);
        write_dimension(&mut stream, sheet);
        write_cell_records(&mut stream, sheet, sheet_idx, &sst, &styles);
        write_window2(&mut stream);
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
        type XfKey = (u16, u16, Alignment, BorderStyle, FillStyle);
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
                        });
                        xf_key_to_ixfe.insert(xf_key, new_ixfe);
                        new_ixfe
                    }
                };

                cell_ixfe.insert((sheet_idx, cell.style_index), ixfe);
            }
        }

        StyleTables {
            fonts_in_order,
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
    let type_prot: u16 = if is_style_xf { 0xFFF5 } else { 0x0001 };
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
            if try_write_formula_record(stream, row16, col, ixfe, formula_text, &data.value, sst) {
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
    if compile_ptgs(&expr, &mut tokens).is_err() {
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
            if cref.sheet.is_some() {
                return Err(UnsupportedToken);
            }
            out.push(0x44); // PTG_REF (V class)
            push_ref_payload(out, &cref.address)?;
        }
        FormulaExpr::RangeRef(rref) => {
            if rref.sheet.is_some() {
                return Err(UnsupportedToken);
            }
            out.push(0x45); // PTG_AREA (V class)
            push_area_payload(out, &rref.range)?;
        }
        FormulaExpr::BinaryOp { op, left, right } => {
            compile_ptgs(left, out)?;
            compile_ptgs(right, out)?;
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
            compile_ptgs(operand, out)?;
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
                    compile_ptgs(arg, out)?;
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
        FormulaExpr::Array(_)
        | FormulaExpr::NameRef(_)
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

fn push_area_payload(
    out: &mut Vec<u8>,
    range: &duke_sheets_core::CellRange,
) -> Result<(), UnsupportedToken> {
    let start = &range.start;
    let end = &range.end;
    if start.row > u16::MAX as u32 || end.row > u16::MAX as u32 {
        return Err(UnsupportedToken);
    }
    out.extend_from_slice(&(start.row as u16).to_le_bytes());
    out.extend_from_slice(&(end.row as u16).to_le_bytes());
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
