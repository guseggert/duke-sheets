use std::collections::HashMap;
use std::io::{Read, Seek};

use duke_sheets_core::style::{
    Alignment, BorderEdge, BorderLineStyle, BorderStyle, Color, FillStyle, FontStyle, GradientStop,
    GradientType, HorizontalAlignment, NumberFormat, PatternType, Protection, ReadingOrder, Style,
    Underline, VerticalAlignment,
};


use crate::biff12::parser;
use crate::biff12::records;
use crate::biff12::RecordIter;
use crate::error::XlsbResult;

pub(crate) struct StylesData {
    pub styles: Vec<Style>,
    pub fonts: Vec<FontStyle>,
    pub dxf_styles: Vec<Style>,
}

pub(crate) fn read_styles<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> XlsbResult<StylesData> {
    let file = match archive.by_name("xl/styles.bin") {
        Ok(f) => f,
        Err(_) => {
            return Ok(StylesData {
                styles: vec![Style::default()],
                fonts: Vec::new(),
                dxf_styles: Vec::new(),
            })
        }
    };

    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(4096);

    let mut numfmts: HashMap<u32, String> = HashMap::new();
    let mut fonts: Vec<FontStyle> = Vec::new();
    let mut fills: Vec<FillStyle> = Vec::new();
    let mut borders: Vec<BorderStyle> = Vec::new();
    let mut cell_style_xfs: Vec<RawXf> = Vec::new();
    let mut cell_xfs: Vec<RawXf> = Vec::new();
    let mut dxf_styles: Vec<Style> = Vec::new();

    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    enum Section {
        None,
        Fonts,
        Fills,
        Borders,
        Fmts,
        CellStyleXfs,
        CellXfs,
        Dxfs,
    }
    let mut section = Section::None;

    loop {
        let (typ, len) = match iter.next_record(&mut buf) {
            Ok(r) => r,
            Err(_) => break,
        };

        match typ {
            records::BRT_BEGIN_FONTS => section = Section::Fonts,
            records::BRT_END_FONTS => section = Section::None,
            records::BRT_BEGIN_FILLS => section = Section::Fills,
            records::BRT_END_FILLS => section = Section::None,
            records::BRT_BEGIN_BORDERS => section = Section::Borders,
            records::BRT_END_BORDERS => section = Section::None,
            records::BRT_BEGIN_FMTS => section = Section::Fmts,
            records::BRT_END_FMTS => section = Section::None,
            records::BRT_BEGIN_CELL_STYLE_XFS => section = Section::CellStyleXfs,
            records::BRT_END_CELL_STYLE_XFS => section = Section::None,
            records::BRT_BEGIN_CELL_XFS => section = Section::CellXfs,
            records::BRT_END_CELL_XFS => section = Section::None,
            0x01F9 => section = Section::Dxfs, // BrtBeginDXFs
            0x01FA => section = Section::None, // BrtEndDXFs

            records::BRT_FMT if section == Section::Fmts => {
                if len >= 2 {
                    parse_fmt(&buf[..len], &mut numfmts);
                }
            }
            records::BRT_FONT if section == Section::Fonts => {
                fonts.push(parse_font(&buf[..len]));
            }
            records::BRT_FILL if section == Section::Fills => {
                fills.push(parse_fill(&buf[..len]));
            }
            records::BRT_BORDER if section == Section::Borders => {
                borders.push(parse_border(&buf[..len]));
            }
            records::BRT_XF if section == Section::CellStyleXfs => {
                if len >= 16 {
                    cell_style_xfs.push(parse_xf(&buf[..len]));
                }
            }
            records::BRT_XF if section == Section::CellXfs => {
                if len >= 16 {
                    cell_xfs.push(parse_xf(&buf[..len]));
                }
            }
            0x01FB if section == Section::Dxfs => dxf_styles.push(parse_dxf(&buf[..len])),
            _ => {}
        }
    }

    let resolved_bases: Vec<Style> = cell_style_xfs
        .iter()
        .map(|xf| resolve_style(xf, &numfmts, &fonts, &fills, &borders))
        .collect();

    let styles: Vec<Style> = if cell_xfs.is_empty() {
        vec![Style::default()]
    } else {
        cell_xfs
            .iter()
            .map(|xf| {
                let resolved = resolve_style(xf, &numfmts, &fonts, &fills, &borders);
                if let Some(base) = resolved_bases.get(xf.xf_id as usize).cloned() {
                    merge_xf_with_base(base, &resolved, xf)
                } else {
                    resolved
                }
            })
            .collect()
    };

    Ok(StylesData {
        styles,
        fonts,
        dxf_styles,
    })
}

#[derive(Debug, Clone)]
struct RawXf {
    xf_id: u16,
    num_fmt_id: u16,
    font_id: u16,
    fill_id: u16,
    border_id: u16,
    alignment: Alignment,
    protection: Protection,
    apply_number_format: bool,
    apply_font: bool,
    apply_fill: bool,
    apply_border: bool,
    apply_alignment: bool,
    apply_protection: bool,
}

fn resolve_style(
    xf: &RawXf,
    numfmts: &HashMap<u32, String>,
    fonts: &[FontStyle],
    fills: &[FillStyle],
    borders: &[BorderStyle],
) -> Style {
    let num_fmt_id = xf.num_fmt_id as u32;
    let number_format = if num_fmt_id == 0 {
        NumberFormat::General
    } else if let Some(code) = numfmts.get(&num_fmt_id) {
        NumberFormat::Custom(code.clone())
    } else {
        NumberFormat::BuiltIn(num_fmt_id)
    };

    Style {
        font: fonts.get(xf.font_id as usize).cloned().unwrap_or_default(),
        fill: fills.get(xf.fill_id as usize).cloned().unwrap_or_default(),
        border: borders
            .get(xf.border_id as usize)
            .cloned()
            .unwrap_or_default(),
        alignment: xf.alignment.clone(),
        protection: xf.protection,
        number_format,
    }
}

fn merge_xf_with_base(base: Style, resolved: &Style, xf: &RawXf) -> Style {
    let mut out = base;
    if xf.apply_number_format || xf.num_fmt_id != 0 {
        out.number_format = resolved.number_format.clone();
    }
    if xf.apply_font || xf.font_id != 0 {
        out.font = resolved.font.clone();
    }
    if xf.apply_fill || xf.fill_id != 0 {
        out.fill = resolved.fill.clone();
    }
    if xf.apply_border || xf.border_id != 0 {
        out.border = resolved.border.clone();
    }
    if xf.apply_alignment || xf.alignment != Alignment::default() {
        out.alignment = resolved.alignment.clone();
    }
    if xf.apply_protection || xf.protection != Protection::default() {
        out.protection = resolved.protection;
    }
    out
}

fn parse_fmt(buf: &[u8], numfmts: &mut HashMap<u32, String>) {
    if buf.len() < 4 {
        return;
    }
    let id = parser::read_u16(buf, 0) as u32;
    if let Ok((code, _)) = parser::wide_str(buf, 2) {
        numfmts.insert(id, code);
    }
}

/// Parse BrtXF record.
///
/// Layout per [MS-XLSB] 2.4.876:
///   0..2   ixfeParent (u16) - parent cellStyleXf index (0xFFFF = none)
///   2..4   iFmt (u16) - number format id
///   4..6   iFont (u16)
///   6..8   iFill (u16)
///   8..10  ixBorder (u16)
///  10      trot (u8) - text rotation
///  11      indent (u8, lower 4 bits)
///  11      bit fields: alc (3 bits), alcv (3 bits), fWrap, fJustLast, fShrinkToFit, ...
///  12      more flags: fMergeCell, readingOrder, ...
///  13      xfGrbitAtr (6 bits) - apply flags
///  14..16  unused
fn parse_xf(buf: &[u8]) -> RawXf {
    let xf_id = parser::read_u16(buf, 0);
    let num_fmt_id = parser::read_u16(buf, 2);
    let font_id = parser::read_u16(buf, 4);
    let fill_id = parser::read_u16(buf, 6);
    let border_id = parser::read_u16(buf, 8);
    let trot = buf[10];
    let indent_byte = buf[11];
    let indent = indent_byte & 0x0F;

    let align_byte = buf[12];
    let alc = align_byte & 0x07;
    let alcv = (align_byte >> 3) & 0x07;
    let f_wrap = (align_byte >> 6) & 0x01 != 0;

    let flags_byte = buf[13];
    let f_shrink = flags_byte & 0x01 != 0;
    let reading_order_val = (flags_byte >> 2) & 0x03;

    let atr = if buf.len() > 14 { buf[14] } else { 0 };
    let apply_number_format = atr & 0x01 != 0;
    let apply_font = atr & 0x02 != 0;
    let apply_alignment = atr & 0x04 != 0;
    let apply_border = atr & 0x08 != 0;
    let apply_fill = atr & 0x10 != 0;
    let apply_protection = atr & 0x20 != 0;

    let horizontal = match alc {
        0 => HorizontalAlignment::General,
        1 => HorizontalAlignment::Left,
        2 => HorizontalAlignment::Center,
        3 => HorizontalAlignment::Right,
        4 => HorizontalAlignment::Fill,
        5 => HorizontalAlignment::Justify,
        6 => HorizontalAlignment::CenterContinuous,
        7 => HorizontalAlignment::Distributed,
        _ => HorizontalAlignment::General,
    };

    let vertical = match alcv {
        0 => VerticalAlignment::Top,
        1 => VerticalAlignment::Center,
        2 => VerticalAlignment::Bottom,
        3 => VerticalAlignment::Justify,
        4 => VerticalAlignment::Distributed,
        _ => VerticalAlignment::Bottom,
    };

    let rotation = if trot == 255 {
        255i16
    } else if trot <= 90 {
        trot as i16
    } else if trot <= 180 {
        -((trot as i16) - 90)
    } else {
        0
    };

    let reading_order = match reading_order_val {
        1 => ReadingOrder::LeftToRight,
        2 => ReadingOrder::RightToLeft,
        _ => ReadingOrder::ContextDependent,
    };

    let prot_byte = if buf.len() > 15 { buf[15] } else { 0 };
    let locked = prot_byte & 0x01 != 0;
    let hidden = prot_byte & 0x02 != 0;

    RawXf {
        xf_id,
        num_fmt_id,
        font_id,
        fill_id,
        border_id,
        alignment: Alignment {
            horizontal,
            vertical,
            wrap_text: f_wrap,
            shrink_to_fit: f_shrink,
            indent,
            rotation,
            reading_order,
        },
        protection: Protection { locked, hidden },
        apply_number_format,
        apply_font,
        apply_fill,
        apply_border,
        apply_alignment,
        apply_protection,
    }
}

/// Parse BrtFont record.
///
/// Layout per [MS-XLSB] 2.4.690:
///   0..2   dyHeight (u16) - height in twips (1/20 pt)
///   2..4   grbit (u16) - flags: bit 0=bold, bit 1=italic
///   4..6   bls (u16) - font weight (400=normal, 700=bold)
///   6..8   sss (u16) - superscript/subscript (0=none, 1=super, 2=sub)
///   8      uls (u8) - underline style
///   9      bFamily (u8) - font family
///  10      bCharSet (u8) - charset
///  11      unused
///  12..16  brtColor - font color (xColorType u8, index/theme u8, tint i16, rgba 4 bytes)
///  20      bFontScheme (u8)
///  21..    name (XLWideString)
fn parse_font(buf: &[u8]) -> FontStyle {
    let mut font = FontStyle::default();
    if buf.len() < 21 {
        return font;
    }

    let dy_height = parser::read_u16(buf, 0);
    font.size = dy_height as f64 / 20.0;

    let grbit = parser::read_u16(buf, 2);
    font.italic = grbit & 0x02 != 0;
    font.strikethrough = grbit & 0x08 != 0;

    let bls = parser::read_u16(buf, 4);
    font.bold = bls >= 700;

    let sss = parser::read_u16(buf, 6);
    font.vertical_align = match sss {
        1 => duke_sheets_core::style::FontVerticalAlign::Superscript,
        2 => duke_sheets_core::style::FontVerticalAlign::Subscript,
        _ => duke_sheets_core::style::FontVerticalAlign::Baseline,
    };

    let uls = buf[8];
    font.underline = match uls {
        0 => Underline::None,
        1 => Underline::Single,
        2 => Underline::Double,
        0x21 => Underline::SingleAccounting,
        0x22 => Underline::DoubleAccounting,
        _ => Underline::Single,
    };

    let b_family = buf[9];
    if b_family != 0 {
        font.family = Some(b_family);
    }

    let b_charset = buf[10];
    if b_charset != 0 {
        font.charset = Some(b_charset);
    }

    if buf.len() >= 20 {
        font.color = parse_brt_color(buf, 12);
    }

    if buf.len() >= 21 {
        let b_font_scheme = buf[20];
        font.scheme = match b_font_scheme {
            1 => Some("major".to_string()),
            2 => Some("minor".to_string()),
            _ => None,
        };
    }

    if buf.len() > 21 {
        if let Ok((name, _)) = parser::wide_str(buf, 21) {
            if !name.is_empty() {
                font.name = name;
            }
        }
    }

    font
}

/// Parse a BrtColor structure (8 bytes starting at `off`).
///
/// Layout:
///   off+0: xColorType (u8) - 0=auto, 1=indexed, 2=rgb, 3=theme
///   off+1: index (u8) - indexed color or theme index
///   off+2..off+4: nTintAndShade (i16) - tint value
///   off+4..off+8: bRed, bGreen, bBlue, bAlpha
fn parse_brt_color(buf: &[u8], off: usize) -> Color {
    if off + 8 > buf.len() {
        return Color::Auto;
    }
    let color_type = buf[off];
    let index = buf[off + 1];
    let tint_raw = i16::from_le_bytes([buf[off + 2], buf[off + 3]]);
    let r = buf[off + 4];
    let g = buf[off + 5];
    let b = buf[off + 6];
    let a = buf[off + 7];

    match color_type {
        0 => Color::Auto,
        1 => {
            if index == 64 {
                Color::Auto
            } else {
                Color::Indexed(index)
            }
        }
        2 => {
            if a == 0xFF {
                Color::Rgb { r, g, b }
            } else {
                Color::Argb { a, r, g, b }
            }
        }
        3 => {
            let tint_i8 = if tint_raw == 0 {
                0i8
            } else {
                ((tint_raw as f64 / 32767.0) * 100.0).round() as i8
            };
            Color::Theme {
                index,
                tint: tint_i8,
            }
        }
        _ => Color::Auto,
    }
}

/// Parse BrtFill record.
///
/// Layout per [MS-XLSB] 2.4.681:
///   0..4   fls (u32) - fill pattern type
///   4..12  BrtColor foreground
///  12..20  BrtColor background
fn parse_fill(buf: &[u8]) -> FillStyle {
    if buf.len() < 4 {
        return FillStyle::None;
    }

    let fls = parser::read_u32(buf, 0);
    let pattern = match fls {
        0 => PatternType::None,
        1 => PatternType::Solid,
        2 => PatternType::MediumGray,
        3 => PatternType::DarkGray,
        4 => PatternType::LightGray,
        5 => PatternType::DarkHorizontal,
        6 => PatternType::DarkVertical,
        7 => PatternType::DarkDown,
        8 => PatternType::DarkUp,
        9 => PatternType::DarkGrid,
        10 => PatternType::DarkTrellis,
        11 => PatternType::LightHorizontal,
        12 => PatternType::LightVertical,
        13 => PatternType::LightDown,
        14 => PatternType::LightUp,
        15 => PatternType::LightGrid,
        16 => PatternType::LightTrellis,
        17 => PatternType::Gray125,
        18 => PatternType::Gray0625,
        _ => PatternType::None,
    };

    let fg = if buf.len() >= 12 {
        parse_brt_color(buf, 4)
    } else {
        Color::Auto
    };
    let bg = if buf.len() >= 20 {
        parse_brt_color(buf, 12)
    } else {
        Color::Auto
    };

    if fls == 40 && buf.len() >= 68 {
        let gradient_type = parser::read_u32(buf, 20);
        let angle = f64::from_le_bytes(buf[24..32].try_into().unwrap_or([0; 8]));
        let c_stops = parser::read_u32(buf, 64) as usize;
        let mut stops = Vec::with_capacity(c_stops);
        let mut pos = 68;
        for _ in 0..c_stops {
            if pos + 16 > buf.len() {
                break;
            }
            let position = f64::from_le_bytes(buf[pos..pos + 8].try_into().unwrap_or([0; 8]));
            let color = parse_brt_color(buf, pos + 8);
            stops.push(GradientStop { position, color });
            pos += 16;
        }
        return FillStyle::Gradient {
            gradient_type: if gradient_type == 1 {
                GradientType::Path
            } else {
                GradientType::Linear
            },
            angle,
            stops,
        };
    }

    match pattern {
        PatternType::None => FillStyle::None,
        PatternType::Solid => {
            let color = if matches!(fg, Color::Auto) { bg } else { fg };
            FillStyle::Solid { color }
        }
        PatternType::Gray125 => FillStyle::None,
        p => FillStyle::Pattern {
            pattern: p,
            foreground: fg,
            background: bg,
        },
    }
}

/// Parse BrtBorder record.
///
/// Layout per [MS-XLSB] 2.4.314:
///   0      flags (u8) - bit 0: diagonalDown, bit 1: diagonalUp
///   1..10  top edge (1 byte style + 8 bytes BrtColor)
///   10..19 bottom edge
///   19..28 left edge
///   28..37 right edge
///   37..46 diagonal edge
fn parse_border(buf: &[u8]) -> BorderStyle {
    let mut border = BorderStyle::default();
    if buf.len() < 1 {
        return border;
    }

    let flags = buf[0];
    if flags & 0x01 != 0 {
        border.diagonal_direction = duke_sheets_core::style::DiagonalDirection::Down;
    }
    if flags & 0x02 != 0 {
        if border.diagonal_direction == duke_sheets_core::style::DiagonalDirection::Down {
            border.diagonal_direction = duke_sheets_core::style::DiagonalDirection::Both;
        } else {
            border.diagonal_direction = duke_sheets_core::style::DiagonalDirection::Up;
        }
    }

    let edge_size = 10;
    let base = 1;

    if buf.len() >= base + edge_size {
        border.top = parse_border_edge(buf, base);
    }
    if buf.len() >= base + edge_size * 2 {
        border.bottom = parse_border_edge(buf, base + edge_size);
    }
    if buf.len() >= base + edge_size * 3 {
        border.left = parse_border_edge(buf, base + edge_size * 2);
    }
    if buf.len() >= base + edge_size * 4 {
        border.right = parse_border_edge(buf, base + edge_size * 3);
    }
    if buf.len() >= base + edge_size * 5 {
        border.diagonal = parse_border_edge(buf, base + edge_size * 4);
    }

    border
}

fn parse_border_edge(buf: &[u8], off: usize) -> Option<BorderEdge> {
    if off + 10 > buf.len() {
        return None;
    }
    let style_byte = buf[off];
    let line_style = match style_byte {
        0 => return None,
        1 => BorderLineStyle::Thin,
        2 => BorderLineStyle::Medium,
        3 => BorderLineStyle::Dashed,
        4 => BorderLineStyle::Dotted,
        5 => BorderLineStyle::Thick,
        6 => BorderLineStyle::Double,
        7 => BorderLineStyle::Hair,
        8 => BorderLineStyle::MediumDashed,
        9 => BorderLineStyle::DashDot,
        10 => BorderLineStyle::MediumDashDot,
        11 => BorderLineStyle::DashDotDot,
        12 => BorderLineStyle::MediumDashDotDot,
        13 => BorderLineStyle::SlantDashDot,
        _ => return None,
    };
    let color = parse_brt_color(buf, off + 2);
    Some(BorderEdge {
        style: line_style,
        color,
    })
}

/// Parse a BrtDxf record into a differential Style.
///
/// The BrtDxf payload uses a sequence of optional sub-structures, each preceded
/// by flag bits indicating which parts are present. The layout:
///   [0..4] flags (u32) - bit mask of present parts
///     bit 0: has font, bit 1: has numfmt, bit 2: has fill, bit 3: has alignment
///     bit 4: has border, bit 5: has protection
///   Then each present sub-structure follows in order.
///
/// For simplicity we parse font (bold, italic, color) and fill (solid color).
fn parse_dxf(buf: &[u8]) -> Style {
    let mut style = Style::default();
    if buf.len() < 4 {
        return style;
    }

    let flags = parser::read_u32(buf, 0);
    let mut pos = 4;

    let has_font = (flags & 0x01) != 0;
    let has_numfmt = (flags & 0x02) != 0;
    let has_fill = (flags & 0x04) != 0;
    let has_alignment = (flags & 0x08) != 0;
    let has_border = (flags & 0x10) != 0;
    let _has_protection = (flags & 0x20) != 0;

    if has_font {
        pos = parse_dxf_font(buf, pos, &mut style);
    }
    if has_numfmt {
        pos = skip_dxf_numfmt(buf, pos);
    }
    if has_fill {
        pos = parse_dxf_fill(buf, pos, &mut style);
    }
    if has_alignment {
        pos = skip_dxf_alignment(buf, pos);
    }
    if has_border {
        let _ = skip_dxf_border(buf, pos);
    }

    style
}

fn parse_dxf_font(buf: &[u8], mut pos: usize, style: &mut Style) -> usize {
    if pos >= buf.len() {
        return pos;
    }

    let font_flags = if pos + 4 <= buf.len() {
        let f = parser::read_u32(buf, pos);
        pos += 4;
        f
    } else {
        return pos;
    };

    // bit 0: bold/weight present
    if (font_flags & 0x01) != 0 {
        if pos + 2 <= buf.len() {
            let bls = parser::read_u16(buf, pos);
            style.font.bold = bls >= 700;
            pos += 2;
        }
    }
    // bit 1: italic present
    if (font_flags & 0x02) != 0 {
        if pos < buf.len() {
            style.font.italic = buf[pos] != 0;
            pos += 1;
        }
    }
    // bit 2: underline present
    if (font_flags & 0x04) != 0 {
        if pos < buf.len() {
            style.font.underline = match buf[pos] {
                0 => Underline::None,
                1 => Underline::Single,
                2 => Underline::Double,
                _ => Underline::Single,
            };
            pos += 1;
        }
    }
    // bit 3: strikethrough present
    if (font_flags & 0x08) != 0 {
        if pos < buf.len() {
            style.font.strikethrough = buf[pos] != 0;
            pos += 1;
        }
    }
    // bit 4: color present
    if (font_flags & 0x10) != 0 {
        if pos + 8 <= buf.len() {
            style.font.color = parse_brt_color(buf, pos);
            pos += 8;
        }
    }
    // bit 5: size present
    if (font_flags & 0x20) != 0 {
        if pos + 2 <= buf.len() {
            let dy_height = parser::read_u16(buf, pos);
            style.font.size = dy_height as f64 / 20.0;
            pos += 2;
        }
    }
    // bit 6: name present
    if (font_flags & 0x40) != 0 {
        if let Ok((name, consumed)) = parser::wide_str(buf, pos) {
            if !name.is_empty() {
                style.font.name = name;
            }
            pos += consumed;
        }
    }

    pos
}

fn parse_dxf_fill(buf: &[u8], mut pos: usize, style: &mut Style) -> usize {
    if pos + 4 > buf.len() {
        return pos;
    }
    let fill_flags = parser::read_u32(buf, pos);
    pos += 4;

    // bit 0: pattern type present
    let pattern_type = if (fill_flags & 0x01) != 0 {
        if pos + 4 <= buf.len() {
            let fls = parser::read_u32(buf, pos);
            pos += 4;
            fls
        } else {
            return pos;
        }
    } else {
        0
    };

    // bit 1: foreground color present
    let fg = if (fill_flags & 0x02) != 0 {
        if pos + 8 <= buf.len() {
            let c = parse_brt_color(buf, pos);
            pos += 8;
            c
        } else {
            Color::Auto
        }
    } else {
        Color::Auto
    };

    // bit 2: background color present
    let bg = if (fill_flags & 0x04) != 0 {
        if pos + 8 <= buf.len() {
            let c = parse_brt_color(buf, pos);
            pos += 8;
            c
        } else {
            Color::Auto
        }
    } else {
        Color::Auto
    };

    if pattern_type == 1 {
        let color = if matches!(fg, Color::Auto) { bg } else { fg };
        style.fill = FillStyle::Solid { color };
    } else if pattern_type > 1 {
        let pat = match pattern_type {
            2 => PatternType::MediumGray,
            3 => PatternType::DarkGray,
            4 => PatternType::LightGray,
            _ => PatternType::None,
        };
        if pat != PatternType::None {
            style.fill = FillStyle::Pattern {
                pattern: pat,
                foreground: fg,
                background: bg,
            };
        }
    }

    pos
}

fn skip_dxf_numfmt(buf: &[u8], mut pos: usize) -> usize {
    // numfmt: u16 id + XLWideString
    if pos + 2 > buf.len() {
        return pos;
    }
    pos += 2; // skip id
    if let Ok((_, consumed)) = parser::wide_str(buf, pos) {
        pos += consumed;
    }
    pos
}

fn skip_dxf_alignment(buf: &[u8], mut pos: usize) -> usize {
    // alignment: 8 bytes of alignment data
    if pos + 8 <= buf.len() {
        pos += 8;
    }
    pos
}

fn skip_dxf_border(buf: &[u8], mut pos: usize) -> usize {
    // border: 1 flags byte + 5 edges * (2 style + 8 color) = 51 bytes
    if pos + 51 <= buf.len() {
        pos += 51;
    }
    pos
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    use crate::biff12::build_record;

    fn utf16le(s: &str) -> Vec<u8> {
        s.encode_utf16().flat_map(|c| c.to_le_bytes()).collect()
    }

    fn xlwide(s: &str) -> Vec<u8> {
        let encoded = utf16le(s);
        let char_count = (encoded.len() / 2) as u32;
        let mut out = Vec::new();
        out.extend_from_slice(&char_count.to_le_bytes());
        out.extend_from_slice(&encoded);
        out
    }

    fn build_styles_bin(
        fmts: &[(u16, &str)],
        fonts: &[Vec<u8>],
        fills: &[Vec<u8>],
        borders: &[Vec<u8>],
        cell_style_xfs: &[Vec<u8>],
        cell_xfs: &[Vec<u8>],
    ) -> Vec<u8> {
        let mut data = Vec::new();

        if !fmts.is_empty() {
            data.extend_from_slice(&build_record(
                records::BRT_BEGIN_FMTS,
                &(fmts.len() as u32).to_le_bytes(),
            ));
            for (id, code) in fmts {
                let mut payload = Vec::new();
                payload.extend_from_slice(&id.to_le_bytes());
                payload.extend_from_slice(&xlwide(code));
                data.extend_from_slice(&build_record(records::BRT_FMT, &payload));
            }
            data.extend_from_slice(&build_record(records::BRT_END_FMTS, &[]));
        }

        if !fonts.is_empty() {
            data.extend_from_slice(&build_record(
                records::BRT_BEGIN_FONTS,
                &(fonts.len() as u32).to_le_bytes(),
            ));
            for font_payload in fonts {
                data.extend_from_slice(&build_record(records::BRT_FONT, font_payload));
            }
            data.extend_from_slice(&build_record(records::BRT_END_FONTS, &[]));
        }

        if !fills.is_empty() {
            data.extend_from_slice(&build_record(
                records::BRT_BEGIN_FILLS,
                &(fills.len() as u32).to_le_bytes(),
            ));
            for fill_payload in fills {
                data.extend_from_slice(&build_record(records::BRT_FILL, fill_payload));
            }
            data.extend_from_slice(&build_record(records::BRT_END_FILLS, &[]));
        }

        if !borders.is_empty() {
            data.extend_from_slice(&build_record(
                records::BRT_BEGIN_BORDERS,
                &(borders.len() as u32).to_le_bytes(),
            ));
            for border_payload in borders {
                data.extend_from_slice(&build_record(records::BRT_BORDER, border_payload));
            }
            data.extend_from_slice(&build_record(records::BRT_END_BORDERS, &[]));
        }

        if !cell_style_xfs.is_empty() {
            data.extend_from_slice(&build_record(
                records::BRT_BEGIN_CELL_STYLE_XFS,
                &(cell_style_xfs.len() as u32).to_le_bytes(),
            ));
            for xf_payload in cell_style_xfs {
                data.extend_from_slice(&build_record(records::BRT_XF, xf_payload));
            }
            data.extend_from_slice(&build_record(records::BRT_END_CELL_STYLE_XFS, &[]));
        }

        if !cell_xfs.is_empty() {
            data.extend_from_slice(&build_record(
                records::BRT_BEGIN_CELL_XFS,
                &(cell_xfs.len() as u32).to_le_bytes(),
            ));
            for xf_payload in cell_xfs {
                data.extend_from_slice(&build_record(records::BRT_XF, xf_payload));
            }
            data.extend_from_slice(&build_record(records::BRT_END_CELL_XFS, &[]));
        }

        data
    }

    fn make_zip_with_styles(styles_bin: &[u8]) -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = SimpleFileOptions::default();
            zip.start_file("xl/styles.bin", opts).unwrap();
            zip.write_all(styles_bin).unwrap();
            zip.finish().unwrap();
        }
        buf
    }

    fn make_xf_payload(
        xf_id: u16,
        num_fmt_id: u16,
        font_id: u16,
        fill_id: u16,
        border_id: u16,
        apply_bits: u8,
    ) -> Vec<u8> {
        let mut payload = vec![0u8; 16];
        payload[0..2].copy_from_slice(&xf_id.to_le_bytes());
        payload[2..4].copy_from_slice(&num_fmt_id.to_le_bytes());
        payload[4..6].copy_from_slice(&font_id.to_le_bytes());
        payload[6..8].copy_from_slice(&fill_id.to_le_bytes());
        payload[8..10].copy_from_slice(&border_id.to_le_bytes());
        payload[14] = apply_bits;
        payload
    }

    fn make_font_payload(bold: bool, size_twips: u16, name: &str) -> Vec<u8> {
        let mut payload = vec![0u8; 21];
        payload[0..2].copy_from_slice(&size_twips.to_le_bytes());
        payload[2..4].copy_from_slice(&0u16.to_le_bytes());
        let bls: u16 = if bold { 700 } else { 400 };
        payload[4..6].copy_from_slice(&bls.to_le_bytes());
        payload.extend_from_slice(&xlwide(name));
        payload
    }

    #[test]
    fn read_styles_number_format() {
        let xf = make_xf_payload(0, 164, 0, 0, 0, 0x01);
        let styles_bin = build_styles_bin(
            &[(164, "yyyy-mm-dd")],
            &[],
            &[],
            &[],
            &[make_xf_payload(0xFFFF, 0, 0, 0, 0, 0)],
            &[xf],
        );
        let zip_data = make_zip_with_styles(&styles_bin);
        let mut archive = zip::ZipArchive::new(Cursor::new(zip_data)).unwrap();
        let data = read_styles(&mut archive).unwrap();
        assert_eq!(data.styles.len(), 1);
        assert_eq!(
            data.styles[0].number_format,
            NumberFormat::Custom("yyyy-mm-dd".to_string())
        );
    }

    #[test]
    fn read_styles_builtin_number_format() {
        let xf = make_xf_payload(0, 14, 0, 0, 0, 0x01);
        let styles_bin = build_styles_bin(
            &[],
            &[],
            &[],
            &[],
            &[make_xf_payload(0xFFFF, 0, 0, 0, 0, 0)],
            &[xf],
        );
        let zip_data = make_zip_with_styles(&styles_bin);
        let mut archive = zip::ZipArchive::new(Cursor::new(zip_data)).unwrap();
        let data = read_styles(&mut archive).unwrap();
        assert_eq!(data.styles[0].number_format, NumberFormat::BuiltIn(14));
    }

    #[test]
    fn read_styles_bold_font() {
        let font = make_font_payload(true, 220, "Calibri");
        let xf = make_xf_payload(0, 0, 0, 0, 0, 0x02);
        let styles_bin = build_styles_bin(
            &[],
            &[font],
            &[],
            &[],
            &[make_xf_payload(0xFFFF, 0, 0, 0, 0, 0)],
            &[xf],
        );
        let zip_data = make_zip_with_styles(&styles_bin);
        let mut archive = zip::ZipArchive::new(Cursor::new(zip_data)).unwrap();
        let data = read_styles(&mut archive).unwrap();
        assert!(data.styles[0].font.bold);
        assert_eq!(data.styles[0].font.size, 11.0);
        assert_eq!(data.styles[0].font.name, "Calibri");
    }

    #[test]
    fn read_styles_solid_fill() {
        let mut fill_payload = vec![0u8; 20];
        fill_payload[0..4].copy_from_slice(&1u32.to_le_bytes());
        fill_payload[4] = 2;
        fill_payload[5] = 0;
        fill_payload[6..8].copy_from_slice(&0i16.to_le_bytes());
        fill_payload[8] = 255;
        fill_payload[9] = 0;
        fill_payload[10] = 0;
        fill_payload[11] = 255;

        let xf = make_xf_payload(0, 0, 0, 0, 0, 0x10);
        let styles_bin = build_styles_bin(
            &[],
            &[],
            &[fill_payload],
            &[],
            &[make_xf_payload(0xFFFF, 0, 0, 0, 0, 0)],
            &[xf],
        );
        let zip_data = make_zip_with_styles(&styles_bin);
        let mut archive = zip::ZipArchive::new(Cursor::new(zip_data)).unwrap();
        let data = read_styles(&mut archive).unwrap();
        assert_eq!(
            data.styles[0].fill,
            FillStyle::Solid {
                color: Color::Rgb { r: 255, g: 0, b: 0 }
            }
        );
    }

    #[test]
    fn read_styles_no_styles_file() {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = SimpleFileOptions::default();
            zip.start_file("xl/workbook.bin", opts).unwrap();
            zip.write_all(&[]).unwrap();
            zip.finish().unwrap();
        }
        let mut archive = zip::ZipArchive::new(Cursor::new(buf)).unwrap();
        let data = read_styles(&mut archive).unwrap();
        assert_eq!(data.styles.len(), 1);
        assert_eq!(data.styles[0], Style::default());
    }
}
