//! BIFF8 style record parsing.
//!
//! Parses FONT, FORMAT, XF, and PALETTE records from the workbook globals
//! stream and resolves them into `duke_sheets_core::Style` objects.

use std::collections::HashMap;

use duke_sheets_core::style::{
    DiagonalDirection, FontVerticalAlign, PatternType, Protection, ReadingOrder, Underline,
};
use duke_sheets_core::{
    Alignment, BorderEdge, BorderLineStyle, BorderStyle, Color, FillStyle, FontStyle,
    HorizontalAlignment, NumberFormat, Style, VerticalAlignment,
};

use crate::biff::parser::{read_u16, read_u32};
use crate::biff::strings::{read_short_string, read_unicode_string};
use crate::error::{XlsError, XlsResult};

// ============================================================================
// Default BIFF8 color palette (56 entries, indices 8–63)
// ============================================================================

/// The standard BIFF8 color palette.  Indices 8–63 in the workbook map to
/// entries 0–55 here.  A PALETTE record can override individual entries.
pub(crate) const DEFAULT_PALETTE: [(u8, u8, u8); 56] = [
    (0, 0, 0),       //  8: Black
    (255, 255, 255), //  9: White
    (255, 0, 0),     // 10: Red
    (0, 255, 0),     // 11: Bright Green
    (0, 0, 255),     // 12: Blue
    (255, 255, 0),   // 13: Yellow
    (255, 0, 255),   // 14: Pink
    (0, 255, 255),   // 15: Turquoise
    (128, 0, 0),     // 16: Dark Red
    (0, 128, 0),     // 17: Green
    (0, 0, 128),     // 18: Dark Blue
    (128, 128, 0),   // 19: Dark Yellow
    (128, 0, 128),   // 20: Violet
    (0, 128, 128),   // 21: Teal
    (192, 192, 192), // 22: Silver (25% Gray)
    (128, 128, 128), // 23: Gray (50% Gray)
    (153, 153, 255), // 24: Periwinkle
    (153, 51, 102),  // 25: Plum
    (255, 255, 204), // 26: Ivory
    (204, 255, 255), // 27: Light Turquoise
    (102, 0, 102),   // 28: Dark Purple
    (255, 128, 128), // 29: Coral
    (0, 102, 204),   // 30: Ocean Blue
    (204, 204, 255), // 31: Ice Blue
    (0, 0, 128),     // 32: Dark Blue (dup)
    (255, 0, 255),   // 33: Pink (dup)
    (255, 255, 0),   // 34: Yellow (dup)
    (0, 255, 255),   // 35: Turquoise (dup)
    (128, 0, 128),   // 36: Violet (dup)
    (128, 0, 0),     // 37: Dark Red (dup)
    (0, 128, 128),   // 38: Teal (dup)
    (0, 0, 255),     // 39: Blue (dup)
    (0, 204, 255),   // 40: Sky Blue
    (204, 255, 255), // 41: Light Turquoise (dup)
    (204, 255, 204), // 42: Light Green
    (255, 255, 153), // 43: Light Yellow
    (153, 204, 255), // 44: Pale Blue
    (255, 153, 204), // 45: Rose
    (204, 153, 255), // 46: Lavender
    (255, 204, 153), // 47: Tan
    (51, 102, 255),  // 48: Light Blue
    (51, 204, 204),  // 49: Aqua
    (153, 204, 0),   // 50: Lime
    (255, 204, 0),   // 51: Gold
    (255, 153, 0),   // 52: Light Orange
    (255, 102, 0),   // 53: Orange
    (102, 102, 153), // 54: Blue-Gray
    (150, 150, 150), // 55: 40% Gray
    (0, 51, 102),    // 56: Dark Teal
    (51, 153, 102),  // 57: Sea Green
    (0, 51, 0),      // 58: Dark Green
    (51, 51, 0),     // 59: Olive Green
    (153, 51, 0),    // 60: Brown
    (153, 51, 51),   // 61: Dark Rose
    (51, 51, 153),   // 62: Indigo
    (51, 51, 51),    // 63: 80% Gray
];

// ============================================================================
// Intermediate BIFF types
// ============================================================================

/// Parsed FONT record data.
#[derive(Debug, Clone)]
pub(crate) struct BiffFont {
    /// Font height in twips (1/20 of a point).
    pub height_twips: u16,
    pub bold: bool,
    pub italic: bool,
    pub underline: u8,
    pub strikethrough: bool,
    /// Palette color index for the font.
    pub color_index: u16,
    /// 0 = baseline, 1 = superscript, 2 = subscript.
    pub superscript: u16,
    pub name: String,
}

/// Parsed XF record data (20 bytes in BIFF8).
#[derive(Debug, Clone)]
pub(crate) struct BiffXf {
    pub font_index: u16,
    pub format_index: u16,
    pub locked: bool,
    pub hidden: bool,
    #[allow(dead_code)]
    pub is_style_xf: bool,
    // Alignment
    pub hor_align: u8,
    pub vert_align: u8,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
    pub indent: u8,
    pub rotation: u8,
    pub reading_order: u8,
    // Borders — line style codes (0–13)
    pub border_left: u8,
    pub border_right: u8,
    pub border_top: u8,
    pub border_bottom: u8,
    pub border_diag: u8,
    // Border color indices
    pub icv_left: u16,
    pub icv_right: u16,
    pub icv_top: u16,
    pub icv_bottom: u16,
    pub icv_diag: u16,
    pub diagonal_dir: u8,
    // Fill
    pub fill_pattern: u8,
    pub icv_fore: u16,
    pub icv_back: u16,
}

/// All style data collected from the workbook globals stream.
pub(crate) struct StyleContext {
    pub fonts: Vec<BiffFont>,
    pub formats: HashMap<u16, String>,
    pub xfs: Vec<BiffXf>,
    pub palette: [(u8, u8, u8); 56],
}

impl StyleContext {
    pub fn new() -> Self {
        Self {
            fonts: Vec::new(),
            formats: HashMap::new(),
            xfs: Vec::new(),
            palette: DEFAULT_PALETTE,
        }
    }

    /// Build the resolved style table (one `Style` per XF record).
    pub fn build_style_table(&self) -> Vec<Style> {
        self.xfs.iter().map(|xf| self.resolve_xf(xf)).collect()
    }

    /// Resolve a single XF record into a core `Style`.
    fn resolve_xf(&self, xf: &BiffXf) -> Style {
        Style {
            font: self.resolve_font(xf.font_index),
            fill: self.resolve_fill(xf),
            border: self.resolve_border(xf),
            alignment: self.resolve_alignment(xf),
            number_format: self.resolve_number_format(xf.format_index),
            protection: Protection {
                locked: xf.locked,
                hidden: xf.hidden,
            },
        }
    }

    // ── Font resolution ─────────────────────────────────────────────────

    fn resolve_font(&self, font_index: u16) -> FontStyle {
        // BIFF8 quirk: font index 4 is skipped in the file.
        // Indices 0–3 map directly; index 5 → fonts[4], index 6 → fonts[5], etc.
        let actual = if font_index >= 5 {
            (font_index - 1) as usize
        } else {
            font_index as usize
        };

        let bf = match self.fonts.get(actual) {
            Some(f) => f,
            None => return FontStyle::default(),
        };

        FontStyle {
            name: bf.name.clone(),
            size: bf.height_twips as f64 / 20.0,
            bold: bf.bold,
            italic: bf.italic,
            underline: match bf.underline {
                0x01 => Underline::Single,
                0x02 => Underline::Double,
                0x21 => Underline::SingleAccounting,
                0x22 => Underline::DoubleAccounting,
                _ => Underline::None,
            },
            strikethrough: bf.strikethrough,
            color: self.resolve_color(bf.color_index),
            vertical_align: match bf.superscript {
                1 => FontVerticalAlign::Superscript,
                2 => FontVerticalAlign::Subscript,
                _ => FontVerticalAlign::Baseline,
            },
        }
    }

    // ── Fill resolution ─────────────────────────────────────────────────

    fn resolve_fill(&self, xf: &BiffXf) -> FillStyle {
        let pat = pattern_from_biff(xf.fill_pattern);

        match pat {
            PatternType::None => FillStyle::None,
            PatternType::Solid => {
                // Solid fill: foreground color is the fill color.
                let color = self.resolve_color(xf.icv_fore);
                if color.is_auto() {
                    FillStyle::None
                } else {
                    FillStyle::Solid { color }
                }
            }
            _ => {
                let fg = self.resolve_color(xf.icv_fore);
                let bg = self.resolve_color(xf.icv_back);
                FillStyle::Pattern {
                    pattern: pat,
                    foreground: fg,
                    background: bg,
                }
            }
        }
    }

    // ── Border resolution ───────────────────────────────────────────────

    fn resolve_border(&self, xf: &BiffXf) -> BorderStyle {
        let make_edge = |line_code: u8, icv: u16| -> Option<BorderEdge> {
            let ls = border_line_from_biff(line_code);
            if matches!(ls, BorderLineStyle::None) {
                None
            } else {
                Some(BorderEdge::new(ls, self.resolve_color(icv)))
            }
        };

        let diag_dir = match xf.diagonal_dir {
            1 => DiagonalDirection::Down,
            2 => DiagonalDirection::Up,
            3 => DiagonalDirection::Both,
            _ => DiagonalDirection::None,
        };

        BorderStyle {
            left: make_edge(xf.border_left, xf.icv_left),
            right: make_edge(xf.border_right, xf.icv_right),
            top: make_edge(xf.border_top, xf.icv_top),
            bottom: make_edge(xf.border_bottom, xf.icv_bottom),
            diagonal: make_edge(xf.border_diag, xf.icv_diag),
            diagonal_direction: diag_dir,
        }
    }

    // ── Alignment resolution ────────────────────────────────────────────

    fn resolve_alignment(&self, xf: &BiffXf) -> Alignment {
        let horizontal = match xf.hor_align {
            1 => HorizontalAlignment::Left,
            2 => HorizontalAlignment::Center,
            3 => HorizontalAlignment::Right,
            4 => HorizontalAlignment::Fill,
            5 => HorizontalAlignment::Justify,
            6 => HorizontalAlignment::CenterContinuous,
            7 => HorizontalAlignment::Distributed,
            _ => HorizontalAlignment::General,
        };

        let vertical = match xf.vert_align {
            0 => VerticalAlignment::Top,
            1 => VerticalAlignment::Center,
            2 => VerticalAlignment::Bottom,
            3 => VerticalAlignment::Justify,
            4 => VerticalAlignment::Distributed,
            _ => VerticalAlignment::Bottom,
        };

        // BIFF rotation: 0 = none, 1–90 = CCW degrees, 91–180 = CW as -(val-90),
        // 255 = vertical text.
        let rotation = match xf.rotation {
            0 => 0i16,
            r @ 1..=90 => r as i16,
            r @ 91..=180 => -((r as i16) - 90),
            255 => 255,
            _ => 0,
        };

        let reading_order = match xf.reading_order {
            1 => ReadingOrder::LeftToRight,
            2 => ReadingOrder::RightToLeft,
            _ => ReadingOrder::ContextDependent,
        };

        Alignment {
            horizontal,
            vertical,
            wrap_text: xf.wrap_text,
            shrink_to_fit: xf.shrink_to_fit,
            indent: xf.indent,
            rotation,
            reading_order,
        }
    }

    // ── Number format resolution ────────────────────────────────────────

    fn resolve_number_format(&self, fmt_id: u16) -> NumberFormat {
        if fmt_id == 0 {
            return NumberFormat::General;
        }
        // Custom format string from FORMAT records?
        if let Some(code) = self.formats.get(&fmt_id) {
            return NumberFormat::Custom(code.clone());
        }
        // Built-in format ID (1–49 are well-known).
        NumberFormat::BuiltIn(fmt_id as u32)
    }

    // ── Color resolution ────────────────────────────────────────────────

    pub(crate) fn resolve_color(&self, icv: u16) -> Color {
        match icv {
            8..=63 => {
                let idx = (icv - 8) as usize;
                let (r, g, b) = self.palette[idx];
                Color::Rgb { r, g, b }
            }
            // 0x0040 = default foreground (system window text → black)
            0x0040 => Color::Rgb { r: 0, g: 0, b: 0 },
            // 0x0041 = default background (system window → white)
            0x0041 => Color::Rgb {
                r: 255,
                g: 255,
                b: 255,
            },
            // 0x7FFF = automatic
            0x7FFF => Color::Auto,
            // Indices 0–7: EGA colors (rarely referenced directly in BIFF8,
            // but some writers use them). Map to the same values as 8–15.
            0..=7 => {
                let ega: [(u8, u8, u8); 8] = [
                    (0, 0, 0),
                    (255, 255, 255),
                    (255, 0, 0),
                    (0, 255, 0),
                    (0, 0, 255),
                    (255, 255, 0),
                    (255, 0, 255),
                    (0, 255, 255),
                ];
                let (r, g, b) = ega[icv as usize];
                Color::Rgb { r, g, b }
            }
            _ => Color::Auto,
        }
    }
}

// ============================================================================
// Record parsers
// ============================================================================

/// Parse a FONT record (0x0031).
///
/// Layout:
///   0  u16  dyHeight   — font height in twips (1/20 pt)
///   2  u16  grbit      — flags (bit 1 = italic, bit 3 = strikethrough)
///   4  u16  icv        — color index
///   6  u16  bls        — bold weight (400 = normal, 700 = bold)
///   8  u16  sss        — super/subscript (0/1/2)
///  10  u8   uls        — underline type
///  11  u8   bFamily    — font family (ignored)
///  12  u8   bCharSet   — character set (ignored)
///  13  u8   reserved
///  14  ...  font name  — short string (1-byte length prefix)
pub(crate) fn parse_font(data: &[u8]) -> XlsResult<BiffFont> {
    if data.len() < 15 {
        return Err(XlsError::Parse("FONT record too short".into()));
    }

    let mut off = 0;
    let height = read_u16(data, &mut off)?;
    let grbit = read_u16(data, &mut off)?;
    let icv = read_u16(data, &mut off)?;
    let bls = read_u16(data, &mut off)?;
    let sss = read_u16(data, &mut off)?;
    let uls = data[off];
    off += 1;
    let _family = data[off];
    off += 1;
    let _charset = data[off];
    off += 1;
    let _reserved = data[off];
    off += 1;

    let name = if off < data.len() {
        read_short_string(data, &mut off).unwrap_or_default()
    } else {
        String::new()
    };

    Ok(BiffFont {
        height_twips: height,
        italic: (grbit & 0x0002) != 0,
        strikethrough: (grbit & 0x0008) != 0,
        bold: bls >= 700,
        underline: uls,
        color_index: icv,
        superscript: sss,
        name,
    })
}

/// Parse a FORMAT record (0x041E).
///
/// Layout:
///   0  u16  ifmt   — format index
///   2  ...  format string (unicode string, 2-byte length prefix)
pub(crate) fn parse_format(data: &[u8]) -> XlsResult<(u16, String)> {
    let mut off = 0;
    let ifmt = read_u16(data, &mut off)?;
    let s = read_unicode_string(data, &mut off)?;
    Ok((ifmt, s))
}

/// Parse an XF record (0x00E0, always 20 bytes in BIFF8).
///
/// Layout (see [MS-XLS] §2.4.353):
///   0   u16  ifnt          — font index
///   2   u16  ifmt          — format index
///   4   u16  type/protect  — bits 0-1 lock/hidden, bit 2 style-xf
///   6   u8   alignment1    — bits 0-2 halign, bit 3 wrap, bits 4-6 valign
///   7   u8   trot          — text rotation
///   8   u8   alignment2    — bits 0-3 indent, bit 4 shrink, bits 6-7 reading order
///   9   u8   used_attribs  — (ignored)
///  10   u32  border lines/colors 1
///  14   u32  border lines/colors 2 + fill pattern
///  18   u16  fill colors
pub(crate) fn parse_xf(data: &[u8]) -> XlsResult<BiffXf> {
    if data.len() < 20 {
        return Err(XlsError::Parse(format!(
            "XF record too short: {} bytes (expected 20)",
            data.len()
        )));
    }

    let mut off = 0;
    let ifnt = read_u16(data, &mut off)?;
    let ifmt = read_u16(data, &mut off)?;
    let type_prot = read_u16(data, &mut off)?;

    let locked = (type_prot & 0x0001) != 0;
    let hidden = (type_prot & 0x0002) != 0;
    let is_style_xf = (type_prot & 0x0004) != 0;

    // Byte 6: alignment byte 1
    let align1 = data[off];
    off += 1;
    let hor_align = align1 & 0x07;
    let wrap_text = (align1 & 0x08) != 0;
    let vert_align = (align1 >> 4) & 0x07;

    // Byte 7: rotation
    let rotation = data[off];
    off += 1;

    // Byte 8: alignment byte 2
    let align2 = data[off];
    off += 1;
    let indent = align2 & 0x0F;
    let shrink_to_fit = (align2 & 0x10) != 0;
    let reading_order = (align2 >> 6) & 0x03;

    // Byte 9: used attributes (skip)
    off += 1;

    // Bytes 10–13: border & color block 1
    let border1 = read_u32(data, &mut off)?;
    let border_left = (border1 & 0x0F) as u8;
    let border_right = ((border1 >> 4) & 0x0F) as u8;
    let border_top = ((border1 >> 8) & 0x0F) as u8;
    let border_bottom = ((border1 >> 12) & 0x0F) as u8;
    let icv_left = ((border1 >> 16) & 0x7F) as u16;
    let icv_right = ((border1 >> 23) & 0x7F) as u16;
    let diagonal_dir = ((border1 >> 30) & 0x03) as u8;

    // Bytes 14–17: border & color block 2 + fill pattern
    let border2 = read_u32(data, &mut off)?;
    let icv_top = (border2 & 0x7F) as u16;
    let icv_bottom = ((border2 >> 7) & 0x7F) as u16;
    let icv_diag = ((border2 >> 14) & 0x7F) as u16;
    let border_diag = ((border2 >> 21) & 0x0F) as u8;
    let fill_pattern = ((border2 >> 26) & 0x3F) as u8;

    // Bytes 18–19: fill colors
    let fill_colors = read_u16(data, &mut off)?;
    let icv_fore = fill_colors & 0x7F;
    let icv_back = (fill_colors >> 7) & 0x7F;

    Ok(BiffXf {
        font_index: ifnt,
        format_index: ifmt,
        locked,
        hidden,
        is_style_xf,
        hor_align,
        vert_align,
        wrap_text,
        shrink_to_fit,
        indent,
        rotation,
        reading_order,
        border_left,
        border_right,
        border_top,
        border_bottom,
        border_diag,
        icv_left,
        icv_right,
        icv_top,
        icv_bottom,
        icv_diag,
        diagonal_dir,
        fill_pattern,
        icv_fore,
        icv_back,
    })
}

/// Apply a PALETTE record to the style context.
///
/// Layout:
///   0  u16  ccv    — number of colors (typically 56)
///   2  ...  colors — array of ccv × 4-byte entries (R, G, B, 0x00)
pub(crate) fn apply_palette(data: &[u8], palette: &mut [(u8, u8, u8); 56]) -> XlsResult<()> {
    if data.len() < 2 {
        return Err(XlsError::Parse("PALETTE record too short".into()));
    }

    let mut off = 0;
    let count = read_u16(data, &mut off)? as usize;
    let max = count.min(56);

    for i in 0..max {
        if off + 4 > data.len() {
            break;
        }
        palette[i] = (data[off], data[off + 1], data[off + 2]);
        off += 4; // skip the 4th byte (always 0x00)
    }

    Ok(())
}

// ============================================================================
// Mapping helpers
// ============================================================================

/// Map a BIFF border line code (0–13) to a `BorderLineStyle`.
fn border_line_from_biff(code: u8) -> BorderLineStyle {
    match code {
        0 => BorderLineStyle::None,
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
        _ => BorderLineStyle::None,
    }
}

/// Map a BIFF fill pattern code (0–18) to a `PatternType`.
fn pattern_from_biff(code: u8) -> PatternType {
    match code {
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
    }
}

// ============================================================================
// Unit tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_color_palette() {
        let ctx = StyleContext::new();
        // Index 8 → palette[0] = (0,0,0) = black
        assert_eq!(ctx.resolve_color(8), Color::Rgb { r: 0, g: 0, b: 0 });
        // Index 10 → palette[2] = (255,0,0) = red
        assert_eq!(ctx.resolve_color(10), Color::Rgb { r: 255, g: 0, b: 0 });
        // Index 63 → palette[55]
        assert_eq!(
            ctx.resolve_color(63),
            Color::Rgb {
                r: 51,
                g: 51,
                b: 51
            }
        );
    }

    #[test]
    fn test_resolve_color_special() {
        let ctx = StyleContext::new();
        assert_eq!(ctx.resolve_color(0x0040), Color::Rgb { r: 0, g: 0, b: 0 });
        assert_eq!(
            ctx.resolve_color(0x0041),
            Color::Rgb {
                r: 255,
                g: 255,
                b: 255
            }
        );
        assert_eq!(ctx.resolve_color(0x7FFF), Color::Auto);
    }

    #[test]
    fn test_resolve_color_ega() {
        let ctx = StyleContext::new();
        assert_eq!(ctx.resolve_color(0), Color::Rgb { r: 0, g: 0, b: 0 });
        assert_eq!(
            ctx.resolve_color(1),
            Color::Rgb {
                r: 255,
                g: 255,
                b: 255
            }
        );
        assert_eq!(ctx.resolve_color(2), Color::Rgb { r: 255, g: 0, b: 0 });
    }

    #[test]
    fn test_parse_font_basic() {
        // Minimal FONT record: height=220 (11pt), flags=0, color=auto(0x7FFF),
        // bls=400 (normal), sss=0, uls=0, family=0, charset=0, reserved=0,
        // font name "Arial" (short string: len=5, flags=0, "Arial")
        let mut data = Vec::new();
        data.extend_from_slice(&220u16.to_le_bytes()); // height = 220 twips = 11pt
        data.extend_from_slice(&0u16.to_le_bytes()); // grbit = 0
        data.extend_from_slice(&0x7FFFu16.to_le_bytes()); // icv = auto
        data.extend_from_slice(&400u16.to_le_bytes()); // bls = normal
        data.extend_from_slice(&0u16.to_le_bytes()); // sss = none
        data.push(0x00); // uls = none
        data.push(0); // family
        data.push(0); // charset
        data.push(0); // reserved
                      // Short string: length=5, flags=0 (compressed), "Arial"
        data.push(5);
        data.push(0x00); // compressed (Latin-1)
        data.extend_from_slice(b"Arial");

        let font = parse_font(&data).unwrap();
        assert_eq!(font.height_twips, 220);
        assert!(!font.bold);
        assert!(!font.italic);
        assert!(!font.strikethrough);
        assert_eq!(font.underline, 0);
        assert_eq!(font.color_index, 0x7FFF);
        assert_eq!(font.superscript, 0);
        assert_eq!(font.name, "Arial");
    }

    #[test]
    fn test_parse_font_bold_italic() {
        let mut data = Vec::new();
        data.extend_from_slice(&240u16.to_le_bytes()); // 12pt
        data.extend_from_slice(&0x0002u16.to_le_bytes()); // italic
        data.extend_from_slice(&10u16.to_le_bytes()); // red
        data.extend_from_slice(&700u16.to_le_bytes()); // bold
        data.extend_from_slice(&1u16.to_le_bytes()); // superscript
        data.push(0x01); // single underline
        data.push(0);
        data.push(0);
        data.push(0);
        data.push(0); // empty name
        data.push(0x00);

        let font = parse_font(&data).unwrap();
        assert!(font.bold);
        assert!(font.italic);
        assert_eq!(font.underline, 0x01);
        assert_eq!(font.superscript, 1);
        assert_eq!(font.color_index, 10);
        assert_eq!(font.height_twips, 240);
    }

    #[test]
    fn test_parse_xf_basic() {
        // 20-byte XF record: all zeros = default style XF
        let mut data = [0u8; 20];
        // font_index = 0, format_index = 0, type_prot = 0x0004 (style XF, unlocked)
        data[4] = 0x04;
        data[5] = 0x00;

        let xf = parse_xf(&data).unwrap();
        assert_eq!(xf.font_index, 0);
        assert_eq!(xf.format_index, 0);
        assert!(!xf.locked);
        assert!(!xf.hidden);
        assert!(xf.is_style_xf);
        assert_eq!(xf.hor_align, 0);
        assert_eq!(xf.vert_align, 0);
        assert!(!xf.wrap_text);
        assert_eq!(xf.fill_pattern, 0);
    }

    #[test]
    fn test_border_line_mapping() {
        assert_eq!(border_line_from_biff(0), BorderLineStyle::None);
        assert_eq!(border_line_from_biff(1), BorderLineStyle::Thin);
        assert_eq!(border_line_from_biff(2), BorderLineStyle::Medium);
        assert_eq!(border_line_from_biff(5), BorderLineStyle::Thick);
        assert_eq!(border_line_from_biff(6), BorderLineStyle::Double);
        assert_eq!(border_line_from_biff(7), BorderLineStyle::Hair);
        assert_eq!(border_line_from_biff(13), BorderLineStyle::SlantDashDot);
        assert_eq!(border_line_from_biff(99), BorderLineStyle::None);
    }

    #[test]
    fn test_pattern_mapping() {
        assert_eq!(pattern_from_biff(0), PatternType::None);
        assert_eq!(pattern_from_biff(1), PatternType::Solid);
        assert_eq!(pattern_from_biff(2), PatternType::MediumGray);
        assert_eq!(pattern_from_biff(18), PatternType::Gray0625);
        assert_eq!(pattern_from_biff(255), PatternType::None);
    }

    #[test]
    fn test_rotation_mapping() {
        let ctx = StyleContext::new();
        // Helper to test rotation through a minimal XF
        let make_xf = |rot: u8| BiffXf {
            font_index: 0,
            format_index: 0,
            locked: false,
            hidden: false,
            is_style_xf: false,
            hor_align: 0,
            vert_align: 0,
            wrap_text: false,
            shrink_to_fit: false,
            indent: 0,
            rotation: rot,
            reading_order: 0,
            border_left: 0,
            border_right: 0,
            border_top: 0,
            border_bottom: 0,
            border_diag: 0,
            icv_left: 0,
            icv_right: 0,
            icv_top: 0,
            icv_bottom: 0,
            icv_diag: 0,
            diagonal_dir: 0,
            fill_pattern: 0,
            icv_fore: 0,
            icv_back: 0,
        };

        assert_eq!(ctx.resolve_alignment(&make_xf(0)).rotation, 0);
        assert_eq!(ctx.resolve_alignment(&make_xf(45)).rotation, 45);
        assert_eq!(ctx.resolve_alignment(&make_xf(90)).rotation, 90);
        assert_eq!(ctx.resolve_alignment(&make_xf(91)).rotation, -1);
        assert_eq!(ctx.resolve_alignment(&make_xf(180)).rotation, -90);
        assert_eq!(ctx.resolve_alignment(&make_xf(255)).rotation, 255);
    }

    #[test]
    fn test_font_index_4_skipped() {
        let mut ctx = StyleContext::new();
        // Add 5 fonts (indices 0,1,2,3 then 5 → stored as [0..4])
        for i in 0..5 {
            ctx.fonts.push(BiffFont {
                height_twips: 200 + i * 20,
                bold: false,
                italic: false,
                underline: 0,
                strikethrough: false,
                color_index: 0x7FFF,
                superscript: 0,
                name: format!("Font{}", i),
            });
        }
        // Font index 0 → fonts[0] = "Font0"
        assert_eq!(ctx.resolve_font(0).name, "Font0");
        // Font index 3 → fonts[3] = "Font3"
        assert_eq!(ctx.resolve_font(3).name, "Font3");
        // Font index 5 → fonts[4] = "Font4" (skip index 4)
        assert_eq!(ctx.resolve_font(5).name, "Font4");
        // Font index 6 → fonts[5] = out of bounds → default
        assert_eq!(ctx.resolve_font(6).name, "Calibri"); // FontStyle::default()
    }

    #[test]
    fn test_number_format_resolution() {
        let mut ctx = StyleContext::new();
        ctx.formats.insert(164, "yyyy-mm-dd".into());
        ctx.formats.insert(165, "0.00%".into());

        assert_eq!(ctx.resolve_number_format(0), NumberFormat::General);
        assert_eq!(ctx.resolve_number_format(14), NumberFormat::BuiltIn(14));
        assert_eq!(
            ctx.resolve_number_format(164),
            NumberFormat::Custom("yyyy-mm-dd".into())
        );
        assert_eq!(
            ctx.resolve_number_format(165),
            NumberFormat::Custom("0.00%".into())
        );
    }

    #[test]
    fn test_apply_palette() {
        let mut palette = DEFAULT_PALETTE;
        // Fake PALETTE record overriding 2 entries
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_le_bytes()); // 2 colors
        data.extend_from_slice(&[0xAA, 0xBB, 0xCC, 0x00]); // entry 0
        data.extend_from_slice(&[0x11, 0x22, 0x33, 0x00]); // entry 1

        apply_palette(&data, &mut palette).unwrap();
        assert_eq!(palette[0], (0xAA, 0xBB, 0xCC));
        assert_eq!(palette[1], (0x11, 0x22, 0x33));
        // Entry 2 should still be the default
        assert_eq!(palette[2], DEFAULT_PALETTE[2]);
    }
}
