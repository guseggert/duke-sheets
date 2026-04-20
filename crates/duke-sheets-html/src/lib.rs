//! HTML export for duke-sheets workbooks.
//!
//! Converts worksheets into HTML tables, preserving cell styles (fonts, fills,
//! borders, alignment), merged cells, column widths, and row heights.
//!
//! # Example
//!
//! ```rust,no_run
//! use duke_sheets_html::{worksheet_to_html, HtmlOptions};
//! use duke_sheets_core::Worksheet;
//!
//! let sheet = Worksheet::new("Demo");
//! let html = worksheet_to_html(&sheet, &HtmlOptions::default());
//! println!("{}", html);
//! ```

use std::collections::HashMap;
use std::fmt::Write;

use duke_sheets_core::style::{
    Alignment, BorderEdge, BorderLineStyle, FillStyle, FontStyle, FontVerticalAlign,
    HorizontalAlignment, Underline, VerticalAlignment,
};
use duke_sheets_core::{CellRange, CellValue, Style, Worksheet};

// Public API

/// Options controlling HTML generation.
#[derive(Debug, Clone)]
pub struct HtmlOptions {
    /// Wrap the `<table>` in a full `<!DOCTYPE html>` document.
    /// When `false`, only the bare `<table>` element is emitted.
    pub full_document: bool,

    /// Document `<title>` (defaults to the sheet name).
    pub title: Option<String>,

    /// Use Excel number-formatted values (dates, percentages, currencies ...).
    /// When `false`, raw values are emitted (numbers as plain floats, dates as
    /// serial numbers).
    pub formatted: bool,
}

impl Default for HtmlOptions {
    fn default() -> Self {
        Self {
            full_document: true,
            title: None,
            formatted: true,
        }
    }
}

/// Convert a single worksheet to an HTML string.
pub fn worksheet_to_html(sheet: &Worksheet, options: &HtmlOptions) -> String {
    let mut buf = String::with_capacity(4096);

    if options.full_document {
        write_document_open(&mut buf, options, sheet.name());
    }

    write_table(&mut buf, sheet, options);

    if options.full_document {
        buf.push_str("</body>\n</html>\n");
    }

    buf
}

// Document wrapper

fn write_document_open(buf: &mut String, options: &HtmlOptions, sheet_name: &str) {
    buf.push_str("<!DOCTYPE html>\n<html>\n<head>\n");
    buf.push_str("<meta charset=\"utf-8\">\n");

    let title = options.title.as_deref().unwrap_or(sheet_name);
    let _ = writeln!(buf, "<title>{}</title>", escape_html(title));

    buf.push_str("<style>\n");
    buf.push_str("table { border-collapse: collapse; }\n");
    buf.push_str("td { padding: 1px 4px; vertical-align: bottom; }\n");
    buf.push_str("</style>\n");
    buf.push_str("</head>\n<body>\n");
}

// Table generation

fn write_table(buf: &mut String, sheet: &Worksheet, options: &HtmlOptions) {
    let used_range = match sheet.used_range() {
        Some(r) => r,
        None => {
            buf.push_str("<table></table>\n");
            return;
        }
    };

    let merge_map = build_merge_map(sheet.merged_regions(), &used_range);

    buf.push_str("<table>\n");

    // Column widths via <colgroup>
    write_colgroup(buf, sheet, used_range.end.col);

    // Rows
    for row in 0..=used_range.end.row {
        if sheet.is_row_hidden(row) {
            continue;
        }

        let height = sheet.row_height(row);
        let _ = writeln!(buf, "<tr style=\"height:{:.0}px\">", pt_to_px(height));

        for col in 0..=used_range.end.col {
            if sheet.is_column_hidden(col) {
                continue;
            }

            match merge_map.get(&(row, col)) {
                Some(MergeInfo::Interior) => continue,
                Some(MergeInfo::Origin { colspan, rowspan }) => {
                    write_cell(buf, sheet, row, col, Some((*colspan, *rowspan)), options);
                }
                None => {
                    write_cell(buf, sheet, row, col, None, options);
                }
            }
        }

        buf.push_str("</tr>\n");
    }

    buf.push_str("</table>\n");
}

fn write_colgroup(buf: &mut String, sheet: &Worksheet, max_col: u16) {
    buf.push_str("<colgroup>\n");
    for col in 0..=max_col {
        if sheet.is_column_hidden(col) {
            buf.push_str("<col style=\"display:none\">\n");
        } else {
            let width_px = col_width_to_px(sheet.column_width(col));
            let _ = writeln!(buf, "<col style=\"width:{width_px}px\">");
        }
    }
    buf.push_str("</colgroup>\n");
}

// Cell rendering

fn write_cell(
    buf: &mut String,
    sheet: &Worksheet,
    row: u32,
    col: u16,
    span: Option<(u16, u32)>,
    options: &HtmlOptions,
) {
    buf.push_str("<td");

    // colspan / rowspan
    if let Some((cs, rs)) = span {
        if cs > 1 {
            let _ = write!(buf, " colspan=\"{cs}\"");
        }
        if rs > 1 {
            let _ = write!(buf, " rowspan=\"{rs}\"");
        }
    }

    // Inline styles from the cell style
    let style = sheet.cell_style_at(row, col);
    let css = style.map(style_to_css).unwrap_or_default();
    if !css.is_empty() {
        let _ = write!(buf, " style=\"{css}\"");
    }

    buf.push('>');

    // Cell content
    let text = cell_display_text(sheet, row, col, options);
    if !text.is_empty() {
        // Check for rich text - if the raw value is RichText, render with spans
        let raw = sheet.get_value_at(row, col);
        if let CellValue::RichText(runs) = &raw {
            write_rich_text(buf, runs);
        } else {
            buf.push_str(&escape_html(&text));
        }
    }

    buf.push_str("</td>\n");
}

fn cell_display_text(sheet: &Worksheet, row: u32, col: u16, options: &HtmlOptions) -> String {
    let value = sheet.get_value_at(row, col);

    if options.formatted {
        return sheet.formatted_value_at(row, col);
    }

    cell_value_to_string(&value)
}

fn cell_value_to_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{}", *n as i64)
            } else {
                format!("{n}")
            }
        }
        CellValue::String(s) => s.to_string(),
        CellValue::RichText(runs) => duke_sheets_core::rich_text_to_plain(runs),
        CellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::Error(e) => e.to_string(),
        CellValue::SpillTarget { .. } => String::new(),
    }
}

fn write_rich_text(buf: &mut String, runs: &[duke_sheets_core::RichTextRun]) {
    for run in runs {
        match &run.font {
            Some(font) if !font.is_empty() => {
                let css = run_font_to_css(font);
                if css.is_empty() {
                    buf.push_str(&escape_html(&run.text));
                } else {
                    let _ = write!(
                        buf,
                        "<span style=\"{css}\">{}</span>",
                        escape_html(&run.text)
                    );
                }
            }
            _ => buf.push_str(&escape_html(&run.text)),
        }
    }
}

fn run_font_to_css(font: &duke_sheets_core::RunFont) -> String {
    let mut css = String::new();

    if let Some(true) = font.bold {
        css.push_str("font-weight:bold;");
    }
    if let Some(true) = font.italic {
        css.push_str("font-style:italic;");
    }

    // text-decoration: combine underline + strikethrough
    let ul = matches!(
        font.underline,
        Some(Underline::Single | Underline::SingleAccounting)
    );
    let st = matches!(font.strikethrough, Some(true));
    push_text_decoration(&mut css, ul, st);

    if let Some(size) = font.size {
        let _ = write!(css, "font-size:{size}pt;");
    }
    if let Some(ref name) = font.name {
        let _ = write!(css, "font-family:{};", escape_html(name));
    }
    if let Some(color) = &font.color {
        if !color.is_auto() && color.to_rgb() != (0, 0, 0) {
            let _ = write!(css, "color:#{};", css_hex(color));
        }
    }
    if let Some(FontVerticalAlign::Superscript) = font.vertical_align {
        css.push_str("vertical-align:super;font-size:smaller;");
    }
    if let Some(FontVerticalAlign::Subscript) = font.vertical_align {
        css.push_str("vertical-align:sub;font-size:smaller;");
    }

    css
}

// Style → CSS conversion

fn style_to_css(style: &Style) -> String {
    let mut css = String::new();

    // Font
    font_to_css(&mut css, &style.font);

    // Fill
    fill_to_css(&mut css, &style.fill);

    // Borders
    borders_to_css(&mut css, &style.border);

    // Alignment
    alignment_to_css(&mut css, &style.alignment);

    css
}

fn font_to_css(css: &mut String, font: &FontStyle) {
    if font.bold {
        css.push_str("font-weight:bold;");
    }
    if font.italic {
        css.push_str("font-style:italic;");
    }

    // text-decoration: combine underline + strikethrough
    let has_underline = !matches!(font.underline, Underline::None);
    push_text_decoration(css, has_underline, font.strikethrough);

    if (font.size - 11.0).abs() > 0.01 {
        let _ = write!(css, "font-size:{}pt;", font.size);
    }
    if font.name != "Calibri" {
        let _ = write!(css, "font-family:{};", escape_html(&font.name));
    }
    if !font.color.is_auto() && font.color.to_rgb() != (0, 0, 0) {
        let _ = write!(css, "color:#{};", css_hex(&font.color));
    }
    if matches!(font.vertical_align, FontVerticalAlign::Superscript) {
        css.push_str("vertical-align:super;font-size:smaller;");
    }
    if matches!(font.vertical_align, FontVerticalAlign::Subscript) {
        css.push_str("vertical-align:sub;font-size:smaller;");
    }
}

fn fill_to_css(css: &mut String, fill: &FillStyle) {
    match fill {
        FillStyle::Solid { color } if !color.is_auto() => {
            let _ = write!(css, "background-color:#{};", css_hex(color));
        }
        FillStyle::Pattern {
            foreground,
            background,
            ..
        } => {
            // Use foreground as background-color (closest HTML approximation)
            if !foreground.is_auto() {
                let _ = write!(css, "background-color:#{};", css_hex(foreground));
            } else if !background.is_auto() {
                let _ = write!(css, "background-color:#{};", css_hex(background));
            }
        }
        _ => {}
    }
}

fn borders_to_css(css: &mut String, border: &duke_sheets_core::BorderStyle) {
    if let Some(ref edge) = border.top {
        push_border(css, "border-top", edge);
    }
    if let Some(ref edge) = border.right {
        push_border(css, "border-right", edge);
    }
    if let Some(ref edge) = border.bottom {
        push_border(css, "border-bottom", edge);
    }
    if let Some(ref edge) = border.left {
        push_border(css, "border-left", edge);
    }
}

fn push_border(css: &mut String, prop: &str, edge: &BorderEdge) {
    if matches!(edge.style, BorderLineStyle::None) {
        return;
    }
    let (width, style) = border_line_to_css(edge.style);
    let hex = css_hex(&edge.color);
    let _ = write!(css, "{prop}:{width} {style} #{hex};");
}

fn border_line_to_css(line: BorderLineStyle) -> (&'static str, &'static str) {
    match line {
        BorderLineStyle::None => ("0", "none"),
        BorderLineStyle::Thin | BorderLineStyle::Hair => ("1px", "solid"),
        BorderLineStyle::Medium => ("2px", "solid"),
        BorderLineStyle::Thick => ("3px", "solid"),
        BorderLineStyle::Dashed => ("1px", "dashed"),
        BorderLineStyle::MediumDashed => ("2px", "dashed"),
        BorderLineStyle::Dotted => ("1px", "dotted"),
        BorderLineStyle::Double => ("3px", "double"),
        BorderLineStyle::DashDot | BorderLineStyle::SlantDashDot => ("1px", "dashed"),
        BorderLineStyle::MediumDashDot => ("2px", "dashed"),
        BorderLineStyle::DashDotDot => ("1px", "dotted"),
        BorderLineStyle::MediumDashDotDot => ("2px", "dotted"),
    }
}

fn alignment_to_css(css: &mut String, align: &Alignment) {
    match align.horizontal {
        HorizontalAlignment::Left => css.push_str("text-align:left;"),
        HorizontalAlignment::Center | HorizontalAlignment::CenterContinuous => {
            css.push_str("text-align:center;")
        }
        HorizontalAlignment::Right => css.push_str("text-align:right;"),
        HorizontalAlignment::Justify | HorizontalAlignment::Distributed => {
            css.push_str("text-align:justify;")
        }
        HorizontalAlignment::General | HorizontalAlignment::Fill => {}
    }

    match align.vertical {
        VerticalAlignment::Top => css.push_str("vertical-align:top;"),
        VerticalAlignment::Center => css.push_str("vertical-align:middle;"),
        VerticalAlignment::Bottom => {} // default in our base CSS
        VerticalAlignment::Justify | VerticalAlignment::Distributed => {}
    }

    if align.wrap_text {
        css.push_str("white-space:normal;word-wrap:break-word;");
    }

    if align.indent > 0 {
        let _ = write!(css, "padding-left:{}px;", align.indent as u32 * 8);
    }

    if align.rotation != 0 && align.rotation != 255 {
        let _ = write!(
            css,
            "writing-mode:vertical-lr;transform:rotate({}deg);",
            -align.rotation
        );
    } else if align.rotation == 255 {
        css.push_str("writing-mode:vertical-rl;");
    }
}

fn push_text_decoration(css: &mut String, underline: bool, strikethrough: bool) {
    match (underline, strikethrough) {
        (true, true) => css.push_str("text-decoration:underline line-through;"),
        (true, false) => css.push_str("text-decoration:underline;"),
        (false, true) => css.push_str("text-decoration:line-through;"),
        (false, false) => {}
    }
}

// Merged-cell bookkeeping

enum MergeInfo {
    /// Top-left cell of a merged region.
    Origin { colspan: u16, rowspan: u32 },
    /// Interior cell that should be skipped in output.
    Interior,
}

fn build_merge_map(
    merged_regions: &[CellRange],
    used_range: &CellRange,
) -> HashMap<(u32, u16), MergeInfo> {
    let mut map = HashMap::new();
    for region in merged_regions {
        let start_row = region.start.row;
        let start_col = region.start.col;
        let end_row = region.end.row.min(used_range.end.row);
        let end_col = region.end.col.min(used_range.end.col);

        let colspan = end_col - start_col + 1;
        let rowspan = end_row - start_row + 1;

        map.insert(
            (start_row, start_col),
            MergeInfo::Origin { colspan, rowspan },
        );

        for r in start_row..=end_row {
            for c in start_col..=end_col {
                if r == start_row && c == start_col {
                    continue;
                }
                map.insert((r, c), MergeInfo::Interior);
            }
        }
    }
    map
}

// Unit conversions

/// Convert Excel column-width (in character units) to pixels.
///
/// Excel defines column width as the number of `0` characters that fit using
/// the default font.  For Calibri 11pt the maximum digit width (MDW) is ~7px,
/// giving roughly `width * 7 + 5` pixels of padding.
fn col_width_to_px(width: f64) -> u32 {
    (width * 7.0 + 5.0).round() as u32
}

/// Convert points to CSS pixels (at 96 dpi: 1pt ≈ 1.333px).
fn pt_to_px(pt: f64) -> f64 {
    pt * 96.0 / 72.0
}

/// Convert a `Color` to a 6-character RGB hex string for CSS.
///
/// OOXML stores ARGB as `AARRGGBB` but CSS 8-char hex is `RRGGBBAA`.
/// Rather than swapping, we just drop the alpha and emit plain RGB.
fn css_hex(color: &duke_sheets_core::Color) -> String {
    let (r, g, b) = color.to_rgb();
    format!("{r:02X}{g:02X}{b:02X}")
}

// HTML escaping

fn escape_html(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(ch),
        }
    }
    out
}

// Tests

#[cfg(test)]
mod tests {
    use super::*;
    use duke_sheets_core::Color;

    #[test]
    fn empty_sheet_produces_empty_table() {
        let sheet = Worksheet::new("Empty");
        let html = worksheet_to_html(
            &sheet,
            &HtmlOptions {
                full_document: false,
                ..Default::default()
            },
        );
        assert_eq!(html, "<table></table>\n");
    }

    #[test]
    fn escape_html_entities() {
        assert_eq!(escape_html("<b>A&B</b>"), "&lt;b&gt;A&amp;B&lt;/b&gt;");
        assert_eq!(escape_html("\"hi\""), "&quot;hi&quot;");
    }

    #[test]
    fn basic_cell_renders() {
        let mut sheet = Worksheet::new("Test");
        sheet.set_cell_value("A1", "Hello").unwrap();
        sheet.set_cell_value("B1", 42.0).unwrap();

        let html = worksheet_to_html(
            &sheet,
            &HtmlOptions {
                full_document: false,
                formatted: false,
                ..Default::default()
            },
        );
        assert!(html.contains("<td>Hello</td>"));
        assert!(html.contains("<td>42</td>"));
    }

    #[test]
    fn full_document_wraps_html() {
        let sheet = Worksheet::new("Demo");
        let html = worksheet_to_html(&sheet, &HtmlOptions::default());
        assert!(html.starts_with("<!DOCTYPE html>"));
        assert!(html.contains("<title>Demo</title>"));
        assert!(html.contains("</html>"));
    }

    #[test]
    fn col_width_conversion() {
        // Default 8.43 chars → ~64px
        assert_eq!(col_width_to_px(8.43), 64);
    }

    #[test]
    fn border_css_generation() {
        let edge = BorderEdge::new(BorderLineStyle::Thin, Color::BLACK);
        let mut css = String::new();
        push_border(&mut css, "border-bottom", &edge);
        assert_eq!(css, "border-bottom:1px solid #000000;");
    }
}
