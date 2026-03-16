//! NAPI object types for the Node.js binding.
//!
//! These are plain JS objects (`#[napi(object)]`) used as return types
//! from the read-only API. Each has a `From` impl to convert from the
//! corresponding core Rust type.

use napi_derive::napi;

use duke_sheets_core::{
    self as core,
    style::{
        Alignment as CoreAlignment, BorderEdge as CoreBorderEdge,
        BorderLineStyle as CoreBorderLineStyle, BorderStyle as CoreBorderStyle, Color as CoreColor,
        DiagonalDirection, FillStyle as CoreFillStyle, FontStyle as CoreFontStyle,
        FontVerticalAlign, GradientType, HorizontalAlignment, NumberFormat as CoreNumberFormat,
        PatternType, ReadingOrder, Style as CoreStyle, Underline, VerticalAlignment,
    },
};

/// A single cell within a sparse row.
#[napi(object)]
pub struct JsRowCell {
    /// Column index (0-based).
    pub col: u32,
    /// String representation of the cell value.
    pub value: String,
    /// Cell style (when includeStyles is set).
    pub style: Option<JsStyle>,
    /// Merge span for merge-origin cells (when includeMergeInfo is set).
    pub merge_span: Option<JsMergeSpan>,
    /// Whether this cell is a non-origin member of a merge (when includeMergeInfo is set).
    pub is_merged_secondary: Option<bool>,
    /// Hyperlink (when includeHyperlinks is set).
    pub hyperlink: Option<JsHyperlink>,
    /// Comment (when includeComments is set).
    pub comment: Option<JsComment>,
    /// Formula text (when includeFormulas is set).
    pub formula: Option<String>,
    /// IMAGE() metadata (when includeImages is set).
    pub image: Option<JsImageInfo>,
}

/// A sparse row containing only non-empty cells.
#[napi(object)]
pub struct JsRow {
    /// Row index (0-based).
    pub index: u32,
    /// Non-empty cells in this row, sorted by column.
    pub cells: Vec<JsRowCell>,
}

/// Options for row iteration.
#[napi(object)]
pub struct JsRowsOptions {
    /// Use display-formatted values (e.g., "$1,234.56") instead of raw values.
    pub use_formatted_values: Option<bool>,
    /// Use calculated values for formula cells (requires prior calculate() call).
    pub use_calculated_values: Option<bool>,
    /// Include cell styles.
    pub include_styles: Option<bool>,
    /// Include merge info (mergeSpan + isMergedSecondary).
    pub include_merge_info: Option<bool>,
    /// Include hyperlinks.
    pub include_hyperlinks: Option<bool>,
    /// Include comments.
    pub include_comments: Option<bool>,
    /// Include formula text.
    pub include_formulas: Option<bool>,
    /// Include IMAGE() metadata.
    pub include_images: Option<bool>,
}

/// Color representation. The `colorType` field indicates the variant:
/// `"auto"`, `"rgb"`, `"argb"`, `"theme"`, or `"indexed"`.
/// The `hex` field always contains the resolved 6- or 8-char hex string.
#[napi(object)]
pub struct JsColor {
    pub color_type: String,
    /// Resolved hex string (6 or 8 chars, no `#` prefix).
    pub hex: String,
    pub r: Option<u32>,
    pub g: Option<u32>,
    pub b: Option<u32>,
    pub a: Option<u32>,
    /// Theme color index (0–9), present when `colorType === "theme"`.
    pub theme_index: Option<u32>,
    /// Tint percentage (-100 to 100), present when `colorType === "theme"`.
    pub tint: Option<i32>,
    /// Palette index, present when `colorType === "indexed"`.
    pub palette_index: Option<u32>,
}

impl From<&CoreColor> for JsColor {
    fn from(c: &CoreColor) -> Self {
        let hex = c.to_hex();
        match c {
            CoreColor::Auto => JsColor {
                color_type: "auto".into(),
                hex,
                r: None,
                g: None,
                b: None,
                a: None,
                theme_index: None,
                tint: None,
                palette_index: None,
            },
            CoreColor::Rgb { r, g, b } => JsColor {
                color_type: "rgb".into(),
                hex,
                r: Some(*r as u32),
                g: Some(*g as u32),
                b: Some(*b as u32),
                a: None,
                theme_index: None,
                tint: None,
                palette_index: None,
            },
            CoreColor::Argb { a, r, g, b } => JsColor {
                color_type: "argb".into(),
                hex,
                r: Some(*r as u32),
                g: Some(*g as u32),
                b: Some(*b as u32),
                a: Some(*a as u32),
                theme_index: None,
                tint: None,
                palette_index: None,
            },
            CoreColor::Theme { index, tint } => JsColor {
                color_type: "theme".into(),
                hex,
                r: None,
                g: None,
                b: None,
                a: None,
                theme_index: Some(*index as u32),
                tint: Some(*tint as i32),
                palette_index: None,
            },
            CoreColor::Indexed(i) => JsColor {
                color_type: "indexed".into(),
                hex,
                r: None,
                g: None,
                b: None,
                a: None,
                theme_index: None,
                tint: None,
                palette_index: Some(*i as u32),
            },
        }
    }
}

/// Font style settings.
#[napi(object)]
pub struct JsFontStyle {
    pub name: String,
    pub size: f64,
    pub bold: bool,
    pub italic: bool,
    /// One of: `"none"`, `"single"`, `"double"`, `"singleAccounting"`, `"doubleAccounting"`.
    pub underline: String,
    pub strikethrough: bool,
    pub color: JsColor,
    /// One of: `"baseline"`, `"superscript"`, `"subscript"`.
    pub vertical_align: String,
    pub family: Option<u32>,
    pub charset: Option<u32>,
    pub scheme: Option<String>,
}

fn underline_to_string(u: &Underline) -> &'static str {
    match u {
        Underline::None => "none",
        Underline::Single => "single",
        Underline::Double => "double",
        Underline::SingleAccounting => "singleAccounting",
        Underline::DoubleAccounting => "doubleAccounting",
    }
}

fn font_valign_to_string(v: &FontVerticalAlign) -> &'static str {
    match v {
        FontVerticalAlign::Baseline => "baseline",
        FontVerticalAlign::Superscript => "superscript",
        FontVerticalAlign::Subscript => "subscript",
    }
}

impl From<&CoreFontStyle> for JsFontStyle {
    fn from(f: &CoreFontStyle) -> Self {
        JsFontStyle {
            name: f.name.clone(),
            size: f.size,
            bold: f.bold,
            italic: f.italic,
            underline: underline_to_string(&f.underline).into(),
            strikethrough: f.strikethrough,
            color: JsColor::from(&f.color),
            vertical_align: font_valign_to_string(&f.vertical_align).into(),
            family: f.family.map(|v| v as u32),
            charset: f.charset.map(|v| v as u32),
            scheme: f.scheme.clone(),
        }
    }
}

/// Gradient color stop.
#[napi(object)]
pub struct JsGradientStop {
    pub position: f64,
    pub color: JsColor,
}

fn pattern_type_to_string(p: &PatternType) -> &'static str {
    match p {
        PatternType::None => "none",
        PatternType::Solid => "solid",
        PatternType::MediumGray => "mediumGray",
        PatternType::DarkGray => "darkGray",
        PatternType::LightGray => "lightGray",
        PatternType::DarkHorizontal => "darkHorizontal",
        PatternType::DarkVertical => "darkVertical",
        PatternType::DarkDown => "darkDown",
        PatternType::DarkUp => "darkUp",
        PatternType::DarkGrid => "darkGrid",
        PatternType::DarkTrellis => "darkTrellis",
        PatternType::LightHorizontal => "lightHorizontal",
        PatternType::LightVertical => "lightVertical",
        PatternType::LightDown => "lightDown",
        PatternType::LightUp => "lightUp",
        PatternType::LightGrid => "lightGrid",
        PatternType::LightTrellis => "lightTrellis",
        PatternType::Gray125 => "gray125",
        PatternType::Gray0625 => "gray0625",
    }
}

/// Fill/background style. The `fillType` field indicates the variant:
/// `"none"`, `"solid"`, `"pattern"`, or `"gradient"`.
#[napi(object)]
pub struct JsFillStyle {
    pub fill_type: String,
    /// Solid fill color (present when `fillType === "solid"`).
    pub color: Option<JsColor>,
    /// Pattern type string (present when `fillType === "pattern"`).
    pub pattern: Option<String>,
    /// Pattern foreground color.
    pub foreground: Option<JsColor>,
    /// Pattern background color.
    pub background: Option<JsColor>,
    /// Gradient type: `"linear"` or `"path"` (present when `fillType === "gradient"`).
    pub gradient_type: Option<String>,
    /// Gradient angle in degrees.
    pub angle: Option<f64>,
    /// Gradient color stops.
    pub stops: Option<Vec<JsGradientStop>>,
}

impl From<&CoreFillStyle> for JsFillStyle {
    fn from(f: &CoreFillStyle) -> Self {
        match f {
            CoreFillStyle::None => JsFillStyle {
                fill_type: "none".into(),
                color: None,
                pattern: None,
                foreground: None,
                background: None,
                gradient_type: None,
                angle: None,
                stops: None,
            },
            CoreFillStyle::Solid { color } => JsFillStyle {
                fill_type: "solid".into(),
                color: Some(JsColor::from(color)),
                pattern: None,
                foreground: None,
                background: None,
                gradient_type: None,
                angle: None,
                stops: None,
            },
            CoreFillStyle::Pattern {
                pattern,
                foreground,
                background,
            } => JsFillStyle {
                fill_type: "pattern".into(),
                color: None,
                pattern: Some(pattern_type_to_string(pattern).into()),
                foreground: Some(JsColor::from(foreground)),
                background: Some(JsColor::from(background)),
                gradient_type: None,
                angle: None,
                stops: None,
            },
            CoreFillStyle::Gradient {
                gradient_type,
                angle,
                stops,
            } => JsFillStyle {
                fill_type: "gradient".into(),
                color: None,
                pattern: None,
                foreground: None,
                background: None,
                gradient_type: Some(
                    match gradient_type {
                        GradientType::Linear => "linear",
                        GradientType::Path => "path",
                    }
                    .into(),
                ),
                angle: Some(*angle),
                stops: Some(
                    stops
                        .iter()
                        .map(|s| JsGradientStop {
                            position: s.position,
                            color: JsColor::from(&s.color),
                        })
                        .collect(),
                ),
            },
        }
    }
}

fn border_line_style_to_string(s: &CoreBorderLineStyle) -> &'static str {
    match s {
        CoreBorderLineStyle::None => "none",
        CoreBorderLineStyle::Thin => "thin",
        CoreBorderLineStyle::Medium => "medium",
        CoreBorderLineStyle::Thick => "thick",
        CoreBorderLineStyle::Dashed => "dashed",
        CoreBorderLineStyle::Dotted => "dotted",
        CoreBorderLineStyle::Double => "double",
        CoreBorderLineStyle::Hair => "hair",
        CoreBorderLineStyle::MediumDashed => "mediumDashed",
        CoreBorderLineStyle::DashDot => "dashDot",
        CoreBorderLineStyle::MediumDashDot => "mediumDashDot",
        CoreBorderLineStyle::DashDotDot => "dashDotDot",
        CoreBorderLineStyle::MediumDashDotDot => "mediumDashDotDot",
        CoreBorderLineStyle::SlantDashDot => "slantDashDot",
    }
}

/// A single border edge (line style + color).
#[napi(object)]
pub struct JsBorderEdge {
    /// One of: `"none"`, `"thin"`, `"medium"`, `"thick"`, `"dashed"`, `"dotted"`,
    /// `"double"`, `"hair"`, `"mediumDashed"`, `"dashDot"`, `"mediumDashDot"`,
    /// `"dashDotDot"`, `"mediumDashDotDot"`, `"slantDashDot"`.
    pub style: String,
    pub color: JsColor,
}

impl From<&CoreBorderEdge> for JsBorderEdge {
    fn from(e: &CoreBorderEdge) -> Self {
        JsBorderEdge {
            style: border_line_style_to_string(&e.style).into(),
            color: JsColor::from(&e.color),
        }
    }
}

/// Cell border style.
#[napi(object)]
pub struct JsBorderStyle {
    pub left: Option<JsBorderEdge>,
    pub right: Option<JsBorderEdge>,
    pub top: Option<JsBorderEdge>,
    pub bottom: Option<JsBorderEdge>,
    pub diagonal: Option<JsBorderEdge>,
    /// One of: `"none"`, `"down"`, `"up"`, `"both"`.
    pub diagonal_direction: String,
}

impl From<&CoreBorderStyle> for JsBorderStyle {
    fn from(b: &CoreBorderStyle) -> Self {
        JsBorderStyle {
            left: b.left.as_ref().map(JsBorderEdge::from),
            right: b.right.as_ref().map(JsBorderEdge::from),
            top: b.top.as_ref().map(JsBorderEdge::from),
            bottom: b.bottom.as_ref().map(JsBorderEdge::from),
            diagonal: b.diagonal.as_ref().map(JsBorderEdge::from),
            diagonal_direction: match b.diagonal_direction {
                DiagonalDirection::None => "none",
                DiagonalDirection::Down => "down",
                DiagonalDirection::Up => "up",
                DiagonalDirection::Both => "both",
            }
            .into(),
        }
    }
}

fn horizontal_alignment_to_string(a: &HorizontalAlignment) -> &'static str {
    match a {
        HorizontalAlignment::General => "general",
        HorizontalAlignment::Left => "left",
        HorizontalAlignment::Center => "center",
        HorizontalAlignment::Right => "right",
        HorizontalAlignment::Fill => "fill",
        HorizontalAlignment::Justify => "justify",
        HorizontalAlignment::CenterContinuous => "centerContinuous",
        HorizontalAlignment::Distributed => "distributed",
    }
}

fn vertical_alignment_to_string(a: &VerticalAlignment) -> &'static str {
    match a {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Bottom => "bottom",
        VerticalAlignment::Justify => "justify",
        VerticalAlignment::Distributed => "distributed",
    }
}

fn reading_order_to_string(r: &ReadingOrder) -> &'static str {
    match r {
        ReadingOrder::ContextDependent => "contextDependent",
        ReadingOrder::LeftToRight => "leftToRight",
        ReadingOrder::RightToLeft => "rightToLeft",
    }
}

/// Text alignment settings.
#[napi(object)]
pub struct JsAlignment {
    pub horizontal: String,
    pub vertical: String,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
    pub indent: u32,
    /// Rotation in degrees (-90 to 90, or 255 for vertical text).
    pub rotation: i32,
    pub reading_order: String,
}

impl From<&CoreAlignment> for JsAlignment {
    fn from(a: &CoreAlignment) -> Self {
        JsAlignment {
            horizontal: horizontal_alignment_to_string(&a.horizontal).into(),
            vertical: vertical_alignment_to_string(&a.vertical).into(),
            wrap_text: a.wrap_text,
            shrink_to_fit: a.shrink_to_fit,
            indent: a.indent as u32,
            rotation: a.rotation as i32,
            reading_order: reading_order_to_string(&a.reading_order).into(),
        }
    }
}

/// Number format. The `formatType` field indicates the variant:
/// `"general"`, `"builtin"`, or `"custom"`.
#[napi(object)]
pub struct JsNumberFormat {
    pub format_type: String,
    /// Built-in format ID (present when `formatType === "builtin"`).
    pub id: Option<u32>,
    /// The resolved format string (always present).
    pub format_string: String,
    /// Whether this format represents a date/time.
    pub is_date_format: bool,
}

impl From<&CoreNumberFormat> for JsNumberFormat {
    fn from(n: &CoreNumberFormat) -> Self {
        JsNumberFormat {
            format_type: match n {
                CoreNumberFormat::General => "general",
                CoreNumberFormat::BuiltIn(_) => "builtin",
                CoreNumberFormat::Custom(_) => "custom",
            }
            .into(),
            id: match n {
                CoreNumberFormat::BuiltIn(id) => Some(*id),
                _ => None,
            },
            format_string: n.format_string().to_string(),
            is_date_format: n.is_date_format(),
        }
    }
}

/// Cell protection settings.
#[napi(object)]
pub struct JsCellProtection {
    pub locked: bool,
    pub hidden: bool,
}

/// Complete cell style including font, fill, border, alignment, number format,
/// and protection settings.
#[napi(object)]
pub struct JsStyle {
    pub font: JsFontStyle,
    pub fill: JsFillStyle,
    pub border: JsBorderStyle,
    pub alignment: JsAlignment,
    pub number_format: JsNumberFormat,
    pub protection: JsCellProtection,
}

impl From<&CoreStyle> for JsStyle {
    fn from(s: &CoreStyle) -> Self {
        JsStyle {
            font: JsFontStyle::from(&s.font),
            fill: JsFillStyle::from(&s.fill),
            border: JsBorderStyle::from(&s.border),
            alignment: JsAlignment::from(&s.alignment),
            number_format: JsNumberFormat::from(&s.number_format),
            protection: JsCellProtection {
                locked: s.protection.locked,
                hidden: s.protection.hidden,
            },
        }
    }
}

/// A hyperlink attached to a cell.
#[napi(object)]
pub struct JsHyperlink {
    /// Target URL (external) or cell reference (internal).
    pub target: String,
    /// Display text (shown in cell; `null` means cell value is used).
    pub display: Option<String>,
    /// Tooltip shown on hover.
    pub tooltip: Option<String>,
    /// Location within target (e.g., sheet reference for internal links).
    pub location: Option<String>,
}

impl From<&core::Hyperlink> for JsHyperlink {
    fn from(h: &core::Hyperlink) -> Self {
        JsHyperlink {
            target: h.target.clone(),
            display: h.display.clone(),
            tooltip: h.tooltip.clone(),
            location: h.location.clone(),
        }
    }
}

/// A cell comment/note.
#[napi(object)]
pub struct JsComment {
    pub author: String,
    pub text: String,
    pub visible: bool,
}

impl From<&core::CellComment> for JsComment {
    fn from(c: &core::CellComment) -> Self {
        JsComment {
            author: c.author.clone(),
            text: c.text.clone(),
            visible: c.visible,
        }
    }
}

/// A comment with its cell address.
#[napi(object)]
pub struct JsCommentEntry {
    pub row: u32,
    pub col: u32,
    pub comment: JsComment,
}

/// Freeze pane settings.
#[napi(object)]
pub struct JsFreezePanes {
    /// First unfrozen row.
    pub row: u32,
    /// First unfrozen column.
    pub col: u32,
}

impl From<&core::FreezePanes> for JsFreezePanes {
    fn from(f: &core::FreezePanes) -> Self {
        JsFreezePanes {
            row: f.row,
            col: f.col as u32,
        }
    }
}

/// Split pane settings.
#[napi(object)]
pub struct JsSplitPanes {
    pub x_split: f64,
    pub y_split: f64,
    pub top_left_row: Option<u32>,
    pub top_left_col: Option<u32>,
    pub active_pane: Option<String>,
}

impl From<&core::SplitPanes> for JsSplitPanes {
    fn from(s: &core::SplitPanes) -> Self {
        JsSplitPanes {
            x_split: s.x_split,
            y_split: s.y_split,
            top_left_row: s.top_left.map(|(r, _)| r),
            top_left_col: s.top_left.map(|(_, c)| c as u32),
            active_pane: s.active_pane.clone(),
        }
    }
}

/// A selection within a sheet view.
#[napi(object)]
pub struct JsSelection {
    pub pane: Option<String>,
    pub active_cell: Option<String>,
    pub sqref: Option<String>,
}

impl From<&core::Selection> for JsSelection {
    fn from(s: &core::Selection) -> Self {
        JsSelection {
            pane: s.pane.clone(),
            active_cell: s.active_cell.clone(),
            sqref: s.sqref.clone(),
        }
    }
}

/// Sheet protection settings.
#[napi(object)]
pub struct JsSheetProtection {
    pub protected: bool,
    pub select_locked_cells: bool,
    pub select_unlocked_cells: bool,
    pub format_cells: bool,
    pub format_columns: bool,
    pub format_rows: bool,
    pub insert_columns: bool,
    pub insert_rows: bool,
    pub insert_hyperlinks: bool,
    pub delete_columns: bool,
    pub delete_rows: bool,
    pub sort: bool,
    pub auto_filter: bool,
    pub pivot_tables: bool,
}

impl From<&core::SheetProtection> for JsSheetProtection {
    fn from(p: &core::SheetProtection) -> Self {
        JsSheetProtection {
            protected: p.protected,
            select_locked_cells: p.select_locked_cells,
            select_unlocked_cells: p.select_unlocked_cells,
            format_cells: p.format_cells,
            format_columns: p.format_columns,
            format_rows: p.format_rows,
            insert_columns: p.insert_columns,
            insert_rows: p.insert_rows,
            insert_hyperlinks: p.insert_hyperlinks,
            delete_columns: p.delete_columns,
            delete_rows: p.delete_rows,
            sort: p.sort,
            auto_filter: p.auto_filter,
            pivot_tables: p.pivot_tables,
        }
    }
}

/// Page setup / print settings.
#[napi(object)]
pub struct JsPageSetup {
    /// Paper size (1 = Letter, 9 = A4, etc.).
    pub paper_size: u32,
    /// `"portrait"` or `"landscape"`.
    pub orientation: String,
    /// Scale percentage (10–400).
    pub scale: u32,
    pub fit_to_width: Option<u32>,
    pub fit_to_height: Option<u32>,
    pub top_margin: f64,
    pub bottom_margin: f64,
    pub left_margin: f64,
    pub right_margin: f64,
    pub header_margin: f64,
    pub footer_margin: f64,
    pub print_gridlines: bool,
    pub print_headings: bool,
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
    pub different_odd_even: bool,
    pub different_first: bool,
    pub scale_with_doc: bool,
    pub align_with_margins: bool,
}

impl From<&core::PageSetup> for JsPageSetup {
    fn from(p: &core::PageSetup) -> Self {
        JsPageSetup {
            paper_size: p.paper_size as u32,
            orientation: match p.orientation {
                core::PageOrientation::Portrait => "portrait",
                core::PageOrientation::Landscape => "landscape",
            }
            .into(),
            scale: p.scale as u32,
            fit_to_width: p.fit_to_width.map(|v| v as u32),
            fit_to_height: p.fit_to_height.map(|v| v as u32),
            top_margin: p.top_margin,
            bottom_margin: p.bottom_margin,
            left_margin: p.left_margin,
            right_margin: p.right_margin,
            header_margin: p.header_margin,
            footer_margin: p.footer_margin,
            print_gridlines: p.print_gridlines,
            print_headings: p.print_headings,
            odd_header: p.odd_header.clone(),
            odd_footer: p.odd_footer.clone(),
            even_header: p.even_header.clone(),
            even_footer: p.even_footer.clone(),
            first_header: p.first_header.clone(),
            first_footer: p.first_footer.clone(),
            different_odd_even: p.different_odd_even,
            different_first: p.different_first,
            scale_with_doc: p.scale_with_doc,
            align_with_margins: p.align_with_margins,
        }
    }
}

/// A manual page break (row or column).
#[napi(object)]
pub struct JsPageBreak {
    /// Row index (for row breaks) or column index (for col breaks), 0-based.
    pub id: u32,
    pub min: u32,
    pub max: u32,
    /// Whether this is a manual break.
    pub manual: bool,
}

impl From<&core::PageBreak> for JsPageBreak {
    fn from(b: &core::PageBreak) -> Self {
        JsPageBreak {
            id: b.id,
            min: b.min,
            max: b.max,
            manual: b.man,
        }
    }
}

/// Workbook-level settings.
#[napi(object)]
pub struct JsWorkbookSettings {
    /// Whether the 1904 date system is used (macOS default).
    pub date_1904: bool,
    /// Whether the workbook structure is protected.
    pub protected: bool,
    /// Calculate formulas on open.
    pub calc_on_open: bool,
    pub theme: Option<String>,
}

impl From<&core::WorkbookSettings> for JsWorkbookSettings {
    fn from(s: &core::WorkbookSettings) -> Self {
        JsWorkbookSettings {
            date_1904: s.date_1904,
            protected: s.protected,
            calc_on_open: s.calc_on_open,
            theme: s.theme.clone(),
        }
    }
}

/// A named range definition.
#[napi(object)]
pub struct JsNamedRange {
    pub name: String,
    /// `"workbook"` or `"sheet"`.
    pub scope: String,
    /// Sheet index when `scope === "sheet"`.
    pub sheet_index: Option<u32>,
    /// The formula/reference the name refers to.
    pub refers_to: String,
    pub comment: Option<String>,
    pub hidden: bool,
}

/// An Excel table (ListObject).
#[napi(object)]
pub struct JsTable {
    pub id: u32,
    pub name: String,
    pub display_name: String,
    /// Range string (e.g., `"A1:D10"`).
    pub reference: String,
    pub columns: Vec<JsTableColumn>,
    pub style_info: Option<JsTableStyleInfo>,
    pub header_row_count: u32,
    pub totals_row_count: u32,
    pub totals_row_shown: bool,
}

impl From<&core::Table> for JsTable {
    fn from(t: &core::Table) -> Self {
        JsTable {
            id: t.id,
            name: t.name.clone(),
            display_name: t.display_name.clone(),
            reference: t.reference.to_string(),
            columns: t.columns.iter().map(JsTableColumn::from).collect(),
            style_info: t.style_info.as_ref().map(JsTableStyleInfo::from),
            header_row_count: t.header_row_count,
            totals_row_count: t.totals_row_count,
            totals_row_shown: t.totals_row_shown,
        }
    }
}

/// A column within a table.
#[napi(object)]
pub struct JsTableColumn {
    pub id: u32,
    pub name: String,
    /// One of: `"average"`, `"count"`, `"countNums"`, `"max"`, `"min"`,
    /// `"sum"`, `"stdDev"`, `"var"`, `"custom"`, `"none"`, or `null`.
    pub totals_row_function: Option<String>,
    pub totals_row_formula: Option<String>,
    pub totals_row_label: Option<String>,
    pub calculated_column_formula: Option<String>,
}

impl From<&core::TableColumn> for JsTableColumn {
    fn from(c: &core::TableColumn) -> Self {
        JsTableColumn {
            id: c.id,
            name: c.name.clone(),
            totals_row_function: c.totals_row_function.as_ref().map(|f| f.to_ooxml().into()),
            totals_row_formula: c.totals_row_formula.clone(),
            totals_row_label: c.totals_row_label.clone(),
            calculated_column_formula: c.calculated_column_formula.clone(),
        }
    }
}

/// Table style configuration.
#[napi(object)]
pub struct JsTableStyleInfo {
    pub name: Option<String>,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
}

impl From<&core::TableStyleInfo> for JsTableStyleInfo {
    fn from(s: &core::TableStyleInfo) -> Self {
        JsTableStyleInfo {
            name: s.name.clone(),
            show_first_column: s.show_first_column,
            show_last_column: s.show_last_column,
            show_row_stripes: s.show_row_stripes,
            show_column_stripes: s.show_column_stripes,
        }
    }
}

/// A standalone auto-filter on a worksheet.
#[napi(object)]
pub struct JsAutoFilter {
    /// Range string the filter covers (e.g., `"A1:D10"`).
    pub range: String,
    pub filter_columns: Vec<JsFilterColumn>,
}

impl From<&core::AutoFilter> for JsAutoFilter {
    fn from(af: &core::AutoFilter) -> Self {
        JsAutoFilter {
            range: af.range.to_string(),
            filter_columns: af.filter_columns.iter().map(JsFilterColumn::from).collect(),
        }
    }
}

/// A filter on a single column.
#[napi(object)]
pub struct JsFilterColumn {
    pub col_id: u32,
    pub hidden_button: bool,
    pub show_button: bool,
    /// The type of filter: `"values"`, `"custom"`, `"top10"`, `"dynamic"`, or `"color"`.
    pub filter_type: String,
    /// Values for a discrete value filter (present when `filterType === "values"`).
    pub values: Option<Vec<String>>,
    /// Include blanks (present when `filterType === "values"`).
    pub blank: Option<bool>,
}

impl From<&core::FilterColumn> for JsFilterColumn {
    fn from(fc: &core::FilterColumn) -> Self {
        let (filter_type, values, blank) = match &fc.filter {
            core::ColumnFilter::Values(vf) => ("values", Some(vf.values.clone()), Some(vf.blank)),
            core::ColumnFilter::Custom(_) => ("custom", None, None),
            core::ColumnFilter::Top10(_) => ("top10", None, None),
            core::ColumnFilter::Dynamic(_) => ("dynamic", None, None),
            core::ColumnFilter::Color(_) => ("color", None, None),
        };
        JsFilterColumn {
            col_id: fc.col_id,
            hidden_button: fc.hidden_button,
            show_button: fc.show_button,
            filter_type: filter_type.into(),
            values,
            blank,
        }
    }
}

/// A data validation rule.
#[napi(object)]
pub struct JsDataValidation {
    /// The type of validation: `"none"`, `"whole"`, `"decimal"`, `"list"`,
    /// `"date"`, `"time"`, `"textLength"`, `"custom"`.
    pub validation_type: String,
    /// Ranges this validation applies to (as range strings).
    pub ranges: Vec<String>,
    pub allow_blank: bool,
    pub show_dropdown: bool,
    pub show_input_message: bool,
    pub input_title: Option<String>,
    pub input_message: Option<String>,
    pub show_error_alert: bool,
    /// `"stop"`, `"warning"`, or `"information"`.
    pub error_style: String,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
    /// Operator (present for numeric/date/time/textLength validations):
    /// `"between"`, `"notBetween"`, `"equal"`, `"notEqual"`, `"greaterThan"`,
    /// `"lessThan"`, `"greaterThanOrEqual"`, `"lessThanOrEqual"`.
    pub operator: Option<String>,
    /// First value/formula (present for most validation types).
    pub value1: Option<String>,
    /// Second value/formula (present for `between`/`notBetween`).
    pub value2: Option<String>,
    /// List source string (present when `validationType === "list"`).
    pub list_source: Option<String>,
    /// Custom formula (present when `validationType === "custom"`).
    pub formula: Option<String>,
}

fn validation_operator_to_string(op: &core::ValidationOperator) -> &'static str {
    match op {
        core::ValidationOperator::Between => "between",
        core::ValidationOperator::NotBetween => "notBetween",
        core::ValidationOperator::Equal => "equal",
        core::ValidationOperator::NotEqual => "notEqual",
        core::ValidationOperator::GreaterThan => "greaterThan",
        core::ValidationOperator::LessThan => "lessThan",
        core::ValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        core::ValidationOperator::LessThanOrEqual => "lessThanOrEqual",
    }
}

impl From<&core::DataValidation> for JsDataValidation {
    fn from(dv: &core::DataValidation) -> Self {
        let (vtype, operator, value1, value2, list_source, formula) = match &dv.validation_type {
            core::ValidationType::None => ("none", None, None, None, None, None),
            core::ValidationType::Whole {
                operator,
                value1,
                value2,
            } => (
                "whole",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::Decimal {
                operator,
                value1,
                value2,
            } => (
                "decimal",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::List { source } => {
                ("list", None, None, None, Some(source.clone()), None)
            }
            core::ValidationType::Date {
                operator,
                value1,
                value2,
            } => (
                "date",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::Time {
                operator,
                value1,
                value2,
            } => (
                "time",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::TextLength {
                operator,
                value1,
                value2,
            } => (
                "textLength",
                Some(validation_operator_to_string(operator)),
                Some(value1.clone()),
                value2.clone(),
                None,
                None,
            ),
            core::ValidationType::Custom { formula } => {
                ("custom", None, None, None, None, Some(formula.clone()))
            }
        };

        JsDataValidation {
            validation_type: vtype.into(),
            ranges: dv.ranges.iter().map(|r| r.to_string()).collect(),
            allow_blank: dv.allow_blank,
            show_dropdown: dv.show_dropdown,
            show_input_message: dv.show_input_message,
            input_title: dv.input_title.clone(),
            input_message: dv.input_message.clone(),
            show_error_alert: dv.show_error_alert,
            error_style: match dv.error_style {
                core::ValidationErrorStyle::Stop => "stop",
                core::ValidationErrorStyle::Warning => "warning",
                core::ValidationErrorStyle::Information => "information",
            }
            .into(),
            error_title: dv.error_title.clone(),
            error_message: dv.error_message.clone(),
            operator: operator.map(Into::into),
            value1,
            value2,
            list_source,
            formula,
        }
    }
}

/// A conditional formatting rule.
#[napi(object)]
pub struct JsConditionalFormatRule {
    /// The type of rule: `"cellIs"`, `"expression"`, `"colorScale"`, `"dataBar"`,
    /// `"iconSet"`, `"top10"`, `"aboveAverage"`, `"containsText"`, `"beginsWith"`,
    /// `"endsWith"`, `"duplicateValues"`, `"uniqueValues"`, `"containsBlanks"`,
    /// `"notContainsBlanks"`, `"containsErrors"`, `"notContainsErrors"`, `"timePeriod"`.
    pub rule_type: String,
    /// Ranges this rule applies to (as range strings).
    pub ranges: Vec<String>,
    /// Lower number = higher priority.
    pub priority: u32,
    pub stop_if_true: bool,
    /// Operator (present for `cellIs` rules).
    pub operator: Option<String>,
    /// Formula/value (present for `cellIs`, `expression` rules).
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    /// Text value (present for `containsText`, `beginsWith`, `endsWith`).
    pub text: Option<String>,
    /// Rank for top/bottom N rules.
    pub rank: Option<u32>,
    /// Whether the top/bottom N rule uses percentages.
    pub percent: Option<bool>,
    /// Whether it's a "bottom N" rule (vs top N).
    pub bottom: Option<bool>,
}

impl From<&core::ConditionalFormatRule> for JsConditionalFormatRule {
    fn from(r: &core::ConditionalFormatRule) -> Self {
        let (rule_type, operator, formula1, formula2, text, rank, percent, bottom) =
            match &r.rule_type {
                core::CfRuleType::CellIs {
                    operator,
                    formula1,
                    formula2,
                } => {
                    let op = match operator {
                        core::CfOperator::Between => "between",
                        core::CfOperator::NotBetween => "notBetween",
                        core::CfOperator::Equal => "equal",
                        core::CfOperator::NotEqual => "notEqual",
                        core::CfOperator::GreaterThan => "greaterThan",
                        core::CfOperator::LessThan => "lessThan",
                        core::CfOperator::GreaterThanOrEqual => "greaterThanOrEqual",
                        core::CfOperator::LessThanOrEqual => "lessThanOrEqual",
                    };
                    (
                        "cellIs",
                        Some(op),
                        Some(formula1.clone()),
                        formula2.clone(),
                        None,
                        None,
                        None,
                        None,
                    )
                }
                core::CfRuleType::Expression { formula } => (
                    "expression",
                    None,
                    Some(formula.clone()),
                    None,
                    None,
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::ColorScale { .. } => {
                    ("colorScale", None, None, None, None, None, None, None)
                }
                core::CfRuleType::DataBar { .. } => {
                    ("dataBar", None, None, None, None, None, None, None)
                }
                core::CfRuleType::IconSet { .. } => {
                    ("iconSet", None, None, None, None, None, None, None)
                }
                core::CfRuleType::Top10 {
                    rank,
                    percent,
                    bottom,
                } => (
                    "top10",
                    None,
                    None,
                    None,
                    None,
                    Some(*rank),
                    Some(*percent),
                    Some(*bottom),
                ),
                core::CfRuleType::AboveAverage { .. } => {
                    ("aboveAverage", None, None, None, None, None, None, None)
                }
                core::CfRuleType::ContainsText { text } => (
                    "containsText",
                    None,
                    None,
                    None,
                    Some(text.clone()),
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::BeginsWith { text } => (
                    "beginsWith",
                    None,
                    None,
                    None,
                    Some(text.clone()),
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::EndsWith { text } => (
                    "endsWith",
                    None,
                    None,
                    None,
                    Some(text.clone()),
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::DuplicateValues => {
                    ("duplicateValues", None, None, None, None, None, None, None)
                }
                core::CfRuleType::UniqueValues => {
                    ("uniqueValues", None, None, None, None, None, None, None)
                }
                core::CfRuleType::ContainsBlanks => {
                    ("containsBlanks", None, None, None, None, None, None, None)
                }
                core::CfRuleType::NotContainsBlanks => (
                    "notContainsBlanks",
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::ContainsErrors => {
                    ("containsErrors", None, None, None, None, None, None, None)
                }
                core::CfRuleType::NotContainsErrors => (
                    "notContainsErrors",
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                ),
                core::CfRuleType::TimePeriod { .. } => {
                    ("timePeriod", None, None, None, None, None, None, None)
                }
            };

        JsConditionalFormatRule {
            rule_type: rule_type.into(),
            ranges: r.ranges.iter().map(|rng| rng.to_string()).collect(),
            priority: r.priority,
            stop_if_true: r.stop_if_true,
            operator: operator.map(Into::into),
            formula1,
            formula2,
            text,
            rank,
            percent,
            bottom,
        }
    }
}

/// A single run of rich text.
#[napi(object)]
pub struct JsRichTextRun {
    pub text: String,
    pub font: Option<JsRunFont>,
}

impl From<&core::RichTextRun> for JsRichTextRun {
    fn from(r: &core::RichTextRun) -> Self {
        JsRichTextRun {
            text: r.text.clone(),
            font: r.font.as_ref().map(JsRunFont::from),
        }
    }
}

/// Font properties for a rich text run (all fields optional — unset inherits cell style).
#[napi(object)]
pub struct JsRunFont {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub size: Option<f64>,
    pub color: Option<JsColor>,
    pub name: Option<String>,
    pub underline: Option<String>,
    pub strikethrough: Option<bool>,
    pub vertical_align: Option<String>,
}

impl From<&core::RunFont> for JsRunFont {
    fn from(f: &core::RunFont) -> Self {
        JsRunFont {
            bold: f.bold,
            italic: f.italic,
            size: f.size,
            color: f.color.as_ref().map(JsColor::from),
            name: f.name.clone(),
            underline: f.underline.as_ref().map(|u| underline_to_string(u).into()),
            strikethrough: f.strikethrough,
            vertical_align: f
                .vertical_align
                .as_ref()
                .map(|v| font_valign_to_string(v).into()),
        }
    }
}

/// A hyperlink with its cell address.
#[napi(object)]
pub struct JsHyperlinkEntry {
    pub address: String,
    pub hyperlink: JsHyperlink,
}

/// A formula cell with address.
#[napi(object)]
pub struct JsFormulaCell {
    pub row: u32,
    pub col: u32,
    pub formula: String,
}

/// A cell with address and value.
#[napi(object)]
pub struct JsSpillSource {
    pub row: u32,
    pub col: u32,
}

/// A merged cell region with structured coordinates.
#[napi(object)]
pub struct JsMergedRegion {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
    /// The range as an A1-style string (e.g., "A1:C3").
    pub range: String,
}

/// The row/column span of a merged region's origin cell.
#[napi(object)]
pub struct JsMergeSpan {
    pub row_span: u32,
    pub col_span: u32,
}

/// IMAGE() metadata captured during calculation.
#[napi(object)]
pub struct JsImageInfo {
    /// IMAGE source URL or path.
    pub source: String,
    /// IMAGE alternate text.
    pub alt_text: String,
    /// 0=FitCell, 1=FillCell, 2=OriginalSize, 3=Custom
    pub sizing: u32,
    /// Optional custom width.
    pub width: Option<f64>,
    /// Optional custom height.
    pub height: Option<f64>,
}
