use serde::{Deserialize, Serialize};

use duke_sheets_core::{
    self as core,
    named_range::{NameScope, NamedRange},
    style::{
        Alignment as CoreAlignment, BorderEdge as CoreBorderEdge,
        BorderLineStyle as CoreBorderLineStyle, BorderStyle as CoreBorderStyle, Color as CoreColor,
        DiagonalDirection, FillStyle as CoreFillStyle, FontStyle as CoreFontStyle,
        FontVerticalAlign, GradientType, HorizontalAlignment, NumberFormat as CoreNumberFormat,
        PatternType, ReadingOrder, Style as CoreStyle, Underline, VerticalAlignment,
    },
};

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmRowCell {
    pub col: u32,
    pub value: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<WasmStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub merge_span: Option<WasmMergeSpan>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_merged_secondary: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<WasmHyperlink>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comment: Option<WasmComment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image: Option<WasmImageInfo>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmRow {
    pub index: u32,
    pub cells: Vec<WasmRowCell>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmRowsOptions {
    pub use_formatted_values: Option<bool>,
    pub use_calculated_values: Option<bool>,
    pub include_styles: Option<bool>,
    pub include_merge_info: Option<bool>,
    pub include_hyperlinks: Option<bool>,
    pub include_comments: Option<bool>,
    pub include_formulas: Option<bool>,
    pub include_images: Option<bool>,
    pub skip_empty_values: Option<bool>,
    pub skip_blank_values: Option<bool>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmImageInfo {
    pub source: String,
    pub alt_text: String,
    pub sizing: u32,
    pub width: Option<f64>,
    pub height: Option<f64>,
}

impl From<core::ImageInfo> for WasmImageInfo {
    fn from(info: core::ImageInfo) -> Self {
        Self {
            source: info.source,
            alt_text: info.alt_text,
            sizing: match info.sizing {
                core::ImageSizing::FitCell => 0,
                core::ImageSizing::FillCell => 1,
                core::ImageSizing::OriginalSize => 2,
                core::ImageSizing::Custom => 3,
            },
            width: info.width,
            height: info.height,
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmColor {
    pub color_type: String,
    pub hex: String,
    pub r: Option<u32>,
    pub g: Option<u32>,
    pub b: Option<u32>,
    pub a: Option<u32>,
    pub theme_index: Option<u32>,
    pub tint: Option<i32>,
    pub palette_index: Option<u32>,
}

impl From<&CoreColor> for WasmColor {
    fn from(c: &CoreColor) -> Self {
        let hex = c.to_hex();
        match c {
            CoreColor::Auto => Self {
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
            CoreColor::Rgb { r, g, b } => Self {
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
            CoreColor::Argb { a, r, g, b } => Self {
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
            CoreColor::Theme { index, tint } => Self {
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
            CoreColor::Indexed(i) => Self {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmFontStyle {
    pub name: String,
    pub size: f64,
    pub bold: bool,
    pub italic: bool,
    pub underline: String,
    pub strikethrough: bool,
    pub color: WasmColor,
    pub vertical_align: String,
    pub family: Option<u32>,
    pub charset: Option<u32>,
    pub scheme: Option<String>,
}

pub(crate) fn underline_to_string(u: &Underline) -> &'static str {
    match u {
        Underline::None => "none",
        Underline::Single => "single",
        Underline::Double => "double",
        Underline::SingleAccounting => "singleAccounting",
        Underline::DoubleAccounting => "doubleAccounting",
    }
}

pub(crate) fn font_valign_to_string(v: &FontVerticalAlign) -> &'static str {
    match v {
        FontVerticalAlign::Baseline => "baseline",
        FontVerticalAlign::Superscript => "superscript",
        FontVerticalAlign::Subscript => "subscript",
    }
}

impl From<&CoreFontStyle> for WasmFontStyle {
    fn from(f: &CoreFontStyle) -> Self {
        Self {
            name: f.name.clone(),
            size: f.size,
            bold: f.bold,
            italic: f.italic,
            underline: underline_to_string(&f.underline).into(),
            strikethrough: f.strikethrough,
            color: WasmColor::from(&f.color),
            vertical_align: font_valign_to_string(&f.vertical_align).into(),
            family: f.family.map(|v| v as u32),
            charset: f.charset.map(|v| v as u32),
            scheme: f.scheme.clone(),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmGradientStop {
    pub position: f64,
    pub color: WasmColor,
}

pub(crate) fn pattern_type_to_string(p: &PatternType) -> &'static str {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmFillStyle {
    pub fill_type: String,
    pub color: Option<WasmColor>,
    pub pattern: Option<String>,
    pub foreground: Option<WasmColor>,
    pub background: Option<WasmColor>,
    pub gradient_type: Option<String>,
    pub angle: Option<f64>,
    pub stops: Option<Vec<WasmGradientStop>>,
}

impl From<&CoreFillStyle> for WasmFillStyle {
    fn from(f: &CoreFillStyle) -> Self {
        match f {
            CoreFillStyle::None => Self {
                fill_type: "none".into(),
                color: None,
                pattern: None,
                foreground: None,
                background: None,
                gradient_type: None,
                angle: None,
                stops: None,
            },
            CoreFillStyle::Solid { color } => Self {
                fill_type: "solid".into(),
                color: Some(WasmColor::from(color)),
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
            } => Self {
                fill_type: "pattern".into(),
                color: None,
                pattern: Some(pattern_type_to_string(pattern).into()),
                foreground: Some(WasmColor::from(foreground)),
                background: Some(WasmColor::from(background)),
                gradient_type: None,
                angle: None,
                stops: None,
            },
            CoreFillStyle::Gradient {
                gradient_type,
                angle,
                stops,
            } => Self {
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
                        .map(|s| WasmGradientStop {
                            position: s.position,
                            color: WasmColor::from(&s.color),
                        })
                        .collect(),
                ),
            },
        }
    }
}

pub(crate) fn border_line_style_to_string(s: &CoreBorderLineStyle) -> &'static str {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmBorderEdge {
    pub style: String,
    pub color: WasmColor,
}

impl From<&CoreBorderEdge> for WasmBorderEdge {
    fn from(e: &CoreBorderEdge) -> Self {
        Self {
            style: border_line_style_to_string(&e.style).into(),
            color: WasmColor::from(&e.color),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmBorderStyle {
    pub left: Option<WasmBorderEdge>,
    pub right: Option<WasmBorderEdge>,
    pub top: Option<WasmBorderEdge>,
    pub bottom: Option<WasmBorderEdge>,
    pub diagonal: Option<WasmBorderEdge>,
    pub diagonal_direction: String,
}

impl From<&CoreBorderStyle> for WasmBorderStyle {
    fn from(b: &CoreBorderStyle) -> Self {
        Self {
            left: b.left.as_ref().map(WasmBorderEdge::from),
            right: b.right.as_ref().map(WasmBorderEdge::from),
            top: b.top.as_ref().map(WasmBorderEdge::from),
            bottom: b.bottom.as_ref().map(WasmBorderEdge::from),
            diagonal: b.diagonal.as_ref().map(WasmBorderEdge::from),
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

pub(crate) fn horizontal_alignment_to_string(a: &HorizontalAlignment) -> &'static str {
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

pub(crate) fn vertical_alignment_to_string(a: &VerticalAlignment) -> &'static str {
    match a {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Bottom => "bottom",
        VerticalAlignment::Justify => "justify",
        VerticalAlignment::Distributed => "distributed",
    }
}

pub(crate) fn reading_order_to_string(r: &ReadingOrder) -> &'static str {
    match r {
        ReadingOrder::ContextDependent => "contextDependent",
        ReadingOrder::LeftToRight => "leftToRight",
        ReadingOrder::RightToLeft => "rightToLeft",
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmAlignment {
    pub horizontal: String,
    pub vertical: String,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
    pub indent: u32,
    pub rotation: i32,
    pub reading_order: String,
}

impl From<&CoreAlignment> for WasmAlignment {
    fn from(a: &CoreAlignment) -> Self {
        Self {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmNumberFormat {
    pub format_type: String,
    pub id: Option<u32>,
    pub format_string: String,
    pub is_date_format: bool,
}

impl From<&CoreNumberFormat> for WasmNumberFormat {
    fn from(n: &CoreNumberFormat) -> Self {
        Self {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmCellProtection {
    pub locked: bool,
    pub hidden: bool,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmStyle {
    pub font: WasmFontStyle,
    pub fill: WasmFillStyle,
    pub border: WasmBorderStyle,
    pub alignment: WasmAlignment,
    pub number_format: WasmNumberFormat,
    pub protection: WasmCellProtection,
}

impl From<&CoreStyle> for WasmStyle {
    fn from(s: &CoreStyle) -> Self {
        Self {
            font: WasmFontStyle::from(&s.font),
            fill: WasmFillStyle::from(&s.fill),
            border: WasmBorderStyle::from(&s.border),
            alignment: WasmAlignment::from(&s.alignment),
            number_format: WasmNumberFormat::from(&s.number_format),
            protection: WasmCellProtection {
                locked: s.protection.locked,
                hidden: s.protection.hidden,
            },
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmHyperlink {
    pub target: String,
    pub display: Option<String>,
    pub tooltip: Option<String>,
    pub location: Option<String>,
}

impl From<&core::Hyperlink> for WasmHyperlink {
    fn from(h: &core::Hyperlink) -> Self {
        Self {
            target: h.target.clone(),
            display: h.display.clone(),
            tooltip: h.tooltip.clone(),
            location: h.location.clone(),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmComment {
    pub author: String,
    pub text: String,
    pub visible: bool,
}

impl From<&core::CellComment> for WasmComment {
    fn from(c: &core::CellComment) -> Self {
        Self {
            author: c.author.clone(),
            text: c.text.clone(),
            visible: c.visible,
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmCommentEntry {
    pub row: u32,
    pub col: u32,
    pub comment: WasmComment,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmFreezePanes {
    pub row: u32,
    pub col: u32,
}

impl From<&core::FreezePanes> for WasmFreezePanes {
    fn from(f: &core::FreezePanes) -> Self {
        Self {
            row: f.row,
            col: f.col as u32,
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmSplitPanes {
    pub x_split: f64,
    pub y_split: f64,
    pub top_left_row: Option<u32>,
    pub top_left_col: Option<u32>,
    pub active_pane: Option<String>,
}

impl From<&core::SplitPanes> for WasmSplitPanes {
    fn from(s: &core::SplitPanes) -> Self {
        Self {
            x_split: s.x_split,
            y_split: s.y_split,
            top_left_row: s.top_left.map(|(r, _)| r),
            top_left_col: s.top_left.map(|(_, c)| c as u32),
            active_pane: s.active_pane.clone(),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmSelection {
    pub pane: Option<String>,
    pub active_cell: Option<String>,
    pub sqref: Option<String>,
}

impl From<&core::Selection> for WasmSelection {
    fn from(s: &core::Selection) -> Self {
        Self {
            pane: s.pane.clone(),
            active_cell: s.active_cell.clone(),
            sqref: s.sqref.clone(),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmSheetProtection {
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

impl From<&core::SheetProtection> for WasmSheetProtection {
    fn from(p: &core::SheetProtection) -> Self {
        Self {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPageSetup {
    pub paper_size: u32,
    pub orientation: String,
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

impl From<&core::PageSetup> for WasmPageSetup {
    fn from(p: &core::PageSetup) -> Self {
        Self {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPageBreak {
    pub id: u32,
    pub min: u32,
    pub max: u32,
    pub manual: bool,
}

impl From<&core::PageBreak> for WasmPageBreak {
    fn from(b: &core::PageBreak) -> Self {
        Self {
            id: b.id,
            min: b.min,
            max: b.max,
            manual: b.man,
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmWorkbookSettings {
    pub date_1904: bool,
    pub protected: bool,
    pub calc_on_open: bool,
    pub theme: Option<String>,
}

impl From<&core::WorkbookSettings> for WasmWorkbookSettings {
    fn from(s: &core::WorkbookSettings) -> Self {
        Self {
            date_1904: s.date_1904,
            protected: s.protected,
            calc_on_open: s.calc_on_open,
            theme: s.theme.clone(),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmNamedRange {
    pub name: String,
    pub scope: String,
    pub sheet_index: Option<u32>,
    pub refers_to: String,
    pub comment: Option<String>,
    pub hidden: bool,
}

impl From<&NamedRange> for WasmNamedRange {
    fn from(nr: &NamedRange) -> Self {
        Self {
            name: nr.name.clone(),
            scope: match &nr.scope {
                NameScope::Workbook => "workbook".into(),
                NameScope::Sheet(_) => "sheet".into(),
            },
            sheet_index: match &nr.scope {
                NameScope::Workbook => None,
                NameScope::Sheet(idx) => Some(*idx as u32),
            },
            refers_to: nr.refers_to.clone(),
            comment: nr.comment.clone(),
            hidden: nr.hidden,
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmTable {
    pub id: u32,
    pub name: String,
    pub display_name: String,
    pub reference: String,
    pub columns: Vec<WasmTableColumn>,
    pub style_info: Option<WasmTableStyleInfo>,
    pub header_row_count: u32,
    pub totals_row_count: u32,
    pub totals_row_shown: bool,
}

impl From<&core::Table> for WasmTable {
    fn from(t: &core::Table) -> Self {
        Self {
            id: t.id,
            name: t.name.clone(),
            display_name: t.display_name.clone(),
            reference: t.reference.to_string(),
            columns: t.columns.iter().map(WasmTableColumn::from).collect(),
            style_info: t.style_info.as_ref().map(WasmTableStyleInfo::from),
            header_row_count: t.header_row_count,
            totals_row_count: t.totals_row_count,
            totals_row_shown: t.totals_row_shown,
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmTableColumn {
    pub id: u32,
    pub name: String,
    pub totals_row_function: Option<String>,
    pub totals_row_formula: Option<String>,
    pub totals_row_label: Option<String>,
    pub calculated_column_formula: Option<String>,
}

impl From<&core::TableColumn> for WasmTableColumn {
    fn from(c: &core::TableColumn) -> Self {
        Self {
            id: c.id,
            name: c.name.clone(),
            totals_row_function: c.totals_row_function.as_ref().map(|f| f.to_ooxml().into()),
            totals_row_formula: c.totals_row_formula.clone(),
            totals_row_label: c.totals_row_label.clone(),
            calculated_column_formula: c.calculated_column_formula.clone(),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmTableStyleInfo {
    pub name: Option<String>,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
}

impl From<&core::TableStyleInfo> for WasmTableStyleInfo {
    fn from(s: &core::TableStyleInfo) -> Self {
        Self {
            name: s.name.clone(),
            show_first_column: s.show_first_column,
            show_last_column: s.show_last_column,
            show_row_stripes: s.show_row_stripes,
            show_column_stripes: s.show_column_stripes,
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmAutoFilter {
    pub range: String,
    pub filter_columns: Vec<WasmFilterColumn>,
}

impl From<&core::AutoFilter> for WasmAutoFilter {
    fn from(af: &core::AutoFilter) -> Self {
        Self {
            range: af.range.to_string(),
            filter_columns: af
                .filter_columns
                .iter()
                .map(WasmFilterColumn::from)
                .collect(),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmFilterColumn {
    pub col_id: u32,
    pub hidden_button: bool,
    pub show_button: bool,
    pub filter_type: String,
    pub values: Option<Vec<String>>,
    pub blank: Option<bool>,
}

impl From<&core::FilterColumn> for WasmFilterColumn {
    fn from(fc: &core::FilterColumn) -> Self {
        let (filter_type, values, blank) = match &fc.filter {
            core::ColumnFilter::Values(vf) => ("values", Some(vf.values.clone()), Some(vf.blank)),
            core::ColumnFilter::Custom(_) => ("custom", None, None),
            core::ColumnFilter::Top10(_) => ("top10", None, None),
            core::ColumnFilter::Dynamic(_) => ("dynamic", None, None),
            core::ColumnFilter::Color(_) => ("color", None, None),
        };
        Self {
            col_id: fc.col_id,
            hidden_button: fc.hidden_button,
            show_button: fc.show_button,
            filter_type: filter_type.into(),
            values,
            blank,
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmDataValidation {
    pub validation_type: String,
    pub ranges: Vec<String>,
    pub allow_blank: bool,
    pub show_dropdown: bool,
    pub show_input_message: bool,
    pub input_title: Option<String>,
    pub input_message: Option<String>,
    pub show_error_alert: bool,
    pub error_style: String,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
    pub operator: Option<String>,
    pub value1: Option<String>,
    pub value2: Option<String>,
    pub list_source: Option<String>,
    pub formula: Option<String>,
}

pub(crate) fn validation_operator_to_string(op: &core::ValidationOperator) -> &'static str {
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

impl From<&core::DataValidation> for WasmDataValidation {
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

        Self {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmConditionalFormatRule {
    pub rule_type: String,
    pub ranges: Vec<String>,
    pub priority: u32,
    pub stop_if_true: bool,
    pub operator: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub text: Option<String>,
    pub rank: Option<u32>,
    pub percent: Option<bool>,
    pub bottom: Option<bool>,
}

impl From<&core::ConditionalFormatRule> for WasmConditionalFormatRule {
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

        Self {
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmRichTextRun {
    pub text: String,
    pub font: Option<WasmRunFont>,
}

impl From<&core::RichTextRun> for WasmRichTextRun {
    fn from(r: &core::RichTextRun) -> Self {
        Self {
            text: r.text.clone(),
            font: r.font.as_ref().map(WasmRunFont::from),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmRunFont {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub size: Option<f64>,
    pub color: Option<WasmColor>,
    pub name: Option<String>,
    pub underline: Option<String>,
    pub strikethrough: Option<bool>,
    pub vertical_align: Option<String>,
}

impl From<&core::RunFont> for WasmRunFont {
    fn from(f: &core::RunFont) -> Self {
        Self {
            bold: f.bold,
            italic: f.italic,
            size: f.size,
            color: f.color.as_ref().map(WasmColor::from),
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmHyperlinkEntry {
    pub address: String,
    pub hyperlink: WasmHyperlink,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmFormulaCell {
    pub row: u32,
    pub col: u32,
    pub formula: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmSpillSource {
    pub row: u32,
    pub col: u32,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmMergedRegion {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
    /// The range as an A1-style string (e.g., "A1:C3").
    pub range: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmMergeSpan {
    pub row_span: u32,
    pub col_span: u32,
}


#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartAnchor {
    pub from_col: u16,
    pub from_row: u32,
    pub from_col_offset: i64,
    pub from_row_offset: i64,
    pub to_col: u16,
    pub to_row: u32,
    pub to_col_offset: i64,
    pub to_row_offset: i64,
}

impl From<&duke_sheets_chart::ChartAnchor> for WasmChartAnchor {
    fn from(a: &duke_sheets_chart::ChartAnchor) -> Self {
        Self {
            from_col: a.from_col,
            from_row: a.from_row,
            from_col_offset: a.from_col_offset,
            from_row_offset: a.from_row_offset,
            to_col: a.to_col,
            to_row: a.to_row,
            to_col_offset: a.to_col_offset,
            to_row_offset: a.to_row_offset,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmDataReference {
    pub ref_type: String,
    pub formula: Option<String>,
    pub numbers: Option<Vec<f64>>,
    pub strings: Option<Vec<String>>,
}

impl From<&duke_sheets_chart::DataReference> for WasmDataReference {
    fn from(r: &duke_sheets_chart::DataReference) -> Self {
        match r {
            duke_sheets_chart::DataReference::Formula(f) => Self {
                ref_type: "formula".into(),
                formula: Some(f.clone()),
                numbers: None,
                strings: None,
            },
            duke_sheets_chart::DataReference::Numbers(n) => Self {
                ref_type: "numbers".into(),
                formula: None,
                numbers: Some(n.clone()),
                strings: None,
            },
            duke_sheets_chart::DataReference::Strings(s) => Self {
                ref_type: "strings".into(),
                formula: None,
                numbers: None,
                strings: Some(s.clone()),
            },
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartNumberFormat {
    pub format_code: String,
    pub source_linked: Option<bool>,
}

impl From<&duke_sheets_chart::NumberFormat> for WasmChartNumberFormat {
    fn from(n: &duke_sheets_chart::NumberFormat) -> Self {
        Self {
            format_code: n.format_code.clone(),
            source_linked: n.source_linked,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartShapeProperties {
    pub solid_fill_hex: Option<String>,
    pub no_fill: bool,
    pub line_width: Option<i64>,
    pub line_color_hex: Option<String>,
    pub line_no_fill: bool,
    pub line_dash_style: Option<String>,
}

impl From<&duke_sheets_chart::ChartShapeProperties> for WasmChartShapeProperties {
    fn from(sp: &duke_sheets_chart::ChartShapeProperties) -> Self {
        Self {
            solid_fill_hex: sp.solid_fill.as_ref().map(|c| c.hex.clone()),
            no_fill: sp.no_fill,
            line_width: sp.line.as_ref().and_then(|l| l.width),
            line_color_hex: sp.line.as_ref().and_then(|l| l.solid_fill.as_ref().map(|c| c.hex.clone())),
            line_no_fill: sp.line.as_ref().map(|l| l.no_fill).unwrap_or(false),
            line_dash_style: sp.line.as_ref().and_then(|l| l.dash_style.clone()),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmDataLabels {
    pub show_legend_key: Option<bool>,
    pub show_value: Option<bool>,
    pub show_category_name: Option<bool>,
    pub show_series_name: Option<bool>,
    pub show_percent: Option<bool>,
    pub show_bubble_size: Option<bool>,
    pub separator: Option<String>,
    pub position: Option<String>,
    pub number_format: Option<WasmChartNumberFormat>,
    pub show_leader_lines: Option<bool>,
}

impl From<&duke_sheets_chart::DataLabels> for WasmDataLabels {
    fn from(d: &duke_sheets_chart::DataLabels) -> Self {
        use duke_sheets_chart::DataLabelPosition;
        Self {
            show_legend_key: d.show_legend_key,
            show_value: d.show_value,
            show_category_name: d.show_category_name,
            show_series_name: d.show_series_name,
            show_percent: d.show_percent,
            show_bubble_size: d.show_bubble_size,
            separator: d.separator.clone(),
            position: d.position.as_ref().map(|p| {
                match p {
                    DataLabelPosition::BestFit => "bestFit",
                    DataLabelPosition::Bottom => "bottom",
                    DataLabelPosition::Center => "center",
                    DataLabelPosition::InsideBase => "insideBase",
                    DataLabelPosition::InsideEnd => "insideEnd",
                    DataLabelPosition::Left => "left",
                    DataLabelPosition::OutsideEnd => "outsideEnd",
                    DataLabelPosition::Right => "right",
                    DataLabelPosition::Top => "top",
                }
                .into()
            }),
            number_format: d.number_format.as_ref().map(WasmChartNumberFormat::from),
            show_leader_lines: d.show_leader_lines,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmTrendline {
    pub trendline_type: String,
    pub name: Option<String>,
    pub order: Option<u32>,
    pub period: Option<u32>,
    pub forward: Option<f64>,
    pub backward: Option<f64>,
    pub intercept: Option<f64>,
    pub display_r_squared: Option<bool>,
    pub display_equation: Option<bool>,
}

impl From<&duke_sheets_chart::Trendline> for WasmTrendline {
    fn from(t: &duke_sheets_chart::Trendline) -> Self {
        use duke_sheets_chart::TrendlineType;
        Self {
            trendline_type: match t.trendline_type {
                TrendlineType::Linear => "linear",
                TrendlineType::Exponential => "exponential",
                TrendlineType::Logarithmic => "logarithmic",
                TrendlineType::MovingAverage => "movingAverage",
                TrendlineType::Polynomial => "polynomial",
                TrendlineType::Power => "power",
            }
            .into(),
            name: t.name.clone(),
            order: t.order,
            period: t.period,
            forward: t.forward,
            backward: t.backward,
            intercept: t.intercept,
            display_r_squared: t.display_r_squared,
            display_equation: t.display_equation,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmErrorBars {
    pub direction: String,
    pub bar_type: String,
    pub value_type: String,
    pub value: Option<f64>,
    pub no_end_cap: Option<bool>,
}

impl From<&duke_sheets_chart::ErrorBars> for WasmErrorBars {
    fn from(e: &duke_sheets_chart::ErrorBars) -> Self {
        use duke_sheets_chart::{ErrorBarDirection, ErrorBarType, ErrorValueType};
        Self {
            direction: match e.direction {
                ErrorBarDirection::X => "x",
                ErrorBarDirection::Y => "y",
            }
            .into(),
            bar_type: match e.bar_type {
                ErrorBarType::Both => "both",
                ErrorBarType::Minus => "minus",
                ErrorBarType::Plus => "plus",
            }
            .into(),
            value_type: match e.value_type {
                ErrorValueType::Custom => "custom",
                ErrorValueType::FixedValue => "fixedValue",
                ErrorValueType::Percentage => "percentage",
                ErrorValueType::StandardDeviation => "standardDeviation",
                ErrorValueType::StandardError => "standardError",
            }
            .into(),
            value: e.value,
            no_end_cap: e.no_end_cap,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmMarker {
    pub symbol: Option<String>,
    pub size: Option<u8>,
}

impl From<&duke_sheets_chart::Marker> for WasmMarker {
    fn from(m: &duke_sheets_chart::Marker) -> Self {
        use duke_sheets_chart::MarkerSymbol;
        Self {
            symbol: m.symbol.as_ref().map(|s| {
                match s {
                    MarkerSymbol::Circle => "circle",
                    MarkerSymbol::Dash => "dash",
                    MarkerSymbol::Diamond => "diamond",
                    MarkerSymbol::Dot => "dot",
                    MarkerSymbol::None => "none",
                    MarkerSymbol::Picture => "picture",
                    MarkerSymbol::Plus => "plus",
                    MarkerSymbol::Square => "square",
                    MarkerSymbol::Star => "star",
                    MarkerSymbol::Triangle => "triangle",
                    MarkerSymbol::X => "x",
                    MarkerSymbol::Auto => "auto",
                }
                .into()
            }),
            size: m.size,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmDataPoint {
    pub index: u32,
    pub marker: Option<WasmMarker>,
    pub explosion: Option<u32>,
}

impl From<&duke_sheets_chart::DataPoint> for WasmDataPoint {
    fn from(dp: &duke_sheets_chart::DataPoint) -> Self {
        Self {
            index: dp.index,
            marker: dp.marker.as_ref().map(WasmMarker::from),
            explosion: dp.explosion,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmView3D {
    pub rotate_x: Option<i32>,
    pub rotate_y: Option<i32>,
    pub depth_percent: Option<u32>,
    pub height_percent: Option<u32>,
    pub perspective: Option<u32>,
    pub right_angle_axes: Option<bool>,
}

impl From<&duke_sheets_chart::View3D> for WasmView3D {
    fn from(v: &duke_sheets_chart::View3D) -> Self {
        Self {
            rotate_x: v.rotate_x,
            rotate_y: v.rotate_y,
            depth_percent: v.depth_percent,
            height_percent: v.height_percent,
            perspective: v.perspective,
            right_angle_axes: v.right_angle_axes,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartDataTable {
    pub show_horizontal_border: Option<bool>,
    pub show_vertical_border: Option<bool>,
    pub show_outline: Option<bool>,
    pub show_keys: Option<bool>,
}

impl From<&duke_sheets_chart::ChartDataTable> for WasmChartDataTable {
    fn from(dt: &duke_sheets_chart::ChartDataTable) -> Self {
        Self {
            show_horizontal_border: dt.show_horizontal_border,
            show_vertical_border: dt.show_vertical_border,
            show_outline: dt.show_outline,
            show_keys: dt.show_keys,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmManualLayout {
    pub x: Option<f64>,
    pub y: Option<f64>,
    pub width: Option<f64>,
    pub height: Option<f64>,
}

impl From<&duke_sheets_chart::ManualLayout> for WasmManualLayout {
    fn from(ml: &duke_sheets_chart::ManualLayout) -> Self {
        Self {
            x: ml.x,
            y: ml.y,
            width: ml.width,
            height: ml.height,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmLayout {
    pub manual_layout: Option<WasmManualLayout>,
}

impl From<&duke_sheets_chart::Layout> for WasmLayout {
    fn from(l: &duke_sheets_chart::Layout) -> Self {
        Self {
            manual_layout: l.manual_layout.as_ref().map(WasmManualLayout::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmDataSeries {
    pub name: Option<String>,
    pub values: WasmDataReference,
    pub categories: Option<WasmDataReference>,
    pub data_labels: Option<WasmDataLabels>,
    pub trendline: Option<WasmTrendline>,
    pub error_bars: Option<WasmErrorBars>,
    pub marker: Option<WasmMarker>,
    pub data_points: Vec<WasmDataPoint>,
    pub smooth: Option<bool>,
    pub explosion: Option<u32>,
    pub invert_if_negative: Option<bool>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::DataSeries> for WasmDataSeries {
    fn from(s: &duke_sheets_chart::DataSeries) -> Self {
        Self {
            name: s.name.clone(),
            values: WasmDataReference::from(&s.values),
            categories: s.categories.as_ref().map(WasmDataReference::from),
            data_labels: s.data_labels.as_ref().map(WasmDataLabels::from),
            trendline: s.trendline.as_ref().map(WasmTrendline::from),
            error_bars: s.error_bars.as_ref().map(WasmErrorBars::from),
            marker: s.marker.as_ref().map(WasmMarker::from),
            data_points: s.data_points.iter().map(WasmDataPoint::from).collect(),
            smooth: s.smooth,
            explosion: s.explosion,
            invert_if_negative: s.invert_if_negative,
            shape_properties: s.shape_properties.as_ref().map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmAxis {
    pub title: Option<String>,
    pub minimum: Option<f64>,
    pub maximum: Option<f64>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
    pub position: String,
    pub number_format: Option<WasmChartNumberFormat>,
    pub major_gridlines: bool,
    pub minor_gridlines: bool,
    pub major_tick_mark: Option<String>,
    pub minor_tick_mark: Option<String>,
    pub label_position: Option<String>,
    pub delete: Option<bool>,
    pub crosses: Option<String>,
    pub cross_between: Option<String>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::Axis> for WasmAxis {
    fn from(a: &duke_sheets_chart::Axis) -> Self {
        use duke_sheets_chart::{AxisCrosses, CrossBetween, TickLabelPosition, TickMark};
        Self {
            title: a.title.clone(),
            minimum: a.minimum,
            maximum: a.maximum,
            major_unit: a.major_unit,
            minor_unit: a.minor_unit,
            position: match a.position {
                duke_sheets_chart::AxisPosition::Bottom => "bottom",
                duke_sheets_chart::AxisPosition::Top => "top",
                duke_sheets_chart::AxisPosition::Left => "left",
                duke_sheets_chart::AxisPosition::Right => "right",
            }
            .into(),
            number_format: a.number_format.as_ref().map(WasmChartNumberFormat::from),
            major_gridlines: a.major_gridlines,
            minor_gridlines: a.minor_gridlines,
            major_tick_mark: a.major_tick_mark.as_ref().map(|t| {
                match t {
                    TickMark::Cross => "cross",
                    TickMark::Inside => "inside",
                    TickMark::None => "none",
                    TickMark::Outside => "outside",
                }
                .into()
            }),
            minor_tick_mark: a.minor_tick_mark.as_ref().map(|t| {
                match t {
                    TickMark::Cross => "cross",
                    TickMark::Inside => "inside",
                    TickMark::None => "none",
                    TickMark::Outside => "outside",
                }
                .into()
            }),
            label_position: a.label_position.as_ref().map(|p| {
                match p {
                    TickLabelPosition::High => "high",
                    TickLabelPosition::Low => "low",
                    TickLabelPosition::NextTo => "nextTo",
                    TickLabelPosition::None => "none",
                }
                .into()
            }),
            delete: a.delete,
            crosses: a.crosses.as_ref().map(|c| {
                match c {
                    AxisCrosses::AutoZero => "autoZero",
                    AxisCrosses::Min => "min",
                    AxisCrosses::Max => "max",
                }
                .into()
            }),
            cross_between: a.cross_between.as_ref().map(|cb| {
                match cb {
                    CrossBetween::Between => "between",
                    CrossBetween::MidCat => "midCat",
                }
                .into()
            }),
            shape_properties: a.shape_properties.as_ref().map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmLegend {
    pub position: String,
    pub overlay: bool,
}

impl From<&duke_sheets_chart::Legend> for WasmLegend {
    fn from(l: &duke_sheets_chart::Legend) -> Self {
        Self {
            position: match l.position {
                duke_sheets_chart::LegendPosition::Right => "right",
                duke_sheets_chart::LegendPosition::Top => "top",
                duke_sheets_chart::LegendPosition::Bottom => "bottom",
                duke_sheets_chart::LegendPosition::Left => "left",
                duke_sheets_chart::LegendPosition::TopRight => "topRight",
            }
            .into(),
            overlay: l.overlay,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartTypeGroup {
    pub chart_type: String,
    pub is_3d: bool,
    pub series: Vec<WasmDataSeries>,
    pub data_labels: Option<WasmDataLabels>,
    pub vary_colors: Option<bool>,
    pub gap_width: Option<u32>,
    pub overlap: Option<i32>,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
    pub bubble_scale: Option<u32>,
    pub show_negative_bubbles: Option<bool>,
    pub radar_style: Option<String>,
    pub wireframe: Option<bool>,
    pub axis_ids: Vec<u32>,
    pub drop_lines: Option<WasmChartLines>,
    pub high_low_lines: Option<WasmChartLines>,
    pub series_lines: Option<WasmChartLines>,
    pub up_down_bars: Option<WasmUpDownBars>,
}

impl From<&duke_sheets_chart::ChartTypeGroup> for WasmChartTypeGroup {
    fn from(g: &duke_sheets_chart::ChartTypeGroup) -> Self {
        Self {
            chart_type: format!("{:?}", g.chart_type),
            is_3d: g.is_3d,
            series: g.series.iter().map(WasmDataSeries::from).collect(),
            data_labels: g.data_labels.as_ref().map(WasmDataLabels::from),
            vary_colors: g.vary_colors,
            gap_width: g.gap_width,
            overlap: g.overlap,
            first_slice_angle: g.first_slice_angle,
            hole_size: g.hole_size,
            bubble_scale: g.bubble_scale,
            show_negative_bubbles: g.show_negative_bubbles,
            radar_style: g.radar_style.clone(),
            wireframe: g.wireframe,
            axis_ids: g.axis_ids.clone(),
            drop_lines: g.drop_lines.as_ref().map(WasmChartLines::from),
            high_low_lines: g.high_low_lines.as_ref().map(WasmChartLines::from),
            series_lines: g.series_lines.as_ref().map(WasmChartLines::from),
            up_down_bars: g.up_down_bars.as_ref().map(WasmUpDownBars::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartAxis {
    pub id: u32,
    pub cross_id: u32,
    pub axis: WasmAxis,
}

impl From<&duke_sheets_chart::ChartAxis> for WasmChartAxis {
    fn from(a: &duke_sheets_chart::ChartAxis) -> Self {
        Self {
            id: a.id,
            cross_id: a.cross_id,
            axis: WasmAxis::from(&a.axis),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChart {
    pub chart_type: String,
    pub title: Option<String>,
    pub series: Vec<WasmDataSeries>,
    pub category_axis: Option<WasmAxis>,
    pub value_axis: Option<WasmAxis>,
    pub legend: Option<WasmLegend>,
    pub anchor: WasmChartAnchor,
    pub data_labels: Option<WasmDataLabels>,
    pub view_3d: Option<WasmView3D>,
    pub data_table: Option<WasmChartDataTable>,
    pub display_blanks_as: Option<String>,
    pub plot_visible_only: Option<bool>,
    pub layout: Option<WasmLayout>,
    pub is_3d: bool,
    pub vary_colors: Option<bool>,
    pub gap_width: Option<u32>,
    pub overlap: Option<i32>,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
    pub bubble_scale: Option<u32>,
    pub show_negative_bubbles: Option<bool>,
    pub auto_title_deleted: Option<bool>,
    pub rounded_corners: Option<bool>,
    pub show_dlbls_over_max: Option<bool>,
    pub wireframe: Option<bool>,
    pub radar_style: Option<String>,
    pub type_groups: Vec<WasmChartTypeGroup>,
    pub axes: Vec<WasmChartAxis>,
    pub drop_lines: Option<WasmChartLines>,
    pub high_low_lines: Option<WasmChartLines>,
    pub series_lines: Option<WasmChartLines>,
    pub up_down_bars: Option<WasmUpDownBars>,
}

impl From<&duke_sheets_chart::Chart> for WasmChart {
    fn from(c: &duke_sheets_chart::Chart) -> Self {
        let chart_type = match &c.chart_type {
            duke_sheets_chart::ChartType::Unsupported(tag) => format!("Unsupported({})", tag),
            other => format!("{:?}", other),
        };
        Self {
            chart_type,
            title: c.title.clone(),
            series: c.series.iter().map(WasmDataSeries::from).collect(),
            category_axis: c.category_axis.as_ref().map(WasmAxis::from),
            value_axis: c.value_axis.as_ref().map(WasmAxis::from),
            legend: c.legend.as_ref().map(WasmLegend::from),
            anchor: WasmChartAnchor::from(&c.anchor),
            data_labels: c.data_labels.as_ref().map(WasmDataLabels::from),
            view_3d: c.view_3d.as_ref().map(WasmView3D::from),
            data_table: c.data_table.as_ref().map(WasmChartDataTable::from),
            display_blanks_as: c.display_blanks_as.as_ref().map(|d| {
                use duke_sheets_chart::DisplayBlanksAs;
                match d {
                    DisplayBlanksAs::Gap => "gap",
                    DisplayBlanksAs::Span => "span",
                    DisplayBlanksAs::Zero => "zero",
                }
                .into()
            }),
            plot_visible_only: c.plot_visible_only,
            layout: c.layout.as_ref().map(WasmLayout::from),
            is_3d: c.is_3d,
            vary_colors: c.vary_colors,
            gap_width: c.gap_width,
            overlap: c.overlap,
            first_slice_angle: c.first_slice_angle,
            hole_size: c.hole_size,
            bubble_scale: c.bubble_scale,
            show_negative_bubbles: c.show_negative_bubbles,
            auto_title_deleted: c.auto_title_deleted,
            rounded_corners: c.rounded_corners,
            show_dlbls_over_max: c.show_dlbls_over_max,
            wireframe: c.wireframe,
            radar_style: c.radar_style.clone(),
            type_groups: c.type_groups.iter().map(WasmChartTypeGroup::from).collect(),
            axes: c.axes.iter().map(WasmChartAxis::from).collect(),
            drop_lines: c.drop_lines.as_ref().map(WasmChartLines::from),
            high_low_lines: c.high_low_lines.as_ref().map(WasmChartLines::from),
            series_lines: c.series_lines.as_ref().map(WasmChartLines::from),
            up_down_bars: c.up_down_bars.as_ref().map(WasmUpDownBars::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartLines {
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartLines> for WasmChartLines {
    fn from(cl: &duke_sheets_chart::ChartLines) -> Self {
        Self {
            shape_properties: cl.shape_properties.as_ref().map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmUpDownBars {
    pub gap_width: Option<u32>,
    pub up_bars: Option<WasmChartLines>,
    pub down_bars: Option<WasmChartLines>,
}

impl From<&duke_sheets_chart::UpDownBars> for WasmUpDownBars {
    fn from(ud: &duke_sheets_chart::UpDownBars) -> Self {
        Self {
            gap_width: ud.gap_width,
            up_bars: ud.up_bars.as_ref().map(WasmChartLines::from),
            down_bars: ud.down_bars.as_ref().map(WasmChartLines::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartSheet {
    pub name: String,
    pub chart: WasmChart,
    pub visibility: String,
}

impl From<&core::ChartSheet> for WasmChartSheet {
    fn from(cs: &core::ChartSheet) -> Self {
        Self {
            name: cs.name.clone(),
            chart: WasmChart::from(&cs.chart),
            visibility: match cs.visibility {
                core::worksheet::SheetVisibility::Visible => "visible",
                core::worksheet::SheetVisibility::Hidden => "hidden",
                core::worksheet::SheetVisibility::VeryHidden => "veryHidden",
            }
            .into(),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmSheetSlot {
    pub slot_type: String,
    pub index: u32,
}

impl From<&core::SheetSlot> for WasmSheetSlot {
    fn from(slot: &core::SheetSlot) -> Self {
        match slot {
            core::SheetSlot::Worksheet(idx) => Self {
                slot_type: "worksheet".into(),
                index: *idx as u32,
            },
            core::SheetSlot::ChartSheet(idx) => Self {
                slot_type: "chartsheet".into(),
                index: *idx as u32,
            },
        }
    }
}