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

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotMeasureOptions {
    pub field: String,
    pub aggregate: Option<String>,
    pub name: Option<String>,
    pub show_as: Option<String>,
    pub base_field: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotFilterOptions {
    pub kind: Option<String>,
    pub field: String,
    pub items: Option<Vec<String>>,
    pub operator: Option<String>,
    pub text: Option<String>,
    pub measure: Option<WasmPivotMeasureOptions>,
    pub value: Option<f64>,
    pub n: Option<u32>,
    pub top: Option<bool>,
    pub percent: Option<bool>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotCalculatedFieldOptions {
    pub name: String,
    pub formula: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotGroupingOptions {
    pub field: String,
    pub kind: String,
    pub start: Option<f64>,
    pub end: Option<f64>,
    pub interval: Option<f64>,
    pub units: Option<Vec<String>>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotFieldOptions {
    pub field: String,
    pub sort: Option<String>,
    pub subtotal: Option<String>,
    pub show_empty_items: Option<bool>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotRefreshPolicyOptions {
    pub refresh_on_open: Option<bool>,
    pub preserve_formatting: Option<bool>,
    pub background_query: Option<bool>,
    pub missing_items_limit: Option<u32>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotLayoutOptions {
    pub kind: Option<String>,
    pub show_row_grand_totals: Option<bool>,
    pub show_column_grand_totals: Option<bool>,
    pub show_field_headers: Option<bool>,
    pub repeat_item_labels: Option<bool>,
    pub show_expand_collapse: Option<bool>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotStyleOptions {
    pub name: Option<String>,
    pub show_row_headers: Option<bool>,
    pub show_column_headers: Option<bool>,
    pub show_row_stripes: Option<bool>,
    pub show_column_stripes: Option<bool>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotTableOptions {
    pub name: String,
    pub source_range: Option<String>,
    pub source_sheet: Option<String>,
    pub table_name: Option<String>,
    pub target: String,
    pub rows: Option<Vec<String>>,
    pub columns: Option<Vec<String>>,
    pub pages: Option<Vec<String>>,
    pub row_fields: Option<Vec<WasmPivotFieldOptions>>,
    pub column_fields: Option<Vec<WasmPivotFieldOptions>>,
    pub page_fields: Option<Vec<WasmPivotFieldOptions>>,
    pub measures: Vec<WasmPivotMeasureOptions>,
    pub filters: Option<Vec<WasmPivotFilterOptions>>,
    pub calculated_fields: Option<Vec<WasmPivotCalculatedFieldOptions>>,
    pub groupings: Option<Vec<WasmPivotGroupingOptions>>,
    pub refresh_policy: Option<WasmPivotRefreshPolicyOptions>,
    pub layout: Option<WasmPivotLayoutOptions>,
    pub style: Option<WasmPivotStyleOptions>,
    pub overwrite_policy: Option<String>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmPivotRefreshStats {
    pub pivot_count: usize,
    pub pivots_refreshed: usize,
    pub source_rows: usize,
    pub output_cells: usize,
    pub cache_hits: usize,
    pub cache_misses: usize,
}

impl From<duke_sheets::PivotRefreshStats> for WasmPivotRefreshStats {
    fn from(stats: duke_sheets::PivotRefreshStats) -> Self {
        Self {
            pivot_count: stats.pivot_count,
            pivots_refreshed: stats.pivots_refreshed,
            source_rows: stats.source_rows,
            output_cells: stats.output_cells,
            cache_hits: stats.cache_hits,
            cache_misses: stats.cache_misses,
        }
    }
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

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmColorInput {
    pub color_type: Option<String>,
    pub hex: Option<String>,
    pub r: Option<u32>,
    pub g: Option<u32>,
    pub b: Option<u32>,
    pub a: Option<u32>,
    pub theme_index: Option<u32>,
    pub tint: Option<i32>,
    pub palette_index: Option<u32>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmFontStylePatch {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<String>,
    pub strikethrough: Option<bool>,
    pub color: Option<WasmColorInput>,
    pub vertical_align: Option<String>,
    pub family: Option<u32>,
    pub charset: Option<u32>,
    pub scheme: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmGradientStopInput {
    pub position: f64,
    pub color: WasmColorInput,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmFillStylePatch {
    pub fill_type: Option<String>,
    pub color: Option<WasmColorInput>,
    pub pattern: Option<String>,
    pub foreground: Option<WasmColorInput>,
    pub background: Option<WasmColorInput>,
    pub gradient_type: Option<String>,
    pub angle: Option<f64>,
    pub stops: Option<Vec<WasmGradientStopInput>>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmBorderEdgePatch {
    pub style: Option<String>,
    pub color: Option<WasmColorInput>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmBorderStylePatch {
    pub left: Option<WasmBorderEdgePatch>,
    pub right: Option<WasmBorderEdgePatch>,
    pub top: Option<WasmBorderEdgePatch>,
    pub bottom: Option<WasmBorderEdgePatch>,
    pub diagonal: Option<WasmBorderEdgePatch>,
    pub diagonal_direction: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmAlignmentPatch {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: Option<bool>,
    pub shrink_to_fit: Option<bool>,
    pub indent: Option<u32>,
    pub rotation: Option<i32>,
    pub reading_order: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmNumberFormatPatch {
    pub format_type: Option<String>,
    pub id: Option<u32>,
    pub format_string: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmCellProtectionPatch {
    pub locked: Option<bool>,
    pub hidden: Option<bool>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WasmStylePatch {
    pub font: Option<WasmFontStylePatch>,
    pub fill: Option<WasmFillStylePatch>,
    pub border: Option<WasmBorderStylePatch>,
    pub alignment: Option<WasmAlignmentPatch>,
    pub number_format: Option<WasmNumberFormatPatch>,
    pub protection: Option<WasmCellProtectionPatch>,
}

fn u32_to_u8(value: u32, field: &str) -> Result<u8, String> {
    u8::try_from(value).map_err(|_| format!("{field} must be between 0 and 255"))
}

fn i32_to_i8(value: i32, field: &str) -> Result<i8, String> {
    i8::try_from(value).map_err(|_| format!("{field} must be between -128 and 127"))
}

fn parse_color_hex(hex: &str) -> Result<CoreColor, String> {
    CoreColor::from_hex(hex).ok_or_else(|| {
        "color hex must be 6 or 8 hexadecimal characters, with optional # prefix".to_string()
    })
}

fn parse_rgb_hex(hex: &str) -> Result<CoreColor, String> {
    match parse_color_hex(hex)? {
        CoreColor::Rgb { r, g, b } => Ok(CoreColor::Rgb { r, g, b }),
        CoreColor::Argb { r, g, b, .. } => Ok(CoreColor::Rgb { r, g, b }),
        other => Ok(other),
    }
}

fn parse_argb_hex(hex: &str) -> Result<CoreColor, String> {
    match parse_color_hex(hex)? {
        CoreColor::Rgb { r, g, b } => Ok(CoreColor::Argb { a: 255, r, g, b }),
        CoreColor::Argb { a, r, g, b } => Ok(CoreColor::Argb { a, r, g, b }),
        other => Ok(other),
    }
}

impl WasmColorInput {
    fn to_core_color(&self) -> Result<CoreColor, String> {
        match self.color_type.as_deref() {
            Some("auto") => Ok(CoreColor::Auto),
            Some("rgb") => {
                if let Some(hex) = &self.hex {
                    parse_rgb_hex(hex)
                } else {
                    Ok(CoreColor::Rgb {
                        r: u32_to_u8(
                            self.r.ok_or_else(|| "rgb color requires r".to_string())?,
                            "r",
                        )?,
                        g: u32_to_u8(
                            self.g.ok_or_else(|| "rgb color requires g".to_string())?,
                            "g",
                        )?,
                        b: u32_to_u8(
                            self.b.ok_or_else(|| "rgb color requires b".to_string())?,
                            "b",
                        )?,
                    })
                }
            }
            Some("argb") => {
                if let Some(hex) = &self.hex {
                    parse_argb_hex(hex)
                } else {
                    Ok(CoreColor::Argb {
                        a: u32_to_u8(self.a.unwrap_or(255), "a")?,
                        r: u32_to_u8(
                            self.r.ok_or_else(|| "argb color requires r".to_string())?,
                            "r",
                        )?,
                        g: u32_to_u8(
                            self.g.ok_or_else(|| "argb color requires g".to_string())?,
                            "g",
                        )?,
                        b: u32_to_u8(
                            self.b.ok_or_else(|| "argb color requires b".to_string())?,
                            "b",
                        )?,
                    })
                }
            }
            Some("theme") => Ok(CoreColor::Theme {
                index: u32_to_u8(
                    self.theme_index
                        .ok_or_else(|| "theme color requires themeIndex".to_string())?,
                    "themeIndex",
                )?,
                tint: i32_to_i8(self.tint.unwrap_or(0), "tint")?,
            }),
            Some("indexed") => Ok(CoreColor::Indexed(u32_to_u8(
                self.palette_index
                    .ok_or_else(|| "indexed color requires paletteIndex".to_string())?,
                "paletteIndex",
            )?)),
            Some(other) => Err(format!("unknown colorType {other:?}")),
            None => {
                if let Some(hex) = &self.hex {
                    parse_color_hex(hex)
                } else if self.r.is_some() || self.g.is_some() || self.b.is_some() {
                    Ok(CoreColor::Rgb {
                        r: u32_to_u8(
                            self.r.ok_or_else(|| "rgb color requires r".to_string())?,
                            "r",
                        )?,
                        g: u32_to_u8(
                            self.g.ok_or_else(|| "rgb color requires g".to_string())?,
                            "g",
                        )?,
                        b: u32_to_u8(
                            self.b.ok_or_else(|| "rgb color requires b".to_string())?,
                            "b",
                        )?,
                    })
                } else if let Some(theme_index) = self.theme_index {
                    Ok(CoreColor::Theme {
                        index: u32_to_u8(theme_index, "themeIndex")?,
                        tint: i32_to_i8(self.tint.unwrap_or(0), "tint")?,
                    })
                } else if let Some(palette_index) = self.palette_index {
                    Ok(CoreColor::Indexed(u32_to_u8(
                        palette_index,
                        "paletteIndex",
                    )?))
                } else {
                    Err("color requires colorType, hex, rgb, themeIndex, or paletteIndex".into())
                }
            }
        }
    }
}

fn parse_underline_input(value: &str) -> Result<Underline, String> {
    match value {
        "none" => Ok(Underline::None),
        "single" => Ok(Underline::Single),
        "double" => Ok(Underline::Double),
        "singleAccounting" => Ok(Underline::SingleAccounting),
        "doubleAccounting" => Ok(Underline::DoubleAccounting),
        other => Err(format!("unknown underline {other:?}")),
    }
}

fn parse_font_vertical_align_input(value: &str) -> Result<FontVerticalAlign, String> {
    match value {
        "baseline" => Ok(FontVerticalAlign::Baseline),
        "superscript" => Ok(FontVerticalAlign::Superscript),
        "subscript" => Ok(FontVerticalAlign::Subscript),
        other => Err(format!("unknown verticalAlign {other:?}")),
    }
}

impl WasmFontStylePatch {
    fn is_full_font(&self) -> bool {
        self.name.is_some()
            && self.size.is_some()
            && self.bold.is_some()
            && self.italic.is_some()
            && self.underline.is_some()
            && self.strikethrough.is_some()
            && self.color.is_some()
            && self.vertical_align.is_some()
    }

    fn apply_to_core_font(&self, font: &mut CoreFontStyle) -> Result<(), String> {
        if let Some(name) = &self.name {
            font.name = name.clone();
        }
        if let Some(size) = self.size {
            font.size = size;
        }
        if let Some(bold) = self.bold {
            font.bold = bold;
        }
        if let Some(italic) = self.italic {
            font.italic = italic;
        }
        if let Some(underline) = &self.underline {
            font.underline = parse_underline_input(underline)?;
        }
        if let Some(strikethrough) = self.strikethrough {
            font.strikethrough = strikethrough;
        }
        if let Some(color) = &self.color {
            font.color = color.to_core_color()?;
        }
        if let Some(vertical_align) = &self.vertical_align {
            font.vertical_align = parse_font_vertical_align_input(vertical_align)?;
        }
        if let Some(family) = self.family {
            font.family = Some(u32_to_u8(family, "family")?);
        }
        if let Some(charset) = self.charset {
            font.charset = Some(u32_to_u8(charset, "charset")?);
        }
        if let Some(scheme) = &self.scheme {
            font.scheme = Some(scheme.clone());
        }
        Ok(())
    }
}

fn parse_pattern_type_input(value: &str) -> Result<PatternType, String> {
    match value {
        "none" => Ok(PatternType::None),
        "solid" => Ok(PatternType::Solid),
        "mediumGray" => Ok(PatternType::MediumGray),
        "darkGray" => Ok(PatternType::DarkGray),
        "lightGray" => Ok(PatternType::LightGray),
        "darkHorizontal" => Ok(PatternType::DarkHorizontal),
        "darkVertical" => Ok(PatternType::DarkVertical),
        "darkDown" => Ok(PatternType::DarkDown),
        "darkUp" => Ok(PatternType::DarkUp),
        "darkGrid" => Ok(PatternType::DarkGrid),
        "darkTrellis" => Ok(PatternType::DarkTrellis),
        "lightHorizontal" => Ok(PatternType::LightHorizontal),
        "lightVertical" => Ok(PatternType::LightVertical),
        "lightDown" => Ok(PatternType::LightDown),
        "lightUp" => Ok(PatternType::LightUp),
        "lightGrid" => Ok(PatternType::LightGrid),
        "lightTrellis" => Ok(PatternType::LightTrellis),
        "gray125" => Ok(PatternType::Gray125),
        "gray0625" => Ok(PatternType::Gray0625),
        other => Err(format!("unknown fill pattern {other:?}")),
    }
}

fn parse_gradient_type_input(value: &str) -> Result<GradientType, String> {
    match value {
        "linear" => Ok(GradientType::Linear),
        "path" => Ok(GradientType::Path),
        other => Err(format!("unknown gradientType {other:?}")),
    }
}

impl WasmFillStylePatch {
    fn to_core_fill(&self) -> Result<CoreFillStyle, String> {
        match self.fill_type.as_deref() {
            Some("none") => Ok(CoreFillStyle::None),
            Some("solid") | None if self.color.is_some() => Ok(CoreFillStyle::Solid {
                color: self
                    .color
                    .as_ref()
                    .ok_or_else(|| "solid fill requires color".to_string())?
                    .to_core_color()?,
            }),
            Some("pattern") => Ok(CoreFillStyle::Pattern {
                pattern: parse_pattern_type_input(
                    self.pattern
                        .as_deref()
                        .ok_or_else(|| "pattern fill requires pattern".to_string())?,
                )?,
                foreground: self
                    .foreground
                    .as_ref()
                    .ok_or_else(|| "pattern fill requires foreground".to_string())?
                    .to_core_color()?,
                background: self
                    .background
                    .as_ref()
                    .ok_or_else(|| "pattern fill requires background".to_string())?
                    .to_core_color()?,
            }),
            Some("gradient") => Ok(CoreFillStyle::Gradient {
                gradient_type: parse_gradient_type_input(
                    self.gradient_type.as_deref().unwrap_or("linear"),
                )?,
                angle: self.angle.unwrap_or(0.0),
                stops: self
                    .stops
                    .as_ref()
                    .ok_or_else(|| "gradient fill requires stops".to_string())?
                    .iter()
                    .map(|stop| {
                        Ok(duke_sheets_core::style::GradientStop {
                            position: stop.position,
                            color: stop.color.to_core_color()?,
                        })
                    })
                    .collect::<Result<Vec<_>, String>>()?,
            }),
            Some(other) => Err(format!("unknown fillType {other:?}")),
            None => Err("fill patch requires fillType or color".into()),
        }
    }
}

fn parse_border_line_style_input(value: &str) -> Result<CoreBorderLineStyle, String> {
    match value {
        "none" => Ok(CoreBorderLineStyle::None),
        "thin" => Ok(CoreBorderLineStyle::Thin),
        "medium" => Ok(CoreBorderLineStyle::Medium),
        "thick" => Ok(CoreBorderLineStyle::Thick),
        "dashed" => Ok(CoreBorderLineStyle::Dashed),
        "dotted" => Ok(CoreBorderLineStyle::Dotted),
        "double" => Ok(CoreBorderLineStyle::Double),
        "hair" => Ok(CoreBorderLineStyle::Hair),
        "mediumDashed" => Ok(CoreBorderLineStyle::MediumDashed),
        "dashDot" => Ok(CoreBorderLineStyle::DashDot),
        "mediumDashDot" => Ok(CoreBorderLineStyle::MediumDashDot),
        "dashDotDot" => Ok(CoreBorderLineStyle::DashDotDot),
        "mediumDashDotDot" => Ok(CoreBorderLineStyle::MediumDashDotDot),
        "slantDashDot" => Ok(CoreBorderLineStyle::SlantDashDot),
        other => Err(format!("unknown border style {other:?}")),
    }
}

fn parse_diagonal_direction_input(value: &str) -> Result<DiagonalDirection, String> {
    match value {
        "none" => Ok(DiagonalDirection::None),
        "down" => Ok(DiagonalDirection::Down),
        "up" => Ok(DiagonalDirection::Up),
        "both" => Ok(DiagonalDirection::Both),
        other => Err(format!("unknown diagonalDirection {other:?}")),
    }
}

impl WasmBorderEdgePatch {
    fn apply_to_edge(
        &self,
        existing: Option<&CoreBorderEdge>,
    ) -> Result<Option<CoreBorderEdge>, String> {
        let parsed_style = self
            .style
            .as_deref()
            .map(parse_border_line_style_input)
            .transpose()?;
        if parsed_style == Some(CoreBorderLineStyle::None) {
            return Ok(None);
        }

        let mut edge = existing
            .cloned()
            .unwrap_or_else(|| CoreBorderEdge::new(CoreBorderLineStyle::Thin, CoreColor::BLACK));
        if let Some(style) = parsed_style {
            edge.style = style;
        }
        if let Some(color) = &self.color {
            edge.color = color.to_core_color()?;
        }
        Ok(Some(edge))
    }
}

impl WasmBorderStylePatch {
    fn is_full_border(&self) -> bool {
        self.diagonal_direction.is_some()
    }

    fn apply_to_core_border(&self, border: &mut CoreBorderStyle) -> Result<(), String> {
        if let Some(edge) = &self.left {
            border.left = edge.apply_to_edge(border.left.as_ref())?;
        }
        if let Some(edge) = &self.right {
            border.right = edge.apply_to_edge(border.right.as_ref())?;
        }
        if let Some(edge) = &self.top {
            border.top = edge.apply_to_edge(border.top.as_ref())?;
        }
        if let Some(edge) = &self.bottom {
            border.bottom = edge.apply_to_edge(border.bottom.as_ref())?;
        }
        if let Some(edge) = &self.diagonal {
            border.diagonal = edge.apply_to_edge(border.diagonal.as_ref())?;
        }
        if let Some(direction) = &self.diagonal_direction {
            border.diagonal_direction = parse_diagonal_direction_input(direction)?;
        }
        Ok(())
    }
}

fn parse_horizontal_alignment_input(value: &str) -> Result<HorizontalAlignment, String> {
    match value {
        "general" => Ok(HorizontalAlignment::General),
        "left" => Ok(HorizontalAlignment::Left),
        "center" => Ok(HorizontalAlignment::Center),
        "right" => Ok(HorizontalAlignment::Right),
        "fill" => Ok(HorizontalAlignment::Fill),
        "justify" => Ok(HorizontalAlignment::Justify),
        "centerContinuous" => Ok(HorizontalAlignment::CenterContinuous),
        "distributed" => Ok(HorizontalAlignment::Distributed),
        other => Err(format!("unknown horizontal alignment {other:?}")),
    }
}

fn parse_vertical_alignment_input(value: &str) -> Result<VerticalAlignment, String> {
    match value {
        "top" => Ok(VerticalAlignment::Top),
        "center" => Ok(VerticalAlignment::Center),
        "bottom" => Ok(VerticalAlignment::Bottom),
        "justify" => Ok(VerticalAlignment::Justify),
        "distributed" => Ok(VerticalAlignment::Distributed),
        other => Err(format!("unknown vertical alignment {other:?}")),
    }
}

fn parse_reading_order_input(value: &str) -> Result<ReadingOrder, String> {
    match value {
        "contextDependent" => Ok(ReadingOrder::ContextDependent),
        "leftToRight" => Ok(ReadingOrder::LeftToRight),
        "rightToLeft" => Ok(ReadingOrder::RightToLeft),
        other => Err(format!("unknown readingOrder {other:?}")),
    }
}

impl WasmAlignmentPatch {
    fn is_full_alignment(&self) -> bool {
        self.horizontal.is_some()
            && self.vertical.is_some()
            && self.wrap_text.is_some()
            && self.shrink_to_fit.is_some()
            && self.indent.is_some()
            && self.rotation.is_some()
            && self.reading_order.is_some()
    }

    fn apply_to_core_alignment(&self, alignment: &mut CoreAlignment) -> Result<(), String> {
        if let Some(horizontal) = &self.horizontal {
            alignment.horizontal = parse_horizontal_alignment_input(horizontal)?;
        }
        if let Some(vertical) = &self.vertical {
            alignment.vertical = parse_vertical_alignment_input(vertical)?;
        }
        if let Some(wrap_text) = self.wrap_text {
            alignment.wrap_text = wrap_text;
        }
        if let Some(shrink_to_fit) = self.shrink_to_fit {
            alignment.shrink_to_fit = shrink_to_fit;
        }
        if let Some(indent) = self.indent {
            alignment.indent = u32_to_u8(indent, "indent")?;
        }
        if let Some(rotation) = self.rotation {
            if !((-90..=90).contains(&rotation) || rotation == 255) {
                return Err("rotation must be between -90 and 90, or 255".into());
            }
            alignment.rotation = rotation as i16;
        }
        if let Some(reading_order) = &self.reading_order {
            alignment.reading_order = parse_reading_order_input(reading_order)?;
        }
        Ok(())
    }
}

impl WasmNumberFormatPatch {
    fn to_core_number_format(&self) -> Result<CoreNumberFormat, String> {
        match self.format_type.as_deref() {
            Some("general") => Ok(CoreNumberFormat::General),
            Some("builtin") => {
                Ok(CoreNumberFormat::BuiltIn(self.id.ok_or_else(|| {
                    "builtin number format requires id".to_string()
                })?))
            }
            Some("custom") => Ok(CoreNumberFormat::Custom(
                self.format_string
                    .clone()
                    .ok_or_else(|| "custom number format requires formatString".to_string())?,
            )),
            Some(other) => Err(format!("unknown formatType {other:?}")),
            None if self.id.is_some() => Ok(CoreNumberFormat::BuiltIn(self.id.unwrap())),
            None if self.format_string.is_some() => Ok(CoreNumberFormat::Custom(
                self.format_string.clone().unwrap(),
            )),
            None => Err("numberFormat requires formatType, id, or formatString".into()),
        }
    }
}

impl WasmCellProtectionPatch {
    fn apply_to_core_protection(&self, protection: &mut duke_sheets_core::style::Protection) {
        if let Some(locked) = self.locked {
            protection.locked = locked;
        }
        if let Some(hidden) = self.hidden {
            protection.hidden = hidden;
        }
    }
}

impl WasmStylePatch {
    pub fn apply_to_core_style(&self, style: &mut CoreStyle) -> Result<(), String> {
        if let Some(font_patch) = &self.font {
            if font_patch.is_full_font() {
                let mut font = CoreFontStyle::default();
                font_patch.apply_to_core_font(&mut font)?;
                style.font = font;
            } else {
                font_patch.apply_to_core_font(&mut style.font)?;
            }
        }

        if let Some(fill_patch) = &self.fill {
            style.fill = fill_patch.to_core_fill()?;
        }

        if let Some(border_patch) = &self.border {
            if border_patch.is_full_border() {
                let mut border = CoreBorderStyle::default();
                border_patch.apply_to_core_border(&mut border)?;
                style.border = border;
            } else {
                border_patch.apply_to_core_border(&mut style.border)?;
            }
        }

        if let Some(alignment_patch) = &self.alignment {
            if alignment_patch.is_full_alignment() {
                let mut alignment = CoreAlignment::default();
                alignment_patch.apply_to_core_alignment(&mut alignment)?;
                style.alignment = alignment;
            } else {
                alignment_patch.apply_to_core_alignment(&mut style.alignment)?;
            }
        }

        if let Some(number_format_patch) = &self.number_format {
            style.number_format = number_format_patch.to_core_number_format()?;
        }

        if let Some(protection_patch) = &self.protection {
            protection_patch.apply_to_core_protection(&mut style.protection);
        }

        Ok(())
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
pub struct WasmDrawingAnchor {
    pub from_col: u16,
    pub from_row: u32,
    pub from_col_offset: i64,
    pub from_row_offset: i64,
    pub to_col: u16,
    pub to_row: u32,
    pub to_col_offset: i64,
    pub to_row_offset: i64,
}

impl From<&duke_sheets_chart::DrawingAnchor> for WasmDrawingAnchor {
    fn from(a: &duke_sheets_chart::DrawingAnchor) -> Self {
        match a {
            duke_sheets_chart::DrawingAnchor::TwoCell { from, to, .. } => Self {
                from_col: from.col,
                from_row: from.row,
                from_col_offset: from.col_offset_emu,
                from_row_offset: from.row_offset_emu,
                to_col: to.col,
                to_row: to.row,
                to_col_offset: to.col_offset_emu,
                to_row_offset: to.row_offset_emu,
            },
            _ => Self {
                from_col: 0,
                from_row: 0,
                from_col_offset: 0,
                from_row_offset: 0,
                to_col: 0,
                to_row: 0,
                to_col_offset: 0,
                to_row_offset: 0,
            },
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
            line_color_hex: sp
                .line
                .as_ref()
                .and_then(|l| l.solid_fill.as_ref().map(|c| c.hex.clone())),
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
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::DataPoint> for WasmDataPoint {
    fn from(dp: &duke_sheets_chart::DataPoint) -> Self {
        Self {
            index: dp.index,
            marker: dp.marker.as_ref().map(WasmMarker::from),
            explosion: dp.explosion,
            shape_properties: dp
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
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
            shape_properties: s
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
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
    pub major_gridlines_shape_properties: Option<WasmChartShapeProperties>,
    pub minor_gridlines_shape_properties: Option<WasmChartShapeProperties>,
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
            major_gridlines_shape_properties: a
                .major_gridlines_shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
            minor_gridlines_shape_properties: a
                .minor_gridlines_shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
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
            shape_properties: a
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
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
pub struct WasmPivotChartSource {
    pub name: String,
    pub format_id: u32,
}

impl From<&duke_sheets_chart::PivotChartSource> for WasmPivotChartSource {
    fn from(s: &duke_sheets_chart::PivotChartSource) -> Self {
        Self {
            name: s.name.clone(),
            format_id: s.format_id,
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
    pub anchor: WasmDrawingAnchor,
    pub data_labels: Option<WasmDataLabels>,
    pub view_3d: Option<WasmView3D>,
    pub data_table: Option<WasmChartDataTable>,
    pub display_blanks_as: Option<String>,
    pub plot_visible_only: Option<bool>,
    pub layout: Option<WasmLayout>,
    pub shape_properties: Option<WasmChartShapeProperties>,
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
    pub pivot_source: Option<WasmPivotChartSource>,
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
            anchor: WasmDrawingAnchor::from(&c.anchor),
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
            shape_properties: c
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
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
            pivot_source: c.pivot_source.as_ref().map(WasmPivotChartSource::from),
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
            shape_properties: cl
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
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

fn chart_ex_layout_to_string(layout: &duke_sheets_chart::ChartExLayout) -> &'static str {
    match layout {
        duke_sheets_chart::ChartExLayout::Waterfall => "waterfall",
        duke_sheets_chart::ChartExLayout::Treemap => "treemap",
        duke_sheets_chart::ChartExLayout::Sunburst => "sunburst",
        duke_sheets_chart::ChartExLayout::Funnel => "funnel",
        duke_sheets_chart::ChartExLayout::Histogram => "histogram",
        duke_sheets_chart::ChartExLayout::BoxWhisker => "boxWhisker",
        duke_sheets_chart::ChartExLayout::ParetoLine => "paretoLine",
        duke_sheets_chart::ChartExLayout::RegionMap => "regionMap",
        duke_sheets_chart::ChartExLayout::ClusteredColumn => "clusteredColumn",
        duke_sheets_chart::ChartExLayout::Unknown(_) => "unknown",
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExOffset {
    pub top: Option<f64>,
    pub left: Option<f64>,
}

impl From<&duke_sheets_chart::ChartExOffset> for WasmChartExOffset {
    fn from(o: &duke_sheets_chart::ChartExOffset) -> Self {
        Self {
            top: o.top,
            left: o.left,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExText {
    pub formula: Option<String>,
    pub value: Option<String>,
}

impl From<&duke_sheets_chart::ChartExText> for WasmChartExText {
    fn from(t: &duke_sheets_chart::ChartExText) -> Self {
        Self {
            formula: t.data.as_ref().and_then(|d| d.formula.clone()),
            value: t.data.as_ref().and_then(|d| d.value.clone()),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExColorPosition {
    pub position_type: String,
    pub value: Option<f64>,
}

impl From<&duke_sheets_chart::ChartExColorPosition> for WasmChartExColorPosition {
    fn from(p: &duke_sheets_chart::ChartExColorPosition) -> Self {
        match p {
            duke_sheets_chart::ChartExColorPosition::ExtremeValue => Self {
                position_type: "extremeValue".into(),
                value: None,
            },
            duke_sheets_chart::ChartExColorPosition::Number(v) => Self {
                position_type: "number".into(),
                value: Some(*v),
            },
            duke_sheets_chart::ChartExColorPosition::Percent(v) => Self {
                position_type: "percent".into(),
                value: Some(*v),
            },
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExValueColorPositions {
    pub count: Option<u32>,
    pub min: Option<WasmChartExColorPosition>,
    pub mid: Option<WasmChartExColorPosition>,
    pub max: Option<WasmChartExColorPosition>,
}

impl From<&duke_sheets_chart::ChartExValueColorPositions> for WasmChartExValueColorPositions {
    fn from(p: &duke_sheets_chart::ChartExValueColorPositions) -> Self {
        Self {
            count: p.count,
            min: p.min.as_ref().map(WasmChartExColorPosition::from),
            mid: p.mid.as_ref().map(WasmChartExColorPosition::from),
            max: p.max.as_ref().map(WasmChartExColorPosition::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExScaling {
    pub scaling_type: String,
    pub gap_width: Option<f64>,
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
}

impl From<&duke_sheets_chart::ChartExScaling> for WasmChartExScaling {
    fn from(s: &duke_sheets_chart::ChartExScaling) -> Self {
        match s {
            duke_sheets_chart::ChartExScaling::Category { gap_width } => Self {
                scaling_type: "category".into(),
                gap_width: *gap_width,
                min: None,
                max: None,
                major_unit: None,
                minor_unit: None,
            },
            duke_sheets_chart::ChartExScaling::Value {
                min,
                max,
                major_unit,
                minor_unit,
            } => Self {
                scaling_type: "value".into(),
                gap_width: None,
                min: *min,
                max: *max,
                major_unit: *major_unit,
                minor_unit: *minor_unit,
            },
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExAxisTitle {
    pub text: Option<String>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExAxisTitle> for WasmChartExAxisTitle {
    fn from(t: &duke_sheets_chart::ChartExAxisTitle) -> Self {
        Self {
            text: t.text.as_ref().and_then(|tx| {
                tx.data
                    .as_ref()
                    .and_then(|d| d.value.clone().or_else(|| d.formula.clone()))
            }),
            shape_properties: t
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExAxisUnits {
    pub unit: Option<String>,
}

impl From<&duke_sheets_chart::ChartExAxisUnits> for WasmChartExAxisUnits {
    fn from(u: &duke_sheets_chart::ChartExAxisUnits) -> Self {
        Self {
            unit: u.unit.clone(),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExSeriesVisibility {
    pub connector_lines: Option<bool>,
    pub mean_line: Option<bool>,
    pub mean_marker: Option<bool>,
    pub nonoutliers: Option<bool>,
    pub outliers: Option<bool>,
}

impl From<&duke_sheets_chart::ChartExSeriesVisibility> for WasmChartExSeriesVisibility {
    fn from(v: &duke_sheets_chart::ChartExSeriesVisibility) -> Self {
        Self {
            connector_lines: v.connector_lines,
            mean_line: v.mean_line,
            mean_marker: v.mean_marker,
            nonoutliers: v.nonoutliers,
            outliers: v.outliers,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExBinning {
    pub interval_closed: Option<String>,
    pub underflow: Option<String>,
    pub overflow: Option<String>,
    pub bin_size: Option<f64>,
    pub bin_count: Option<u32>,
}

impl From<&duke_sheets_chart::ChartExBinning> for WasmChartExBinning {
    fn from(b: &duke_sheets_chart::ChartExBinning) -> Self {
        Self {
            interval_closed: b.interval_closed.clone(),
            underflow: b.underflow.clone(),
            overflow: b.overflow.clone(),
            bin_size: b.bin_size,
            bin_count: b.bin_count,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExGeography {
    pub projection_type: Option<String>,
    pub viewed_region_type: Option<String>,
    pub culture_language: Option<String>,
    pub culture_region: Option<String>,
    pub attribution: Option<String>,
}

impl From<&duke_sheets_chart::ChartExGeography> for WasmChartExGeography {
    fn from(g: &duke_sheets_chart::ChartExGeography) -> Self {
        Self {
            projection_type: g.projection_type.clone(),
            viewed_region_type: g.viewed_region_type.clone(),
            culture_language: g.culture_language.clone(),
            culture_region: g.culture_region.clone(),
            attribution: g.attribution.clone(),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExStatistics {
    pub quartile_method: Option<String>,
}

impl From<&duke_sheets_chart::ChartExStatistics> for WasmChartExStatistics {
    fn from(s: &duke_sheets_chart::ChartExStatistics) -> Self {
        Self {
            quartile_method: s.quartile_method.clone(),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExDataPoint {
    pub idx: u32,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExDataPoint> for WasmChartExDataPoint {
    fn from(p: &duke_sheets_chart::ChartExDataPoint) -> Self {
        Self {
            idx: p.idx,
            shape_properties: p
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExDataLabel {
    pub idx: u32,
    pub position: Option<String>,
    pub visibility_series_name: Option<bool>,
    pub visibility_category_name: Option<bool>,
    pub visibility_value: Option<bool>,
    pub number_format: Option<WasmChartNumberFormat>,
    pub separator: Option<String>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExDataLabel> for WasmChartExDataLabel {
    fn from(l: &duke_sheets_chart::ChartExDataLabel) -> Self {
        Self {
            idx: l.idx,
            position: l.position.clone(),
            visibility_series_name: l.visibility_series_name,
            visibility_category_name: l.visibility_category_name,
            visibility_value: l.visibility_value,
            number_format: l.number_format.as_ref().map(WasmChartNumberFormat::from),
            separator: l.separator.clone(),
            shape_properties: l
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExFormatOverride {
    pub idx: u32,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExFormatOverride> for WasmChartExFormatOverride {
    fn from(o: &duke_sheets_chart::ChartExFormatOverride) -> Self {
        Self {
            idx: o.idx,
            shape_properties: o
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExHeaderFooter {
    pub align_with_margins: Option<bool>,
    pub different_odd_even: Option<bool>,
    pub different_first: Option<bool>,
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
}

impl From<&duke_sheets_chart::ChartExHeaderFooter> for WasmChartExHeaderFooter {
    fn from(h: &duke_sheets_chart::ChartExHeaderFooter) -> Self {
        Self {
            align_with_margins: h.align_with_margins,
            different_odd_even: h.different_odd_even,
            different_first: h.different_first,
            odd_header: h.odd_header.clone(),
            odd_footer: h.odd_footer.clone(),
            even_header: h.even_header.clone(),
            even_footer: h.even_footer.clone(),
            first_header: h.first_header.clone(),
            first_footer: h.first_footer.clone(),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExPageMargins {
    pub left: Option<f64>,
    pub right: Option<f64>,
    pub top: Option<f64>,
    pub bottom: Option<f64>,
    pub header: Option<f64>,
    pub footer: Option<f64>,
}

impl From<&duke_sheets_chart::ChartExPageMargins> for WasmChartExPageMargins {
    fn from(m: &duke_sheets_chart::ChartExPageMargins) -> Self {
        Self {
            left: m.left,
            right: m.right,
            top: m.top,
            bottom: m.bottom,
            header: m.header,
            footer: m.footer,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExPageSetup {
    pub paper_size: Option<u32>,
    pub first_page_number: Option<u32>,
    pub orientation: Option<String>,
    pub black_and_white: Option<bool>,
    pub draft: Option<bool>,
    pub use_first_page_number: Option<bool>,
    pub horizontal_dpi: Option<u32>,
    pub vertical_dpi: Option<u32>,
    pub copies: Option<u32>,
}

impl From<&duke_sheets_chart::ChartExPageSetup> for WasmChartExPageSetup {
    fn from(p: &duke_sheets_chart::ChartExPageSetup) -> Self {
        Self {
            paper_size: p.paper_size,
            first_page_number: p.first_page_number,
            orientation: p.orientation.clone(),
            black_and_white: p.black_and_white,
            draft: p.draft,
            use_first_page_number: p.use_first_page_number,
            horizontal_dpi: p.horizontal_dpi,
            vertical_dpi: p.vertical_dpi,
            copies: p.copies,
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExPrintSettings {
    pub header_footer: Option<WasmChartExHeaderFooter>,
    pub page_margins: Option<WasmChartExPageMargins>,
    pub page_setup: Option<WasmChartExPageSetup>,
}

impl From<&duke_sheets_chart::ChartExPrintSettings> for WasmChartExPrintSettings {
    fn from(p: &duke_sheets_chart::ChartExPrintSettings) -> Self {
        Self {
            header_footer: p.header_footer.as_ref().map(WasmChartExHeaderFooter::from),
            page_margins: p.page_margins.as_ref().map(WasmChartExPageMargins::from),
            page_setup: p.page_setup.as_ref().map(WasmChartExPageSetup::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExPlotArea {
    pub plot_surface: Option<WasmChartShapeProperties>,
    pub series: Vec<WasmChartExSeries>,
    pub axes: Vec<WasmChartExAxis>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExPlotArea> for WasmChartExPlotArea {
    fn from(p: &duke_sheets_chart::ChartExPlotArea) -> Self {
        Self {
            plot_surface: p.plot_surface.as_ref().map(WasmChartShapeProperties::from),
            series: p.series.iter().map(WasmChartExSeries::from).collect(),
            axes: p.axes.iter().map(WasmChartExAxis::from).collect(),
            shape_properties: p
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExDimension {
    pub dim_type: String,
    pub formula: Option<String>,
    pub nf_formula: Option<String>,
}

impl From<&duke_sheets_chart::ChartExDimension> for WasmChartExDimension {
    fn from(d: &duke_sheets_chart::ChartExDimension) -> Self {
        match d {
            duke_sheets_chart::ChartExDimension::String {
                dim_type,
                formula,
                nf_formula,
                ..
            } => Self {
                dim_type: match dim_type {
                    duke_sheets_chart::StringDimType::Cat => "cat".into(),
                    duke_sheets_chart::StringDimType::ColorStr => "colorStr".into(),
                    duke_sheets_chart::StringDimType::EntityId => "entityId".into(),
                },
                formula: formula.clone(),
                nf_formula: nf_formula.clone(),
            },
            duke_sheets_chart::ChartExDimension::Numeric {
                dim_type,
                formula,
                nf_formula,
                ..
            } => Self {
                dim_type: match dim_type {
                    duke_sheets_chart::NumericDimType::Val => "val".into(),
                    duke_sheets_chart::NumericDimType::X => "x".into(),
                    duke_sheets_chart::NumericDimType::Y => "y".into(),
                    duke_sheets_chart::NumericDimType::Size => "size".into(),
                    duke_sheets_chart::NumericDimType::ColorVal => "colorVal".into(),
                },
                formula: formula.clone(),
                nf_formula: nf_formula.clone(),
            },
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExData {
    pub id: u32,
    pub dimensions: Vec<WasmChartExDimension>,
}

impl From<&duke_sheets_chart::ChartExData> for WasmChartExData {
    fn from(d: &duke_sheets_chart::ChartExData) -> Self {
        Self {
            id: d.id,
            dimensions: d
                .dimensions
                .iter()
                .map(WasmChartExDimension::from)
                .collect(),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExDataLabels {
    pub position: Option<String>,
    pub visibility_series_name: Option<bool>,
    pub visibility_category_name: Option<bool>,
    pub visibility_value: Option<bool>,
    pub number_format: Option<WasmChartNumberFormat>,
    pub separator: Option<String>,
    pub shape_properties: Option<WasmChartShapeProperties>,
    pub overrides: Vec<WasmChartExDataLabel>,
    pub hidden_labels: Vec<u32>,
}

impl From<&duke_sheets_chart::ChartExDataLabels> for WasmChartExDataLabels {
    fn from(l: &duke_sheets_chart::ChartExDataLabels) -> Self {
        Self {
            position: l.position.clone(),
            visibility_series_name: l.visibility_series_name,
            visibility_category_name: l.visibility_category_name,
            visibility_value: l.visibility_value,
            number_format: l.number_format.as_ref().map(WasmChartNumberFormat::from),
            separator: l.separator.clone(),
            shape_properties: l
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
            overrides: l.overrides.iter().map(WasmChartExDataLabel::from).collect(),
            hidden_labels: l.hidden_labels.clone(),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExTitle {
    pub text: Option<String>,
    pub position: Option<String>,
    pub align: Option<String>,
    pub overlay: Option<bool>,
    pub offset: Option<WasmChartExOffset>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExTitle> for WasmChartExTitle {
    fn from(t: &duke_sheets_chart::ChartExTitle) -> Self {
        Self {
            text: t.text.clone(),
            position: t.position.clone(),
            align: t.align.clone(),
            overlay: t.overlay,
            offset: t.offset.as_ref().map(WasmChartExOffset::from),
            shape_properties: t
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExLegend {
    pub position: Option<String>,
    pub align: Option<String>,
    pub overlay: Option<bool>,
    pub offset: Option<WasmChartExOffset>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExLegend> for WasmChartExLegend {
    fn from(l: &duke_sheets_chart::ChartExLegend) -> Self {
        Self {
            position: l.position.clone(),
            align: l.align.clone(),
            overlay: l.overlay,
            offset: l.offset.as_ref().map(WasmChartExOffset::from),
            shape_properties: l
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExLayoutPr {
    pub parent_label_layout: Option<String>,
    pub region_label_layout: Option<String>,
    pub visibility: Option<WasmChartExSeriesVisibility>,
    pub aggregation: bool,
    pub binning: Option<WasmChartExBinning>,
    pub geography: Option<WasmChartExGeography>,
    pub statistics: Option<WasmChartExStatistics>,
    pub subtotals: Vec<u32>,
}

impl From<&duke_sheets_chart::ChartExLayoutPr> for WasmChartExLayoutPr {
    fn from(l: &duke_sheets_chart::ChartExLayoutPr) -> Self {
        Self {
            parent_label_layout: l.parent_label_layout.clone(),
            region_label_layout: l.region_label_layout.clone(),
            visibility: l.visibility.as_ref().map(WasmChartExSeriesVisibility::from),
            aggregation: l.aggregation,
            binning: l.binning.as_ref().map(WasmChartExBinning::from),
            geography: l.geography.as_ref().map(WasmChartExGeography::from),
            statistics: l.statistics.as_ref().map(WasmChartExStatistics::from),
            subtotals: l.subtotals.clone(),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExAxis {
    pub id: u32,
    pub hidden: Option<bool>,
    pub scaling: WasmChartExScaling,
    pub title: Option<WasmChartExAxisTitle>,
    pub units: Option<WasmChartExAxisUnits>,
    pub major_gridlines: Option<WasmChartShapeProperties>,
    pub minor_gridlines: Option<WasmChartShapeProperties>,
    pub major_tick_marks: Option<String>,
    pub minor_tick_marks: Option<String>,
    pub tick_labels: bool,
    pub number_format: Option<WasmChartNumberFormat>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExAxis> for WasmChartExAxis {
    fn from(a: &duke_sheets_chart::ChartExAxis) -> Self {
        Self {
            id: a.id,
            hidden: a.hidden,
            scaling: WasmChartExScaling::from(&a.scaling),
            title: a.title.as_ref().map(WasmChartExAxisTitle::from),
            units: a.units.as_ref().map(WasmChartExAxisUnits::from),
            major_gridlines: a
                .major_gridlines
                .as_ref()
                .map(WasmChartShapeProperties::from),
            minor_gridlines: a
                .minor_gridlines
                .as_ref()
                .map(WasmChartShapeProperties::from),
            major_tick_marks: a.major_tick_marks.clone(),
            minor_tick_marks: a.minor_tick_marks.clone(),
            tick_labels: a.tick_labels,
            number_format: a.number_format.as_ref().map(WasmChartNumberFormat::from),
            shape_properties: a
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartExSeries {
    pub layout: String,
    pub data_id: u32,
    pub unique_id: Option<String>,
    pub hidden: Option<bool>,
    pub owner_idx: Option<u32>,
    pub format_idx: Option<u32>,
    pub text: Option<WasmChartExText>,
    pub data_labels: Option<WasmChartExDataLabels>,
    pub data_points: Vec<WasmChartExDataPoint>,
    pub layout_properties: Option<WasmChartExLayoutPr>,
    pub axis_ids: Vec<u32>,
    pub value_colors: bool,
    pub value_color_positions: Option<WasmChartExValueColorPositions>,
    pub shape_properties: Option<WasmChartShapeProperties>,
}

impl From<&duke_sheets_chart::ChartExSeries> for WasmChartExSeries {
    fn from(s: &duke_sheets_chart::ChartExSeries) -> Self {
        Self {
            layout: chart_ex_layout_to_string(&s.layout).into(),
            data_id: s.data_id,
            unique_id: s.unique_id.clone(),
            hidden: s.hidden,
            owner_idx: s.owner_idx,
            format_idx: s.format_idx,
            text: s.text.as_ref().map(WasmChartExText::from),
            data_labels: s.data_labels.as_ref().map(WasmChartExDataLabels::from),
            data_points: s
                .data_points
                .iter()
                .map(WasmChartExDataPoint::from)
                .collect(),
            layout_properties: s.layout_properties.as_ref().map(WasmChartExLayoutPr::from),
            axis_ids: s.axis_ids.clone(),
            value_colors: s.value_colors.is_some(),
            value_color_positions: s
                .value_color_positions
                .as_ref()
                .map(WasmChartExValueColorPositions::from),
            shape_properties: s
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmChartEx {
    pub layout: String,
    pub version: Option<String>,
    pub feature_list: Option<String>,
    pub fallback_img: Option<String>,
    pub title: Option<WasmChartExTitle>,
    pub data: Vec<WasmChartExData>,
    pub plot_area: WasmChartExPlotArea,
    pub legend: Option<WasmChartExLegend>,
    pub anchor: WasmDrawingAnchor,
    pub shape_properties: Option<WasmChartShapeProperties>,
    pub format_overrides: Vec<WasmChartExFormatOverride>,
    pub print_settings: Option<WasmChartExPrintSettings>,
    pub external_data_rel_id: Option<String>,
    pub external_data_auto_update: Option<bool>,
}

impl From<&duke_sheets_chart::ChartEx> for WasmChartEx {
    fn from(c: &duke_sheets_chart::ChartEx) -> Self {
        let layout = c
            .plot_area
            .series
            .first()
            .map(|s| chart_ex_layout_to_string(&s.layout))
            .unwrap_or("unknown");
        Self {
            layout: layout.into(),
            version: c.version.clone(),
            feature_list: c.feature_list.clone(),
            fallback_img: c.fallback_img.clone(),
            title: c.title.as_ref().map(WasmChartExTitle::from),
            data: c.data.iter().map(WasmChartExData::from).collect(),
            plot_area: WasmChartExPlotArea::from(&c.plot_area),
            legend: c.legend.as_ref().map(WasmChartExLegend::from),
            anchor: WasmDrawingAnchor::from(&c.anchor),
            shape_properties: c
                .shape_properties
                .as_ref()
                .map(WasmChartShapeProperties::from),
            format_overrides: c
                .format_overrides
                .iter()
                .map(WasmChartExFormatOverride::from)
                .collect(),
            print_settings: c
                .print_settings
                .as_ref()
                .map(WasmChartExPrintSettings::from),
            external_data_rel_id: c.external_data.as_ref().map(|e| e.rel_id.clone()),
            external_data_auto_update: c.external_data.as_ref().and_then(|e| e.auto_update),
        }
    }
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WasmEmbeddedImage {
    pub id: u32,
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    pub anchor: WasmDrawingAnchor,
    pub format: String,
    pub media_path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub svg_media_path: Option<String>,
    pub width_emu: i64,
    pub height_emu: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<i32>,
    pub flip_h: bool,
    pub flip_v: bool,
    pub data: Vec<u8>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub svg_data: Option<Vec<u8>>,
}

impl From<&duke_sheets_chart::EmbeddedImage> for WasmEmbeddedImage {
    fn from(img: &duke_sheets_chart::EmbeddedImage) -> Self {
        WasmEmbeddedImage {
            id: img.id,
            name: img.name.clone(),
            description: img.description.clone(),
            anchor: WasmDrawingAnchor::from(&img.anchor),
            format: img.format.as_str().to_string(),
            media_path: img.media_path.clone(),
            svg_media_path: img.svg_media_path.clone(),
            width_emu: img.width_emu,
            height_emu: img.height_emu,
            rotation: img.rotation,
            flip_h: img.flip_h,
            flip_v: img.flip_v,
            data: img.data().to_vec(),
            svg_data: img.svg_data().map(|b| b.to_vec()),
        }
    }
}
