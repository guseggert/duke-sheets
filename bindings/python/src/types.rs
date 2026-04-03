use duke_sheets_chart as chart;
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
use pyo3::prelude::*;

use crate::PyCalculationImage;

#[pyclass(name = "Color")]
#[derive(Clone)]
pub struct PyColor {
    #[pyo3(get)]
    pub color_type: String,
    #[pyo3(get)]
    pub hex: String,
    #[pyo3(get)]
    pub r: Option<u32>,
    #[pyo3(get)]
    pub g: Option<u32>,
    #[pyo3(get)]
    pub b: Option<u32>,
    #[pyo3(get)]
    pub a: Option<u32>,
    #[pyo3(get)]
    pub theme_index: Option<u32>,
    #[pyo3(get)]
    pub tint: Option<i32>,
    #[pyo3(get)]
    pub palette_index: Option<u32>,
}

impl From<&CoreColor> for PyColor {
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

#[pyclass(name = "FontStyle")]
#[derive(Clone)]
pub struct PyFontStyle {
    #[pyo3(get)]
    pub name: String,
    #[pyo3(get)]
    pub size: f64,
    #[pyo3(get)]
    pub bold: bool,
    #[pyo3(get)]
    pub italic: bool,
    #[pyo3(get)]
    pub underline: String,
    #[pyo3(get)]
    pub strikethrough: bool,
    #[pyo3(get)]
    pub color: PyColor,
    #[pyo3(get)]
    pub vertical_align: String,
    #[pyo3(get)]
    pub family: Option<u32>,
    #[pyo3(get)]
    pub charset: Option<u32>,
    #[pyo3(get)]
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

impl From<&CoreFontStyle> for PyFontStyle {
    fn from(f: &CoreFontStyle) -> Self {
        Self {
            name: f.name.clone(),
            size: f.size,
            bold: f.bold,
            italic: f.italic,
            underline: underline_to_string(&f.underline).into(),
            strikethrough: f.strikethrough,
            color: PyColor::from(&f.color),
            vertical_align: font_valign_to_string(&f.vertical_align).into(),
            family: f.family.map(|v| v as u32),
            charset: f.charset.map(|v| v as u32),
            scheme: f.scheme.clone(),
        }
    }
}

#[pyclass(name = "GradientStop")]
#[derive(Clone)]
pub struct PyGradientStop {
    #[pyo3(get)]
    pub position: f64,
    #[pyo3(get)]
    pub color: PyColor,
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

#[pyclass(name = "FillStyle")]
#[derive(Clone)]
pub struct PyFillStyle {
    #[pyo3(get)]
    pub fill_type: String,
    #[pyo3(get)]
    pub color: Option<PyColor>,
    #[pyo3(get)]
    pub pattern: Option<String>,
    #[pyo3(get)]
    pub foreground: Option<PyColor>,
    #[pyo3(get)]
    pub background: Option<PyColor>,
    #[pyo3(get)]
    pub gradient_type: Option<String>,
    #[pyo3(get)]
    pub angle: Option<f64>,
    #[pyo3(get)]
    pub stops: Option<Vec<PyGradientStop>>,
}

impl From<&CoreFillStyle> for PyFillStyle {
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
                color: Some(PyColor::from(color)),
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
                foreground: Some(PyColor::from(foreground)),
                background: Some(PyColor::from(background)),
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
                        .map(|s| PyGradientStop {
                            position: s.position,
                            color: PyColor::from(&s.color),
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

#[pyclass(name = "BorderEdge")]
#[derive(Clone)]
pub struct PyBorderEdge {
    #[pyo3(get)]
    pub style: String,
    #[pyo3(get)]
    pub color: PyColor,
}

impl From<&CoreBorderEdge> for PyBorderEdge {
    fn from(e: &CoreBorderEdge) -> Self {
        Self {
            style: border_line_style_to_string(&e.style).into(),
            color: PyColor::from(&e.color),
        }
    }
}

#[pyclass(name = "BorderStyle")]
#[derive(Clone)]
pub struct PyBorderStyle {
    #[pyo3(get)]
    pub left: Option<PyBorderEdge>,
    #[pyo3(get)]
    pub right: Option<PyBorderEdge>,
    #[pyo3(get)]
    pub top: Option<PyBorderEdge>,
    #[pyo3(get)]
    pub bottom: Option<PyBorderEdge>,
    #[pyo3(get)]
    pub diagonal: Option<PyBorderEdge>,
    #[pyo3(get)]
    pub diagonal_direction: String,
}

impl From<&CoreBorderStyle> for PyBorderStyle {
    fn from(b: &CoreBorderStyle) -> Self {
        Self {
            left: b.left.as_ref().map(PyBorderEdge::from),
            right: b.right.as_ref().map(PyBorderEdge::from),
            top: b.top.as_ref().map(PyBorderEdge::from),
            bottom: b.bottom.as_ref().map(PyBorderEdge::from),
            diagonal: b.diagonal.as_ref().map(PyBorderEdge::from),
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

#[pyclass(name = "Alignment")]
#[derive(Clone)]
pub struct PyAlignment {
    #[pyo3(get)]
    pub horizontal: String,
    #[pyo3(get)]
    pub vertical: String,
    #[pyo3(get)]
    pub wrap_text: bool,
    #[pyo3(get)]
    pub shrink_to_fit: bool,
    #[pyo3(get)]
    pub indent: u32,
    #[pyo3(get)]
    pub rotation: i32,
    #[pyo3(get)]
    pub reading_order: String,
}

impl From<&CoreAlignment> for PyAlignment {
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

#[pyclass(name = "NumberFormat")]
#[derive(Clone)]
pub struct PyNumberFormat {
    #[pyo3(get)]
    pub format_type: String,
    #[pyo3(get)]
    pub id: Option<u32>,
    #[pyo3(get)]
    pub format_string: String,
    #[pyo3(get)]
    pub is_date_format: bool,
}

impl From<&CoreNumberFormat> for PyNumberFormat {
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

#[pyclass(name = "CellProtection")]
#[derive(Clone)]
pub struct PyCellProtection {
    #[pyo3(get)]
    pub locked: bool,
    #[pyo3(get)]
    pub hidden: bool,
}

#[pyclass(name = "Style")]
#[derive(Clone)]
pub struct PyStyle {
    #[pyo3(get)]
    pub font: PyFontStyle,
    #[pyo3(get)]
    pub fill: PyFillStyle,
    #[pyo3(get)]
    pub border: PyBorderStyle,
    #[pyo3(get)]
    pub alignment: PyAlignment,
    #[pyo3(get)]
    pub number_format: PyNumberFormat,
    #[pyo3(get)]
    pub protection: PyCellProtection,
}

impl From<&CoreStyle> for PyStyle {
    fn from(s: &CoreStyle) -> Self {
        Self {
            font: PyFontStyle::from(&s.font),
            fill: PyFillStyle::from(&s.fill),
            border: PyBorderStyle::from(&s.border),
            alignment: PyAlignment::from(&s.alignment),
            number_format: PyNumberFormat::from(&s.number_format),
            protection: PyCellProtection {
                locked: s.protection.locked,
                hidden: s.protection.hidden,
            },
        }
    }
}

#[pyclass(name = "Hyperlink")]
#[derive(Clone)]
pub struct PyHyperlink {
    #[pyo3(get)]
    pub target: String,
    #[pyo3(get)]
    pub display: Option<String>,
    #[pyo3(get)]
    pub tooltip: Option<String>,
    #[pyo3(get)]
    pub location: Option<String>,
}

impl From<&core::Hyperlink> for PyHyperlink {
    fn from(h: &core::Hyperlink) -> Self {
        Self {
            target: h.target.clone(),
            display: h.display.clone(),
            tooltip: h.tooltip.clone(),
            location: h.location.clone(),
        }
    }
}

#[pyclass(name = "Comment")]
#[derive(Clone)]
pub struct PyComment {
    #[pyo3(get)]
    pub author: String,
    #[pyo3(get)]
    pub text: String,
    #[pyo3(get)]
    pub visible: bool,
}

impl From<&core::CellComment> for PyComment {
    fn from(c: &core::CellComment) -> Self {
        Self {
            author: c.author.clone(),
            text: c.text.clone(),
            visible: c.visible,
        }
    }
}

#[pyclass(name = "CommentEntry")]
#[derive(Clone)]
pub struct PyCommentEntry {
    #[pyo3(get)]
    pub row: u32,
    #[pyo3(get)]
    pub col: u32,
    #[pyo3(get)]
    pub comment: PyComment,
}

#[pyclass(name = "FreezePanes")]
#[derive(Clone)]
pub struct PyFreezePanes {
    #[pyo3(get)]
    pub row: u32,
    #[pyo3(get)]
    pub col: u32,
}

impl From<&core::FreezePanes> for PyFreezePanes {
    fn from(f: &core::FreezePanes) -> Self {
        Self {
            row: f.row,
            col: f.col as u32,
        }
    }
}

#[pyclass(name = "SplitPanes")]
#[derive(Clone)]
pub struct PySplitPanes {
    #[pyo3(get)]
    pub x_split: f64,
    #[pyo3(get)]
    pub y_split: f64,
    #[pyo3(get)]
    pub top_left_row: Option<u32>,
    #[pyo3(get)]
    pub top_left_col: Option<u32>,
    #[pyo3(get)]
    pub active_pane: Option<String>,
}

impl From<&core::SplitPanes> for PySplitPanes {
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

#[pyclass(name = "Selection")]
#[derive(Clone)]
pub struct PySelection {
    #[pyo3(get)]
    pub pane: Option<String>,
    #[pyo3(get)]
    pub active_cell: Option<String>,
    #[pyo3(get)]
    pub sqref: Option<String>,
}

impl From<&core::Selection> for PySelection {
    fn from(s: &core::Selection) -> Self {
        Self {
            pane: s.pane.clone(),
            active_cell: s.active_cell.clone(),
            sqref: s.sqref.clone(),
        }
    }
}

#[pyclass(name = "SheetProtection")]
#[derive(Clone)]
pub struct PySheetProtection {
    #[pyo3(get)]
    pub protected: bool,
    #[pyo3(get)]
    pub select_locked_cells: bool,
    #[pyo3(get)]
    pub select_unlocked_cells: bool,
    #[pyo3(get)]
    pub format_cells: bool,
    #[pyo3(get)]
    pub format_columns: bool,
    #[pyo3(get)]
    pub format_rows: bool,
    #[pyo3(get)]
    pub insert_columns: bool,
    #[pyo3(get)]
    pub insert_rows: bool,
    #[pyo3(get)]
    pub insert_hyperlinks: bool,
    #[pyo3(get)]
    pub delete_columns: bool,
    #[pyo3(get)]
    pub delete_rows: bool,
    #[pyo3(get)]
    pub sort: bool,
    #[pyo3(get)]
    pub auto_filter: bool,
    #[pyo3(get)]
    pub pivot_tables: bool,
}

impl From<&core::SheetProtection> for PySheetProtection {
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

#[pyclass(name = "PageSetup")]
#[derive(Clone)]
pub struct PyPageSetup {
    #[pyo3(get)]
    pub paper_size: u32,
    #[pyo3(get)]
    pub orientation: String,
    #[pyo3(get)]
    pub scale: u32,
    #[pyo3(get)]
    pub fit_to_width: Option<u32>,
    #[pyo3(get)]
    pub fit_to_height: Option<u32>,
    #[pyo3(get)]
    pub top_margin: f64,
    #[pyo3(get)]
    pub bottom_margin: f64,
    #[pyo3(get)]
    pub left_margin: f64,
    #[pyo3(get)]
    pub right_margin: f64,
    #[pyo3(get)]
    pub header_margin: f64,
    #[pyo3(get)]
    pub footer_margin: f64,
    #[pyo3(get)]
    pub print_gridlines: bool,
    #[pyo3(get)]
    pub print_headings: bool,
    #[pyo3(get)]
    pub odd_header: Option<String>,
    #[pyo3(get)]
    pub odd_footer: Option<String>,
    #[pyo3(get)]
    pub even_header: Option<String>,
    #[pyo3(get)]
    pub even_footer: Option<String>,
    #[pyo3(get)]
    pub first_header: Option<String>,
    #[pyo3(get)]
    pub first_footer: Option<String>,
    #[pyo3(get)]
    pub different_odd_even: bool,
    #[pyo3(get)]
    pub different_first: bool,
    #[pyo3(get)]
    pub scale_with_doc: bool,
    #[pyo3(get)]
    pub align_with_margins: bool,
}

impl From<&core::PageSetup> for PyPageSetup {
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

#[pyclass(name = "PageBreak")]
#[derive(Clone)]
pub struct PyPageBreak {
    #[pyo3(get)]
    pub id: u32,
    #[pyo3(get)]
    pub min: u32,
    #[pyo3(get)]
    pub max: u32,
    #[pyo3(get)]
    pub manual: bool,
}

impl From<&core::PageBreak> for PyPageBreak {
    fn from(b: &core::PageBreak) -> Self {
        Self {
            id: b.id,
            min: b.min,
            max: b.max,
            manual: b.man,
        }
    }
}

#[pyclass(name = "WorkbookSettings")]
#[derive(Clone)]
pub struct PyWorkbookSettings {
    #[pyo3(get)]
    pub date_1904: bool,
    #[pyo3(get)]
    pub protected: bool,
    #[pyo3(get)]
    pub calc_on_open: bool,
    #[pyo3(get)]
    pub theme: Option<String>,
}

impl From<&core::WorkbookSettings> for PyWorkbookSettings {
    fn from(s: &core::WorkbookSettings) -> Self {
        Self {
            date_1904: s.date_1904,
            protected: s.protected,
            calc_on_open: s.calc_on_open,
            theme: s.theme.clone(),
        }
    }
}

#[pyclass(name = "NamedRange")]
#[derive(Clone)]
pub struct PyNamedRange {
    #[pyo3(get)]
    pub name: String,
    #[pyo3(get)]
    pub scope: String,
    #[pyo3(get)]
    pub sheet_index: Option<u32>,
    #[pyo3(get)]
    pub refers_to: String,
    #[pyo3(get)]
    pub comment: Option<String>,
    #[pyo3(get)]
    pub hidden: bool,
}

#[pyclass(name = "Table")]
#[derive(Clone)]
pub struct PyTable {
    #[pyo3(get)]
    pub id: u32,
    #[pyo3(get)]
    pub name: String,
    #[pyo3(get)]
    pub display_name: String,
    #[pyo3(get)]
    pub reference: String,
    #[pyo3(get)]
    pub columns: Vec<PyTableColumn>,
    #[pyo3(get)]
    pub style_info: Option<PyTableStyleInfo>,
    #[pyo3(get)]
    pub header_row_count: u32,
    #[pyo3(get)]
    pub totals_row_count: u32,
    #[pyo3(get)]
    pub totals_row_shown: bool,
}

impl From<&core::Table> for PyTable {
    fn from(t: &core::Table) -> Self {
        Self {
            id: t.id,
            name: t.name.clone(),
            display_name: t.display_name.clone(),
            reference: t.reference.to_string(),
            columns: t.columns.iter().map(PyTableColumn::from).collect(),
            style_info: t.style_info.as_ref().map(PyTableStyleInfo::from),
            header_row_count: t.header_row_count,
            totals_row_count: t.totals_row_count,
            totals_row_shown: t.totals_row_shown,
        }
    }
}

#[pyclass(name = "TableColumn")]
#[derive(Clone)]
pub struct PyTableColumn {
    #[pyo3(get)]
    pub id: u32,
    #[pyo3(get)]
    pub name: String,
    #[pyo3(get)]
    pub totals_row_function: Option<String>,
    #[pyo3(get)]
    pub totals_row_formula: Option<String>,
    #[pyo3(get)]
    pub totals_row_label: Option<String>,
    #[pyo3(get)]
    pub calculated_column_formula: Option<String>,
}

impl From<&core::TableColumn> for PyTableColumn {
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

#[pyclass(name = "TableStyleInfo")]
#[derive(Clone)]
pub struct PyTableStyleInfo {
    #[pyo3(get)]
    pub name: Option<String>,
    #[pyo3(get)]
    pub show_first_column: bool,
    #[pyo3(get)]
    pub show_last_column: bool,
    #[pyo3(get)]
    pub show_row_stripes: bool,
    #[pyo3(get)]
    pub show_column_stripes: bool,
}

impl From<&core::TableStyleInfo> for PyTableStyleInfo {
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

#[pyclass(name = "AutoFilter")]
#[derive(Clone)]
pub struct PyAutoFilter {
    #[pyo3(get)]
    pub range: String,
    #[pyo3(get)]
    pub filter_columns: Vec<PyFilterColumn>,
}

impl From<&core::AutoFilter> for PyAutoFilter {
    fn from(af: &core::AutoFilter) -> Self {
        Self {
            range: af.range.to_string(),
            filter_columns: af.filter_columns.iter().map(PyFilterColumn::from).collect(),
        }
    }
}

#[pyclass(name = "FilterColumn")]
#[derive(Clone)]
pub struct PyFilterColumn {
    #[pyo3(get)]
    pub col_id: u32,
    #[pyo3(get)]
    pub hidden_button: bool,
    #[pyo3(get)]
    pub show_button: bool,
    #[pyo3(get)]
    pub filter_type: String,
    #[pyo3(get)]
    pub values: Option<Vec<String>>,
    #[pyo3(get)]
    pub blank: Option<bool>,
}

impl From<&core::FilterColumn> for PyFilterColumn {
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

#[pyclass(name = "DataValidation")]
#[derive(Clone)]
pub struct PyDataValidation {
    #[pyo3(get)]
    pub validation_type: String,
    #[pyo3(get)]
    pub ranges: Vec<String>,
    #[pyo3(get)]
    pub allow_blank: bool,
    #[pyo3(get)]
    pub show_dropdown: bool,
    #[pyo3(get)]
    pub show_input_message: bool,
    #[pyo3(get)]
    pub input_title: Option<String>,
    #[pyo3(get)]
    pub input_message: Option<String>,
    #[pyo3(get)]
    pub show_error_alert: bool,
    #[pyo3(get)]
    pub error_style: String,
    #[pyo3(get)]
    pub error_title: Option<String>,
    #[pyo3(get)]
    pub error_message: Option<String>,
    #[pyo3(get)]
    pub operator: Option<String>,
    #[pyo3(get)]
    pub value1: Option<String>,
    #[pyo3(get)]
    pub value2: Option<String>,
    #[pyo3(get)]
    pub list_source: Option<String>,
    #[pyo3(get)]
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

impl From<&core::DataValidation> for PyDataValidation {
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

#[pyclass(name = "ConditionalFormatRule")]
#[derive(Clone)]
pub struct PyConditionalFormatRule {
    #[pyo3(get)]
    pub rule_type: String,
    #[pyo3(get)]
    pub ranges: Vec<String>,
    #[pyo3(get)]
    pub priority: u32,
    #[pyo3(get)]
    pub stop_if_true: bool,
    #[pyo3(get)]
    pub operator: Option<String>,
    #[pyo3(get)]
    pub formula1: Option<String>,
    #[pyo3(get)]
    pub formula2: Option<String>,
    #[pyo3(get)]
    pub text: Option<String>,
    #[pyo3(get)]
    pub rank: Option<u32>,
    #[pyo3(get)]
    pub percent: Option<bool>,
    #[pyo3(get)]
    pub bottom: Option<bool>,
}

impl From<&core::ConditionalFormatRule> for PyConditionalFormatRule {
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

#[pyclass(name = "RichTextRun")]
#[derive(Clone)]
pub struct PyRichTextRun {
    #[pyo3(get)]
    pub text: String,
    #[pyo3(get)]
    pub font: Option<PyRunFont>,
}

impl From<&core::RichTextRun> for PyRichTextRun {
    fn from(r: &core::RichTextRun) -> Self {
        Self {
            text: r.text.clone(),
            font: r.font.as_ref().map(PyRunFont::from),
        }
    }
}

#[pyclass(name = "RunFont")]
#[derive(Clone)]
pub struct PyRunFont {
    #[pyo3(get)]
    pub bold: Option<bool>,
    #[pyo3(get)]
    pub italic: Option<bool>,
    #[pyo3(get)]
    pub size: Option<f64>,
    #[pyo3(get)]
    pub color: Option<PyColor>,
    #[pyo3(get)]
    pub name: Option<String>,
    #[pyo3(get)]
    pub underline: Option<String>,
    #[pyo3(get)]
    pub strikethrough: Option<bool>,
    #[pyo3(get)]
    pub vertical_align: Option<String>,
}

impl From<&core::RunFont> for PyRunFont {
    fn from(f: &core::RunFont) -> Self {
        Self {
            bold: f.bold,
            italic: f.italic,
            size: f.size,
            color: f.color.as_ref().map(PyColor::from),
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

#[pyclass(name = "HyperlinkEntry")]
#[derive(Clone)]
pub struct PyHyperlinkEntry {
    #[pyo3(get)]
    pub address: String,
    #[pyo3(get)]
    pub hyperlink: PyHyperlink,
}

#[pyclass(name = "RowCell")]
#[derive(Clone)]
pub struct PyRowCell {
    #[pyo3(get)]
    pub col: u32,
    #[pyo3(get)]
    pub value: String,
    #[pyo3(get)]
    pub style: Option<PyStyle>,
    #[pyo3(get)]
    pub merge_span: Option<PyMergeSpan>,
    #[pyo3(get)]
    pub is_merged_secondary: Option<bool>,
    #[pyo3(get)]
    pub hyperlink: Option<PyHyperlink>,
    #[pyo3(get)]
    pub comment: Option<PyComment>,
    #[pyo3(get)]
    pub formula: Option<String>,
    #[pyo3(get)]
    pub image: Option<PyCalculationImage>,
}

#[pyclass(name = "Row")]
#[derive(Clone)]
pub struct PyRow {
    #[pyo3(get)]
    pub index: u32,
    #[pyo3(get)]
    pub cells: Vec<PyRowCell>,
}

#[pyclass(name = "FormulaCell")]
#[derive(Clone)]
pub struct PyFormulaCell {
    #[pyo3(get)]
    pub row: u32,
    #[pyo3(get)]
    pub col: u32,
    #[pyo3(get)]
    pub formula: String,
}

#[pyclass(name = "SpillSource")]
#[derive(Clone)]
pub struct PySpillSource {
    #[pyo3(get)]
    pub row: u32,
    #[pyo3(get)]
    pub col: u32,
}

#[pyclass(name = "MergedRegion")]
#[derive(Clone)]
pub struct PyMergedRegion {
    #[pyo3(get)]
    pub start_row: u32,
    #[pyo3(get)]
    pub start_col: u32,
    #[pyo3(get)]
    pub end_row: u32,
    #[pyo3(get)]
    pub end_col: u32,
    /// The range as an A1-style string (e.g., "A1:C3").
    #[pyo3(get)]
    pub range: String,
}

#[pyclass(name = "MergeSpan")]
#[derive(Clone)]
pub struct PyMergeSpan {
    #[pyo3(get)]
    pub row_span: u32,
    #[pyo3(get)]
    pub col_span: u32,
}

#[pyclass(name = "ChartAnchor")]
#[derive(Clone)]
pub struct PyChartAnchor {
    #[pyo3(get)]
    pub from_col: u16,
    #[pyo3(get)]
    pub from_row: u32,
    #[pyo3(get)]
    pub from_col_offset: i64,
    #[pyo3(get)]
    pub from_row_offset: i64,
    #[pyo3(get)]
    pub to_col: u16,
    #[pyo3(get)]
    pub to_row: u32,
    #[pyo3(get)]
    pub to_col_offset: i64,
    #[pyo3(get)]
    pub to_row_offset: i64,
}

impl From<&duke_sheets::ChartAnchor> for PyChartAnchor {
    fn from(a: &duke_sheets::ChartAnchor) -> Self {
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

#[pyclass(name = "DataReference")]
#[derive(Clone)]
pub struct PyDataReference {
    #[pyo3(get)]
    pub ref_type: String,
    #[pyo3(get)]
    pub formula: Option<String>,
    #[pyo3(get)]
    pub numbers: Option<Vec<f64>>,
    #[pyo3(get)]
    pub strings: Option<Vec<String>>,
}

impl From<&duke_sheets::DataReference> for PyDataReference {
    fn from(r: &duke_sheets::DataReference) -> Self {
        match r {
            duke_sheets::DataReference::Formula(f) => Self {
                ref_type: "formula".into(),
                formula: Some(f.clone()),
                numbers: None,
                strings: None,
            },
            duke_sheets::DataReference::Numbers(ns) => Self {
                ref_type: "numbers".into(),
                formula: None,
                numbers: Some(ns.clone()),
                strings: None,
            },
            duke_sheets::DataReference::Strings(ss) => Self {
                ref_type: "strings".into(),
                formula: None,
                numbers: None,
                strings: Some(ss.clone()),
            },
        }
    }
}

#[pyclass(name = "ChartShapeProperties")]
#[derive(Clone)]
pub struct PyChartShapeProperties {
    #[pyo3(get)]
    pub solid_fill_hex: Option<String>,
    #[pyo3(get)]
    pub no_fill: bool,
    #[pyo3(get)]
    pub line_width: Option<i64>,
    #[pyo3(get)]
    pub line_color_hex: Option<String>,
    #[pyo3(get)]
    pub line_no_fill: bool,
    #[pyo3(get)]
    pub line_dash_style: Option<String>,
}

impl From<&chart::ChartShapeProperties> for PyChartShapeProperties {
    fn from(sp: &chart::ChartShapeProperties) -> Self {
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

#[pyclass(name = "DataSeries")]
#[derive(Clone)]
pub struct PyDataSeries {
    #[pyo3(get)]
    pub name: Option<String>,
    #[pyo3(get)]
    pub values: PyDataReference,
    #[pyo3(get)]
    pub categories: Option<PyDataReference>,
    #[pyo3(get)]
    pub data_labels: Option<PyDataLabels>,
    #[pyo3(get)]
    pub trendline: Option<PyTrendline>,
    #[pyo3(get)]
    pub error_bars: Option<PyErrorBars>,
    #[pyo3(get)]
    pub marker: Option<PyMarker>,
    #[pyo3(get)]
    pub data_points: Vec<PyDataPoint>,
    #[pyo3(get)]
    pub smooth: Option<bool>,
    #[pyo3(get)]
    pub explosion: Option<u32>,
    #[pyo3(get)]
    pub invert_if_negative: Option<bool>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&duke_sheets::DataSeries> for PyDataSeries {
    fn from(s: &duke_sheets::DataSeries) -> Self {
        Self {
            name: s.name.clone(),
            values: PyDataReference::from(&s.values),
            categories: s.categories.as_ref().map(PyDataReference::from),
            data_labels: s.data_labels.as_ref().map(PyDataLabels::from),
            trendline: s.trendline.as_ref().map(PyTrendline::from),
            error_bars: s.error_bars.as_ref().map(PyErrorBars::from),
            marker: s.marker.as_ref().map(PyMarker::from),
            data_points: s.data_points.iter().map(PyDataPoint::from).collect(),
            smooth: s.smooth,
            explosion: s.explosion,
            invert_if_negative: s.invert_if_negative,
            shape_properties: s.shape_properties.as_ref().map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "Axis")]
#[derive(Clone)]
pub struct PyAxis {
    #[pyo3(get)]
    pub title: Option<String>,
    #[pyo3(get)]
    pub minimum: Option<f64>,
    #[pyo3(get)]
    pub maximum: Option<f64>,
    #[pyo3(get)]
    pub major_unit: Option<f64>,
    #[pyo3(get)]
    pub minor_unit: Option<f64>,
    #[pyo3(get)]
    pub position: String,
    #[pyo3(get)]
    pub number_format: Option<PyChartNumberFormat>,
    #[pyo3(get)]
    pub major_gridlines: bool,
    #[pyo3(get)]
    pub minor_gridlines: bool,
    #[pyo3(get)]
    pub major_tick_mark: Option<String>,
    #[pyo3(get)]
    pub minor_tick_mark: Option<String>,
    #[pyo3(get)]
    pub label_position: Option<String>,
    #[pyo3(get)]
    pub delete: Option<bool>,
    #[pyo3(get)]
    pub crosses: Option<String>,
    #[pyo3(get)]
    pub cross_between: Option<String>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&duke_sheets::Axis> for PyAxis {
    fn from(a: &duke_sheets::Axis) -> Self {
        Self {
            title: a.title.clone(),
            minimum: a.minimum,
            maximum: a.maximum,
            major_unit: a.major_unit,
            minor_unit: a.minor_unit,
            position: match a.position {
                duke_sheets::AxisPosition::Bottom => "bottom",
                duke_sheets::AxisPosition::Top => "top",
                duke_sheets::AxisPosition::Left => "left",
                duke_sheets::AxisPosition::Right => "right",
            }
            .into(),
            number_format: a.number_format.as_ref().map(PyChartNumberFormat::from),
            major_gridlines: a.major_gridlines,
            minor_gridlines: a.minor_gridlines,
            major_tick_mark: a.major_tick_mark.as_ref().map(|t| match t {
                chart::TickMark::Cross => "cross",
                chart::TickMark::Inside => "inside",
                chart::TickMark::None => "none",
                chart::TickMark::Outside => "outside",
            }.into()),
            minor_tick_mark: a.minor_tick_mark.as_ref().map(|t| match t {
                chart::TickMark::Cross => "cross",
                chart::TickMark::Inside => "inside",
                chart::TickMark::None => "none",
                chart::TickMark::Outside => "outside",
            }.into()),
            label_position: a.label_position.as_ref().map(|p| match p {
                chart::TickLabelPosition::High => "high",
                chart::TickLabelPosition::Low => "low",
                chart::TickLabelPosition::NextTo => "nextTo",
                chart::TickLabelPosition::None => "none",
            }.into()),
            delete: a.delete,
            crosses: a.crosses.as_ref().map(|c| match c {
                chart::AxisCrosses::AutoZero => "autoZero",
                chart::AxisCrosses::Min => "min",
                chart::AxisCrosses::Max => "max",
            }.into()),
            cross_between: a.cross_between.as_ref().map(|c| match c {
                chart::CrossBetween::Between => "between",
                chart::CrossBetween::MidCat => "midCat",
            }.into()),
            shape_properties: a.shape_properties.as_ref().map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "Legend")]
#[derive(Clone)]
pub struct PyLegend {
    #[pyo3(get)]
    pub position: String,
    #[pyo3(get)]
    pub overlay: bool,
}

impl From<&duke_sheets::Legend> for PyLegend {
    fn from(l: &duke_sheets::Legend) -> Self {
        let position = match format!("{:?}", l.position).as_str() {
            "Right" => "right",
            "Top" => "top",
            "Bottom" => "bottom",
            "Left" => "left",
            "TopRight" => "topRight",
            _ => "right",
        }
        .to_string();
        Self {
            position,
            overlay: l.overlay,
        }
    }
}

#[pyclass(name = "ChartTypeGroup")]
#[derive(Clone)]
pub struct PyChartTypeGroup {
    #[pyo3(get)]
    pub chart_type: String,
    #[pyo3(get)]
    pub is_3d: bool,
    #[pyo3(get)]
    pub series: Vec<PyDataSeries>,
    #[pyo3(get)]
    pub data_labels: Option<PyDataLabels>,
    #[pyo3(get)]
    pub vary_colors: Option<bool>,
    #[pyo3(get)]
    pub gap_width: Option<u32>,
    #[pyo3(get)]
    pub overlap: Option<i32>,
    #[pyo3(get)]
    pub first_slice_angle: Option<u32>,
    #[pyo3(get)]
    pub hole_size: Option<u32>,
    #[pyo3(get)]
    pub bubble_scale: Option<u32>,
    #[pyo3(get)]
    pub show_negative_bubbles: Option<bool>,
    #[pyo3(get)]
    pub radar_style: Option<String>,
    #[pyo3(get)]
    pub wireframe: Option<bool>,
    #[pyo3(get)]
    pub axis_ids: Vec<u32>,
    #[pyo3(get)]
    pub drop_lines: Option<PyChartLines>,
    #[pyo3(get)]
    pub high_low_lines: Option<PyChartLines>,
    #[pyo3(get)]
    pub series_lines: Option<PyChartLines>,
    #[pyo3(get)]
    pub up_down_bars: Option<PyUpDownBars>,
}

impl From<&chart::ChartTypeGroup> for PyChartTypeGroup {
    fn from(g: &chart::ChartTypeGroup) -> Self {
        Self {
            chart_type: format!("{:?}", g.chart_type),
            is_3d: g.is_3d,
            series: g.series.iter().map(PyDataSeries::from).collect(),
            data_labels: g.data_labels.as_ref().map(PyDataLabels::from),
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
            drop_lines: g.drop_lines.as_ref().map(PyChartLines::from),
            high_low_lines: g.high_low_lines.as_ref().map(PyChartLines::from),
            series_lines: g.series_lines.as_ref().map(PyChartLines::from),
            up_down_bars: g.up_down_bars.as_ref().map(PyUpDownBars::from),
        }
    }
}


#[pyclass(name = "ChartAxis")]
#[derive(Clone)]
pub struct PyChartAxis {
    #[pyo3(get)]
    pub id: u32,
    #[pyo3(get)]
    pub cross_id: u32,
    #[pyo3(get)]
    pub axis: PyAxis,
}

impl From<&chart::ChartAxis> for PyChartAxis {
    fn from(a: &chart::ChartAxis) -> Self {
        Self {
            id: a.id,
            cross_id: a.cross_id,
            axis: PyAxis::from(&a.axis),
        }
    }
}

#[pyclass(name = "Chart")]
#[derive(Clone)]
pub struct PyChart {
    #[pyo3(get)]
    pub chart_type: String,
    #[pyo3(get)]
    pub title: Option<String>,
    #[pyo3(get)]
    pub series: Vec<PyDataSeries>,
    #[pyo3(get)]
    pub category_axis: Option<PyAxis>,
    #[pyo3(get)]
    pub value_axis: Option<PyAxis>,
    #[pyo3(get)]
    pub legend: Option<PyLegend>,
    #[pyo3(get)]
    pub anchor: PyChartAnchor,
    #[pyo3(get)]
    pub data_labels: Option<PyDataLabels>,
    #[pyo3(get)]
    pub view_3d: Option<PyView3D>,
    #[pyo3(get)]
    pub data_table: Option<PyChartDataTable>,
    #[pyo3(get)]
    pub display_blanks_as: Option<String>,
    #[pyo3(get)]
    pub plot_visible_only: Option<bool>,
    #[pyo3(get)]
    pub layout: Option<PyLayout>,
    #[pyo3(get)]
    pub is_3d: bool,
    #[pyo3(get)]
    pub vary_colors: Option<bool>,
    #[pyo3(get)]
    pub gap_width: Option<u32>,
    #[pyo3(get)]
    pub overlap: Option<i32>,
    #[pyo3(get)]
    pub first_slice_angle: Option<u32>,
    #[pyo3(get)]
    pub hole_size: Option<u32>,
    #[pyo3(get)]
    pub bubble_scale: Option<u32>,
    #[pyo3(get)]
    pub show_negative_bubbles: Option<bool>,
    #[pyo3(get)]
    pub auto_title_deleted: Option<bool>,
    #[pyo3(get)]
    pub rounded_corners: Option<bool>,
    #[pyo3(get)]
    pub show_dlbls_over_max: Option<bool>,
    #[pyo3(get)]
    pub wireframe: Option<bool>,
    #[pyo3(get)]
    pub radar_style: Option<String>,
    #[pyo3(get)]
    pub type_groups: Vec<PyChartTypeGroup>,
    #[pyo3(get)]
    pub axes: Vec<PyChartAxis>,
    #[pyo3(get)]
    pub drop_lines: Option<PyChartLines>,
    #[pyo3(get)]
    pub high_low_lines: Option<PyChartLines>,
    #[pyo3(get)]
    pub series_lines: Option<PyChartLines>,
    #[pyo3(get)]
    pub up_down_bars: Option<PyUpDownBars>,
}

impl From<&duke_sheets::Chart> for PyChart {
    fn from(c: &duke_sheets::Chart) -> Self {
        let chart_type = match &c.chart_type {
            duke_sheets::ChartType::Unsupported(tag) => format!("Unsupported({})", tag),
            other => format!("{:?}", other),
        };
        Self {
            chart_type,
            title: c.title.clone(),
            series: c.series.iter().map(PyDataSeries::from).collect(),
            category_axis: c.category_axis.as_ref().map(PyAxis::from),
            value_axis: c.value_axis.as_ref().map(PyAxis::from),
            legend: c.legend.as_ref().map(PyLegend::from),
            anchor: PyChartAnchor::from(&c.anchor),
            data_labels: c.data_labels.as_ref().map(PyDataLabels::from),
            view_3d: c.view_3d.as_ref().map(PyView3D::from),
            data_table: c.data_table.as_ref().map(PyChartDataTable::from),
            display_blanks_as: c.display_blanks_as.as_ref().map(|d| match d {
                chart::DisplayBlanksAs::Gap => "gap",
                chart::DisplayBlanksAs::Span => "span",
                chart::DisplayBlanksAs::Zero => "zero",
            }.into()),
            plot_visible_only: c.plot_visible_only,
            layout: c.layout.as_ref().map(PyLayout::from),
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
            type_groups: c.type_groups.iter().map(PyChartTypeGroup::from).collect(),
            axes: c.axes.iter().map(PyChartAxis::from).collect(),
            drop_lines: c.drop_lines.as_ref().map(PyChartLines::from),
            high_low_lines: c.high_low_lines.as_ref().map(PyChartLines::from),
            series_lines: c.series_lines.as_ref().map(PyChartLines::from),
            up_down_bars: c.up_down_bars.as_ref().map(PyUpDownBars::from),
        }
    }
}

#[pyclass(name = "ChartNumberFormat")]
#[derive(Clone)]
pub struct PyChartNumberFormat {
    #[pyo3(get)]
    pub format_code: String,
    #[pyo3(get)]
    pub source_linked: Option<bool>,
}

impl From<&chart::NumberFormat> for PyChartNumberFormat {
    fn from(n: &chart::NumberFormat) -> Self {
        Self {
            format_code: n.format_code.clone(),
            source_linked: n.source_linked,
        }
    }
}

#[pyclass(name = "DataLabels")]
#[derive(Clone)]
pub struct PyDataLabels {
    #[pyo3(get)]
    pub show_legend_key: Option<bool>,
    #[pyo3(get)]
    pub show_value: Option<bool>,
    #[pyo3(get)]
    pub show_category_name: Option<bool>,
    #[pyo3(get)]
    pub show_series_name: Option<bool>,
    #[pyo3(get)]
    pub show_percent: Option<bool>,
    #[pyo3(get)]
    pub show_bubble_size: Option<bool>,
    #[pyo3(get)]
    pub separator: Option<String>,
    #[pyo3(get)]
    pub position: Option<String>,
    #[pyo3(get)]
    pub number_format: Option<PyChartNumberFormat>,
    #[pyo3(get)]
    pub show_leader_lines: Option<bool>,
}

impl From<&chart::DataLabels> for PyDataLabels {
    fn from(d: &chart::DataLabels) -> Self {
        Self {
            show_legend_key: d.show_legend_key,
            show_value: d.show_value,
            show_category_name: d.show_category_name,
            show_series_name: d.show_series_name,
            show_percent: d.show_percent,
            show_bubble_size: d.show_bubble_size,
            separator: d.separator.clone(),
            position: d.position.as_ref().map(|p| match p {
                chart::DataLabelPosition::BestFit => "bestFit",
                chart::DataLabelPosition::Bottom => "bottom",
                chart::DataLabelPosition::Center => "center",
                chart::DataLabelPosition::InsideBase => "insideBase",
                chart::DataLabelPosition::InsideEnd => "insideEnd",
                chart::DataLabelPosition::Left => "left",
                chart::DataLabelPosition::OutsideEnd => "outsideEnd",
                chart::DataLabelPosition::Right => "right",
                chart::DataLabelPosition::Top => "top",
            }.into()),
            number_format: d.number_format.as_ref().map(PyChartNumberFormat::from),
            show_leader_lines: d.show_leader_lines,
        }
    }
}

#[pyclass(name = "Trendline")]
#[derive(Clone)]
pub struct PyTrendline {
    #[pyo3(get)]
    pub trendline_type: String,
    #[pyo3(get)]
    pub name: Option<String>,
    #[pyo3(get)]
    pub order: Option<u32>,
    #[pyo3(get)]
    pub period: Option<u32>,
    #[pyo3(get)]
    pub forward: Option<f64>,
    #[pyo3(get)]
    pub backward: Option<f64>,
    #[pyo3(get)]
    pub intercept: Option<f64>,
    #[pyo3(get)]
    pub display_r_squared: Option<bool>,
    #[pyo3(get)]
    pub display_equation: Option<bool>,
}

impl From<&chart::Trendline> for PyTrendline {
    fn from(t: &chart::Trendline) -> Self {
        Self {
            trendline_type: match t.trendline_type {
                chart::TrendlineType::Linear => "linear",
                chart::TrendlineType::Exponential => "exponential",
                chart::TrendlineType::Logarithmic => "logarithmic",
                chart::TrendlineType::MovingAverage => "movingAverage",
                chart::TrendlineType::Polynomial => "polynomial",
                chart::TrendlineType::Power => "power",
            }.into(),
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

#[pyclass(name = "ErrorBars")]
#[derive(Clone)]
pub struct PyErrorBars {
    #[pyo3(get)]
    pub direction: String,
    #[pyo3(get)]
    pub bar_type: String,
    #[pyo3(get)]
    pub value_type: String,
    #[pyo3(get)]
    pub value: Option<f64>,
    #[pyo3(get)]
    pub no_end_cap: Option<bool>,
}

impl From<&chart::ErrorBars> for PyErrorBars {
    fn from(e: &chart::ErrorBars) -> Self {
        Self {
            direction: match e.direction {
                chart::ErrorBarDirection::X => "x",
                chart::ErrorBarDirection::Y => "y",
            }.into(),
            bar_type: match e.bar_type {
                chart::ErrorBarType::Both => "both",
                chart::ErrorBarType::Minus => "minus",
                chart::ErrorBarType::Plus => "plus",
            }.into(),
            value_type: match e.value_type {
                chart::ErrorValueType::Custom => "custom",
                chart::ErrorValueType::FixedValue => "fixedValue",
                chart::ErrorValueType::Percentage => "percentage",
                chart::ErrorValueType::StandardDeviation => "standardDeviation",
                chart::ErrorValueType::StandardError => "standardError",
            }.into(),
            value: e.value,
            no_end_cap: e.no_end_cap,
        }
    }
}

#[pyclass(name = "Marker")]
#[derive(Clone)]
pub struct PyMarker {
    #[pyo3(get)]
    pub symbol: Option<String>,
    #[pyo3(get)]
    pub size: Option<u8>,
}

impl From<&chart::Marker> for PyMarker {
    fn from(m: &chart::Marker) -> Self {
        Self {
            symbol: m.symbol.as_ref().map(|s| match s {
                chart::MarkerSymbol::Circle => "circle",
                chart::MarkerSymbol::Dash => "dash",
                chart::MarkerSymbol::Diamond => "diamond",
                chart::MarkerSymbol::Dot => "dot",
                chart::MarkerSymbol::None => "none",
                chart::MarkerSymbol::Picture => "picture",
                chart::MarkerSymbol::Plus => "plus",
                chart::MarkerSymbol::Square => "square",
                chart::MarkerSymbol::Star => "star",
                chart::MarkerSymbol::Triangle => "triangle",
                chart::MarkerSymbol::X => "x",
                chart::MarkerSymbol::Auto => "auto",
            }.into()),
            size: m.size,
        }
    }
}

#[pyclass(name = "DataPoint")]
#[derive(Clone)]
pub struct PyDataPoint {
    #[pyo3(get)]
    pub index: u32,
    #[pyo3(get)]
    pub marker: Option<PyMarker>,
    #[pyo3(get)]
    pub explosion: Option<u32>,
}

impl From<&chart::DataPoint> for PyDataPoint {
    fn from(p: &chart::DataPoint) -> Self {
        Self {
            index: p.index,
            marker: p.marker.as_ref().map(PyMarker::from),
            explosion: p.explosion,
        }
    }
}

#[pyclass(name = "View3D")]
#[derive(Clone)]
pub struct PyView3D {
    #[pyo3(get)]
    pub rotate_x: Option<i32>,
    #[pyo3(get)]
    pub rotate_y: Option<i32>,
    #[pyo3(get)]
    pub depth_percent: Option<u32>,
    #[pyo3(get)]
    pub height_percent: Option<u32>,
    #[pyo3(get)]
    pub perspective: Option<u32>,
    #[pyo3(get)]
    pub right_angle_axes: Option<bool>,
}

impl From<&chart::View3D> for PyView3D {
    fn from(v: &chart::View3D) -> Self {
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

#[pyclass(name = "ChartDataTable")]
#[derive(Clone)]
pub struct PyChartDataTable {
    #[pyo3(get)]
    pub show_horizontal_border: Option<bool>,
    #[pyo3(get)]
    pub show_vertical_border: Option<bool>,
    #[pyo3(get)]
    pub show_outline: Option<bool>,
    #[pyo3(get)]
    pub show_keys: Option<bool>,
}

impl From<&chart::ChartDataTable> for PyChartDataTable {
    fn from(t: &chart::ChartDataTable) -> Self {
        Self {
            show_horizontal_border: t.show_horizontal_border,
            show_vertical_border: t.show_vertical_border,
            show_outline: t.show_outline,
            show_keys: t.show_keys,
        }
    }
}

#[pyclass(name = "ManualLayout")]
#[derive(Clone)]
pub struct PyManualLayout {
    #[pyo3(get)]
    pub x: Option<f64>,
    #[pyo3(get)]
    pub y: Option<f64>,
    #[pyo3(get)]
    pub width: Option<f64>,
    #[pyo3(get)]
    pub height: Option<f64>,
}

impl From<&chart::ManualLayout> for PyManualLayout {
    fn from(m: &chart::ManualLayout) -> Self {
        Self {
            x: m.x,
            y: m.y,
            width: m.width,
            height: m.height,
        }
    }
}

#[pyclass(name = "Layout")]
#[derive(Clone)]
pub struct PyLayout {
    #[pyo3(get)]
    pub manual_layout: Option<PyManualLayout>,
}

impl From<&chart::Layout> for PyLayout {
    fn from(l: &chart::Layout) -> Self {
        Self {
            manual_layout: l.manual_layout.as_ref().map(PyManualLayout::from),
        }
    }
}

#[pyclass(name = "ChartLines")]
#[derive(Clone)]
pub struct PyChartLines {
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartLines> for PyChartLines {
    fn from(cl: &chart::ChartLines) -> Self {
        Self {
            shape_properties: cl.shape_properties.as_ref().map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "UpDownBars")]
#[derive(Clone)]
pub struct PyUpDownBars {
    #[pyo3(get)]
    pub gap_width: Option<u32>,
    #[pyo3(get)]
    pub up_bars: Option<PyChartLines>,
    #[pyo3(get)]
    pub down_bars: Option<PyChartLines>,
}

impl From<&chart::UpDownBars> for PyUpDownBars {
    fn from(ud: &chart::UpDownBars) -> Self {
        Self {
            gap_width: ud.gap_width,
            up_bars: ud.up_bars.as_ref().map(PyChartLines::from),
            down_bars: ud.down_bars.as_ref().map(PyChartLines::from),
        }
    }
}

#[pyclass(name = "ChartSheet")]
#[derive(Clone)]
pub struct PyChartSheet {
    #[pyo3(get)]
    pub name: String,
    #[pyo3(get)]
    pub chart: PyChart,
    #[pyo3(get)]
    pub visibility: String,
}

impl From<&core::ChartSheet> for PyChartSheet {
    fn from(cs: &core::ChartSheet) -> Self {
        Self {
            name: cs.name.clone(),
            chart: PyChart::from(&cs.chart),
            visibility: match cs.visibility {
                core::worksheet::SheetVisibility::Visible => "visible",
                core::worksheet::SheetVisibility::Hidden => "hidden",
                core::worksheet::SheetVisibility::VeryHidden => "veryHidden",
            }
            .into(),
        }
    }
}

#[pyclass(name = "SheetSlot")]
#[derive(Clone)]
pub struct PySheetSlot {
    #[pyo3(get)]
    pub slot_type: String,
    #[pyo3(get)]
    pub index: u32,
}

impl From<&core::SheetSlot> for PySheetSlot {
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
