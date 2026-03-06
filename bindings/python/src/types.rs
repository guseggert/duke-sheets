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
