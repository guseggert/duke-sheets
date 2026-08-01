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
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList};

use crate::{PyCalculationImage, PyDrawingText};

#[pyclass(name = "Color")]
#[derive(Clone)]
pub struct PyColor {
    #[pyo3(get)]
    pub color_type: String,
    /// Context-free hex string; ``None`` for auto and theme colors,
    /// which resolve through :meth:`Workbook.resolve_color`.
    #[pyo3(get)]
    pub hex: Option<String>,
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
    pub tint: Option<f64>,
    #[pyo3(get)]
    pub palette_index: Option<u32>,
}

#[pymethods]
impl PyColor {
    /// Build a color for :meth:`Workbook.resolve_color`, e.g.
    /// ``Color("theme", theme_index=4, tint=0.5)`` or
    /// ``Color("rgb", r=255, g=0, b=0)``.
    #[new]
    #[pyo3(signature=(color_type, *, r=None, g=None, b=None, a=None, theme_index=None, tint=None, palette_index=None))]
    #[allow(clippy::too_many_arguments)]
    fn new(
        color_type: String,
        r: Option<u32>,
        g: Option<u32>,
        b: Option<u32>,
        a: Option<u32>,
        theme_index: Option<u32>,
        tint: Option<f64>,
        palette_index: Option<u32>,
    ) -> PyResult<Self> {
        let color = Self {
            color_type,
            hex: None,
            r,
            g,
            b,
            a,
            theme_index,
            tint,
            palette_index,
        };
        // Validate eagerly and stamp the context-free hex.
        let core = color.to_core()?;
        Ok(Self {
            hex: core.to_hex(),
            ..color
        })
    }
}

impl PyColor {
    /// Reconstruct the core color this DTO was built from.
    pub(crate) fn to_core(&self) -> PyResult<CoreColor> {
        let channel = |value: Option<u32>, name: &str| -> PyResult<u8> {
            let value = value
                .ok_or_else(|| PyValueError::new_err(format!("color is missing {name}")))?;
            u8::try_from(value)
                .map_err(|_| PyValueError::new_err(format!("color {name} out of range")))
        };
        match self.color_type.as_str() {
            "auto" => Ok(CoreColor::Auto),
            "rgb" => Ok(CoreColor::Rgb {
                r: channel(self.r, "r")?,
                g: channel(self.g, "g")?,
                b: channel(self.b, "b")?,
            }),
            "argb" => Ok(CoreColor::Argb {
                a: channel(self.a, "a")?,
                r: channel(self.r, "r")?,
                g: channel(self.g, "g")?,
                b: channel(self.b, "b")?,
            }),
            "theme" => {
                let index = channel(self.theme_index, "theme_index")?;
                let tint = tint_fraction(self.tint)?;
                Ok(CoreColor::Theme { index, tint })
            }
            "indexed" => Ok(CoreColor::Indexed(channel(self.palette_index, "palette_index")?)),
            other => Err(PyValueError::new_err(format!("unknown color type {other:?}"))),
        }
    }
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
                tint: Some(*tint),
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

fn style_input_error(message: impl Into<String>) -> PyErr {
    PyValueError::new_err(message.into())
}

fn dict_get<'py>(dict: &Bound<'py, PyDict>, key: &str) -> PyResult<Option<Bound<'py, PyAny>>> {
    Ok(dict.get_item(key)?.filter(|value| !value.is_none()))
}

fn dict_has(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<bool> {
    Ok(dict.get_item(key)?.is_some())
}

fn dict_get_string(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    dict_get(dict, key)?
        .map(|value| value.extract::<String>())
        .transpose()
}

fn dict_get_bool(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<bool>> {
    dict_get(dict, key)?
        .map(|value| value.extract::<bool>())
        .transpose()
}

fn dict_get_f64(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<f64>> {
    dict_get(dict, key)?
        .map(|value| value.extract::<f64>())
        .transpose()
}

fn dict_get_u32(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<u32>> {
    dict_get(dict, key)?
        .map(|value| value.extract::<u32>())
        .transpose()
}

fn dict_get_u32_alias(
    dict: &Bound<'_, PyDict>,
    primary: &str,
    alias: &str,
) -> PyResult<Option<u32>> {
    match dict_get_u32(dict, primary)? {
        Some(value) => Ok(Some(value)),
        None => dict_get_u32(dict, alias),
    }
}

fn dict_get_i32(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<i32>> {
    dict_get(dict, key)?
        .map(|value| value.extract::<i32>())
        .transpose()
}

fn u32_to_u8(value: u32, field: &str) -> PyResult<u8> {
    u8::try_from(value).map_err(|_| style_input_error(format!("{field} must be between 0 and 255")))
}

fn tint_fraction(value: Option<f64>) -> PyResult<f64> {
    let tint = value.unwrap_or(0.0);
    if !tint.is_finite() || !(-1.0..=1.0).contains(&tint) {
        return Err(style_input_error("tint must be between -1.0 and 1.0"));
    }
    Ok(tint)
}

fn parse_color_hex(hex: &str) -> PyResult<CoreColor> {
    CoreColor::from_hex(hex).ok_or_else(|| {
        style_input_error("color hex must be 6 or 8 hexadecimal characters, with optional # prefix")
    })
}

fn parse_rgb_hex(hex: &str) -> PyResult<CoreColor> {
    match parse_color_hex(hex)? {
        CoreColor::Rgb { r, g, b } => Ok(CoreColor::Rgb { r, g, b }),
        CoreColor::Argb { r, g, b, .. } => Ok(CoreColor::Rgb { r, g, b }),
        other => Ok(other),
    }
}

fn parse_argb_hex(hex: &str) -> PyResult<CoreColor> {
    match parse_color_hex(hex)? {
        CoreColor::Rgb { r, g, b } => Ok(CoreColor::Argb { a: 255, r, g, b }),
        CoreColor::Argb { a, r, g, b } => Ok(CoreColor::Argb { a, r, g, b }),
        other => Ok(other),
    }
}

fn color_parts_to_core(
    color_type: Option<&str>,
    hex: Option<&str>,
    r: Option<u32>,
    g: Option<u32>,
    b: Option<u32>,
    a: Option<u32>,
    theme_index: Option<u32>,
    tint: Option<f64>,
    palette_index: Option<u32>,
) -> PyResult<CoreColor> {
    match color_type {
        Some("auto") => Ok(CoreColor::Auto),
        Some("rgb") => {
            if let Some(hex) = hex {
                parse_rgb_hex(hex)
            } else {
                Ok(CoreColor::Rgb {
                    r: u32_to_u8(
                        r.ok_or_else(|| style_input_error("rgb color requires r"))?,
                        "r",
                    )?,
                    g: u32_to_u8(
                        g.ok_or_else(|| style_input_error("rgb color requires g"))?,
                        "g",
                    )?,
                    b: u32_to_u8(
                        b.ok_or_else(|| style_input_error("rgb color requires b"))?,
                        "b",
                    )?,
                })
            }
        }
        Some("argb") => {
            if let Some(hex) = hex {
                parse_argb_hex(hex)
            } else {
                Ok(CoreColor::Argb {
                    a: u32_to_u8(a.unwrap_or(255), "a")?,
                    r: u32_to_u8(
                        r.ok_or_else(|| style_input_error("argb color requires r"))?,
                        "r",
                    )?,
                    g: u32_to_u8(
                        g.ok_or_else(|| style_input_error("argb color requires g"))?,
                        "g",
                    )?,
                    b: u32_to_u8(
                        b.ok_or_else(|| style_input_error("argb color requires b"))?,
                        "b",
                    )?,
                })
            }
        }
        Some("theme") => Ok(CoreColor::Theme {
            index: u32_to_u8(
                theme_index.ok_or_else(|| style_input_error("theme color requires theme_index"))?,
                "theme_index",
            )?,
            tint: tint_fraction(tint)?,
        }),
        Some("indexed") => Ok(CoreColor::Indexed(u32_to_u8(
            palette_index
                .ok_or_else(|| style_input_error("indexed color requires palette_index"))?,
            "palette_index",
        )?)),
        Some(other) => Err(style_input_error(format!("unknown color_type {other:?}"))),
        None => {
            if let Some(hex) = hex {
                parse_color_hex(hex)
            } else if r.is_some() || g.is_some() || b.is_some() {
                Ok(CoreColor::Rgb {
                    r: u32_to_u8(
                        r.ok_or_else(|| style_input_error("rgb color requires r"))?,
                        "r",
                    )?,
                    g: u32_to_u8(
                        g.ok_or_else(|| style_input_error("rgb color requires g"))?,
                        "g",
                    )?,
                    b: u32_to_u8(
                        b.ok_or_else(|| style_input_error("rgb color requires b"))?,
                        "b",
                    )?,
                })
            } else if let Some(theme_index) = theme_index {
                Ok(CoreColor::Theme {
                    index: u32_to_u8(theme_index, "theme_index")?,
                    tint: tint_fraction(tint)?,
                })
            } else if let Some(palette_index) = palette_index {
                Ok(CoreColor::Indexed(u32_to_u8(
                    palette_index,
                    "palette_index",
                )?))
            } else {
                Err(style_input_error(
                    "color requires color_type, hex, rgb, theme_index, or palette_index",
                ))
            }
        }
    }
}

pub(crate) fn py_color_to_core(color: &PyColor) -> PyResult<CoreColor> {
    color_parts_to_core(
        Some(color.color_type.as_str()),
        color.hex.as_deref(),
        color.r,
        color.g,
        color.b,
        color.a,
        color.theme_index,
        color.tint,
        color.palette_index,
    )
}

pub(crate) fn color_input_to_core(input: &Bound<'_, PyAny>) -> PyResult<CoreColor> {
    if let Ok(color) = input.extract::<PyRef<PyColor>>() {
        return py_color_to_core(&color);
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("color must be a Color or dict"))?;
    let color_type = dict_get_string(dict, "color_type")?;
    let hex = dict_get_string(dict, "hex")?;
    color_parts_to_core(
        color_type.as_deref(),
        hex.as_deref(),
        dict_get_u32(dict, "r")?,
        dict_get_u32(dict, "g")?,
        dict_get_u32(dict, "b")?,
        dict_get_u32(dict, "a")?,
        dict_get_u32(dict, "theme_index")?,
        dict_get_f64(dict, "tint")?,
        dict_get_u32(dict, "palette_index")?,
    )
}

fn parse_underline_input(value: &str) -> PyResult<Underline> {
    match value {
        "none" => Ok(Underline::None),
        "single" => Ok(Underline::Single),
        "double" => Ok(Underline::Double),
        "singleAccounting" => Ok(Underline::SingleAccounting),
        "doubleAccounting" => Ok(Underline::DoubleAccounting),
        other => Err(style_input_error(format!("unknown underline {other:?}"))),
    }
}

fn parse_font_vertical_align_input(value: &str) -> PyResult<FontVerticalAlign> {
    match value {
        "baseline" => Ok(FontVerticalAlign::Baseline),
        "superscript" => Ok(FontVerticalAlign::Superscript),
        "subscript" => Ok(FontVerticalAlign::Subscript),
        other => Err(style_input_error(format!(
            "unknown vertical_align {other:?}"
        ))),
    }
}

fn py_font_to_core(font: &PyFontStyle) -> PyResult<CoreFontStyle> {
    Ok(CoreFontStyle {
        name: font.name.clone(),
        size: font.size,
        bold: font.bold,
        italic: font.italic,
        underline: parse_underline_input(&font.underline)?,
        strikethrough: font.strikethrough,
        color: py_color_to_core(&font.color)?,
        vertical_align: parse_font_vertical_align_input(&font.vertical_align)?,
        family: font.family.map(|v| u32_to_u8(v, "family")).transpose()?,
        charset: font.charset.map(|v| u32_to_u8(v, "charset")).transpose()?,
        scheme: font.scheme.clone(),
    })
}

fn is_full_font_dict(dict: &Bound<'_, PyDict>) -> PyResult<bool> {
    Ok(dict_has(dict, "name")?
        && dict_has(dict, "size")?
        && dict_has(dict, "bold")?
        && dict_has(dict, "italic")?
        && dict_has(dict, "underline")?
        && dict_has(dict, "strikethrough")?
        && dict_has(dict, "color")?
        && dict_has(dict, "vertical_align")?)
}

fn apply_font_dict(dict: &Bound<'_, PyDict>, font: &mut CoreFontStyle) -> PyResult<()> {
    if let Some(name) = dict_get_string(dict, "name")? {
        font.name = name;
    }
    if let Some(size) = dict_get_f64(dict, "size")? {
        font.size = size;
    }
    if let Some(bold) = dict_get_bool(dict, "bold")? {
        font.bold = bold;
    }
    if let Some(italic) = dict_get_bool(dict, "italic")? {
        font.italic = italic;
    }
    if let Some(underline) = dict_get_string(dict, "underline")? {
        font.underline = parse_underline_input(&underline)?;
    }
    if let Some(strikethrough) = dict_get_bool(dict, "strikethrough")? {
        font.strikethrough = strikethrough;
    }
    if let Some(color) = dict_get(dict, "color")? {
        font.color = color_input_to_core(&color)?;
    }
    if let Some(vertical_align) = dict_get_string(dict, "vertical_align")? {
        font.vertical_align = parse_font_vertical_align_input(&vertical_align)?;
    }
    if let Some(family) = dict_get_u32(dict, "family")? {
        font.family = Some(u32_to_u8(family, "family")?);
    }
    if let Some(charset) = dict_get_u32(dict, "charset")? {
        font.charset = Some(u32_to_u8(charset, "charset")?);
    }
    if let Some(scheme) = dict_get_string(dict, "scheme")? {
        font.scheme = Some(scheme);
    }
    Ok(())
}

fn apply_font_input(input: &Bound<'_, PyAny>, font: &mut CoreFontStyle) -> PyResult<()> {
    if let Ok(py_font) = input.extract::<PyRef<PyFontStyle>>() {
        *font = py_font_to_core(&py_font)?;
        return Ok(());
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("font must be a FontStyle or dict"))?;
    if is_full_font_dict(dict)? {
        let mut next = CoreFontStyle::default();
        apply_font_dict(dict, &mut next)?;
        *font = next;
    } else {
        apply_font_dict(dict, font)?;
    }
    Ok(())
}

fn parse_pattern_type_input(value: &str) -> PyResult<PatternType> {
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
        other => Err(style_input_error(format!("unknown fill pattern {other:?}"))),
    }
}

fn parse_gradient_type_input(value: &str) -> PyResult<GradientType> {
    match value {
        "linear" => Ok(GradientType::Linear),
        "path" => Ok(GradientType::Path),
        other => Err(style_input_error(format!(
            "unknown gradient_type {other:?}"
        ))),
    }
}

fn py_fill_to_core(fill: &PyFillStyle) -> PyResult<CoreFillStyle> {
    match fill.fill_type.as_str() {
        "none" => Ok(CoreFillStyle::None),
        "solid" => Ok(CoreFillStyle::Solid {
            color: py_color_to_core(
                fill.color
                    .as_ref()
                    .ok_or_else(|| style_input_error("solid fill requires color"))?,
            )?,
        }),
        "pattern" => Ok(CoreFillStyle::Pattern {
            pattern: parse_pattern_type_input(
                fill.pattern
                    .as_deref()
                    .ok_or_else(|| style_input_error("pattern fill requires pattern"))?,
            )?,
            foreground: py_color_to_core(
                fill.foreground
                    .as_ref()
                    .ok_or_else(|| style_input_error("pattern fill requires foreground"))?,
            )?,
            background: py_color_to_core(
                fill.background
                    .as_ref()
                    .ok_or_else(|| style_input_error("pattern fill requires background"))?,
            )?,
        }),
        "gradient" => Ok(CoreFillStyle::Gradient {
            gradient_type: parse_gradient_type_input(
                fill.gradient_type.as_deref().unwrap_or("linear"),
            )?,
            angle: fill.angle.unwrap_or(0.0),
            stops: fill
                .stops
                .as_ref()
                .ok_or_else(|| style_input_error("gradient fill requires stops"))?
                .iter()
                .map(|stop| {
                    Ok(core::style::GradientStop {
                        position: stop.position,
                        color: py_color_to_core(&stop.color)?,
                    })
                })
                .collect::<PyResult<Vec<_>>>()?,
        }),
        other => Err(style_input_error(format!("unknown fill_type {other:?}"))),
    }
}

fn fill_dict_to_core(dict: &Bound<'_, PyDict>) -> PyResult<CoreFillStyle> {
    match dict_get_string(dict, "fill_type")?.as_deref() {
        Some("none") => Ok(CoreFillStyle::None),
        Some("solid") | None if dict_get(dict, "color")?.is_some() => Ok(CoreFillStyle::Solid {
            color: color_input_to_core(
                &dict_get(dict, "color")?
                    .ok_or_else(|| style_input_error("solid fill requires color"))?,
            )?,
        }),
        Some("pattern") => Ok(CoreFillStyle::Pattern {
            pattern: parse_pattern_type_input(
                dict_get_string(dict, "pattern")?
                    .as_deref()
                    .ok_or_else(|| style_input_error("pattern fill requires pattern"))?,
            )?,
            foreground: color_input_to_core(
                &dict_get(dict, "foreground")?
                    .ok_or_else(|| style_input_error("pattern fill requires foreground"))?,
            )?,
            background: color_input_to_core(
                &dict_get(dict, "background")?
                    .ok_or_else(|| style_input_error("pattern fill requires background"))?,
            )?,
        }),
        Some("gradient") => {
            let stops_any = dict_get(dict, "stops")?
                .ok_or_else(|| style_input_error("gradient fill requires stops"))?
                .downcast_into::<PyList>()
                .map_err(|_| style_input_error("gradient fill stops must be a list"))?;
            let stops = stops_any
                .iter()
                .map(|stop| {
                    let stop = stop
                        .downcast::<PyDict>()
                        .map_err(|_| style_input_error("gradient stop must be a dict"))?;
                    let position = dict_get_f64(&stop, "position")?
                        .ok_or_else(|| style_input_error("gradient stop requires position"))?;
                    let color = color_input_to_core(
                        &dict_get(&stop, "color")?
                            .ok_or_else(|| style_input_error("gradient stop requires color"))?,
                    )?;
                    Ok(core::style::GradientStop { position, color })
                })
                .collect::<PyResult<Vec<_>>>()?;
            Ok(CoreFillStyle::Gradient {
                gradient_type: parse_gradient_type_input(
                    dict_get_string(dict, "gradient_type")?
                        .as_deref()
                        .unwrap_or("linear"),
                )?,
                angle: dict_get_f64(dict, "angle")?.unwrap_or(0.0),
                stops,
            })
        }
        Some(other) => Err(style_input_error(format!("unknown fill_type {other:?}"))),
        None => Err(style_input_error("fill patch requires fill_type or color")),
    }
}

fn apply_fill_input(input: &Bound<'_, PyAny>, fill: &mut CoreFillStyle) -> PyResult<()> {
    if let Ok(py_fill) = input.extract::<PyRef<PyFillStyle>>() {
        *fill = py_fill_to_core(&py_fill)?;
        return Ok(());
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("fill must be a FillStyle or dict"))?;
    *fill = fill_dict_to_core(dict)?;
    Ok(())
}

fn parse_border_line_style_input(value: &str) -> PyResult<CoreBorderLineStyle> {
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
        other => Err(style_input_error(format!("unknown border style {other:?}"))),
    }
}

fn parse_diagonal_direction_input(value: &str) -> PyResult<DiagonalDirection> {
    match value {
        "none" => Ok(DiagonalDirection::None),
        "down" => Ok(DiagonalDirection::Down),
        "up" => Ok(DiagonalDirection::Up),
        "both" => Ok(DiagonalDirection::Both),
        other => Err(style_input_error(format!(
            "unknown diagonal_direction {other:?}"
        ))),
    }
}

fn py_border_edge_to_core(edge: &PyBorderEdge) -> PyResult<Option<CoreBorderEdge>> {
    let style = parse_border_line_style_input(&edge.style)?;
    if style == CoreBorderLineStyle::None {
        Ok(None)
    } else {
        Ok(Some(CoreBorderEdge::new(
            style,
            py_color_to_core(&edge.color)?,
        )))
    }
}

fn apply_border_edge_input(
    input: &Bound<'_, PyAny>,
    existing: Option<&CoreBorderEdge>,
) -> PyResult<Option<CoreBorderEdge>> {
    if let Ok(py_edge) = input.extract::<PyRef<PyBorderEdge>>() {
        return py_border_edge_to_core(&py_edge);
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("border edge must be a BorderEdge or dict"))?;
    let parsed_style = dict_get_string(dict, "style")?
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
    if let Some(color) = dict_get(dict, "color")? {
        edge.color = color_input_to_core(&color)?;
    }
    Ok(Some(edge))
}

fn py_border_to_core(border: &PyBorderStyle) -> PyResult<CoreBorderStyle> {
    Ok(CoreBorderStyle {
        left: border
            .left
            .as_ref()
            .map(py_border_edge_to_core)
            .transpose()?
            .flatten(),
        right: border
            .right
            .as_ref()
            .map(py_border_edge_to_core)
            .transpose()?
            .flatten(),
        top: border
            .top
            .as_ref()
            .map(py_border_edge_to_core)
            .transpose()?
            .flatten(),
        bottom: border
            .bottom
            .as_ref()
            .map(py_border_edge_to_core)
            .transpose()?
            .flatten(),
        diagonal: border
            .diagonal
            .as_ref()
            .map(py_border_edge_to_core)
            .transpose()?
            .flatten(),
        diagonal_direction: parse_diagonal_direction_input(&border.diagonal_direction)?,
    })
}

fn apply_border_dict(dict: &Bound<'_, PyDict>, border: &mut CoreBorderStyle) -> PyResult<()> {
    if let Some(edge) = dict_get(dict, "left")? {
        border.left = apply_border_edge_input(&edge, border.left.as_ref())?;
    }
    if let Some(edge) = dict_get(dict, "right")? {
        border.right = apply_border_edge_input(&edge, border.right.as_ref())?;
    }
    if let Some(edge) = dict_get(dict, "top")? {
        border.top = apply_border_edge_input(&edge, border.top.as_ref())?;
    }
    if let Some(edge) = dict_get(dict, "bottom")? {
        border.bottom = apply_border_edge_input(&edge, border.bottom.as_ref())?;
    }
    if let Some(edge) = dict_get(dict, "diagonal")? {
        border.diagonal = apply_border_edge_input(&edge, border.diagonal.as_ref())?;
    }
    if let Some(direction) = dict_get_string(dict, "diagonal_direction")? {
        border.diagonal_direction = parse_diagonal_direction_input(&direction)?;
    }
    Ok(())
}

fn apply_border_input(input: &Bound<'_, PyAny>, border: &mut CoreBorderStyle) -> PyResult<()> {
    if let Ok(py_border) = input.extract::<PyRef<PyBorderStyle>>() {
        *border = py_border_to_core(&py_border)?;
        return Ok(());
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("border must be a BorderStyle or dict"))?;
    if dict_has(dict, "diagonal_direction")? {
        let mut next = CoreBorderStyle::default();
        apply_border_dict(dict, &mut next)?;
        *border = next;
    } else {
        apply_border_dict(dict, border)?;
    }
    Ok(())
}

pub(crate) fn parse_horizontal_alignment_input(value: &str) -> PyResult<HorizontalAlignment> {
    match value {
        "general" => Ok(HorizontalAlignment::General),
        "left" => Ok(HorizontalAlignment::Left),
        "center" => Ok(HorizontalAlignment::Center),
        "right" => Ok(HorizontalAlignment::Right),
        "fill" => Ok(HorizontalAlignment::Fill),
        "justify" => Ok(HorizontalAlignment::Justify),
        "centerContinuous" => Ok(HorizontalAlignment::CenterContinuous),
        "distributed" => Ok(HorizontalAlignment::Distributed),
        other => Err(style_input_error(format!(
            "unknown horizontal alignment {other:?}"
        ))),
    }
}

pub(crate) fn parse_vertical_alignment_input(value: &str) -> PyResult<VerticalAlignment> {
    match value {
        "top" => Ok(VerticalAlignment::Top),
        "center" => Ok(VerticalAlignment::Center),
        "bottom" => Ok(VerticalAlignment::Bottom),
        "justify" => Ok(VerticalAlignment::Justify),
        "distributed" => Ok(VerticalAlignment::Distributed),
        other => Err(style_input_error(format!(
            "unknown vertical alignment {other:?}"
        ))),
    }
}

fn parse_reading_order_input(value: &str) -> PyResult<ReadingOrder> {
    match value {
        "contextDependent" => Ok(ReadingOrder::ContextDependent),
        "leftToRight" => Ok(ReadingOrder::LeftToRight),
        "rightToLeft" => Ok(ReadingOrder::RightToLeft),
        other => Err(style_input_error(format!(
            "unknown reading_order {other:?}"
        ))),
    }
}

fn py_alignment_to_core(alignment: &PyAlignment) -> PyResult<CoreAlignment> {
    Ok(CoreAlignment {
        horizontal: parse_horizontal_alignment_input(&alignment.horizontal)?,
        vertical: parse_vertical_alignment_input(&alignment.vertical)?,
        wrap_text: alignment.wrap_text,
        shrink_to_fit: alignment.shrink_to_fit,
        indent: u32_to_u8(alignment.indent, "indent")?,
        rotation: alignment.rotation as i16,
        reading_order: parse_reading_order_input(&alignment.reading_order)?,
    })
}

fn is_full_alignment_dict(dict: &Bound<'_, PyDict>) -> PyResult<bool> {
    Ok(dict_has(dict, "horizontal")?
        && dict_has(dict, "vertical")?
        && dict_has(dict, "wrap_text")?
        && dict_has(dict, "shrink_to_fit")?
        && dict_has(dict, "indent")?
        && dict_has(dict, "rotation")?
        && dict_has(dict, "reading_order")?)
}

fn apply_alignment_dict(dict: &Bound<'_, PyDict>, alignment: &mut CoreAlignment) -> PyResult<()> {
    if let Some(horizontal) = dict_get_string(dict, "horizontal")? {
        alignment.horizontal = parse_horizontal_alignment_input(&horizontal)?;
    }
    if let Some(vertical) = dict_get_string(dict, "vertical")? {
        alignment.vertical = parse_vertical_alignment_input(&vertical)?;
    }
    if let Some(wrap_text) = dict_get_bool(dict, "wrap_text")? {
        alignment.wrap_text = wrap_text;
    }
    if let Some(shrink_to_fit) = dict_get_bool(dict, "shrink_to_fit")? {
        alignment.shrink_to_fit = shrink_to_fit;
    }
    if let Some(indent) = dict_get_u32(dict, "indent")? {
        alignment.indent = u32_to_u8(indent, "indent")?;
    }
    if let Some(rotation) = dict_get_i32(dict, "rotation")? {
        if !((-90..=90).contains(&rotation) || rotation == 255) {
            return Err(style_input_error(
                "rotation must be between -90 and 90, or 255",
            ));
        }
        alignment.rotation = rotation as i16;
    }
    if let Some(reading_order) = dict_get_string(dict, "reading_order")? {
        alignment.reading_order = parse_reading_order_input(&reading_order)?;
    }
    Ok(())
}

fn apply_alignment_input(input: &Bound<'_, PyAny>, alignment: &mut CoreAlignment) -> PyResult<()> {
    if let Ok(py_alignment) = input.extract::<PyRef<PyAlignment>>() {
        *alignment = py_alignment_to_core(&py_alignment)?;
        return Ok(());
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("alignment must be an Alignment or dict"))?;
    if is_full_alignment_dict(dict)? {
        let mut next = CoreAlignment::default();
        apply_alignment_dict(dict, &mut next)?;
        *alignment = next;
    } else {
        apply_alignment_dict(dict, alignment)?;
    }
    Ok(())
}

fn py_number_format_to_core(number_format: &PyNumberFormat) -> PyResult<CoreNumberFormat> {
    match number_format.format_type.as_str() {
        "general" => Ok(CoreNumberFormat::General),
        "builtin" => Ok(CoreNumberFormat::BuiltIn(number_format.id.ok_or_else(
            || style_input_error("builtin number format requires id"),
        )?)),
        "custom" => Ok(CoreNumberFormat::Custom(
            number_format.format_string.clone(),
        )),
        other => Err(style_input_error(format!("unknown format_type {other:?}"))),
    }
}

fn number_format_dict_to_core(dict: &Bound<'_, PyDict>) -> PyResult<CoreNumberFormat> {
    match dict_get_string(dict, "format_type")?.as_deref() {
        Some("general") => Ok(CoreNumberFormat::General),
        Some("builtin") => Ok(CoreNumberFormat::BuiltIn(
            dict_get_u32(dict, "id")?
                .ok_or_else(|| style_input_error("builtin number format requires id"))?,
        )),
        Some("custom") => Ok(CoreNumberFormat::Custom(
            dict_get_string(dict, "format_string")?
                .ok_or_else(|| style_input_error("custom number format requires format_string"))?,
        )),
        Some(other) => Err(style_input_error(format!("unknown format_type {other:?}"))),
        None if dict_get_u32(dict, "id")?.is_some() => Ok(CoreNumberFormat::BuiltIn(
            dict_get_u32(dict, "id")?.unwrap(),
        )),
        None if dict_get_string(dict, "format_string")?.is_some() => Ok(CoreNumberFormat::Custom(
            dict_get_string(dict, "format_string")?.unwrap(),
        )),
        None => Err(style_input_error(
            "number_format requires format_type, id, or format_string",
        )),
    }
}

fn apply_number_format_input(
    input: &Bound<'_, PyAny>,
    number_format: &mut CoreNumberFormat,
) -> PyResult<()> {
    if let Ok(py_number_format) = input.extract::<PyRef<PyNumberFormat>>() {
        *number_format = py_number_format_to_core(&py_number_format)?;
        return Ok(());
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("number_format must be a NumberFormat or dict"))?;
    *number_format = number_format_dict_to_core(dict)?;
    Ok(())
}

fn py_style_to_core(style: &PyStyle) -> PyResult<CoreStyle> {
    Ok(CoreStyle {
        font: py_font_to_core(&style.font)?,
        fill: py_fill_to_core(&style.fill)?,
        border: py_border_to_core(&style.border)?,
        alignment: py_alignment_to_core(&style.alignment)?,
        number_format: py_number_format_to_core(&style.number_format)?,
        protection: core::style::Protection {
            locked: style.protection.locked,
            hidden: style.protection.hidden,
        },
    })
}

fn apply_protection_input(
    input: &Bound<'_, PyAny>,
    protection: &mut core::style::Protection,
) -> PyResult<()> {
    if let Ok(py_protection) = input.extract::<PyRef<PyCellProtection>>() {
        protection.locked = py_protection.locked;
        protection.hidden = py_protection.hidden;
        return Ok(());
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("protection must be a CellProtection or dict"))?;
    if let Some(locked) = dict_get_bool(dict, "locked")? {
        protection.locked = locked;
    }
    if let Some(hidden) = dict_get_bool(dict, "hidden")? {
        protection.hidden = hidden;
    }
    Ok(())
}

pub(crate) fn apply_style_input_to_core(
    input: &Bound<'_, PyAny>,
    style: &mut CoreStyle,
) -> PyResult<()> {
    if let Ok(py_style) = input.extract::<PyRef<PyStyle>>() {
        *style = py_style_to_core(&py_style)?;
        return Ok(());
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| style_input_error("style must be a Style or dict"))?;

    if let Some(font) = dict_get(dict, "font")? {
        apply_font_input(&font, &mut style.font)?;
    }
    if let Some(fill) = dict_get(dict, "fill")? {
        apply_fill_input(&fill, &mut style.fill)?;
    }
    if let Some(border) = dict_get(dict, "border")? {
        apply_border_input(&border, &mut style.border)?;
    }
    if let Some(alignment) = dict_get(dict, "alignment")? {
        apply_alignment_input(&alignment, &mut style.alignment)?;
    }
    if let Some(number_format) = dict_get(dict, "number_format")? {
        apply_number_format_input(&number_format, &mut style.number_format)?;
    }
    if let Some(protection) = dict_get(dict, "protection")? {
        apply_protection_input(&protection, &mut style.protection)?;
    }
    Ok(())
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
}

impl From<&core::CellComment> for PyComment {
    fn from(c: &core::CellComment) -> Self {
        Self {
            author: c.author.clone(),
            text: c.plain_text(),
        }
    }
}

#[pymethods]
impl PyComment {
    #[new]
    #[pyo3(signature=(text, *, author=None))]
    fn new(text: String, author: Option<String>) -> Self {
        Self {
            author: author.unwrap_or_default(),
            text,
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
    pub password_hash: Option<u32>,
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
            password_hash: p.password_hash.map(|h| h as u32),
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

#[pyclass(name = "WorkbookProtection")]
#[derive(Clone)]
pub struct PyWorkbookProtection {
    #[pyo3(get)]
    pub structure: bool,
    #[pyo3(get)]
    pub windows: bool,
    #[pyo3(get)]
    pub password_hash: Option<u32>,
}

impl From<&core::WorkbookProtection> for PyWorkbookProtection {
    fn from(p: &core::WorkbookProtection) -> Self {
        Self {
            structure: p.structure,
            windows: p.windows,
            password_hash: p.password_hash.map(|h| h as u32),
        }
    }
}

#[pyclass(name = "ProtectedRange")]
#[derive(Clone)]
pub struct PyProtectedRange {
    #[pyo3(get)]
    pub name: String,
    #[pyo3(get)]
    pub ranges: Vec<String>,
    #[pyo3(get)]
    pub password_hash: Option<u32>,
    #[pyo3(get)]
    pub security_descriptor: Option<String>,
}

impl From<&core::ProtectedRange> for PyProtectedRange {
    fn from(p: &core::ProtectedRange) -> Self {
        Self {
            name: p.name.clone(),
            ranges: p.ranges.iter().map(ToString::to_string).collect(),
            password_hash: p.password_hash.map(|h| h as u32),
            security_descriptor: p.security_descriptor.clone(),
        }
    }
}

fn protection_password_hash(
    password: Option<String>,
    password_hash: Option<u32>,
) -> PyResult<Option<u16>> {
    match (password, password_hash) {
        (Some(_), Some(_)) => Err(PyValueError::new_err(
            "Specify either password or password_hash, not both",
        )),
        (Some(password), None) => Ok(Some(core::hash_legacy_protection_password(&password))),
        (None, Some(hash)) => u16::try_from(hash)
            .map(Some)
            .map_err(|_| PyValueError::new_err("password_hash must be between 0 and 65535")),
        (None, None) => Ok(None),
    }
}

fn py_sheet_protection_to_core(p: &PySheetProtection) -> core::SheetProtection {
    core::SheetProtection {
        protected: p.protected,
        password_hash: p.password_hash.map(|h| h as u16),
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

pub(crate) fn sheet_protection_input_to_core(
    input: &Bound<'_, PyAny>,
) -> PyResult<Option<core::SheetProtection>> {
    if input.is_none() {
        return Ok(None);
    }
    if let Ok(py_protection) = input.extract::<PyRef<PySheetProtection>>() {
        return Ok(Some(py_sheet_protection_to_core(&py_protection)));
    }

    let dict = input.downcast::<PyDict>().map_err(|_| {
        PyValueError::new_err("protection must be a dict, SheetProtection, or None")
    })?;
    Ok(Some(core::SheetProtection {
        protected: dict_get_bool(dict, "protected")?.unwrap_or(true),
        password_hash: protection_password_hash(
            dict_get_string(dict, "password")?,
            dict_get_u32_alias(dict, "password_hash", "passwordHash")?,
        )?,
        select_locked_cells: dict_get_bool(dict, "select_locked_cells")?.unwrap_or(true),
        select_unlocked_cells: dict_get_bool(dict, "select_unlocked_cells")?.unwrap_or(true),
        format_cells: dict_get_bool(dict, "format_cells")?.unwrap_or(false),
        format_columns: dict_get_bool(dict, "format_columns")?.unwrap_or(false),
        format_rows: dict_get_bool(dict, "format_rows")?.unwrap_or(false),
        insert_columns: dict_get_bool(dict, "insert_columns")?.unwrap_or(false),
        insert_rows: dict_get_bool(dict, "insert_rows")?.unwrap_or(false),
        insert_hyperlinks: dict_get_bool(dict, "insert_hyperlinks")?.unwrap_or(false),
        delete_columns: dict_get_bool(dict, "delete_columns")?.unwrap_or(false),
        delete_rows: dict_get_bool(dict, "delete_rows")?.unwrap_or(false),
        sort: dict_get_bool(dict, "sort")?.unwrap_or(false),
        auto_filter: dict_get_bool(dict, "auto_filter")?.unwrap_or(false),
        pivot_tables: dict_get_bool(dict, "pivot_tables")?.unwrap_or(false),
    }))
}

pub(crate) fn workbook_protection_input_to_core(
    input: &Bound<'_, PyAny>,
) -> PyResult<Option<core::WorkbookProtection>> {
    if input.is_none() {
        return Ok(None);
    }
    if let Ok(py_protection) = input.extract::<PyRef<PyWorkbookProtection>>() {
        return Ok(Some(core::WorkbookProtection {
            structure: py_protection.structure,
            windows: py_protection.windows,
            password_hash: py_protection.password_hash.map(|h| h as u16),
        }));
    }

    let dict = input.downcast::<PyDict>().map_err(|_| {
        PyValueError::new_err("workbook protection must be a dict, WorkbookProtection, or None")
    })?;
    Ok(Some(core::WorkbookProtection {
        structure: dict_get_bool(dict, "structure")?.unwrap_or(true),
        windows: dict_get_bool(dict, "windows")?.unwrap_or(false),
        password_hash: protection_password_hash(
            dict_get_string(dict, "password")?,
            dict_get_u32_alias(dict, "password_hash", "passwordHash")?,
        )?,
    }))
}

fn protected_range_input_to_core(input: &Bound<'_, PyAny>) -> PyResult<core::ProtectedRange> {
    if let Ok(py_range) = input.extract::<PyRef<PyProtectedRange>>() {
        let mut ranges = Vec::with_capacity(py_range.ranges.len());
        for range in &py_range.ranges {
            ranges.push(core::CellRange::parse(range).map_err(|e| {
                PyValueError::new_err(format!("Invalid protected range '{}': {}", range, e))
            })?);
        }
        return Ok(core::ProtectedRange {
            name: py_range.name.clone(),
            ranges,
            password_hash: py_range.password_hash.map(|h| h as u16),
            security_descriptor: py_range.security_descriptor.clone(),
        });
    }

    let dict = input
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("protected range must be a dict or ProtectedRange"))?;
    let name = dict_get_string(dict, "name")?
        .ok_or_else(|| PyValueError::new_err("protected range requires name"))?;
    let ranges_any = dict_get(dict, "ranges")?
        .ok_or_else(|| PyValueError::new_err("protected range requires ranges"))?;
    let range_strings = ranges_any.extract::<Vec<String>>()?;
    let mut ranges = Vec::with_capacity(range_strings.len());
    for range in range_strings {
        ranges.push(core::CellRange::parse(&range).map_err(|e| {
            PyValueError::new_err(format!("Invalid protected range '{}': {}", range, e))
        })?);
    }
    Ok(core::ProtectedRange {
        name,
        ranges,
        password_hash: protection_password_hash(
            dict_get_string(dict, "password")?,
            dict_get_u32_alias(dict, "password_hash", "passwordHash")?,
        )?,
        security_descriptor: dict_get_string(dict, "security_descriptor")?,
    })
}

pub(crate) fn protected_ranges_input_to_core(
    input: &Bound<'_, PyAny>,
) -> PyResult<Vec<core::ProtectedRange>> {
    let list = input
        .downcast::<PyList>()
        .map_err(|_| PyValueError::new_err("protected ranges must be a list"))?;
    list.iter()
        .map(|item| protected_range_input_to_core(&item))
        .collect()
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

impl PyRichTextRun {
    pub(crate) fn to_core(&self) -> PyResult<core::RichTextRun> {
        Ok(core::RichTextRun {
            text: self.text.clone(),
            font: self.font.as_ref().map(PyRunFont::to_core).transpose()?,
        })
    }
}

#[pymethods]
impl PyRichTextRun {
    #[new]
    #[pyo3(signature=(text, *, font=None))]
    fn new(text: String, font: Option<PyRef<'_, PyRunFont>>) -> Self {
        Self {
            text,
            font: font.map(|font| font.clone()),
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
    #[pyo3(get)]
    pub family: Option<u8>,
    #[pyo3(get)]
    pub charset: Option<u8>,
    #[pyo3(get)]
    pub scheme: Option<String>,
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
            family: f.family,
            charset: f.charset,
            scheme: f.scheme.clone(),
        }
    }
}

impl PyRunFont {
    pub(crate) fn to_core(&self) -> PyResult<core::RunFont> {
        let underline = self
            .underline
            .as_deref()
            .map(|value| match value {
                "none" => Ok(Underline::None),
                "single" => Ok(Underline::Single),
                "double" => Ok(Underline::Double),
                "singleAccounting" => Ok(Underline::SingleAccounting),
                "doubleAccounting" => Ok(Underline::DoubleAccounting),
                _ => Err(PyValueError::new_err(
                    "underline must be none, single, double, singleAccounting, or doubleAccounting",
                )),
            })
            .transpose()?;
        let vertical_align = self
            .vertical_align
            .as_deref()
            .map(|value| match value {
                "baseline" => Ok(FontVerticalAlign::Baseline),
                "superscript" => Ok(FontVerticalAlign::Superscript),
                "subscript" => Ok(FontVerticalAlign::Subscript),
                _ => Err(PyValueError::new_err(
                    "vertical_align must be baseline, superscript, or subscript",
                )),
            })
            .transpose()?;
        Ok(core::RunFont {
            bold: self.bold,
            italic: self.italic,
            size: self.size,
            color: self.color.as_ref().map(py_color_to_core).transpose()?,
            name: self.name.clone(),
            underline,
            strikethrough: self.strikethrough,
            vertical_align,
            family: self.family,
            charset: self.charset,
            scheme: self.scheme.clone(),
        })
    }
}

#[pymethods]
impl PyRunFont {
    #[new]
    #[pyo3(signature=(*, bold=None, italic=None, size=None, color=None, name=None, underline=None, strikethrough=None, vertical_align=None, family=None, charset=None, scheme=None))]
    fn new(
        bold: Option<bool>,
        italic: Option<bool>,
        size: Option<f64>,
        color: Option<&Bound<'_, PyAny>>,
        name: Option<String>,
        underline: Option<String>,
        strikethrough: Option<bool>,
        vertical_align: Option<String>,
        family: Option<u8>,
        charset: Option<u8>,
        scheme: Option<String>,
    ) -> PyResult<Self> {
        let font = Self {
            bold,
            italic,
            size,
            color: color.map(color_input_to_core).transpose()?.as_ref().map(PyColor::from),
            name,
            underline,
            strikethrough,
            vertical_align,
            family,
            charset,
            scheme,
        };
        font.to_core()?;
        Ok(font)
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

/// A drawing anchor in its native variant: two-cell, one-cell, or
/// absolute. Anchors read from files keep their variant; writers
/// flatten only when a target format requires it.
#[pyclass(name = "DrawingAnchor")]
#[derive(Clone)]
pub struct PyDrawingAnchor {
    pub(crate) inner: duke_sheets::DrawingAnchor,
}

impl From<&duke_sheets::DrawingAnchor> for PyDrawingAnchor {
    fn from(a: &duke_sheets::DrawingAnchor) -> Self {
        Self { inner: a.clone() }
    }
}

impl PyDrawingAnchor {
    pub(crate) fn to_core(&self) -> PyResult<duke_sheets::DrawingAnchor> {
        Ok(self.inner.clone())
    }

    fn from_marker(&self) -> Option<&duke_sheets::CellMarker> {
        match &self.inner {
            duke_sheets::DrawingAnchor::TwoCell { from, .. }
            | duke_sheets::DrawingAnchor::OneCell { from, .. } => Some(from),
            duke_sheets::DrawingAnchor::Absolute { .. } => None,
        }
    }

    fn to_marker(&self) -> Option<&duke_sheets::CellMarker> {
        match &self.inner {
            duke_sheets::DrawingAnchor::TwoCell { to, .. } => Some(to),
            _ => None,
        }
    }
}

fn parse_edit_as(value: &str) -> PyResult<duke_sheets::EditAs> {
    match value {
        "two_cell" => Ok(duke_sheets::EditAs::TwoCell),
        "one_cell" => Ok(duke_sheets::EditAs::OneCell),
        "absolute" => Ok(duke_sheets::EditAs::Absolute),
        other => Err(PyValueError::new_err(format!(
            "invalid edit_as {other:?}; expected two_cell, one_cell, or absolute"
        ))),
    }
}

#[pymethods]
impl PyDrawingAnchor {
    /// Two-cell anchor (the default Excel drawing placement).
    #[new]
    #[pyo3(signature=(from_row, from_col, to_row, to_col, *, from_row_offset=0, from_col_offset=0, to_row_offset=0, to_col_offset=0, edit_as="two_cell"))]
    fn new(
        from_row: u32,
        from_col: u16,
        to_row: u32,
        to_col: u16,
        from_row_offset: i64,
        from_col_offset: i64,
        to_row_offset: i64,
        to_col_offset: i64,
        edit_as: &str,
    ) -> PyResult<Self> {
        Self::two_cell(
            from_row,
            from_col,
            to_row,
            to_col,
            from_row_offset,
            from_col_offset,
            to_row_offset,
            to_col_offset,
            edit_as,
        )
    }

    /// Anchor spanning two cell markers; `edit_as` records the sizing
    /// behavior Excel applies when cells move or resize.
    #[staticmethod]
    #[pyo3(signature=(from_row, from_col, to_row, to_col, *, from_row_offset=0, from_col_offset=0, to_row_offset=0, to_col_offset=0, edit_as="two_cell"))]
    fn two_cell(
        from_row: u32,
        from_col: u16,
        to_row: u32,
        to_col: u16,
        from_row_offset: i64,
        from_col_offset: i64,
        to_row_offset: i64,
        to_col_offset: i64,
        edit_as: &str,
    ) -> PyResult<Self> {
        let edit_as = parse_edit_as(edit_as)?;
        Ok(Self {
            inner: duke_sheets::DrawingAnchor::TwoCell {
                from: duke_sheets::CellMarker {
                    col: from_col,
                    col_offset_emu: from_col_offset,
                    row: from_row,
                    row_offset_emu: from_row_offset,
                },
                to: duke_sheets::CellMarker {
                    col: to_col,
                    col_offset_emu: to_col_offset,
                    row: to_row,
                    row_offset_emu: to_row_offset,
                },
                edit_as: Some(edit_as),
            },
        })
    }

    /// Anchor pinned to one cell with a fixed EMU extent.
    #[staticmethod]
    #[pyo3(signature=(from_row, from_col, *, width_emu, height_emu, from_row_offset=0, from_col_offset=0))]
    fn one_cell(
        from_row: u32,
        from_col: u16,
        width_emu: i64,
        height_emu: i64,
        from_row_offset: i64,
        from_col_offset: i64,
    ) -> PyResult<Self> {
        if width_emu < 0 || height_emu < 0 {
            return Err(PyValueError::new_err(
                "drawing anchor dimensions cannot be negative",
            ));
        }
        Ok(Self {
            inner: duke_sheets::DrawingAnchor::OneCell {
                from: duke_sheets::CellMarker {
                    col: from_col,
                    col_offset_emu: from_col_offset,
                    row: from_row,
                    row_offset_emu: from_row_offset,
                },
                width_emu,
                height_emu,
            },
        })
    }

    /// Anchor at a fixed EMU position and extent, independent of cells.
    #[staticmethod]
    fn absolute(x_emu: i64, y_emu: i64, width_emu: i64, height_emu: i64) -> PyResult<Self> {
        if x_emu < 0 || y_emu < 0 || width_emu < 0 || height_emu < 0 {
            return Err(PyValueError::new_err(
                "absolute drawing anchor values cannot be negative",
            ));
        }
        Ok(Self {
            inner: duke_sheets::DrawingAnchor::Absolute {
                x_emu,
                y_emu,
                width_emu,
                height_emu,
            },
        })
    }

    /// "two_cell", "one_cell", or "absolute".
    #[getter]
    fn anchor_type(&self) -> &'static str {
        match &self.inner {
            duke_sheets::DrawingAnchor::TwoCell { .. } => "two_cell",
            duke_sheets::DrawingAnchor::OneCell { .. } => "one_cell",
            duke_sheets::DrawingAnchor::Absolute { .. } => "absolute",
        }
    }

    /// From-marker row (two-cell and one-cell anchors).
    #[getter]
    fn from_row(&self) -> Option<u32> {
        self.from_marker().map(|marker| marker.row)
    }

    /// From-marker column (two-cell and one-cell anchors).
    #[getter]
    fn from_col(&self) -> Option<u16> {
        self.from_marker().map(|marker| marker.col)
    }

    #[getter]
    fn from_row_offset(&self) -> Option<i64> {
        self.from_marker().map(|marker| marker.row_offset_emu)
    }

    #[getter]
    fn from_col_offset(&self) -> Option<i64> {
        self.from_marker().map(|marker| marker.col_offset_emu)
    }

    /// To-marker row (two-cell anchors only).
    #[getter]
    fn to_row(&self) -> Option<u32> {
        self.to_marker().map(|marker| marker.row)
    }

    /// To-marker column (two-cell anchors only).
    #[getter]
    fn to_col(&self) -> Option<u16> {
        self.to_marker().map(|marker| marker.col)
    }

    #[getter]
    fn to_row_offset(&self) -> Option<i64> {
        self.to_marker().map(|marker| marker.row_offset_emu)
    }

    #[getter]
    fn to_col_offset(&self) -> Option<i64> {
        self.to_marker().map(|marker| marker.col_offset_emu)
    }

    /// "two_cell", "one_cell", or "absolute" when the two-cell anchor
    /// carries an explicit editing behavior; None otherwise.
    #[getter]
    fn edit_as(&self) -> Option<&'static str> {
        match &self.inner {
            duke_sheets::DrawingAnchor::TwoCell { edit_as, .. } => {
                edit_as.as_ref().map(|edit_as| match edit_as {
                    duke_sheets::EditAs::TwoCell => "two_cell",
                    duke_sheets::EditAs::OneCell => "one_cell",
                    duke_sheets::EditAs::Absolute => "absolute",
                })
            }
            _ => None,
        }
    }

    /// Fixed EMU width (one-cell and absolute anchors).
    #[getter]
    fn width_emu(&self) -> Option<i64> {
        match &self.inner {
            duke_sheets::DrawingAnchor::OneCell { width_emu, .. }
            | duke_sheets::DrawingAnchor::Absolute { width_emu, .. } => Some(*width_emu),
            _ => None,
        }
    }

    /// Fixed EMU height (one-cell and absolute anchors).
    #[getter]
    fn height_emu(&self) -> Option<i64> {
        match &self.inner {
            duke_sheets::DrawingAnchor::OneCell { height_emu, .. }
            | duke_sheets::DrawingAnchor::Absolute { height_emu, .. } => Some(*height_emu),
            _ => None,
        }
    }

    /// Absolute EMU x position (absolute anchors only).
    #[getter]
    fn x_emu(&self) -> Option<i64> {
        match &self.inner {
            duke_sheets::DrawingAnchor::Absolute { x_emu, .. } => Some(*x_emu),
            _ => None,
        }
    }

    /// Absolute EMU y position (absolute anchors only).
    #[getter]
    fn y_emu(&self) -> Option<i64> {
        match &self.inner {
            duke_sheets::DrawingAnchor::Absolute { y_emu, .. } => Some(*y_emu),
            _ => None,
        }
    }

    fn __repr__(&self) -> String {
        format!("DrawingAnchor(anchor_type={:?})", self.anchor_type())
    }
}

#[pyclass(name = "FormControl")]
#[derive(Clone)]
pub struct PyFormControl {
    pub(crate) inner: core::FormControl,
}

impl From<&core::FormControl> for PyFormControl {
    fn from(control: &core::FormControl) -> Self {
        Self {
            inner: control.clone(),
        }
    }
}

impl PyFormControl {
    fn create(
        kind: core::FormControlKind,
        macro_name: Option<String>,
    ) -> PyResult<Self> {
        let mut control = core::FormControl::new(kind);
        control.macro_name = macro_name;
        control
            .validate()
            .map_err(|err| PyValueError::new_err(err.to_string()))?;
        Ok(Self { inner: control })
    }
}

#[pymethods]
impl PyFormControl {
    #[staticmethod]
    #[pyo3(signature=(caption, *, macro_name=None))]
    fn button(caption: &Bound<'_, PyAny>, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::Button {
                caption: drawing_text_input_to_core(caption)?,
            },
            macro_name,
        )
    }

    #[staticmethod]
    #[pyo3(signature=(caption, *, state="unchecked", cell_link=None, no_3d=false, macro_name=None))]
    fn checkbox(caption: &Bound<'_, PyAny>, state: &str, cell_link: Option<String>, no_3d: bool, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::Checkbox {
                caption: drawing_text_input_to_core(caption)?,
                state: parse_check_state(state)?,
                cell_link,
                no_3d,
            },
            macro_name,
        )
    }

    /// Create an option (radio) button. Radio grouping is derived
    /// from group-box containment when writing, so there is no
    /// grouping argument; `first_in_group` is read-side information.
    #[staticmethod]
    #[pyo3(signature=(caption, *, state="unchecked", cell_link=None, no_3d=false, macro_name=None))]
    fn option_button(caption: &Bound<'_, PyAny>, state: &str, cell_link: Option<String>, no_3d: bool, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::OptionButton {
                caption: drawing_text_input_to_core(caption)?,
                state: parse_check_state(state)?,
                cell_link,
                first_in_group: false,
                no_3d,
            },
            macro_name,
        )
    }

    #[staticmethod]
    #[pyo3(signature=(caption, *, macro_name=None))]
    fn label(caption: &Bound<'_, PyAny>, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::Label {
                caption: drawing_text_input_to_core(caption)?,
            },
            macro_name,
        )
    }

    #[staticmethod]
    #[pyo3(signature=(caption, *, no_3d=false, macro_name=None))]
    fn group_box(caption: &Bound<'_, PyAny>, no_3d: bool, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::GroupBox {
                caption: drawing_text_input_to_core(caption)?,
                no_3d,
            },
            macro_name,
        )
    }

    /// Create a list box. `selected` holds zero-based item indexes,
    /// sorted ascending; a linked cell still receives Excel's
    /// one-based value at runtime.
    #[staticmethod]
    #[pyo3(signature=(*, input_range=None, cell_link=None, selection="single", selected=None, no_3d=false, macro_name=None))]
    fn list_box(input_range: Option<String>, cell_link: Option<String>, selection: &str, selected: Option<Vec<u16>>, no_3d: bool, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::ListBox {
                input_range,
                cell_link,
                selection: parse_list_selection(selection)?,
                selected: selected.unwrap_or_default(),
                no_3d,
            },
            macro_name,
        )
    }

    /// Create a dropdown (combo box). `selected` is a zero-based item
    /// index; a linked cell still receives Excel's one-based value at
    /// runtime.
    #[staticmethod]
    #[pyo3(signature=(*, input_range=None, cell_link=None, selected=None, lines=8, no_3d=false, macro_name=None))]
    fn dropdown(input_range: Option<String>, cell_link: Option<String>, selected: Option<u16>, lines: u16, no_3d: bool, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::Dropdown {
                input_range,
                cell_link,
                selected,
                lines,
                no_3d,
            },
            macro_name,
        )
    }

    #[staticmethod]
    #[pyo3(signature=(*, value, min, max, increment=1, page=1, horizontal=false, cell_link=None, macro_name=None))]
    fn scrollbar(value: u16, min: u16, max: u16, increment: u16, page: u16, horizontal: bool, cell_link: Option<String>, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::Scrollbar {
                value,
                min,
                max,
                increment,
                page,
                horizontal,
                cell_link,
            },
            macro_name,
        )
    }

    #[staticmethod]
    #[pyo3(signature=(*, value, min, max, increment=1, cell_link=None, macro_name=None))]
    fn spinner(value: u16, min: u16, max: u16, increment: u16, cell_link: Option<String>, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::Spinner {
                value,
                min,
                max,
                increment,
                cell_link,
            },
            macro_name,
        )
    }

    #[staticmethod]
    #[pyo3(signature=(object_type, caption, *, legacy_object_type=None, macro_name=None))]
    fn unknown(object_type: String, caption: &Bound<'_, PyAny>, legacy_object_type: Option<u16>, macro_name: Option<String>) -> PyResult<Self> {
        Self::create(
            core::FormControlKind::Unknown {
                object_type,
                legacy_object_type,
                caption: drawing_text_input_to_core(caption)?,
            },
            macro_name,
        )
    }

    #[getter]
    fn kind(&self) -> &'static str { form_control_kind_name(&self.inner.kind) }
    #[getter]
    fn macro_name(&self) -> Option<String> { self.inner.macro_name.clone() }
    #[getter]
    fn caption(&self) -> Option<PyDrawingText> { self.inner.caption().map(PyDrawingText::from) }
    #[getter]
    fn caption_text(&self) -> Option<String> { self.inner.caption_text() }
    #[getter]
    fn state(&self) -> Option<&'static str> { form_control_state(&self.inner.kind) }
    #[getter]
    fn cell_link(&self) -> Option<String> { self.inner.cell_link().map(str::to_string) }
    #[getter]
    fn input_range(&self) -> Option<String> { match &self.inner.kind { core::FormControlKind::ListBox { input_range, .. } | core::FormControlKind::Dropdown { input_range, .. } => input_range.clone(), _ => None } }
    #[getter]
    fn no_3d(&self) -> Option<bool> { match &self.inner.kind { core::FormControlKind::Checkbox { no_3d, .. } | core::FormControlKind::OptionButton { no_3d, .. } | core::FormControlKind::GroupBox { no_3d, .. } | core::FormControlKind::ListBox { no_3d, .. } | core::FormControlKind::Dropdown { no_3d, .. } => Some(*no_3d), _ => None } }
    /// Whether this radio heads its group (writers recompute this
    /// from group-box containment).
    #[getter]
    fn first_in_group(&self) -> Option<bool> { match &self.inner.kind { core::FormControlKind::OptionButton { first_in_group, .. } => Some(*first_in_group), _ => None } }
    #[getter]
    fn selection(&self) -> Option<&'static str> { match &self.inner.kind { core::FormControlKind::ListBox { selection, .. } => Some(match selection { core::ListSelection::Single => "single", core::ListSelection::Multi => "multi", core::ListSelection::Extend => "extend" }), _ => None } }
    /// Zero-based selected item indexes (list boxes and dropdowns).
    #[getter]
    fn selected(&self) -> Option<Vec<u16>> { match &self.inner.kind { core::FormControlKind::ListBox { selected, .. } => Some(selected.clone()), core::FormControlKind::Dropdown { selected, .. } => Some(selected.iter().copied().collect()), _ => None } }
    #[getter]
    fn lines(&self) -> Option<u16> { match &self.inner.kind { core::FormControlKind::Dropdown { lines, .. } => Some(*lines), _ => None } }
    #[getter]
    fn value(&self) -> Option<u16> { match &self.inner.kind { core::FormControlKind::Scrollbar { value, .. } | core::FormControlKind::Spinner { value, .. } => Some(*value), _ => None } }
    #[getter]
    fn min(&self) -> Option<u16> { match &self.inner.kind { core::FormControlKind::Scrollbar { min, .. } | core::FormControlKind::Spinner { min, .. } => Some(*min), _ => None } }
    #[getter]
    fn max(&self) -> Option<u16> { match &self.inner.kind { core::FormControlKind::Scrollbar { max, .. } | core::FormControlKind::Spinner { max, .. } => Some(*max), _ => None } }
    #[getter]
    fn increment(&self) -> Option<u16> { match &self.inner.kind { core::FormControlKind::Scrollbar { increment, .. } | core::FormControlKind::Spinner { increment, .. } => Some(*increment), _ => None } }
    #[getter]
    fn page(&self) -> Option<u16> { match &self.inner.kind { core::FormControlKind::Scrollbar { page, .. } => Some(*page), _ => None } }
    #[getter]
    fn horizontal(&self) -> Option<bool> { match &self.inner.kind { core::FormControlKind::Scrollbar { horizontal, .. } => Some(*horizontal), _ => None } }
    #[getter]
    fn object_type(&self) -> Option<String> { match &self.inner.kind { core::FormControlKind::Unknown { object_type, .. } => Some(object_type.clone()), _ => None } }
    #[getter]
    fn legacy_object_type(&self) -> Option<u16> { match &self.inner.kind { core::FormControlKind::Unknown { legacy_object_type, .. } => *legacy_object_type, _ => None } }
    /// Unmodeled XLSX `formControlPr` attributes preserved on any
    /// control kind; read-only passthrough echoed back on write.
    #[getter]
    fn raw_properties(&self) -> Vec<(String, String)> { self.inner.raw_properties.clone() }
    /// Unmodeled VML `ClientData` fragments preserved on any control
    /// kind; read-only passthrough echoed back on write.
    #[getter]
    fn raw_client_data(&self, py: Python<'_>) -> Vec<Py<PyBytes>> { self.inner.raw_client_data.iter().map(|bytes| PyBytes::new_bound(py, bytes).unbind()).collect() }
    /// Original BIFF OBJ body for XLS passthrough of unknown
    /// controls; read-only passthrough echoed back on write.
    #[getter]
    fn raw_obj(&self, py: Python<'_>) -> Option<Py<PyBytes>> { self.inner.raw_obj.as_ref().map(|bytes| PyBytes::new_bound(py, bytes).unbind()) }
}

pub(crate) fn drawing_text_input_to_core(value: &Bound<'_, PyAny>) -> PyResult<core::DrawingText> {
    if let Ok(text) = value.extract::<String>() {
        return Ok(core::DrawingText::plain(text));
    }
    value
        .extract::<PyRef<'_, PyDrawingText>>()
        .map(|text| text.inner.clone())
        .map_err(|_| PyValueError::new_err("caption must be str or DrawingText"))
}

fn parse_check_state(value: &str) -> PyResult<core::CheckState> {
    match value {
        "unchecked" => Ok(core::CheckState::Unchecked),
        "checked" => Ok(core::CheckState::Checked),
        "mixed" => Ok(core::CheckState::Mixed),
        _ => Err(PyValueError::new_err("state must be unchecked, checked, or mixed")),
    }
}

fn parse_list_selection(value: &str) -> PyResult<core::ListSelection> {
    match value {
        "single" => Ok(core::ListSelection::Single),
        "multi" => Ok(core::ListSelection::Multi),
        "extend" => Ok(core::ListSelection::Extend),
        _ => Err(PyValueError::new_err("selection must be single, multi, or extend")),
    }
}

fn form_control_kind_name(kind: &core::FormControlKind) -> &'static str {
    match kind {
        core::FormControlKind::Button { .. } => "button",
        core::FormControlKind::Checkbox { .. } => "checkbox",
        core::FormControlKind::OptionButton { .. } => "option_button",
        core::FormControlKind::Label { .. } => "label",
        core::FormControlKind::GroupBox { .. } => "group_box",
        core::FormControlKind::ListBox { .. } => "list_box",
        core::FormControlKind::Dropdown { .. } => "dropdown",
        core::FormControlKind::Scrollbar { .. } => "scrollbar",
        core::FormControlKind::Spinner { .. } => "spinner",
        core::FormControlKind::Unknown { .. } => "unknown",
    }
}

fn form_control_state(kind: &core::FormControlKind) -> Option<&'static str> {
    let state = match kind {
        core::FormControlKind::Checkbox { state, .. }
        | core::FormControlKind::OptionButton { state, .. } => state,
        _ => return None,
    };
    Some(match state {
        core::CheckState::Unchecked => "unchecked",
        core::CheckState::Checked => "checked",
        core::CheckState::Mixed => "mixed",
    })
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
            line_color_hex: sp
                .line
                .as_ref()
                .and_then(|l| l.solid_fill.as_ref().map(|c| c.hex.clone())),
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
            shape_properties: s
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
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
    /// `bottom`, `top`, `left` or `right`; unset when the source
    /// omitted it and the writer will supply the conventional one.
    pub position: Option<String>,
    #[pyo3(get)]
    pub number_format: Option<PyChartNumberFormat>,
    #[pyo3(get)]
    pub major_gridlines: bool,
    #[pyo3(get)]
    pub minor_gridlines: bool,
    #[pyo3(get)]
    pub major_gridlines_shape_properties: Option<PyChartShapeProperties>,
    #[pyo3(get)]
    pub minor_gridlines_shape_properties: Option<PyChartShapeProperties>,
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
            position: a.position.map(|position| {
                match position {
                    duke_sheets::AxisPosition::Bottom => "bottom",
                    duke_sheets::AxisPosition::Top => "top",
                    duke_sheets::AxisPosition::Left => "left",
                    duke_sheets::AxisPosition::Right => "right",
                }
                .to_string()
            }),
            number_format: a.number_format.as_ref().map(PyChartNumberFormat::from),
            major_gridlines: a.major_gridlines,
            minor_gridlines: a.minor_gridlines,
            major_gridlines_shape_properties: a
                .major_gridlines_shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
            minor_gridlines_shape_properties: a
                .minor_gridlines_shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
            major_tick_mark: a.major_tick_mark.as_ref().map(|t| {
                match t {
                    chart::TickMark::Cross => "cross",
                    chart::TickMark::Inside => "inside",
                    chart::TickMark::None => "none",
                    chart::TickMark::Outside => "outside",
                }
                .into()
            }),
            minor_tick_mark: a.minor_tick_mark.as_ref().map(|t| {
                match t {
                    chart::TickMark::Cross => "cross",
                    chart::TickMark::Inside => "inside",
                    chart::TickMark::None => "none",
                    chart::TickMark::Outside => "outside",
                }
                .into()
            }),
            label_position: a.label_position.as_ref().map(|p| {
                match p {
                    chart::TickLabelPosition::High => "high",
                    chart::TickLabelPosition::Low => "low",
                    chart::TickLabelPosition::NextTo => "nextTo",
                    chart::TickLabelPosition::None => "none",
                }
                .into()
            }),
            delete: a.delete,
            crosses: a.crosses.as_ref().map(|c| {
                match c {
                    chart::AxisCrosses::AutoZero => "autoZero",
                    chart::AxisCrosses::Min => "min",
                    chart::AxisCrosses::Max => "max",
                }
                .into()
            }),
            cross_between: a.cross_between.as_ref().map(|c| {
                match c {
                    chart::CrossBetween::Between => "between",
                    chart::CrossBetween::MidCat => "midCat",
                }
                .into()
            }),
            shape_properties: a
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
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
    pub(crate) inner: duke_sheets::Chart,
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
    pub shape_properties: Option<PyChartShapeProperties>,
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
    #[pyo3(get)]
    pub style: Option<PyChartStyle>,
    #[pyo3(get)]
    pub color_style: Option<PyChartColorStyle>,
}

impl From<&duke_sheets::Chart> for PyChart {
    fn from(c: &duke_sheets::Chart) -> Self {
        let chart_type = match &c.chart_type {
            duke_sheets::ChartType::Unsupported(tag) => format!("Unsupported({})", tag),
            other => format!("{:?}", other),
        };
        Self {
            inner: c.clone(),
            chart_type,
            title: c.title.clone(),
            series: c.series.iter().map(PyDataSeries::from).collect(),
            category_axis: c.category_axis.as_ref().map(PyAxis::from),
            value_axis: c.value_axis.as_ref().map(PyAxis::from),
            legend: c.legend.as_ref().map(PyLegend::from),
            data_labels: c.data_labels.as_ref().map(PyDataLabels::from),
            view_3d: c.view_3d.as_ref().map(PyView3D::from),
            data_table: c.data_table.as_ref().map(PyChartDataTable::from),
            display_blanks_as: c.display_blanks_as.as_ref().map(|d| {
                match d {
                    chart::DisplayBlanksAs::Gap => "gap",
                    chart::DisplayBlanksAs::Span => "span",
                    chart::DisplayBlanksAs::Zero => "zero",
                }
                .into()
            }),
            plot_visible_only: c.plot_visible_only,
            layout: c.layout.as_ref().map(PyLayout::from),
            shape_properties: c
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
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
            style: c.style.as_ref().map(PyChartStyle::from),
            color_style: c.color_style.as_ref().map(PyChartColorStyle::from),
        }
    }
}

fn chart_type_from_python(value: &str) -> PyResult<duke_sheets::ChartType> {
    use duke_sheets::ChartType;
    Ok(match value {
        "ColumnClustered" | "column_clustered" => ChartType::ColumnClustered,
        "ColumnStacked" | "column_stacked" => ChartType::ColumnStacked,
        "ColumnPercentStacked" | "column_percent_stacked" => ChartType::ColumnPercentStacked,
        "BarClustered" | "bar_clustered" => ChartType::BarClustered,
        "BarStacked" | "bar_stacked" => ChartType::BarStacked,
        "BarPercentStacked" | "bar_percent_stacked" => ChartType::BarPercentStacked,
        "Line" | "line" => ChartType::Line,
        "LineStacked" | "line_stacked" => ChartType::LineStacked,
        "Pie" | "pie" => ChartType::Pie,
        "PieExploded" | "pie_exploded" => ChartType::PieExploded,
        "Doughnut" | "doughnut" => ChartType::Doughnut,
        "Area" | "area" => ChartType::Area,
        "AreaStacked" | "area_stacked" => ChartType::AreaStacked,
        "AreaPercentStacked" | "area_percent_stacked" => ChartType::AreaPercentStacked,
        "ScatterMarkers" | "scatter_markers" => ChartType::ScatterMarkers,
        "ScatterSmooth" | "scatter_smooth" => ChartType::ScatterSmooth,
        "ScatterLines" | "scatter_lines" => ChartType::ScatterLines,
        "Bubble" | "bubble" => ChartType::Bubble,
        "Radar" | "radar" => ChartType::Radar,
        "Stock" | "stock" => ChartType::Stock,
        "Surface" | "surface" => ChartType::Surface,
        other => {
            return Err(PyValueError::new_err(format!(
                "unsupported chart type {other:?}"
            )))
        }
    })
}

#[pymethods]
impl PyChart {
    #[new]
    #[pyo3(signature=(chart_type, *, title=None))]
    fn new(chart_type: &str, title: Option<String>) -> PyResult<Self> {
        let mut chart = duke_sheets::Chart::new(chart_type_from_python(chart_type)?);
        chart.title = title;
        Ok(Self::from(&chart))
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
            position: d.position.as_ref().map(|p| {
                match p {
                    chart::DataLabelPosition::BestFit => "bestFit",
                    chart::DataLabelPosition::Bottom => "bottom",
                    chart::DataLabelPosition::Center => "center",
                    chart::DataLabelPosition::InsideBase => "insideBase",
                    chart::DataLabelPosition::InsideEnd => "insideEnd",
                    chart::DataLabelPosition::Left => "left",
                    chart::DataLabelPosition::OutsideEnd => "outsideEnd",
                    chart::DataLabelPosition::Right => "right",
                    chart::DataLabelPosition::Top => "top",
                }
                .into()
            }),
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
            }
            .into(),
            bar_type: match e.bar_type {
                chart::ErrorBarType::Both => "both",
                chart::ErrorBarType::Minus => "minus",
                chart::ErrorBarType::Plus => "plus",
            }
            .into(),
            value_type: match e.value_type {
                chart::ErrorValueType::Custom => "custom",
                chart::ErrorValueType::FixedValue => "fixedValue",
                chart::ErrorValueType::Percentage => "percentage",
                chart::ErrorValueType::StandardDeviation => "standardDeviation",
                chart::ErrorValueType::StandardError => "standardError",
            }
            .into(),
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
            symbol: m.symbol.as_ref().map(|s| {
                match s {
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
                }
                .into()
            }),
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
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::DataPoint> for PyDataPoint {
    fn from(p: &chart::DataPoint) -> Self {
        Self {
            index: p.index,
            marker: p.marker.as_ref().map(PyMarker::from),
            explosion: p.explosion,
            shape_properties: p
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
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
            shape_properties: cl
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
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

fn chart_ex_layout_to_string(layout: &chart::ChartExLayout) -> &'static str {
    match layout {
        chart::ChartExLayout::Waterfall => "waterfall",
        chart::ChartExLayout::Treemap => "treemap",
        chart::ChartExLayout::Sunburst => "sunburst",
        chart::ChartExLayout::Funnel => "funnel",
        chart::ChartExLayout::Histogram => "histogram",
        chart::ChartExLayout::BoxWhisker => "boxWhisker",
        chart::ChartExLayout::ParetoLine => "paretoLine",
        chart::ChartExLayout::RegionMap => "regionMap",
        chart::ChartExLayout::ClusteredColumn => "clusteredColumn",
        chart::ChartExLayout::Unknown(_) => "unknown",
    }
}

#[pyclass(name = "ChartExOffset")]
#[derive(Clone)]
pub struct PyChartExOffset {
    #[pyo3(get)]
    pub top: Option<f64>,
    #[pyo3(get)]
    pub left: Option<f64>,
}

impl From<&chart::ChartExOffset> for PyChartExOffset {
    fn from(o: &chart::ChartExOffset) -> Self {
        Self {
            top: o.top,
            left: o.left,
        }
    }
}

#[pyclass(name = "ChartExText")]
#[derive(Clone)]
pub struct PyChartExText {
    #[pyo3(get)]
    pub formula: Option<String>,
    #[pyo3(get)]
    pub value: Option<String>,
}

impl From<&chart::ChartExText> for PyChartExText {
    fn from(t: &chart::ChartExText) -> Self {
        Self {
            formula: t.data.as_ref().and_then(|d| d.formula.clone()),
            value: t.data.as_ref().and_then(|d| d.value.clone()),
        }
    }
}

#[pyclass(name = "ChartExColorPosition")]
#[derive(Clone)]
pub struct PyChartExColorPosition {
    #[pyo3(get)]
    pub position_type: String,
    #[pyo3(get)]
    pub value: Option<f64>,
}

impl From<&chart::ChartExColorPosition> for PyChartExColorPosition {
    fn from(p: &chart::ChartExColorPosition) -> Self {
        match p {
            chart::ChartExColorPosition::ExtremeValue => Self {
                position_type: "extremeValue".into(),
                value: None,
            },
            chart::ChartExColorPosition::Number(v) => Self {
                position_type: "number".into(),
                value: Some(*v),
            },
            chart::ChartExColorPosition::Percent(v) => Self {
                position_type: "percent".into(),
                value: Some(*v),
            },
        }
    }
}

#[pyclass(name = "ChartExValueColorPositions")]
#[derive(Clone)]
pub struct PyChartExValueColorPositions {
    #[pyo3(get)]
    pub count: Option<u32>,
    #[pyo3(get)]
    pub min: Option<PyChartExColorPosition>,
    #[pyo3(get)]
    pub mid: Option<PyChartExColorPosition>,
    #[pyo3(get)]
    pub max: Option<PyChartExColorPosition>,
}

impl From<&chart::ChartExValueColorPositions> for PyChartExValueColorPositions {
    fn from(p: &chart::ChartExValueColorPositions) -> Self {
        Self {
            count: p.count,
            min: p.min.as_ref().map(PyChartExColorPosition::from),
            mid: p.mid.as_ref().map(PyChartExColorPosition::from),
            max: p.max.as_ref().map(PyChartExColorPosition::from),
        }
    }
}

#[pyclass(name = "ChartExScaling")]
#[derive(Clone)]
pub struct PyChartExScaling {
    #[pyo3(get)]
    pub scaling_type: String,
    #[pyo3(get)]
    pub gap_width: Option<f64>,
    #[pyo3(get)]
    pub min: Option<f64>,
    #[pyo3(get)]
    pub max: Option<f64>,
    #[pyo3(get)]
    pub major_unit: Option<f64>,
    #[pyo3(get)]
    pub minor_unit: Option<f64>,
}

impl From<&chart::ChartExScaling> for PyChartExScaling {
    fn from(s: &chart::ChartExScaling) -> Self {
        match s {
            chart::ChartExScaling::Category { gap_width } => Self {
                scaling_type: "category".into(),
                gap_width: *gap_width,
                min: None,
                max: None,
                major_unit: None,
                minor_unit: None,
            },
            chart::ChartExScaling::Value {
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

#[pyclass(name = "ChartExAxisTitle")]
#[derive(Clone)]
pub struct PyChartExAxisTitle {
    #[pyo3(get)]
    pub text: Option<String>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExAxisTitle> for PyChartExAxisTitle {
    fn from(t: &chart::ChartExAxisTitle) -> Self {
        Self {
            text: t.text.as_ref().and_then(|tx| {
                tx.data
                    .as_ref()
                    .and_then(|d| d.value.clone().or_else(|| d.formula.clone()))
            }),
            shape_properties: t
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExAxisUnits")]
#[derive(Clone)]
pub struct PyChartExAxisUnits {
    #[pyo3(get)]
    pub unit: Option<String>,
}

impl From<&chart::ChartExAxisUnits> for PyChartExAxisUnits {
    fn from(u: &chart::ChartExAxisUnits) -> Self {
        Self {
            unit: u.unit.clone(),
        }
    }
}

#[pyclass(name = "ChartExSeriesVisibility")]
#[derive(Clone)]
pub struct PyChartExSeriesVisibility {
    #[pyo3(get)]
    pub connector_lines: Option<bool>,
    #[pyo3(get)]
    pub mean_line: Option<bool>,
    #[pyo3(get)]
    pub mean_marker: Option<bool>,
    #[pyo3(get)]
    pub nonoutliers: Option<bool>,
    #[pyo3(get)]
    pub outliers: Option<bool>,
}

impl From<&chart::ChartExSeriesVisibility> for PyChartExSeriesVisibility {
    fn from(v: &chart::ChartExSeriesVisibility) -> Self {
        Self {
            connector_lines: v.connector_lines,
            mean_line: v.mean_line,
            mean_marker: v.mean_marker,
            nonoutliers: v.nonoutliers,
            outliers: v.outliers,
        }
    }
}

#[pyclass(name = "ChartExBinning")]
#[derive(Clone)]
pub struct PyChartExBinning {
    #[pyo3(get)]
    pub interval_closed: Option<String>,
    #[pyo3(get)]
    pub underflow: Option<String>,
    #[pyo3(get)]
    pub overflow: Option<String>,
    #[pyo3(get)]
    pub bin_size: Option<f64>,
    #[pyo3(get)]
    pub bin_count: Option<u32>,
}

impl From<&chart::ChartExBinning> for PyChartExBinning {
    fn from(b: &chart::ChartExBinning) -> Self {
        Self {
            interval_closed: b.interval_closed.clone(),
            underflow: b.underflow.clone(),
            overflow: b.overflow.clone(),
            bin_size: b.bin_size,
            bin_count: b.bin_count,
        }
    }
}

#[pyclass(name = "ChartExGeography")]
#[derive(Clone)]
pub struct PyChartExGeography {
    #[pyo3(get)]
    pub projection_type: Option<String>,
    #[pyo3(get)]
    pub viewed_region_type: Option<String>,
    #[pyo3(get)]
    pub culture_language: Option<String>,
    #[pyo3(get)]
    pub culture_region: Option<String>,
    #[pyo3(get)]
    pub attribution: Option<String>,
}

impl From<&chart::ChartExGeography> for PyChartExGeography {
    fn from(g: &chart::ChartExGeography) -> Self {
        Self {
            projection_type: g.projection_type.clone(),
            viewed_region_type: g.viewed_region_type.clone(),
            culture_language: g.culture_language.clone(),
            culture_region: g.culture_region.clone(),
            attribution: g.attribution.clone(),
        }
    }
}

#[pyclass(name = "ChartExGridlines")]
#[derive(Clone)]
pub struct PyChartExGridlines {
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExGridlines> for PyChartExGridlines {
    fn from(g: &chart::ChartExGridlines) -> Self {
        Self {
            shape_properties: g.shape_properties.as_ref().map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExStatistics")]
#[derive(Clone)]
pub struct PyChartExStatistics {
    #[pyo3(get)]
    pub quartile_method: Option<String>,
}

impl From<&chart::ChartExStatistics> for PyChartExStatistics {
    fn from(s: &chart::ChartExStatistics) -> Self {
        Self {
            quartile_method: s.quartile_method.clone(),
        }
    }
}

#[pyclass(name = "ChartExDataPoint")]
#[derive(Clone)]
pub struct PyChartExDataPoint {
    #[pyo3(get)]
    pub idx: u32,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExDataPoint> for PyChartExDataPoint {
    fn from(p: &chart::ChartExDataPoint) -> Self {
        Self {
            idx: p.idx,
            shape_properties: p
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExDataLabel")]
#[derive(Clone)]
pub struct PyChartExDataLabel {
    #[pyo3(get)]
    pub idx: u32,
    #[pyo3(get)]
    pub position: Option<String>,
    #[pyo3(get)]
    pub visibility_series_name: Option<bool>,
    #[pyo3(get)]
    pub visibility_category_name: Option<bool>,
    #[pyo3(get)]
    pub visibility_value: Option<bool>,
    #[pyo3(get)]
    pub number_format: Option<PyChartNumberFormat>,
    #[pyo3(get)]
    pub separator: Option<String>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExDataLabel> for PyChartExDataLabel {
    fn from(l: &chart::ChartExDataLabel) -> Self {
        Self {
            idx: l.idx,
            position: l.position.clone(),
            visibility_series_name: l.visibility_series_name,
            visibility_category_name: l.visibility_category_name,
            visibility_value: l.visibility_value,
            number_format: l.number_format.as_ref().map(PyChartNumberFormat::from),
            separator: l.separator.clone(),
            shape_properties: l
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExFormatOverride")]
#[derive(Clone)]
pub struct PyChartExFormatOverride {
    #[pyo3(get)]
    pub idx: u32,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExFormatOverride> for PyChartExFormatOverride {
    fn from(o: &chart::ChartExFormatOverride) -> Self {
        Self {
            idx: o.idx,
            shape_properties: o
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExHeaderFooter")]
#[derive(Clone)]
pub struct PyChartExHeaderFooter {
    #[pyo3(get)]
    pub align_with_margins: Option<bool>,
    #[pyo3(get)]
    pub different_odd_even: Option<bool>,
    #[pyo3(get)]
    pub different_first: Option<bool>,
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
}

impl From<&chart::ChartExHeaderFooter> for PyChartExHeaderFooter {
    fn from(h: &chart::ChartExHeaderFooter) -> Self {
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

#[pyclass(name = "ChartExPageMargins")]
#[derive(Clone)]
pub struct PyChartExPageMargins {
    #[pyo3(get)]
    pub left: Option<f64>,
    #[pyo3(get)]
    pub right: Option<f64>,
    #[pyo3(get)]
    pub top: Option<f64>,
    #[pyo3(get)]
    pub bottom: Option<f64>,
    #[pyo3(get)]
    pub header: Option<f64>,
    #[pyo3(get)]
    pub footer: Option<f64>,
}

impl From<&chart::ChartExPageMargins> for PyChartExPageMargins {
    fn from(m: &chart::ChartExPageMargins) -> Self {
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

#[pyclass(name = "ChartExPageSetup")]
#[derive(Clone)]
pub struct PyChartExPageSetup {
    #[pyo3(get)]
    pub paper_size: Option<u32>,
    #[pyo3(get)]
    pub first_page_number: Option<u32>,
    #[pyo3(get)]
    pub orientation: Option<String>,
    #[pyo3(get)]
    pub black_and_white: Option<bool>,
    #[pyo3(get)]
    pub draft: Option<bool>,
    #[pyo3(get)]
    pub use_first_page_number: Option<bool>,
    #[pyo3(get)]
    pub horizontal_dpi: Option<u32>,
    #[pyo3(get)]
    pub vertical_dpi: Option<u32>,
    #[pyo3(get)]
    pub copies: Option<u32>,
}

impl From<&chart::ChartExPageSetup> for PyChartExPageSetup {
    fn from(p: &chart::ChartExPageSetup) -> Self {
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

#[pyclass(name = "ChartExPrintSettings")]
#[derive(Clone)]
pub struct PyChartExPrintSettings {
    #[pyo3(get)]
    pub header_footer: Option<PyChartExHeaderFooter>,
    #[pyo3(get)]
    pub page_margins: Option<PyChartExPageMargins>,
    #[pyo3(get)]
    pub page_setup: Option<PyChartExPageSetup>,
}

impl From<&chart::ChartExPrintSettings> for PyChartExPrintSettings {
    fn from(p: &chart::ChartExPrintSettings) -> Self {
        Self {
            header_footer: p.header_footer.as_ref().map(PyChartExHeaderFooter::from),
            page_margins: p.page_margins.as_ref().map(PyChartExPageMargins::from),
            page_setup: p.page_setup.as_ref().map(PyChartExPageSetup::from),
        }
    }
}

#[pyclass(name = "ChartExPlotArea")]
#[derive(Clone)]
pub struct PyChartExPlotArea {
    #[pyo3(get)]
    pub plot_surface: Option<PyChartShapeProperties>,
    #[pyo3(get)]
    pub series: Vec<PyChartExSeries>,
    #[pyo3(get)]
    pub axes: Vec<PyChartExAxis>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExPlotArea> for PyChartExPlotArea {
    fn from(p: &chart::ChartExPlotArea) -> Self {
        Self {
            plot_surface: p.plot_surface.as_ref().map(PyChartShapeProperties::from),
            series: p.series.iter().map(PyChartExSeries::from).collect(),
            axes: p.axes.iter().map(PyChartExAxis::from).collect(),
            shape_properties: p
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExDimension")]
#[derive(Clone)]
pub struct PyChartExDimension {
    #[pyo3(get)]
    pub dim_type: String,
    #[pyo3(get)]
    pub formula: Option<String>,
    #[pyo3(get)]
    pub nf_formula: Option<String>,
}

impl From<&chart::ChartExDimension> for PyChartExDimension {
    fn from(d: &chart::ChartExDimension) -> Self {
        match d {
            chart::ChartExDimension::String {
                dim_type,
                formula,
                nf_formula,
                ..
            } => Self {
                dim_type: match dim_type {
                    chart::StringDimType::Cat => "cat".into(),
                    chart::StringDimType::ColorStr => "colorStr".into(),
                    chart::StringDimType::EntityId => "entityId".into(),
                },
                formula: formula.clone(),
                nf_formula: nf_formula.clone(),
            },
            chart::ChartExDimension::Numeric {
                dim_type,
                formula,
                nf_formula,
                ..
            } => Self {
                dim_type: match dim_type {
                    chart::NumericDimType::Val => "val".into(),
                    chart::NumericDimType::X => "x".into(),
                    chart::NumericDimType::Y => "y".into(),
                    chart::NumericDimType::Size => "size".into(),
                    chart::NumericDimType::ColorVal => "colorVal".into(),
                },
                formula: formula.clone(),
                nf_formula: nf_formula.clone(),
            },
        }
    }
}

#[pyclass(name = "ChartExData")]
#[derive(Clone)]
pub struct PyChartExData {
    #[pyo3(get)]
    pub id: u32,
    #[pyo3(get)]
    pub dimensions: Vec<PyChartExDimension>,
}

impl From<&chart::ChartExData> for PyChartExData {
    fn from(d: &chart::ChartExData) -> Self {
        Self {
            id: d.id,
            dimensions: d.dimensions.iter().map(PyChartExDimension::from).collect(),
        }
    }
}

#[pyclass(name = "ChartExDataLabels")]
#[derive(Clone)]
pub struct PyChartExDataLabels {
    #[pyo3(get)]
    pub position: Option<String>,
    #[pyo3(get)]
    pub visibility_series_name: Option<bool>,
    #[pyo3(get)]
    pub visibility_category_name: Option<bool>,
    #[pyo3(get)]
    pub visibility_value: Option<bool>,
    #[pyo3(get)]
    pub number_format: Option<PyChartNumberFormat>,
    #[pyo3(get)]
    pub separator: Option<String>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
    #[pyo3(get)]
    pub overrides: Vec<PyChartExDataLabel>,
    #[pyo3(get)]
    pub hidden_labels: Vec<u32>,
}

impl From<&chart::ChartExDataLabels> for PyChartExDataLabels {
    fn from(l: &chart::ChartExDataLabels) -> Self {
        Self {
            position: l.position.clone(),
            visibility_series_name: l.visibility_series_name,
            visibility_category_name: l.visibility_category_name,
            visibility_value: l.visibility_value,
            number_format: l.number_format.as_ref().map(PyChartNumberFormat::from),
            separator: l.separator.clone(),
            shape_properties: l
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
            overrides: l.overrides.iter().map(PyChartExDataLabel::from).collect(),
            hidden_labels: l.hidden_labels.clone(),
        }
    }
}

#[pyclass(name = "ChartExTitle")]
#[derive(Clone)]
pub struct PyChartExTitle {
    #[pyo3(get)]
    pub text: Option<String>,
    #[pyo3(get)]
    pub position: Option<String>,
    #[pyo3(get)]
    pub align: Option<String>,
    #[pyo3(get)]
    pub overlay: Option<bool>,
    #[pyo3(get)]
    pub offset: Option<PyChartExOffset>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExTitle> for PyChartExTitle {
    fn from(t: &chart::ChartExTitle) -> Self {
        Self {
            text: t.text.clone(),
            position: t.position.clone(),
            align: t.align.clone(),
            overlay: t.overlay,
            offset: t.offset.as_ref().map(PyChartExOffset::from),
            shape_properties: t
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExLegend")]
#[derive(Clone)]
pub struct PyChartExLegend {
    #[pyo3(get)]
    pub position: Option<String>,
    #[pyo3(get)]
    pub align: Option<String>,
    #[pyo3(get)]
    pub overlay: Option<bool>,
    #[pyo3(get)]
    pub offset: Option<PyChartExOffset>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExLegend> for PyChartExLegend {
    fn from(l: &chart::ChartExLegend) -> Self {
        Self {
            position: l.position.clone(),
            align: l.align.clone(),
            overlay: l.overlay,
            offset: l.offset.as_ref().map(PyChartExOffset::from),
            shape_properties: l
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExLayoutPr")]
#[derive(Clone)]
pub struct PyChartExLayoutPr {
    #[pyo3(get)]
    pub parent_label_layout: Option<String>,
    #[pyo3(get)]
    pub region_label_layout: Option<String>,
    #[pyo3(get)]
    pub visibility: Option<PyChartExSeriesVisibility>,
    #[pyo3(get)]
    pub aggregation: bool,
    #[pyo3(get)]
    pub binning: Option<PyChartExBinning>,
    #[pyo3(get)]
    pub geography: Option<PyChartExGeography>,
    #[pyo3(get)]
    pub statistics: Option<PyChartExStatistics>,
    #[pyo3(get)]
    pub subtotals: Option<Vec<u32>>,
}

impl From<&chart::ChartExLayoutPr> for PyChartExLayoutPr {
    fn from(l: &chart::ChartExLayoutPr) -> Self {
        Self {
            parent_label_layout: l.parent_label_layout.clone(),
            region_label_layout: l.region_label_layout.clone(),
            visibility: l.visibility.as_ref().map(PyChartExSeriesVisibility::from),
            aggregation: l.aggregation,
            binning: l.binning.as_ref().map(PyChartExBinning::from),
            geography: l.geography.as_ref().map(PyChartExGeography::from),
            statistics: l.statistics.as_ref().map(PyChartExStatistics::from),
            subtotals: l.subtotals.clone(),
        }
    }
}

#[pyclass(name = "ChartExAxis")]
#[derive(Clone)]
pub struct PyChartExAxis {
    #[pyo3(get)]
    pub id: u32,
    #[pyo3(get)]
    pub hidden: Option<bool>,
    #[pyo3(get)]
    pub scaling: PyChartExScaling,
    #[pyo3(get)]
    pub title: Option<PyChartExAxisTitle>,
    #[pyo3(get)]
    pub units: Option<PyChartExAxisUnits>,
    #[pyo3(get)]
    pub major_gridlines: Option<PyChartExGridlines>,
    #[pyo3(get)]
    pub minor_gridlines: Option<PyChartExGridlines>,
    #[pyo3(get)]
    pub major_tick_marks: Option<String>,
    #[pyo3(get)]
    pub minor_tick_marks: Option<String>,
    #[pyo3(get)]
    pub tick_labels: bool,
    #[pyo3(get)]
    pub number_format: Option<PyChartNumberFormat>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExAxis> for PyChartExAxis {
    fn from(a: &chart::ChartExAxis) -> Self {
        Self {
            id: a.id,
            hidden: a.hidden,
            scaling: PyChartExScaling::from(&a.scaling),
            title: a.title.as_ref().map(PyChartExAxisTitle::from),
            units: a.units.as_ref().map(PyChartExAxisUnits::from),
            major_gridlines: a.major_gridlines.as_ref().map(PyChartExGridlines::from),
            minor_gridlines: a.minor_gridlines.as_ref().map(PyChartExGridlines::from),
            major_tick_marks: a.major_tick_marks.clone(),
            minor_tick_marks: a.minor_tick_marks.clone(),
            tick_labels: a.tick_labels,
            number_format: a.number_format.as_ref().map(PyChartNumberFormat::from),
            shape_properties: a
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartExSeries")]
#[derive(Clone)]
pub struct PyChartExSeries {
    #[pyo3(get)]
    pub layout: String,
    #[pyo3(get)]
    pub data_id: u32,
    #[pyo3(get)]
    pub unique_id: Option<String>,
    #[pyo3(get)]
    pub hidden: Option<bool>,
    #[pyo3(get)]
    pub owner_idx: Option<u32>,
    #[pyo3(get)]
    pub format_idx: Option<u32>,
    #[pyo3(get)]
    pub text: Option<PyChartExText>,
    #[pyo3(get)]
    pub data_labels: Option<PyChartExDataLabels>,
    #[pyo3(get)]
    pub data_points: Vec<PyChartExDataPoint>,
    #[pyo3(get)]
    pub layout_properties: Option<PyChartExLayoutPr>,
    #[pyo3(get)]
    pub axis_ids: Vec<u32>,
    #[pyo3(get)]
    pub value_colors: bool,
    #[pyo3(get)]
    pub value_color_positions: Option<PyChartExValueColorPositions>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
}

impl From<&chart::ChartExSeries> for PyChartExSeries {
    fn from(s: &chart::ChartExSeries) -> Self {
        Self {
            layout: chart_ex_layout_to_string(&s.layout).into(),
            data_id: s.data_id,
            unique_id: s.unique_id.clone(),
            hidden: s.hidden,
            owner_idx: s.owner_idx,
            format_idx: s.format_idx,
            text: s.text.as_ref().map(PyChartExText::from),
            data_labels: s.data_labels.as_ref().map(PyChartExDataLabels::from),
            data_points: s.data_points.iter().map(PyChartExDataPoint::from).collect(),
            layout_properties: s.layout_properties.as_ref().map(PyChartExLayoutPr::from),
            axis_ids: s.axis_ids.clone(),
            value_colors: s.value_colors.is_some(),
            value_color_positions: s
                .value_color_positions
                .as_ref()
                .map(PyChartExValueColorPositions::from),
            shape_properties: s
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
        }
    }
}

#[pyclass(name = "ChartStyleReference")]
#[derive(Clone)]
pub struct PyChartStyleReference {
    #[pyo3(get)]
    pub idx: u32,
    /// The colour override, as the XML it was read as.
    #[pyo3(get)]
    pub color: Option<String>,
}

impl From<&chart::StyleReference> for PyChartStyleReference {
    fn from(r: &chart::StyleReference) -> Self {
        Self {
            idx: r.idx,
            color: r.color.as_ref().map(|b| String::from_utf8_lossy(b).into_owned()),
        }
    }
}

#[pyclass(name = "ChartStyleEntry")]
#[derive(Clone)]
pub struct PyChartStyleEntry {
    #[pyo3(get)]
    pub line_reference: PyChartStyleReference,
    #[pyo3(get)]
    pub line_width_scale: Option<f64>,
    #[pyo3(get)]
    pub fill_reference: PyChartStyleReference,
    #[pyo3(get)]
    pub effect_reference: PyChartStyleReference,
    /// `major`, `minor` or `none`.
    #[pyo3(get)]
    pub font_collection: String,
    #[pyo3(get)]
    pub font_color: Option<String>,
    /// DrawingML kept as the XML it was read as.
    #[pyo3(get)]
    pub shape_properties: Option<String>,
    #[pyo3(get)]
    pub default_run_properties: Option<String>,
    #[pyo3(get)]
    pub body_properties: Option<String>,
    #[pyo3(get)]
    pub mods: Option<String>,
}

impl From<&chart::StyleEntry> for PyChartStyleEntry {
    fn from(e: &chart::StyleEntry) -> Self {
        let text =
            |b: &Option<Vec<u8>>| b.as_ref().map(|b| String::from_utf8_lossy(b).into_owned());
        Self {
            line_reference: (&e.line_reference).into(),
            line_width_scale: e.line_width_scale,
            fill_reference: (&e.fill_reference).into(),
            effect_reference: (&e.effect_reference).into(),
            font_collection: e.font_reference.collection.as_str().to_string(),
            font_color: text(&e.font_reference.color),
            shape_properties: text(&e.shape_properties),
            default_run_properties: text(&e.default_run_properties),
            body_properties: text(&e.body_properties),
            mods: e.mods.clone(),
        }
    }
}

/// A chart style part. `entries` is keyed by the element name the entry
/// belongs to; `raw` is set instead when the part could not be modelled
/// and is being replayed as read.
#[pyclass(name = "ChartStyle")]
#[derive(Clone)]
pub struct PyChartStyle {
    #[pyo3(get)]
    pub id: Option<u32>,
    #[pyo3(get)]
    pub entries: std::collections::BTreeMap<String, PyChartStyleEntry>,
    #[pyo3(get)]
    pub marker_symbol: Option<String>,
    #[pyo3(get)]
    pub marker_size: Option<u32>,
    #[pyo3(get)]
    pub raw: Option<String>,
}

impl From<&chart::ChartStylePart> for PyChartStyle {
    fn from(part: &chart::ChartStylePart) -> Self {
        match part {
            chart::ChartStylePart::Raw(bytes) => Self {
                id: None,
                entries: std::collections::BTreeMap::new(),
                marker_symbol: None,
                marker_size: None,
                raw: Some(String::from_utf8_lossy(bytes).into_owned()),
            },
            chart::ChartStylePart::Typed(style) => Self {
                id: Some(style.id),
                entries: chart::chart_style::entries_by_name(style)
                    .into_iter()
                    .map(|(name, entry)| (name.to_string(), entry.into()))
                    .collect(),
                marker_symbol: style
                    .data_point_marker_layout
                    .as_ref()
                    .and_then(|m| m.symbol.clone()),
                marker_size: style.data_point_marker_layout.as_ref().and_then(|m| m.size),
                raw: None,
            },
        }
    }
}

/// A chart colour style part. `raw` is set instead when the part could
/// not be modelled.
#[pyclass(name = "ChartColorStyle")]
#[derive(Clone)]
pub struct PyChartColorStyle {
    #[pyo3(get)]
    pub method: Option<String>,
    #[pyo3(get)]
    pub id: Option<u32>,
    #[pyo3(get)]
    pub colors: Vec<String>,
    #[pyo3(get)]
    pub variations: Vec<String>,
    #[pyo3(get)]
    pub raw: Option<String>,
}

impl From<&chart::ChartColorStylePart> for PyChartColorStyle {
    fn from(part: &chart::ChartColorStylePart) -> Self {
        match part {
            chart::ChartColorStylePart::Raw(bytes) => Self {
                method: None,
                id: None,
                colors: Vec::new(),
                variations: Vec::new(),
                raw: Some(String::from_utf8_lossy(bytes).into_owned()),
            },
            chart::ChartColorStylePart::Typed(style) => Self {
                method: Some(style.method.as_str().to_string()),
                id: style.id,
                colors: style
                    .colors
                    .iter()
                    .map(|b| String::from_utf8_lossy(b).into_owned())
                    .collect(),
                variations: style
                    .variations
                    .iter()
                    .map(|b| String::from_utf8_lossy(b).into_owned())
                    .collect(),
                raw: None,
            },
        }
    }
}

#[pyclass(name = "ChartEx")]
#[derive(Clone)]
pub struct PyChartEx {
    pub(crate) inner: chart::ChartEx,
    #[pyo3(get)]
    pub layout: String,
    #[pyo3(get)]
    pub version: Option<String>,
    #[pyo3(get)]
    pub feature_list: Option<String>,
    #[pyo3(get)]
    pub fallback_img: Option<String>,
    #[pyo3(get)]
    pub title: Option<PyChartExTitle>,
    #[pyo3(get)]
    pub data: Vec<PyChartExData>,
    #[pyo3(get)]
    pub plot_area: PyChartExPlotArea,
    #[pyo3(get)]
    pub legend: Option<PyChartExLegend>,
    #[pyo3(get)]
    pub shape_properties: Option<PyChartShapeProperties>,
    #[pyo3(get)]
    pub format_overrides: Vec<PyChartExFormatOverride>,
    #[pyo3(get)]
    pub print_settings: Option<PyChartExPrintSettings>,
    pub external_data_rel_id: Option<String>,
    pub external_data_auto_update: Option<bool>,
    #[pyo3(get)]
    pub style: Option<PyChartStyle>,
    #[pyo3(get)]
    pub color_style: Option<PyChartColorStyle>,
}

impl From<&chart::ChartEx> for PyChartEx {
    fn from(c: &chart::ChartEx) -> Self {
        let layout = c
            .plot_area
            .series
            .first()
            .map(|s| chart_ex_layout_to_string(&s.layout))
            .unwrap_or("unknown");
        Self {
            inner: c.clone(),
            layout: layout.into(),
            version: c.version.clone(),
            feature_list: c.feature_list.clone(),
            fallback_img: c.fallback_img.clone(),
            title: c.title.as_ref().map(PyChartExTitle::from),
            data: c.data.iter().map(PyChartExData::from).collect(),
            plot_area: PyChartExPlotArea::from(&c.plot_area),
            legend: c.legend.as_ref().map(PyChartExLegend::from),
            shape_properties: c
                .shape_properties
                .as_ref()
                .map(PyChartShapeProperties::from),
            format_overrides: c
                .format_overrides
                .iter()
                .map(PyChartExFormatOverride::from)
                .collect(),
            print_settings: c.print_settings.as_ref().map(PyChartExPrintSettings::from),
            external_data_rel_id: c.external_data.as_ref().map(|e| e.rel_id.clone()),
            external_data_auto_update: c.external_data.as_ref().and_then(|e| e.auto_update),
            style: c.style.as_ref().map(PyChartStyle::from),
            color_style: c.color_style.as_ref().map(PyChartColorStyle::from),
        }
    }
}

fn chart_ex_layout_from_python(value: &str) -> PyResult<chart::ChartExLayout> {
    Ok(match value {
        "waterfall" => chart::ChartExLayout::Waterfall,
        "treemap" => chart::ChartExLayout::Treemap,
        "sunburst" => chart::ChartExLayout::Sunburst,
        "funnel" => chart::ChartExLayout::Funnel,
        "histogram" => chart::ChartExLayout::Histogram,
        "box_whisker" | "boxWhisker" => chart::ChartExLayout::BoxWhisker,
        "pareto_line" | "paretoLine" => chart::ChartExLayout::ParetoLine,
        "region_map" | "regionMap" => chart::ChartExLayout::RegionMap,
        "clustered_column" | "clusteredColumn" => chart::ChartExLayout::ClusteredColumn,
        other => {
            return Err(PyValueError::new_err(format!(
                "unsupported ChartEx layout {other:?}"
            )))
        }
    })
}

#[pymethods]
impl PyChartEx {
    #[new]
    #[pyo3(signature=(layout, *, title=None, version=None))]
    fn new(layout: &str, title: Option<String>, version: Option<String>) -> PyResult<Self> {
        let series = chart::ChartExSeries {
            layout: chart_ex_layout_from_python(layout)?,
            unique_id: None,
            hidden: None,
            owner_idx: None,
            format_idx: None,
            text: None,
            data_id: 0,
            data_labels: None,
            data_points: Vec::new(),
            layout_properties: None,
            axis_ids: Vec::new(),
            value_colors: None,
            value_color_positions: None,
            shape_properties: None,
            extensions: None,
        };
        let chart = chart::ChartEx {
            version,
            feature_list: None,
            fallback_img: None,
            title: title.map(|text| chart::ChartExTitle {
                text: Some(text),
                ..Default::default()
            }),
            data: Vec::new(),
            external_data: None,
            plot_area: chart::ChartExPlotArea {
                series: vec![series],
                ..Default::default()
            },
            legend: None,
            shape_properties: None,
            text_properties: None,
            color_map_override: None,
            format_overrides: Vec::new(),
            print_settings: None,
            style: None,
        preserved_rels: Vec::new(),
            color_style: None,
            extensions: None,
            raw_extensions: std::collections::HashMap::new(),
            raw_mc_fallback: None,
        };
        Ok(Self::from(&chart))
    }
}
