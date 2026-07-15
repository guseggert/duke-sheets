use duke_sheets_chart as chart;
use duke_sheets_core as core;
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyBytes};

use crate::types::{
    color_input_to_core, horizontal_alignment_to_string, parse_horizontal_alignment_input,
    parse_vertical_alignment_input, vertical_alignment_to_string,
};
use crate::{
    to_py_err, PyChart, PyChartEx, PyColor, PyComment, PyDrawingAnchor, PyFormControl,
    PyRichTextRun, PyWorkbook, PyWorksheet,
};

#[pyclass(name = "DrawingText")]
#[derive(Clone)]
pub struct PyDrawingText {
    pub(crate) inner: core::DrawingText,
}

impl From<&core::DrawingText> for PyDrawingText {
    fn from(text: &core::DrawingText) -> Self {
        Self {
            inner: text.clone(),
        }
    }
}

pub(crate) fn drawing_text_from_input(value: &Bound<'_, PyAny>) -> PyResult<core::DrawingText> {
    if let Ok(text) = value.extract::<String>() {
        return Ok(core::DrawingText::plain(text));
    }
    value
        .extract::<PyRef<'_, PyDrawingText>>()
        .map(|text| text.inner.clone())
        .map_err(|_| PyTypeError::new_err("expected str or DrawingText"))
}

#[pymethods]
impl PyDrawingText {
    #[new]
    #[pyo3(signature=(runs, *, horizontal_alignment=None, vertical_alignment=None))]
    fn new(
        runs: &Bound<'_, PyAny>,
        horizontal_alignment: Option<&str>,
        vertical_alignment: Option<&str>,
    ) -> PyResult<Self> {
        let runs = if let Ok(text) = runs.extract::<String>() {
            vec![core::RichTextRun::plain(text)]
        } else {
            let mut parsed = Vec::new();
            for item in runs.iter().map_err(|_| {
                PyTypeError::new_err("runs must be a string or iterable of RichTextRun")
            })? {
                let item = item?;
                let run = item.extract::<PyRef<'_, PyRichTextRun>>().map_err(|_| {
                    PyTypeError::new_err("runs must contain only RichTextRun objects")
                })?;
                parsed.push(run.to_core()?);
            }
            parsed
        };
        Ok(Self {
            inner: core::DrawingText {
                runs,
                horizontal_alignment: horizontal_alignment
                    .map(parse_horizontal_alignment_input)
                    .transpose()?,
                vertical_alignment: vertical_alignment
                    .map(parse_vertical_alignment_input)
                    .transpose()?,
            },
        })
    }

    #[getter]
    fn runs(&self) -> Vec<PyRichTextRun> {
        self.inner.runs.iter().map(PyRichTextRun::from).collect()
    }

    #[getter]
    fn horizontal_alignment(&self) -> Option<&'static str> {
        self.inner
            .horizontal_alignment
            .as_ref()
            .map(horizontal_alignment_to_string)
    }

    #[getter]
    fn vertical_alignment(&self) -> Option<&'static str> {
        self.inner
            .vertical_alignment
            .as_ref()
            .map(vertical_alignment_to_string)
    }

    #[getter]
    fn plain_text(&self) -> String {
        self.inner.plain_text()
    }

    fn __str__(&self) -> String {
        self.inner.plain_text()
    }
}

#[pyclass(name = "DrawingMeta")]
#[derive(Clone)]
pub struct PyDrawingMeta {
    pub(crate) name: Option<String>,
    pub(crate) hidden: Option<bool>,
    pub(crate) locked: bool,
    pub(crate) printable: bool,
    pub(crate) alt_text: Option<String>,
    pub(crate) title: Option<String>,
}

impl From<&core::DrawingMeta> for PyDrawingMeta {
    fn from(meta: &core::DrawingMeta) -> Self {
        Self {
            name: meta.name.clone(),
            hidden: Some(meta.hidden),
            locked: meta.locked,
            printable: meta.printable,
            alt_text: meta.alt_text.clone(),
            title: meta.title.clone(),
        }
    }
}

impl Default for PyDrawingMeta {
    fn default() -> Self {
        Self {
            name: None,
            hidden: None,
            locked: true,
            printable: true,
            alt_text: None,
            title: None,
        }
    }
}

impl PyDrawingMeta {
    /// Resolve to the core meta, applying the kind-specific hidden
    /// default (comments hide by default) when unset.
    fn to_core(&self, hidden_default: bool) -> core::DrawingMeta {
        core::DrawingMeta {
            name: self.name.clone(),
            hidden: self.hidden.unwrap_or(hidden_default),
            locked: self.locked,
            printable: self.printable,
            alt_text: self.alt_text.clone(),
            title: self.title.clone(),
        }
    }
}

#[pymethods]
impl PyDrawingMeta {
    #[new]
    #[pyo3(signature=(*, name=None, hidden=None, locked=true, printable=true, alt_text=None, title=None))]
    fn new(
        name: Option<String>,
        hidden: Option<bool>,
        locked: bool,
        printable: bool,
        alt_text: Option<String>,
        title: Option<String>,
    ) -> Self {
        Self {
            name,
            hidden,
            locked,
            printable,
            alt_text,
            title,
        }
    }

    #[getter]
    fn name(&self) -> Option<String> {
        self.name.clone()
    }

    /// None means the default for the drawing kind is applied when the
    /// meta is attached (True for comments, False otherwise).
    #[getter]
    fn hidden(&self) -> Option<bool> {
        self.hidden
    }

    #[getter]
    fn locked(&self) -> bool {
        self.locked
    }

    #[getter]
    fn printable(&self) -> bool {
        self.printable
    }

    #[getter]
    fn alt_text(&self) -> Option<String> {
        self.alt_text.clone()
    }

    #[getter]
    fn title(&self) -> Option<String> {
        self.title.clone()
    }
}

#[pyclass(name = "ChildTransform")]
#[derive(Clone)]
pub struct PyChildTransform {
    pub(crate) inner: chart::ChildTransform,
}

impl From<&chart::ChildTransform> for PyChildTransform {
    fn from(transform: &chart::ChildTransform) -> Self {
        Self {
            inner: transform.clone(),
        }
    }
}

#[pymethods]
impl PyChildTransform {
    #[new]
    #[pyo3(signature=(*, x_emu=0, y_emu=0, cx_emu=0, cy_emu=0, rotation=0, flip_h=false, flip_v=false))]
    fn new(
        x_emu: i64,
        y_emu: i64,
        cx_emu: i64,
        cy_emu: i64,
        rotation: i32,
        flip_h: bool,
        flip_v: bool,
    ) -> PyResult<Self> {
        if cx_emu < 0 || cy_emu < 0 {
            return Err(PyValueError::new_err(
                "child transform extents cannot be negative",
            ));
        }
        Ok(Self {
            inner: chart::ChildTransform {
                x_emu,
                y_emu,
                cx_emu,
                cy_emu,
                rotation,
                flip_h,
                flip_v,
            },
        })
    }

    #[getter]
    fn x_emu(&self) -> i64 {
        self.inner.x_emu
    }

    #[getter]
    fn y_emu(&self) -> i64 {
        self.inner.y_emu
    }

    #[getter]
    fn cx_emu(&self) -> i64 {
        self.inner.cx_emu
    }

    #[getter]
    fn cy_emu(&self) -> i64 {
        self.inner.cy_emu
    }

    #[getter]
    fn rotation(&self) -> i32 {
        self.inner.rotation
    }

    #[getter]
    fn flip_h(&self) -> bool {
        self.inner.flip_h
    }

    #[getter]
    fn flip_v(&self) -> bool {
        self.inner.flip_v
    }
}

#[pyclass(name = "GroupTransform")]
#[derive(Clone)]
pub struct PyGroupTransform {
    pub(crate) inner: chart::GroupTransform,
}

impl From<&chart::GroupTransform> for PyGroupTransform {
    fn from(transform: &chart::GroupTransform) -> Self {
        Self {
            inner: transform.clone(),
        }
    }
}

#[pymethods]
impl PyGroupTransform {
    #[new]
    #[pyo3(signature=(*, x_emu=0, y_emu=0, cx_emu=0, cy_emu=0, child_x_emu=0, child_y_emu=0, child_cx_emu=0, child_cy_emu=0, rotation=0, flip_h=false, flip_v=false))]
    fn new(
        x_emu: i64,
        y_emu: i64,
        cx_emu: i64,
        cy_emu: i64,
        child_x_emu: i64,
        child_y_emu: i64,
        child_cx_emu: i64,
        child_cy_emu: i64,
        rotation: i32,
        flip_h: bool,
        flip_v: bool,
    ) -> Self {
        Self {
            inner: chart::GroupTransform {
                x_emu,
                y_emu,
                cx_emu,
                cy_emu,
                child_x_emu,
                child_y_emu,
                child_cx_emu,
                child_cy_emu,
                rotation,
                flip_h,
                flip_v,
            },
        }
    }

    #[getter]
    fn x_emu(&self) -> i64 {
        self.inner.x_emu
    }

    #[getter]
    fn y_emu(&self) -> i64 {
        self.inner.y_emu
    }

    #[getter]
    fn cx_emu(&self) -> i64 {
        self.inner.cx_emu
    }

    #[getter]
    fn cy_emu(&self) -> i64 {
        self.inner.cy_emu
    }

    #[getter]
    fn child_x_emu(&self) -> i64 {
        self.inner.child_x_emu
    }

    #[getter]
    fn child_y_emu(&self) -> i64 {
        self.inner.child_y_emu
    }

    #[getter]
    fn child_cx_emu(&self) -> i64 {
        self.inner.child_cx_emu
    }

    #[getter]
    fn child_cy_emu(&self) -> i64 {
        self.inner.child_cy_emu
    }

    #[getter]
    fn rotation(&self) -> i32 {
        self.inner.rotation
    }

    #[getter]
    fn flip_h(&self) -> bool {
        self.inner.flip_h
    }

    #[getter]
    fn flip_v(&self) -> bool {
        self.inner.flip_v
    }
}

#[pyclass(name = "ShapeFill")]
#[derive(Clone)]
pub struct PyShapeFill {
    pub(crate) inner: core::ShapeFill,
}

impl From<&core::ShapeFill> for PyShapeFill {
    fn from(fill: &core::ShapeFill) -> Self {
        Self {
            inner: fill.clone(),
        }
    }
}

#[pymethods]
impl PyShapeFill {
    #[staticmethod]
    fn none() -> Self {
        Self {
            inner: core::ShapeFill::None,
        }
    }

    #[staticmethod]
    fn solid(color: &Bound<'_, PyAny>) -> PyResult<Self> {
        Ok(Self {
            inner: core::ShapeFill::Solid(color_input_to_core(color)?),
        })
    }

    #[getter]
    fn kind(&self) -> &'static str {
        match self.inner {
            core::ShapeFill::None => "none",
            core::ShapeFill::Solid(_) => "solid",
        }
    }

    #[getter]
    fn color(&self) -> Option<PyColor> {
        match &self.inner {
            core::ShapeFill::None => None,
            core::ShapeFill::Solid(color) => Some(PyColor::from(color)),
        }
    }
}

#[pyclass(name = "ShapeLine")]
#[derive(Clone)]
pub struct PyShapeLine {
    pub(crate) inner: core::ShapeLine,
}

impl From<&core::ShapeLine> for PyShapeLine {
    fn from(line: &core::ShapeLine) -> Self {
        Self {
            inner: line.clone(),
        }
    }
}

#[pymethods]
impl PyShapeLine {
    #[new]
    #[pyo3(signature=(*, color=None, width_emu=None, dash_style=None, no_fill=false))]
    fn new(
        color: Option<&Bound<'_, PyAny>>,
        width_emu: Option<i64>,
        dash_style: Option<String>,
        no_fill: bool,
    ) -> PyResult<Self> {
        if width_emu.is_some_and(|width| width < 0) {
            return Err(PyValueError::new_err(
                "shape line width cannot be negative",
            ));
        }
        Ok(Self {
            inner: core::ShapeLine {
                color: color.map(color_input_to_core).transpose()?,
                width_emu,
                dash_style,
                no_fill,
            },
        })
    }

    #[getter]
    fn color(&self) -> Option<PyColor> {
        self.inner.color.as_ref().map(PyColor::from)
    }

    #[getter]
    fn width_emu(&self) -> Option<i64> {
        self.inner.width_emu
    }

    #[getter]
    fn dash_style(&self) -> Option<String> {
        self.inner.dash_style.clone()
    }

    #[getter]
    fn no_fill(&self) -> bool {
        self.inner.no_fill
    }
}

#[pyclass(name = "Shape")]
#[derive(Clone)]
pub struct PyShape {
    pub(crate) inner: core::Shape,
}

impl From<&core::Shape> for PyShape {
    fn from(shape: &core::Shape) -> Self {
        Self {
            inner: shape.clone(),
        }
    }
}

#[pymethods]
impl PyShape {
    #[new]
    #[pyo3(signature=(geometry="rect", *, fill=None, line=None, text=None, rotation=0, flip_h=false, flip_v=false))]
    fn new(
        geometry: &str,
        fill: Option<PyRef<'_, PyShapeFill>>,
        line: Option<PyRef<'_, PyShapeLine>>,
        text: Option<&Bound<'_, PyAny>>,
        rotation: i32,
        flip_h: bool,
        flip_v: bool,
    ) -> PyResult<Self> {
        if geometry.trim().is_empty() {
            return Err(PyValueError::new_err("shape geometry cannot be empty"));
        }
        Ok(Self {
            inner: core::Shape {
                geometry: core::ShapeGeometry::Preset(geometry.to_string()),
                fill: fill
                    .map(|fill| fill.inner.clone())
                    .unwrap_or(core::ShapeFill::None),
                line: line
                    .map(|line| line.inner.clone())
                    .unwrap_or_default(),
                text: text.map(drawing_text_from_input).transpose()?,
                rotation,
                flip_h,
                flip_v,
                ..core::Shape::default()
            },
        })
    }

    #[getter]
    fn geometry(&self) -> String {
        match &self.inner.geometry {
            core::ShapeGeometry::Preset(name) => name.clone(),
        }
    }

    #[getter]
    fn fill(&self) -> PyShapeFill {
        PyShapeFill::from(&self.inner.fill)
    }

    #[getter]
    fn line(&self) -> PyShapeLine {
        PyShapeLine::from(&self.inner.line)
    }

    #[getter]
    fn text(&self) -> Option<PyDrawingText> {
        self.inner.text.as_ref().map(PyDrawingText::from)
    }

    #[getter]
    fn rotation(&self) -> i32 {
        self.inner.rotation
    }

    #[getter]
    fn flip_h(&self) -> bool {
        self.inner.flip_h
    }

    #[getter]
    fn flip_v(&self) -> bool {
        self.inner.flip_v
    }
}

#[pyclass(name = "EmbeddedImage")]
#[derive(Clone)]
pub struct PyEmbeddedImage {
    pub(crate) inner: chart::EmbeddedImage,
    input_has_data: bool,
}

impl PyEmbeddedImage {
    fn from_metadata(image: &chart::EmbeddedImage) -> Self {
        Self {
            inner: chart::EmbeddedImage {
                format: image.format,
                media_path: image.media_path.clone(),
                svg_media_path: image.svg_media_path.clone(),
                width_emu: image.width_emu,
                height_emu: image.height_emu,
                rotation: image.rotation,
                flip_h: image.flip_h,
                flip_v: image.flip_v,
                data: Vec::new(),
                svg_data: None,
            },
            input_has_data: false,
        }
    }

    fn to_core(&self) -> PyResult<chart::EmbeddedImage> {
        if !self.input_has_data {
            return Err(PyValueError::new_err(
                "image input requires bytes; fetch them with drawing_image_data(path)",
            ));
        }
        Ok(self.inner.clone())
    }
}

#[pymethods]
impl PyEmbeddedImage {
    #[new]
    #[pyo3(signature=(data, format, *, media_path=None, svg_data=None, svg_media_path=None, width_emu=0, height_emu=0, rotation=None, flip_h=false, flip_v=false))]
    fn new(
        data: &Bound<'_, PyBytes>,
        format: &str,
        media_path: Option<String>,
        svg_data: Option<&Bound<'_, PyBytes>>,
        svg_media_path: Option<String>,
        width_emu: i64,
        height_emu: i64,
        rotation: Option<i32>,
        flip_h: bool,
        flip_v: bool,
    ) -> PyResult<Self> {
        let image_format = chart::ImageFormat::from_extension(format).ok_or_else(|| {
            PyValueError::new_err(
                "format must be png, jpeg, gif, bmp, emf, wmf, tiff, or svg",
            )
        })?;
        let extension = image_format.as_str();
        Ok(Self {
            inner: chart::EmbeddedImage {
                format: image_format,
                media_path: media_path.unwrap_or_else(|| format!("image.{extension}")),
                svg_media_path,
                width_emu,
                height_emu,
                rotation,
                flip_h,
                flip_v,
                data: data.as_bytes().to_vec(),
                svg_data: svg_data.map(|bytes| bytes.as_bytes().to_vec()),
            },
            input_has_data: true,
        })
    }

    #[getter]
    fn format(&self) -> &'static str {
        self.inner.format.as_str()
    }

    #[getter]
    fn media_path(&self) -> String {
        self.inner.media_path.clone()
    }

    #[getter]
    fn svg_media_path(&self) -> Option<String> {
        self.inner.svg_media_path.clone()
    }

    #[getter]
    fn width_emu(&self) -> i64 {
        self.inner.width_emu
    }

    #[getter]
    fn height_emu(&self) -> i64 {
        self.inner.height_emu
    }

    #[getter]
    fn rotation(&self) -> Option<i32> {
        self.inner.rotation
    }

    #[getter]
    fn flip_h(&self) -> bool {
        self.inner.flip_h
    }

    #[getter]
    fn flip_v(&self) -> bool {
        self.inner.flip_v
    }

    #[getter]
    fn has_svg(&self) -> bool {
        self.inner.svg_media_path.is_some() || self.inner.svg_data.is_some()
    }
}

#[pyclass(name = "DrawingComment")]
#[derive(Clone)]
pub struct PyDrawingComment {
    row: u32,
    col: u16,
    comment: core::CellComment,
}

#[pymethods]
impl PyDrawingComment {
    #[new]
    #[pyo3(signature=(row, col, text, *, author=None))]
    fn new(row: u32, col: u16, text: String, author: Option<String>) -> Self {
        Self {
            row,
            col,
            comment: core::CellComment::new(author.unwrap_or_default(), text),
        }
    }

    #[getter]
    fn row(&self) -> u32 {
        self.row
    }

    #[getter]
    fn col(&self) -> u16 {
        self.col
    }

    #[getter]
    fn author(&self) -> String {
        self.comment.author.clone()
    }

    #[getter]
    fn text(&self) -> String {
        self.comment.text.clone()
    }

    #[getter]
    fn comment(&self) -> PyComment {
        PyComment::from(&self.comment)
    }
}

#[pyclass(name = "RawDrawingRelationship")]
#[derive(Clone)]
pub struct PyRawDrawingRelationship {
    #[pyo3(get)]
    id: String,
    #[pyo3(get)]
    rel_type: String,
    #[pyo3(get)]
    target: String,
    #[pyo3(get)]
    external: bool,
    #[pyo3(get)]
    has_part: bool,
}

#[pyclass(name = "RawDrawing")]
#[derive(Clone)]
pub struct PyRawDrawing {
    #[pyo3(get)]
    byte_length: usize,
    #[pyo3(get)]
    relationships: Vec<PyRawDrawingRelationship>,
}

impl From<&core::RawDrawing> for PyRawDrawing {
    fn from(raw: &core::RawDrawing) -> Self {
        Self {
            byte_length: raw.bytes.len(),
            relationships: raw
                .rels
                .iter()
                .map(|rel| PyRawDrawingRelationship {
                    id: rel.id.clone(),
                    rel_type: rel.rel_type.clone(),
                    target: rel.target.clone(),
                    external: rel.external,
                    has_part: rel.part.is_some(),
                })
                .collect(),
        }
    }
}

#[pyclass(name = "DrawingGroup")]
#[derive(Clone)]
pub struct PyDrawingGroup {
    transform: PyGroupTransform,
    children: Vec<PyDrawing>,
}

#[pymethods]
impl PyDrawingGroup {
    #[new]
    fn new(
        transform: PyRef<'_, PyGroupTransform>,
        children: &Bound<'_, PyAny>,
    ) -> PyResult<Self> {
        let mut parsed = Vec::new();
        for item in children
            .iter()
            .map_err(|_| PyTypeError::new_err("children must be an iterable of Drawing"))?
        {
            let item = item?;
            let drawing = item
                .extract::<PyRef<'_, PyDrawing>>()
                .map_err(|_| PyTypeError::new_err("children must contain only Drawing objects"))?;
            drawing.to_group_child()?;
            parsed.push(drawing.clone());
        }
        Ok(Self {
            transform: transform.clone(),
            children: parsed,
        })
    }

    #[getter]
    fn transform(&self) -> PyGroupTransform {
        self.transform.clone()
    }

    #[getter]
    fn children(&self) -> Vec<PyDrawing> {
        self.children.clone()
    }
}

#[pyclass(name = "RectEmu")]
#[derive(Clone, Copy)]
pub struct PyRectEmu {
    #[pyo3(get)]
    pub x_emu: i64,
    #[pyo3(get)]
    pub y_emu: i64,
    #[pyo3(get)]
    pub width_emu: i64,
    #[pyo3(get)]
    pub height_emu: i64,
}

#[pymethods]
impl PyRectEmu {
    fn __repr__(&self) -> String {
        format!(
            "RectEmu(x_emu={}, y_emu={}, width_emu={}, height_emu={})",
            self.x_emu, self.y_emu, self.width_emu, self.height_emu
        )
    }
}

impl PyRectEmu {
    fn from_core(rect: core::drawing::RectEmu) -> Self {
        Self {
            x_emu: rect.x_emu,
            y_emu: rect.y_emu,
            width_emu: rect.width_emu,
            height_emu: rect.height_emu,
        }
    }
}

#[pyclass(name = "Drawing")]
#[derive(Clone)]
pub struct PyDrawing {
    drawing_path: Vec<usize>,
    absolute_rect_emu: Option<PyRectEmu>,
    meta: PyDrawingMeta,
    anchor: Option<PyDrawingAnchor>,
    transform: Option<PyChildTransform>,
    image: Option<PyEmbeddedImage>,
    chart: Option<PyChart>,
    chart_ex: Option<PyChartEx>,
    form_control: Option<PyFormControl>,
    comment: Option<PyDrawingComment>,
    shape: Option<PyShape>,
    group: Option<PyDrawingGroup>,
    raw: Option<PyRawDrawing>,
}

impl PyDrawing {
    fn empty(
        meta: PyDrawingMeta,
        anchor: Option<PyDrawingAnchor>,
        transform: Option<PyChildTransform>,
    ) -> Self {
        Self {
            drawing_path: Vec::new(),
            absolute_rect_emu: None,
            meta,
            anchor,
            transform,
            image: None,
            chart: None,
            chart_ex: None,
            form_control: None,
            comment: None,
            shape: None,
            group: None,
            raw: None,
        }
    }

    fn from_kind(
        sheet: &core::Worksheet,
        meta: &core::DrawingMeta,
        kind: &core::DrawingKind,
        path: Vec<usize>,
        anchor: Option<&chart::DrawingAnchor>,
        transform: Option<&chart::ChildTransform>,
    ) -> Self {
        let mut drawing = Self::empty(
            PyDrawingMeta::from(meta),
            anchor.map(PyDrawingAnchor::from),
            transform.map(PyChildTransform::from),
        );
        drawing.drawing_path = path.clone();
        drawing.absolute_rect_emu = sheet.drawing_rect_emu(&path).map(PyRectEmu::from_core);
        match kind {
            core::DrawingKind::Image(image) => {
                drawing.image = Some(PyEmbeddedImage::from_metadata(image));
            }
            core::DrawingKind::Chart(chart) => {
                drawing.chart = Some(PyChart::from(chart.as_ref()));
            }
            core::DrawingKind::ChartEx(chart) => {
                drawing.chart_ex = Some(PyChartEx::from(chart.as_ref()));
            }
            core::DrawingKind::FormControl(control) => {
                drawing.form_control = Some(PyFormControl::from(control));
            }
            core::DrawingKind::Comment { row, col, comment } => {
                drawing.comment = Some(PyDrawingComment {
                    row: *row,
                    col: *col,
                    comment: comment.clone(),
                });
            }
            core::DrawingKind::Shape(shape) => {
                drawing.shape = Some(PyShape::from(shape.as_ref()));
            }
            core::DrawingKind::Group(group) => {
                let children = group
                    .children
                    .iter()
                    .enumerate()
                    .map(|(index, child)| {
                        let mut child_path = path.clone();
                        child_path.push(index);
                        Self::from_kind(
                            sheet,
                            &child.meta,
                            &child.kind,
                            child_path,
                            None,
                            Some(&child.transform),
                        )
                    })
                    .collect();
                drawing.group = Some(PyDrawingGroup {
                    transform: PyGroupTransform::from(&group.transform),
                    children,
                });
            }
            core::DrawingKind::Raw(raw) => {
                drawing.raw = Some(PyRawDrawing::from(raw));
            }
        }
        drawing
    }

    pub(crate) fn from_top(
        sheet: &core::Worksheet,
        object: &core::DrawingObject,
        index: usize,
    ) -> Self {
        Self::from_kind(
            sheet,
            &object.meta,
            &object.kind,
            vec![index],
            Some(&object.anchor),
            None,
        )
    }

    fn kind_to_core(&self) -> PyResult<core::DrawingKind> {
        if let Some(image) = &self.image {
            return Ok(core::DrawingKind::Image(image.to_core()?));
        }
        if let Some(chart) = &self.chart {
            return Ok(core::DrawingKind::Chart(Box::new(chart.inner.clone())));
        }
        if let Some(chart) = &self.chart_ex {
            return Ok(core::DrawingKind::ChartEx(Box::new(chart.inner.clone())));
        }
        if let Some(control) = &self.form_control {
            return Ok(core::DrawingKind::FormControl(control.inner.clone()));
        }
        if let Some(comment) = &self.comment {
            return Ok(core::DrawingKind::Comment {
                row: comment.row,
                col: comment.col,
                comment: comment.comment.clone(),
            });
        }
        if let Some(shape) = &self.shape {
            return Ok(core::DrawingKind::Shape(Box::new(shape.inner.clone())));
        }
        if let Some(group) = &self.group {
            let children = group
                .children
                .iter()
                .map(PyDrawing::to_group_child)
                .collect::<PyResult<Vec<_>>>()?;
            return Ok(core::DrawingKind::Group(Box::new(core::Group {
                transform: group.transform.inner.clone(),
                children,
            })));
        }
        if self.raw.is_some() {
            return Err(PyValueError::new_err(
                "raw drawings are read-only and cannot be used as mutation input",
            ));
        }
        Err(PyValueError::new_err("drawing has no payload"))
    }

    pub(crate) fn to_top_level(&self) -> PyResult<core::DrawingObject> {
        let anchor = self.anchor.as_ref().ok_or_else(|| {
            PyValueError::new_err("top-level drawing input requires an anchor")
        })?;
        if self.transform.is_some() {
            return Err(PyValueError::new_err(
                "top-level drawing input cannot have a child transform",
            ));
        }
        let object = core::DrawingObject {
            meta: self.meta.to_core(self.comment.is_some()),
            anchor: anchor.to_core()?,
            kind: self.kind_to_core()?,
        };
        object
            .validate()
            .map_err(|error| PyValueError::new_err(error.to_string()))?;
        Ok(object)
    }

    fn to_group_child(&self) -> PyResult<core::GroupChild> {
        let transform = self.transform.as_ref().ok_or_else(|| {
            PyValueError::new_err("group child drawing input requires a child transform")
        })?;
        if self.anchor.is_some() {
            return Err(PyValueError::new_err(
                "group child drawing input cannot have a sheet anchor",
            ));
        }
        let child = core::GroupChild {
            meta: self.meta.to_core(self.comment.is_some()),
            transform: transform.inner.clone(),
            kind: self.kind_to_core()?,
        };
        let validator = core::DrawingObject::group(core::Group {
            transform: chart::GroupTransform::default(),
            children: vec![child.clone()],
        });
        validator
            .validate()
            .map_err(|error| PyValueError::new_err(error.to_string()))?;
        Ok(child)
    }

    fn kind_name(&self) -> &'static str {
        if self.image.is_some() {
            "image"
        } else if self.chart.is_some() {
            "chart"
        } else if self.chart_ex.is_some() {
            "chart_ex"
        } else if self.form_control.is_some() {
            "form_control"
        } else if self.comment.is_some() {
            "comment"
        } else if self.shape.is_some() {
            "shape"
        } else if self.group.is_some() {
            "group"
        } else if self.raw.is_some() {
            "raw"
        } else {
            "unknown"
        }
    }
}

#[pymethods]
impl PyDrawing {
    #[new]
    #[pyo3(signature=(payload, *, anchor=None, transform=None, meta=None))]
    fn new(
        payload: &Bound<'_, PyAny>,
        anchor: Option<PyRef<'_, PyDrawingAnchor>>,
        transform: Option<PyRef<'_, PyChildTransform>>,
        meta: Option<PyRef<'_, PyDrawingMeta>>,
    ) -> PyResult<Self> {
        if anchor.is_some() == transform.is_some() {
            return Err(PyValueError::new_err(
                "drawing requires exactly one of anchor or transform",
            ));
        }

        let comment_payload = payload.extract::<PyRef<'_, PyDrawingComment>>().ok();
        let mut drawing = Self::empty(
            meta.map(|meta| meta.clone()).unwrap_or_default(),
            anchor.map(|anchor| anchor.clone()),
            transform.map(|transform| transform.clone()),
        );

        if let Ok(image) = payload.extract::<PyRef<'_, PyEmbeddedImage>>() {
            drawing.image = Some(image.clone());
        } else if let Ok(chart) = payload.extract::<PyRef<'_, PyChart>>() {
            drawing.chart = Some(chart.clone());
        } else if let Ok(chart) = payload.extract::<PyRef<'_, PyChartEx>>() {
            drawing.chart_ex = Some(chart.clone());
        } else if let Ok(control) = payload.extract::<PyRef<'_, PyFormControl>>() {
            drawing.form_control = Some(control.clone());
        } else if let Some(comment) = comment_payload {
            drawing.comment = Some(comment.clone());
        } else if let Ok(shape) = payload.extract::<PyRef<'_, PyShape>>() {
            drawing.shape = Some(shape.clone());
        } else if let Ok(group) = payload.extract::<PyRef<'_, PyDrawingGroup>>() {
            drawing.group = Some(group.clone());
        } else if payload.extract::<PyRef<'_, PyRawDrawing>>().is_ok() {
            return Err(PyValueError::new_err(
                "raw drawings are read-only and cannot be used as mutation input",
            ));
        } else {
            return Err(PyTypeError::new_err(
                "payload must be EmbeddedImage, Chart, ChartEx, FormControl, DrawingComment, Shape, or DrawingGroup",
            ));
        }

        if drawing.anchor.is_some() {
            drawing.to_top_level()?;
        } else {
            drawing.to_group_child()?;
        }
        Ok(drawing)
    }

    #[getter]
    fn drawing_path(&self) -> Vec<usize> {
        self.drawing_path.clone()
    }

    /// Resolved on-sheet placement in EMU: the anchor rectangle for
    /// top-level drawings, the group-mapped (rotation/flip aware)
    /// rectangle for group children. ``None`` for drawings
    /// constructed in Python rather than read from a sheet.
    #[getter]
    fn absolute_rect_emu(&self) -> Option<PyRectEmu> {
        self.absolute_rect_emu
    }

    #[getter]
    fn kind(&self) -> &'static str {
        self.kind_name()
    }

    #[getter]
    fn meta(&self) -> PyDrawingMeta {
        self.meta.clone()
    }

    #[getter]
    fn name(&self) -> Option<String> {
        self.meta.name.clone()
    }

    /// Resolved visibility: unset meta defaults by kind (comments
    /// hide by default).
    #[getter]
    fn hidden(&self) -> bool {
        self.meta.hidden.unwrap_or(self.comment.is_some())
    }

    #[getter]
    fn locked(&self) -> bool {
        self.meta.locked
    }

    #[getter]
    fn printable(&self) -> bool {
        self.meta.printable
    }

    #[getter]
    fn alt_text(&self) -> Option<String> {
        self.meta.alt_text.clone()
    }

    #[getter]
    fn title(&self) -> Option<String> {
        self.meta.title.clone()
    }

    #[getter]
    fn anchor(&self) -> Option<PyDrawingAnchor> {
        self.anchor.clone()
    }

    #[getter]
    fn transform(&self) -> Option<PyChildTransform> {
        self.transform.clone()
    }

    #[getter]
    fn image(&self) -> Option<PyEmbeddedImage> {
        self.image.clone()
    }

    #[getter]
    fn chart(&self) -> Option<PyChart> {
        self.chart.clone()
    }

    #[getter]
    fn chart_ex(&self) -> Option<PyChartEx> {
        self.chart_ex.clone()
    }

    #[getter]
    fn form_control(&self) -> Option<PyFormControl> {
        self.form_control.clone()
    }

    #[getter]
    fn control(&self) -> Option<PyFormControl> {
        self.form_control.clone()
    }

    #[getter]
    fn comment(&self) -> Option<PyDrawingComment> {
        self.comment.clone()
    }

    #[getter]
    fn shape(&self) -> Option<PyShape> {
        self.shape.clone()
    }

    #[getter]
    fn group(&self) -> Option<PyDrawingGroup> {
        self.group.clone()
    }

    #[getter]
    fn raw(&self) -> Option<PyRawDrawing> {
        self.raw.clone()
    }

    fn __repr__(&self) -> String {
        format!(
            "Drawing(kind={:?}, drawing_path={:?})",
            self.kind_name(),
            self.drawing_path
        )
    }
}

#[pyclass(name = "FormControlInteractionResult")]
#[derive(Clone)]
pub struct PyFormControlInteractionResult {
    #[pyo3(get)]
    controls_changed: usize,
    #[pyo3(get)]
    linked_cells_changed: usize,
}

/// Map a core drawing-mutation error to the Python exception type the
/// binding has always raised: positional problems (bad index or path)
/// as `IndexError`, content problems as `ValueError`.
fn drawing_mutation_err(error: core::Error) -> PyErr {
    let text = error.to_string();
    if text.contains("out of bounds") || text.contains("path") || text.contains("not a group") {
        PyIndexError::new_err(text)
    } else {
        PyValueError::new_err(text)
    }
}

fn collect_filtered_drawings(
    sheet: &core::Worksheet,
    object: &core::DrawingObject,
    top_index: usize,
    predicate: fn(&core::DrawingKind) -> bool,
    output: &mut Vec<PyDrawing>,
) {
    #[allow(clippy::too_many_arguments)]
    fn walk(
        sheet: &core::Worksheet,
        meta: &core::DrawingMeta,
        kind: &core::DrawingKind,
        path: Vec<usize>,
        anchor: Option<&chart::DrawingAnchor>,
        transform: Option<&chart::ChildTransform>,
        predicate: fn(&core::DrawingKind) -> bool,
        output: &mut Vec<PyDrawing>,
    ) {
        if predicate(kind) {
            output.push(PyDrawing::from_kind(
                sheet, meta, kind, path.clone(), anchor, transform,
            ));
        }
        if let core::DrawingKind::Group(group) = kind {
            for (index, child) in group.children.iter().enumerate() {
                let mut child_path = path.clone();
                child_path.push(index);
                walk(
                    sheet,
                    &child.meta,
                    &child.kind,
                    child_path,
                    None,
                    Some(&child.transform),
                    predicate,
                    output,
                );
            }
        }
    }

    walk(
        sheet,
        &object.meta,
        &object.kind,
        vec![top_index],
        Some(&object.anchor),
        None,
        predicate,
        output,
    );
}

fn count_nested_kind(kind: &core::DrawingKind, predicate: fn(&core::DrawingKind) -> bool) -> usize {
    let own = usize::from(predicate(kind));
    let children = match kind {
        core::DrawingKind::Group(group) => group
            .children
            .iter()
            .map(|child| count_nested_kind(&child.kind, predicate))
            .sum(),
        _ => 0,
    };
    own + children
}

#[pymethods]
impl PyWorksheet {
    /// Drawing objects in z-order, with groups represented recursively.
    #[getter]
    fn drawings(&self) -> PyResult<Vec<PyDrawing>> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(worksheet
            .drawings()
            .iter()
            .enumerate()
            .map(|(index, object)| PyDrawing::from_top(worksheet, object, index))
            .collect())
    }

    /// Images at any group depth, in depth-first drawing order.
    #[getter]
    fn images(&self) -> PyResult<Vec<PyDrawing>> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let mut images = Vec::new();
        for (index, object) in worksheet.drawings().iter().enumerate() {
            collect_filtered_drawings(
                worksheet,
                object,
                index,
                |kind| matches!(kind, core::DrawingKind::Image(_)),
                &mut images,
            );
        }
        Ok(images)
    }

    #[getter]
    fn image_count(&self) -> PyResult<usize> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(worksheet
            .drawings()
            .iter()
            .map(|object| {
                count_nested_kind(&object.kind, |kind| {
                    matches!(kind, core::DrawingKind::Image(_))
                })
            })
            .sum())
    }

    /// Form controls at any group depth, in depth-first drawing order.
    #[getter]
    fn form_controls(&self) -> PyResult<Vec<PyDrawing>> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let mut controls = Vec::new();
        for (index, object) in worksheet.drawings().iter().enumerate() {
            collect_filtered_drawings(
                worksheet,
                object,
                index,
                |kind| matches!(kind, core::DrawingKind::FormControl(_)),
                &mut controls,
            );
        }
        Ok(controls)
    }

    #[getter]
    fn form_control_count(&self) -> PyResult<usize> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(worksheet
            .drawings()
            .iter()
            .map(|object| {
                count_nested_kind(&object.kind, |kind| {
                    matches!(kind, core::DrawingKind::FormControl(_))
                })
            })
            .sum())
    }

    /// Standard charts at any group depth, in depth-first drawing order.
    #[getter]
    fn charts(&self) -> PyResult<Vec<PyDrawing>> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let mut charts = Vec::new();
        for (index, object) in worksheet.drawings().iter().enumerate() {
            collect_filtered_drawings(
                worksheet,
                object,
                index,
                |kind| matches!(kind, core::DrawingKind::Chart(_)),
                &mut charts,
            );
        }
        Ok(charts)
    }

    #[getter]
    fn chart_count(&self) -> PyResult<usize> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(worksheet
            .drawings()
            .iter()
            .map(|object| {
                count_nested_kind(&object.kind, |kind| {
                    matches!(kind, core::DrawingKind::Chart(_))
                })
            })
            .sum())
    }

    /// ChartEx charts at any group depth, in depth-first drawing order.
    #[getter]
    fn charts_ex(&self) -> PyResult<Vec<PyDrawing>> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let mut charts = Vec::new();
        for (index, object) in worksheet.drawings().iter().enumerate() {
            collect_filtered_drawings(
                worksheet,
                object,
                index,
                |kind| matches!(kind, core::DrawingKind::ChartEx(_)),
                &mut charts,
            );
        }
        Ok(charts)
    }

    #[getter]
    fn chart_ex_count(&self) -> PyResult<usize> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(worksheet
            .drawings()
            .iter()
            .map(|object| {
                count_nested_kind(&object.kind, |kind| {
                    matches!(kind, core::DrawingKind::ChartEx(_))
                })
            })
            .sum())
    }

    /// Validate and append a top-level drawing, returning its z-order index.
    fn add_drawing(&self, drawing: PyRef<'_, PyDrawing>) -> PyResult<usize> {
        let object = drawing.to_top_level()?;
        let mut workbook = self.workbook.write().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        worksheet.add_drawing(object).map_err(drawing_mutation_err)
    }

    /// Insert a top-level drawing at a z-order index. Drawing paths
    /// are positional; mutating the list invalidates previously
    /// returned paths.
    fn insert_drawing(&self, index: usize, drawing: PyRef<'_, PyDrawing>) -> PyResult<()> {
        let object = drawing.to_top_level()?;
        let mut workbook = self.workbook.write().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        worksheet
            .insert_drawing(index, object)
            .map_err(drawing_mutation_err)
    }

    /// Replace a top-level drawing or nested group child by path.
    /// Drawing paths are positional; mutating the list invalidates
    /// previously returned paths.
    fn set_drawing(&self, path: Vec<usize>, drawing: PyRef<'_, PyDrawing>) -> PyResult<()> {
        let (&top_index, rest) = path
            .split_first()
            .ok_or_else(|| PyIndexError::new_err("drawing path cannot be empty"))?;
        let mut workbook = self.workbook.write().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        if rest.is_empty() {
            let object = drawing.to_top_level()?;
            return worksheet
                .set_drawing(top_index, object)
                .map_err(drawing_mutation_err);
        }

        let child = drawing.to_group_child()?;
        worksheet
            .set_group_child(&path, child)
            .map_err(drawing_mutation_err)
    }

    /// Remove a top-level drawing or nested group child by path.
    /// Drawing paths are positional; mutating the list invalidates
    /// previously returned paths.
    fn remove_drawing(&self, path: Vec<usize>) -> PyResult<()> {
        let (&top_index, rest) = path
            .split_first()
            .ok_or_else(|| PyIndexError::new_err("drawing path cannot be empty"))?;
        let mut workbook = self.workbook.write().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        if rest.is_empty() {
            return worksheet
                .remove_drawing(top_index)
                .map(|_| ())
                .map_err(|error| PyIndexError::new_err(error.to_string()));
        }
        worksheet
            .remove_group_child(&path)
            .map(|_| ())
            .map_err(drawing_mutation_err)
    }

    /// Move a top-level drawing within the z-order list. Drawing
    /// paths are positional; mutating the list invalidates previously
    /// returned paths.
    fn move_drawing(&self, from_index: usize, to_index: usize) -> PyResult<()> {
        let mut workbook = self.workbook.write().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        worksheet
            .move_drawing(from_index, to_index)
            .map_err(|error| PyIndexError::new_err(error.to_string()))
    }

    /// Copy an image's encoded bytes into a Python bytes object on
    /// demand. Paths are positional; mutating the drawing list
    /// invalidates previously returned paths.
    fn drawing_image_data(
        &self,
        py: Python<'_>,
        path: Vec<usize>,
    ) -> PyResult<Py<PyBytes>> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let node = worksheet.drawing_at_path(&path).ok_or_else(|| {
            PyIndexError::new_err(format!("no drawing at path {path:?}"))
        })?;
        let core::DrawingKind::Image(image) = node.kind else {
            return Err(PyValueError::new_err(format!(
                "drawing at path {path:?} is not an image"
            )));
        };
        Ok(PyBytes::new_bound(py, image.data()).unbind())
    }

    /// Copy an image's SVG variant on demand, if present. Paths are
    /// positional; mutating the drawing list invalidates previously
    /// returned paths.
    fn drawing_svg_data(
        &self,
        py: Python<'_>,
        path: Vec<usize>,
    ) -> PyResult<Option<Py<PyBytes>>> {
        let workbook = self.workbook.read().map_err(to_py_err)?;
        let worksheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let node = worksheet.drawing_at_path(&path).ok_or_else(|| {
            PyIndexError::new_err(format!("no drawing at path {path:?}"))
        })?;
        let core::DrawingKind::Image(image) = node.kind else {
            return Err(PyValueError::new_err(format!(
                "drawing at path {path:?} is not an image"
            )));
        };
        Ok(image
            .svg_data()
            .map(|bytes| PyBytes::new_bound(py, bytes).unbind()))
    }

    /// Apply checkbox/radio semantics and immediately update linked cells.
    fn set_form_control_check_state(
        &self,
        path: Vec<usize>,
        state: &str,
    ) -> PyResult<PyFormControlInteractionResult> {
        let state = match state {
            "unchecked" => core::CheckState::Unchecked,
            "checked" => core::CheckState::Checked,
            "mixed" => core::CheckState::Mixed,
            _ => {
                return Err(PyValueError::new_err(
                    "state must be unchecked, checked, or mixed",
                ))
            }
        };
        let mut workbook = self.workbook.write().map_err(to_py_err)?;
        let result = workbook
            .set_form_control_check_state(self.sheet_index, &path, state)
            .map_err(|error| PyValueError::new_err(error.to_string()))?;
        Ok(PyFormControlInteractionResult {
            controls_changed: result.controls_changed,
            linked_cells_changed: result.linked_cells_changed,
        })
    }
}

#[pymethods]
impl PyWorkbook {
    /// Project all form-control state into linked cells.
    fn sync_form_controls(&self) -> PyResult<usize> {
        let mut workbook = self.inner.write().map_err(to_py_err)?;
        Ok(workbook.sync_form_control_links())
    }

    /// Drive controls from linked cells that contain formulas.
    fn sync_form_controls_from_linked_cells(&self) -> PyResult<usize> {
        let mut workbook = self.inner.write().map_err(to_py_err)?;
        Ok(workbook.sync_form_controls_from_linked_cells())
    }
}
