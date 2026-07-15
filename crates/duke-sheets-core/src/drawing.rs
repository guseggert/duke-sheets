//! Unified drawing model
//!
//! Every floating object on a worksheet - images, charts, basic shapes,
//! form controls, cell comments, shape groups, and preserved raw drawing
//! fragments - lives in one per-sheet list of [`DrawingObject`]s.
//! **List order is z-order**, back to front: index 0 renders at the
//! bottom, the last element on top. This mirrors Excel's own object
//! model (one `Shapes` collection per sheet) and the file formats,
//! where the drawing part's document order (XLSX/XLSB) and the
//! OfficeArt container order (XLS) encode stacking.
//!
//! Shared non-visual properties (name, hidden, locked, printable,
//! alternative text, title) are hoisted into [`DrawingMeta`] on the
//! wrapper; kind-specific payloads live in [`DrawingKind`].
//!
//! Group children are positioned in their group's child coordinate
//! space, not with sheet anchors, so they are represented by
//! [`GroupChild`] (meta + [`ChildTransform`] + kind) rather than
//! [`DrawingObject`]. Use [`crate::Worksheet::placed_form_controls`]
//! or [`crate::Worksheet::drawings_flat`] to traverse nested content
//! with paths and resolved on-sheet rectangles.
//!
//! Cell comments participate in the list (their VML shapes share the
//! z-order domain with form controls in the file formats) while
//! keeping the keyed `(row, col)` accessors on
//! [`crate::Worksheet`] as views over the list.

use duke_sheets_chart::{Chart, ChartEx, DrawingAnchor, EmbeddedImage};
pub use duke_sheets_chart::{ChildTransform, GroupTransform};

use crate::comment::CellComment;
use crate::form_control::FormControl;
use crate::{
    rich_text_to_plain, Color, Error, HorizontalAlignment, Result, RichTextRun, RunFont,
    VerticalAlignment, MAX_COLS, MAX_ROWS,
};

/// Rich text and alignment shared by drawing shapes and form controls.
///
/// Empty text is valid and canonically represented as **zero runs**,
/// never as a single run holding an empty string: readers produce no
/// runs for absent captions and [`DrawingText::plain`] with an empty
/// string produces no runs, so empty captions round-trip equal.
/// Alignments are optional so a format's default can remain implicit.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct DrawingText {
    /// Text runs in display order.
    pub runs: Vec<RichTextRun>,
    /// Explicit horizontal alignment.
    pub horizontal_alignment: Option<HorizontalAlignment>,
    /// Explicit vertical alignment.
    pub vertical_alignment: Option<VerticalAlignment>,
}

impl DrawingText {
    /// Create text containing one unformatted run, or no runs at all
    /// when `text` is empty (the canonical empty-caption form).
    pub fn plain(text: impl Into<String>) -> Self {
        let text = text.into();
        // Empty text is no runs, matching what readers produce for
        // absent captions, so empty round-trips compare equal.
        let runs = if text.is_empty() {
            Vec::new()
        } else {
            vec![RichTextRun::plain(text)]
        };
        Self {
            runs,
            horizontal_alignment: None,
            vertical_alignment: None,
        }
    }

    /// Concatenate all runs without formatting.
    pub fn plain_text(&self) -> String {
        rich_text_to_plain(&self.runs)
    }

    /// Whether the concatenated text is empty.
    pub fn is_empty(&self) -> bool {
        self.runs.iter().all(|run| run.text.is_empty())
    }

    /// Convert to the model-neutral DrawingML text representation used
    /// by the shared XLSX/XLSB drawing-part codec.
    #[doc(hidden)]
    pub fn to_drawing_part_text(&self) -> duke_sheets_chart::drawing_part::TwinText {
        use duke_sheets_chart::drawing_part::{
            TwinHorizontalAlignment as H, TwinRunFont, TwinText, TwinTextRun, TwinUnderline,
            TwinVerticalAlignment as V,
        };
        let horizontal_alignment = self.horizontal_alignment.map(|alignment| match alignment {
            HorizontalAlignment::Center | HorizontalAlignment::CenterContinuous => H::Center,
            HorizontalAlignment::Right => H::Right,
            HorizontalAlignment::Justify => H::Justify,
            HorizontalAlignment::Distributed => H::Distributed,
            HorizontalAlignment::General
            | HorizontalAlignment::Left
            | HorizontalAlignment::Fill => H::Left,
        });
        let vertical_alignment = self.vertical_alignment.map(|alignment| match alignment {
            VerticalAlignment::Top => V::Top,
            VerticalAlignment::Center => V::Center,
            VerticalAlignment::Bottom => V::Bottom,
            VerticalAlignment::Justify => V::Justify,
            VerticalAlignment::Distributed => V::Distributed,
        });
        let runs = self
            .runs
            .iter()
            .map(|run| TwinTextRun {
                text: run.text.clone(),
                font: run.font.as_ref().map(|font| TwinRunFont {
                    name: font.name.clone(),
                    size: font.size,
                    color: font.color.and_then(color_to_drawing_part),
                    bold: font.bold,
                    italic: font.italic,
                    underline: font.underline.and_then(|underline| match underline {
                        crate::style::Underline::None => None,
                        crate::style::Underline::Single => Some(TwinUnderline::Single),
                        crate::style::Underline::Double => Some(TwinUnderline::Double),
                        crate::style::Underline::SingleAccounting => {
                            Some(TwinUnderline::SingleAccounting)
                        }
                        crate::style::Underline::DoubleAccounting => {
                            Some(TwinUnderline::DoubleAccounting)
                        }
                    }),
                    strikethrough: font.strikethrough,
                    baseline: font.vertical_align.map(|alignment| match alignment {
                        crate::style::FontVerticalAlign::Baseline => 0,
                        crate::style::FontVerticalAlign::Superscript => 30_000,
                        crate::style::FontVerticalAlign::Subscript => -25_000,
                    }),
                }),
            })
            .collect();
        TwinText {
            runs,
            horizontal_alignment,
            vertical_alignment,
        }
    }

    /// Convert shared DrawingML text into the public drawing model.
    #[doc(hidden)]
    pub fn from_drawing_part_text(text: &duke_sheets_chart::drawing_part::TwinText) -> Self {
        use duke_sheets_chart::drawing_part::{
            TwinHorizontalAlignment as H, TwinUnderline, TwinVerticalAlignment as V,
        };
        Self {
            runs: text
                .runs
                .iter()
                .map(|run| RichTextRun {
                    text: run.text.clone(),
                    font: run.font.as_ref().map(|font| RunFont {
                        name: font.name.clone(),
                        size: font.size,
                        color: font.color.map(color_from_drawing_part),
                        bold: font.bold,
                        italic: font.italic,
                        underline: font.underline.map(|underline| match underline {
                            TwinUnderline::Single => crate::style::Underline::Single,
                            TwinUnderline::Double => crate::style::Underline::Double,
                            TwinUnderline::SingleAccounting => {
                                crate::style::Underline::SingleAccounting
                            }
                            TwinUnderline::DoubleAccounting => {
                                crate::style::Underline::DoubleAccounting
                            }
                        }),
                        strikethrough: font.strikethrough,
                        vertical_align: font.baseline.map(|baseline| {
                            if baseline > 0 {
                                crate::style::FontVerticalAlign::Superscript
                            } else if baseline < 0 {
                                crate::style::FontVerticalAlign::Subscript
                            } else {
                                crate::style::FontVerticalAlign::Baseline
                            }
                        }),
                        ..RunFont::default()
                    }),
                })
                .collect(),
            horizontal_alignment: text.horizontal_alignment.map(|alignment| match alignment {
                H::Left => HorizontalAlignment::Left,
                H::Center => HorizontalAlignment::Center,
                H::Right => HorizontalAlignment::Right,
                H::Justify => HorizontalAlignment::Justify,
                H::Distributed => HorizontalAlignment::Distributed,
            }),
            vertical_alignment: text.vertical_alignment.map(|alignment| match alignment {
                V::Top => VerticalAlignment::Top,
                V::Center => VerticalAlignment::Center,
                V::Bottom => VerticalAlignment::Bottom,
                V::Justify => VerticalAlignment::Justify,
                V::Distributed => VerticalAlignment::Distributed,
            }),
        }
    }
}

/// Convert a core color to the shared DrawingML representation.
#[doc(hidden)]
pub fn color_to_drawing_part(color: Color) -> Option<duke_sheets_chart::drawing_part::TwinColor> {
    use duke_sheets_chart::drawing_part::TwinColor;
    match color {
        Color::Auto => None,
        Color::Rgb { r, g, b } | Color::Argb { r, g, b, .. } => Some(TwinColor::Rgb { r, g, b }),
        Color::Theme { index, tint } => Some(TwinColor::Theme { index, tint }),
        Color::Indexed(_) => {
            let (r, g, b) = color.to_rgb();
            Some(TwinColor::Rgb { r, g, b })
        }
    }
}

/// Convert a shared DrawingML color to the core representation.
#[doc(hidden)]
pub fn color_from_drawing_part(color: duke_sheets_chart::drawing_part::TwinColor) -> Color {
    match color {
        duke_sheets_chart::drawing_part::TwinColor::Rgb { r, g, b } => Color::rgb(r, g, b),
        duke_sheets_chart::drawing_part::TwinColor::Theme { index, tint } => {
            Color::theme(index, tint)
        }
    }
}

impl From<String> for DrawingText {
    fn from(text: String) -> Self {
        Self::plain(text)
    }
}

impl From<&str> for DrawingText {
    fn from(text: &str) -> Self {
        Self::plain(text)
    }
}

/// The geometry of a worksheet shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ShapeGeometry {
    /// A DrawingML preset geometry name, such as `rect`, `ellipse`, or
    /// `roundRect`.
    Preset(String),
}

impl Default for ShapeGeometry {
    fn default() -> Self {
        Self::Preset("rect".to_string())
    }
}

/// Fill applied to a worksheet shape.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum ShapeFill {
    /// No visible fill.
    #[default]
    None,
    /// A solid fill color.
    Solid(Color),
}

/// Outline applied to a worksheet shape.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ShapeLine {
    /// Outline color. `None` leaves the format default implicit.
    pub color: Option<Color>,
    /// Outline width in English Metric Units (12,700 EMU = 1 point).
    pub width_emu: Option<i64>,
    /// DrawingML preset dash name, such as `solid`, `dash`, or `dot`.
    pub dash_style: Option<String>,
    /// Whether the outline is explicitly disabled.
    pub no_fill: bool,
}

/// A basic worksheet shape.
///
/// Modeled fields are always authoritative when writing. The hidden
/// raw shape properties contain mergeable child fragments from OOXML
/// `spPr`; the raw text body can retain a complete `txBody`. Parsed
/// snapshots detect direct public-field edits, so writers regenerate
/// changed modeled blocks instead of replaying stale XML. Unsupported
/// effects, gradients, custom geometry, and extension children survive
/// an untouched round trip.
#[derive(Debug, Clone)]
pub struct Shape {
    /// Preset geometry.
    pub geometry: ShapeGeometry,
    /// Shape fill.
    pub fill: ShapeFill,
    /// Shape outline.
    pub line: ShapeLine,
    /// Optional rich text body.
    pub text: Option<DrawingText>,
    /// Clockwise rotation in 60,000ths of a degree.
    pub rotation: i32,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
    /// Unmodeled `spPr` child XML fragments preserved by OOXML codecs.
    #[doc(hidden)]
    pub raw_shape_properties: Option<Vec<u8>>,
    /// Complete `txBody` XML preserved by OOXML codecs.
    #[doc(hidden)]
    pub raw_text_body: Option<Vec<u8>>,
    /// Parsed geometry paired with raw fragments for stale-fragment detection.
    #[doc(hidden)]
    pub raw_geometry_snapshot: Option<ShapeGeometry>,
    /// Parsed fill paired with raw fragments for stale-fragment detection.
    #[doc(hidden)]
    pub raw_fill_snapshot: Option<ShapeFill>,
    /// Parsed line paired with raw fragments for stale-fragment detection.
    #[doc(hidden)]
    pub raw_line_snapshot: Option<ShapeLine>,
    /// Parsed text paired with a retained complete OOXML `txBody`.
    #[doc(hidden)]
    pub raw_text_snapshot: Option<Option<DrawingText>>,
}

/// Two shapes are equal when they serialize identically: same modeled
/// fields, same preserved raw bytes, and the same staleness outcome
/// for each preserved block (a stale preserved geometry regenerates
/// while a current one replays, so staleness is emission-relevant
/// even though the snapshots themselves are bookkeeping).
impl PartialEq for Shape {
    fn eq(&self, other: &Self) -> bool {
        self.geometry == other.geometry
            && self.fill == other.fill
            && self.line == other.line
            && self.text == other.text
            && self.rotation == other.rotation
            && self.flip_h == other.flip_h
            && self.flip_v == other.flip_v
            && self.raw_shape_properties == other.raw_shape_properties
            && self.raw_text_body == other.raw_text_body
            && self.preserved_geometry_unchanged() == other.preserved_geometry_unchanged()
            && self.preserved_fill_unchanged() == other.preserved_fill_unchanged()
            && self.preserved_line_unchanged() == other.preserved_line_unchanged()
            && self.preserved_text_unchanged() == other.preserved_text_unchanged()
    }
}

impl Default for Shape {
    fn default() -> Self {
        Self {
            geometry: ShapeGeometry::default(),
            fill: ShapeFill::None,
            line: ShapeLine::default(),
            text: None,
            rotation: 0,
            flip_h: false,
            flip_v: false,
            raw_shape_properties: None,
            raw_text_body: None,
            raw_geometry_snapshot: None,
            raw_fill_snapshot: None,
            raw_line_snapshot: None,
            raw_text_snapshot: None,
        }
    }
}

impl Shape {
    /// Create a rectangle with no explicit fill or outline.
    pub fn rectangle() -> Self {
        Self::default()
    }

    /// Create a shape using a DrawingML preset geometry name.
    pub fn preset(name: impl Into<String>) -> Self {
        Self {
            geometry: ShapeGeometry::Preset(name.into()),
            ..Self::default()
        }
    }

    /// Set the geometry.
    pub fn set_geometry(&mut self, geometry: ShapeGeometry) {
        self.geometry = geometry;
    }

    /// Set the geometry and return the shape.
    pub fn with_geometry(mut self, geometry: ShapeGeometry) -> Self {
        self.set_geometry(geometry);
        self
    }

    /// Set the fill.
    pub fn set_fill(&mut self, fill: ShapeFill) {
        self.fill = fill;
    }

    /// Set the fill and return the shape.
    pub fn with_fill(mut self, fill: ShapeFill) -> Self {
        self.set_fill(fill);
        self
    }

    /// Set the outline.
    pub fn set_line(&mut self, line: ShapeLine) {
        self.line = line;
    }

    /// Set the outline and return the shape.
    pub fn with_line(mut self, line: ShapeLine) -> Self {
        self.set_line(line);
        self
    }

    /// Set the text body.
    pub fn set_text(&mut self, text: impl Into<DrawingText>) {
        self.text = Some(text.into());
    }

    /// Set the text body and return the shape.
    pub fn with_text(mut self, text: impl Into<DrawingText>) -> Self {
        self.set_text(text);
        self
    }

    /// Remove the text body.
    pub fn clear_text(&mut self) {
        self.text = None;
    }

    /// Set clockwise rotation in 60,000ths of a degree.
    pub fn set_rotation(&mut self, rotation: i32) {
        self.rotation = rotation;
    }

    /// Set clockwise rotation and return the shape.
    pub fn with_rotation(mut self, rotation: i32) -> Self {
        self.set_rotation(rotation);
        self
    }

    /// Set the horizontal flip.
    pub fn set_flip_h(&mut self, flip_h: bool) {
        self.flip_h = flip_h;
    }

    /// Set the horizontal flip and return the shape.
    pub fn with_flip_h(mut self, flip_h: bool) -> Self {
        self.set_flip_h(flip_h);
        self
    }

    /// Set the vertical flip.
    pub fn set_flip_v(&mut self, flip_v: bool) {
        self.flip_v = flip_v;
    }

    /// Set the vertical flip and return the shape.
    pub fn with_flip_v(mut self, flip_v: bool) -> Self {
        self.set_flip_v(flip_v);
        self
    }

    /// Record preserved OOXML `spPr` fragments and the modeled state
    /// from which they were parsed.
    #[doc(hidden)]
    pub fn set_preserved_shape_properties(&mut self, fragments: Option<Vec<u8>>) {
        self.raw_shape_properties = fragments;
        self.raw_geometry_snapshot = Some(self.geometry.clone());
        self.raw_fill_snapshot = Some(self.fill.clone());
        self.raw_line_snapshot = Some(self.line.clone());
    }

    /// Record a complete OOXML `txBody` and its parsed modeled text.
    #[doc(hidden)]
    pub fn set_preserved_text_body(&mut self, fragments: Option<Vec<u8>>) {
        self.raw_text_body = fragments;
        self.raw_text_snapshot = Some(self.text.clone());
    }

    /// Whether a preserved custom geometry still corresponds to the
    /// modeled geometry. Caller-supplied fragments have no parsed
    /// snapshot and are treated as current.
    #[doc(hidden)]
    pub fn preserved_geometry_unchanged(&self) -> bool {
        self.raw_geometry_snapshot
            .as_ref()
            .is_none_or(|snapshot| snapshot == &self.geometry)
    }

    /// Whether a preserved unsupported fill still corresponds to the
    /// modeled fill.
    #[doc(hidden)]
    pub fn preserved_fill_unchanged(&self) -> bool {
        self.raw_fill_snapshot
            .as_ref()
            .is_none_or(|snapshot| snapshot == &self.fill)
    }

    /// Whether a preserved line block still corresponds to the modeled line.
    #[doc(hidden)]
    pub fn preserved_line_unchanged(&self) -> bool {
        self.raw_line_snapshot
            .as_ref()
            .is_none_or(|snapshot| snapshot == &self.line)
    }

    /// Whether a preserved complete text body still corresponds to modeled text.
    #[doc(hidden)]
    pub fn preserved_text_unchanged(&self) -> bool {
        self.raw_text_snapshot
            .as_ref()
            .is_none_or(|snapshot| snapshot == &self.text)
    }
}

/// Shared non-visual properties of a drawing object.
///
/// Carrier support varies by kind and format; see FEATURES.md.
/// Comments have no name/alt-text/title carrier in any format, and
/// their `printable` flag is governed by page setup rather than the
/// shape, so those fields are ignored for comment objects.
#[derive(Debug, Clone, PartialEq)]
pub struct DrawingMeta {
    /// Shape name (e.g. "Check Box 1"). `None` lets writers derive one.
    pub name: Option<String>,
    /// Whether the object is hidden. For comments this is the popup
    /// visibility: `hidden = true` means the note is only shown on
    /// hover (Excel's default).
    pub hidden: bool,
    /// Whether the object is locked when the sheet is protected.
    pub locked: bool,
    /// Whether the object is included when the sheet is printed.
    pub printable: bool,
    /// Alternative text for accessibility.
    pub alt_text: Option<String>,
    /// Title (OOXML `cNvPr@title`; no XLS carrier).
    pub title: Option<String>,
}

impl Default for DrawingMeta {
    fn default() -> Self {
        Self {
            name: None,
            hidden: false,
            locked: true,
            printable: true,
            alt_text: None,
            title: None,
        }
    }
}

/// A drawing object anchored on the sheet.
///
/// Position in the worksheet's drawing list is the object's
/// z-position; there is no separate z-index field.
#[derive(Debug, Clone, PartialEq)]
pub struct DrawingObject {
    /// Shared non-visual properties.
    pub meta: DrawingMeta,
    /// Placement on the sheet.
    pub anchor: DrawingAnchor,
    /// Kind-specific payload.
    pub kind: DrawingKind,
}

/// Kind-specific payload of a drawing object.
#[derive(Debug, Clone, PartialEq)]
pub enum DrawingKind {
    /// An embedded picture.
    Image(EmbeddedImage),
    /// An embedded chart.
    Chart(Box<Chart>),
    /// An embedded ChartEx (Office 2016+) chart.
    ChartEx(Box<ChartEx>),
    /// A Forms-toolbar control.
    FormControl(FormControl),
    /// A basic worksheet shape.
    Shape(Box<Shape>),
    /// A cell comment (legacy note). The comment is attached to the
    /// cell at `(row, col)`; the wrapper anchor places the popup box.
    Comment {
        /// Anchored cell row (zero-based).
        row: u32,
        /// Anchored cell column (zero-based).
        col: u16,
        /// Comment content.
        comment: CellComment,
    },
    /// A shape group. Children are positioned in the group's child
    /// coordinate space.
    Group(Box<Group>),
    /// A drawing fragment we do not model, preserved for round-trip.
    Raw(RawDrawing),
}

impl DrawingKind {
    /// Payload as an image, if this is an image.
    pub fn as_image(&self) -> Option<&EmbeddedImage> {
        match self {
            DrawingKind::Image(image) => Some(image),
            _ => None,
        }
    }

    /// Payload as a chart, if this is a chart.
    pub fn as_chart(&self) -> Option<&Chart> {
        match self {
            DrawingKind::Chart(chart) => Some(chart),
            _ => None,
        }
    }

    /// Payload as a ChartEx chart, if this is one.
    pub fn as_chart_ex(&self) -> Option<&ChartEx> {
        match self {
            DrawingKind::ChartEx(chart) => Some(chart),
            _ => None,
        }
    }

    /// Payload as a form control, if this is a form control.
    pub fn as_form_control(&self) -> Option<&FormControl> {
        match self {
            DrawingKind::FormControl(control) => Some(control),
            _ => None,
        }
    }

    /// Mutable payload as a form control, if this is a form control.
    pub fn as_form_control_mut(&mut self) -> Option<&mut FormControl> {
        match self {
            DrawingKind::FormControl(control) => Some(control),
            _ => None,
        }
    }

    /// Payload as a shape, if this is a shape.
    pub fn as_shape(&self) -> Option<&Shape> {
        match self {
            DrawingKind::Shape(shape) => Some(shape),
            _ => None,
        }
    }

    /// Mutable payload as a shape, if this is a shape.
    pub fn as_shape_mut(&mut self) -> Option<&mut Shape> {
        match self {
            DrawingKind::Shape(shape) => Some(shape),
            _ => None,
        }
    }

    /// Comment content and cell, if this is a comment.
    pub fn as_comment(&self) -> Option<(u32, u16, &CellComment)> {
        match self {
            DrawingKind::Comment { row, col, comment } => Some((*row, *col, comment)),
            _ => None,
        }
    }

    /// Payload as a group, if this is a group.
    pub fn as_group(&self) -> Option<&Group> {
        match self {
            DrawingKind::Group(group) => Some(group),
            _ => None,
        }
    }
}

/// A preserved drawing fragment we do not model.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct RawDrawing {
    /// Format-specific serialized bytes (one XLSX/XLSB anchor XML
    /// element).
    #[doc(hidden)]
    pub bytes: Vec<u8>,
    /// Relationships referenced by `bytes`, captured so rewrites do
    /// not emit dangling relationship ids.
    #[doc(hidden)]
    pub rels: Vec<RawRel>,
}

/// A captured relationship referenced by a [`RawDrawing`].
#[derive(Debug, Clone, PartialEq)]
pub struct RawRel {
    /// Relationship id as it appears in the raw bytes (e.g. "rId7").
    pub id: String,
    /// Relationship type URI.
    pub rel_type: String,
    /// Relationship target.
    pub target: String,
    /// True when the target is external (TargetMode="External").
    pub external: bool,
    /// Bytes of the target part, when internal and captured.
    #[doc(hidden)]
    pub part: Option<Vec<u8>>,
}

/// A shape group.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Group {
    /// The group's transform: its own extent in parent space plus the
    /// child coordinate space mapping.
    pub transform: GroupTransform,
    /// Child shapes, in z-order (back to front) within the group.
    pub children: Vec<GroupChild>,
}

/// A shape inside a group, positioned in the group's child space.
#[derive(Debug, Clone, PartialEq)]
pub struct GroupChild {
    /// Shared non-visual properties.
    pub meta: DrawingMeta,
    /// Placement within the group's child coordinate space.
    pub transform: ChildTransform,
    /// Kind-specific payload. Nested groups are allowed.
    pub kind: DrawingKind,
}

/// Path to a drawing node: the first element indexes the worksheet's
/// drawing list, subsequent elements index group children.
pub type DrawingPath = Vec<usize>;

/// Absolute rectangle in the sheet's EMU coordinate space.
///
/// Fields saturate at ±(2^53 − 1) EMU so every value is exactly
/// representable as an IEEE-754 double (and thus a JS number).
/// Real sheet geometry never approaches the bound (2^53 EMU is about
/// 250,000 km); only degenerate group transforms in hostile files
/// saturate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct RectEmu {
    /// Left edge, in EMU from the sheet origin.
    pub x_emu: i64,
    /// Top edge, in EMU from the sheet origin.
    pub y_emu: i64,
    /// Width in EMU. Negative when the source anchor is inverted
    /// (its `to` marker before its `from` marker).
    pub width_emu: i64,
    /// Height in EMU. Negative when the source anchor is inverted.
    pub height_emu: i64,
}

impl RectEmu {
    pub(crate) fn from_corners((x1, y1, x2, y2): CornersEmu) -> Self {
        const JS_SAFE: i128 = (1i128 << 53) - 1;
        let clamp = |v: i128| -> i64 { v.clamp(-JS_SAFE, JS_SAFE) as i64 };
        Self {
            x_emu: clamp(x1),
            y_emu: clamp(y1),
            width_emu: clamp(x2.saturating_sub(x1)),
            height_emu: clamp(y2.saturating_sub(y1)),
        }
    }
}

/// Internal corner-form rectangle `(x1, y1, x2, y2)` used by the
/// group transform math; i128 keeps hostile nested transforms from
/// overflowing before the final [`RectEmu`] saturation.
pub(crate) type CornersEmu = (i128, i128, i128, i128);

/// A filtered view item: a typed payload together with its wrapper
/// object and position in the drawing list. Yielded by the
/// [`crate::Worksheet::images`]/[`crate::Worksheet::charts`]/
/// [`crate::Worksheet::form_controls`] family, which cover top-level
/// objects only (group children have no list index).
#[derive(Debug)]
pub struct Drawn<'a, T> {
    /// Index into the worksheet's drawing list (= z-position).
    pub index: usize,
    /// The wrapper object (meta + anchor).
    pub object: &'a DrawingObject,
    /// The typed payload.
    pub payload: &'a T,
}

/// A comment together with its cell, wrapper object, and position in
/// the drawing list.
#[derive(Debug)]
pub struct CommentRef<'a> {
    /// Index into the worksheet's drawing list (= z-position).
    pub index: usize,
    /// Anchored cell row (zero-based).
    pub row: u32,
    /// Anchored cell column (zero-based).
    pub col: u16,
    /// The wrapper object (meta + anchor).
    pub object: &'a DrawingObject,
    /// Comment content.
    pub comment: &'a CellComment,
}

/// A borrowed view of one drawing node (top-level or group child):
/// the shared metadata and the kind payload.
#[derive(Debug)]
pub struct DrawingNodeRef<'a> {
    /// Shared non-visual properties.
    pub meta: &'a DrawingMeta,
    /// Kind-specific payload.
    pub kind: &'a DrawingKind,
}

/// A mutable view of one drawing node (top-level or group child).
#[derive(Debug)]
pub struct DrawingNodeMut<'a> {
    /// Shared non-visual properties.
    pub meta: &'a mut DrawingMeta,
    /// Kind-specific payload.
    pub kind: &'a mut DrawingKind,
}

impl DrawingObject {
    /// Create an object of the given kind with default metadata and a
    /// default anchor. Comments default to `hidden = true` (popup
    /// shown on hover only), matching Excel.
    pub fn new(kind: DrawingKind) -> Self {
        let hidden = matches!(kind, DrawingKind::Comment { .. });
        Self {
            meta: DrawingMeta {
                hidden,
                ..DrawingMeta::default()
            },
            anchor: DrawingAnchor::default(),
            kind,
        }
    }

    /// Create an image object.
    pub fn image(image: EmbeddedImage) -> Self {
        Self::new(DrawingKind::Image(image))
    }

    /// Create a chart object.
    pub fn chart(chart: Chart) -> Self {
        Self::new(DrawingKind::Chart(Box::new(chart)))
    }

    /// Create a ChartEx object.
    pub fn chart_ex(chart: ChartEx) -> Self {
        Self::new(DrawingKind::ChartEx(Box::new(chart)))
    }

    /// Create a form control object.
    pub fn form_control(control: FormControl) -> Self {
        Self::new(DrawingKind::FormControl(control))
    }

    /// Create a basic shape object.
    pub fn shape(shape: Shape) -> Self {
        Self::new(DrawingKind::Shape(Box::new(shape)))
    }

    /// Create a comment object for the cell at `(row, col)`, anchored
    /// at Excel's default popup placement and hidden by default.
    pub fn comment(row: u32, col: u16, comment: CellComment) -> Self {
        let mut object = Self::new(DrawingKind::Comment { row, col, comment });
        object.anchor = default_comment_anchor(row, col);
        object
    }

    /// Create a group object.
    pub fn group(group: Group) -> Self {
        Self::new(DrawingKind::Group(Box::new(group)))
    }

    /// Create a raw passthrough object.
    pub fn raw(raw: RawDrawing) -> Self {
        Self::new(DrawingKind::Raw(raw))
    }

    /// Set the anchor.
    pub fn with_anchor(mut self, anchor: DrawingAnchor) -> Self {
        self.anchor = anchor;
        self
    }

    /// Set the shape name.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.meta.name = Some(name.into());
        self
    }

    /// Set the hidden flag.
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.meta.hidden = hidden;
        self
    }

    /// Validate the anchor and kind-specific invariants.
    pub fn validate(&self) -> Result<()> {
        validate_anchor(&self.anchor)?;
        validate_kind(&self.kind)
    }
}

fn validate_kind(kind: &DrawingKind) -> Result<()> {
    match kind {
        DrawingKind::FormControl(control) => control.validate()?,
        DrawingKind::Comment { row, col, .. } => {
            if *row >= MAX_ROWS {
                return Err(Error::RowOutOfBounds(*row, MAX_ROWS - 1));
            }
            if *col >= MAX_COLS {
                return Err(Error::ColumnOutOfBounds(*col, MAX_COLS - 1));
            }
        }
        DrawingKind::Group(group) => {
            for child in &group.children {
                if child.transform.cx_emu < 0 || child.transform.cy_emu < 0 {
                    return Err(Error::other("group child extents cannot be negative"));
                }
                if matches!(child.kind, DrawingKind::Comment { .. }) {
                    // Comments are cell-keyed top-level objects; no
                    // format nests them in shape groups.
                    return Err(Error::other("comments cannot be group children"));
                }
                if matches!(child.kind, DrawingKind::Raw(_)) {
                    // Raw payloads are whole anchor elements; no format
                    // can nest them inside a grpSp, and no reader
                    // produces them there.
                    return Err(Error::other("raw drawings cannot be group children"));
                }
                validate_kind(&child.kind)?;
            }
        }
        DrawingKind::Shape(shape) => {
            if matches!(&shape.geometry, ShapeGeometry::Preset(name) if name.trim().is_empty()) {
                return Err(Error::other("shape preset name cannot be empty"));
            }
            if shape.line.width_emu.is_some_and(|width| width < 0) {
                return Err(Error::other("shape line width cannot be negative"));
            }
        }
        DrawingKind::Image(_)
        | DrawingKind::Chart(_)
        | DrawingKind::ChartEx(_)
        | DrawingKind::Raw(_) => {}
    }
    Ok(())
}

/// Validate that an anchor is on the grid with coherent extents.
pub fn validate_anchor(anchor: &DrawingAnchor) -> Result<()> {
    let validate_marker = |marker: &duke_sheets_chart::CellMarker| -> Result<()> {
        if marker.row >= MAX_ROWS {
            return Err(Error::RowOutOfBounds(marker.row, MAX_ROWS - 1));
        }
        if marker.col >= MAX_COLS {
            return Err(Error::ColumnOutOfBounds(marker.col, MAX_COLS - 1));
        }
        Ok(())
    };
    match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            validate_marker(from)?;
            validate_marker(to)?;
            if (to.col, to.col_offset_emu) < (from.col, from.col_offset_emu)
                || (to.row, to.row_offset_emu) < (from.row, from.row_offset_emu)
            {
                return Err(Error::other("drawing anchor endpoints are reversed"));
            }
        }
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            validate_marker(from)?;
            if *width_emu < 0 || *height_emu < 0 {
                return Err(Error::other("drawing anchor dimensions cannot be negative"));
            }
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            if *x_emu < 0 || *y_emu < 0 || *width_emu < 0 || *height_emu < 0 {
                return Err(Error::other(
                    "absolute drawing anchor values cannot be negative",
                ));
            }
        }
    }
    Ok(())
}

/// Absolute EMU rectangle for a drawing anchor at Excel's default
/// cell metrics.
#[cfg(test)]
pub(crate) fn anchor_rect_emu(anchor: &DrawingAnchor) -> RectEmu {
    RectEmu::from_corners(anchor_rect_emu_with_metrics(
        anchor,
        &duke_sheets_chart::DefaultDrawingMetrics,
    ))
}

/// Absolute EMU corner rectangle using the supplied worksheet metrics.
pub(crate) fn anchor_rect_emu_with_metrics(
    anchor: &DrawingAnchor,
    metrics: &(impl duke_sheets_chart::DrawingMetrics + ?Sized),
) -> CornersEmu {
    match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            let (x1, y1) = duke_sheets_chart::marker_position_emu(from, metrics);
            let (x2, y2) = duke_sheets_chart::marker_position_emu(to, metrics);
            (x1, y1, x2, y2)
        }
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            let (x1, y1) = duke_sheets_chart::marker_position_emu(from, metrics);
            (
                x1,
                y1,
                x1 + (*width_emu).max(0) as i128,
                y1 + (*height_emu).max(0) as i128,
            )
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => (
            *x_emu as i128,
            *y_emu as i128,
            *x_emu as i128 + (*width_emu).max(0) as i128,
            *y_emu as i128 + (*height_emu).max(0) as i128,
        ),
    }
}

/// Map a child-space rectangle into the parent-space frame `outer`
/// through a group's child coordinate mapping.
///
/// The group's flips mirror and its rotation turns the child about
/// the outer frame's center (flips first, then rotation, matching
/// DrawingML transform order); the result is the axis-aligned
/// bounding rect of the transformed child frame. A child's own
/// rotation keeps its unrotated frame, like top-level rotated shapes;
/// nested group rotation applies through recursion via the inner
/// group's transform.
pub(crate) fn map_child_rect(
    outer: CornersEmu,
    transform: &GroupTransform,
    child: &ChildTransform,
) -> CornersEmu {
    let (ox1, oy1, ox2, oy2) = outer;
    let (ow, oh) = (
        ox2.saturating_sub(ox1).max(0),
        oy2.saturating_sub(oy1).max(0),
    );
    let ch_w = i128::from(transform.child_cx_emu).max(1);
    let ch_h = i128::from(transform.child_cy_emu).max(1);
    // Hostile files can make delta * extent exceed i128; saturate
    // instead of panicking, keeping the map monotonic.
    let scale = |delta: i128, extent: i128, child_extent: i128| -> i128 {
        match delta.checked_mul(extent) {
            Some(product) => product / child_extent,
            None if delta < 0 => i128::MIN,
            None => i128::MAX,
        }
    };
    let map_x = |x: i128| {
        ox1.saturating_add(scale(
            x.saturating_sub(i128::from(transform.child_x_emu)),
            ow,
            ch_w,
        ))
    };
    let map_y = |y: i128| {
        oy1.saturating_add(scale(
            y.saturating_sub(i128::from(transform.child_y_emu)),
            oh,
            ch_h,
        ))
    };
    let x1 = i128::from(child.x_emu);
    let y1 = i128::from(child.y_emu);
    let x2 = x1 + i128::from(child.cx_emu).max(0);
    let y2 = y1 + i128::from(child.cy_emu).max(0);
    let scaled = (map_x(x1), map_y(y1), map_x(x2), map_y(y2));
    if transform.rotation == 0 && !transform.flip_h && !transform.flip_v {
        return scaled;
    }

    // Flip, then rotate, each corner about the outer frame center;
    // rot is in 60,000ths of a degree, clockwise in y-down coords.
    // Sum in f64: extreme outer frames overflow an i128 addition.
    let center_x = (ox1 as f64 + ox2 as f64) / 2.0;
    let center_y = (oy1 as f64 + oy2 as f64) / 2.0;
    let radians = f64::from(transform.rotation) / 60_000.0 * std::f64::consts::PI / 180.0;
    let (sin, cos) = radians.sin_cos();
    let (sx1, sy1, sx2, sy2) = scaled;
    let corners = [(sx1, sy1), (sx2, sy1), (sx1, sy2), (sx2, sy2)];
    let mut min_x = f64::INFINITY;
    let mut min_y = f64::INFINITY;
    let mut max_x = f64::NEG_INFINITY;
    let mut max_y = f64::NEG_INFINITY;
    for (x, y) in corners {
        let mut dx = x as f64 - center_x;
        let mut dy = y as f64 - center_y;
        if transform.flip_h {
            dx = -dx;
        }
        if transform.flip_v {
            dy = -dy;
        }
        let rx = center_x + dx * cos - dy * sin;
        let ry = center_y + dx * sin + dy * cos;
        min_x = min_x.min(rx);
        min_y = min_y.min(ry);
        max_x = max_x.max(rx);
        max_y = max_y.max(ry);
    }
    (
        min_x.round() as i128,
        min_y.round() as i128,
        max_x.round() as i128,
        max_y.round() as i128,
    )
}

/// Excel's default comment popup placement for a cell: one column to
/// the right, from just above the cell's row to three rows below.
/// When the popup cannot fit to the right (cell within three columns
/// of the grid edge), it goes to the left of the cell, as in Excel,
/// instead of collapsing against the last column.
pub fn default_comment_anchor(row: u32, col: u16) -> DrawingAnchor {
    const PX_EMU: i64 = 9_525;
    let last_col = u32::from(MAX_COLS - 1);
    let cell_col = u32::from(col).min(last_col);
    let (from_col, to_col) = if cell_col + 3 <= last_col {
        (cell_col + 1, cell_col + 3)
    } else {
        (cell_col.saturating_sub(3), cell_col.saturating_sub(1))
    };
    let clamp_row = |r: u32| -> u32 { r.min(MAX_ROWS - 1) };
    DrawingAnchor::TwoCell {
        from: duke_sheets_chart::CellMarker {
            col: from_col as u16,
            col_offset_emu: 15 * PX_EMU,
            row: row.saturating_sub(1),
            row_offset_emu: 10 * PX_EMU,
        },
        to: duke_sheets_chart::CellMarker {
            col: to_col as u16,
            col_offset_emu: 15 * PX_EMU,
            row: clamp_row(row + 3),
            row_offset_emu: 4 * PX_EMU,
        },
        edit_as: None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::form_control::FormControlKind;

    #[test]
    fn meta_defaults_match_excel() {
        let meta = DrawingMeta::default();
        assert!(!meta.hidden);
        assert!(meta.locked);
        assert!(meta.printable);
        assert!(meta.name.is_none());
    }

    #[test]
    fn comment_objects_default_hidden() {
        let object = DrawingObject::comment(4, 2, CellComment::new("a", "t"));
        assert!(object.meta.hidden);
        match &object.anchor {
            DrawingAnchor::TwoCell { from, to, .. } => {
                assert_eq!((from.col, from.row), (3, 3));
                assert_eq!((to.col, to.row), (5, 7));
            }
            other => panic!("expected two-cell anchor, got {other:?}"),
        }
    }

    #[test]
    fn non_comment_objects_default_visible() {
        let object = DrawingObject::form_control(FormControl::new(FormControlKind::Button {
            caption: "Run".into(),
        }));
        assert!(!object.meta.hidden);
        assert!(object.meta.locked);
    }

    #[test]
    fn default_comment_anchor_flips_left_at_the_right_grid_edge() {
        let last_col = MAX_COLS - 1;
        // Excel places the popup left of the cell when it cannot fit
        // on the right; it must never collapse to zero width.
        for col in [last_col, MAX_COLS - 2, MAX_COLS - 3] {
            match default_comment_anchor(5, col) {
                DrawingAnchor::TwoCell { from, to, .. } => {
                    assert_eq!((from.col, to.col), (col - 3, col - 1), "col {col}");
                    assert_eq!((from.row, to.row), (4, 8));
                }
                other => panic!("expected two-cell anchor, got {other:?}"),
            }
        }
        // The last column where the popup still fits on the right.
        match default_comment_anchor(5, MAX_COLS - 4) {
            DrawingAnchor::TwoCell { from, to, .. } => {
                assert_eq!((from.col, to.col), (MAX_COLS - 3, last_col));
            }
            other => panic!("expected two-cell anchor, got {other:?}"),
        }
    }

    #[test]
    fn validate_rejects_reversed_anchor() {
        let object = DrawingObject::form_control(FormControl::new(FormControlKind::Button {
            caption: "b".into(),
        }))
        .with_anchor(DrawingAnchor::TwoCell {
            from: duke_sheets_chart::CellMarker {
                col: 4,
                col_offset_emu: 0,
                row: 4,
                row_offset_emu: 0,
            },
            to: duke_sheets_chart::CellMarker {
                col: 2,
                col_offset_emu: 0,
                row: 2,
                row_offset_emu: 0,
            },
            edit_as: None,
        });
        assert!(object
            .validate()
            .unwrap_err()
            .to_string()
            .contains("reversed"));
    }

    #[test]
    fn validate_recurses_into_groups() {
        let child = GroupChild {
            meta: DrawingMeta::default(),
            transform: ChildTransform::default(),
            kind: DrawingKind::FormControl(FormControl::new(FormControlKind::Dropdown {
                input_range: None,
                cell_link: None,
                selected: None,
                lines: 0, // invalid
                no_3d: false,
            })),
        };
        let object = DrawingObject::group(Group {
            transform: GroupTransform::default(),
            children: vec![child],
        });
        assert!(object.validate().unwrap_err().to_string().contains("lines"));
    }

    #[test]
    fn validate_rejects_raw_group_children() {
        let object = DrawingObject::group(Group {
            transform: GroupTransform::default(),
            children: vec![GroupChild {
                meta: DrawingMeta::default(),
                transform: ChildTransform::default(),
                kind: DrawingKind::Raw(RawDrawing::default()),
            }],
        });
        assert!(object
            .validate()
            .unwrap_err()
            .to_string()
            .contains("raw drawings cannot be group children"));
    }

    #[test]
    fn map_child_rect_scales_into_outer_frame() {
        let outer: CornersEmu = (1000, 2000, 3000, 4000);
        let transform = GroupTransform {
            child_x_emu: 0,
            child_y_emu: 0,
            child_cx_emu: 200,
            child_cy_emu: 100,
            ..GroupTransform::default()
        };
        let child = ChildTransform {
            x_emu: 100,
            y_emu: 50,
            cx_emu: 100,
            cy_emu: 50,
            ..ChildTransform::default()
        };
        // Child occupies the lower-right quadrant of child space.
        assert_eq!(
            map_child_rect(outer, &transform, &child),
            (2000, 3000, 3000, 4000)
        );
    }

    #[test]
    fn map_child_rect_rotates_about_the_outer_frame_center() {
        let outer: CornersEmu = (0, 0, 1000, 1000);
        let child = ChildTransform {
            x_emu: 0,
            y_emu: 0,
            cx_emu: 50,
            cy_emu: 50,
            ..ChildTransform::default()
        };
        let base = GroupTransform {
            child_x_emu: 0,
            child_y_emu: 0,
            child_cx_emu: 100,
            child_cy_emu: 100,
            ..GroupTransform::default()
        };

        // 90 degrees clockwise: the top-left quadrant lands top-right.
        let quarter = GroupTransform {
            rotation: 5_400_000,
            ..base.clone()
        };
        assert_eq!(map_child_rect(outer, &quarter, &child), (500, 0, 1000, 500));

        // 180 degrees: the top-left quadrant lands bottom-right.
        let half = GroupTransform {
            rotation: 10_800_000,
            ..base
        };
        assert_eq!(
            map_child_rect(outer, &half, &child),
            (500, 500, 1000, 1000)
        );
    }

    #[test]
    fn map_child_rect_flips_before_rotating() {
        let outer: CornersEmu = (0, 0, 1000, 1000);
        let child = ChildTransform {
            x_emu: 0,
            y_emu: 0,
            cx_emu: 50,
            cy_emu: 50,
            ..ChildTransform::default()
        };
        // The top-left child quadrant, flipped horizontally (lands
        // top-right), then rotated 90 degrees clockwise about the
        // frame center (lands bottom-right). Rotate-then-flip would
        // land it top-left, (0, 0, 500, 500), so this pins DrawingML's
        // flip-first order.
        let transform = GroupTransform {
            child_x_emu: 0,
            child_y_emu: 0,
            child_cx_emu: 100,
            child_cy_emu: 100,
            rotation: 5_400_000,
            flip_h: true,
            ..GroupTransform::default()
        };
        assert_eq!(
            map_child_rect(outer, &transform, &child),
            (500, 500, 1000, 1000)
        );
    }

    #[test]
    fn map_child_rect_survives_hostile_extents() {
        // ow ~ 2^77 against a child delta ~ 2^64 overflows a naive
        // i128 multiply; the mapping must saturate, not panic.
        let outer: CornersEmu = (0, 0, 1_i128 << 77, 1_i128 << 77);
        let transform = GroupTransform {
            child_x_emu: i64::MIN,
            child_y_emu: i64::MIN,
            child_cx_emu: 1,
            child_cy_emu: 1,
            ..GroupTransform::default()
        };
        let child = ChildTransform {
            x_emu: i64::MAX,
            y_emu: i64::MAX,
            cx_emu: i64::MAX,
            cy_emu: i64::MAX,
            ..ChildTransform::default()
        };
        let (x1, y1, x2, y2) = map_child_rect(outer, &transform, &child);
        assert!(x1 <= x2 && y1 <= y2);

        // The rotation path must survive an extreme outer frame, as
        // produced by saturated nested mappings.
        let extreme: CornersEmu = (i128::MIN, i128::MIN, i128::MAX, i128::MAX);
        let rotated = GroupTransform {
            rotation: 5_400_000,
            flip_h: true,
            ..transform
        };
        let (x1, y1, x2, y2) = map_child_rect(extreme, &rotated, &child);
        assert!(x1 <= x2 && y1 <= y2);
    }

    #[test]
    fn map_child_rect_applies_group_flips_about_the_outer_frame_center() {
        let outer: CornersEmu = (0, 0, 1000, 2000);
        let child = ChildTransform {
            x_emu: 0,
            y_emu: 0,
            cx_emu: 50,
            cy_emu: 100,
            ..ChildTransform::default()
        };
        let base = GroupTransform {
            child_x_emu: 0,
            child_y_emu: 0,
            child_cx_emu: 100,
            child_cy_emu: 200,
            ..GroupTransform::default()
        };

        let flipped_h = GroupTransform {
            flip_h: true,
            ..base.clone()
        };
        assert_eq!(
            map_child_rect(outer, &flipped_h, &child),
            (500, 0, 1000, 1000),
            "flipH mirrors across the vertical center line"
        );

        let flipped_v = GroupTransform {
            flip_v: true,
            ..base
        };
        assert_eq!(
            map_child_rect(outer, &flipped_v, &child),
            (0, 1000, 500, 2000),
            "flipV mirrors across the horizontal center line"
        );
    }
}
