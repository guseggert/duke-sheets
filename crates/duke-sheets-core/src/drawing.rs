//! Unified drawing model
//!
//! Every floating object on a worksheet - images, charts, form
//! controls, cell comments, shape groups, and preserved raw drawing
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

use crate::comment::CellComment;
use crate::form_control::FormControl;
use crate::{Error, Result, MAX_COLS, MAX_ROWS};

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
    /// element, or an XLSB drawing bundle during migration).
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

/// Transform of a shape group (DrawingML `a:xfrm` with `chOff`/`chExt`).
///
/// `x/y/cx/cy` place the group in its parent's coordinate space (for
/// a top-level group this duplicates the anchor's extent in absolute
/// EMU; both are preserved on round-trip). `child_*` define the
/// coordinate space the children are expressed in.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct GroupTransform {
    /// Group offset in parent space (EMU).
    pub x_emu: i64,
    /// Group offset in parent space (EMU).
    pub y_emu: i64,
    /// Group extent in parent space (EMU).
    pub cx_emu: i64,
    /// Group extent in parent space (EMU).
    pub cy_emu: i64,
    /// Child-space origin (EMU).
    pub child_x_emu: i64,
    /// Child-space origin (EMU).
    pub child_y_emu: i64,
    /// Child-space extent (EMU).
    pub child_cx_emu: i64,
    /// Child-space extent (EMU).
    pub child_cy_emu: i64,
    /// Rotation in 60,000ths of a degree, clockwise.
    pub rotation: i32,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
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

/// Placement of a group child in its group's child coordinate space.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChildTransform {
    /// Offset in child space (EMU).
    pub x_emu: i64,
    /// Offset in child space (EMU).
    pub y_emu: i64,
    /// Extent in child space (EMU).
    pub cx_emu: i64,
    /// Extent in child space (EMU).
    pub cy_emu: i64,
    /// Rotation in 60,000ths of a degree, clockwise.
    pub rotation: i32,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
}

/// Path to a drawing node: the first element indexes the worksheet's
/// drawing list, subsequent elements index group children.
pub type DrawingPath = Vec<usize>;

/// Absolute rectangle in EMU at Excel's default cell metrics:
/// `(x1, y1, x2, y2)`.
pub type RectEmu = (i128, i128, i128, i128);

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
                    return Err(Error::other(
                        "group child extents cannot be negative",
                    ));
                }
                validate_kind(&child.kind)?;
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
                return Err(Error::other(
                    "drawing anchor dimensions cannot be negative",
                ));
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

/// Absolute EMU rectangle (x1, y1, x2, y2) for a drawing anchor at
/// Excel's default cell metrics (609,600 EMU per column, 190,500 EMU
/// per row).
pub(crate) fn anchor_rect_emu(anchor: &DrawingAnchor) -> RectEmu {
    const COL_EMU: i128 = 609_600;
    const ROW_EMU: i128 = 190_500;
    match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => (
            from.col as i128 * COL_EMU + from.col_offset_emu as i128,
            from.row as i128 * ROW_EMU + from.row_offset_emu as i128,
            to.col as i128 * COL_EMU + to.col_offset_emu as i128,
            to.row as i128 * ROW_EMU + to.row_offset_emu as i128,
        ),
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            let x1 = from.col as i128 * COL_EMU + from.col_offset_emu as i128;
            let y1 = from.row as i128 * ROW_EMU + from.row_offset_emu as i128;
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
pub(crate) fn map_child_rect(
    outer: RectEmu,
    transform: &GroupTransform,
    child: &ChildTransform,
) -> RectEmu {
    let (ox1, oy1, ox2, oy2) = outer;
    let (ow, oh) = ((ox2 - ox1).max(0), (oy2 - oy1).max(0));
    let ch_w = i128::from(transform.child_cx_emu).max(1);
    let ch_h = i128::from(transform.child_cy_emu).max(1);
    let map_x = |x: i128| ox1 + (x - i128::from(transform.child_x_emu)) * ow / ch_w;
    let map_y = |y: i128| oy1 + (y - i128::from(transform.child_y_emu)) * oh / ch_h;
    let x1 = i128::from(child.x_emu);
    let y1 = i128::from(child.y_emu);
    let x2 = x1 + i128::from(child.cx_emu).max(0);
    let y2 = y1 + i128::from(child.cy_emu).max(0);
    (map_x(x1), map_y(y1), map_x(x2), map_y(y2))
}

/// Excel's default comment popup placement for a cell: one column to
/// the right, from just above the cell's row to three rows below.
pub fn default_comment_anchor(row: u32, col: u16) -> DrawingAnchor {
    const PX_EMU: i64 = 9_525;
    let clamp_col = |c: u32| -> u16 { c.min(u32::from(MAX_COLS - 1)) as u16 };
    let clamp_row = |r: u32| -> u32 { r.min(MAX_ROWS - 1) };
    DrawingAnchor::TwoCell {
        from: duke_sheets_chart::CellMarker {
            col: clamp_col(u32::from(col) + 1),
            col_offset_emu: 15 * PX_EMU,
            row: row.saturating_sub(1),
            row_offset_emu: 10 * PX_EMU,
        },
        to: duke_sheets_chart::CellMarker {
            col: clamp_col(u32::from(col) + 3),
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
            caption: "Run".to_string(),
        }));
        assert!(!object.meta.hidden);
        assert!(object.meta.locked);
    }

    #[test]
    fn validate_rejects_reversed_anchor() {
        let object = DrawingObject::form_control(FormControl::new(FormControlKind::Button {
            caption: "b".to_string(),
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
        assert!(object
            .validate()
            .unwrap_err()
            .to_string()
            .contains("lines"));
    }

    #[test]
    fn map_child_rect_scales_into_outer_frame() {
        let outer: RectEmu = (1000, 2000, 3000, 4000);
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
}
