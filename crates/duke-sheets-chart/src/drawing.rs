//! Drawing placement primitives shared by the worksheet drawing
//! model (duke-sheets-core) and the OOXML drawing-part codec
//! ([`crate::drawing_part`]).

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
