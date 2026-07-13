//! Drawing placement primitives shared by the worksheet drawing
//! model (duke-sheets-core) and the OOXML drawing-part codec
//! ([`crate::drawing_part`]).

use crate::CellMarker;

/// English Metric Units per 96-dpi screen pixel.
pub const EMU_PER_PIXEL: i64 = 9_525;
/// English Metric Units per point.
pub const EMU_PER_POINT: i64 = 12_700;

/// Worksheet row and column metrics used to resolve drawing anchors.
///
/// Implementations should return the visible cell extent in EMU. The
/// position methods have general summing defaults and can be overridden
/// when the backing model can compute prefix sums more efficiently.
pub trait DrawingMetrics {
    /// Width of a zero-based worksheet column in EMU.
    fn column_width_emu(&self, col: u16) -> i64;

    /// Height of a zero-based worksheet row in EMU.
    fn row_height_emu(&self, row: u32) -> i64;

    /// Absolute left edge of a zero-based worksheet column in EMU.
    fn column_position_emu(&self, col: u16) -> i128 {
        (0..col)
            .map(|index| i128::from(self.column_width_emu(index).max(0)))
            .sum()
    }

    /// Absolute top edge of a zero-based worksheet row in EMU.
    fn row_position_emu(&self, row: u32) -> i128 {
        (0..row)
            .map(|index| i128::from(self.row_height_emu(index).max(0)))
            .sum()
    }
}

/// Excel-compatible conversion from a stored column width in characters
/// to EMU.
///
/// OOXML defines `col@width` in character units (ECMA-376 Part 1,
/// section 18.3.1.13), but a worksheet does not retain the font-specific
/// maximum digit width needed for the full display calculation. This uses
/// the conventional Calibri-11 compatibility approximation used for Excel
/// workbooks: widths below 1 use `floor(width * 12 + 0.5)` pixels and other
/// widths use `floor(width * 7 + 5)` pixels, at 96 dpi. Thus the default
/// width 8.43 is 64 px, or 609,600 EMU.
pub fn column_width_to_emu(width: f64) -> i64 {
    if !width.is_finite() || width <= 0.0 {
        return 0;
    }
    let pixels = if width < 1.0 {
        (width * 12.0 + 0.5).floor()
    } else {
        (width * 7.0 + 5.0).floor()
    };
    let pixels = pixels.clamp(0.0, (i64::MAX / EMU_PER_PIXEL) as f64) as i64;
    pixels.saturating_mul(EMU_PER_PIXEL)
}

/// Convert a row height in points to EMU (12,700 EMU per point).
pub fn row_height_to_emu(points: f64) -> i64 {
    if !points.is_finite() || points <= 0.0 {
        return 0;
    }
    (points * EMU_PER_POINT as f64)
        .round()
        .clamp(0.0, i64::MAX as f64) as i64
}

/// Excel's default worksheet metrics: 8.43-character columns and
/// 15-point rows.
#[derive(Debug, Clone, Copy, Default)]
pub struct DefaultDrawingMetrics;

impl DrawingMetrics for DefaultDrawingMetrics {
    fn column_width_emu(&self, _col: u16) -> i64 {
        609_600
    }

    fn row_height_emu(&self, _row: u32) -> i64 {
        190_500
    }

    fn column_position_emu(&self, col: u16) -> i128 {
        i128::from(col) * 609_600
    }

    fn row_position_emu(&self, row: u32) -> i128 {
        i128::from(row) * 190_500
    }
}

/// Absolute worksheet position of a cell marker under `metrics`.
#[doc(hidden)]
pub fn marker_position_emu(
    marker: &CellMarker,
    metrics: &(impl DrawingMetrics + ?Sized),
) -> (i128, i128) {
    (
        metrics.column_position_emu(marker.col) + i128::from(marker.col_offset_emu),
        metrics.row_position_emu(marker.row) + i128::from(marker.row_offset_emu),
    )
}

fn column_at_emu(x: i128, metrics: &(impl DrawingMetrics + ?Sized)) -> (u16, i64) {
    // Off-grid positions clamp to the grid origin.
    if x <= 0 {
        return (0, 0);
    }
    let mut low = 0u32;
    let mut high = u32::from(u16::MAX) + 1;
    while low < high {
        let mid = low + (high - low) / 2;
        let position = metrics.column_position_emu(mid as u16);
        if position <= x {
            low = mid + 1;
        } else {
            high = mid;
        }
    }
    let mut col = low.saturating_sub(1).min(u32::from(u16::MAX)) as u16;
    // Among zero-width columns sharing this position (degenerate
    // metrics), resolve to the first of the run.
    let start = metrics.column_position_emu(col);
    let mut first = 0u32;
    let mut last = u32::from(col);
    while first < last {
        let mid = first + (last - first) / 2;
        if metrics.column_position_emu(mid as u16) == start {
            last = mid;
        } else {
            first = mid + 1;
        }
    }
    col = first as u16;
    let width = i128::from(metrics.column_width_emu(col).max(0));
    let max_offset = width.saturating_sub(1).max(0);
    let offset = (x - start).clamp(0, max_offset).min(i128::from(i64::MAX)) as i64;
    (col, offset)
}

fn row_at_emu(y: i128, metrics: &(impl DrawingMetrics + ?Sized)) -> (u32, i64) {
    // Off-grid positions clamp to the grid origin.
    if y <= 0 {
        return (0, 0);
    }
    let mut low = 0u64;
    let mut high = u64::from(u32::MAX) + 1;
    while low < high {
        let mid = low + (high - low) / 2;
        let position = metrics.row_position_emu(mid as u32);
        if position <= y {
            low = mid + 1;
        } else {
            high = mid;
        }
    }
    let mut row = low.saturating_sub(1).min(u64::from(u32::MAX)) as u32;
    // Among zero-height rows sharing this position (degenerate
    // metrics), resolve to the first of the run.
    let start = metrics.row_position_emu(row);
    let mut first = 0u64;
    let mut last = u64::from(row);
    while first < last {
        let mid = first + (last - first) / 2;
        if metrics.row_position_emu(mid as u32) == start {
            last = mid;
        } else {
            first = mid + 1;
        }
    }
    row = first as u32;
    let height = i128::from(metrics.row_height_emu(row).max(0));
    let max_offset = height.saturating_sub(1).max(0);
    let offset = (y - start).clamp(0, max_offset).min(i128::from(i64::MAX)) as i64;
    (row, offset)
}

/// Resolve an absolute worksheet position to a cell marker.
#[doc(hidden)]
pub fn marker_at_emu(
    x: i128,
    y: i128,
    metrics: &(impl DrawingMetrics + ?Sized),
) -> CellMarker {
    let (col, col_offset_emu) = column_at_emu(x, metrics);
    let (row, row_offset_emu) = row_at_emu(y, metrics);
    CellMarker {
        col,
        col_offset_emu,
        row,
        row_offset_emu,
    }
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
