//! Chart-level configuration types (3D view, data table, layout)

/// 3D view rotation and perspective settings
#[derive(Debug, Clone, Default, PartialEq)]
pub struct View3D {
    pub rotate_x: Option<i32>,
    pub rotate_y: Option<i32>,
    pub depth_percent: Option<u32>,
    pub height_percent: Option<u32>,
    pub perspective: Option<u32>,
    pub right_angle_axes: Option<bool>,
}

/// Data table displayed beneath the chart
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartDataTable {
    pub show_horizontal_border: Option<bool>,
    pub show_vertical_border: Option<bool>,
    pub show_outline: Option<bool>,
    pub show_keys: Option<bool>,
}

/// How blank cells are rendered in the chart
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DisplayBlanksAs {
    Gap,
    Span,
    Zero,
}

/// Layout container
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Layout {
    pub manual_layout: Option<ManualLayout>,
}

/// Manual positioning (fractional coordinates 0.0–1.0)
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ManualLayout {
    pub x: Option<f64>,
    pub y: Option<f64>,
    pub width: Option<f64>,
    pub height: Option<f64>,
}
