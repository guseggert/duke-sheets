//! Chart types

use std::collections::HashMap;


use crate::axis::Axis;
use crate::config::{ChartDataTable, DisplayBlanksAs, Layout, View3D};
use crate::data_labels::DataLabels;
use crate::formatting::ChartShapeProperties;
use crate::legend::Legend;
use crate::series::DataSeries;
use crate::text_properties::TextProperties;


/// Chart types
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartType {
    // Column/Bar
    ColumnClustered,
    ColumnStacked,
    ColumnPercentStacked,
    BarClustered,
    BarStacked,
    BarPercentStacked,

    // Line
    Line,
    LineStacked,

    // Pie
    Pie,
    PieExploded,
    Doughnut,

    // Area
    Area,
    AreaStacked,
    AreaPercentStacked,

    // Scatter
    ScatterMarkers,
    ScatterSmooth,
    ScatterLines,

    // Other
    Bubble,
    Radar,
    Stock,
    Surface,

    /// Imported chart type not yet mapped to a known variant.
    /// Contains the original OOXML element tag (e.g. "c:surface3DChart").
    Unsupported(String),
}

/// Chart line overlay (drop lines, high-low lines, series lines, leader lines).
/// All share the same structure: optional shape properties for formatting.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartLines {
    pub shape_properties: Option<ChartShapeProperties>,
}

/// Up-down bars (stock charts).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct UpDownBars {
    pub gap_width: Option<u32>,
    pub up_bars: Option<ChartLines>,
    pub down_bars: Option<ChartLines>,
}

/// Represents one chart type block (e.g. barChart, lineChart) within a plotArea.
/// Used for combo charts where multiple chart types share axes.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartTypeGroup {
    pub chart_type: ChartType,
    pub is_3d: bool,
    pub series: Vec<DataSeries>,
    pub data_labels: Option<DataLabels>,
    pub vary_colors: Option<bool>,
    pub gap_width: Option<u32>,
    pub overlap: Option<i32>,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
    pub bubble_scale: Option<u32>,
    pub show_negative_bubbles: Option<bool>,
    pub radar_style: Option<String>,
    pub wireframe: Option<bool>,
    pub drop_lines: Option<ChartLines>,
    pub high_low_lines: Option<ChartLines>,
    pub series_lines: Option<ChartLines>,
    pub up_down_bars: Option<UpDownBars>,
    /// axId values this group references (e.g. [1, 2] or [1, 3])
    pub of_pie_type: Option<OfPieType>,
    pub split_type: Option<SplitType>,
    pub split_pos: Option<f64>,
    pub second_pie_size: Option<u32>,
    pub bar_shape: Option<BarShape>,
    pub floor: Option<Surface>,
    pub side_wall: Option<Surface>,
    pub back_wall: Option<Surface>,
    pub axis_ids: Vec<u32>,
    #[doc(hidden)]
    pub raw_ext: Option<Vec<u8>>,
}

/// Represents one axis in the plotArea with its ID and cross-reference.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartAxis {
    pub id: u32,
    pub cross_id: u32,
    pub axis: Axis,
}

/// Chart definition
#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    /// Chart type
    pub chart_type: ChartType,
    /// Chart title
    pub title: Option<String>,
    /// Data series
    pub series: Vec<DataSeries>,
    /// Category axis (X)
    pub category_axis: Option<Axis>,
    /// Value axis (Y)
    pub value_axis: Option<Axis>,
    /// Series axis (Z, for 3D charts)
    pub series_axis: Option<Axis>,
    /// Legend
    pub legend: Option<Legend>,
    /// Position anchor
    pub anchor: ChartAnchor,
    pub data_labels: Option<DataLabels>,
    pub view_3d: Option<View3D>,
    pub data_table: Option<ChartDataTable>,
    pub display_blanks_as: Option<DisplayBlanksAs>,
    pub plot_visible_only: Option<bool>,
    pub layout: Option<Layout>,
    pub shape_properties: Option<ChartShapeProperties>,
    pub vary_colors: Option<bool>,
    pub is_3d: bool,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
    pub bubble_scale: Option<u32>,
    pub show_negative_bubbles: Option<bool>,
    pub radar_style: Option<String>,
    pub auto_title_deleted: Option<bool>,
    pub rounded_corners: Option<bool>,
    pub show_dlbls_over_max: Option<bool>,
    pub wireframe: Option<bool>,
    pub drop_lines: Option<ChartLines>,
    pub high_low_lines: Option<ChartLines>,
    pub up_down_bars: Option<UpDownBars>,
    pub series_lines: Option<ChartLines>,
    pub gap_width: Option<u32>,
    pub overlap: Option<i32>,
    pub show_marker: Option<bool>,
    pub of_pie_type: Option<OfPieType>,
    pub split_type: Option<SplitType>,
    pub split_pos: Option<f64>,
    pub second_pie_size: Option<u32>,
    pub bar_shape: Option<BarShape>,
    pub floor: Option<Surface>,
    pub side_wall: Option<Surface>,
    pub back_wall: Option<Surface>,
    pub text_properties: Option<TextProperties>,
    /// Raw XML fragments for extension lists we don't parse but need to preserve.
    /// Key is the context path (e.g. "chartSpace", "chart", "plotArea", "chartType").
    #[doc(hidden)]
    pub raw_extensions: HashMap<String, Vec<u8>>,
    /// Raw chart style XML (Office 2013+), preserved for roundtrip.
    #[doc(hidden)]
    pub raw_chart_style: Option<Vec<u8>>,
    /// Raw chart color style XML (Office 2013+), preserved for roundtrip.
    #[doc(hidden)]
    pub raw_chart_color_style: Option<Vec<u8>>,
    /// Chart type groups for combo charts (multiple chart types in one plotArea).
    /// When non-empty, the writer uses these instead of the legacy single-group fields.
    pub type_groups: Vec<ChartTypeGroup>,
    /// Axes for combo charts, each with its ID and cross-reference.
    pub axes: Vec<ChartAxis>,
}

impl Chart {
    /// Create a new chart
    pub fn new(chart_type: ChartType) -> Self {
        Self {
            chart_type,
            title: None,
            series: Vec::new(),
            category_axis: None,
            value_axis: None,
            series_axis: None,
            legend: None,
            anchor: ChartAnchor::default(),
            data_labels: None,
            view_3d: None,
            data_table: None,
            display_blanks_as: None,
            plot_visible_only: None,
            layout: None,
            shape_properties: None,
            vary_colors: None,
            gap_width: None,
            overlap: None,
            is_3d: false,
            first_slice_angle: None,
            hole_size: None,
            bubble_scale: None,
            show_negative_bubbles: None,
            radar_style: None,
            auto_title_deleted: None,
            rounded_corners: None,
            show_dlbls_over_max: None,
            wireframe: None,
            drop_lines: None,
            high_low_lines: None,
            up_down_bars: None,
            series_lines: None,
            show_marker: None,
            of_pie_type: None,
            split_type: None,
            split_pos: None,
            second_pie_size: None,
            bar_shape: None,
            floor: None,
            side_wall: None,
            back_wall: None,
            text_properties: None,
            raw_extensions: HashMap::new(),
            raw_chart_style: None,
            raw_chart_color_style: None,
            type_groups: Vec::new(),
            axes: Vec::new(),
        }
    }


    /// Set chart title
    pub fn with_title<S: Into<String>>(mut self, title: S) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Add a data series
    pub fn add_series(&mut self, series: DataSeries) {
        self.series.push(series);
    }
}

/// Chart anchor position in a worksheet.
///
/// Cell positions use zero-based column and row indices.
/// Offsets are in EMU (English Metric Units, 1 inch = 914400 EMU)
/// and represent sub-cell positioning within the from/to cells.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartAnchor {
    /// Start column (zero-based)
    pub from_col: u16,
    /// Start row (zero-based)
    pub from_row: u32,
    /// Offset within start column (EMU)
    pub from_col_offset: i64,
    /// Offset within start row (EMU)
    pub from_row_offset: i64,
    /// End column (zero-based)
    pub to_col: u16,
    /// End row (zero-based)
    pub to_row: u32,
    /// Offset within end column (EMU)
    pub to_col_offset: i64,
    /// Offset within end row (EMU)
    pub to_row_offset: i64,
}

/// 3D chart surface (floor, side wall, back wall)
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Surface {
    pub thickness: Option<u32>,
    pub shape_properties: Option<ChartShapeProperties>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OfPieType {
    Pie,
    Bar,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SplitType {
    Auto,
    Custom,
    Percent,
    Position,
    Value,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarShape {
    Box,
    Cone,
    ConeToMax,
    Cylinder,
    Pyramid,
    PyramidToMax,
}
