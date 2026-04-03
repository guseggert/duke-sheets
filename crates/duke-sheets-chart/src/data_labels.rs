//! Data label types

use crate::chart::ChartLines;
use crate::config::Layout;
use crate::formatting::{ChartShapeProperties, NumberFormat};
use crate::marker::Marker;
use crate::text_properties::TextProperties;

/// Data labels configuration for a chart or series
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DataLabels {
    pub show_legend_key: Option<bool>,
    pub show_value: Option<bool>,
    pub show_category_name: Option<bool>,
    pub show_series_name: Option<bool>,
    pub show_percent: Option<bool>,
    pub show_bubble_size: Option<bool>,
    pub separator: Option<String>,
    pub position: Option<DataLabelPosition>,
    pub number_format: Option<NumberFormat>,
    pub show_leader_lines: Option<bool>,
    pub leader_lines: Option<ChartLines>,
    pub text_properties: Option<TextProperties>,
    pub data_label_overrides: Vec<DataLabel>,
}

/// Position of a data label relative to its data point
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataLabelPosition {
    BestFit,
    Bottom,
    Center,
    InsideBase,
    InsideEnd,
    Left,
    OutsideEnd,
    Right,
    Top,
}

/// Override formatting/marker for an individual data point
#[derive(Debug, Clone, PartialEq)]
pub struct DataPoint {
    pub index: u32,
    pub marker: Option<Marker>,
    /// Pie explosion percent
    pub explosion: Option<u32>,
}

/// Per-point data label override
#[derive(Debug, Clone, PartialEq, Default)]
pub struct DataLabel {
    pub index: u32,
    pub layout: Option<Layout>,
    /// Override text
    pub text: Option<String>,
    pub number_format: Option<NumberFormat>,
    pub shape_properties: Option<ChartShapeProperties>,
    pub show_legend_key: Option<bool>,
    pub show_value: Option<bool>,
    pub show_category_name: Option<bool>,
    pub show_series_name: Option<bool>,
    pub show_percent: Option<bool>,
    pub show_bubble_size: Option<bool>,
    pub separator: Option<String>,
    pub position: Option<DataLabelPosition>,
}
