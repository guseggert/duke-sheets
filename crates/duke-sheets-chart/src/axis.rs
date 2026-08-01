//! Axis types

use crate::formatting::{ChartShapeProperties, NumberFormat};
use crate::text_properties::TextProperties;

/// Distinguishes the XML axis element used for serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AxisType {
    #[default]
    Category,
    Value,
    Date,
    Series,
}

/// Chart axis
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Axis {
    /// Axis title
    pub title: Option<String>,
    /// Which XML element to use when writing (catAx, valAx, dateAx, serAx).
    pub axis_type: AxisType,
    /// Minimum value
    pub minimum: Option<f64>,
    /// Maximum value
    pub maximum: Option<f64>,
    /// Major unit
    pub major_unit: Option<f64>,
    /// Minor unit
    pub minor_unit: Option<f64>,
    /// Where the axis sits (`c:axPos`). `None` when the source omitted
    /// it, which the writer fills with the position conventional for the
    /// axis's role. Distinguishing that from an explicit `Bottom` is what
    /// keeps a bar chart's value axis from migrating to the left.
    pub position: Option<AxisPosition>,
    pub number_format: Option<NumberFormat>,
    pub major_gridlines: bool,
    pub minor_gridlines: bool,
    pub major_gridlines_shape_properties: Option<ChartShapeProperties>,
    pub minor_gridlines_shape_properties: Option<ChartShapeProperties>,
    pub major_tick_mark: Option<TickMark>,
    pub minor_tick_mark: Option<TickMark>,
    pub label_position: Option<TickLabelPosition>,
    /// When true, the axis is hidden
    pub delete: Option<bool>,
    pub crosses: Option<AxisCrosses>,
    pub cross_between: Option<CrossBetween>,
    pub shape_properties: Option<ChartShapeProperties>,
    /// Category axis label offset (0-1000)
    pub label_offset: Option<u32>,
    pub auto_labeled: Option<bool>,
    pub text_properties: Option<TextProperties>,
    /// Raw extLst XML to preserve on roundtrip.
    #[doc(hidden)]
    pub raw_ext: Option<Vec<u8>>,
}

impl Axis {
    /// Create a new axis
    pub fn new() -> Self {
        Self {
            title: None,
            axis_type: AxisType::default(),
            minimum: None,
            maximum: None,
            major_unit: None,
            minor_unit: None,
            position: None,
            number_format: None,
            major_gridlines: false,
            minor_gridlines: false,
            major_gridlines_shape_properties: None,
            minor_gridlines_shape_properties: None,
            major_tick_mark: None,
            minor_tick_mark: None,
            label_position: None,
            delete: None,
            crosses: None,
            cross_between: None,
            shape_properties: None,
            raw_ext: None,
            label_offset: None,
            auto_labeled: None,
            text_properties: None,
        }
    }

    /// Set axis title
    pub fn with_title<S: Into<String>>(mut self, title: S) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Set axis bounds
    pub fn with_bounds(mut self, min: f64, max: f64) -> Self {
        self.minimum = Some(min);
        self.maximum = Some(max);
        self
    }
}

/// Axis position
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AxisPosition {
    #[default]
    Bottom,
    Top,
    Left,
    Right,
}

/// Tick mark style
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TickMark {
    Cross,
    Inside,
    None,
    Outside,
}

/// Position of tick labels relative to the axis
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TickLabelPosition {
    High,
    Low,
    NextTo,
    None,
}

/// Where the perpendicular axis crosses
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisCrosses {
    AutoZero,
    Min,
    Max,
}

/// Whether categories are crossed between or at midpoint
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CrossBetween {
    Between,
    MidCat,
}
