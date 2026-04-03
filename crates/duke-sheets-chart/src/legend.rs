//! Legend types

use crate::text_properties::TextProperties;
/// Chart legend
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Legend {
    /// Position
    pub position: LegendPosition,
    /// Whether legend overlays the chart
    pub overlay: bool,
    /// Shape properties (fill, line)
    pub shape_properties: Option<crate::formatting::ChartShapeProperties>,
    pub text_properties: Option<TextProperties>,
    pub entries: Vec<LegendEntry>,
}

impl Legend {
    /// Create a new legend
    pub fn new(position: LegendPosition) -> Self {
        Self {
            position,
            overlay: false,
            shape_properties: None,
            text_properties: None,
            entries: Vec::new(),
        }
    }
}

/// Legend position
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LegendPosition {
    #[default]
    Right,
    Top,
    Bottom,
    Left,
    TopRight,
}

/// Per-entry override (e.g. hiding a specific legend entry).
#[derive(Debug, Clone, PartialEq)]
pub struct LegendEntry {
    pub index: u32,
    pub delete: Option<bool>,
    pub text_properties: Option<TextProperties>,
}
