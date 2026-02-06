//! Legend types

/// Chart legend
#[derive(Debug, Clone, Default)]
pub struct Legend {
    /// Position
    pub position: LegendPosition,
    /// Whether legend overlays the chart
    pub overlay: bool,
}

impl Legend {
    /// Create a new legend
    pub fn new(position: LegendPosition) -> Self {
        Self {
            position,
            overlay: false,
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
