//! Axis types

/// Chart axis
#[derive(Debug, Clone, Default)]
pub struct Axis {
    /// Axis title
    pub title: Option<String>,
    /// Minimum value
    pub minimum: Option<f64>,
    /// Maximum value
    pub maximum: Option<f64>,
    /// Major unit
    pub major_unit: Option<f64>,
    /// Minor unit
    pub minor_unit: Option<f64>,
    /// Position
    pub position: AxisPosition,
}

impl Axis {
    /// Create a new axis
    pub fn new() -> Self {
        Self::default()
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
