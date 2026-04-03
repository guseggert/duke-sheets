//! Marker types for data points on line/scatter charts

/// Marker symbol displayed at data points
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Marker {
    pub symbol: Option<MarkerSymbol>,
    /// Size in points (2–72)
    pub size: Option<u8>,
}

/// Marker symbol shape
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkerSymbol {
    Circle,
    Dash,
    Diamond,
    Dot,
    None,
    Picture,
    Plus,
    Square,
    Star,
    Triangle,
    X,
    Auto,
}
