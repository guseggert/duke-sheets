//! Chart formatting types (fill, line, number format)

/// Number format for axis labels, data labels, etc.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberFormat {
    pub format_code: String,
    pub source_linked: Option<bool>,
}

impl Default for NumberFormat {
    fn default() -> Self {
        Self {
            format_code: "General".to_string(),
            source_linked: None,
        }
    }
}

/// Simplified shape properties for charts (fill + line)
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartShapeProperties {
    pub solid_fill: Option<ChartColor>,
    pub no_fill: bool,
    pub line: Option<ChartLine>,
}

/// A color specified as a hex RGB string
#[derive(Debug, Clone, PartialEq)]
pub struct ChartColor {
    /// Hex RGB value, e.g. "FF0000"
    pub hex: String,
}

/// Line/outline properties
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartLine {
    /// Width in EMU
    pub width: Option<i64>,
    pub solid_fill: Option<ChartColor>,
    pub no_fill: bool,
    /// Dash style: "solid", "dash", "dot", etc.
    pub dash_style: Option<String>,
}

/// Picture fill options for data series/points
#[derive(Debug, Clone, PartialEq, Default)]
pub struct PictureOptions {
    pub apply_to_front: Option<bool>,
    pub apply_to_sides: Option<bool>,
    pub apply_to_end: Option<bool>,
}
