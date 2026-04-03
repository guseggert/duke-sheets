//! Trendline types

use crate::config::Layout;
use crate::formatting::{ChartShapeProperties, NumberFormat};
use crate::text_properties::TextProperties;
/// A trendline attached to a data series
#[derive(Debug, Clone, PartialEq)]
pub struct Trendline {
    pub trendline_type: TrendlineType,
    pub name: Option<String>,
    /// Polynomial order (when type is Polynomial)
    pub order: Option<u32>,
    /// Moving average period (when type is MovingAverage)
    pub period: Option<u32>,
    /// Forecast forward periods
    pub forward: Option<f64>,
    /// Forecast backward periods
    pub backward: Option<f64>,
    pub intercept: Option<f64>,
    pub display_r_squared: Option<bool>,
    pub display_equation: Option<bool>,
    pub label: Option<TrendlineLabel>,
}

/// Type of trendline regression
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrendlineType {
    Linear,
    Exponential,
    Logarithmic,
    MovingAverage,
    Polynomial,
    Power,
}

/// Label displayed on a trendline
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TrendlineLabel {
    pub layout: Option<Layout>,
    pub text: Option<String>,
    pub number_format: Option<NumberFormat>,
    pub shape_properties: Option<ChartShapeProperties>,
    pub text_properties: Option<TextProperties>,
}
