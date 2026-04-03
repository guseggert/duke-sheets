//! Error bar types

use crate::series::DataReference;
/// Error bars attached to a data series
#[derive(Debug, Clone, PartialEq)]
pub struct ErrorBars {
    pub direction: ErrorBarDirection,
    pub bar_type: ErrorBarType,
    pub value_type: ErrorValueType,
    pub value: Option<f64>,
    pub no_end_cap: Option<bool>,
    /// Custom positive error values (when value_type is Custom)
    pub plus: Option<DataReference>,
    /// Custom negative error values (when value_type is Custom)
    pub minus: Option<DataReference>,
}

/// Which axis the error bars apply to
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorBarDirection {
    X,
    Y,
}

/// Whether error bars extend in both directions or only one
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorBarType {
    Both,
    Minus,
    Plus,
}

/// How the error bar magnitude is determined
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorValueType {
    Custom,
    FixedValue,
    Percentage,
    StandardDeviation,
    StandardError,
}
