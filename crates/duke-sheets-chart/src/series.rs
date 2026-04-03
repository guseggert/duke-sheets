//! Data series types

use crate::data_labels::{DataLabels, DataPoint};
use crate::error_bars::ErrorBars;
use crate::formatting::ChartShapeProperties;
use crate::marker::Marker;
use crate::trendline::Trendline;

/// Data series for a chart
#[derive(Debug, Clone, PartialEq)]
pub struct DataSeries {
    /// Series name
    pub name: Option<String>,
    /// Values (Y data)
    pub values: DataReference,
    /// Categories (X data)
    pub categories: Option<DataReference>,
    pub data_labels: Option<DataLabels>,
    pub trendline: Option<Trendline>,
    pub error_bars: Option<ErrorBars>,
    pub marker: Option<Marker>,
    pub data_points: Vec<DataPoint>,
    pub shape_properties: Option<ChartShapeProperties>,
    /// Smooth line (for line/scatter charts)
    pub smooth: Option<bool>,
    /// Pie explosion percent
    pub explosion: Option<u32>,
    pub invert_if_negative: Option<bool>,
    pub bubble_3d: Option<bool>,
    /// Raw extLst XML to preserve on roundtrip.
    #[doc(hidden)]
    pub raw_ext: Option<Vec<u8>>,
}

impl DataSeries {
    /// Create a new data series
    pub fn new(values: DataReference) -> Self {
        Self {
            name: None,
            values,
            categories: None,
            data_labels: None,
            trendline: None,
            error_bars: None,
            marker: None,
            data_points: Vec::new(),
            shape_properties: None,
            smooth: None,
            explosion: None,
            raw_ext: None,
            invert_if_negative: None,
            bubble_3d: None,
        }
    }

    /// Set series name
    pub fn with_name<S: Into<String>>(mut self, name: S) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set categories
    pub fn with_categories(mut self, categories: DataReference) -> Self {
        self.categories = Some(categories);
        self
    }
}

/// Reference to chart data
#[derive(Debug, Clone, PartialEq)]
pub enum DataReference {
    /// Formula reference (e.g., "Sheet1!$A$1:$A$10")
    Formula(String),
    /// Literal numeric values
    Numbers(Vec<f64>),
    /// Literal string values (for categories)
    Strings(Vec<String>),
}

impl DataReference {
    /// Create a formula reference
    pub fn formula<S: Into<String>>(formula: S) -> Self {
        DataReference::Formula(formula.into())
    }

    /// Create from numeric values
    pub fn numbers(values: Vec<f64>) -> Self {
        DataReference::Numbers(values)
    }

    /// Create from string values
    pub fn strings(values: Vec<String>) -> Self {
        DataReference::Strings(values)
    }
}
