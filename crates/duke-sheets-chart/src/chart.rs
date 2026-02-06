//! Chart types

use crate::axis::Axis;
use crate::legend::Legend;
use crate::series::DataSeries;

/// Chart types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartType {
    // Column/Bar
    ColumnClustered,
    ColumnStacked,
    ColumnPercentStacked,
    BarClustered,
    BarStacked,
    BarPercentStacked,

    // Line
    Line,
    LineStacked,
    LineMarkers,

    // Pie
    Pie,
    PieExploded,
    Doughnut,

    // Area
    Area,
    AreaStacked,
    AreaPercentStacked,

    // Scatter
    ScatterMarkers,
    ScatterSmooth,
    ScatterLines,

    // Other
    Bubble,
    Radar,
    Stock,
    Surface,
}

/// Chart definition
#[derive(Debug, Clone)]
pub struct Chart {
    /// Chart type
    pub chart_type: ChartType,
    /// Chart title
    pub title: Option<String>,
    /// Data series
    pub series: Vec<DataSeries>,
    /// Category axis (X)
    pub category_axis: Option<Axis>,
    /// Value axis (Y)
    pub value_axis: Option<Axis>,
    /// Legend
    pub legend: Option<Legend>,
    /// Position anchor
    pub anchor: ChartAnchor,
}

impl Chart {
    /// Create a new chart
    pub fn new(chart_type: ChartType) -> Self {
        Self {
            chart_type,
            title: None,
            series: Vec::new(),
            category_axis: None,
            value_axis: None,
            legend: None,
            anchor: ChartAnchor::default(),
        }
    }

    /// Set chart title
    pub fn with_title<S: Into<String>>(mut self, title: S) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Add a data series
    pub fn add_series(&mut self, series: DataSeries) {
        self.series.push(series);
    }
}

/// Chart anchor position
#[derive(Debug, Clone, Default)]
pub struct ChartAnchor {
    /// Start column
    pub from_col: u16,
    /// Start row
    pub from_row: u32,
    /// End column
    pub to_col: u16,
    /// End row
    pub to_row: u32,
}
