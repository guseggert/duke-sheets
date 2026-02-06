//! Data series types

/// Data series for a chart
#[derive(Debug, Clone)]
pub struct DataSeries {
    /// Series name
    pub name: Option<String>,
    /// Values (Y data)
    pub values: DataReference,
    /// Categories (X data)
    pub categories: Option<DataReference>,
}

impl DataSeries {
    /// Create a new data series
    pub fn new(values: DataReference) -> Self {
        Self {
            name: None,
            values,
            categories: None,
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
#[derive(Debug, Clone)]
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
