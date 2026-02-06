//! Column types

/// Column metadata
#[derive(Debug, Clone)]
pub struct Column {
    /// Column index (0-based)
    pub index: u16,
    /// Custom width (None = default)
    pub width: Option<f64>,
    /// Column is hidden
    pub hidden: bool,
    /// Outline/grouping level (0-7)
    pub outline_level: u8,
    /// Column-level style index (None = no column style)
    pub style_index: Option<u32>,
    /// Column is collapsed (in outline)
    pub collapsed: bool,
    /// Best fit (auto-sized)
    pub best_fit: bool,
}

impl Column {
    /// Create a new column with default settings
    pub fn new(index: u16) -> Self {
        Self {
            index,
            width: None,
            hidden: false,
            outline_level: 0,
            style_index: None,
            collapsed: false,
            best_fit: false,
        }
    }

    /// Check if this column has any custom settings
    pub fn has_custom_settings(&self) -> bool {
        self.width.is_some()
            || self.hidden
            || self.outline_level > 0
            || self.style_index.is_some()
            || self.collapsed
            || self.best_fit
    }
}

/// Column data (for reader/writer use)
#[derive(Debug, Clone)]
pub struct ColumnData {
    /// Start column index
    pub min: u16,
    /// End column index (inclusive)
    pub max: u16,
    /// Width
    pub width: Option<f64>,
    /// Hidden
    pub hidden: bool,
    /// Outline level
    pub outline_level: u8,
    /// Style index
    pub style_index: Option<u32>,
    /// Collapsed
    pub collapsed: bool,
    /// Best fit
    pub best_fit: bool,
}

impl ColumnData {
    /// Create column data for a single column
    pub fn single(index: u16) -> Self {
        Self {
            min: index,
            max: index,
            width: None,
            hidden: false,
            outline_level: 0,
            style_index: None,
            collapsed: false,
            best_fit: false,
        }
    }

    /// Create column data for a range of columns
    pub fn range(min: u16, max: u16) -> Self {
        Self {
            min,
            max,
            width: None,
            hidden: false,
            outline_level: 0,
            style_index: None,
            collapsed: false,
            best_fit: false,
        }
    }

    /// Set width
    pub fn with_width(mut self, width: f64) -> Self {
        self.width = Some(width);
        self
    }

    /// Set hidden
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }
}
