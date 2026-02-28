/// A hyperlink attached to a cell
#[derive(Debug, Clone, PartialEq)]
pub struct Hyperlink {
    /// Target URL (external) or cell reference (internal)
    pub target: String,
    /// Display text (shown in cell, optional - cell value used if None)
    pub display: Option<String>,
    /// Tooltip shown on hover
    pub tooltip: Option<String>,
    /// Location within target (e.g., sheet reference for internal links)
    pub location: Option<String>,
}
