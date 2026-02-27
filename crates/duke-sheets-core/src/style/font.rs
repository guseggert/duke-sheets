//! Font style types

use super::Color;

/// Font style settings
#[derive(Debug, Clone, PartialEq)]
pub struct FontStyle {
    /// Font family name (e.g., "Calibri", "Arial")
    pub name: String,
    /// Font size in points
    pub size: f64,
    /// Bold
    pub bold: bool,
    /// Italic
    pub italic: bool,
    /// Underline style
    pub underline: Underline,
    /// Strikethrough
    pub strikethrough: bool,
    /// Font color
    pub color: Color,
    /// Superscript/subscript
    pub vertical_align: FontVerticalAlign,
    /// Font family classification (OOXML `family` value)
    pub family: Option<u8>,
    /// Font charset (OOXML `charset` value)
    pub charset: Option<u8>,
    /// Font scheme (`major`/`minor`)
    pub scheme: Option<String>,
}

impl Default for FontStyle {
    fn default() -> Self {
        Self {
            name: "Calibri".to_string(),
            size: 11.0,
            bold: false,
            italic: false,
            underline: Underline::None,
            strikethrough: false,
            color: Color::Auto,
            vertical_align: FontVerticalAlign::Baseline,
            family: None,
            charset: None,
            scheme: None,
        }
    }
}

impl FontStyle {
    /// Create a new default font
    pub fn new() -> Self {
        Self::default()
    }

    /// Set font name
    pub fn with_name<S: Into<String>>(mut self, name: S) -> Self {
        self.name = name.into();
        self
    }

    /// Set font size
    pub fn with_size(mut self, size: f64) -> Self {
        self.size = size;
        self
    }

    /// Set bold
    pub fn with_bold(mut self, bold: bool) -> Self {
        self.bold = bold;
        self
    }

    /// Set italic
    pub fn with_italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }

    /// Set underline
    pub fn with_underline(mut self, underline: Underline) -> Self {
        self.underline = underline;
        self
    }

    /// Set strikethrough
    pub fn with_strikethrough(mut self, strikethrough: bool) -> Self {
        self.strikethrough = strikethrough;
        self
    }

    /// Set color
    pub fn with_color(mut self, color: Color) -> Self {
        self.color = color;
        self
    }

    /// Set font family classification
    pub fn with_family(mut self, family: Option<u8>) -> Self {
        self.family = family;
        self
    }

    /// Set font charset
    pub fn with_charset(mut self, charset: Option<u8>) -> Self {
        self.charset = charset;
        self
    }

    /// Set font scheme
    pub fn with_scheme<S: Into<String>>(mut self, scheme: Option<S>) -> Self {
        self.scheme = scheme.map(Into::into);
        self
    }
}

impl std::hash::Hash for FontStyle {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.name.hash(state);
        self.size.to_bits().hash(state);
        self.bold.hash(state);
        self.italic.hash(state);
        self.underline.hash(state);
        self.strikethrough.hash(state);
        self.color.hash(state);
        self.vertical_align.hash(state);
        self.family.hash(state);
        self.charset.hash(state);
        self.scheme.hash(state);
    }
}

impl Eq for FontStyle {}

/// Underline style
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Underline {
    /// No underline
    #[default]
    None,
    /// Single underline
    Single,
    /// Double underline
    Double,
    /// Single accounting underline (extends to cell width)
    SingleAccounting,
    /// Double accounting underline
    DoubleAccounting,
}

/// Font vertical alignment (superscript/subscript)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum FontVerticalAlign {
    /// Normal baseline
    #[default]
    Baseline,
    /// Superscript
    Superscript,
    /// Subscript
    Subscript,
}
