//! Text alignment types

/// Text alignment settings
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Alignment {
    /// Horizontal alignment
    pub horizontal: HorizontalAlignment,
    /// Vertical alignment
    pub vertical: VerticalAlignment,
    /// Wrap text
    pub wrap_text: bool,
    /// Shrink to fit
    pub shrink_to_fit: bool,
    /// Indent level (0-250)
    pub indent: u8,
    /// Text rotation in degrees (-90 to 90, or 255 for vertical)
    pub rotation: i16,
    /// Reading order
    pub reading_order: ReadingOrder,
}

impl Alignment {
    /// Create a new default alignment
    pub fn new() -> Self {
        Self::default()
    }

    /// Set horizontal alignment
    pub fn with_horizontal(mut self, align: HorizontalAlignment) -> Self {
        self.horizontal = align;
        self
    }

    /// Set vertical alignment
    pub fn with_vertical(mut self, align: VerticalAlignment) -> Self {
        self.vertical = align;
        self
    }

    /// Enable text wrapping
    pub fn with_wrap(mut self, wrap: bool) -> Self {
        self.wrap_text = wrap;
        self
    }

    /// Set indent level
    pub fn with_indent(mut self, indent: u8) -> Self {
        self.indent = indent;
        self
    }

    /// Set rotation angle
    pub fn with_rotation(mut self, degrees: i16) -> Self {
        self.rotation = degrees.clamp(-90, 90);
        self
    }

    /// Set vertical text (rotation = 255)
    pub fn vertical_text(mut self) -> Self {
        self.rotation = 255;
        self
    }
}

/// Horizontal alignment options
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum HorizontalAlignment {
    /// General alignment (text left, numbers right)
    #[default]
    General,
    /// Left aligned
    Left,
    /// Center aligned
    Center,
    /// Right aligned
    Right,
    /// Fill (repeat content to fill cell width)
    Fill,
    /// Justify (stretch to fit width)
    Justify,
    /// Center across selection
    CenterContinuous,
    /// Distributed (like justify, but for East Asian text)
    Distributed,
}

/// Vertical alignment options
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum VerticalAlignment {
    /// Top aligned
    Top,
    /// Center aligned
    Center,
    /// Bottom aligned (default)
    #[default]
    Bottom,
    /// Justify
    Justify,
    /// Distributed
    Distributed,
}

/// Reading order for text
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ReadingOrder {
    /// Context dependent
    #[default]
    ContextDependent,
    /// Left to right
    LeftToRight,
    /// Right to left
    RightToLeft,
}
