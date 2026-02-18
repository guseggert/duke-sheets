//! Cell styling types
//!
//! This module contains types for cell formatting:
//! - [`Style`] - Complete cell style
//! - [`FontStyle`] - Font settings
//! - [`FillStyle`] - Background fill
//! - [`BorderStyle`] - Cell borders
//! - [`Alignment`] - Text alignment
//! - [`Color`] - Color representation

mod alignment;
mod border;
mod color;
mod fill;
mod font;
mod number_format;
mod pool;

pub use alignment::{Alignment, HorizontalAlignment, ReadingOrder, VerticalAlignment};
pub use border::{BorderEdge, BorderLineStyle, BorderStyle, DiagonalDirection};
pub use color::Color;
pub use fill::{FillStyle, GradientStop, GradientType, PatternType};
pub use font::{FontStyle, FontVerticalAlign, Underline};
pub use number_format::NumberFormat;
pub use pool::StylePool;

/// Complete cell style
///
/// Styles are typically deduplicated via [`StylePool`] to save memory.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Style {
    /// Font settings
    pub font: FontStyle,
    /// Fill/background settings
    pub fill: FillStyle,
    /// Border settings
    pub border: BorderStyle,
    /// Text alignment
    pub alignment: Alignment,
    /// Number format
    pub number_format: NumberFormat,
    /// Cell protection
    pub protection: Protection,
}

impl Style {
    /// Create a new default style
    pub fn new() -> Self {
        Self::default()
    }

    /// Set font to bold
    pub fn bold(mut self, bold: bool) -> Self {
        self.font.bold = bold;
        self
    }

    /// Set font to italic
    pub fn italic(mut self, italic: bool) -> Self {
        self.font.italic = italic;
        self
    }

    /// Set font size in points
    pub fn font_size(mut self, size: f64) -> Self {
        self.font.size = size;
        self
    }

    /// Set font name
    pub fn font_name<S: Into<String>>(mut self, name: S) -> Self {
        self.font.name = name.into();
        self
    }

    /// Set font color
    pub fn font_color(mut self, color: Color) -> Self {
        self.font.color = color;
        self
    }

    /// Set fill color (solid fill)
    pub fn fill_color(mut self, color: Color) -> Self {
        self.fill = FillStyle::Solid { color };
        self
    }

    /// Set number format string
    pub fn number_format<S: Into<String>>(mut self, format: S) -> Self {
        self.number_format = NumberFormat::Custom(format.into());
        self
    }

    /// Set horizontal alignment
    pub fn horizontal_alignment(mut self, align: HorizontalAlignment) -> Self {
        self.alignment.horizontal = align;
        self
    }

    /// Set vertical alignment
    pub fn vertical_alignment(mut self, align: VerticalAlignment) -> Self {
        self.alignment.vertical = align;
        self
    }

    /// Enable text wrapping
    pub fn wrap_text(mut self, wrap: bool) -> Self {
        self.alignment.wrap_text = wrap;
        self
    }

    /// Get a mutable reference to font settings
    pub fn font_mut(&mut self) -> &mut FontStyle {
        &mut self.font
    }

    /// Get a mutable reference to fill settings
    pub fn fill_mut(&mut self) -> &mut FillStyle {
        &mut self.fill
    }

    /// Get a mutable reference to border settings
    pub fn border_mut(&mut self) -> &mut BorderStyle {
        &mut self.border
    }

    /// Get a mutable reference to alignment settings
    pub fn alignment_mut(&mut self) -> &mut Alignment {
        &mut self.alignment
    }
}

/// Cell protection settings
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Protection {
    /// Cell is locked (protected when sheet is protected)
    pub locked: bool,
    /// Formula is hidden when sheet is protected
    pub hidden: bool,
}

impl Protection {
    /// Create default protection (locked, not hidden)
    pub fn new() -> Self {
        Self {
            locked: true,
            hidden: false,
        }
    }

    /// Create unlocked protection
    pub fn unlocked() -> Self {
        Self {
            locked: false,
            hidden: false,
        }
    }
}

impl std::hash::Hash for Style {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.font.hash(state);
        self.fill.hash(state);
        self.border.hash(state);
        self.alignment.hash(state);
        self.number_format.hash(state);
        self.protection.locked.hash(state);
        self.protection.hidden.hash(state);
    }
}

impl Eq for Style {}
