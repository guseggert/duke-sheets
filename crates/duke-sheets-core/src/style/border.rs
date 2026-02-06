//! Border style types

use super::Color;

/// Border style for a cell
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct BorderStyle {
    /// Left border
    pub left: Option<BorderEdge>,
    /// Right border
    pub right: Option<BorderEdge>,
    /// Top border
    pub top: Option<BorderEdge>,
    /// Bottom border
    pub bottom: Option<BorderEdge>,
    /// Diagonal border
    pub diagonal: Option<BorderEdge>,
    /// Diagonal border direction
    pub diagonal_direction: DiagonalDirection,
}

impl BorderStyle {
    /// Create a new border style with no borders
    pub fn new() -> Self {
        Self::default()
    }

    /// Set all borders to the same style
    pub fn all(style: BorderLineStyle, color: Color) -> Self {
        let edge = Some(BorderEdge::new(style, color));
        Self {
            left: edge.clone(),
            right: edge.clone(),
            top: edge.clone(),
            bottom: edge,
            diagonal: None,
            diagonal_direction: DiagonalDirection::None,
        }
    }

    /// Set the left border
    pub fn with_left(mut self, style: BorderLineStyle, color: Color) -> Self {
        self.left = Some(BorderEdge::new(style, color));
        self
    }

    /// Set the right border
    pub fn with_right(mut self, style: BorderLineStyle, color: Color) -> Self {
        self.right = Some(BorderEdge::new(style, color));
        self
    }

    /// Set the top border
    pub fn with_top(mut self, style: BorderLineStyle, color: Color) -> Self {
        self.top = Some(BorderEdge::new(style, color));
        self
    }

    /// Set the bottom border
    pub fn with_bottom(mut self, style: BorderLineStyle, color: Color) -> Self {
        self.bottom = Some(BorderEdge::new(style, color));
        self
    }

    /// Set outline borders (left, right, top, bottom)
    pub fn outline(style: BorderLineStyle, color: Color) -> Self {
        Self::all(style, color)
    }

    /// Check if all borders are empty
    pub fn is_empty(&self) -> bool {
        self.left.is_none()
            && self.right.is_none()
            && self.top.is_none()
            && self.bottom.is_none()
            && self.diagonal.is_none()
    }
}

/// A single border edge
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct BorderEdge {
    /// Line style
    pub style: BorderLineStyle,
    /// Line color
    pub color: Color,
}

impl BorderEdge {
    /// Create a new border edge
    pub fn new(style: BorderLineStyle, color: Color) -> Self {
        Self { style, color }
    }

    /// Create a thin black border
    pub fn thin() -> Self {
        Self::new(BorderLineStyle::Thin, Color::BLACK)
    }

    /// Create a medium black border
    pub fn medium() -> Self {
        Self::new(BorderLineStyle::Medium, Color::BLACK)
    }

    /// Create a thick black border
    pub fn thick() -> Self {
        Self::new(BorderLineStyle::Thick, Color::BLACK)
    }
}

/// Border line styles
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum BorderLineStyle {
    /// No border
    #[default]
    None,
    /// Thin line
    Thin,
    /// Medium line
    Medium,
    /// Thick line
    Thick,
    /// Dashed line
    Dashed,
    /// Dotted line
    Dotted,
    /// Double line
    Double,
    /// Hair line (very thin)
    Hair,
    /// Medium dashed
    MediumDashed,
    /// Dash-dot
    DashDot,
    /// Medium dash-dot
    MediumDashDot,
    /// Dash-dot-dot
    DashDotDot,
    /// Medium dash-dot-dot
    MediumDashDotDot,
    /// Slant dash-dot
    SlantDashDot,
}

/// Diagonal border direction
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum DiagonalDirection {
    /// No diagonal
    #[default]
    None,
    /// Diagonal from top-left to bottom-right
    Down,
    /// Diagonal from bottom-left to top-right
    Up,
    /// Both diagonals
    Both,
}
