//! Fill/background style types

use super::Color;

/// Fill style for cell background
#[derive(Debug, Clone, PartialEq, Default)]
pub enum FillStyle {
    /// No fill (transparent)
    #[default]
    None,

    /// Solid color fill
    Solid { color: Color },

    /// Pattern fill
    Pattern {
        pattern: PatternType,
        foreground: Color,
        background: Color,
    },

    /// Gradient fill
    Gradient {
        gradient_type: GradientType,
        angle: f64,
        stops: Vec<GradientStop>,
    },
}

impl FillStyle {
    /// Create a solid fill with the given color
    pub fn solid(color: Color) -> Self {
        FillStyle::Solid { color }
    }

    /// Create a pattern fill
    pub fn pattern(pattern: PatternType, foreground: Color, background: Color) -> Self {
        FillStyle::Pattern {
            pattern,
            foreground,
            background,
        }
    }

    /// Create a linear gradient fill
    pub fn linear_gradient(angle: f64, stops: Vec<GradientStop>) -> Self {
        FillStyle::Gradient {
            gradient_type: GradientType::Linear,
            angle,
            stops,
        }
    }

    /// Check if this is a "no fill"
    pub fn is_none(&self) -> bool {
        matches!(self, FillStyle::None)
    }
}

impl std::hash::Hash for FillStyle {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        std::mem::discriminant(self).hash(state);
        match self {
            FillStyle::None => {}
            FillStyle::Solid { color } => {
                color.hash(state);
            }
            FillStyle::Pattern {
                pattern,
                foreground,
                background,
            } => {
                pattern.hash(state);
                foreground.hash(state);
                background.hash(state);
            }
            FillStyle::Gradient {
                gradient_type,
                angle,
                stops,
            } => {
                gradient_type.hash(state);
                angle.to_bits().hash(state);
                for stop in stops {
                    stop.position.to_bits().hash(state);
                    stop.color.hash(state);
                }
            }
        }
    }
}

impl Eq for FillStyle {}

/// Pattern fill types
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum PatternType {
    /// No pattern
    #[default]
    None,
    /// Solid (100% foreground)
    Solid,
    /// 50% gray
    MediumGray,
    /// 75% gray
    DarkGray,
    /// 25% gray
    LightGray,
    /// Horizontal stripe
    DarkHorizontal,
    /// Vertical stripe
    DarkVertical,
    /// Diagonal stripe (down)
    DarkDown,
    /// Diagonal stripe (up)
    DarkUp,
    /// Grid
    DarkGrid,
    /// Trellis
    DarkTrellis,
    /// Thin horizontal stripe
    LightHorizontal,
    /// Thin vertical stripe
    LightVertical,
    /// Thin diagonal stripe (down)
    LightDown,
    /// Thin diagonal stripe (up)
    LightUp,
    /// Thin grid
    LightGrid,
    /// Thin trellis
    LightTrellis,
    /// 12.5% gray
    Gray125,
    /// 6.25% gray
    Gray0625,
}

/// Gradient types
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum GradientType {
    /// Linear gradient
    #[default]
    Linear,
    /// Radial/path gradient
    Path,
}

/// Gradient stop (position and color)
#[derive(Debug, Clone, PartialEq)]
pub struct GradientStop {
    /// Position (0.0 to 1.0)
    pub position: f64,
    /// Color at this position
    pub color: Color,
}

impl GradientStop {
    /// Create a new gradient stop
    pub fn new(position: f64, color: Color) -> Self {
        Self { position, color }
    }
}
