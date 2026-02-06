//! Color representation

use std::fmt;

/// Color representation
///
/// Supports RGB, ARGB, theme colors, and indexed colors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Color {
    /// Automatic/default color
    #[default]
    Auto,

    /// RGB color (no alpha)
    Rgb { r: u8, g: u8, b: u8 },

    /// ARGB color with alpha channel
    Argb { a: u8, r: u8, g: u8, b: u8 },

    /// Theme color with optional tint
    ///
    /// Theme indices:
    /// 0 = Background 1 (light)
    /// 1 = Text 1 (dark)
    /// 2 = Background 2
    /// 3 = Text 2
    /// 4-9 = Accent 1-6
    Theme {
        /// Theme color index (0-9)
        index: u8,
        /// Tint value (-1.0 to 1.0, stored as i8 percentage)
        tint: i8,
    },

    /// Indexed color (legacy Excel palette)
    Indexed(u8),
}

impl Color {
    /// Create an RGB color
    pub const fn rgb(r: u8, g: u8, b: u8) -> Self {
        Color::Rgb { r, g, b }
    }

    /// Create an ARGB color
    pub const fn argb(a: u8, r: u8, g: u8, b: u8) -> Self {
        Color::Argb { a, r, g, b }
    }

    /// Create a theme color
    pub const fn theme(index: u8, tint: i8) -> Self {
        Color::Theme { index, tint }
    }

    /// Create from a hex string (e.g., "#FF0000" or "FF0000")
    pub fn from_hex(hex: &str) -> Option<Self> {
        let hex = hex.trim_start_matches('#');

        match hex.len() {
            6 => {
                let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
                let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
                let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
                Some(Color::Rgb { r, g, b })
            }
            8 => {
                let a = u8::from_str_radix(&hex[0..2], 16).ok()?;
                let r = u8::from_str_radix(&hex[2..4], 16).ok()?;
                let g = u8::from_str_radix(&hex[4..6], 16).ok()?;
                let b = u8::from_str_radix(&hex[6..8], 16).ok()?;
                Some(Color::Argb { a, r, g, b })
            }
            _ => None,
        }
    }

    /// Convert to hex string (without # prefix)
    pub fn to_hex(&self) -> String {
        match self {
            Color::Auto => "000000".to_string(),
            Color::Rgb { r, g, b } => format!("{:02X}{:02X}{:02X}", r, g, b),
            Color::Argb { a, r, g, b } => format!("{:02X}{:02X}{:02X}{:02X}", a, r, g, b),
            Color::Theme { index, .. } => {
                // Return a placeholder for theme colors
                format!("theme{}", index)
            }
            Color::Indexed(i) => {
                // Return indexed color from standard palette
                let (r, g, b) = Self::indexed_to_rgb(*i);
                format!("{:02X}{:02X}{:02X}", r, g, b)
            }
        }
    }

    /// Convert to ARGB hex string (8 characters, used by XLSX)
    ///
    /// Always returns an 8-character string with alpha, e.g., "FFFF0000" for opaque red.
    pub fn to_argb_hex(&self) -> String {
        match self {
            Color::Auto => "FF000000".to_string(),
            Color::Rgb { r, g, b } => format!("FF{:02X}{:02X}{:02X}", r, g, b),
            Color::Argb { a, r, g, b } => format!("{:02X}{:02X}{:02X}{:02X}", a, r, g, b),
            Color::Theme { index, .. } => {
                // Convert theme to RGB first, then format
                let (r, g, b) = Self::theme_to_rgb(*index);
                format!("FF{:02X}{:02X}{:02X}", r, g, b)
            }
            Color::Indexed(i) => {
                let (r, g, b) = Self::indexed_to_rgb(*i);
                format!("FF{:02X}{:02X}{:02X}", r, g, b)
            }
        }
    }

    /// Convert to RGB tuple
    pub fn to_rgb(&self) -> (u8, u8, u8) {
        match self {
            Color::Auto => (0, 0, 0),
            Color::Rgb { r, g, b } => (*r, *g, *b),
            Color::Argb { r, g, b, .. } => (*r, *g, *b),
            Color::Theme { index, tint } => {
                // Get base theme color and apply tint
                let base = Self::theme_to_rgb(*index);
                Self::apply_tint(base, *tint)
            }
            Color::Indexed(i) => Self::indexed_to_rgb(*i),
        }
    }

    /// Check if color is automatic/default
    pub fn is_auto(&self) -> bool {
        matches!(self, Color::Auto)
    }

    /// Get RGB for indexed color
    fn indexed_to_rgb(index: u8) -> (u8, u8, u8) {
        // Standard Excel color palette (first 56 colors)
        const PALETTE: [(u8, u8, u8); 56] = [
            (0, 0, 0),       // 0: Black
            (255, 255, 255), // 1: White
            (255, 0, 0),     // 2: Red
            (0, 255, 0),     // 3: Bright Green
            (0, 0, 255),     // 4: Blue
            (255, 255, 0),   // 5: Yellow
            (255, 0, 255),   // 6: Pink
            (0, 255, 255),   // 7: Turquoise
            (0, 0, 0),       // 8: Black
            (255, 255, 255), // 9: White
            (255, 0, 0),     // 10: Red
            (0, 255, 0),     // 11: Bright Green
            (0, 0, 255),     // 12: Blue
            (255, 255, 0),   // 13: Yellow
            (255, 0, 255),   // 14: Pink
            (0, 255, 255),   // 15: Turquoise
            (128, 0, 0),     // 16: Dark Red
            (0, 128, 0),     // 17: Green
            (0, 0, 128),     // 18: Dark Blue
            (128, 128, 0),   // 19: Dark Yellow
            (128, 0, 128),   // 20: Violet
            (0, 128, 128),   // 21: Teal
            (192, 192, 192), // 22: 25% Gray
            (128, 128, 128), // 23: 50% Gray
            (153, 153, 255), // 24: Periwinkle
            (153, 51, 102),  // 25: Plum
            (255, 255, 204), // 26: Ivory
            (204, 255, 255), // 27: Light Turquoise
            (102, 0, 102),   // 28: Dark Purple
            (255, 128, 128), // 29: Coral
            (0, 102, 204),   // 30: Ocean Blue
            (204, 204, 255), // 31: Ice Blue
            (0, 0, 128),     // 32: Dark Blue
            (255, 0, 255),   // 33: Pink
            (255, 255, 0),   // 34: Yellow
            (0, 255, 255),   // 35: Turquoise
            (128, 0, 128),   // 36: Violet
            (128, 0, 0),     // 37: Dark Red
            (0, 128, 128),   // 38: Teal
            (0, 0, 255),     // 39: Blue
            (0, 204, 255),   // 40: Sky Blue
            (204, 255, 255), // 41: Light Turquoise
            (204, 255, 204), // 42: Light Green
            (255, 255, 153), // 43: Light Yellow
            (153, 204, 255), // 44: Pale Blue
            (255, 153, 204), // 45: Rose
            (204, 153, 255), // 46: Lavender
            (255, 204, 153), // 47: Tan
            (51, 102, 255),  // 48: Light Blue
            (51, 204, 204),  // 49: Aqua
            (153, 204, 0),   // 50: Lime
            (255, 204, 0),   // 51: Gold
            (255, 153, 0),   // 52: Light Orange
            (255, 102, 0),   // 53: Orange
            (102, 102, 153), // 54: Blue-Gray
            (150, 150, 150), // 55: 40% Gray
        ];

        if (index as usize) < PALETTE.len() {
            PALETTE[index as usize]
        } else {
            (0, 0, 0)
        }
    }

    /// Get RGB for theme color (using default Office theme)
    fn theme_to_rgb(index: u8) -> (u8, u8, u8) {
        match index {
            0 => (255, 255, 255), // Background 1 (white)
            1 => (0, 0, 0),       // Text 1 (black)
            2 => (238, 236, 225), // Background 2
            3 => (31, 73, 125),   // Text 2
            4 => (79, 129, 189),  // Accent 1
            5 => (192, 80, 77),   // Accent 2
            6 => (155, 187, 89),  // Accent 3
            7 => (128, 100, 162), // Accent 4
            8 => (75, 172, 198),  // Accent 5
            9 => (247, 150, 70),  // Accent 6
            _ => (0, 0, 0),
        }
    }

    /// Apply tint to a color
    fn apply_tint(color: (u8, u8, u8), tint: i8) -> (u8, u8, u8) {
        let tint_float = tint as f64 / 100.0;

        let apply = |c: u8| -> u8 {
            let c = c as f64;
            let result = if tint_float < 0.0 {
                c * (1.0 + tint_float)
            } else {
                c + (255.0 - c) * tint_float
            };
            result.clamp(0.0, 255.0) as u8
        };

        (apply(color.0), apply(color.1), apply(color.2))
    }

    // Common colors
    pub const BLACK: Color = Color::Rgb { r: 0, g: 0, b: 0 };
    pub const WHITE: Color = Color::Rgb {
        r: 255,
        g: 255,
        b: 255,
    };
    pub const RED: Color = Color::Rgb { r: 255, g: 0, b: 0 };
    pub const GREEN: Color = Color::Rgb { r: 0, g: 255, b: 0 };
    pub const BLUE: Color = Color::Rgb { r: 0, g: 0, b: 255 };
    pub const YELLOW: Color = Color::Rgb {
        r: 255,
        g: 255,
        b: 0,
    };
    pub const CYAN: Color = Color::Rgb {
        r: 0,
        g: 255,
        b: 255,
    };
    pub const MAGENTA: Color = Color::Rgb {
        r: 255,
        g: 0,
        b: 255,
    };
    pub const GRAY: Color = Color::Rgb {
        r: 128,
        g: 128,
        b: 128,
    };
    pub const LIGHT_GRAY: Color = Color::Rgb {
        r: 192,
        g: 192,
        b: 192,
    };
    pub const DARK_GRAY: Color = Color::Rgb {
        r: 64,
        g: 64,
        b: 64,
    };
}

impl fmt::Display for Color {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Color::Auto => write!(f, "auto"),
            Color::Rgb { r, g, b } => write!(f, "#{:02X}{:02X}{:02X}", r, g, b),
            Color::Argb { a, r, g, b } => write!(f, "#{:02X}{:02X}{:02X}{:02X}", a, r, g, b),
            Color::Theme { index, tint } => write!(f, "theme({}, {}%)", index, tint),
            Color::Indexed(i) => write!(f, "indexed({})", i),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_from_hex() {
        assert_eq!(
            Color::from_hex("#FF0000"),
            Some(Color::Rgb { r: 255, g: 0, b: 0 })
        );
        assert_eq!(
            Color::from_hex("00FF00"),
            Some(Color::Rgb { r: 0, g: 255, b: 0 })
        );
        assert_eq!(
            Color::from_hex("#80FFFFFF"),
            Some(Color::Argb {
                a: 128,
                r: 255,
                g: 255,
                b: 255
            })
        );
    }

    #[test]
    fn test_to_hex() {
        assert_eq!(Color::Rgb { r: 255, g: 0, b: 0 }.to_hex(), "FF0000");
        assert_eq!(
            Color::Argb {
                a: 128,
                r: 255,
                g: 255,
                b: 255
            }
            .to_hex(),
            "80FFFFFF"
        );
    }

    #[test]
    fn test_to_rgb() {
        assert_eq!(Color::RED.to_rgb(), (255, 0, 0));
        assert_eq!(Color::Indexed(2).to_rgb(), (255, 0, 0));
    }
}
