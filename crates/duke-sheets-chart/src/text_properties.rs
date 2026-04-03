//! Text properties for chart text elements (titles, labels, axes).

/// Simplified text properties for chart text (titles, labels, axes).
/// Full DrawingML text model is huge — we capture the most common properties.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TextProperties {
    /// Rotation in 60,000ths of a degree
    pub rotation: Option<i32>,
    pub vertical: Option<TextVertical>,
    pub anchor: Option<TextAnchor>,
    pub anchor_center: Option<bool>,
    pub wrap: Option<TextWrap>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextVertical {
    Horizontal,
    Vertical,
    Vertical270,
    WordArt,
    WordArtVertical,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAnchor {
    Top,
    Middle,
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextWrap {
    None,
    Square,
}
