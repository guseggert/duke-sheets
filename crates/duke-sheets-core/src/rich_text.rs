//! Rich text run types for per-character formatting within a cell.

use crate::style::{Color, FontVerticalAlign, Underline};

/// Font properties for a single rich text run.
///
/// All fields are optional — unset properties inherit from the cell's style.
/// This mirrors the OOXML `CT_RPrElt` (run properties element) where each
/// property is independently optional.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct RunFont {
    /// Bold
    pub bold: Option<bool>,
    /// Italic
    pub italic: Option<bool>,
    /// Font size in points
    pub size: Option<f64>,
    /// Text color
    pub color: Option<Color>,
    /// Font face name (e.g., "Calibri", "Arial")
    pub name: Option<String>,
    /// Underline style
    pub underline: Option<Underline>,
    /// Strikethrough
    pub strikethrough: Option<bool>,
    /// Superscript/subscript
    pub vertical_align: Option<FontVerticalAlign>,
    /// Font family classification (OOXML `family` value: 0–5)
    pub family: Option<u8>,
    /// Font charset (Windows charset ID)
    pub charset: Option<u8>,
    /// Font scheme (`major`/`minor`/`none`)
    pub scheme: Option<String>,
}

impl RunFont {
    /// Returns true if all properties are None (no formatting specified).
    pub fn is_empty(&self) -> bool {
        self.bold.is_none()
            && self.italic.is_none()
            && self.size.is_none()
            && self.color.is_none()
            && self.name.is_none()
            && self.underline.is_none()
            && self.strikethrough.is_none()
            && self.vertical_align.is_none()
            && self.family.is_none()
            && self.charset.is_none()
            && self.scheme.is_none()
    }
}

/// A single run of text with optional per-run formatting.
///
/// Rich text in OOXML is represented as a sequence of runs (`<r>` elements),
/// each with optional run properties (`<rPr>`) and required text (`<t>`).
/// A run without properties inherits the cell's style.
#[derive(Debug, Clone, PartialEq)]
pub struct RichTextRun {
    /// The text content of this run.
    pub text: String,
    /// Optional font properties (None = inherit cell style).
    pub font: Option<RunFont>,
}

impl RichTextRun {
    /// Create a plain run with no formatting (inherits cell style).
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            font: None,
        }
    }

    /// Create a run with specific font properties.
    pub fn with_font(text: impl Into<String>, font: RunFont) -> Self {
        Self {
            text: text.into(),
            font: Some(font),
        }
    }
}

/// Extract the plain text from a slice of rich text runs by concatenating
/// all run text segments.
pub fn rich_text_to_plain(runs: &[RichTextRun]) -> String {
    let total: usize = runs.iter().map(|r| r.text.len()).sum();
    let mut result = String::with_capacity(total);
    for run in runs {
        result.push_str(&run.text);
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn plain_run() {
        let run = RichTextRun::plain("Hello");
        assert_eq!(run.text, "Hello");
        assert!(run.font.is_none());
    }

    #[test]
    fn run_with_font() {
        let font = RunFont {
            bold: Some(true),
            size: Some(14.0),
            ..Default::default()
        };
        let run = RichTextRun::with_font("Bold", font);
        assert_eq!(run.text, "Bold");
        assert_eq!(run.font.as_ref().unwrap().bold, Some(true));
        assert_eq!(run.font.as_ref().unwrap().size, Some(14.0));
        assert!(run.font.as_ref().unwrap().italic.is_none());
    }

    #[test]
    fn empty_run_font() {
        let font = RunFont::default();
        assert!(font.is_empty());

        let font = RunFont {
            bold: Some(true),
            ..Default::default()
        };
        assert!(!font.is_empty());
    }

    #[test]
    fn rich_text_plain_extraction() {
        let runs = vec![
            RichTextRun::plain("Hello "),
            RichTextRun::with_font(
                "World",
                RunFont {
                    bold: Some(true),
                    ..Default::default()
                },
            ),
        ];
        assert_eq!(rich_text_to_plain(&runs), "Hello World");
    }
}
