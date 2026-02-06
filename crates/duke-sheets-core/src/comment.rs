//! Cell comments (notes)
//!
//! This module provides support for cell comments in worksheets.
//!
//! ## Example
//!
//! ```rust
//! use duke_sheets_core::{Workbook, CellComment};
//!
//! let mut workbook = Workbook::new();
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! // Add a comment to cell A1
//! sheet.set_comment("A1", CellComment::new("Author", "This is a note")).unwrap();
//!
//! // Get the comment back
//! let comment = sheet.comment("A1").unwrap();
//! assert!(comment.is_some());
//! ```

/// A cell comment/note
///
/// Comments are annotations attached to cells that can contain
/// author information and text content.
#[derive(Debug, Clone, PartialEq)]
pub struct CellComment {
    /// Author of the comment
    pub author: String,
    /// Comment text content
    pub text: String,
    /// Whether the comment box is visible by default
    pub visible: bool,
}

impl CellComment {
    /// Create a new comment with the given author and text
    ///
    /// # Example
    ///
    /// ```rust
    /// use duke_sheets_core::CellComment;
    ///
    /// let comment = CellComment::new("John Doe", "Review this value");
    /// assert_eq!(comment.author, "John Doe");
    /// assert_eq!(comment.text, "Review this value");
    /// assert!(!comment.visible);
    /// ```
    pub fn new(author: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            text: text.into(),
            visible: false,
        }
    }

    /// Create a comment with just text (empty author)
    pub fn text_only(text: impl Into<String>) -> Self {
        Self {
            author: String::new(),
            text: text.into(),
            visible: false,
        }
    }

    /// Set whether the comment is visible by default
    pub fn with_visible(mut self, visible: bool) -> Self {
        self.visible = visible;
        self
    }

    /// Check if this comment has an author
    pub fn has_author(&self) -> bool {
        !self.author.is_empty()
    }
}

impl Default for CellComment {
    fn default() -> Self {
        Self {
            author: String::new(),
            text: String::new(),
            visible: false,
        }
    }
}

impl std::fmt::Display for CellComment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        if self.has_author() {
            write!(f, "[{}]: {}", self.author, self.text)
        } else {
            write!(f, "{}", self.text)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_comment() {
        let comment = CellComment::new("Author", "Text");
        assert_eq!(comment.author, "Author");
        assert_eq!(comment.text, "Text");
        assert!(!comment.visible);
    }

    #[test]
    fn test_text_only() {
        let comment = CellComment::text_only("Just text");
        assert_eq!(comment.author, "");
        assert_eq!(comment.text, "Just text");
        assert!(!comment.has_author());
    }

    #[test]
    fn test_with_visible() {
        let comment = CellComment::new("A", "B").with_visible(true);
        assert!(comment.visible);
    }

    #[test]
    fn test_display() {
        let with_author = CellComment::new("John", "Hello");
        assert_eq!(format!("{}", with_author), "[John]: Hello");

        let without_author = CellComment::text_only("Hello");
        assert_eq!(format!("{}", without_author), "Hello");
    }
}
