//! Cell comments (notes)
//!
//! This module provides support for cell comments in worksheets.
//!
//! A comment is a drawing object: its popup placement and visibility
//! live on the wrapping [`crate::DrawingObject`] in the worksheet's
//! z-ordered drawing list (`hidden = true`, the default, shows the
//! note only on hover). [`CellComment`] carries the content, and the
//! keyed accessors on [`crate::Worksheet`] provide `(row, col)`
//! lookup over the list.
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
#[derive(Default)]
pub struct CellComment {
    /// Author of the comment
    pub author: String,
    /// Comment text content
    pub text: String,
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
    /// ```
    pub fn new(author: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            text: text.into(),
        }
    }

    /// Create a comment with just text (empty author)
    pub fn text_only(text: impl Into<String>) -> Self {
        Self {
            author: String::new(),
            text: text.into(),
        }
    }

    /// Check if this comment has an author
    pub fn has_author(&self) -> bool {
        !self.author.is_empty()
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
    }

    #[test]
    fn test_text_only() {
        let comment = CellComment::text_only("Just text");
        assert_eq!(comment.author, "");
        assert_eq!(comment.text, "Just text");
        assert!(!comment.has_author());
    }

    #[test]
    fn test_display() {
        let with_author = CellComment::new("John", "Hello");
        assert_eq!(format!("{}", with_author), "[John]: Hello");

        let without_author = CellComment::text_only("Hello");
        assert_eq!(format!("{}", without_author), "Hello");
    }
}
