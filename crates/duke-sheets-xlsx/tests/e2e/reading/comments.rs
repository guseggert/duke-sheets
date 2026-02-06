//! Tests for reading cell comments from PyUNO fixtures.
//!
//! Fixture: `comments.xlsx`

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_comments_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("comments.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open comments.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_comments_present() {
    skip_if_no_fixtures!();

    let path = fixture_path("comments.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with comments
    let mut found_comment = false;

    for row in 0..30 {
        for col in 0..5 {
            if sheet.comment_at(row, col).is_some() {
                found_comment = true;
                break;
            }
        }
        if found_comment {
            break;
        }
    }

    assert!(found_comment, "Should find at least one cell with comment");
}

#[test]
fn test_comment_text() {
    skip_if_no_fixtures!();

    let path = fixture_path("comments.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Find a comment and verify it has text
    for row in 0..30 {
        for col in 0..5 {
            if let Some(comment) = sheet.comment_at(row, col) {
                assert!(
                    !comment.text.is_empty(),
                    "Comment should have non-empty text"
                );
                return; // Test passed
            }
        }
    }

    panic!("No comments found to test");
}
