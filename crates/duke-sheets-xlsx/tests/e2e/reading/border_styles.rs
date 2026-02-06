//! Tests for reading border styles from PyUNO fixtures.
//!
//! Fixture: `border_styles.xlsx`

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_border_styles_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("border_styles.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open border_styles.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_cell_borders() {
    skip_if_no_fixtures!();

    let path = fixture_path("border_styles.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with borders
    let mut found_border = false;

    for row in 0..30 {
        for col in 0..5 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                // BorderStyle has Option<BorderEdge> fields for each side
                if style.border.left.is_some()
                    || style.border.right.is_some()
                    || style.border.top.is_some()
                    || style.border.bottom.is_some()
                {
                    found_border = true;
                    break;
                }
            }
        }
        if found_border {
            break;
        }
    }

    assert!(found_border, "Should find at least one cell with border");
}
