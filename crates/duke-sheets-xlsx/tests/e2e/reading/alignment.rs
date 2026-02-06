//! Tests for reading alignment from PyUNO fixtures.
//!
//! Fixture: `alignment.xlsx`, `alignment_rotation.xlsx`

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_core::HorizontalAlignment;
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_alignment_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("alignment.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open alignment.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_horizontal_alignment() {
    skip_if_no_fixtures!();

    let path = fixture_path("alignment.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with alignment settings
    let mut found_alignment = false;

    for row in 0..30 {
        for col in 0..5 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                // Check for non-General horizontal alignment
                if style.alignment.horizontal != HorizontalAlignment::General {
                    found_alignment = true;
                    break;
                }
            }
        }
        if found_alignment {
            break;
        }
    }

    assert!(
        found_alignment,
        "Should find at least one cell with horizontal alignment"
    );
}

#[test]
#[ignore = "Text rotation reading not yet implemented"]
fn test_text_rotation() {
    skip_if_no_fixtures!();

    let path = fixture_path("alignment_rotation.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let _sheet = workbook.worksheet(0).expect("No worksheet");

    // TODO: Test text rotation once implemented
}
