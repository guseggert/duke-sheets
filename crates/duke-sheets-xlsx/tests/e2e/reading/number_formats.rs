//! Tests for reading number formats from PyUNO fixtures.
//!
//! Fixture: `number_formats.xlsx`

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_core::NumberFormat;
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_number_formats_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("number_formats.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open number_formats.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_number_format_applied() {
    skip_if_no_fixtures!();

    let path = fixture_path("number_formats.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with number formats
    let mut found_format = false;

    for row in 0..50 {
        for col in 0..5 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                // Check for non-General formats
                match &style.number_format {
                    NumberFormat::General => {}
                    NumberFormat::BuiltIn(_) | NumberFormat::Custom(_) => {
                        found_format = true;
                        break;
                    }
                }
            }
        }
        if found_format {
            break;
        }
    }

    assert!(
        found_format,
        "Should find at least one cell with number format"
    );
}
