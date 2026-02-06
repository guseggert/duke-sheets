//! Tests for reading fill/background styles from PyUNO fixtures.
//!
//! Fixture: `fill_styles.xlsx`

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_core::FillStyle;
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_fill_styles_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("fill_styles.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open fill_styles.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_solid_fill() {
    skip_if_no_fixtures!();

    let path = fixture_path("fill_styles.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with fill colors
    let mut found_fill = false;

    for row in 0..30 {
        for col in 0..5 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                // Check if fill is not None (has a solid color or pattern)
                match &style.fill {
                    FillStyle::Solid { color } => {
                        let (r, g, b) = color.to_rgb();
                        // Non-white fill (white is often default)
                        if r < 255 || g < 255 || b < 255 {
                            found_fill = true;
                            break;
                        }
                    }
                    FillStyle::Pattern { .. } | FillStyle::Gradient { .. } => {
                        found_fill = true;
                        break;
                    }
                    FillStyle::None => {}
                }
            }
        }
        if found_fill {
            break;
        }
    }

    assert!(found_fill, "Should find at least one cell with fill color");
}
