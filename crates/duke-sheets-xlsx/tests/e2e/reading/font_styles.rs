//! Tests for reading font styles from PyUNO fixtures.
//!
//! Fixture: `font_styles.xlsx`
//! - Bold, italic, underline, strikethrough
//! - Font colors (RGB)
//! - Font sizes
//! - Font names/families

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_font_styles_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("font_styles.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open font_styles.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_bold_style() {
    skip_if_no_fixtures!();

    let path = fixture_path("font_styles.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with bold styling
    let mut found_bold = false;

    for row in 0..20 {
        for col in 0..5 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                if style.font.bold {
                    found_bold = true;
                    break;
                }
            }
        }
        if found_bold {
            break;
        }
    }

    assert!(found_bold, "Should find at least one bold cell");
}

#[test]
fn test_italic_style() {
    skip_if_no_fixtures!();

    let path = fixture_path("font_styles.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with italic styling
    let mut found_italic = false;

    for row in 0..20 {
        for col in 0..5 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                if style.font.italic {
                    found_italic = true;
                    break;
                }
            }
        }
        if found_italic {
            break;
        }
    }

    assert!(found_italic, "Should find at least one italic cell");
}

#[test]
fn test_font_color() {
    skip_if_no_fixtures!();

    let path = fixture_path("font_styles.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with non-default font colors
    let mut found_colored = false;

    for row in 0..30 {
        for col in 0..5 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                // Check for non-auto colors (any explicit color set)
                if !style.font.color.is_auto() {
                    let (r, g, b) = style.font.color.to_rgb();
                    // Check for non-black colors
                    if r != 0 || g != 0 || b != 0 {
                        found_colored = true;
                        break;
                    }
                }
            }
        }
        if found_colored {
            break;
        }
    }

    assert!(
        found_colored,
        "Should find at least one cell with colored font"
    );
}

#[test]
fn test_font_size() {
    skip_if_no_fixtures!();

    let path = fixture_path("font_styles.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with non-default font sizes (default is 11.0)
    let mut found_large = false;
    let mut found_small = false;

    for row in 0..30 {
        for col in 0..5 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                let size = style.font.size;
                if size > 14.0 {
                    found_large = true;
                }
                if size < 10.0 {
                    found_small = true;
                }
            }
        }
    }

    assert!(
        found_large || found_small,
        "Should find cells with non-default font sizes"
    );
}
