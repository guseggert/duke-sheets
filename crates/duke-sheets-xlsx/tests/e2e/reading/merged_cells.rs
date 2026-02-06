//! Tests for reading merged cells from PyUNO fixtures.
//!
//! Fixture: `merged_cells.xlsx`

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_merged_cells_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("merged_cells.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open merged_cells.xlsx: {:?}",
        result.err()
    );
}

#[test]
#[ignore = "Merged cell reading not yet implemented"]
fn test_merged_regions_present() {
    skip_if_no_fixtures!();

    let path = fixture_path("merged_cells.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // TODO: Once merged cell reading is implemented, verify merges
    let merges = sheet.merged_regions();
    assert!(!merges.is_empty(), "Should have at least one merged region");
}
