//! Round-trip tests for the XLS writer's ROW (0x0208) and COLINFO
//! (0x007D) records: custom row heights, custom column widths,
//! hidden rows / columns, outline levels, and collapsed state.

use std::io::Cursor;

use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn custom_row_height_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_row_height(0, 30.0);
    ws.set_row_height(2, 50.0);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(
        (sheet.row_height(0) - 30.0).abs() < 1e-3,
        "row 0 height = {}",
        sheet.row_height(0)
    );
    assert!(
        (sheet.row_height(2) - 50.0).abs() < 1e-3,
        "row 2 height = {}",
        sheet.row_height(2)
    );
}

#[test]
fn custom_column_width_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_column_width(0, 25.0);
    ws.set_column_width(3, 12.5);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(
        (sheet.column_width(0) - 25.0).abs() < 0.05,
        "col A width = {}",
        sheet.column_width(0)
    );
    assert!(
        (sheet.column_width(3) - 12.5).abs() < 0.05,
        "col D width = {}",
        sheet.column_width(3)
    );
}

#[test]
fn hidden_rows_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_cell_value("A3", 3.0).expect("A3");
    ws.set_row_hidden(1, true);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(sheet.is_row_hidden(1), "row 1 must be hidden");
    assert!(!sheet.is_row_hidden(0), "row 0 must remain visible");
}

#[test]
fn hidden_columns_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_column_hidden(2, true);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(sheet.is_column_hidden(2), "col C must be hidden");
    assert!(!sheet.is_column_hidden(0), "col A must remain visible");
}

#[test]
fn row_outline_level_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_row_outline_level(2, 1);
    ws.set_row_outline_level(3, 2);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(sheet.row_outline_level(2), 1);
    assert_eq!(sheet.row_outline_level(3), 2);
}

#[test]
fn column_outline_level_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_column_outline_level(1, 1);
    ws.set_column_outline_level(2, 2);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(sheet.column_outline_level(1), 1);
    assert_eq!(sheet.column_outline_level(2), 2);
}

#[test]
fn row_collapsed_state_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_row_outline_level(2, 1);
    ws.set_row_collapsed(2, true);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(sheet.is_row_collapsed(2));
}

#[test]
fn column_collapsed_state_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_column_outline_level(2, 1);
    ws.set_column_collapsed(2, true);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(sheet.is_column_collapsed(2));
}

#[test]
fn combined_row_height_and_hidden_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_row_height(0, 40.0);
    ws.set_row_hidden(0, true);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(
        (sheet.row_height(0) - 40.0).abs() < 1e-3,
        "row 0 height = {}",
        sheet.row_height(0)
    );
    assert!(sheet.is_row_hidden(0));
}

#[test]
fn combined_column_width_and_hidden_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_column_width(1, 18.0);
    ws.set_column_hidden(1, true);

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(
        (sheet.column_width(1) - 18.0).abs() < 0.05,
        "col B width = {}",
        sheet.column_width(1)
    );
    assert!(sheet.is_column_hidden(1));
}
