#![allow(clippy::approx_constant)]
//! Round-trip tests for the XLS writer's cell-record emission
//! (slice 2: BLANK / NUMBER / BOOLERR records, no strings yet).

use std::io::Cursor;

use duke_sheets_core::style::Style;
use duke_sheets_core::{CellError, CellValue, Workbook, Worksheet};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn number_at(sheet: &Worksheet, addr: &str) -> Option<f64> {
    sheet.get_value(addr).ok().and_then(|v| v.as_number())
}

fn bool_at(sheet: &Worksheet, addr: &str) -> Option<bool> {
    sheet.get_value(addr).ok().and_then(|v| v.as_bool())
}

#[test]
fn number_value_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("set A1");
    ws.set_cell_value("B1", 3.14159).expect("set B1");
    ws.set_cell_value("C1", -1234.5).expect("set C1");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(number_at(sheet, "A1"), Some(42.0));
    assert_eq!(number_at(sheet, "B1"), Some(3.14159));
    assert_eq!(number_at(sheet, "C1"), Some(-1234.5));
}

#[test]
fn boolean_value_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", true).expect("set A1");
    ws.set_cell_value("A2", false).expect("set A2");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(bool_at(sheet, "A1"), Some(true));
    assert_eq!(bool_at(sheet, "A2"), Some(false));
}

#[test]
fn error_value_round_trips_for_standard_codes() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", CellValue::Error(CellError::Div0))
        .expect("set A1");
    ws.set_cell_value("A2", CellValue::Error(CellError::Value))
        .expect("set A2");
    ws.set_cell_value("A3", CellValue::Error(CellError::Ref))
        .expect("set A3");
    ws.set_cell_value("A4", CellValue::Error(CellError::Name))
        .expect("set A4");
    ws.set_cell_value("A5", CellValue::Error(CellError::Num))
        .expect("set A5");
    ws.set_cell_value("A6", CellValue::Error(CellError::Na))
        .expect("set A6");
    ws.set_cell_value("A7", CellValue::Error(CellError::Null))
        .expect("set A7");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let err_at = |addr: &str| match sheet.get_value(addr).expect(addr) {
        CellValue::Error(e) => e,
        other => panic!("{addr} expected Error, got {other:?}"),
    };
    assert_eq!(err_at("A1"), CellError::Div0);
    assert_eq!(err_at("A2"), CellError::Value);
    assert_eq!(err_at("A3"), CellError::Ref);
    assert_eq!(err_at("A4"), CellError::Name);
    assert_eq!(err_at("A5"), CellError::Num);
    assert_eq!(err_at("A6"), CellError::Na);
    assert_eq!(err_at("A7"), CellError::Null);
}

#[test]
fn cells_round_trip_when_set_in_scrambled_order() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("C2", 30.0).expect("set C2");
    ws.set_cell_value("A1", 10.0).expect("set A1");
    ws.set_cell_value("B1", 20.0).expect("set B1");
    ws.set_cell_value("A2", 40.0).expect("set A2");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(number_at(sheet, "A1"), Some(10.0));
    assert_eq!(number_at(sheet, "B1"), Some(20.0));
    assert_eq!(number_at(sheet, "A2"), Some(40.0));
    assert_eq!(number_at(sheet, "C2"), Some(30.0));
}

#[test]
fn mixed_numbers_and_strings_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    ws.set_cell_value("A2", "hello").expect("set A2");
    ws.set_cell_value("A3", 3.0).expect("set A3");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(number_at(sheet, "A1"), Some(1.0));
    let a2 = sheet.get_value("A2").expect("A2");
    assert!(
        matches!(&a2, CellValue::String(s) if s.as_ref() == "hello"),
        "got {a2:?}"
    );
    assert_eq!(number_at(sheet, "A3"), Some(3.0));
}

#[test]
fn empty_cell_with_format_only_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let bold = Style::new().bold(true);
    ws.set_cell_style("A1", &bold)
        .expect("set style on empty cell");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let style = sheet
        .cell_style("A1")
        .expect("cell_style ok")
        .cloned()
        .expect("style present on formatted-but-empty cell");
    assert!(
        style.font.bold,
        "format-only cell must preserve its style on round-trip"
    );
}

#[test]
fn large_row_and_column_indices_round_trip() {
    // BIFF8 caps at row 65535 (XLS-Excel display 65536) and col 255
    // (column IV). The writer should accept cells right up against
    // those limits.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value_at(0, 0, 1.0).expect("origin");
    ws.set_cell_value_at(65535, 0, 2.0)
        .expect("last row, col A");
    ws.set_cell_value_at(0, 255, 3.0).expect("row 0, last col");
    ws.set_cell_value_at(65535, 255, 4.0).expect("far corner");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(sheet.get_value_at(0, 0).as_number(), Some(1.0));
    assert_eq!(sheet.get_value_at(65535, 0).as_number(), Some(2.0));
    assert_eq!(sheet.get_value_at(0, 255).as_number(), Some(3.0));
    assert_eq!(sheet.get_value_at(65535, 255).as_number(), Some(4.0));
}

#[test]
fn many_rows_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for row in 0..100u32 {
        ws.set_cell_value(&format!("A{}", row + 1), row as f64)
            .expect("set");
    }
    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    for row in 0..100u32 {
        let addr = format!("A{}", row + 1);
        assert_eq!(number_at(sheet, &addr), Some(row as f64), "row {row}");
    }
}
