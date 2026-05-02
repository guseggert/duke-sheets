//! Round-trip tests for the XLS writer's built-in NAME records
//! (Print_Area, index 0x06; Print_Titles, index 0x07).

use std::io::Cursor;

use duke_sheets_core::{CellAddress, CellRange, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn range(start_addr: &str, end_addr: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start_addr).expect("parse start"),
        CellAddress::parse(end_addr).expect("parse end"),
    )
}

#[test]
fn print_area_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.print_area = Some(range("A1", "E20"));
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let area = parsed
        .worksheet(0)
        .unwrap()
        .page_setup()
        .print_area
        .clone()
        .expect("print area present after round-trip");
    assert_eq!(area.start.row, 0);
    assert_eq!(area.start.col, 0);
    assert_eq!(area.end.row, 19);
    assert_eq!(area.end.col, 4);
}

#[test]
fn print_area_persists_per_sheet() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First").expect("rename");
    wb.add_worksheet_with_name("Second").expect("add");
    let mut ps_first = wb.worksheet(0).unwrap().page_setup().clone();
    ps_first.print_area = Some(range("A1", "C3"));
    wb.worksheet_mut(0).unwrap().set_page_setup(ps_first);
    let mut ps_second = wb.worksheet(1).unwrap().page_setup().clone();
    ps_second.print_area = Some(range("D4", "F6"));
    wb.worksheet_mut(1).unwrap().set_page_setup(ps_second);

    let parsed = write_then_read(&wb);
    let first = parsed
        .worksheet_by_name("First")
        .unwrap()
        .page_setup()
        .print_area
        .clone()
        .expect("first");
    let second = parsed
        .worksheet_by_name("Second")
        .unwrap()
        .page_setup()
        .print_area
        .clone()
        .expect("second");
    assert_eq!(first.end.row, 2);
    assert_eq!(second.start.row, 3);
    assert_eq!(second.end.col, 5);
}

#[test]
fn repeat_rows_only_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.repeat_rows = Some((0, 1));
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let ps = parsed.worksheet(0).unwrap().page_setup();
    assert_eq!(ps.repeat_rows, Some((0, 1)));
    assert_eq!(ps.repeat_cols, None);
}

#[test]
fn repeat_cols_only_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.repeat_cols = Some((0, 0));
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let ps = parsed.worksheet(0).unwrap().page_setup();
    assert_eq!(ps.repeat_cols, Some((0, 0)));
    assert_eq!(ps.repeat_rows, None);
}

#[test]
fn repeat_rows_and_cols_together_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.repeat_rows = Some((0, 2));
    ps.repeat_cols = Some((0, 1));
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let ps = parsed.worksheet(0).unwrap().page_setup();
    assert_eq!(ps.repeat_rows, Some((0, 2)));
    assert_eq!(ps.repeat_cols, Some((0, 1)));
}

#[test]
fn no_print_area_means_no_name_record() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("A1");

    let parsed = write_then_read(&wb);
    assert!(parsed
        .worksheet(0)
        .unwrap()
        .page_setup()
        .print_area
        .is_none());
    assert!(parsed
        .worksheet(0)
        .unwrap()
        .page_setup()
        .repeat_rows
        .is_none());
}
