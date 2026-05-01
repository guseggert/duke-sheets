//! Round-trip tests for the XLS writer's WINDOW1 record (active
//! sheet index, MS-XLS §2.4.346).

use std::io::Cursor;

use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn default_active_sheet_is_first() {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Second").expect("add");
    wb.add_worksheet_with_name("Third").expect("add");

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.active_sheet(), 0);
}

#[test]
fn middle_sheet_active_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Alpha").expect("rename");
    wb.add_worksheet_with_name("Beta").expect("add");
    wb.add_worksheet_with_name("Gamma").expect("add");
    wb.set_active_sheet(1).expect("set active");

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.active_sheet(), 1);
    assert_eq!(parsed.worksheet(1).unwrap().name(), "Beta");
}

#[test]
fn last_sheet_active_round_trips() {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Two").expect("add");
    wb.add_worksheet_with_name("Three").expect("add");
    wb.add_worksheet_with_name("Four").expect("add");
    wb.set_active_sheet(3).expect("set active");

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.active_sheet(), 3);
}
