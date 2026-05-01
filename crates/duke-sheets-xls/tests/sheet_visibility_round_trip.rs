//! Round-trip tests for the XLS writer's BoundSheet8.hsState field
//! (sheet visibility: visible / hidden / very hidden).

use std::io::Cursor;

use duke_sheets_core::worksheet::SheetVisibility;
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn default_visibility_round_trips() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("set");

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed.worksheet(0).unwrap().visibility(),
        SheetVisibility::Visible
    );
}

#[test]
fn hidden_sheet_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Visible").expect("rename");
    wb.add_worksheet_with_name("Hidden").expect("add");
    wb.worksheet_mut(1)
        .unwrap()
        .set_visibility(SheetVisibility::Hidden);

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed.worksheet_by_name("Visible").unwrap().visibility(),
        SheetVisibility::Visible
    );
    assert_eq!(
        parsed.worksheet_by_name("Hidden").unwrap().visibility(),
        SheetVisibility::Hidden
    );
}

#[test]
fn very_hidden_sheet_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Public").expect("rename");
    wb.add_worksheet_with_name("Internal").expect("add");
    wb.worksheet_mut(1)
        .unwrap()
        .set_visibility(SheetVisibility::VeryHidden);

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed.worksheet_by_name("Internal").unwrap().visibility(),
        SheetVisibility::VeryHidden
    );
}

#[test]
fn mixed_visibility_states_in_one_workbook_round_trip() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First").expect("rename");
    wb.add_worksheet_with_name("Second").expect("Second");
    wb.add_worksheet_with_name("Third").expect("Third");

    wb.worksheet_mut(1)
        .unwrap()
        .set_visibility(SheetVisibility::Hidden);
    wb.worksheet_mut(2)
        .unwrap()
        .set_visibility(SheetVisibility::VeryHidden);

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed.worksheet_by_name("First").unwrap().visibility(),
        SheetVisibility::Visible
    );
    assert_eq!(
        parsed.worksheet_by_name("Second").unwrap().visibility(),
        SheetVisibility::Hidden
    );
    assert_eq!(
        parsed.worksheet_by_name("Third").unwrap().visibility(),
        SheetVisibility::VeryHidden
    );
}
