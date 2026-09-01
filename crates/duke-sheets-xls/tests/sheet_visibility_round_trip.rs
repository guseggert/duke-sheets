//! Round-trip tests for the XLS writer's BoundSheet8.hsState field
//! (sheet visibility: visible / hidden / very hidden).

use std::io::Cursor;

use duke_sheets_core::worksheet::SheetVisibility;
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

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

/// LibreOffice must accept BoundSheet8.hsState = Hidden / VeryHidden.
#[test]
fn lo_can_read_sheet_visibility_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Public").expect("rename");
    wb.add_worksheet_with_name("Hidden").expect("add hidden");
    wb.add_worksheet_with_name("VeryHidden").expect("add vh");
    wb.worksheet_mut(0).unwrap().set_cell_value("A1", "public").expect("A1");
    wb.worksheet_mut(1)
        .unwrap()
        .set_visibility(SheetVisibility::Hidden);
    wb.worksheet_mut(2)
        .unwrap()
        .set_visibility(SheetVisibility::VeryHidden);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_visibility_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<String, String> = rt.block_on(async {
        let mut bridge = duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect(
            "127.0.0.1",
            2002,
        )
        .await
        .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        wb.get_cell_string("A1")
            .await
            .map_err(|e| format!("A1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    assert_eq!(outcome.expect("LO must open visibility workbook"), "public");
}
