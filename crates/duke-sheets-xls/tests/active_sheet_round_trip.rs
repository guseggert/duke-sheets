//! Round-trip tests for the XLS writer's WINDOW1 record (active
//! sheet index, MS-XLS §2.4.346).

use std::io::Cursor;

use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

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

/// LibreOffice must accept the WINDOW1 record's iTabCur (active
/// sheet index) without rejecting the workbook.
#[test]
fn lo_can_read_active_sheet_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First").expect("rename");
    wb.add_worksheet_with_name("Second").expect("add");
    wb.add_worksheet_with_name("Third").expect("add");
    wb.worksheet_mut(0).unwrap().set_cell_value("A1", "alpha").expect("A1");
    wb.set_active_sheet(2).expect("set active");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_active_{pid}.xls");
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
    assert_eq!(outcome.expect("LO must open active-sheet workbook"), "alpha");
}
