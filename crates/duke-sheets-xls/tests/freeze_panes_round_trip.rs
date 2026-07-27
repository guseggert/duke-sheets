//! Round-trip tests for the XLS writer's PANE record (0x0041)
//! plus the WINDOW2 fFrozen / fFrozenNoSplit grbit flags.

use std::io::Cursor;

use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn freeze_first_two_rows_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_freeze_panes(2, 0);

    let parsed = write_then_read(&wb);
    let freeze = parsed
        .worksheet(0)
        .unwrap()
        .freeze_panes()
        .copied()
        .expect("freeze present after round-trip");
    assert_eq!(freeze.row, 2);
    assert_eq!(freeze.col, 0);
}

#[test]
fn freeze_first_column_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_freeze_panes(0, 1);

    let parsed = write_then_read(&wb);
    let freeze = parsed
        .worksheet(0)
        .unwrap()
        .freeze_panes()
        .copied()
        .expect("freeze present");
    assert_eq!(freeze.row, 0);
    assert_eq!(freeze.col, 1);
}

#[test]
fn freeze_both_row_and_column_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_freeze_panes(3, 2);

    let parsed = write_then_read(&wb);
    let freeze = parsed
        .worksheet(0)
        .unwrap()
        .freeze_panes()
        .copied()
        .expect("freeze present");
    assert_eq!(freeze.row, 3);
    assert_eq!(freeze.col, 2);
}

#[test]
fn no_freeze_means_no_pane_record() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");

    let parsed = write_then_read(&wb);
    assert!(parsed.worksheet(0).unwrap().freeze_panes().is_none());
}

#[test]
fn freeze_persists_per_sheet() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Frozen").expect("rename");
    wb.add_worksheet_with_name("Plain").expect("add");

    wb.worksheet_mut(0).unwrap().set_freeze_panes(1, 0);

    let parsed = write_then_read(&wb);
    let frozen = parsed.worksheet_by_name("Frozen").unwrap();
    let plain = parsed.worksheet_by_name("Plain").unwrap();
    assert_eq!(frozen.freeze_panes().map(|f| (f.row, f.col)), Some((1, 0)));
    assert!(plain.freeze_panes().is_none());
}

/// LibreOffice must accept our PANE + WINDOW2 records.
#[test]
fn lo_can_read_freeze_panes_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "header1").expect("A1");
    ws.set_cell_value("A2", "data").expect("A2");
    ws.set_freeze_panes(1, 1);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_freeze_{pid}.xls");
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
    assert_eq!(outcome.expect("LO must open freeze-panes workbook"), "header1");
}
