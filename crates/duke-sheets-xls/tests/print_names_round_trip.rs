//! Round-trip tests for the XLS writer's built-in NAME records
//! (Print_Area, index 0x06; Print_Titles, index 0x07).

use std::io::Cursor;

use duke_sheets_core::{CellAddress, CellRange, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

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

/// LibreOffice must accept our built-in NAME records (Print_Area
/// index 0x06, Print_Titles index 0x07). The built-in NAME body
/// uses tArea3D ptg with `ixti` referencing an EXTERNSHEET entry
/// pointing at the local sheet — wrong byte order or off-by-one
/// indices would render the file as a non-printable workbook.
#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_read_print_names_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "header").expect("A1");
    ws.set_cell_value("B1", "value").expect("B1");
    ws.set_cell_value("A2", "row1").expect("A2");
    ws.set_cell_value("B2", 100.0).expect("B2");
    let mut ps = ws.page_setup().clone();
    ps.print_area = Some(range("A1", "B2"));
    ps.repeat_rows = Some((0, 0));
    ws.set_page_setup(ps);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_print_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<(String, f64), String> = rt.block_on(async {
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
        let a1 = wb
            .get_cell_string("A1")
            .await
            .map_err(|e| format!("A1: {e}"))?;
        let b2 = wb
            .get_cell_value("B2")
            .await
            .map_err(|e| format!("B2: {e}"))?;
        Ok((a1, b2))
    });
    let _ = std::fs::remove_file(&path);
    let (a1, b2) = outcome.expect("LO must open print-names workbook");
    assert_eq!(a1, "header");
    assert!((b2 - 100.0).abs() < 1e-9, "B2 = {b2}");
}
