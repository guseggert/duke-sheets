//! Round-trip tests for the XLS writer's MERGECELLS emission
//! (slice 6: merged cell ranges, MS-XLS §2.4.169).

use std::io::Cursor;

use duke_sheets_core::{CellAddress, CellRange, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn merge_range(start_addr: &str, end_addr: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start_addr).expect("parse start"),
        CellAddress::parse(end_addr).expect("parse end"),
    )
}

#[test]
fn horizontal_merge_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "merged across columns")
        .expect("set A1");
    ws.merge_cells(&merge_range("A1", "C1"))
        .expect("merge A1:C1");

    let parsed = write_then_read(&wb);
    let regions = parsed.worksheet(0).unwrap().merged_regions();
    assert_eq!(regions.len(), 1);
    assert_eq!(regions[0].start, CellAddress::parse("A1").unwrap());
    assert_eq!(regions[0].end, CellAddress::parse("C1").unwrap());
}

#[test]
fn vertical_merge_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("B2", "merged down").expect("set B2");
    ws.merge_cells(&merge_range("B2", "B5"))
        .expect("merge B2:B5");

    let parsed = write_then_read(&wb);
    let regions = parsed.worksheet(0).unwrap().merged_regions();
    assert_eq!(regions.len(), 1);
    assert_eq!(regions[0].start, CellAddress::parse("B2").unwrap());
    assert_eq!(regions[0].end, CellAddress::parse("B5").unwrap());
}

#[test]
fn block_merge_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("D4", "merged block").expect("set D4");
    ws.merge_cells(&merge_range("D4", "F8"))
        .expect("merge D4:F8");

    let parsed = write_then_read(&wb);
    let regions = parsed.worksheet(0).unwrap().merged_regions();
    assert_eq!(regions.len(), 1);
    assert_eq!(regions[0].start, CellAddress::parse("D4").unwrap());
    assert_eq!(regions[0].end, CellAddress::parse("F8").unwrap());
}

#[test]
fn multiple_disjoint_merges_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "row 1 header").expect("set A1");
    ws.set_cell_value("A3", "row 3 header").expect("set A3");
    ws.set_cell_value("A5", "row 5 header").expect("set A5");
    ws.merge_cells(&merge_range("A1", "C1")).expect("merge 1");
    ws.merge_cells(&merge_range("A3", "C3")).expect("merge 2");
    ws.merge_cells(&merge_range("A5", "C5")).expect("merge 3");

    let parsed = write_then_read(&wb);
    let regions = parsed.worksheet(0).unwrap().merged_regions();
    assert_eq!(regions.len(), 3);
    let mut starts: Vec<_> = regions.iter().map(|r| r.start).collect();
    starts.sort_by_key(|a| (a.row, a.col));
    assert_eq!(starts[0], CellAddress::parse("A1").unwrap());
    assert_eq!(starts[1], CellAddress::parse("A3").unwrap());
    assert_eq!(starts[2], CellAddress::parse("A5").unwrap());
}

#[test]
fn merges_persist_across_multiple_sheets() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First").expect("rename");
    wb.add_worksheet_with_name("Second").expect("add");

    wb.worksheet_mut(0)
        .unwrap()
        .merge_cells(&merge_range("A1", "B2"))
        .expect("merge first");
    wb.worksheet_mut(1)
        .unwrap()
        .merge_cells(&merge_range("C3", "D4"))
        .expect("merge second");

    let parsed = write_then_read(&wb);
    let first = parsed.worksheet_by_name("First").unwrap();
    let second = parsed.worksheet_by_name("Second").unwrap();
    assert_eq!(first.merged_regions().len(), 1);
    assert_eq!(
        first.merged_regions()[0].start,
        CellAddress::parse("A1").unwrap()
    );
    assert_eq!(
        first.merged_regions()[0].end,
        CellAddress::parse("B2").unwrap()
    );
    assert_eq!(second.merged_regions().len(), 1);
    assert_eq!(
        second.merged_regions()[0].start,
        CellAddress::parse("C3").unwrap()
    );
    assert_eq!(
        second.merged_regions()[0].end,
        CellAddress::parse("D4").unwrap()
    );
}

#[test]
fn no_merges_emits_no_record_and_round_trips_clean() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("set");
    let parsed = write_then_read(&wb);
    assert!(parsed.worksheet(0).unwrap().merged_regions().is_empty());
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_read_merged_cells_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "merged").expect("set A1");
    ws.merge_cells(&merge_range("A1", "C1"))
        .expect("merge A1:C1");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_merge_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<String, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
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
    let a1 = outcome.expect("LO must read merged anchor");
    assert_eq!(a1, "merged");
}
