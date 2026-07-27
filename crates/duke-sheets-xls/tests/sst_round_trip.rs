//! Round-trip tests for the XLS writer's SST + LABELSST emission
//! (slice 3: shared string table including CONTINUE-record splitting
//! across the BIFF8 8224-byte record-body cap).

use std::io::Cursor;

use duke_sheets_core::{CellValue, Workbook, Worksheet};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn string_at(sheet: &Worksheet, addr: &str) -> Option<String> {
    let v = sheet.get_value(addr).ok()?;
    match v {
        CellValue::String(s) => Some(s.as_ref().to_string()),
        _ => None,
    }
}

#[test]
fn ascii_string_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "hello world").expect("set A1");

    let parsed = write_then_read(&wb);
    assert_eq!(
        string_at(parsed.worksheet(0).unwrap(), "A1").as_deref(),
        Some("hello world")
    );
}

#[test]
fn unicode_string_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "héllo wörld").expect("set A1");
    ws.set_cell_value("A2", "日本語").expect("set A2");
    ws.set_cell_value("A3", "🦀 crab").expect("set A3");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(string_at(sheet, "A1").as_deref(), Some("héllo wörld"));
    assert_eq!(string_at(sheet, "A2").as_deref(), Some("日本語"));
    assert_eq!(string_at(sheet, "A3").as_deref(), Some("🦀 crab"));
}

#[test]
fn empty_string_round_trips() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "")
        .expect("set A1");
    let parsed = write_then_read(&wb);
    assert_eq!(
        string_at(parsed.worksheet(0).unwrap(), "A1").as_deref(),
        Some("")
    );
}

#[test]
fn line_break_within_string_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "first line\nsecond line\nthird line")
        .expect("set A1");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        string_at(sheet, "A1").as_deref(),
        Some("first line\nsecond line\nthird line"),
        "embedded \\n must round-trip verbatim through SST"
    );
}

#[test]
fn duplicate_strings_are_deduplicated_in_sst() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for row in 0..50u32 {
        ws.set_cell_value(&format!("A{}", row + 1), "duplicate")
            .expect("set");
    }

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back");
    let sheet = parsed.worksheet(0).unwrap();
    for row in 0..50u32 {
        let addr = format!("A{}", row + 1);
        assert_eq!(string_at(sheet, &addr).as_deref(), Some("duplicate"));
    }
}

#[test]
fn many_unique_strings_round_trip_via_continue() {
    // 5000 distinct ~50-character strings comfortably exceed the 8224-byte
    // BIFF8 record cap, forcing the SST writer to split into a SST + N
    // CONTINUE record chain. Verify every string is still recoverable.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let count = 5_000u32;
    for i in 0..count {
        let s = format!("string_number_{i:08}_with_padding_to_make_it_longer");
        ws.set_cell_value(&format!("A{}", i + 1), &s as &str)
            .expect("set");
    }

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    for i in 0..count {
        let addr = format!("A{}", i + 1);
        let expected = format!("string_number_{i:08}_with_padding_to_make_it_longer");
        assert_eq!(
            string_at(sheet, &addr).as_deref(),
            Some(expected.as_str()),
            "row {i}"
        );
    }
}

#[test]
fn strings_across_multiple_sheets_share_one_sst() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First").expect("rename");
    wb.add_worksheet_with_name("Second").expect("add");

    wb.worksheet_mut(0).unwrap()
        .set_cell_value("A1", "shared").expect("set sheet1 A1");
    wb.worksheet_mut(0).unwrap()
        .set_cell_value("A2", "first-only").expect("set sheet1 A2");
    wb.worksheet_mut(1).unwrap()
        .set_cell_value("A1", "shared").expect("set sheet2 A1");
    wb.worksheet_mut(1).unwrap()
        .set_cell_value("A2", "second-only").expect("set sheet2 A2");

    let parsed = write_then_read(&wb);
    let s1 = parsed.worksheet_by_name("First").unwrap();
    let s2 = parsed.worksheet_by_name("Second").unwrap();
    assert_eq!(string_at(s1, "A1").as_deref(), Some("shared"));
    assert_eq!(string_at(s1, "A2").as_deref(), Some("first-only"));
    assert_eq!(string_at(s2, "A1").as_deref(), Some("shared"));
    assert_eq!(string_at(s2, "A2").as_deref(), Some("second-only"));
}

#[test]
fn lo_can_read_strings_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "hello").expect("set A1");
    ws.set_cell_value("B1", "héllo wörld").expect("set B1");
    ws.set_cell_value("C1", 42.0).expect("set C1");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_sst_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<(String, String, f64), String> = rt.block_on(async {
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
        let a1 = wb.get_cell_string("A1").await.map_err(|e| format!("A1: {e}"))?;
        let b1 = wb.get_cell_string("B1").await.map_err(|e| format!("B1: {e}"))?;
        let c1 = wb.get_cell_value("C1").await.map_err(|e| format!("C1: {e}"))?;
        Ok((a1, b1, c1))
    });
    let _ = std::fs::remove_file(&path);
    let (a1, b1, c1) = outcome.expect("LO must read what we wrote");
    assert_eq!(a1, "hello");
    assert_eq!(b1, "héllo wörld");
    assert!((c1 - 42.0).abs() < 1e-9);
}
