//! Round-trip tests for the XLS skeleton writer.
//!
//! Build an empty `Workbook`, write it to BIFF8 bytes, read it back
//! through `XlsReader`, and confirm the structure (sheet count, sheet
//! names) round-trips. The skeleton writer doesn't yet emit cells,
//! formatting, or formulas — those land in subsequent slices.

use std::io::Cursor;

use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

#[test]
fn empty_default_workbook_round_trips_via_reader() {
    let wb = Workbook::new();
    let original_count = wb.sheet_count();
    let original_name = wb.worksheet(0).unwrap().name().to_string();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize empty workbook");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back via XlsReader");

    assert_eq!(parsed.sheet_count(), original_count);
    assert_eq!(parsed.worksheet(0).unwrap().name(), original_name);
}

#[test]
fn renamed_sheet_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "CustomName").expect("rename");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back");

    assert_eq!(parsed.sheet_count(), 1);
    assert_eq!(parsed.worksheet(0).unwrap().name(), "CustomName");
}

#[test]
fn multi_sheet_round_trips_with_all_names() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Alpha").expect("rename Sheet1");
    wb.add_worksheet_with_name("Beta").expect("add Beta");
    wb.add_worksheet_with_name("Gamma").expect("add Gamma");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back");

    assert_eq!(parsed.sheet_count(), 3);
    let names: Vec<_> = parsed.worksheets().map(|s| s.name().to_string()).collect();
    assert_eq!(names, vec!["Alpha", "Beta", "Gamma"]);
}

#[test]
fn writes_cfb_v3_envelope() {
    let wb = Workbook::new();
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    assert_eq!(
        &bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "CFB magic header (MS-CFB §2.2)"
    );
    assert_eq!(
        u16::from_le_bytes([bytes[26], bytes[27]]),
        0x0003,
        "major version must be 3 (512-byte sectors) for .xls"
    );
}

#[test]
fn special_and_unicode_sheet_names_round_trip() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First & Last").expect("rename Sheet1");
    wb.add_worksheet_with_name("with 'apostrophe'")
        .expect("apostrophe sheet");
    wb.add_worksheet_with_name("日本語データ")
        .expect("unicode sheet");
    wb.add_worksheet_with_name("dash-dot.dot")
        .expect("dash-dot sheet");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back");

    let names: Vec<_> = parsed.worksheets().map(|s| s.name().to_string()).collect();
    assert_eq!(
        names,
        vec![
            "First & Last",
            "with 'apostrophe'",
            "日本語データ",
            "dash-dot.dot",
        ]
    );
}

#[test]
fn write_to_bytes_then_read_file_round_trips() {
    let wb = Workbook::new();
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    let temp = tempfile::NamedTempFile::new().expect("temp file");
    std::fs::write(temp.path(), &bytes).expect("write to temp");
    let parsed = XlsReader::read_file(temp.path()).expect("read back from disk");

    assert_eq!(parsed.sheet_count(), 1);
}

/// Probe whether LibreOffice's loadenv accepts our skeleton output.
/// Useful for empirical viability checks during writer development.
/// `#[ignore]`-gated because it needs a running LO container.
#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_skeleton_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "ProbeSheet").expect("rename");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_skeleton_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
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
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the skeleton workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_read_cell_values_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Probe").expect("rename");
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("set A1");
    ws.set_cell_value("B2", -3.14).expect("set B2");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_cells_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<(f64, f64), String> = rt.block_on(async {
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
            .get_cell_value("A1")
            .await
            .map_err(|e| format!("read A1: {e}"))?;
        let b2 = wb
            .get_cell_value("B2")
            .await
            .map_err(|e| format!("read B2: {e}"))?;
        Ok((a1, b2))
    });
    let _ = std::fs::remove_file(&path);
    let (a1, b2) = outcome.expect("LO must read cells we wrote");
    assert!((a1 - 42.0).abs() < 1e-9, "A1 = {a1}");
    assert!((b2 - -3.14).abs() < 1e-9, "B2 = {b2}");
}
