#![allow(clippy::approx_constant)]
//! Round-trip tests for the XLS skeleton writer.
//!
//! Build an empty `Workbook`, write it to BIFF8 bytes, read it back
//! through `XlsReader`, and confirm the structure (sheet count, sheet
//! names) round-trips. The skeleton writer doesn't yet emit cells,
//! formatting, or formulas — those land in subsequent slices.

use std::io::Cursor;

use duke_sheets_core::{CellRange, PivotAggregate, PivotFilter, PivotTable, Workbook};
use duke_sheets_xls::{cfb::CompoundFile, XlsReader, XlsWriter};

fn add_test_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("BasicPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_average_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("AveragePivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Average, "Average Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_column_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Quarter").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Q1").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "East").unwrap();
    ws.set_cell_value("B3", "Q2").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "West").unwrap();
    ws.set_cell_value("B4", "Q1").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("RevenueByQuarter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_page_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Salesperson").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Ada").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Ben").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();

    let pivot = PivotTable::builder("RevenueByRep")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .page("Salesperson")
        .filter(PivotFilter::field_items("Salesperson", ["Ada"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

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
fn semantic_pivot_tables_emit_native_biff8_streams() {
    let mut wb = Workbook::new();
    add_test_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    assert!(cfb.exists("/_SX_DB_CUR/0001"));

    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = record_types(&workbook);
    assert!(workbook_records.contains(&0x00D5));
    assert!(workbook_records.contains(&0x00E3));
    assert!(workbook_records.contains(&0x0051));
    assert!(workbook_records.contains(&0x0160));
    assert!(workbook_records.contains(&0x089A));
    assert!(workbook_records.contains(&0x0802));
    assert!(workbook_records.contains(&0x0864));
    assert!(
        record_position(&workbook_records, 0x00D5) < record_position(&workbook_records, 0x0085),
        "SXIDSTM must be emitted before BoundSheet8 so Excel can bind the cache stream"
    );

    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = record_types(&cache);
    assert!(cache_records.contains(&0x00C6));
    assert_eq!(
        cache_records.iter().filter(|&&typ| typ == 0x00C7).count(),
        2
    );
    assert!(cache_records.contains(&0x00CD));
    assert!(cache_records.contains(&0x00C9));
}

#[test]
fn semantic_pivot_tables_emit_xls_average_data_field() {
    let mut wb = Workbook::new();
    add_average_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxdi = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C5).then_some(payload))
        .expect("SXDI data field record");

    assert_eq!(
        u16::from_le_bytes(sxdi[2..4].try_into().unwrap()),
        2,
        "BIFF8 data function 2 is Average"
    );
    assert!(
        String::from_utf8_lossy(sxdi).contains("Average Revenue"),
        "SXDI caption should come from the semantic pivot measure"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_column_axis_records() {
    let mut wb = Workbook::new();
    add_column_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let sxview = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW record");
    assert_eq!(u16::from_le_bytes(sxview[22..24].try_into().unwrap()), 3);
    assert_eq!(u16::from_le_bytes(sxview[24..26].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[26..28].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[34..36].try_into().unwrap()), 3);

    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxvd_axes, vec![0x0001, 0x0002, 0x0008]);

    let axis_declarations = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B4).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        axis_declarations,
        vec![vec![0, 0], vec![1, 0]],
        "row and column axes should declare Region then Quarter"
    );

    let sxli_lengths = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload.len()))
        .collect::<Vec<_>>();
    assert_eq!(sxli_lengths, vec![30, 30]);
}

#[test]
fn semantic_pivot_tables_emit_xls_page_axis_records() {
    let mut wb = Workbook::new();
    add_page_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let sxview = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW record");
    assert_eq!(u16::from_le_bytes(sxview[22..24].try_into().unwrap()), 3);
    assert_eq!(u16::from_le_bytes(sxview[24..26].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[26..28].try_into().unwrap()), 0);
    assert_eq!(u16::from_le_bytes(sxview[28..30].try_into().unwrap()), 1);

    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxvd_axes, vec![0x0001, 0x0004, 0x0008]);

    let page_fields = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B6).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(page_fields, vec![vec![1, 0, 0, 0, 1, 0]]);
}

fn record_types(stream: &[u8]) -> Vec<u16> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    while pos + 4 <= stream.len() {
        let typ = u16::from_le_bytes([stream[pos], stream[pos + 1]]);
        let len = u16::from_le_bytes([stream[pos + 2], stream[pos + 3]]) as usize;
        out.push(typ);
        pos += 4 + len;
    }
    out
}

fn records_with_payload(stream: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    while pos + 4 <= stream.len() {
        let typ = u16::from_le_bytes([stream[pos], stream[pos + 1]]);
        let len = u16::from_le_bytes([stream[pos + 2], stream[pos + 3]]) as usize;
        let start = pos + 4;
        let end = start + len;
        out.push((typ, stream[start..end].to_vec()));
        pos = end;
    }
    out
}

fn record_position(records: &[u16], typ: u16) -> usize {
    records
        .iter()
        .position(|&record| record == typ)
        .unwrap_or(usize::MAX)
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
    wb.rename_worksheet(0, "First & Last")
        .expect("rename Sheet1");
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
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
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
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
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
