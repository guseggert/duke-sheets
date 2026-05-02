//! Round-trip tests for the XLS writer's HLINK record (0x01B8) -
//! external URLs and internal `#Sheet!Cell` targets.

use std::io::Cursor;

use duke_sheets_core::Hyperlink;
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn external_url_hyperlink_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "click me").expect("A1");
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "https://example.com/path".to_string(),
            display: Some("click me".into()),
            tooltip: None,
            location: None,
        },
    )
    .expect("set hyperlink");

    let parsed = write_then_read(&wb);
    let hl = parsed
        .worksheet(0)
        .unwrap()
        .hyperlink("A1")
        .expect("hyperlink present after round-trip");
    assert_eq!(hl.target, "https://example.com/path");
    assert_eq!(hl.display.as_deref(), Some("click me"));
}

#[test]
fn internal_hash_target_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Sheet1").expect("rename");
    wb.add_worksheet_with_name("Other").expect("add");

    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "go").expect("A1");
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "#Other!B5".to_string(),
            display: Some("go".into()),
            tooltip: None,
            location: None,
        },
    )
    .expect("set hyperlink");

    let parsed = write_then_read(&wb);
    let hl = parsed
        .worksheet_by_name("Sheet1")
        .unwrap()
        .hyperlink("A1")
        .expect("hyperlink present");
    assert_eq!(hl.target, "#Other!B5");
}

#[test]
fn unicode_url_hyperlink_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "link").expect("A1");
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "https://例え.test/日本語".to_string(),
            display: Some("link".into()),
            tooltip: None,
            location: None,
        },
    )
    .expect("set hyperlink");

    let parsed = write_then_read(&wb);
    let hl = parsed
        .worksheet(0)
        .unwrap()
        .hyperlink("A1")
        .expect("hyperlink present");
    assert_eq!(hl.target, "https://例え.test/日本語");
}

#[test]
fn url_without_display_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "link").expect("A1");
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "https://example.com".to_string(),
            display: None,
            tooltip: None,
            location: None,
        },
    )
    .expect("set hyperlink");

    let parsed = write_then_read(&wb);
    let hl = parsed
        .worksheet(0)
        .unwrap()
        .hyperlink("A1")
        .expect("hyperlink present");
    assert_eq!(hl.target, "https://example.com");
}

#[test]
fn multiple_hyperlinks_round_trip_sorted() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "first").expect("A1");
    ws.set_cell_value("B2", "second").expect("B2");
    ws.set_cell_value("C3", "third").expect("C3");
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "https://first.example".into(),
            display: Some("first".into()),
            tooltip: None,
            location: None,
        },
    )
    .expect("A1");
    ws.set_hyperlink(
        "B2",
        Hyperlink {
            target: "https://second.example".into(),
            display: Some("second".into()),
            tooltip: None,
            location: None,
        },
    )
    .expect("B2");
    ws.set_hyperlink(
        "C3",
        Hyperlink {
            target: "https://third.example".into(),
            display: Some("third".into()),
            tooltip: None,
            location: None,
        },
    )
    .expect("C3");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        sheet.hyperlink("A1").map(|h| h.target.as_str()),
        Some("https://first.example")
    );
    assert_eq!(
        sheet.hyperlink("B2").map(|h| h.target.as_str()),
        Some("https://second.example")
    );
    assert_eq!(
        sheet.hyperlink("C3").map(|h| h.target.as_str()),
        Some("https://third.example")
    );
}
