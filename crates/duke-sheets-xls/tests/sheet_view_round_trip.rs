//! Round-trip tests for the XLS writer's SELECTION (0x001D) and
//! SCL (0x00A0) records covering active cell, multi-range selection,
//! and zoom level.

use std::io::Cursor;

use duke_sheets_core::worksheet::Selection;
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn active_cell_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.add_selection(Selection {
        pane: None,
        active_cell: Some("C5".into()),
        sqref: Some("C5".into()),
    });

    let parsed = write_then_read(&wb);
    let sels = parsed.worksheet(0).unwrap().selections();
    assert_eq!(sels.len(), 1);
    assert_eq!(sels[0].active_cell.as_deref(), Some("C5"));
}

#[test]
fn single_range_selection_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.add_selection(Selection {
        pane: None,
        active_cell: Some("B2".into()),
        sqref: Some("B2:D5".into()),
    });

    let parsed = write_then_read(&wb);
    let sels = parsed.worksheet(0).unwrap().selections();
    assert_eq!(sels.len(), 1);
    assert_eq!(sels[0].active_cell.as_deref(), Some("B2"));
    assert_eq!(sels[0].sqref.as_deref(), Some("B2:D5"));
}

#[test]
fn multi_range_selection_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.add_selection(Selection {
        pane: None,
        active_cell: Some("A1".into()),
        sqref: Some("A1:B2 D4:E5".into()),
    });

    let parsed = write_then_read(&wb);
    let sels = parsed.worksheet(0).unwrap().selections();
    assert_eq!(sels.len(), 1);
    let sqref = sels[0].sqref.as_deref().expect("sqref present");
    assert!(sqref.contains("A1:B2"), "got {sqref:?}");
    assert!(sqref.contains("D4:E5"), "got {sqref:?}");
}

#[test]
fn zoom_at_75_percent_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_zoom_scale(Some(75));

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.worksheet(0).unwrap().zoom_scale(), Some(75));
}

#[test]
fn zoom_at_200_percent_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_zoom_scale(Some(200));

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.worksheet(0).unwrap().zoom_scale(), Some(200));
}

#[test]
fn zoom_at_100_percent_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_zoom_scale(Some(100));

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.worksheet(0).unwrap().zoom_scale(), Some(100));
}

#[test]
fn no_zoom_means_no_scl_record() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("A1");

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.worksheet(0).unwrap().zoom_scale(), None);
}
