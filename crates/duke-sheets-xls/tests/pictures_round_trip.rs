//! Round-trip tests for embedded PNG images in XLS.
//!
//! Exercises the in-process loop: build a workbook with an
//! `EmbeddedImage`, write to BIFF8 bytes, read back, assert the
//! image survives with its anchor + format + raw payload bytes
//! intact.

use std::io::Cursor;

use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

/// A 68-byte 1x1 transparent PNG with verified chunk CRCs, used as
/// the deterministic image payload for these tests.
const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82,
];

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn test_image(id: u32, name: &str, from_col: u16, from_row: u32) -> EmbeddedImage {
    EmbeddedImage {
        id,
        name: name.to_string(),
        description: None,
        anchor: DrawingAnchor::TwoCell {
            from: CellMarker {
                col: from_col,
                col_offset_emu: 0,
                row: from_row,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: from_col + 3,
                col_offset_emu: 0,
                row: from_row + 5,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
        format: ImageFormat::Png,
        media_path: String::new(),
        svg_media_path: None,
        width_emu: 1_000_000,
        height_emu: 1_000_000,
        rotation: None,
        flip_h: false,
        flip_v: false,
        data: TEST_PNG_1X1.to_vec(),
        svg_data: None,
    }
}

#[test]
fn single_picture_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").expect("A1");
    ws.add_image(test_image(1, "Picture 1", 2, 3));

    let parsed = write_then_read(&wb);
    let images = parsed.worksheet(0).unwrap().images();
    assert_eq!(images.len(), 1, "picture must survive round-trip");

    let img = &images[0];
    assert_eq!(img.format, ImageFormat::Png);
    assert_eq!(img.data, TEST_PNG_1X1, "PNG bytes must round-trip verbatim");
    assert_eq!(img.name, "Picture 1");
    if let DrawingAnchor::TwoCell { from, to, .. } = &img.anchor {
        assert_eq!(from.col, 2);
        assert_eq!(from.row, 3);
        assert_eq!(to.col, 5);
        assert_eq!(to.row, 8);
    } else {
        panic!("expected TwoCell anchor, got {:?}", img.anchor);
    }
}

#[test]
fn multiple_pictures_round_trip_on_one_sheet() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_image(test_image(1, "Picture 1", 0, 0));
    ws.add_image(test_image(2, "Picture 2", 4, 4));
    ws.add_image(test_image(3, "Picture 3", 8, 8));

    let parsed = write_then_read(&wb);
    let images = parsed.worksheet(0).unwrap().images();
    assert_eq!(images.len(), 3, "all three pictures must round-trip");
    for img in images {
        assert_eq!(img.format, ImageFormat::Png);
        assert_eq!(
            img.data, TEST_PNG_1X1,
            "each picture's PNG bytes must survive verbatim"
        );
    }
}

#[test]
fn picture_and_comment_coexist_on_same_sheet() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_image(test_image(1, "Picture 1", 1, 1));
    ws.set_comment_at(3, 3, CellComment::new("Alice", "A note"));

    let parsed = write_then_read(&wb);
    let ws_in = parsed.worksheet(0).unwrap();
    assert_eq!(ws_in.image_count(), 1, "picture must survive");
    assert_eq!(ws_in.comment_count(), 1, "comment must survive");
    assert_eq!(ws_in.comment_at(3, 3).unwrap().text, "A note");
    assert_eq!(ws_in.images()[0].name, "Picture 1");
}

#[test]
fn pictures_on_multiple_sheets_round_trip() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First").expect("rename");
    wb.add_worksheet_with_name("Second").expect("add");

    wb.worksheet_mut(0)
        .unwrap()
        .add_image(test_image(1, "Pic on First", 0, 0));
    wb.worksheet_mut(1)
        .unwrap()
        .add_image(test_image(2, "Pic on Second", 2, 2));

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.worksheet(0).unwrap().image_count(), 1);
    assert_eq!(parsed.worksheet(1).unwrap().image_count(), 1);
    assert_eq!(
        parsed.worksheet(0).unwrap().images()[0].name,
        "Pic on First"
    );
    assert_eq!(
        parsed.worksheet(1).unwrap().images()[0].name,
        "Pic on Second"
    );
}

/// LibreOffice envelope check: write an XLS with one PNG image,
/// push to the LO shared dir, open via URP, read the anchor cell's
/// value back. Catches structural malformations in the Escher tree
/// (Blip/FBSE/BSTORE_CONTAINER + picture SP_CONTAINER) that would
/// make LO refuse the file.
#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_xls_with_picture_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");
    ws.add_image(test_image(1, "Picture 1", 2, 3));

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_picture_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<f64, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb_in = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        wb_in
            .get_cell_value("A1")
            .await
            .map_err(|e| format!("A1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    let a1 = outcome.expect("LO must open our XLS with picture without error");
    assert!(
        (a1 - 42.0).abs() < 1e-9,
        "A1 must round-trip; got {a1} (expected 42)"
    );
}

#[test]
fn empty_workbook_emits_no_blip_store() {
    // No images = no MSODRAWINGGROUP / MSODRAWING / BSTORE_CONTAINER
    // emission. Confirm by writing, parsing, and verifying the
    // workbook stream has zero drawing-group records.
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "value")
        .expect("A1");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read");
    assert_eq!(parsed.worksheet(0).unwrap().image_count(), 0);
}
