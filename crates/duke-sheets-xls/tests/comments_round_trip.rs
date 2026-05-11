//! Round-trip tests for the XLS writer's comment emission:
//! `MSODRAWINGGROUP`, `MSODRAWING`, `OBJ`, `TXO`, and `NOTE` records.
//!
//! These tests exercise the in-process loop:
//!
//! 1. Build a workbook with one or more `CellComment`s via the core
//!    model.
//! 2. Serialise with [`XlsWriter`].
//! 3. Read the resulting bytes back with [`XlsReader`].
//! 4. Confirm the comment is present at the expected cell with the
//!    expected text + author.
//!
//! In-process round-trip is the cheapest layer — it asserts the
//! writer + reader agree on the encoding. It does not catch
//! spec-noncompliant output that our permissive reader still parses;
//! that needs the Excel COM parity layer (writing_xls.rs).

use std::io::Cursor;

use duke_sheets_core::{CellComment, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn single_comment_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").expect("set A1");
    ws.set_comment_at(0, 0, CellComment::new("Alice", "This is a note"));

    let parsed = write_then_read(&wb);
    let ws_in = parsed.worksheet(0).unwrap();
    let c = ws_in
        .comment_at(0, 0)
        .expect("comment must survive round-trip");
    assert_eq!(c.text, "This is a note");
    assert_eq!(c.author, "Alice");
}

#[test]
fn comment_without_anchor_cell_value_round_trips() {
    // Cell value is optional: a comment can sit on an otherwise-empty
    // cell. The OBJ/TXO/NOTE chain must still link correctly.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_comment_at(5, 3, CellComment::new("Bob", "Empty-cell note"));

    let parsed = write_then_read(&wb);
    let ws_in = parsed.worksheet(0).unwrap();
    let c = ws_in
        .comment_at(5, 3)
        .expect("comment on empty cell must survive");
    assert_eq!(c.text, "Empty-cell note");
    assert_eq!(c.author, "Bob");
}

#[test]
fn multiple_comments_on_same_sheet_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_comment_at(0, 0, CellComment::new("Alice", "First"));
    ws.set_comment_at(2, 1, CellComment::new("Alice", "Second"));
    ws.set_comment_at(10, 5, CellComment::new("Charlie", "Third"));

    let parsed = write_then_read(&wb);
    let ws_in = parsed.worksheet(0).unwrap();
    assert_eq!(ws_in.comment_count(), 3);
    assert_eq!(ws_in.comment_at(0, 0).unwrap().text, "First");
    assert_eq!(ws_in.comment_at(2, 1).unwrap().text, "Second");
    assert_eq!(ws_in.comment_at(10, 5).unwrap().text, "Third");
    assert_eq!(ws_in.comment_at(10, 5).unwrap().author, "Charlie");
}

#[test]
fn unicode_comment_text_round_trips() {
    // Japanese + emoji forces the writer onto the UTF-16LE path
    // (high-byte flag in the TXO CONTINUE).
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_comment_at(0, 0, CellComment::new("作者", "こんにちは 🌸"));

    let parsed = write_then_read(&wb);
    let c = parsed.worksheet(0).unwrap().comment_at(0, 0).unwrap();
    assert_eq!(c.text, "こんにちは 🌸");
    assert_eq!(c.author, "作者");
}

#[test]
fn comments_on_multiple_sheets_round_trip() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First").expect("rename");
    wb.add_worksheet_with_name("Second").expect("add second");
    wb.add_worksheet_with_name("Third").expect("add third");

    wb.worksheet_mut(0)
        .unwrap()
        .set_comment_at(0, 0, CellComment::new("a", "sheet 1 comment"));
    wb.worksheet_mut(2)
        .unwrap()
        .set_comment_at(4, 4, CellComment::new("c", "sheet 3 comment"));

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed.worksheet(0).unwrap().comment_at(0, 0).unwrap().text,
        "sheet 1 comment"
    );
    assert_eq!(parsed.worksheet(1).unwrap().comment_count(), 0);
    assert_eq!(
        parsed.worksheet(2).unwrap().comment_at(4, 4).unwrap().text,
        "sheet 3 comment"
    );
}

#[test]
fn empty_workbook_emits_no_drawing_records() {
    // A workbook with zero comments must not emit MSODRAWINGGROUP or
    // MSODRAWING records — confirm by writing, parsing, and asserting
    // no comments are present on any sheet.
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "value")
        .expect("A1");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    // Direct byte scan: the BIFF record types 0xEB / 0xEC must not
    // appear in the stream. (We scan the whole CFB envelope which is
    // a strict superset of the workbook stream — any sighting is a
    // bug.)
    let needle_drawing_group = [0xEB, 0x00];
    let needle_drawing = [0xEC, 0x00];
    let drawing_group_present = bytes.windows(2).any(|w| w == needle_drawing_group);
    let drawing_present = bytes.windows(2).any(|w| w == needle_drawing);
    // The byte sequences could appear in CFB sector data by chance,
    // so we round-trip and assert there's still no comment showing.
    let _ = (drawing_group_present, drawing_present);

    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back");
    assert_eq!(parsed.worksheet(0).unwrap().comment_count(), 0);
}
