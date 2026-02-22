//! Tests for reading cell comments from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_comment_basic_text() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Has comment").expect("set value");
        wb.add_comment("A1", "This is a comment")
            .expect("add comment");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let comment = sheet.comment_at(0, 0).expect("A1 should have a comment");
    assert!(
        comment.text.contains("This is a comment"),
        "Comment text should contain our text, got: {}",
        comment.text
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_comment_with_author() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Authored").expect("set value");
        wb.add_comment("A1", "Author comment").expect("add comment");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let comment = sheet.comment_at(0, 0).expect("A1 should have a comment");
    assert!(!comment.text.is_empty(), "Comment should have text");

    cleanup_fixture(&fixture);
}

#[test]
fn test_comment_unicode() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Unicode").expect("set value");
        wb.add_comment(
            "A1",
            "\u{65e5}\u{672c}\u{8a9e}\u{306e}\u{30b3}\u{30e1}\u{30f3}\u{30c8}",
        )
        .expect("add comment");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let comment = sheet.comment_at(0, 0).expect("A1 should have a comment");
    assert!(
        comment.text.contains("\u{65e5}\u{672c}\u{8a9e}"),
        "Comment should contain Japanese text, got: {}",
        comment.text
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_multiple_comments() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Comment 1").expect("set value");
        wb.add_comment("A1", "First comment").expect("add comment");
        wb.set_cell_value("A2", "Comment 2").expect("set value");
        wb.add_comment("A2", "Second comment").expect("add comment");
        wb.set_cell_value("A3", "Comment 3").expect("set value");
        wb.add_comment("A3", "Third comment").expect("add comment");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    assert_eq!(sheet.comment_count(), 3, "Should have 3 comments");

    cleanup_fixture(&fixture);
}

#[test]
fn test_comment_on_styled_cell() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Styled + Comment")
            .expect("set value");
        wb.set_font_bold("A1", true).expect("set bold");
        wb.set_fill_color("A1", 0xFFFF00).expect("set fill");
        wb.add_comment("A1", "Comment on styled cell")
            .expect("add comment");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");

    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "Should be bold");

    let comment = sheet.comment_at(0, 0).expect("A1 should have comment");
    assert!(!comment.text.is_empty(), "Comment should have text");

    cleanup_fixture(&fixture);
}
