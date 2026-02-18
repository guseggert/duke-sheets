//! Tests for reading cell comments from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_comment_basic_text() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Has comment").await.unwrap();
        wb.add_comment(0, "A1", "This is a comment", None).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let comment = sheet.comment_at(0, 0).expect("A1 should have a comment");
    assert!(
        comment.text.contains("This is a comment"),
        "Comment text should contain our text, got: {}",
        comment.text
    );

    cleanup_fixture(&path);
}

#[test]
fn test_comment_with_author() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Authored").await.unwrap();
        wb.add_comment(0, "A1", "Author comment", Some("Test Author"))
            .await
            .unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let comment = sheet.comment_at(0, 0).expect("A1 should have a comment");
    assert!(!comment.text.is_empty(), "Comment should have text");

    cleanup_fixture(&path);
}

#[test]
fn test_comment_unicode() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Unicode").await.unwrap();
        wb.add_comment(0, "A1", "\u{65e5}\u{672c}\u{8a9e}\u{306e}\u{30b3}\u{30e1}\u{30f3}\u{30c8}", None)
            .await
            .unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let comment = sheet.comment_at(0, 0).expect("A1 should have a comment");
    assert!(
        comment.text.contains("\u{65e5}\u{672c}\u{8a9e}"),
        "Comment should contain Japanese text, got: {}",
        comment.text
    );

    cleanup_fixture(&path);
}

#[test]
fn test_multiple_comments() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Comment 1").await.unwrap();
        wb.add_comment(0, "A1", "First comment", None).await.unwrap();
        wb.set_cell_value("A2", "Comment 2").await.unwrap();
        wb.add_comment(0, "A2", "Second comment", None).await.unwrap();
        wb.set_cell_value("A3", "Comment 3").await.unwrap();
        wb.add_comment(0, "A3", "Third comment", None).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.comment_count(), 3, "Should have 3 comments");

    cleanup_fixture(&path);
}

#[test]
fn test_comment_on_styled_cell() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Styled + Comment").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            bold: true,
            fill_color: Some(0xFFFF00),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.add_comment(0, "A1", "Comment on styled cell", None)
            .await
            .unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "Should be bold");

    let comment = sheet.comment_at(0, 0).expect("A1 should have comment");
    assert!(!comment.text.is_empty(), "Comment should have text");

    cleanup_fixture(&path);
}
