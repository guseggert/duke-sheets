//! Tests for reading cell comments from XLS files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xls::XlsReader;

#[test]
fn test_xls_comment_basic_text() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Has comment").await.unwrap();
        wb.add_comment(0, "A1", "This is a comment", None)
            .await
            .unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
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
fn test_xls_comment_unicode() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Unicode").await.unwrap();
        wb.add_comment(0, "A1", "日本語のコメント", None)
            .await
            .unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let comment = sheet.comment_at(0, 0).expect("A1 should have a comment");
    assert!(
        comment.text.contains("日本語のコメント"),
        "Comment should contain Japanese text, got: {}",
        comment.text
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_multiple_comments() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Comment 1").await.unwrap();
        wb.add_comment(0, "A1", "First comment", None)
            .await
            .unwrap();
        wb.set_cell_value("A2", "Comment 2").await.unwrap();
        wb.add_comment(0, "A2", "Second comment", None)
            .await
            .unwrap();
        wb.set_cell_value("A3", "Comment 3").await.unwrap();
        wb.add_comment(0, "A3", "Third comment", None)
            .await
            .unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.comment_count(), 3, "Should have 3 comments");

    cleanup_fixture(&path);
}
