//! Tests for reading alignment properties from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::{HorizontalAlignment, VerticalAlignment};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_horizontal_center() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Centered").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            horizontal: Some("center".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Center);

    cleanup_fixture(&path);
}

#[test]
fn test_horizontal_right() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Right").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            horizontal: Some("right".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Right);

    cleanup_fixture(&path);
}

#[test]
fn test_vertical_bottom() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Bottom").await.unwrap();
        wb.set_row_height(0, 0, 40.0).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            vertical: Some("bottom".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.vertical, VerticalAlignment::Bottom);

    cleanup_fixture(&path);
}

#[test]
fn test_wrap_text() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "This is a long text that should wrap")
            .await
            .unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            wrap_text: true,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.alignment.wrap_text, "Wrap text should be true");

    cleanup_fixture(&path);
}

#[test]
fn test_shrink_to_fit() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Shrink").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            shrink_to_fit: true,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.alignment.shrink_to_fit, "Shrink to fit should be true");

    cleanup_fixture(&path);
}

#[test]
fn test_rotation() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Rotated 45").await.unwrap();
        wb.set_row_height(0, 0, 60.0).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            rotation: 45,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.rotation, 45, "Rotation should be 45 degrees");

    cleanup_fixture(&path);
}

#[test]
fn test_indent() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Indented").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            indent: 2,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.alignment.indent >= 1, "Indent should be >= 1, got {}", style.alignment.indent);

    cleanup_fixture(&path);
}
