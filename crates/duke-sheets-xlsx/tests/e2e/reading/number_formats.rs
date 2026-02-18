//! Tests for reading number formats from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::NumberFormat;
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_percentage_format() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 0.1234).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("0.00%".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    let fmt = style.number_format.format_string();
    assert!(fmt.contains('%'), "Number format should contain '%', got: {fmt}");

    cleanup_fixture(&path);
}

#[test]
fn test_currency_format() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 1234.56).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("\"$\"#,##0.00".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        style.number_format != NumberFormat::General,
        "Number format should not be General"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_custom_decimal_format() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 1234.5678).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("#,##0.00".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.number_format != NumberFormat::General, "Should not be General");
    let fmt = style.number_format.format_string();
    assert!(
        fmt.contains("#,##0") || fmt.contains("0.00"),
        "Format should be decimal, got: {fmt}"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_date_format() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 45366.0).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("YYYY-MM-DD".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.number_format != NumberFormat::General, "Should not be General");

    cleanup_fixture(&path);
}
