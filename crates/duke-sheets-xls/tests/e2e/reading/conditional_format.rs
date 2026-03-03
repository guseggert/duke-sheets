//! Tests for reading conditional formatting from XLS files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xls::XlsReader;

#[test]
fn test_xls_cf_cell_is_greater_than() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();

        for (i, val) in [10.0, 30.0, 50.0, 70.0, 90.0].iter().enumerate() {
            let cell = format!("B{}", i + 1);
            wb.set_cell_value(&cell, *val).await.unwrap();
        }

        let style = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0x00FF00),
            bold: true,
            ..Default::default()
        };
        wb.add_conditional_format(0, "B1:B5", "greaterThan", "50", "CF_GreenBold", &style)
            .await
            .unwrap();

        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty(), "Should have at least one CF rule");

    let has_cell_is = rules
        .iter()
        .any(|r| matches!(&r.rule_type, duke_sheets_core::CfRuleType::CellIs { .. }));
    assert!(has_cell_is, "Should have a CellIs rule");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_cf_multiple_rules() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();

        for (i, val) in [10.0, 30.0, 50.0, 70.0, 90.0].iter().enumerate() {
            let cell = format!("B{}", i + 1);
            wb.set_cell_value(&cell, *val).await.unwrap();
        }

        let red_style = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0xFF0000),
            ..Default::default()
        };
        wb.add_conditional_format(0, "B1:B5", "greaterThan", "70", "CF_Red", &red_style)
            .await
            .unwrap();

        let green_style = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0x00FF00),
            ..Default::default()
        };
        wb.add_conditional_format(0, "B1:B5", "lessThan", "30", "CF_Green", &green_style)
            .await
            .unwrap();

        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(
        rules.len() >= 2,
        "Should have at least 2 CF rules, got {}",
        rules.len()
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_cf_formula() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();

        wb.set_cell_value("A1", 100.0).await.unwrap();
        let style = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0x00FF00),
            bold: true,
            ..Default::default()
        };
        wb.add_conditional_format(0, "A1", "equal", "100", "CF_Equals", &style)
            .await
            .unwrap();

        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty(), "Should have at least one CF rule");

    cleanup_fixture(&path);
}
