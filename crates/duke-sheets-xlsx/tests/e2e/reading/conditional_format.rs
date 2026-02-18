//! Tests for reading conditional formatting from XLSX files.
//! Includes DXF (differential format) style tests.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::{FillStyle, HorizontalAlignment, NumberFormat};
use duke_sheets_xlsx::XlsxReader;

/// Helper: create a workbook with values in B1:B5 and a conditional format rule.
async fn create_cf_fixture(
    path: &std::path::Path,
    operator: &str,
    formula: &str,
    style: &duke_sheets_libreoffice::StyleSpec,
) {
    let lo = lo_bridge().await.unwrap();
    let mut b = lo.lock().await;
    let mut wb = b.create_workbook().await.unwrap();

    for (i, val) in [10.0, 30.0, 50.0, 70.0, 90.0].iter().enumerate() {
        let cell = format!("B{}", i + 1);
        wb.set_cell_value(&cell, *val).await.unwrap();
    }

    let style_name = format!("CF_{}", path.file_stem().unwrap().to_str().unwrap());
    wb.add_conditional_format(0, "B1:B5", operator, formula, &style_name, style)
        .await
        .unwrap();

    wb.save(path.to_str().unwrap()).await.unwrap();
    wb.close().await.unwrap();
}

#[test]
fn test_cell_is_greater_than() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let style = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0x00FF00),
            bold: true,
            ..Default::default()
        };
        create_cf_fixture(&path, "greaterThan", "50", &style).await;
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
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
fn test_cf_dxf_bold_font() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let style = duke_sheets_libreoffice::StyleSpec {
            bold: true,
            ..Default::default()
        };
        create_cf_fixture(&path, "greaterThan", "50", &style).await;
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let rule = &rules[0];
    let format = rule.format.as_ref().expect("Rule should have a DXF format");
    assert!(format.font.bold, "DXF font should be bold");

    cleanup_fixture(&path);
}

#[test]
fn test_cf_dxf_fill() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let style = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0x00FF00),
            ..Default::default()
        };
        create_cf_fixture(&path, "greaterThan", "50", &style).await;
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let format = rules[0].format.as_ref().expect("Rule should have a DXF format");
    assert!(format.fill != FillStyle::None, "DXF should have non-None fill");

    cleanup_fixture(&path);
}

#[test]
fn test_cf_dxf_alignment() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let style = duke_sheets_libreoffice::StyleSpec {
            horizontal: Some("center".to_string()),
            wrap_text: true,
            ..Default::default()
        };
        create_cf_fixture(&path, "greaterThan", "50", &style).await;
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let format = rules[0].format.as_ref().expect("Rule should have a DXF format");
    assert_eq!(format.alignment.horizontal, HorizontalAlignment::Center);
    assert!(format.alignment.wrap_text, "DXF should have wrap_text");

    cleanup_fixture(&path);
}

#[test]
fn test_cf_dxf_number_format() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();

        for (i, val) in [0.1, 0.3, 0.5, 0.7, 0.9].iter().enumerate() {
            let cell = format!("A{}", i + 1);
            wb.set_cell_value(&cell, *val).await.unwrap();
        }

        let style = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("0.00%".to_string()),
            ..Default::default()
        };
        wb.add_conditional_format(0, "A1:A5", "greaterThan", "0.5", "CF_NumFmt", &style)
            .await
            .unwrap();

        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let format = rules[0].format.as_ref().expect("Rule should have a DXF format");
    assert!(format.number_format != NumberFormat::General, "DXF should have non-General number format");

    cleanup_fixture(&path);
}

#[test]
fn test_cf_dxf_border() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let style = duke_sheets_libreoffice::StyleSpec {
            border_style: Some("thin".to_string()),
            border_color: Some(0x0000FF),
            ..Default::default()
        };
        create_cf_fixture(&path, "greaterThan", "50", &style).await;
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let format = rules[0].format.as_ref().expect("Rule should have a DXF format");
    let has_edges = format.border.left.is_some()
        || format.border.right.is_some()
        || format.border.top.is_some()
        || format.border.bottom.is_some();
    assert!(has_edges, "DXF should have border edges");

    cleanup_fixture(&path);
}

#[test]
fn test_cf_multiple_rules() {
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

        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.conditional_formats();
    assert!(rules.len() >= 2, "Should have at least 2 CF rules, got {}", rules.len());

    for (i, rule) in rules.iter().enumerate() {
        assert!(rule.format.is_some(), "Rule {i} should have a DXF format");
    }

    cleanup_fixture(&path);
}
