//! Tests for reading cell styles from XLS files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::{
    BorderLineStyle, CellValue, FillStyle, HorizontalAlignment, NumberFormat, VerticalAlignment,
};
use duke_sheets_xls::XlsReader;

// ── Font styles ─────────────────────────────────────────────────────────

#[test]
fn test_xls_font_bold() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Bold").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            bold: true,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "A1 should be bold");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_font_italic() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Italic").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            italic: true,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.italic, "A1 should be italic");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_font_size_and_name() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Big Arial").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            font_size: Some(20.0),
            font_name: Some("Arial".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        (style.font.size - 20.0).abs() < 0.5,
        "A1 font size should be ~20pt, got {}",
        style.font.size
    );
    assert_eq!(style.font.name, "Arial", "A1 font should be Arial");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_font_color() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Red text").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            font_color: Some(0xFF0000_u32 as i32), // Red
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    let (r, g, b) = style.font.color.to_rgb();
    assert!(
        r > 200 && g < 50 && b < 50,
        "A1 font color should be red-ish, got ({r}, {g}, {b})"
    );

    cleanup_fixture(&path);
}

// ── Fill styles ─────────────────────────────────────────────────────────

#[test]
fn test_xls_fill_solid_color() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Yellow bg").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0xFFFF00_u32 as i32), // Yellow
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    match &style.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert!(
                r > 200 && g > 200 && b < 50,
                "A1 fill should be yellow-ish, got ({r}, {g}, {b})"
            );
        }
        other => panic!("A1 should have solid fill, got {other:?}"),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_xls_fill_blue_background() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Blue bg").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0x0000FF_u32 as i32), // Blue
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    match &style.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert!(
                r < 50 && g < 50 && b > 200,
                "A1 fill should be blue-ish, got ({r}, {g}, {b})"
            );
        }
        other => panic!("A1 should have solid fill, got {other:?}"),
    }

    cleanup_fixture(&path);
}

// ── Border styles ───────────────────────────────────────────────────────

#[test]
fn test_xls_border_thin() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Thin border").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            border_style: Some("thin".into()),
            border_color: Some(0x000000),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    let border = &style.border;

    // At least one side should have a thin border
    let has_thin = [&border.left, &border.right, &border.top, &border.bottom]
        .iter()
        .any(|edge| {
            edge.as_ref()
                .map(|e| e.style == BorderLineStyle::Thin)
                .unwrap_or(false)
        });
    assert!(has_thin, "A1 should have at least one thin border edge, got {border:?}");

    cleanup_fixture(&path);
}

// ── Alignment ───────────────────────────────────────────────────────────

#[test]
fn test_xls_alignment_center() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Centered").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            horizontal: Some("center".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(
        style.alignment.horizontal,
        HorizontalAlignment::Center,
        "A1 should be center-aligned"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_alignment_right() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Right").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            horizontal: Some("right".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(
        style.alignment.horizontal,
        HorizontalAlignment::Right,
        "A1 should be right-aligned"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_alignment_vertical_center() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "VCenter").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            vertical: Some("center".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(
        style.alignment.vertical,
        VerticalAlignment::Center,
        "A1 should be vertically centered"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_wrap_text() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Long text that wraps")
            .await
            .unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            wrap_text: true,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.alignment.wrap_text, "A1 should have wrap text enabled");

    cleanup_fixture(&path);
}

// ── Number formats ──────────────────────────────────────────────────────

#[test]
fn test_xls_number_format_percentage() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 0.75).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("0%".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    // Could be BuiltIn(9) for "0%" or Custom("0%")
    let is_percent = match &style.number_format {
        NumberFormat::BuiltIn(9) | NumberFormat::BuiltIn(10) => true,
        NumberFormat::Custom(s) => s.contains('%'),
        _ => false,
    };
    assert!(
        is_percent,
        "A1 should have percentage format, got {:?}",
        style.number_format
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_number_format_date() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // Excel serial 45366 = 2024-03-12
        wb.set_cell_value("A1", 45366.0).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("YYYY-MM-DD".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    // Value should be the serial number
    let val = sheet.get_value_at(0, 0);
    match val {
        CellValue::Number(n) => assert!(
            (n - 45366.0).abs() < 0.01,
            "Date serial should be ~45366, got {}",
            n
        ),
        other => panic!("A1 should be Number, got {:?}", other),
    }

    // Format should be identified as a date format
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        style.number_format.is_date_format(),
        "A1 should have date format, got {:?}",
        style.number_format
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_number_format_datetime() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // 45366.5 = 2024-03-12 12:00:00
        wb.set_cell_value("A1", 45366.5).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("YYYY-MM-DD HH:MM:SS".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let val = sheet.get_value_at(0, 0);
    match val {
        CellValue::Number(n) => assert!(
            (n - 45366.5).abs() < 0.01,
            "DateTime serial should be ~45366.5, got {}",
            n
        ),
        other => panic!("A1 should be Number, got {:?}", other),
    }

    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        style.number_format.is_date_format(),
        "A1 should have datetime format, got {:?}",
        style.number_format
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_number_format_currency() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 1234.56).await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            number_format: Some("#,##0.00".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    // Should be either BuiltIn(4) for "#,##0.00" or Custom("#,##0.00")
    let is_number_fmt = match &style.number_format {
        NumberFormat::BuiltIn(4) => true,
        NumberFormat::Custom(s) => s.contains("#,##0"),
        _ => false,
    };
    assert!(
        is_number_fmt,
        "A1 should have #,##0.00 format, got {:?}",
        style.number_format
    );

    cleanup_fixture(&path);
}

// ── Combined styles ─────────────────────────────────────────────────────

#[test]
fn test_xls_combined_styles() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();

        // Cell with multiple style properties
        wb.set_cell_value("A1", "Styled").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            bold: true,
            italic: true,
            font_size: Some(14.0),
            font_color: Some(0x0000FF_u32 as i32), // Blue text
            fill_color: Some(0xFFFF00_u32 as i32),  // Yellow bg
            horizontal: Some("center".into()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();

        // A plain cell for comparison
        wb.set_cell_value("B1", "Plain").await.unwrap();

        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    // Check styled cell
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "A1 should be bold");
    assert!(style.font.italic, "A1 should be italic");
    assert!(
        (style.font.size - 14.0).abs() < 0.5,
        "A1 font size should be ~14pt, got {}",
        style.font.size
    );
    // Font color should be blue-ish
    let (r, g, b) = style.font.color.to_rgb();
    assert!(
        r < 50 && g < 50 && b > 200,
        "A1 font color should be blue-ish, got ({r}, {g}, {b})"
    );
    // Fill should be yellow-ish
    match &style.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert!(
                r > 200 && g > 200 && b < 50,
                "A1 fill should be yellow-ish, got ({r}, {g}, {b})"
            );
        }
        other => panic!("A1 should have solid fill, got {other:?}"),
    }
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Center);

    // Plain cell should have no style (or default)
    let plain_style = sheet.cell_style_at(0, 1);
    if let Some(ps) = plain_style {
        // If it has a style, it should be default-ish (not bold, etc.)
        assert!(!ps.font.bold, "B1 should not be bold");
    }

    cleanup_fixture(&path);
}
