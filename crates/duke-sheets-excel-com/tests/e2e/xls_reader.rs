use crate::{
    cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture_xls,
};
use duke_sheets_core::cell::CellValue;
use duke_sheets_core::style::PatternType;
use duke_sheets_core::FillStyle;
use duke_sheets_excel_com::{ChainStep, SheetRef};
use duke_sheets_xls::XlsReader;

#[test]
fn test_xls_rich_text() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Hello World").expect("set A1");
        wb.set_character_font_property("A1", 7, 5, "Bold", serde_json::Value::from(true))
            .expect("bold World");

        wb.set_cell_value("A2", "Red Blue").expect("set A2");
        wb.set_character_font_property("A2", 1, 3, "Italic", serde_json::Value::from(true))
            .expect("italic Red");
        wb.set_character_font_property("A2", 1, 3, "Color", serde_json::Value::from(0x0000FFi64))
            .expect("color Red");
        wb.set_character_font_property("A2", 5, 4, "Bold", serde_json::Value::from(true))
            .expect("bold Blue");
        wb.set_character_font_property("A2", 5, 4, "Color", serde_json::Value::from(0xFF0000i64))
            .expect("color Blue");

        wb.save(&fixture.vm_path).expect("save xls");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");

    match sheet.get_value_at(0, 0) {
        CellValue::RichText(runs) => {
            let text: String = runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(text, "Hello World", "A1 text");
            let has_bold = runs
                .iter()
                .any(|r| r.font.as_ref().is_some_and(|f| f.bold == Some(true)));
            assert!(has_bold, "A1 should have bold run: {runs:#?}");
        }
        CellValue::String(s) => {
            assert_eq!(s.as_ref(), "Hello World", "A1 collapsed to plain");
        }
        other => panic!("A1: expected RichText/String, got {other:?}"),
    }

    match sheet.get_value_at(1, 0) {
        CellValue::RichText(runs) => {
            let text: String = runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(text, "Red Blue", "A2 text");
            let has_italic = runs
                .iter()
                .any(|r| r.font.as_ref().is_some_and(|f| f.italic == Some(true)));
            let has_bold = runs
                .iter()
                .any(|r| r.font.as_ref().is_some_and(|f| f.bold == Some(true)));
            assert!(has_italic, "A2 should have italic run: {runs:#?}");
            assert!(has_bold, "A2 should have bold run: {runs:#?}");
        }
        CellValue::String(s) => {
            assert_eq!(s.as_ref(), "Red Blue", "A2 collapsed to plain");
        }
        other => panic!("A2: expected RichText/String, got {other:?}"),
    }

    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_zoom() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", "Zoom test").expect("set A1");
        wb.set_page_setup_property("Zoom", serde_json::Value::from(75))
            .expect("set zoom");
        wb.set_page_setup_property("FitToPagesWide", serde_json::Value::from(false))
            .expect("disable FitToPagesWide");
        wb.set_page_setup_property("FitToPagesTall", serde_json::Value::from(false))
            .expect("disable FitToPagesTall");
        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }
    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let ps = sheet.page_setup();
    assert_eq!(ps.scale, 75, "print scale should be 75, got {}", ps.scale);
    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_view_zoom() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", "View zoom test").expect("set A1");
        excel
            .set(
                0,
                vec![duke_sheets_excel_com::ChainStep::Property(
                    "ActiveWindow".to_string(),
                )],
                "Zoom",
                serde_json::Value::from(150),
            )
            .expect("set view zoom");
        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }
    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");
    assert_eq!(sheet.zoom_scale(), Some(150), "view zoom should be 150");
    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_print_gridlines() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", "Gridlines test").expect("set A1");
        wb.set_page_setup_property("PrintGridlines", serde_json::Value::from(true))
            .expect("set print gridlines");
        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }
    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");
    assert!(
        sheet.page_setup().print_gridlines,
        "print_gridlines should be true"
    );
    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_print_headings() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", "Headings test").expect("set A1");
        wb.set_page_setup_property("PrintHeadings", serde_json::Value::from(true))
            .expect("set print headings");
        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }
    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");
    assert!(
        sheet.page_setup().print_headings,
        "print_headings should be true"
    );
    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_print_area() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");
        for r in 0..10 {
            for c in 0..4u16 {
                let cell = format!("{}{}", (b'A' + c as u8) as char, r + 1);
                wb.set_cell_value(&cell, (r * 4 + c as u32) as f64)
                    .expect("set cell");
            }
        }
        wb.set_page_setup_property("PrintArea", serde_json::Value::from("$A$1:$D$10"))
            .expect("set print area");
        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }
    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let area = sheet.print_area().expect("should have print area");
    assert_eq!(area.start.row, 0, "print area first_row");
    assert_eq!(area.end.row, 9, "print area last_row");
    assert_eq!(area.start.col, 0, "print area first_col");
    assert_eq!(area.end.col, 3, "print area last_col");
    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_print_titles() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", "Header 1").expect("set A1");
        wb.set_cell_value("B1", "Header 2").expect("set B1");
        wb.set_cell_value("A2", 100.0).expect("set A2");
        wb.set_page_setup_property("PrintTitleRows", serde_json::Value::from("$1:$2"))
            .expect("set print title rows");
        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }
    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let repeat = sheet.repeat_rows().expect("should have repeat_rows");
    assert_eq!(repeat, (0, 1), "repeat_rows should be (0, 1)");
    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_print_titles_columns() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", "Row header").expect("set A1");
        wb.set_cell_value("B1", "Data").expect("set B1");
        wb.set_page_setup_property("PrintTitleColumns", serde_json::Value::from("$A:$B"))
            .expect("set print title cols");
        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }
    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let repeat = sheet.repeat_cols().expect("should have repeat_cols");
    assert_eq!(repeat, (0, 1), "repeat_cols should be (0, 1)");
    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_shared_formula() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        for (i, v) in [10.0, 20.0, 30.0, 40.0, 50.0].iter().enumerate() {
            let cell = format!("A{}", i + 1);
            wb.set_cell_value(&cell, *v).expect("set A");
        }
        for (i, v) in [1.0, 2.0, 3.0, 4.0, 5.0].iter().enumerate() {
            let cell = format!("B{}", i + 1);
            wb.set_cell_value(&cell, *v).expect("set B");
        }

        wb.set_cell_formula("C1", "=A1+B1").expect("set C1 formula");
        let h = wb.handle();
        excel
            .invoke(
                h,
                vec![
                    SheetRef::Index(0).to_chain_step(),
                    ChainStep::Indexed("Range".to_string(), serde_json::Value::from("C1:C5")),
                ],
                "FillDown",
                vec![],
            )
            .expect("FillDown");

        wb.save(&fixture.vm_path).expect("save xls");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");

    let expected = [11.0, 22.0, 33.0, 44.0, 55.0];
    for (row, expected_value) in expected.iter().enumerate() {
        match sheet.get_value_at(row as u32, 2) {
            CellValue::Formula {
                text, cached_value, ..
            } => {
                if row == 0 {
                    assert!(
                        text.contains("A1") && text.contains("B1"),
                        "C1 formula: {text}"
                    );
                }
                if row == 1 {
                    assert!(
                        text.contains("A2") && text.contains("B2"),
                        "C2 formula: {text}"
                    );
                }
                if row == 4 {
                    assert!(
                        text.contains("A5") && text.contains("B5"),
                        "C5 formula: {text}"
                    );
                }

                let cached = cached_value
                    .as_ref()
                    .unwrap_or_else(|| panic!("C{} missing cached value", row + 1));
                match cached.as_ref() {
                    CellValue::Number(n) => {
                        assert_eq!(*n, *expected_value, "C{} cached value", row + 1);
                    }
                    other => panic!("C{} expected numeric cache, got {other:?}", row + 1),
                }
            }
            other => panic!("C{} expected Formula, got {other:?}", row + 1),
        }
    }

    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_cse_array_formula() {
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        for (i, v) in [2.0, 3.0, 4.0].iter().enumerate() {
            let cell = format!("A{}", i + 1);
            wb.set_cell_value(&cell, *v).expect("set A");
        }
        for (i, v) in [5.0, 6.0, 7.0].iter().enumerate() {
            let cell = format!("B{}", i + 1);
            wb.set_cell_value(&cell, *v).expect("set B");
        }

        let h = wb.handle();
        excel
            .set(
                h,
                vec![
                    SheetRef::Index(0).to_chain_step(),
                    ChainStep::Indexed("Range".to_string(), serde_json::Value::from("C1")),
                ],
                "FormulaArray",
                serde_json::Value::from("=SUM(A1:A3*B1:B3)"),
            )
            .expect("set array formula");

        wb.save(&fixture.vm_path).expect("save xls");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");

    match sheet.get_value_at(0, 2) {
        CellValue::Formula {
            text, cached_value, ..
        } => {
            assert!(text.starts_with("{="), "C1 should be CSE formula: {text}");
            assert!(text.ends_with('}'), "C1 should end with }}: {text}");
            assert!(text.contains("SUM"), "C1 should contain SUM: {text}");
            assert!(text.contains("A1:A3"), "C1 should contain A1:A3: {text}");
            assert!(text.contains("B1:B3"), "C1 should contain B1:B3: {text}");

            let cached = cached_value.as_ref().expect("C1 missing cached value");
            match cached.as_ref() {
                CellValue::Number(n) => assert_eq!(*n, 56.0, "C1 cached value"),
                other => panic!("C1 expected numeric cache, got {other:?}"),
            }
        }
        other => panic!("C1 expected Formula, got {other:?}"),
    }

    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_data_table_formula() {
    // Range.Table() requires Range objects as arguments.  The bridge now
    // supports {"$ref": handle} in invoke args to pass stored COM objects.
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Input").expect("header");
        wb.set_cell_formula("D1", "=A1*2").expect("master formula");
        for (i, v) in [1.0, 2.0, 3.0, 4.0, 5.0].iter().enumerate() {
            let cell = format!("A{}", i + 2);
            wb.set_cell_value(&cell, *v).expect("set input");
        }

        let h = wb.handle();

        // Get a handle to Range("A1") — the column-input cell for the data table.
        // Use navigate() to walk the chain and store the endpoint as a handle.
        let a1_handle = excel
            .navigate(
                h,
                vec![
                    SheetRef::Index(0).to_chain_step(),
                    ChainStep::Indexed("Range".to_string(), serde_json::Value::from("A1")),
                ],
            )
            .expect("navigate to Range A1");

        // Call Range("A1:D6").Table(RowInput:=Nothing, ColumnInput:=Range("A1"))
        // using {"$ref": handle} to pass the Range object
        excel
            .invoke(
                h,
                vec![
                    SheetRef::Index(0).to_chain_step(),
                    ChainStep::Indexed("Range".to_string(), serde_json::Value::from("A1:D6")),
                ],
                "Table",
                vec![
                    serde_json::Value::Null,
                    serde_json::json!({"$ref": a1_handle}),
                ],
            )
            .expect("create data table");

        // Release the Range handle
        let _ = excel.release(a1_handle);

        wb.save(&fixture.vm_path).expect("save xls");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");

    let expected = [2.0, 4.0, 6.0, 8.0, 10.0];
    for (i, expected_value) in expected.iter().enumerate() {
        let row = (i + 1) as u32;
        match sheet.get_value_at(row, 3) {
            CellValue::Formula {
                text, cached_value, ..
            } => {
                let formula_upper = text.to_ascii_uppercase();
                assert!(
                    formula_upper.contains("TABLE"),
                    "D{} should contain TABLE formula text: {}",
                    i + 2,
                    text
                );

                let cached = cached_value
                    .as_ref()
                    .unwrap_or_else(|| panic!("D{} missing cached value", i + 2));
                match cached.as_ref() {
                    CellValue::Number(n) => {
                        assert_eq!(*n, *expected_value, "D{} cached value", i + 2);
                    }
                    other => panic!("D{} expected numeric cache, got {other:?}", i + 2),
                }
            }
            other => panic!("D{} expected Formula, got {other:?}", i + 2),
        }
    }

    cleanup_fixture(&fixture);
}

/// Helper: build the chain to a cell's Interior object.
fn interior_chain(cell: &str) -> Vec<ChainStep> {
    vec![
        SheetRef::Index(0).to_chain_step(),
        ChainStep::Indexed("Range".to_string(), serde_json::Value::from(cell)),
        ChainStep::Property("Interior".to_string()),
    ]
}

/// Helper: convert RGB (0xRRGGBB) to BGR for Excel COM Interior properties.
fn rgb_to_bgr(rgb: u32) -> u32 {
    let r = (rgb >> 16) & 0xFF;
    let g = (rgb >> 8) & 0xFF;
    let b = rgb & 0xFF;
    (b << 16) | (g << 8) | r
}

#[test]
fn test_xls_pattern_fills() {
    // Test multiple pattern fill types with distinct foreground/background colors.
    // Excel COM: Interior.Pattern = xlPattern constant
    //            Interior.PatternColor = pattern line color (→ BIFF icv_fore → our foreground)
    //            Interior.Color = background color (→ BIFF icv_back → our background)
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();

    // (cell, xlPattern constant, pattern_rgb, bg_rgb, expected PatternType)
    let cases: &[(&str, i64, u32, u32, PatternType)] = &[
        // xlPatternGray50 = -4125 → BIFF 2 → MediumGray
        ("A1", -4125, 0xFF0000, 0x0000FF, PatternType::MediumGray),
        // xlPatternGray75 = -4126 → BIFF 3 → DarkGray
        ("A2", -4126, 0x00FF00, 0xFFFF00, PatternType::DarkGray),
        // xlPatternGray25 = -4124 → BIFF 4 → LightGray
        ("A3", -4124, 0x0000FF, 0xFF00FF, PatternType::LightGray),
        // xlPatternHorizontal = -4128 → BIFF 5 → DarkHorizontal
        ("A4", -4128, 0x800000, 0x008080, PatternType::DarkHorizontal),
        // xlPatternVertical = -4166 → BIFF 6 → DarkVertical
        ("A5", -4166, 0x808000, 0x800080, PatternType::DarkVertical),
    ];

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        for &(cell, xl_pattern, pattern_rgb, bg_rgb, _) in cases {
            wb.set_cell_value(cell, format!("Pattern {cell}"))
                .expect("set value");

            let h = wb.handle();
            // Set pattern type
            excel
                .set(
                    h,
                    interior_chain(cell),
                    "Pattern",
                    serde_json::Value::from(xl_pattern),
                )
                .unwrap_or_else(|e| panic!("{cell} set Pattern: {e}"));

            // Set pattern line color (foreground in BIFF terms)
            excel
                .set(
                    h,
                    interior_chain(cell),
                    "PatternColor",
                    serde_json::Value::from(rgb_to_bgr(pattern_rgb)),
                )
                .unwrap_or_else(|e| panic!("{cell} set PatternColor: {e}"));

            // Set background color
            excel
                .set(
                    h,
                    interior_chain(cell),
                    "Color",
                    serde_json::Value::from(rgb_to_bgr(bg_rgb)),
                )
                .unwrap_or_else(|e| panic!("{cell} set Color: {e}"));
        }

        wb.save(&fixture.vm_path).expect("save xls");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");

    for (i, &(cell, _, pattern_rgb, bg_rgb, expected_pattern)) in cases.iter().enumerate() {
        let row = i as u32;
        let style = sheet
            .cell_style_at(row, 0)
            .unwrap_or_else(|| panic!("{cell} should have a style"));

        match &style.fill {
            FillStyle::Pattern {
                pattern,
                foreground,
                background,
            } => {
                assert_eq!(
                    *pattern, expected_pattern,
                    "{cell}: wrong PatternType — got {pattern:?}, expected {expected_pattern:?}"
                );

                // Verify foreground (pattern line) color
                let (pr, pg, pb) = expected_rgb(pattern_rgb);
                let (fr, fg, fb) = foreground.to_rgb();
                assert!(
                    close(fr, pr) && close(fg, pg) && close(fb, pb),
                    "{cell}: foreground expected ~({pr},{pg},{pb}), got ({fr},{fg},{fb})"
                );

                // Verify background color
                let (br_e, bg_e, bb_e) = expected_rgb(bg_rgb);
                let (br, bg_a, bb) = background.to_rgb();
                assert!(
                    close(br, br_e) && close(bg_a, bg_e) && close(bb, bb_e),
                    "{cell}: background expected ~({br_e},{bg_e},{bb_e}), got ({br},{bg_a},{bb})"
                );
            }
            other => panic!("{cell}: expected Pattern fill, got {other:?}"),
        }
    }

    cleanup_fixture(&fixture);
}

#[test]
fn test_xls_solid_fill() {
    // Verify solid fill round-trips through real Excel as XLS.
    let bridge = excel_bridge();
    let fixture = temp_fixture_xls();
    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", "Red solid").expect("set value");
        wb.set_fill_color("A1", 0xFF0000).expect("set fill");
        wb.save(&fixture.vm_path).expect("save xls");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsReader::read_file(&fixture.host_path).expect("XlsReader");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    match &style.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert!(
                r > 200 && g < 50 && b < 50,
                "Expected red solid fill, got ({r}, {g}, {b})"
            );
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }

    cleanup_fixture(&fixture);
}

/// Break an 0xRRGGBB u32 into (r, g, b).
fn expected_rgb(rgb: u32) -> (u8, u8, u8) {
    (
        ((rgb >> 16) & 0xFF) as u8,
        ((rgb >> 8) & 0xFF) as u8,
        (rgb & 0xFF) as u8,
    )
}

/// Allow ±2 tolerance for palette rounding.
fn close(a: u8, b: u8) -> bool {
    (a as i16 - b as i16).unsigned_abs() <= 2
}
