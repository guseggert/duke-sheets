//! Roundtrip fidelity test: write XLSX with duke-sheets, open in real Excel,
//! verify no repair warning, verify data survives.
//!
//! This is the ultimate compatibility test - if Excel can open our file
//! without repairing it, the OOXML we produce is spec-compliant.

use crate::{
    cleanup_fixture, ensure_vm_temp_dir, excel_bridge, push_file_to_vm, temp_fixture,
    temp_fixture_xlsb,
};
use duke_sheets_core::{
    BorderLineStyle, CellRange, CellValue, Color, ConditionalFormatRule, DataValidation,
    HorizontalAlignment, NumberFormat, Style, VerticalAlignment,
};
use duke_sheets_xlsb::XlsbWriter;
use duke_sheets_xlsx::XlsxWriter;
use std::io::Cursor;

/// Write a workbook exercising many writer features, push to the VM,
/// open in real Excel, and verify no repair + data integrity.
#[test]
fn test_roundtrip_no_repair() {
    let fixture = temp_fixture();

    let mut wb = duke_sheets_core::Workbook::new();

    {
        let sheet = wb.worksheet_mut(0).unwrap();

        // Data types
        sheet.set_cell_value("A1", "Hello").unwrap();
        sheet.set_cell_value("A2", 42.0).unwrap();
        sheet.set_cell_value("A3", true).unwrap();
        sheet.set_cell_value("A4", CellValue::Number(2.0)).unwrap();
        sheet.set_cell_formula("A4", "=1+1").unwrap();

        // Styled cell: bold, red font, yellow fill
        let bold_style = Style::new()
            .bold(true)
            .font_color(Color::rgb(255, 0, 0))
            .fill_color(Color::rgb(255, 255, 0));
        sheet.set_cell_value("B1", "Styled").unwrap();
        sheet.set_cell_style("B1", &bold_style).unwrap();

        // Alignment: center + wrap
        let align_style = Style {
            alignment: duke_sheets_core::Alignment {
                horizontal: HorizontalAlignment::Center,
                vertical: VerticalAlignment::Center,
                wrap_text: true,
                ..Default::default()
            },
            ..Default::default()
        };
        sheet.set_cell_value("B2", "Centered\nWrapped").unwrap();
        sheet.set_cell_style("B2", &align_style).unwrap();

        // Border: thin all sides, blue
        let border_style = Style {
            border: duke_sheets_core::BorderStyle::all(
                BorderLineStyle::Thin,
                Color::rgb(0, 0, 255),
            ),
            ..Default::default()
        };
        sheet.set_cell_value("B3", "Bordered").unwrap();
        sheet.set_cell_style("B3", &border_style).unwrap();

        // Number format
        let pct_style = Style {
            number_format: NumberFormat::Custom("0.00%".into()),
            ..Default::default()
        };
        sheet.set_cell_value("C1", 0.1234).unwrap();
        sheet.set_cell_style("C1", &pct_style).unwrap();

        // Row height and column width
        sheet.set_row_height(4, 30.0);
        sheet.set_column_width(3, 20.0);

        // Merged cells: A7:C7
        sheet
            .merge_cells(&CellRange::parse("A7:C7").unwrap())
            .unwrap();
        sheet.set_cell_value("A7", "Merged region").unwrap();

        // NOTE: Comments omitted - our writer produces comments1.xml but not
        // the required VML drawing (vmlDrawing1.vml), which Excel rejects.
        // TODO: add VML drawing support to the writer.

        // Conditional formatting: highlight >50 with green fill
        sheet.set_cell_value("D1", 10.0).unwrap();
        sheet.set_cell_value("D2", 60.0).unwrap();
        sheet.set_cell_value("D3", 90.0).unwrap();

        let cf_rule = ConditionalFormatRule::cell_is_greater_than("50")
            .with_range(CellRange::parse("D1:D3").unwrap())
            .with_format(Style::new().fill_color(Color::rgb(0, 255, 0)));
        sheet.add_conditional_format(cf_rule);

        // Data validation: list in E1
        sheet.set_cell_value("E1", "Apple").unwrap();

        let dv = DataValidation::list("Apple,Banana,Cherry")
            .with_range(CellRange::parse("E1").unwrap())
            .with_input_message("Pick a fruit", "Choose from the list")
            .with_error_message("Invalid", "Must be a fruit");
        sheet.add_data_validation(dv);
    }

    // -- Write to file --
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("write xlsx");
    std::fs::write(&fixture.host_path, &buf)
        .unwrap_or_else(|e| panic!("write {}: {e}", fixture.host_path.display()));

    // -- Push to VM and open in Excel --
    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&fixture.vm_path)
        .expect("Excel should open our file without error");

    // -- Verify no repair --
    let wb_name = opened.name().expect("get workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel repaired the file! Workbook name: {wb_name}"
    );

    let read_only = opened.is_read_only().expect("get ReadOnly");
    assert!(
        !read_only,
        "Excel opened the file as read-only (possible repair)"
    );

    // -- Verify data survived --
    let val = opened.get_cell_value("A1").expect("get A1");
    assert_eq!(
        val,
        excel_com_protocol::CellValue::String("Hello".into()),
        "A1 string"
    );

    let val = opened.get_cell_value("A2").expect("get A2");
    match val {
        excel_com_protocol::CellValue::Number(n) => {
            assert!((n - 42.0).abs() < 0.001, "A2 should be 42.0, got {n}");
        }
        other => panic!("A2 should be Number, got {other:?}"),
    }

    let val = opened.get_cell_value("A3").expect("get A3");
    match val {
        excel_com_protocol::CellValue::Bool(true) => {}
        excel_com_protocol::CellValue::Number(n) if n == 1.0 => {}
        other => panic!("A3 should be true, got {other:?}"),
    }

    let val = opened.get_cell_value("A4").expect("get A4");
    match val {
        excel_com_protocol::CellValue::Number(n) => {
            assert!((n - 2.0).abs() < 0.001, "A4 =1+1 should be 2.0, got {n}");
        }
        other => panic!("A4 should be 2.0, got {other:?}"),
    }

    let val = opened.get_cell_value("A7").expect("get A7");
    assert_eq!(
        val,
        excel_com_protocol::CellValue::String("Merged region".into()),
        "A7 merged region"
    );

    opened.close().expect("close workbook");
    cleanup_fixture(&fixture);
}

#[test]
fn test_xlsb_roundtrip_no_repair() {
    let fixture = temp_fixture_xlsb();

    let mut wb = duke_sheets_core::Workbook::new();
    {
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Hello").unwrap();
        sheet.set_cell_value("A2", 42.0).unwrap();
        sheet.set_cell_value("A3", true).unwrap();
        sheet.set_cell_value("A4", CellValue::Number(2.0)).unwrap();
        sheet.set_cell_formula("A4", "=1+1").unwrap();

        let bold_style = Style::new()
            .bold(true)
            .font_color(Color::rgb(255, 0, 0))
            .fill_color(Color::rgb(255, 255, 0));
        sheet.set_cell_value("B1", "Styled").unwrap();
        sheet.set_cell_style("B1", &bold_style).unwrap();

        sheet.set_cell_value("C1", 0.75).unwrap();

        sheet.set_cell_value("A7", "Merged region").unwrap();
    }

    let mut buf = Vec::new();
    XlsbWriter::write(&wb, Cursor::new(&mut buf)).expect("XlsbWriter::write");
    std::fs::write(&fixture.host_path, &buf).expect("write XLSB fixture");

    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&fixture.vm_path)
        .expect("Excel should open our XLSB without error");

    let wb_name = opened.name().expect("get workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel repaired the XLSB file! Name: {wb_name}"
    );
    let read_only = opened.is_read_only().expect("get ReadOnly");
    assert!(
        !read_only,
        "Excel opened XLSB as read-only (possible repair)"
    );

    let val = opened.get_cell_value("A1").expect("get A1");
    assert_eq!(
        val,
        excel_com_protocol::CellValue::String("Hello".into()),
        "A1 string"
    );

    let val = opened.get_cell_value("A2").expect("get A2");
    match val {
        excel_com_protocol::CellValue::Number(n) => {
            assert!((n - 42.0).abs() < 0.001, "A2 should be 42.0, got {n}");
        }
        other => panic!("A2 should be Number, got {other:?}"),
    }

    let val = opened.get_cell_value("A3").expect("get A3");
    match val {
        excel_com_protocol::CellValue::Bool(true) => {}
        excel_com_protocol::CellValue::Number(n) if n == 1.0 => {}
        other => panic!("A3 should be true, got {other:?}"),
    }

    let val = opened.get_cell_value("A4").expect("get A4");
    match val {
        excel_com_protocol::CellValue::Number(n) => {
            assert!((n - 2.0).abs() < 0.001, "A4 =1+1 should be 2.0, got {n}");
        }
        other => panic!("A4 should be 2.0, got {other:?}"),
    }

    let val = opened.get_cell_value("A7").expect("get A7");
    assert_eq!(
        val,
        excel_com_protocol::CellValue::String("Merged region".into()),
        "A7 merged region"
    );

    opened.close().expect("close workbook");
    cleanup_fixture(&fixture);
}

#[test]
fn test_xlsb_feature_roundtrip() {
    use crate::roundtrip_through_excel_xlsb;
    use duke_sheets_core::worksheet::SheetVisibility;

    let mut wb = duke_sheets_core::Workbook::new();
    wb.add_worksheet_with_name("Sheet2").unwrap();
    wb.add_worksheet_with_name("Hidden").unwrap();
    wb.set_active_sheet(1).unwrap();

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "visible")
        .unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_tab_color(Some(Color::rgb(255, 0, 0)));
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", "s2")
        .unwrap();
    wb.worksheet_mut(2)
        .unwrap()
        .set_visibility(SheetVisibility::Hidden);
    wb.worksheet_mut(2)
        .unwrap()
        .set_cell_value("A1", "hidden")
        .unwrap();

    // Zoom
    wb.worksheet_mut(0).unwrap().set_zoom_scale(Some(150));

    // Selection
    wb.worksheet_mut(0).unwrap().set_selection_active_cell(2, 1);

    // Tab selected (sheet 1 is active, so sheet 0 should still be selectable)
    wb.worksheet_mut(1).unwrap().set_selected(true);

    // Outline levels
    wb.worksheet_mut(0).unwrap().set_row_outline_level(1, 1);
    wb.worksheet_mut(0).unwrap().set_column_outline_level(1, 1);

    wb.worksheet_mut(0).unwrap().set_protection(Some(
        duke_sheets_core::worksheet::SheetProtection {
            protected: true,
            password_hash: None,
            select_locked_cells: true,
            select_unlocked_cells: true,
            format_cells: false,
            format_columns: false,
            format_rows: false,
            insert_columns: false,
            insert_rows: false,
            insert_hyperlinks: false,
            delete_columns: false,
            delete_rows: false,
            sort: false,
            auto_filter: false,
            pivot_tables: false,
        },
    ));
    wb.worksheet_mut(0).unwrap().add_row_break(2);
    wb.worksheet_mut(0)
        .unwrap()
        .set_print_area(CellRange::parse("A1:B3").unwrap());

    let wb2 = roundtrip_through_excel_xlsb(&wb);

    assert_eq!(wb2.active_sheet(), 1, "active sheet");
    assert_eq!(
        wb2.worksheet(0).unwrap().visibility(),
        SheetVisibility::Visible
    );
    assert_eq!(
        wb2.worksheet(2).unwrap().visibility(),
        SheetVisibility::Hidden
    );

    let tab_color = wb2.worksheet(0).unwrap().tab_color();
    assert!(tab_color.is_some(), "tab color should survive round-trip");
    match tab_color.unwrap() {
        Color::Rgb { r, g, b } => assert_eq!((r, g, b), (255, 0, 0), "tab color RGB"),
        Color::Theme { .. } => {}
        other => panic!("unexpected tab color variant: {other:?}"),
    }

    assert_eq!(
        wb2.worksheet(0).unwrap().zoom_scale(),
        Some(150),
        "zoom scale"
    );
    assert!(
        wb2.worksheet(0).unwrap().protection().is_some(),
        "sheet protection"
    );
    assert!(
        !wb2.worksheet(0).unwrap().row_breaks().is_empty(),
        "page breaks"
    );
    assert!(
        wb2.worksheet(0).unwrap().print_area().is_some(),
        "print area"
    );
}

#[test]
fn test_xlsb_named_range_roundtrip() {
    use crate::roundtrip_through_excel_xlsb;
    use duke_sheets_core::named_range::NamedRange;

    let mut wb = duke_sheets_core::Workbook::new();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "header")
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 100.0)
        .unwrap();
    wb.named_ranges_mut()
        .define_or_update(NamedRange::workbook_scope("SalesTotal", "Data!$A$1:$D$10"));
    wb.named_ranges_mut()
        .define_or_update(NamedRange::sheet_scope("LocalVal", "Sheet1!$B$1", 0));

    let wb2 = roundtrip_through_excel_xlsb(&wb);
    let nr = wb2.named_ranges();
    assert!(nr.get("SalesTotal", 0).is_some(), "SalesTotal survives");
    assert!(nr.get("LocalVal", 0).is_some(), "LocalVal survives");
}

#[test]
fn test_xlsb_table_roundtrip() {
    use crate::roundtrip_through_excel_xlsb;
    use duke_sheets_core::table::{Table, TableColumn, TableStyleInfo};

    let mut wb = duke_sheets_core::Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Product").unwrap();
    ws.set_cell_value("B1", "Price").unwrap();
    ws.set_cell_value("A2", "Widget").unwrap();
    ws.set_cell_value("B2", 9.99).unwrap();
    ws.set_cell_value("A3", "Gadget").unwrap();
    ws.set_cell_value("B3", 19.99).unwrap();
    ws.add_table(Table {
        id: 1,
        name: "Products".to_string(),
        display_name: "Products".to_string(),
        reference: CellRange::parse("A1:B3").unwrap(),
        columns: vec![
            TableColumn {
                id: 1,
                name: "Product".to_string(),
                totals_row_function: None,
                totals_row_formula: None,
                totals_row_label: None,
                calculated_column_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Price".to_string(),
                totals_row_function: None,
                totals_row_formula: None,
                totals_row_label: None,
                calculated_column_formula: None,
            },
        ],
        style_info: Some(TableStyleInfo {
            name: Some("TableStyleMedium2".to_string()),
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }),
        header_row_count: 1,
        totals_row_count: 0,
        totals_row_shown: true,
    });

    let wb2 = roundtrip_through_excel_xlsb(&wb);
    let tables = wb2.worksheet(0).unwrap().tables();
    assert_eq!(tables.len(), 1, "table survives xlsb->excel->xlsb");
    assert_eq!(tables[0].name, "Products");
    assert_eq!(tables[0].columns.len(), 2);
    assert_eq!(tables[0].columns[0].name, "Product");
    assert_eq!(tables[0].columns[1].name, "Price");
}

#[test]
fn test_xlsb_cf_roundtrip() {
    use crate::roundtrip_through_excel_xlsb;
    use duke_sheets_core::conditional_format::{
        CfColorValue, CfRuleType, CfValueType, ConditionalFormatRule,
    };

    let mut wb = duke_sheets_core::Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for i in 0..5 {
        ws.set_cell_value_at(i, 0, (i as f64 + 1.0) * 20.0).unwrap();
    }

    let cs_rule = ConditionalFormatRule {
        rule_type: CfRuleType::ColorScale {
            colors: vec![
                CfColorValue::new(CfValueType::Min, None, Color::rgb(255, 0, 0)),
                CfColorValue::new(CfValueType::Max, None, Color::rgb(0, 255, 0)),
            ],
        },
        ranges: vec![CellRange::parse("A1:A5").unwrap()],
        priority: 2,
        stop_if_true: false,
        format: None,
        dxf_id: None,
    };
    ws.add_conditional_format(cs_rule);

    let cell_is_rule = ConditionalFormatRule::cell_is_greater_than("50")
        .with_range(CellRange::parse("A1:A5").unwrap())
        .with_format(Style::new().fill_color(Color::rgb(0, 255, 0)))
        .with_priority(1);
    ws.add_conditional_format(cell_is_rule);

    let wb2 = roundtrip_through_excel_xlsb(&wb);
    let ws2 = wb2.worksheet(0).unwrap();
    let rules = ws2.conditional_formats();

    assert!(
        rules.len() >= 2,
        "expected >=2 CF rules, got {}",
        rules.len()
    );
    let has_color_scale = rules
        .iter()
        .any(|r| matches!(&r.rule_type, CfRuleType::ColorScale { .. }));
    assert!(has_color_scale, "colorScale should survive round-trip");
    let has_cell_is = rules
        .iter()
        .any(|r| matches!(&r.rule_type, CfRuleType::CellIs { .. }));
    assert!(has_cell_is, "cellIs should survive round-trip");
}
