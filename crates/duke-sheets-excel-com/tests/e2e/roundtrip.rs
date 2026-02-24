//! Roundtrip fidelity test: write XLSX with duke-sheets, open in real Excel,
//! verify no repair warning, verify data survives.
//!
//! This is the ultimate compatibility test — if Excel can open our file
//! without repairing it, the OOXML we produce is spec-compliant.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, push_file_to_vm, temp_fixture};
use duke_sheets_core::{
    BorderLineStyle, CellRange, CellValue, Color, ConditionalFormatRule, DataValidation,
    HorizontalAlignment, NumberFormat, Style, VerticalAlignment,
};
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
        sheet
            .set_cell_value(
                "A4",
                CellValue::Formula {
                    text: "=1+1".into(),
                    cached_value: Some(Box::new(CellValue::Number(2.0))),
                    array_result: None,
                },
            )
            .unwrap();

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

        // NOTE: Comments omitted — our writer produces comments1.xml but not
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
