//! Excel COM parity tests for the XLS (BIFF8) writer.
//!
//! Each test builds a workbook in memory, writes it via `XlsWriter`,
//! pushes to the Windows VM, opens in real Excel (asserting no
//! `Repaired` warning), re-saves to a second `.xls`, pulls it back,
//! and reads it with `XlsReader`. The `Repaired` check is the
//! canonical signal for writer bugs — Excel auto-recovers from
//! malformations our permissive reader silently tolerates.
//!
//! All tests are batched into one process to amortise the per-test
//! VM round-trip cost (~15-25s of warm-VM time per test).

use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, Top10Filter};
use duke_sheets_core::conditional_format::ConditionalFormatRule;
use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::Color;
use duke_sheets_core::validation::{DataValidation, ValidationOperator, ValidationType};
use duke_sheets_core::worksheet::{PageOrientation, SheetProtection, SheetVisibility};
use duke_sheets_core::{CellAddress, CellRange, CellValue, Hyperlink, Workbook};

use crate::roundtrip_through_excel_xls;

fn range(start: &str, end: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start).unwrap(),
        CellAddress::parse(end).unwrap(),
    )
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_hyperlinks_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "click").unwrap();
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "https://example.com".into(),
            display: Some("click".into()),
            tooltip: None,
            location: None,
        },
    )
    .unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("click"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_cross_sheet_formulas_we_emit() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();

    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 10.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 20.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 30.0)
        .unwrap();

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=Data!A1").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(10.0))
        .unwrap();
    calc.set_cell_formula("B2", "=SUM(Data!A1:A3)").unwrap();
    calc.set_formula_result(1, 1, CellValue::Number(60.0))
        .unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet_by_name("Calc").unwrap();
    let v1 = s.get_value_at(0, 1);
    match v1.effective_value() {
        CellValue::Number(n) => assert!((n - 10.0).abs() < 1e-9, "B1 = {n}"),
        other => panic!("B1 expected Number(10), got {other:?}"),
    }
    let v2 = s.get_value_at(1, 1);
    match v2.effective_value() {
        CellValue::Number(n) => assert!((n - 60.0).abs() < 1e-9, "B2 = {n}"),
        other => panic!("B2 expected Number(60), got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_autofilter_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Name").unwrap();
    ws.set_cell_value("B1", "Value").unwrap();
    ws.set_cell_value("A2", "alpha").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "beta").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let mut af = AutoFilter::new(range("A1", "B3"));
    af.filter_columns.push(FilterColumn::new(
        1,
        ColumnFilter::Top10(Top10Filter {
            top: true,
            percent: false,
            val: 1.0,
            filter_val: None,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("Name"));
    assert_eq!(s.get_value_at(0, 1).as_string(), Some("Value"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_data_validations_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "list").unwrap();
    ws.set_cell_value("B1", 50.0).unwrap();
    ws.set_cell_value("C1", 7.0).unwrap();

    let mut v_list = DataValidation::list("Red,Green,Blue");
    v_list.ranges = vec![range("A1", "A1")];
    ws.add_data_validation(v_list);

    let mut v_whole = DataValidation::whole_number_between(ValidationOperator::Between, "1", "100");
    v_whole.ranges = vec![range("B1", "B1")];
    ws.add_data_validation(v_whole);

    let mut v_custom = DataValidation::new();
    v_custom.validation_type = ValidationType::Custom {
        formula: "=ISNUMBER(C1)".into(),
    };
    v_custom.ranges = vec![range("C1", "C1")];
    ws.add_data_validation(v_custom);

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("list"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_conditional_formats_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 150.0).unwrap();
    ws.set_cell_value("A2", 50.0).unwrap();
    ws.set_cell_value("B1", 75.0).unwrap();

    let mut r1 = ConditionalFormatRule::cell_is_greater_than("100");
    r1.ranges = vec![range("A1", "A2")];
    ws.add_conditional_format(r1);

    let mut r2 = ConditionalFormatRule::cell_is_between("10", "100");
    r2.ranges = vec![range("B1", "B1")];
    ws.add_conditional_format(r2);

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 0);
    match v.effective_value() {
        CellValue::Number(n) => assert!((n - 150.0).abs() < 1e-9),
        other => panic!("A1 expected Number(150), got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_rich_text_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let bold = RunFont {
        bold: Some(true),
        ..Default::default()
    };
    let italic = RunFont {
        italic: Some(true),
        ..Default::default()
    };
    let red_big = RunFont {
        size: Some(16.0),
        color: Some(Color::Indexed(2)),
        ..Default::default()
    };
    ws.set_cell_value_at(
        0,
        0,
        CellValue::rich_text(vec![
            RichTextRun {
                text: "plain ".into(),
                font: None,
            },
            RichTextRun {
                text: "bold ".into(),
                font: Some(bold),
            },
            RichTextRun {
                text: "italic ".into(),
                font: Some(italic),
            },
            RichTextRun {
                text: "loud".into(),
                font: Some(red_big),
            },
        ]),
    )
    .unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    // Excel may round-trip the cell as either a plain `String` or a
    // `RichText`; `as_string()` only matches the former, so fall back
    // to the Display impl which concatenates rich runs.
    let plain = format!("{}", s.get_value_at(0, 0));
    assert_eq!(plain, "plain bold italic loud");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_named_range_formulas_we_emit() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 5.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 10.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 15.0)
        .unwrap();

    wb.define_name("Numbers", "Data!$A$1:$A$3").unwrap();
    wb.define_name("TaxRate", "0.1").unwrap();

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=SUM(Numbers)").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(30.0))
        .unwrap();
    calc.set_cell_formula("B2", "=B1*TaxRate").unwrap();
    calc.set_formula_result(1, 1, CellValue::Number(3.0))
        .unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet_by_name("Calc").unwrap();
    let v1 = s.get_value_at(0, 1);
    match v1.effective_value() {
        CellValue::Number(n) => assert!((n - 30.0).abs() < 1e-9, "B1 = {n}"),
        other => panic!("B1 expected Number(30), got {other:?}"),
    }
    let v2 = s.get_value_at(1, 1);
    match v2.effective_value() {
        CellValue::Number(n) => assert!((n - 3.0).abs() < 1e-9, "B2 = {n}"),
        other => panic!("B2 expected Number(3), got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_print_names_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "header").unwrap();
    ws.set_cell_value("B2", 100.0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.print_area = Some(range("A1", "B2"));
    ps.repeat_rows = Some((0, 0));
    ws.set_page_setup(ps);

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("header"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_visual_state_we_emit() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Public").unwrap();
    wb.add_worksheet_with_name("Hidden").unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "p")
        .unwrap();
    wb.worksheet_mut(0).unwrap().set_freeze_panes(1, 1);
    wb.worksheet_mut(0).unwrap().set_zoom_scale(Some(125));
    wb.worksheet_mut(1)
        .unwrap()
        .set_visibility(SheetVisibility::Hidden);
    wb.set_active_sheet(0).unwrap();

    let mut ps = wb.worksheet(0).unwrap().page_setup().clone();
    ps.orientation = PageOrientation::Landscape;
    ps.left_margin = 0.5;
    ps.odd_header = Some("Hdr".into());
    ps.print_gridlines = true;
    wb.worksheet_mut(0).unwrap().set_page_setup(ps);

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet_by_name("Public").unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("p"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_protection_state_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "locked").unwrap();
    ws.set_protection(Some(SheetProtection {
        protected: true,
        password_hash: Some(0xCAFE),
        ..Default::default()
    }));

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("locked"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_dimensions_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "tall").unwrap();
    ws.set_cell_value("B1", "wide").unwrap();
    ws.set_row_height(0, 36.0);
    ws.set_column_width(1, 25.0);

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("tall"));
    assert_eq!(s.get_value_at(0, 1).as_string(), Some("wide"));
}
