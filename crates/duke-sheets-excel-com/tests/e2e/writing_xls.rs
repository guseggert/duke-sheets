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
    let hl = s
        .hyperlink("A1")
        .expect("hyperlink must survive Excel round-trip");
    assert!(
        hl.target.contains("example.com"),
        "hyperlink target lost after round-trip: {:?}",
        hl.target
    );
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
    let af = s
        .auto_filter()
        .expect("autofilter must survive Excel round-trip");
    assert_eq!(
        af.range.start,
        CellAddress::parse("A1").unwrap(),
        "autofilter range start lost"
    );
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
    let validations = s.data_validations();
    assert!(
        !validations.is_empty(),
        "data validations must survive Excel round-trip"
    );
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
    let rules = s.conditional_formats();
    assert!(
        !rules.is_empty(),
        "conditional format rules must survive Excel round-trip"
    );
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
    let value = s.get_value_at(0, 0);

    // Concatenated text must round-trip exactly.
    assert_eq!(format!("{value}"), "plain bold italic loud");

    // The cell must come back as RichText with at least the bold,
    // italic, and size+color runs distinguishable from the plain run.
    // If Excel collapsed the runs into a single span the formatting
    // is lost — the writer's RichText emission would be effectively
    // ineffective even though the text reads back correctly.
    let runs = match &value {
        CellValue::RichText(runs) => runs,
        other => panic!("expected RichText after Excel round-trip, got {other:?}"),
    };
    assert!(
        runs.len() >= 4,
        "expected at least 4 runs (plain/bold/italic/loud), got {}: {runs:?}",
        runs.len()
    );

    let has_bold = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if f.bold == Some(true)));
    let has_italic = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if f.italic == Some(true)));
    let has_big = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if matches!(f.size, Some(s) if s >= 14.0)));
    assert!(
        has_bold,
        "expected at least one run with bold=true: {runs:?}"
    );
    assert!(
        has_italic,
        "expected at least one run with italic=true: {runs:?}"
    );
    assert!(has_big, "expected at least one run with size>=14: {runs:?}");
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
    let print_area = s
        .page_setup()
        .print_area
        .as_ref()
        .expect("print_area must survive Excel round-trip");
    assert_eq!(print_area.start, CellAddress::parse("A1").unwrap());
    assert_eq!(print_area.end, CellAddress::parse("B2").unwrap());
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
    let freeze = s.freeze_panes().expect("freeze panes must survive");
    assert_eq!(
        (freeze.row, freeze.col),
        (1, 1),
        "freeze panes lost after round-trip"
    );
    let hidden = result.worksheet_by_name("Hidden").unwrap();
    assert!(
        hidden.visibility() != SheetVisibility::Visible,
        "Hidden sheet must remain non-visible after round-trip; got {:?}",
        hidden.visibility()
    );
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
    let prot = s
        .protection()
        .expect("sheet protection must survive Excel round-trip");
    assert!(prot.protected, "protected flag lost after round-trip");
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
    // Row 0 should retain a non-default height. Excel may round or
    // adjust slightly, so allow ±1 pt slack.
    let h = s.row_height(0);
    assert!((h - 36.0).abs() < 1.0, "row 0 height = {h} (expected ~36)",);
    // Column B (index 1) should retain a non-default width.
    let w = s.column_width(1);
    assert!(w > 20.0 && w < 30.0, "col B width = {w} (expected ~25)",);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_intersection_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, &(a, b, c)) in [(1, 2, 3), (4, 5, 6), (7, 8, 9)].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 0, a as f64).unwrap();
        ws.set_cell_value_at(row as u32, 1, b as f64).unwrap();
        ws.set_cell_value_at(row as u32, 2, c as f64).unwrap();
    }
    ws.set_cell_formula("E1", "=SUM(A1:B3 B2:C3)").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(13.0))
        .unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 4);
    match v.effective_value() {
        CellValue::Number(n) => assert!(
            (n - 13.0).abs() < 1e-9,
            "intersection sum drifted: E1 = {n}"
        ),
        other => panic!("E1 expected Number(13), got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_union_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, &(a, b, c)) in [(1, 2, 3), (4, 5, 6), (7, 8, 9)].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 0, a as f64).unwrap();
        ws.set_cell_value_at(row as u32, 1, b as f64).unwrap();
        ws.set_cell_value_at(row as u32, 2, c as f64).unwrap();
    }
    ws.set_cell_formula("E1", "=SUM((A1:A2,C2:C3))").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(20.0))
        .unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 4);
    match v.effective_value() {
        CellValue::Number(n) => assert!((n - 20.0).abs() < 1e-9, "union sum drifted: E1 = {n}"),
        other => panic!("E1 expected Number(20), got {other:?}"),
    }
}

/// Excel must accept our XLS comment-shape emit (MSODRAWINGGROUP +
/// MSODRAWING + OBJ + TXO + NOTE) and round-trip the comment text +
/// author + anchor cell through SaveAs. The roundtrip helper asserts
/// no `Repaired` warning fires, which is the canonical signal that
/// our Escher record output is spec-compliant.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_comment_we_emit() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.set_comment_at(0, 0, CellComment::new("Alice", "Hello from duke-sheets"));

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    let c = s
        .comment_at(0, 0)
        .expect("comment must survive Excel re-save");
    assert!(
        c.text.contains("Hello from duke-sheets"),
        "comment text lost after Excel round-trip: {:?}",
        c.text
    );
}

/// Multi-comment scenario: Excel must accept multiple SP_CONTAINERs
/// inside one DG_CONTAINER and preserve each comment's text +
/// author + anchor cell. Catches off-by-one bugs in shape ID
/// allocation, OBJ.ftCmo.id, and NOTE.objId linking.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_multiple_comments_we_emit() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "first").unwrap();
    ws.set_comment_at(0, 0, CellComment::new("Alice", "Comment one"));
    ws.set_cell_value("C3", "second").unwrap();
    ws.set_comment_at(2, 2, CellComment::new("Bob", "Comment two body"));
    ws.set_cell_value("E5", "third").unwrap();
    ws.set_comment_at(4, 4, CellComment::new("Carol", "Comment three"));

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.comment_count(), 3, "all three comments must survive");
    assert!(
        s.comment_at(0, 0).unwrap().text.contains("Comment one"),
        "A1 comment lost: {:?}",
        s.comment_at(0, 0).unwrap().text
    );
    assert!(
        s.comment_at(2, 2)
            .unwrap()
            .text
            .contains("Comment two body"),
        "C3 comment lost: {:?}",
        s.comment_at(2, 2).unwrap().text
    );
    assert!(
        s.comment_at(4, 4).unwrap().text.contains("Comment three"),
        "E5 comment lost: {:?}",
        s.comment_at(4, 4).unwrap().text
    );
}

/// Unicode comment text — drives the writer onto the UTF-16LE TXO
/// CONTINUE path. Confirms Excel preserves non-Latin glyphs through
/// the round-trip.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_unicode_comment_we_emit() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "jp").unwrap();
    ws.set_comment_at(0, 0, CellComment::new("作者", "こんにちは"));

    let result = roundtrip_through_excel_xls(&wb);
    let c = result.worksheet(0).unwrap().comment_at(0, 0).unwrap();
    assert!(
        c.text.contains("こんにちは"),
        "Japanese text lost: {:?}",
        c.text
    );
}

/// Multi-sheet comments: exercises the per-drawing 1024-aligned
/// shape-ID cluster allocation. A workbook with comments on
/// non-contiguous sheets must produce one FDGG cluster entry per
/// drawing and unique shape IDs across the workbook.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_comments_across_multiple_sheets_we_emit() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Alpha").unwrap();
    wb.add_worksheet_with_name("Beta").unwrap();
    wb.add_worksheet_with_name("Gamma").unwrap();

    // Alpha gets two comments; Beta is empty; Gamma gets one. This
    // forces the writer to allocate two clusters (dgid 1 and 2)
    // because Beta has no shapes to take a cluster slot.
    wb.worksheet_mut(0)
        .unwrap()
        .set_comment_at(0, 0, CellComment::new("Alice", "Alpha A1"));
    wb.worksheet_mut(0)
        .unwrap()
        .set_comment_at(3, 3, CellComment::new("Alice", "Alpha D4"));
    wb.worksheet_mut(2)
        .unwrap()
        .set_comment_at(5, 5, CellComment::new("Carol", "Gamma F6"));

    let result = roundtrip_through_excel_xls(&wb);
    assert_eq!(result.worksheet(0).unwrap().comment_count(), 2);
    assert_eq!(result.worksheet(1).unwrap().comment_count(), 0);
    assert_eq!(result.worksheet(2).unwrap().comment_count(), 1);
    assert!(result
        .worksheet(0)
        .unwrap()
        .comment_at(0, 0)
        .unwrap()
        .text
        .contains("Alpha A1"));
    assert!(result
        .worksheet(0)
        .unwrap()
        .comment_at(3, 3)
        .unwrap()
        .text
        .contains("Alpha D4"));
    assert!(result
        .worksheet(2)
        .unwrap()
        .comment_at(5, 5)
        .unwrap()
        .text
        .contains("Gamma F6"));
}

/// A 68-byte 1x1 transparent PNG with verified chunk CRCs, used as
/// the deterministic image payload for picture parity tests.
const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82,
];

/// Excel must accept our XLS picture emit (MSODRAWINGGROUP with
/// BSTORE_CONTAINER + per-sheet MSODRAWING with picture
/// SP_CONTAINER + picture OBJ) without triggering the Repaired
/// warning, and the embedded PNG bytes must survive Excel's
/// SaveAs round-trip verbatim.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_xls_png_image_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(EmbeddedImage {
        id: 1,
        name: "Picture 1".into(),
        description: None,
        anchor: DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 2,
                col_offset_emu: 0,
                row: 3,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 5,
                col_offset_emu: 0,
                row: 8,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
        format: ImageFormat::Png,
        media_path: String::new(),
        svg_media_path: None,
        width_emu: 1_000_000,
        height_emu: 1_000_000,
        rotation: None,
        flip_h: false,
        flip_v: false,
        data: TEST_PNG_1X1.to_vec(),
        svg_data: None,
    });

    let result = roundtrip_through_excel_xls(&wb);
    let images = result.worksheet(0).unwrap().images();
    assert_eq!(images.len(), 1, "image must survive Excel re-save");
    let img = &images[0];
    assert_eq!(img.format, ImageFormat::Png);
    assert_eq!(
        img.data, TEST_PNG_1X1,
        "PNG bytes must round-trip through Excel verbatim"
    );
    match &img.anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            // Excel may adjust within-cell EMU offsets when it
            // re-saves; we only assert the cell *range* is preserved
            // (top-left and bottom-right cell indices).
            assert_eq!(from.col, 2, "from col must survive");
            assert_eq!(from.row, 3, "from row must survive");
            assert_eq!(to.col, 5, "to col must survive");
            assert_eq!(to.row, 8, "to row must survive");
        }
        other => panic!("expected TwoCell anchor after round-trip, got {other:?}"),
    }
}
