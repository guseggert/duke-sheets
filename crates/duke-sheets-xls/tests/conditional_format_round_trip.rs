//! Round-trip tests for the XLS writer's CONDFMT (0x01B0) and CF
//! (0x01B1) records.

use std::io::Cursor;

use duke_sheets_core::conditional_format::{CfOperator, CfRuleType, ConditionalFormatRule};
use duke_sheets_core::{CellAddress, CellRange, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn range(start: &str, end: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start).expect("start"),
        CellAddress::parse(end).expect("end"),
    )
}

#[test]
fn cellis_greater_than_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut rule = ConditionalFormatRule::cell_is_greater_than("100");
    rule.ranges = vec![range("A1", "A10")];
    ws.add_conditional_format(rule);

    let parsed = write_then_read(&wb);
    let rules = parsed.worksheet(0).unwrap().conditional_formats();
    assert_eq!(rules.len(), 1);
    match &rules[0].rule_type {
        CfRuleType::CellIs {
            operator, formula1, ..
        } => {
            assert_eq!(*operator, CfOperator::GreaterThan);
            assert_eq!(formula1, "100");
        }
        other => panic!("expected CellIs>, got {other:?}"),
    }
    assert_eq!(rules[0].ranges, vec![range("A1", "A10")]);
}

#[test]
fn cellis_between_two_values_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut rule = ConditionalFormatRule::cell_is_between("10", "100");
    rule.ranges = vec![range("B1", "B5")];
    ws.add_conditional_format(rule);

    let parsed = write_then_read(&wb);
    let rules = parsed.worksheet(0).unwrap().conditional_formats();
    match &rules[0].rule_type {
        CfRuleType::CellIs {
            operator,
            formula1,
            formula2,
        } => {
            assert_eq!(*operator, CfOperator::Between);
            assert_eq!(formula1, "10");
            assert_eq!(formula2.as_deref(), Some("100"));
        }
        other => panic!("expected CellIs Between, got {other:?}"),
    }
}

#[test]
fn cellis_equal_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut rule = ConditionalFormatRule::cell_is_equal_to("0");
    rule.ranges = vec![range("C1", "C3")];
    ws.add_conditional_format(rule);

    let parsed = write_then_read(&wb);
    let rules = parsed.worksheet(0).unwrap().conditional_formats();
    match &rules[0].rule_type {
        CfRuleType::CellIs {
            operator, formula1, ..
        } => {
            assert_eq!(*operator, CfOperator::Equal);
            assert_eq!(formula1, "0");
        }
        other => panic!("expected CellIs =, got {other:?}"),
    }
}

#[test]
fn expression_rule_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut rule = ConditionalFormatRule::new(CfRuleType::Expression {
        formula: "=A1>100".into(),
    });
    rule.ranges = vec![range("A1", "A5")];
    ws.add_conditional_format(rule);

    let parsed = write_then_read(&wb);
    let rules = parsed.worksheet(0).unwrap().conditional_formats();
    match &rules[0].rule_type {
        CfRuleType::Expression { formula } => {
            assert!(formula.contains("A1"), "got {formula:?}");
            assert!(formula.contains('>'), "got {formula:?}");
            assert!(formula.contains("100"), "got {formula:?}");
        }
        other => panic!("expected Expression, got {other:?}"),
    }
}

#[test]
fn multiple_rules_per_sheet_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut r1 = ConditionalFormatRule::cell_is_greater_than("0");
    r1.ranges = vec![range("A1", "A10")];
    let mut r2 = ConditionalFormatRule::cell_is_less_than("0");
    r2.ranges = vec![range("B1", "B10")];
    let mut r3 = ConditionalFormatRule::new(CfRuleType::Expression {
        formula: "=C1=\"flagged\"".into(),
    });
    r3.ranges = vec![range("C1", "C10")];
    ws.add_conditional_format(r1);
    ws.add_conditional_format(r2);
    ws.add_conditional_format(r3);

    let parsed = write_then_read(&wb);
    let rules = parsed.worksheet(0).unwrap().conditional_formats();
    assert_eq!(rules.len(), 3);
}

#[test]
fn no_rules_means_no_condfmt_records() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("A1");

    let parsed = write_then_read(&wb);
    assert!(parsed
        .worksheet(0)
        .unwrap()
        .conditional_formats()
        .is_empty());
}

#[test]
fn rule_with_multiple_ranges_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut rule = ConditionalFormatRule::cell_is_greater_than("50");
    rule.ranges = vec![range("A1", "A5"), range("C1", "C5"), range("E1", "E5")];
    ws.add_conditional_format(rule);

    let parsed = write_then_read(&wb);
    let rules = parsed.worksheet(0).unwrap().conditional_formats();
    assert_eq!(rules.len(), 1);
    assert_eq!(rules[0].ranges.len(), 3);
}
