//! Round-trip tests for the XLS writer's _FilterDatabase NAME +
//! FILTERMODE + AUTOFILTER records.

use std::io::Cursor;

use duke_sheets_core::auto_filter::{
    AutoFilter, ColumnFilter, CustomFilterCondition, CustomFilters, FilterColumn, FilterOperator,
    Top10Filter, ValueFilter,
};
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
fn filter_range_only_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_auto_filter(Some(AutoFilter::new(range("A1", "C10"))));

    let parsed = write_then_read(&wb);
    let af = parsed
        .worksheet(0)
        .unwrap()
        .auto_filter()
        .cloned()
        .expect("auto-filter present after round-trip");
    assert_eq!(af.range.start, CellAddress::parse("A1").unwrap());
    assert_eq!(af.range.end, CellAddress::parse("C10").unwrap());
    assert!(af.filter_columns.is_empty());
}

#[test]
fn top10_filter_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut af = AutoFilter::new(range("A1", "C10"));
    af.filter_columns.push(FilterColumn::new(
        2,
        ColumnFilter::Top10(Top10Filter {
            top: true,
            percent: false,
            val: 5.0,
            filter_val: None,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let parsed = write_then_read(&wb);
    let af = parsed
        .worksheet(0)
        .unwrap()
        .auto_filter()
        .cloned()
        .expect("filter present");
    assert_eq!(af.filter_columns.len(), 1);
    assert_eq!(af.filter_columns[0].col_id, 2);
    match &af.filter_columns[0].filter {
        ColumnFilter::Top10(t) => {
            assert!(t.top);
            assert!(!t.percent);
            assert_eq!(t.val as i32, 5);
        }
        other => panic!("expected Top10, got {other:?}"),
    }
}

#[test]
fn top10_percent_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut af = AutoFilter::new(range("A1", "C10"));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Top10(Top10Filter {
            top: false,
            percent: true,
            val: 25.0,
            filter_val: None,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let parsed = write_then_read(&wb);
    let af = parsed
        .worksheet(0)
        .unwrap()
        .auto_filter()
        .cloned()
        .expect("filter present");
    match &af.filter_columns[0].filter {
        ColumnFilter::Top10(t) => {
            assert!(!t.top);
            assert!(t.percent);
            assert_eq!(t.val as i32, 25);
        }
        other => panic!("expected Top10 percent, got {other:?}"),
    }
}

#[test]
fn custom_numeric_filter_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut af = AutoFilter::new(range("A1", "C10"));
    af.filter_columns.push(FilterColumn::new(
        1,
        ColumnFilter::Custom(CustomFilters {
            and: false,
            conditions: vec![CustomFilterCondition {
                operator: FilterOperator::GreaterThan,
                value: "100".into(),
            }],
        }),
    ));
    ws.set_auto_filter(Some(af));

    let parsed = write_then_read(&wb);
    let af = parsed
        .worksheet(0)
        .unwrap()
        .auto_filter()
        .cloned()
        .expect("filter present");
    match &af.filter_columns[0].filter {
        ColumnFilter::Custom(c) => {
            assert_eq!(c.conditions.len(), 1);
            assert_eq!(c.conditions[0].operator, FilterOperator::GreaterThan);
            assert_eq!(c.conditions[0].value, "100");
        }
        other => panic!("expected Custom, got {other:?}"),
    }
}

#[test]
fn custom_dual_condition_and_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut af = AutoFilter::new(range("A1", "C10"));
    af.filter_columns.push(FilterColumn::new(
        1,
        ColumnFilter::Custom(CustomFilters {
            and: true,
            conditions: vec![
                CustomFilterCondition {
                    operator: FilterOperator::GreaterThanOrEqual,
                    value: "10".into(),
                },
                CustomFilterCondition {
                    operator: FilterOperator::LessThanOrEqual,
                    value: "100".into(),
                },
            ],
        }),
    ));
    ws.set_auto_filter(Some(af));

    let parsed = write_then_read(&wb);
    let af = parsed
        .worksheet(0)
        .unwrap()
        .auto_filter()
        .cloned()
        .expect("filter present");
    match &af.filter_columns[0].filter {
        ColumnFilter::Custom(c) => {
            assert!(c.and, "and join must round-trip");
            assert_eq!(c.conditions.len(), 2);
            assert_eq!(c.conditions[0].operator, FilterOperator::GreaterThanOrEqual);
            assert_eq!(c.conditions[1].operator, FilterOperator::LessThanOrEqual);
        }
        other => panic!("expected Custom AND, got {other:?}"),
    }
}

#[test]
fn value_filter_round_trips_as_equal_or() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut af = AutoFilter::new(range("A1", "C10"));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Values(ValueFilter {
            values: vec!["alpha".into(), "beta".into()],
            blank: false,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let parsed = write_then_read(&wb);
    let af = parsed
        .worksheet(0)
        .unwrap()
        .auto_filter()
        .cloned()
        .expect("filter present");
    match &af.filter_columns[0].filter {
        ColumnFilter::Values(v) => {
            assert_eq!(v.values.len(), 2);
            assert!(v.values.contains(&"alpha".to_string()));
            assert!(v.values.contains(&"beta".to_string()));
        }
        other => panic!("expected Values, got {other:?}"),
    }
}

#[test]
fn multiple_columns_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut af = AutoFilter::new(range("A1", "D10"));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Top10(Top10Filter {
            top: true,
            percent: false,
            val: 3.0,
            filter_val: None,
        }),
    ));
    af.filter_columns.push(FilterColumn::new(
        2,
        ColumnFilter::Custom(CustomFilters {
            and: false,
            conditions: vec![CustomFilterCondition {
                operator: FilterOperator::NotEqual,
                value: "0".into(),
            }],
        }),
    ));
    ws.set_auto_filter(Some(af));

    let parsed = write_then_read(&wb);
    let af = parsed
        .worksheet(0)
        .unwrap()
        .auto_filter()
        .cloned()
        .expect("filter present");
    assert_eq!(af.filter_columns.len(), 2);
    let col_ids: Vec<u32> = af.filter_columns.iter().map(|c| c.col_id).collect();
    assert!(col_ids.contains(&0));
    assert!(col_ids.contains(&2));
}

#[test]
fn no_autofilter_means_no_filtermode_record() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("A1");

    let parsed = write_then_read(&wb);
    assert!(parsed.worksheet(0).unwrap().auto_filter().is_none());
}
