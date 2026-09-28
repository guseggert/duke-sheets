//! Excel COM parity tests for XLSX pivot table writing.

use crate::roundtrip_through_excel;
use duke_sheets_core::{
    CellRange, PivotAggregate, PivotDateGroupUnit, PivotField, PivotFieldRef, PivotFilter,
    PivotGrouping, PivotManualGroup, PivotMeasure, PivotSort, PivotSource, PivotSourceRange,
    PivotStyle, PivotValue, Workbook,
};

fn basic_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("BasicPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn sort_by_measure_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Q1").unwrap();
    sheet.set_cell_value("C3", 50.0).unwrap();

    let mut region =
        PivotField::new("Region").with_sort_by(PivotFieldRef::new("Revenue"), PivotAggregate::Sum);
    region.sort = PivotSort::Descending;
    let pivot = duke_sheets_core::PivotTable::builder("ValueSortedPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Sum of Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn styled_pivot_style() -> PivotStyle {
    PivotStyle {
        name: Some("PivotStyleLight16".to_string()),
        show_row_headers: false,
        show_column_headers: true,
        show_row_stripes: true,
        show_column_stripes: true,
        show_last_column: true,
    }
}

fn styled_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("StyledPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .style(styled_pivot_style())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn top_n_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", 30.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("TopRegions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::TopN {
            field: PivotFieldRef::new("Region"),
            measure: PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue"),
            n: 2,
            top: true,
            percent: false,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn named_consolidation_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    {
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();
    }
    wb.define_name("NamedSource", "Sheet1!$A$1:$B$3").unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("NamedConsolidation")
        .source(PivotSource::Consolidation {
            ranges: vec![PivotSourceRange::named("NamedSource")],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

// features: Pivot cache (source data); Pivot table definition; Row / column / value fields; Filter (page) fields; Aggregate functions (Sum/Count/Avg/...)
#[test]
fn test_write_basic_pivot_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&basic_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
}

#[test]
fn test_write_pivot_sort_by_measure_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&sort_by_measure_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ValueSortedPivot")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.rows[0].sort, PivotSort::Descending);
    let measure = pivot.rows[0]
        .sort_by_measure
        .as_ref()
        .expect("sort-by-measure survives Excel roundtrip");
    assert_eq!(measure.field.name, "Revenue");
    assert_eq!(measure.aggregate, PivotAggregate::Sum);
    assert_eq!(measure.name.as_deref(), Some("Sum of Revenue"));
}

// features: Pivot table styles
#[test]
fn test_write_pivot_style_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&styled_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("StyledPivot")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.style, styled_pivot_style());
}

#[test]
fn test_write_pivot_top_n_filter_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&top_n_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TopRegions")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.filters.len(), 1);
    match &pivot.filters[0] {
        PivotFilter::TopN {
            field,
            measure,
            n,
            top,
            percent,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(measure.field.name, "Revenue");
            assert_eq!(measure.aggregate, PivotAggregate::Sum);
            assert_eq!(measure.name.as_deref(), Some("Total Revenue"));
            assert_eq!(*n, 2);
            assert!(*top);
            assert!(!*percent);
        }
        other => panic!("unexpected pivot filter after Excel roundtrip: {other:?}"),
    }
}

#[test]
fn test_write_named_consolidation_pivot_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&named_consolidation_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("NamedConsolidation")
        .expect("pivot survives Excel roundtrip");

    match &pivot.source {
        PivotSource::Consolidation { ranges } => {
            assert_eq!(ranges.len(), 1);
            assert_eq!(ranges[0].name.as_deref(), Some("NamedSource"));
        }
        other => panic!("unexpected pivot source after Excel roundtrip: {other:?}"),
    }
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

// features: Grouping (dates, numbers, items)
#[test]
fn test_write_manual_pivot_grouping_survives_excel_roundtrip() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "Central").unwrap();
    sheet.set_cell_value("B4", 5.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("ManualGroupedRegions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let result = roundtrip_through_excel(&wb);
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualGroupedRegions")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Manual { field, groups } => {
            assert_eq!(field.name, "Region");
            assert_eq!(groups.len(), 1);
            assert_eq!(groups[0].name, "Coastal");
            assert_eq!(
                groups[0].members,
                vec![
                    PivotValue::String("East".to_string()),
                    PivotValue::String("West".to_string())
                ]
            );
        }
        other => panic!("unexpected grouping after Excel roundtrip: {other:?}"),
    }
}

// features: Grouping (dates, numbers, items)
#[test]
fn test_write_numeric_pivot_grouping_survives_excel_roundtrip() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    for (address, value) in [("A1", "Amount"), ("B1", "Revenue")] {
        sheet.set_cell_value(address, value).unwrap();
    }
    for (row, (amount, revenue)) in [(2.0, 20.0), (12.0, 120.0), (22.0, 220.0)]
        .into_iter()
        .enumerate()
    {
        let row = row + 2;
        sheet.set_cell_value(&format!("A{row}"), amount).unwrap();
        sheet.set_cell_value(&format!("B{row}"), revenue).unwrap();
    }

    let pivot = duke_sheets_core::PivotTable::builder("NumericGroupedSales")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Amount")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Number {
            field: "Amount".into(),
            start: None,
            end: None,
            interval: 10.0,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let result = roundtrip_through_excel(&wb);
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("NumericGroupedSales")
        .expect("numeric grouping survives Excel roundtrip");
    assert!(pivot.groupings.iter().any(|grouping| matches!(
        grouping,
        PivotGrouping::Number { field, start, end, interval }
            if field.name == "Amount"
                && *start == Some(2.0)
                && *end == Some(22.0)
                && *interval == 10.0
    )));
}

// features: Grouping (dates, numbers, items)
#[test]
fn test_write_single_date_pivot_grouping_survives_excel_roundtrip() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "SaleDate").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    for (row, (date, revenue)) in [(45292.0, 10.0), (45323.0, 20.0), (45352.0, 30.0)]
        .into_iter()
        .enumerate()
    {
        let row = row + 2;
        sheet.set_cell_value(&format!("A{row}"), date).unwrap();
        sheet.set_cell_value(&format!("B{row}"), revenue).unwrap();
    }

    let pivot = duke_sheets_core::PivotTable::builder("DateGroupedSales")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("SaleDate")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Date {
            field: "SaleDate".into(),
            units: vec![PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let result = roundtrip_through_excel(&wb);
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("DateGroupedSales")
        .expect("date grouping survives Excel roundtrip");
    assert!(pivot.groupings.iter().any(|grouping| matches!(
        grouping,
        PivotGrouping::Date { field, units }
            if field.name == "SaleDate" && units == &[PivotDateGroupUnit::Months]
    )));
}

// features: Grouping (dates, numbers, items)
#[test]
fn test_write_multi_date_pivot_grouping_survives_excel_roundtrip() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "SaleDate").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    for (row, (date, revenue)) in [(45292.0, 10.0), (45323.0, 20.0), (45658.0, 30.0)]
        .into_iter()
        .enumerate()
    {
        let row = row + 2;
        sheet.set_cell_value(&format!("A{row}"), date).unwrap();
        sheet.set_cell_value(&format!("B{row}"), revenue).unwrap();
    }

    let pivot = duke_sheets_core::PivotTable::builder("MultiDateGroupedSales")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("SaleDate")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Date {
            field: "SaleDate".into(),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let result = roundtrip_through_excel(&wb);
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("MultiDateGroupedSales")
        .expect("multi-unit date grouping survives Excel roundtrip");
    assert!(pivot.groupings.iter().any(|grouping| matches!(
        grouping,
        PivotGrouping::Date { field, units }
            if field.name == "SaleDate"
                && units == &[PivotDateGroupUnit::Years, PivotDateGroupUnit::Months]
    )));
}

// features: Grouping (dates, numbers, items)
#[test]
fn test_write_range_pivot_grouping_survives_excel_roundtrip() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    for (address, value) in [("A1", "Amount"), ("B1", "SaleDate"), ("C1", "Revenue")] {
        sheet.set_cell_value(address, value).unwrap();
    }
    for (row, (amount, date, revenue)) in [
        (2.0, 45292.0, 10.0),
        (12.0, 45323.0, 20.0),
        (22.0, 45352.0, 30.0),
    ]
    .into_iter()
    .enumerate()
    {
        let row = row + 2;
        sheet.set_cell_value(&format!("A{row}"), amount).unwrap();
        sheet.set_cell_value(&format!("B{row}"), date).unwrap();
        sheet.set_cell_value(&format!("C{row}"), revenue).unwrap();
    }

    let pivot = duke_sheets_core::PivotTable::builder("RangeGroupedSales")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Amount")
        .column("SaleDate")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Number {
            field: "Amount".into(),
            start: Some(0.0),
            end: Some(30.0),
            interval: 10.0,
        })
        .grouping(PivotGrouping::Date {
            field: "SaleDate".into(),
            units: vec![PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let result = roundtrip_through_excel(&wb);
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RangeGroupedSales")
        .expect("range groupings survive Excel roundtrip");
    assert!(pivot.groupings.iter().any(|grouping| matches!(
        grouping,
        PivotGrouping::Number { field, start, end, interval }
            if field.name == "Amount"
                && *start == Some(0.0)
                && *end == Some(30.0)
                && *interval == 10.0
    )));
    assert!(pivot.groupings.iter().any(|grouping| matches!(
        grouping,
        PivotGrouping::Date { field, units }
            if field.name == "SaleDate" && units == &[PivotDateGroupUnit::Months]
    )));
}
