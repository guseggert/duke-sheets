use crate::test_support::*;
use pretty_assertions::assert_eq;

#[test]
fn refresh_applies_item_filters() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", 15.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::field_items("Region", ["East"]))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 25.0);
    assert_eq!(text(&workbook, "D3"), "Grand Total");
    assert_eq!(number(&workbook, "E3"), 25.0);
    assert_eq!(text(&workbook, "D4"), "");
}

#[test]
fn refresh_includes_new_items_in_existing_filters() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let mut region = PivotField::new("Region");
    region.include_new_items_in_filter = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row(region)
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::field_items("Region", ["East"]))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 10.0);
    assert_eq!(text(&workbook, "D3"), "Grand Total");
    assert_eq!(number(&workbook, "E3"), 10.0);

    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", 8.0).unwrap();
    sheet.pivot_tables_mut()[0].source = PivotSource::range(CellRange::parse("A1:B4").unwrap());

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 10.0);
    assert_eq!(text(&workbook, "D3"), "North");
    assert_eq!(number(&workbook, "E3"), 8.0);
    assert_eq!(text(&workbook, "D4"), "Grand Total");
    assert_eq!(number(&workbook, "E4"), 18.0);
    assert_eq!(text(&workbook, "D5"), "");
}

#[test]
fn refresh_applies_value_filters() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", 15.0).unwrap();
    sheet.set_cell_value("A5", "North").unwrap();
    sheet.set_cell_value("B5", 5.0).unwrap();

    let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .pivot_measure(measure.clone())
        .filter(PivotFilter::Value {
            field: "Region".into(),
            measure,
            operator: PivotFilterOperator::GreaterThanOrEqual,
            value: 20.0,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 25.0);
    assert_eq!(text(&workbook, "D3"), "West");
    assert_eq!(number(&workbook, "E3"), 20.0);
    assert_eq!(text(&workbook, "D4"), "Grand Total");
    assert_eq!(number(&workbook, "E4"), 45.0);
}

#[test]
fn show_empty_items_respects_value_filters() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 1.0).unwrap();

    let mut region = PivotField::new("Region");
    region.show_empty_items = true;
    let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row(region)
        .pivot_measure(measure.clone())
        .filter(PivotFilter::Value {
            field: "Region".into(),
            measure,
            operator: PivotFilterOperator::GreaterThanOrEqual,
            value: 5.0,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 10.0);
    assert_eq!(text(&workbook, "D3"), "Grand Total");
    assert_eq!(number(&workbook, "E3"), 10.0);
    assert_eq!(text(&workbook, "D4"), "");
}

#[test]
fn refresh_applies_range_filters() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", 15.0).unwrap();
    sheet.set_cell_value("A5", "North").unwrap();
    sheet.set_cell_value("B5", 5.0).unwrap();

    let label_range = PivotTable::builder("LabelRange")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::LabelBetween {
            field: "Region".into(),
            start: "East".into(),
            end: "North".into(),
            not_between: false,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(label_range).unwrap();

    let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
    let value_range = PivotTable::builder("ValueRange")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("G1")
        .unwrap()
        .row("Region")
        .pivot_measure(measure.clone())
        .filter(PivotFilter::ValueBetween {
            field: "Region".into(),
            measure,
            start: 10.0,
            end: 25.0,
            not_between: false,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(value_range).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 25.0);
    assert_eq!(text(&workbook, "D3"), "North");
    assert_eq!(number(&workbook, "E3"), 5.0);
    assert_eq!(text(&workbook, "D4"), "Grand Total");
    assert_eq!(number(&workbook, "E4"), 30.0);

    assert_eq!(text(&workbook, "G2"), "East");
    assert_eq!(number(&workbook, "H2"), 25.0);
    assert_eq!(text(&workbook, "G3"), "West");
    assert_eq!(number(&workbook, "H3"), 20.0);
    assert_eq!(text(&workbook, "G4"), "Grand Total");
    assert_eq!(number(&workbook, "H4"), 45.0);
}

#[test]
fn refresh_applies_date_filters() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Date").unwrap();
    sheet.set_cell_value("B1", "Region").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet
        .set_cell_value("A2", date_to_serial(2024, 1, 1, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B2", "East").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet
        .set_cell_value("A3", date_to_serial(2024, 1, 15, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B3", "East").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet
        .set_cell_value("A4", date_to_serial(2024, 1, 20, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B4", "West").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();
    sheet
        .set_cell_value("A5", date_to_serial(2024, 2, 1, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B5", "North").unwrap();
    sheet.set_cell_value("C5", 30.0).unwrap();
    sheet
        .set_cell_value("A6", date_to_serial(2024, 2, 15, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B6", "West").unwrap();
    sheet.set_cell_value("C6", 40.0).unwrap();

    let after_february = PivotTable::builder("AfterFebruary")
        .source_range(CellRange::parse("A1:C6").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::Date {
            field: "Date".into(),
            operator: PivotFilterOperator::GreaterThanOrEqual,
            value: date_to_serial(2024, 2, 1, DateSystem::Date1900),
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(after_february).unwrap();

    let january_window = PivotTable::builder("JanuaryWindow")
        .source_range(CellRange::parse("A1:C6").unwrap())
        .target_address("H1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::DateBetween {
            field: "Date".into(),
            start: date_to_serial(2024, 1, 10, DateSystem::Date1900),
            end: date_to_serial(2024, 1, 31, DateSystem::Date1900),
            not_between: false,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(january_window).unwrap();

    let this_month = PivotTable::builder("ThisMonth")
        .source_range(CellRange::parse("A1:C6").unwrap())
        .target_address("K1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::DatePeriod {
            field: "Date".into(),
            period: PivotDatePeriod::ThisMonth,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(this_month).unwrap();

    workbook
        .refresh_pivots_with_options(&PivotRefreshOptions {
            today: Some(date_to_serial(2024, 2, 10, DateSystem::Date1900)),
            ..PivotRefreshOptions::default()
        })
        .unwrap();

    assert_eq!(text(&workbook, "E2"), "North");
    assert_eq!(number(&workbook, "F2"), 30.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(number(&workbook, "F3"), 40.0);
    assert_eq!(text(&workbook, "E4"), "Grand Total");
    assert_eq!(number(&workbook, "F4"), 70.0);

    assert_eq!(text(&workbook, "H2"), "East");
    assert_eq!(number(&workbook, "I2"), 20.0);
    assert_eq!(text(&workbook, "H3"), "West");
    assert_eq!(number(&workbook, "I3"), 7.0);
    assert_eq!(text(&workbook, "H4"), "Grand Total");
    assert_eq!(number(&workbook, "I4"), 27.0);

    assert_eq!(text(&workbook, "K2"), "North");
    assert_eq!(number(&workbook, "L2"), 30.0);
    assert_eq!(text(&workbook, "K3"), "West");
    assert_eq!(number(&workbook, "L3"), 40.0);
    assert_eq!(text(&workbook, "K4"), "Grand Total");
    assert_eq!(number(&workbook, "L4"), 70.0);
}

#[test]
fn relative_date_period_filters_require_today() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Date").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet
        .set_cell_value("A2", date_to_serial(2024, 2, 10, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();

    let pivot = PivotTable::builder("TodaySales")
        .source_range(CellRange::parse("A1:B2").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Date")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::DatePeriod {
            field: "Date".into(),
            period: PivotDatePeriod::Today,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let error = workbook.refresh_pivots().unwrap_err().to_string();

    assert!(error.contains("refresh options did not provide today"));
}

#[test]
fn refresh_applies_top_n_filters() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", 15.0).unwrap();
    sheet.set_cell_value("A5", "North").unwrap();
    sheet.set_cell_value("B5", 5.0).unwrap();

    let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .pivot_measure(measure.clone())
        .filter(PivotFilter::TopN {
            field: "Region".into(),
            measure,
            n: 2,
            top: true,
            percent: false,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 25.0);
    assert_eq!(text(&workbook, "D3"), "West");
    assert_eq!(number(&workbook, "E3"), 20.0);
    assert_eq!(text(&workbook, "D4"), "Grand Total");
    assert_eq!(number(&workbook, "E4"), 45.0);
}

#[test]
fn refresh_applies_aggregate_filters_to_column_fields() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", 30.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Q1").unwrap();
    sheet.set_cell_value("C4", 20.0).unwrap();
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "Q2").unwrap();
    sheet.set_cell_value("C5", 5.0).unwrap();

    let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .pivot_measure(measure.clone())
        .filter(PivotFilter::Value {
            field: "Quarter".into(),
            measure,
            operator: PivotFilterOperator::GreaterThan,
            value: 32.0,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Q2");
    assert_eq!(text(&workbook, "G1"), "Grand Total");
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 30.0);
    assert_eq!(number(&workbook, "G2"), 30.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(number(&workbook, "F3"), 5.0);
    assert_eq!(number(&workbook, "G3"), 5.0);
    assert_eq!(text(&workbook, "E4"), "Grand Total");
    assert_eq!(number(&workbook, "F4"), 35.0);
    assert_eq!(number(&workbook, "G4"), 35.0);
}

#[test]
fn refresh_applies_label_filter_to_page_field_caption() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Wholesale").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", "Online").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .page("Segment")
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::Label {
            field: "Segment".into(),
            operator: PivotFilterOperator::Contains,
            value: "sale".to_string(),
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Segment");
    assert_eq!(text(&workbook, "F1"), "Wholesale");
    assert_eq!(text(&workbook, "E3"), "Region");
    assert_eq!(text(&workbook, "E4"), "West");
    assert_eq!(number(&workbook, "F4"), 20.0);
    assert_eq!(text(&workbook, "E5"), "Grand Total");
    assert_eq!(number(&workbook, "F5"), 20.0);
}
