use crate::test_support::*;
use pretty_assertions::assert_eq;

#[test]
fn refreshes_sum_by_row_field() {
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
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.pivot_count, 1);
    assert_eq!(stats.pivots_refreshed, 1);
    assert_eq!(stats.source_rows, 3);
    assert_eq!(text(&workbook, "D1"), "Region");
    assert_eq!(text(&workbook, "E1"), "Revenue");
    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 25.0);
    assert_eq!(text(&workbook, "D3"), "West");
    assert_eq!(number(&workbook, "E3"), 20.0);
    assert_eq!(text(&workbook, "D4"), "Grand Total");
    assert_eq!(number(&workbook, "E4"), 45.0);
}

#[test]
fn refreshes_with_max_thread_option() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let stats = workbook
        .refresh_pivots_with_options(&PivotRefreshOptions {
            max_threads: Some(1),
            ..PivotRefreshOptions::default()
        })
        .unwrap();

    assert_eq!(stats.pivot_count, 1);
    assert_eq!(stats.pivots_refreshed, 1);
    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 10.0);
    assert_eq!(text(&workbook, "D3"), "West");
    assert_eq!(number(&workbook, "E3"), 20.0);
}

#[test]
fn refreshes_sorted_row_fields() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", 15.0).unwrap();

    let mut region = PivotField::new("Region");
    region.sort = PivotSort::Descending;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row(region)
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "West");
    assert_eq!(text(&workbook, "D3"), "North");
    assert_eq!(text(&workbook, "D4"), "East");
}

#[test]
fn refresh_sorts_parent_row_field_by_scoped_measure_total() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "High").unwrap();
    sheet.set_cell_value("C2", 100.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Low").unwrap();
    sheet.set_cell_value("C3", 1.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "High").unwrap();
    sheet.set_cell_value("C4", 60.0).unwrap();
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "Low").unwrap();
    sheet.set_cell_value("C5", 60.0).unwrap();

    let mut region = PivotField::new("Region");
    region.sort = PivotSort::Descending;
    region.sort_by_measure = Some(PivotMeasure::new("Revenue", PivotAggregate::Sum));
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "West");
    assert_eq!(text(&workbook, "F2"), "High");
    assert_eq!(number(&workbook, "G2"), 60.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(text(&workbook, "F3"), "Low");
    assert_eq!(number(&workbook, "G3"), 60.0);
    assert_eq!(text(&workbook, "E4"), "West Total");
    assert_eq!(number(&workbook, "G4"), 120.0);
    assert_eq!(text(&workbook, "E5"), "East");
    assert_eq!(text(&workbook, "F5"), "High");
    assert_eq!(number(&workbook, "G5"), 100.0);
    assert_eq!(text(&workbook, "E6"), "East");
    assert_eq!(text(&workbook, "F6"), "Low");
    assert_eq!(number(&workbook, "G6"), 1.0);
    assert_eq!(text(&workbook, "E7"), "East Total");
    assert_eq!(number(&workbook, "G7"), 101.0);
}

#[test]
fn refreshes_percentage_show_as_calculations() {
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
    sheet.set_cell_value("C4", 10.0).unwrap();
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "Q2").unwrap();
    sheet.set_cell_value("C5", 50.0).unwrap();

    let source = CellRange::parse("A1:C5").unwrap();
    let grand = PivotTable::builder("GrandPercent")
        .source_range(source)
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_show_as(PivotShowAs::PercentOfGrandTotal),
        )
        .build()
        .unwrap();
    let row = PivotTable::builder("RowPercent")
        .source_range(source)
        .target_address("J1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_show_as(PivotShowAs::PercentOfRowTotal),
        )
        .build()
        .unwrap();
    let column = PivotTable::builder("ColumnPercent")
        .source_range(source)
        .target_address("O1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_show_as(PivotShowAs::PercentOfColumnTotal),
        )
        .build()
        .unwrap();
    sheet.add_pivot_table(grand).unwrap();
    sheet.add_pivot_table(row).unwrap();
    sheet.add_pivot_table(column).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_close(number(&workbook, "F2"), 0.1);
    assert_close(number(&workbook, "G2"), 0.3);
    assert_close(number(&workbook, "H2"), 0.4);
    assert_close(number(&workbook, "F4"), 0.2);
    assert_close(number(&workbook, "G4"), 0.8);
    assert_close(number(&workbook, "H4"), 1.0);

    assert_close(number(&workbook, "K2"), 0.25);
    assert_close(number(&workbook, "L2"), 0.75);
    assert_close(number(&workbook, "M2"), 1.0);
    assert_close(number(&workbook, "K3"), 1.0 / 6.0);
    assert_close(number(&workbook, "L3"), 5.0 / 6.0);
    assert_close(number(&workbook, "K4"), 0.2);
    assert_close(number(&workbook, "L4"), 0.8);

    assert_close(number(&workbook, "P2"), 0.5);
    assert_close(number(&workbook, "Q2"), 0.375);
    assert_close(number(&workbook, "R2"), 0.4);
    assert_close(number(&workbook, "P3"), 0.5);
    assert_close(number(&workbook, "Q3"), 0.625);
    assert_close(number(&workbook, "P4"), 1.0);
    assert_close(number(&workbook, "Q4"), 1.0);
}

#[test]
fn refreshes_parent_percentage_show_as_calculations() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Year").unwrap();
    sheet.set_cell_value("D1", "Quarter").unwrap();
    sheet.set_cell_value("E1", "Revenue").unwrap();
    let rows = [
        ("East", "Retail", "2024", "Q1", 10.0),
        ("East", "Online", "2024", "Q1", 30.0),
        ("East", "Retail", "2024", "Q2", 15.0),
        ("East", "Online", "2024", "Q2", 45.0),
        ("West", "Retail", "2024", "Q1", 20.0),
        ("West", "Online", "2024", "Q1", 10.0),
        ("West", "Retail", "2024", "Q2", 30.0),
        ("West", "Online", "2024", "Q2", 20.0),
        ("East", "Online", "2025", "Q1", 60.0),
        ("West", "Online", "2025", "Q1", 40.0),
    ];
    for (index, (region, segment, year, quarter, revenue)) in rows.into_iter().enumerate() {
        let row = (index + 2) as u32;
        sheet.set_cell_value_at(row - 1, 0, region).unwrap();
        sheet.set_cell_value_at(row - 1, 1, segment).unwrap();
        sheet.set_cell_value_at(row - 1, 2, year).unwrap();
        sheet.set_cell_value_at(row - 1, 3, quarter).unwrap();
        sheet.set_cell_value_at(row - 1, 4, revenue).unwrap();
    }

    let source = CellRange::parse("A1:E11").unwrap();
    let parent_row = PivotTable::builder("ParentRow")
        .source_range(source)
        .target_address("G1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_show_as(PivotShowAs::PercentOfParentRowTotal),
        )
        .build()
        .unwrap();
    let parent_total = PivotTable::builder("ParentTotal")
        .source_range(source)
        .target_address("L1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                PivotShowAs::PercentOfParentTotal {
                    base_field: "Region".into(),
                },
            ),
        )
        .build()
        .unwrap();
    let parent_column = PivotTable::builder("ParentColumn")
        .source_range(source)
        .target_address("O1")
        .unwrap()
        .row("Region")
        .column("Year")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_show_as(PivotShowAs::PercentOfParentColumnTotal),
        )
        .build()
        .unwrap();
    sheet.add_pivot_table(parent_row).unwrap();
    sheet.add_pivot_table(parent_total).unwrap();
    sheet.add_pivot_table(parent_column).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "G3"), "Online");
    assert_close(number(&workbook, "H3"), 90.0 / 100.0);
    assert_close(number(&workbook, "I3"), 45.0 / 60.0);
    assert_close(number(&workbook, "J3"), 135.0 / 160.0);
    assert_eq!(text(&workbook, "G4"), "Retail");
    assert_close(number(&workbook, "H4"), 10.0 / 100.0);
    assert_close(number(&workbook, "I4"), 15.0 / 60.0);
    assert_close(number(&workbook, "J4"), 25.0 / 160.0);

    assert_eq!(text(&workbook, "L3"), "Online");
    assert_close(number(&workbook, "M3"), 135.0 / 160.0);
    assert_eq!(text(&workbook, "L4"), "Retail");
    assert_close(number(&workbook, "M4"), 25.0 / 160.0);

    assert_eq!(text(&workbook, "O1"), "Region");
    assert_close(number(&workbook, "P2"), 40.0 / 100.0);
    assert_close(number(&workbook, "Q2"), 60.0 / 100.0);
    assert_close(number(&workbook, "R2"), 100.0 / 280.0);
    assert_close(number(&workbook, "S2"), 60.0 / 60.0);
    assert_close(number(&workbook, "T2"), 60.0 / 280.0);
    assert_close(number(&workbook, "P3"), 30.0 / 80.0);
    assert_close(number(&workbook, "Q3"), 50.0 / 80.0);
    assert_close(number(&workbook, "R3"), 80.0 / 280.0);
    assert_close(number(&workbook, "S3"), 40.0 / 40.0);
    assert_close(number(&workbook, "T3"), 40.0 / 280.0);
}

#[test]
fn refreshes_index_show_as_calculation() {
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
    sheet.set_cell_value("C5", 40.0).unwrap();

    let pivot = PivotTable::builder("IndexShowAs")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(PivotShowAs::Index),
        )
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_close(number(&workbook, "F2"), 10.0 * 100.0 / (40.0 * 30.0));
    assert_close(number(&workbook, "G2"), 30.0 * 100.0 / (40.0 * 70.0));
    assert_close(number(&workbook, "F3"), 20.0 * 100.0 / (60.0 * 30.0));
    assert_close(number(&workbook, "G3"), 40.0 * 100.0 / (60.0 * 70.0));
}

#[test]
fn refreshes_base_field_show_as_calculations() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Period").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", 1.0).unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", 2.0).unwrap();
    sheet.set_cell_value("B3", 15.0).unwrap();
    sheet.set_cell_value("A4", 3.0).unwrap();
    sheet.set_cell_value("B4", 20.0).unwrap();

    let source = CellRange::parse("A1:B4").unwrap();
    let running = PivotTable::builder("Running")
        .source_range(source)
        .target_address("D1")
        .unwrap()
        .row("Period")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                PivotShowAs::RunningTotal {
                    base_field: "Period".into(),
                },
            ),
        )
        .build()
        .unwrap();
    let difference = PivotTable::builder("Difference")
        .source_range(source)
        .target_address("G1")
        .unwrap()
        .row("Period")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                PivotShowAs::DifferenceFrom {
                    base_field: "Period".into(),
                    base_item: PivotValue::Number(1.0),
                },
            ),
        )
        .build()
        .unwrap();
    let percent_difference = PivotTable::builder("PercentDifference")
        .source_range(source)
        .target_address("J1")
        .unwrap()
        .row("Period")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                PivotShowAs::PercentDifferenceFrom {
                    base_field: "Period".into(),
                    base_item: PivotValue::Number(1.0),
                },
            ),
        )
        .build()
        .unwrap();
    let rank = PivotTable::builder("Rank")
        .source_range(source)
        .target_address("M1")
        .unwrap()
        .row("Period")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                PivotShowAs::RankDescending {
                    base_field: "Period".into(),
                },
            ),
        )
        .build()
        .unwrap();
    sheet.add_pivot_table(running).unwrap();
    sheet.add_pivot_table(difference).unwrap();
    sheet.add_pivot_table(percent_difference).unwrap();
    sheet.add_pivot_table(rank).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(number(&workbook, "E2"), 10.0);
    assert_eq!(number(&workbook, "E3"), 25.0);
    assert_eq!(number(&workbook, "E4"), 45.0);
    assert_eq!(text(&workbook, "E5"), "");

    assert_eq!(number(&workbook, "H2"), 0.0);
    assert_eq!(number(&workbook, "H3"), 5.0);
    assert_eq!(number(&workbook, "H4"), 10.0);

    assert_eq!(number(&workbook, "K2"), 0.0);
    assert_eq!(number(&workbook, "K3"), 0.5);
    assert_eq!(number(&workbook, "K4"), 1.0);

    assert_eq!(number(&workbook, "N2"), 3.0);
    assert_eq!(number(&workbook, "N3"), 2.0);
    assert_eq!(number(&workbook, "N4"), 1.0);
}

#[test]
fn refresh_applies_non_visual_totals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Channel").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", "Online").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", "Online").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "Q1").unwrap();
    sheet.set_cell_value("C4", "Store").unwrap();
    sheet.set_cell_value("D4", 99.0).unwrap();
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "Q1").unwrap();
    sheet.set_cell_value("C5", "Online").unwrap();
    sheet.set_cell_value("D5", 3.0).unwrap();
    sheet.set_cell_value("A6", "West").unwrap();
    sheet.set_cell_value("B6", "Q2").unwrap();
    sheet.set_cell_value("C6", "Online").unwrap();
    sheet.set_cell_value("D6", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.visual_totals = false;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D6").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Quarter", ["Q1"]))
        .filter(PivotFilter::field_items("Channel", ["Online"]))
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "Q1");
    assert_eq!(text(&workbook, "H1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(number(&workbook, "G2"), 10.0);
    assert_eq!(number(&workbook, "H2"), 15.0);
    assert_eq!(text(&workbook, "F3"), "West");
    assert_eq!(number(&workbook, "G3"), 3.0);
    assert_eq!(number(&workbook, "H3"), 10.0);
    assert_eq!(text(&workbook, "F4"), "Grand Total");
    assert_eq!(number(&workbook, "G4"), 13.0);
    assert_eq!(number(&workbook, "H4"), 25.0);
}
