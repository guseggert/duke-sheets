use crate::test_support::*;
use pretty_assertions::assert_eq;

#[cfg(feature = "parallel")]
#[test]
fn parallel_refreshes_large_source_snapshot_and_aggregation() {
    let mut workbook = Workbook::new();
    let data_rows = PARALLEL_ROW_THRESHOLD + 7;
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();

    for index in 0..data_rows {
        let row = (index + 2) as u32;
        let region = match index % 3 {
            0 => "East",
            1 => "West",
            _ => "North",
        };
        let quarter = if index % 2 == 0 { "Q1" } else { "Q2" };
        sheet.set_cell_value_at(row - 1, 0, region).unwrap();
        sheet.set_cell_value_at(row - 1, 1, quarter).unwrap();
        sheet.set_cell_value_at(row - 1, 2, 1.0).unwrap();
    }

    let source = CellRange::parse(&format!("A1:C{}", data_rows + 1)).unwrap();
    let pivot = PivotTable::builder("LargeSalesPivot")
        .source_range(source)
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.source_rows, data_rows);
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 8335.0);
    assert_eq!(number(&workbook, "G2"), 8334.0);
    assert_eq!(number(&workbook, "H2"), 16669.0);
    assert_eq!(text(&workbook, "E3"), "North");
    assert_eq!(number(&workbook, "F3"), 8335.0);
    assert_eq!(number(&workbook, "G3"), 8334.0);
    assert_eq!(text(&workbook, "E4"), "West");
    assert_eq!(number(&workbook, "F4"), 8334.0);
    assert_eq!(number(&workbook, "G4"), 8335.0);
    assert_eq!(text(&workbook, "E5"), "Grand Total");
    assert_eq!(number(&workbook, "F5"), 25004.0);
    assert_eq!(number(&workbook, "G5"), 25003.0);
    assert_eq!(number(&workbook, "H5"), data_rows as f64);
}

#[cfg(feature = "parallel")]
#[test]
fn parallel_refreshes_non_visual_totals_with_hidden_total_source() {
    let mut workbook = Workbook::new();
    let repetitions = (PARALLEL_ROW_THRESHOLD / 5) + 1;
    let data_rows = repetitions * 5;
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Channel").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();

    for index in 0..data_rows {
        let row = (index + 1) as u32;
        let (region, quarter, channel, revenue) = match index % 5 {
            0 => ("East", "Q1", "Online", 10.0),
            1 => ("East", "Q2", "Online", 5.0),
            2 => ("East", "Q1", "Store", 99.0),
            3 => ("West", "Q1", "Online", 3.0),
            _ => ("West", "Q2", "Online", 7.0),
        };
        sheet.set_cell_value_at(row, 0, region).unwrap();
        sheet.set_cell_value_at(row, 1, quarter).unwrap();
        sheet.set_cell_value_at(row, 2, channel).unwrap();
        sheet.set_cell_value_at(row, 3, revenue).unwrap();
    }

    let mut layout = PivotLayout::default();
    layout.visual_totals = false;
    let source = CellRange::parse(&format!("A1:D{}", data_rows + 1)).unwrap();
    let pivot = PivotTable::builder("LargeNonVisualTotals")
        .source_range(source)
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

    let stats = workbook
        .refresh_pivots_with_options(&PivotRefreshOptions {
            max_threads: Some(4),
            ..PivotRefreshOptions::default()
        })
        .unwrap();

    let scale = repetitions as f64;
    assert_eq!(stats.source_rows, data_rows);
    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "Q1");
    assert_eq!(text(&workbook, "H1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(number(&workbook, "G2"), 10.0 * scale);
    assert_eq!(number(&workbook, "H2"), 15.0 * scale);
    assert_eq!(text(&workbook, "F3"), "West");
    assert_eq!(number(&workbook, "G3"), 3.0 * scale);
    assert_eq!(number(&workbook, "H3"), 10.0 * scale);
    assert_eq!(text(&workbook, "F4"), "Grand Total");
    assert_eq!(number(&workbook, "G4"), 13.0 * scale);
    assert_eq!(number(&workbook, "H4"), 25.0 * scale);
}

#[test]
fn refreshes_table_sources() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let mut table = Table::new(1, "SalesData", CellRange::parse("A1:B3").unwrap());
    table.columns = vec![
        TableColumn::new(1, "Region"),
        TableColumn::new(2, "Revenue"),
    ];
    sheet.add_table(table);

    let pivot = PivotTable::builder("SalesPivot")
        .table_source("SalesData")
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.source_rows, 2);
    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 10.0);
    assert_eq!(text(&workbook, "D3"), "West");
    assert_eq!(number(&workbook, "E3"), 20.0);
}

#[test]
fn shared_sources_hit_the_internal_snapshot_cache() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();

    let source = CellRange::parse("A1:B2").unwrap();
    let pivot_a = PivotTable::builder("PivotA")
        .source_range(source)
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    let pivot_b = PivotTable::builder("PivotB")
        .source_range(source)
        .target_address("G1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot_a).unwrap();
    sheet.add_pivot_table(pivot_b).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.cache_misses, 1);
    assert_eq!(stats.cache_hits, 1);
    assert_eq!(number(&workbook, "E2"), 10.0);
    assert_eq!(number(&workbook, "H2"), 10.0);
}

#[test]
fn shared_transforms_hit_the_internal_snapshot_cache() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Units").unwrap();
    sheet.set_cell_value("C1", "Price").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 2.0).unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 1.0).unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();

    let source = CellRange::parse("A1:C3").unwrap();
    let grouping = PivotGrouping::Manual {
        field: "Region".into(),
        groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
    };
    let pivot_a = PivotTable::builder("PivotA")
        .source_range(source)
        .target_address("E1")
        .unwrap()
        .row("Region")
        .calculated_field("Revenue", "=Units*Price")
        .grouping(grouping.clone())
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    let pivot_b = PivotTable::builder("PivotB")
        .source_range(source)
        .target_address("H1")
        .unwrap()
        .row("Region")
        .calculated_field("Revenue", "=Units*Price")
        .grouping(grouping)
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot_a).unwrap();
    sheet.add_pivot_table(pivot_b).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.cache_misses, 1);
    assert_eq!(stats.cache_hits, 1);
    assert_eq!(text(&workbook, "E2"), "Coastal");
    assert_eq!(number(&workbook, "F2"), 40.0);
    assert_eq!(text(&workbook, "H2"), "Coastal");
    assert_eq!(number(&workbook, "I2"), 40.0);

    let cache = workbook
        .take_pivot_runtime_cache()
        .unwrap()
        .downcast::<PivotRuntimeCache>()
        .unwrap();
    assert_eq!(cache.snapshot_count(), 1);
    assert_eq!(cache.transformed_snapshot_count(), 1);
}

#[test]
fn external_sources_are_marked_external_without_failing_refresh() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let pivot = PivotTable::builder("ExternalSales")
        .source(PivotSource::External {
            connection_name: "SalesConnection".to_string(),
            command_text: Some("select Region, Revenue from Sales".to_string()),
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.pivot_count, 1);
    assert_eq!(stats.pivots_refreshed, 0);
    assert_eq!(
        workbook.worksheet(0).unwrap().pivot_tables()[0].refresh_status,
        PivotRefreshStatus::External
    );
}

#[test]
fn external_sources_do_not_block_local_pivot_refresh() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();

    let local = PivotTable::builder("LocalSales")
        .source_range(CellRange::parse("A1:B2").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    let external = PivotTable::builder("ExternalSales")
        .source(PivotSource::External {
            connection_name: "SalesConnection".to_string(),
            command_text: Some("select Region, Revenue from Sales".to_string()),
        })
        .target_address("G1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(local).unwrap();
    sheet.add_pivot_table(external).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.pivot_count, 2);
    assert_eq!(stats.pivots_refreshed, 1);
    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 10.0);
    let pivots = workbook.worksheet(0).unwrap().pivot_tables();
    assert_eq!(pivots[0].refresh_status, PivotRefreshStatus::Succeeded);
    assert_eq!(pivots[1].refresh_status, PivotRefreshStatus::External);
}

#[test]
fn refreshes_consolidation_source_ranges() {
    let mut workbook = Workbook::new();
    workbook.add_worksheet_with_name("WestData").unwrap();

    {
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", 15.0).unwrap();
    }
    {
        let sheet = workbook.worksheet_mut(1).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "West").unwrap();
        sheet.set_cell_value("B2", 20.0).unwrap();
    }

    let pivot = PivotTable::builder("ConsolidatedSales")
        .source(PivotSource::Consolidation {
            ranges: vec![
                PivotSourceRange::new("Sheet1", CellRange::parse("A1:B3").unwrap())
                    .with_page_items(["Retail"]),
                PivotSourceRange::new("WestData", CellRange::parse("A1:B2").unwrap())
                    .with_name("West sales"),
            ],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_pivot_table(pivot)
        .unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.pivot_count, 1);
    assert_eq!(stats.pivots_refreshed, 1);
    assert_eq!(stats.source_rows, 3);
    assert_eq!(stats.cache_misses, 1);
    assert_eq!(stats.cache_hits, 0);
    assert_eq!(text(&workbook, "D1"), "Region");
    assert_eq!(text(&workbook, "E1"), "Revenue");
    assert_eq!(text(&workbook, "D2"), "East");
    assert_eq!(number(&workbook, "E2"), 25.0);
    assert_eq!(text(&workbook, "D3"), "West");
    assert_eq!(number(&workbook, "E3"), 20.0);
    assert_eq!(text(&workbook, "D4"), "Grand Total");
    assert_eq!(number(&workbook, "E4"), 45.0);
    assert_eq!(
        workbook.worksheet(0).unwrap().pivot_tables()[0].refresh_status,
        PivotRefreshStatus::Succeeded
    );
}

#[test]
fn named_consolidation_sources_are_marked_external_for_local_refresh() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let pivot = PivotTable::builder("NamedConsolidation")
        .source(PivotSource::Consolidation {
            ranges: vec![PivotSourceRange::named("NamedSource")],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.pivot_count, 1);
    assert_eq!(stats.pivots_refreshed, 0);
    assert_eq!(
        workbook.worksheet(0).unwrap().pivot_tables()[0].refresh_status,
        PivotRefreshStatus::External
    );
}

#[test]
fn shared_consolidation_sources_hit_the_internal_snapshot_cache() {
    let mut workbook = Workbook::new();
    workbook.add_worksheet_with_name("WestData").unwrap();

    {
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
    }
    {
        let sheet = workbook.worksheet_mut(1).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "West").unwrap();
        sheet.set_cell_value("B2", 20.0).unwrap();
    }

    let source = PivotSource::Consolidation {
        ranges: vec![
            PivotSourceRange::new("Sheet1", CellRange::parse("A1:B2").unwrap()),
            PivotSourceRange::new("WestData", CellRange::parse("A1:B2").unwrap()),
        ],
    };
    let pivot_a = PivotTable::builder("PivotA")
        .source(source.clone())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    let pivot_b = PivotTable::builder("PivotB")
        .source(source)
        .target_address("G1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_pivot_table(pivot_a).unwrap();
    sheet.add_pivot_table(pivot_b).unwrap();

    let stats = workbook.refresh_pivots().unwrap();

    assert_eq!(stats.cache_misses, 1);
    assert_eq!(stats.cache_hits, 1);
    assert_eq!(number(&workbook, "E2"), 10.0);
    assert_eq!(number(&workbook, "H2"), 10.0);
}

#[test]
fn internal_snapshot_cache_reuses_then_invalidates_on_source_mutation() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B2").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let first = workbook.refresh_pivots().unwrap();
    assert_eq!(first.cache_misses, 1);
    assert_eq!(first.cache_hits, 0);
    assert_eq!(number(&workbook, "E2"), 10.0);

    let second = workbook.refresh_pivots().unwrap();
    assert_eq!(second.cache_misses, 0);
    assert_eq!(second.cache_hits, 1);
    assert_eq!(number(&workbook, "E2"), 10.0);

    workbook
        .worksheet_mut(0)
        .unwrap()
        .set_cell_value("B2", 15.0)
        .unwrap();

    let third = workbook.refresh_pivots().unwrap();
    assert_eq!(third.cache_misses, 1);
    assert_eq!(third.cache_hits, 0);
    assert_eq!(number(&workbook, "E2"), 15.0);
}
