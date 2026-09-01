use crate::test_support::*;
use pretty_assertions::assert_eq;

#[test]
fn format_pivot_plan_reuses_cache_and_exposes_field_major_items() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let pivot_a = PivotTable::builder("SalesPivotA")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    let pivot_b = PivotTable::builder("SalesPivotB")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("G1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_pivot_table(pivot_a).unwrap();
    sheet.add_pivot_table(pivot_b).unwrap();

    let plan = crate::plan::plan_format_pivots(&workbook).unwrap();

    assert_eq!(plan.caches.len(), 1);
    assert_eq!(plan.tables.len(), 2);
    assert_eq!(plan.tables[0].cache_num, 1);
    assert_eq!(plan.tables[1].cache_num, 1);
    assert_eq!(plan.tables[0].visible_rows, None);
    assert_eq!(plan.tables[1].visible_rows, None);
    assert_eq!(
        plan.tables[0].axis_tuples.rows,
        Some(vec![vec![0], vec![1]])
    );
    assert_eq!(plan.tables[0].axis_tuples.columns, Some(Vec::new()));
    assert_eq!(
        plan.tables[1].axis_tuples.rows,
        Some(vec![vec![0], vec![1]])
    );
    assert_eq!(plan.tables[1].axis_tuples.columns, Some(Vec::new()));

    let cache = &plan.caches[0];
    assert_eq!(cache.row_count, 2);
    match &cache.source {
        FormatPivotSource::Worksheet {
            sheet_index,
            sheet_name,
            range,
            table_name,
        } => {
            assert_eq!(*sheet_index, 0);
            assert_eq!(sheet_name, "Sheet1");
            assert_eq!(*range, CellRange::parse("A1:B3").unwrap());
            assert_eq!(table_name, &None);
        }
        FormatPivotSource::Consolidation { .. }
        | FormatPivotSource::External { .. }
        | FormatPivotSource::Scenario { .. }
        | FormatPivotSource::Olap { .. } => panic!("expected worksheet source"),
    }

    let region = &cache.fields[0];
    assert_eq!(region.name, "Region");
    assert_eq!(
        region.shared_items,
        vec![PivotValue::from("East"), PivotValue::from("West")]
    );
    assert_eq!(region.item_ids, vec![0, 1]);

    let revenue = &cache.fields[1];
    assert_eq!(revenue.name, "Revenue");
    assert_eq!(
        revenue.shared_items,
        vec![PivotValue::Number(10.0), PivotValue::Number(20.0)]
    );
    assert_eq!(revenue.item_ids, vec![0, 1]);
}

#[test]
fn format_pivot_plan_reuses_refreshed_runtime_snapshot() {
    let mut workbook = format_cache_reuse_workbook();
    let refresh_stats = workbook.refresh_pivots().unwrap();
    assert_eq!(refresh_stats.cache_misses, 1);

    let (_, plan_stats) = crate::plan::plan_format_pivots_with_stats(&workbook).unwrap();

    assert_eq!(plan_stats.cache_hits, 1);
    assert_eq!(plan_stats.cache_misses, 0);
}

#[test]
fn format_pivot_plan_rejects_stale_structural_generation() {
    let mut workbook = format_cache_reuse_workbook();
    workbook.refresh_pivots().unwrap();
    workbook.add_worksheet_with_name("Other").unwrap();

    let (_, plan_stats) = crate::plan::plan_format_pivots_with_stats(&workbook).unwrap();

    assert_eq!(plan_stats.cache_hits, 0);
    assert_eq!(plan_stats.cache_misses, 1);
}

#[test]
fn format_pivot_plan_exposes_page_filter_visible_rows() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Channel").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Online").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Store").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", "Online").unwrap();
    sheet.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .page("Channel")
        .filter(PivotFilter::field_items("Channel", ["Online"]))
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_pivot_table(pivot)
        .unwrap();

    let plan = crate::plan_format_pivots(&workbook).unwrap();

    assert_eq!(plan.caches.len(), 1);
    assert_eq!(plan.tables.len(), 1);
    assert_eq!(plan.tables[0].visible_rows, Some(vec![0, 2]));
    assert_eq!(
        plan.tables[0].axis_tuples.rows,
        Some(vec![vec![0], vec![2]])
    );
}

#[test]
fn format_pivot_plan_applies_row_and_column_item_filters_to_axis_tuples() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", "Central").unwrap();
    sheet.set_cell_value("B4", "Q1").unwrap();
    sheet.set_cell_value("C4", 30.0).unwrap();
    sheet.set_cell_value("A5", "East").unwrap();
    sheet.set_cell_value("B5", "Q3").unwrap();
    sheet.set_cell_value("C5", 40.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .filter(PivotFilter::field_items("Quarter", ["Q1", "Q2"]))
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_pivot_table(pivot)
        .unwrap();

    let plan = crate::plan_format_pivots(&workbook).unwrap();

    assert_eq!(plan.tables[0].visible_rows, Some(vec![0, 1]));
    assert_eq!(
        plan.tables[0].axis_tuples.rows,
        Some(vec![vec![0], vec![1]])
    );
    assert_eq!(
        plan.tables[0].axis_tuples.columns,
        Some(vec![vec![0], vec![1]])
    );
}

#[test]
fn format_pivot_plan_precomputes_measure_sorted_axis_tuples() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 30.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", 20.0).unwrap();

    let mut region = PivotField::new("Region");
    region.sort = PivotSort::Descending;
    region.sort_by_measure = Some(PivotMeasure::new("Revenue", PivotAggregate::Sum));
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row(region)
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_pivot_table(pivot)
        .unwrap();

    let plan = crate::plan_format_pivots(&workbook).unwrap();

    assert_eq!(plan.caches.len(), 1);
    assert_eq!(
        plan.caches[0].fields[0].shared_items,
        vec![
            PivotValue::from("East"),
            PivotValue::from("West"),
            PivotValue::from("North"),
        ]
    );
    assert_eq!(
        plan.tables[0].axis_tuples.rows,
        Some(vec![vec![1], vec![2], vec![0]])
    );
}

#[test]
fn format_pivot_plan_sorts_parent_axis_field_by_scoped_measure_total() {
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
        .build()
        .unwrap();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_pivot_table(pivot)
        .unwrap();

    let plan = crate::plan_format_pivots(&workbook).unwrap();

    assert_eq!(
        plan.tables[0].axis_tuples.rows,
        Some(vec![vec![1, 0], vec![1, 1], vec![0, 0], vec![0, 1]])
    );
}

#[test]
fn format_pivot_plan_precomputes_numeric_group_axis_tuples() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Age").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", 21.0).unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", 34.0).unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", 42.0).unwrap();
    sheet.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("AgeBands")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Age")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Number {
            field: "Age".into(),
            start: Some(0.0),
            end: Some(60.0),
            interval: 10.0,
        })
        .build()
        .unwrap();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_pivot_table(pivot)
        .unwrap();

    let plan = crate::plan_format_pivots(&workbook).unwrap();

    assert_eq!(
        plan.caches[0].fields[0].shared_items,
        vec![
            PivotValue::Number(20.0),
            PivotValue::Number(30.0),
            PivotValue::Number(40.0),
        ]
    );
    assert_eq!(
        plan.tables[0].axis_tuples.rows,
        Some(vec![vec![0], vec![1], vec![2]])
    );
}

#[test]
fn format_pivot_plan_leaves_manual_group_axis_tuples_to_writers() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("GroupedRegions")
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
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_pivot_table(pivot)
        .unwrap();

    let plan = crate::plan_format_pivots(&workbook).unwrap();

    assert_eq!(plan.tables[0].axis_tuples.rows, None);
}

#[test]
fn format_pivot_plan_applies_manual_group_item_filters_to_visible_rows() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("GroupedRegions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .filter(PivotFilter::field_items("Region", ["Coastal"]))
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_pivot_table(pivot)
        .unwrap();

    let plan = crate::plan_format_pivots(&workbook).unwrap();

    assert_eq!(plan.tables[0].visible_rows, Some(vec![0, 1]));
    assert_eq!(plan.tables[0].axis_tuples.rows, None);
}

fn format_cache_reuse_workbook() -> Workbook {
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
    workbook
}
