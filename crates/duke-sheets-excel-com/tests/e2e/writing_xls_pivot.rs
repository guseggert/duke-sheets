//! Excel COM parity tests for XLS pivot table writing.

use duke_sheets_core::table::{Table, TableColumn};
use duke_sheets_core::{
    CellError, CellRange, PivotAggregate, PivotCacheSourceKind, PivotDateGroupUnit,
    PivotDatePeriod, PivotField, PivotFieldRef, PivotFilter, PivotFilterOperator, PivotGrouping,
    PivotManualGroup, PivotMeasure, PivotShowAs, PivotSort, PivotSource, PivotSourceRange,
    PivotStyle, PivotSubtotal, PivotTable, PivotValue, PivotValuesAxis, Workbook,
};

use crate::roundtrip_through_excel_xls_bytes;

fn xls_basic_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("BasicPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_data_caption_pivot_workbook() -> Workbook {
    let mut wb = xls_basic_pivot_workbook();
    wb.worksheet_mut(0).unwrap().pivot_tables_mut()[0]
        .layout
        .data_caption = "Measures".to_string();
    wb
}

fn xls_layout_flags_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Salesperson").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Ada").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Ben").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();

    let mut pivot = PivotTable::builder("LayoutFlags")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E2")
        .unwrap()
        .row("Region")
        .page("Salesperson")
        .filter(PivotFilter::field_items("Salesperson", ["Ada"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    pivot.layout.show_row_grand_totals = false;
    pivot.layout.show_column_grand_totals = false;
    pivot.layout.page_wrap = 2;
    pivot.layout.page_over_then_down = true;
    pivot.layout.merge_item_labels = true;
    pivot.layout.error_caption = Some("ERR".to_string());
    pivot.layout.show_error = true;
    pivot.layout.missing_caption = Some("MISS".to_string());
    pivot.layout.show_missing = true;
    pivot.layout.enable_wizard = false;
    pivot.layout.enable_drill = false;
    pivot.layout.enable_field_properties = false;
    pivot.layout.field_print_titles = true;
    pivot.layout.item_print_titles = true;
    pivot.layout.grand_total_caption = Some("Grand".to_string());
    pivot.refresh_policy.preserve_formatting = false;
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_table_source_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let mut table = Table::new(1, "SalesData", CellRange::parse("A1:B3").unwrap());
    table.columns = vec![
        TableColumn::new(1, "Region"),
        TableColumn::new(2, "Revenue"),
    ];
    ws.add_table(table);

    let pivot = PivotTable::builder("TablePivot")
        .table_source("SalesData")
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_consolidation_source_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("WestData").unwrap();
    {
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        ws.set_cell_value("A2", "East").unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();
        ws.set_cell_value("A3", "East").unwrap();
        ws.set_cell_value("B3", 15.0).unwrap();
    }
    {
        let ws = wb.worksheet_mut(1).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        ws.set_cell_value("A2", "West").unwrap();
        ws.set_cell_value("B2", 20.0).unwrap();
    }

    let pivot = PivotTable::builder("ConsolidatedPivot")
        .source(PivotSource::Consolidation {
            ranges: vec![
                PivotSourceRange::new("Sheet1", CellRange::parse("A1:B3").unwrap()),
                PivotSourceRange::new("WestData", CellRange::parse("A1:B2").unwrap()),
            ],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn xls_consolidation_source_pivot_workbook_with_page_items() -> Workbook {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("WestData").unwrap();
    {
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        ws.set_cell_value("A2", "East").unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();
        ws.set_cell_value("A3", "East").unwrap();
        ws.set_cell_value("B3", 15.0).unwrap();
    }
    {
        let ws = wb.worksheet_mut(1).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        ws.set_cell_value("A2", "West").unwrap();
        ws.set_cell_value("B2", 20.0).unwrap();
    }

    let pivot = PivotTable::builder("ConsolidatedPagePivot")
        .source(PivotSource::Consolidation {
            ranges: vec![
                PivotSourceRange::new("Sheet1", CellRange::parse("A1:B3").unwrap())
                    .with_page_items(["Retail"]),
                PivotSourceRange::new("WestData", CellRange::parse("A1:B2").unwrap())
                    .with_page_items(["Wholesale"]),
            ],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn xls_named_consolidation_source_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    {
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        ws.set_cell_value("A2", "East").unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();
    }
    wb.define_name("NamedSalesSource", "Sheet1!$A$1:$B$2")
        .unwrap();

    let pivot = PivotTable::builder("NamedConsolidatedPivot")
        .source(PivotSource::Consolidation {
            ranges: vec![PivotSourceRange::named("NamedSalesSource")],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn xls_external_consolidation_source_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let pivot = PivotTable::builder("ExternalConsolidatedPivot")
        .source(PivotSource::Consolidation {
            ranges: vec![
                PivotSourceRange::new("ExternalData", CellRange::parse("A1:B3").unwrap())
                    .with_external_relationship_target("external_sales.xlsx"),
            ],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn xls_scenario_source_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let pivot = PivotTable::builder("ScenarioPivot")
        .source(PivotSource::Scenario {
            name: String::new(),
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn include_new_items_field(name: &str) -> PivotField {
    let mut field = PivotField::new(name);
    field.include_new_items_in_filter = true;
    field
}

fn xls_average_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("AveragePivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Average, "Average Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_column_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Quarter").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Q1").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "East").unwrap();
    ws.set_cell_value("B3", "Q2").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "West").unwrap();
    ws.set_cell_value("B4", "Q1").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("RevenueByQuarter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_page_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Salesperson").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Ada").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Ben").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();

    let pivot = PivotTable::builder("RevenueByRep")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .page(include_new_items_field("Salesperson"))
        .filter(PivotFilter::field_items("Salesperson", ["Ada"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_multi_select_page_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Salesperson").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Ada").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Ben").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", "Cam").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("MultiPage")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .page(include_new_items_field("Region"))
        .row("Salesperson")
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_top_n_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "North").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", "South").unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder("TopRegions")
        .source_range(CellRange::parse("A1:B5").unwrap())
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
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_percent_top_n_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "North").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", "South").unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder("TopPercentRegions")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::TopN {
            field: PivotFieldRef::new("Region"),
            measure: PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue"),
            n: 50,
            top: true,
            percent: true,
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_label_contains_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Segment").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Enterprise").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Retail").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "North").unwrap();
    ws.set_cell_value("B4", "Education").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("LabelSegments")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::Label {
            field: PivotFieldRef::new("Segment"),
            operator: PivotFilterOperator::Contains,
            value: "Ed".to_string(),
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_label_equals_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "North").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("LabelEquals")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::Label {
            field: PivotFieldRef::new("Region"),
            operator: PivotFilterOperator::Equals,
            value: "East".to_string(),
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_label_prefix_suffix_pivot_workbook(
    name: &str,
    operator: PivotFilterOperator,
    value: &str,
) -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "North").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder(name)
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::Label {
            field: PivotFieldRef::new("Region"),
            operator,
            value: value.to_string(),
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_label_between_pivot_workbook(name: &str, not_between: bool) -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "North").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "West").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder(name)
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::LabelBetween {
            field: PivotFieldRef::new("Region"),
            start: "East".to_string(),
            end: "West".to_string(),
            not_between,
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_value_filter_pivot_workbook(name: &str, filter: PivotFilter) -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "North").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", "South").unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder(name)
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(filter)
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_date_filter_pivot_workbook(name: &str, filter: PivotFilter) -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Order Date").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 44927.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 44958.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", 44986.0).unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", 45016.0).unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder(name)
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Order Date")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(filter)
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_column_value_filter_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Segment").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Enterprise").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Enterprise").unwrap();
    ws.set_cell_value("C3", 30.0).unwrap();
    ws.set_cell_value("A4", "East").unwrap();
    ws.set_cell_value("B4", "Retail").unwrap();
    ws.set_cell_value("C4", 20.0).unwrap();
    ws.set_cell_value("A5", "West").unwrap();
    ws.set_cell_value("B5", "Retail").unwrap();
    ws.set_cell_value("C5", 40.0).unwrap();

    let pivot = PivotTable::builder("ColumnValueFilter")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Segment")
        .column("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::Value {
            field: PivotFieldRef::new("Region"),
            measure: PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue"),
            operator: PivotFilterOperator::GreaterThan,
            value: 25.0,
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_second_measure_value_filter_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("C1", "Cost").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("C2", 4.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("C3", 12.0).unwrap();
    ws.set_cell_value("A4", "North").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("C4", 18.0).unwrap();

    let mut pivot = PivotTable::builder("SecondMeasureValueFilter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .named_measure("Cost", PivotAggregate::Sum, "Total Cost")
        .filter(PivotFilter::Value {
            field: PivotFieldRef::new("Region"),
            measure: PivotMeasure::new("Cost", PivotAggregate::Sum).with_name("Total Cost"),
            operator: PivotFilterOperator::GreaterThan,
            value: 10.0,
        })
        .build()
        .unwrap();
    pivot.layout.values_axis = PivotValuesAxis::Columns;
    pivot.layout.values_axis_position = Some(0);
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_row_item_filter_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("RowItemFilter")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row(include_new_items_field("Region"))
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_row_item_filter_default_include_new_items_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("RowItemFilterDefault")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_column_item_filter_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Salesperson").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Ada").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Ben").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", "Cam").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("ColumnItemFilter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Salesperson")
        .column(include_new_items_field("Region"))
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_source_only_item_filter_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Channel").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Online").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Store").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", "Online").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("SourceOnlyItemFilter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .filter(PivotFilter::field_items("Channel", ["Online"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_grouped_row_item_filter_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("GroupedRowItemFilter")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row(include_new_items_field("Region"))
        .filter(PivotFilter::field_items("Region", ["Coastal"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: PivotFieldRef::new("Region"),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_grouped_column_item_filter_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Salesperson").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Ada").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Ben").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", "Cam").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("GroupedColumnItemFilter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Salesperson")
        .column(include_new_items_field("Region"))
        .filter(PivotFilter::field_items("Region", ["Coastal"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: PivotFieldRef::new("Region"),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_axis_options_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();

    let mut region = PivotField::new("Region")
        .with_caption("Market")
        .with_subtotals([
            PivotSubtotal::Sum,
            PivotSubtotal::Count,
            PivotSubtotal::Average,
            PivotSubtotal::Max,
            PivotSubtotal::Min,
            PivotSubtotal::Product,
            PivotSubtotal::CountNumbers,
            PivotSubtotal::StdDev,
            PivotSubtotal::StdDevP,
            PivotSubtotal::Var,
            PivotSubtotal::VarP,
        ]);
    region.sort = PivotSort::Descending;
    region.subtotal_caption = Some("Regional subtotal".to_string());
    region.show_empty_items = true;
    region.show_drop_downs = false;
    region.insert_blank_row = true;
    region.subtotal_top = false;
    region.insert_page_break = true;
    region.include_new_items_in_filter = true;

    let pivot = PivotTable::builder("AxisOptions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row(region)
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_sort_by_measure_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Quarter").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Q1").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Q1").unwrap();
    ws.set_cell_value("C3", 50.0).unwrap();

    let mut region =
        PivotField::new("Region").with_sort_by(PivotFieldRef::new("Revenue"), PivotAggregate::Sum);
    region.sort = PivotSort::Descending;
    region.include_new_items_in_filter = true;

    let pivot = PivotTable::builder("ValueSortedPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .named_measure("Revenue", PivotAggregate::Sum, "Sum of Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_styled_pivot_style() -> PivotStyle {
    PivotStyle {
        name: Some("PivotStyleLight16".to_string()),
        show_row_headers: false,
        show_column_headers: true,
        show_row_stripes: true,
        show_column_stripes: true,
        show_last_column: true,
    }
}

fn xls_styled_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("StyledPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .style(xls_styled_pivot_style())
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_calculated_field_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Units").unwrap();
    ws.set_cell_value("C1", "Price").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 2.0).unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 7.0).unwrap();
    ws.set_cell_value("C3", 3.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRevenue")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .calculated_field("Revenue", "=Units*Price")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_calculated_item_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRegion")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_item("Region", "Combined", "East+West")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_calculated_item_cell_like_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Quarter").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "Q1").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "Q2").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("CalculatedQuarter")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Quarter")
        .calculated_item("Quarter", "H1", "Q1+Q2")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_calculated_item_string_ref_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRegionStringRef")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_item("Region", "Combined", "\"East\"+West")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_calculated_field_function_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Units").unwrap();
    ws.set_cell_value("C1", "Price").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 2.0).unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 7.0).unwrap();
    ws.set_cell_value("C3", 3.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRevenueFunction")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .calculated_field("Revenue", "=SUM(Units,Price)")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_calculated_item_function_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRegionFunction")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_item("Region", "Combined", "MAX(East,West)")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_multi_measure_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("C1", "Units").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("C2", 2.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("C3", 3.0).unwrap();

    let mut pivot = PivotTable::builder("RevenueAndUnits")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .named_measure("Units", PivotAggregate::Average, "Average Units")
        .build()
        .unwrap();
    pivot.layout.values_axis = PivotValuesAxis::Columns;
    pivot.layout.values_axis_position = Some(0);
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_all_aggregate_cases() -> [(PivotAggregate, &'static str, u16); 11] {
    [
        (PivotAggregate::Sum, "Sum Revenue", 0),
        (PivotAggregate::Count, "Count Revenue", 1),
        (PivotAggregate::Average, "Average Revenue", 2),
        (PivotAggregate::Max, "Max Revenue", 3),
        (PivotAggregate::Min, "Min Revenue", 4),
        (PivotAggregate::Product, "Product Revenue", 5),
        (PivotAggregate::CountNumbers, "Count Numbers Revenue", 6),
        (PivotAggregate::StdDev, "StdDev Revenue", 7),
        (PivotAggregate::StdDevP, "StdDevP Revenue", 8),
        (PivotAggregate::Var, "Var Revenue", 9),
        (PivotAggregate::VarP, "VarP Revenue", 10),
    ]
}

fn xls_all_aggregate_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 2.0).unwrap();
    ws.set_cell_value("A3", "East").unwrap();
    ws.set_cell_value("B3", 3.0).unwrap();
    ws.set_cell_value("A4", "West").unwrap();
    ws.set_cell_value("B4", 5.0).unwrap();
    ws.set_cell_value("A5", "West").unwrap();
    ws.set_cell_value("B5", 7.0).unwrap();

    let mut builder = PivotTable::builder("AllAggregatePivot")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region");
    for (aggregate, caption, _) in xls_all_aggregate_cases() {
        builder = builder.named_measure("Revenue", aggregate, caption);
    }
    let mut pivot = builder.build().unwrap();
    pivot.layout.values_axis = PivotValuesAxis::Columns;
    pivot.layout.values_axis_position = Some(0);
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_show_as_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Quarter").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Q1").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Q1").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "East").unwrap();
    ws.set_cell_value("B4", "Q2").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();
    ws.set_cell_value("A5", "West").unwrap();
    ws.set_cell_value("B5", "Q2").unwrap();
    ws.set_cell_value("C5", 40.0).unwrap();

    let mut pivot = PivotTable::builder("ShowAsPivot")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Pct Row")
                .with_show_as(PivotShowAs::PercentOfRowTotal),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Running Region")
                .with_show_as(PivotShowAs::RunningTotal {
                    base_field: PivotFieldRef::new("Region"),
                }),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Diff East")
                .with_show_as(PivotShowAs::DifferenceFrom {
                    base_field: PivotFieldRef::new("Region"),
                    base_item: PivotValue::String("East".to_string()),
                }),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Pct Diff East")
                .with_show_as(PivotShowAs::PercentDifferenceFrom {
                    base_field: PivotFieldRef::new("Region"),
                    base_item: PivotValue::String("East".to_string()),
                }),
        )
        .build()
        .unwrap();
    pivot.layout.values_axis = PivotValuesAxis::Columns;
    pivot.layout.values_axis_position = Some(0);
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_custom_measure_format_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 1234.5).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 6789.25).unwrap();

    let pivot = PivotTable::builder("FormattedRevenue")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Formatted Revenue")
                .with_number_format("#,##0.0"),
        )
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_numeric_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Age").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 7.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 13.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", 28.0).unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", 42.0).unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder("AgeBands")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Age")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Number {
            field: "Age".into(),
            start: Some(0.0),
            end: Some(60.0),
            interval: 10.0,
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_date_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Date").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 43831.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 43831.0).unwrap();
    ws.set_cell_value("B3", 15.0).unwrap();
    ws.set_cell_value("A4", 43862.0).unwrap();
    ws.set_cell_value("B4", 20.0).unwrap();
    ws.set_cell_value("A5", 43862.0).unwrap();
    ws.set_cell_value("B5", 25.0).unwrap();
    let pivot = PivotTable::builder("MonthlyRevenue")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Date")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: "Date".into(),
            units: vec![PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_multi_unit_date_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "SaleDate").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 45292.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 45323.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", 45658.0).unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("GroupedDates")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("SaleDate")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: "SaleDate".into(),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_manual_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "Central").unwrap();
    ws.set_cell_value("B2", 30.0).unwrap();
    ws.set_cell_value("A3", "East").unwrap();
    ws.set_cell_value("B3", 10.0).unwrap();
    ws.set_cell_value("A4", "West").unwrap();
    ws.set_cell_value("B4", 20.0).unwrap();
    ws.set_cell_value("A5", "North").unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();
    ws.set_cell_value("A6", "South").unwrap();
    ws.set_cell_value("B6", 50.0).unwrap();
    ws.set_cell_value("A7", "International").unwrap();
    ws.set_cell_value("B7", 60.0).unwrap();

    let pivot = PivotTable::builder("ManualGroupedRegions")
        .source_range(CellRange::parse("A1:B7").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![
                PivotManualGroup::new("Coastal", ["East", "West", "South"]),
                PivotManualGroup::new("Interior", ["Central", "North"]),
            ],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_numeric_manual_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Age").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 7.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 13.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", 21.0).unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", 34.0).unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();
    ws.set_cell_value("A6", 55.0).unwrap();
    ws.set_cell_value("B6", 50.0).unwrap();

    let pivot = PivotTable::builder("ManualAgeGroups")
        .source_range(CellRange::parse("A1:B6").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Age")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: "Age".into(),
            groups: vec![
                PivotManualGroup::new("Young", [PivotValue::Number(7.0), PivotValue::Number(13.0)]),
                PivotManualGroup::new(
                    "Adult",
                    [PivotValue::Number(21.0), PivotValue::Number(34.0)],
                ),
            ],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_bool_error_manual_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Flag").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", true).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", false).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", CellError::Na).unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", CellError::Div0).unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder("ManualBoolErrorGroups")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Flag")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: "Flag".into(),
            groups: vec![
                PivotManualGroup::new(
                    "Booleans",
                    [PivotValue::Boolean(true), PivotValue::Boolean(false)],
                ),
                PivotManualGroup::new(
                    "Errors",
                    [
                        PivotValue::Error(CellError::Na),
                        PivotValue::Error(CellError::Div0),
                    ],
                ),
            ],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_manual_column_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Quarter").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "Central").unwrap();
    ws.set_cell_value("B2", "Q1").unwrap();
    ws.set_cell_value("C2", 30.0).unwrap();
    ws.set_cell_value("A3", "East").unwrap();
    ws.set_cell_value("B3", "Q1").unwrap();
    ws.set_cell_value("C3", 10.0).unwrap();
    ws.set_cell_value("A4", "West").unwrap();
    ws.set_cell_value("B4", "Q2").unwrap();
    ws.set_cell_value("C4", 20.0).unwrap();
    ws.set_cell_value("A5", "North").unwrap();
    ws.set_cell_value("B5", "Q2").unwrap();
    ws.set_cell_value("C5", 40.0).unwrap();
    ws.set_cell_value("A6", "South").unwrap();
    ws.set_cell_value("B6", "Q3").unwrap();
    ws.set_cell_value("C6", 50.0).unwrap();
    ws.set_cell_value("A7", "International").unwrap();
    ws.set_cell_value("B7", "Q3").unwrap();
    ws.set_cell_value("C7", 60.0).unwrap();

    let pivot = PivotTable::builder("ManualColumnGroups")
        .source_range(CellRange::parse("A1:C7").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Quarter")
        .column("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![
                PivotManualGroup::new("Coastal", ["East", "West", "South"]),
                PivotManualGroup::new("Interior", ["Central", "North"]),
            ],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xls_manual_page_grouped_pivot_workbook() -> Workbook {
    xls_manual_page_grouped_pivot_workbook_with_filter_items(&["Coastal"])
}

fn xls_manual_page_grouped_pivot_workbook_with_filter_items(allowed_items: &[&str]) -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Salesperson").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "Central").unwrap();
    ws.set_cell_value("B2", "Dia").unwrap();
    ws.set_cell_value("C2", 30.0).unwrap();
    ws.set_cell_value("A3", "East").unwrap();
    ws.set_cell_value("B3", "Ada").unwrap();
    ws.set_cell_value("C3", 10.0).unwrap();
    ws.set_cell_value("A4", "West").unwrap();
    ws.set_cell_value("B4", "Ben").unwrap();
    ws.set_cell_value("C4", 20.0).unwrap();
    ws.set_cell_value("A5", "North").unwrap();
    ws.set_cell_value("B5", "Eli").unwrap();
    ws.set_cell_value("C5", 40.0).unwrap();
    ws.set_cell_value("A6", "South").unwrap();
    ws.set_cell_value("B6", "Cora").unwrap();
    ws.set_cell_value("C6", 50.0).unwrap();
    ws.set_cell_value("A7", "International").unwrap();
    ws.set_cell_value("B7", "Fay").unwrap();
    ws.set_cell_value("C7", 60.0).unwrap();

    let mut pivot = PivotTable::builder("ManualPageRegions")
        .source_range(CellRange::parse("A1:C7").unwrap())
        .target_address("E2")
        .unwrap()
        .row("Salesperson")
        .page(include_new_items_field("Region"));
    if !allowed_items.is_empty() {
        pivot = pivot.filter(PivotFilter::field_items(
            "Region",
            allowed_items.iter().copied(),
        ));
    }
    let pivot = pivot
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![
                PivotManualGroup::new("Coastal", ["East", "West", "South"]),
                PivotManualGroup::new("Interior", ["Central", "North"]),
            ],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

// features: Pivot cache (source data); Pivot table definition; Row / column / value fields; Filter (page) fields; Aggregate functions (Sum/Count/Avg/...)
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_opens_xls_with_native_pivot_table() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_basic_pivot_workbook());
    assert!(xls_cfb_has_stream(&excel_bytes, "/_SX_DB_CUR/0001"));

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .expect("BasicPivot pivot after Excel re-save");
    assert!(!pivot.rows[0].include_new_items_in_filter);

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let row_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD record after Excel re-save");
    let row_sxvdex = records[row_sxvd_position + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(row_sxvdex[0..4].try_into().unwrap()) & 0x8000,
        0,
        "Excel-compatible XLS output should leave fHideNewItems clear"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_data_caption() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_data_caption_pivot_workbook());

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let sxview = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxview[42..44].try_into().unwrap()),
        8,
        "Excel should preserve the custom SXVIEW data-caption length"
    );
    assert!(
        String::from_utf8_lossy(sxview).contains("Measures"),
        "Excel should preserve the custom SXVIEW data caption"
    );

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .expect("BasicPivot pivot after Excel re-save");
    assert_eq!(pivot.layout.data_caption, "Measures");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_table_layout_flags() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_layout_flags_pivot_workbook());

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let sxview = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxview[36..38].try_into().unwrap()) & 0x0003,
        0,
        "Excel should preserve disabled row/column grand totals"
    );
    let sxex = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F1).then_some(payload))
        .expect("SXEX after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxex[14..16].try_into().unwrap()) & 0x01FF,
        0x0005,
        "Excel should preserve page-over-then-down plus page wrap"
    );
    let sxex_flags = u16::from_le_bytes(sxex[16..18].try_into().unwrap());
    assert_eq!(sxex_flags & 0x0010, 0x0010, "merge labels");
    assert_eq!(sxex_flags & 0x0020, 0x0020, "display error caption");
    assert_eq!(sxex_flags & 0x0040, 0x0040, "display missing caption");
    assert_eq!(sxex_flags & 0x0007, 0, "wizard/drill/field dialog disabled");
    assert_eq!(sxex_flags & 0x0008, 0, "preserve formatting disabled");
    assert!(String::from_utf8_lossy(sxex).contains("ERR"));
    assert!(String::from_utf8_lossy(sxex).contains("MISS"));
    let sxviewex9 = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0810).then_some(payload))
        .expect("SXVIEWEX9 after Excel re-save");
    let sxviewex9_flags = u32::from_le_bytes(sxviewex9[8..12].try_into().unwrap());
    assert_eq!(sxviewex9_flags & 0x0002, 0x0002, "print titles");
    assert_eq!(
        sxviewex9_flags & 0x0020,
        0x0020,
        "repeat item labels when printed"
    );
    assert!(String::from_utf8_lossy(sxviewex9).contains("Grand"));

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("LayoutFlags")
        .expect("LayoutFlags pivot after Excel re-save");
    assert!(!pivot.layout.show_row_grand_totals);
    assert!(!pivot.layout.show_column_grand_totals);
    assert_eq!(pivot.layout.page_wrap, 2);
    assert!(pivot.layout.page_over_then_down);
    assert!(pivot.layout.merge_item_labels);
    assert_eq!(pivot.layout.error_caption.as_deref(), Some("ERR"));
    assert!(pivot.layout.show_error);
    assert_eq!(pivot.layout.missing_caption.as_deref(), Some("MISS"));
    assert!(pivot.layout.show_missing);
    assert!(!pivot.layout.enable_wizard);
    assert!(!pivot.layout.enable_drill);
    assert!(!pivot.layout.enable_field_properties);
    assert!(pivot.layout.field_print_titles);
    assert!(pivot.layout.item_print_titles);
    assert_eq!(pivot.layout.grand_total_caption.as_deref(), Some("Grand"));
    assert!(!pivot.refresh_policy.preserve_formatting);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_table_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_table_source_pivot_workbook());
    assert!(xls_cfb_has_stream(&excel_bytes, "/_SX_DB_CUR/0001"));

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TablePivot")
        .expect("TablePivot pivot after Excel re-save");
    assert!(matches!(
        &pivot.source,
        PivotSource::Table { name } if name == "SalesData"
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_consolidation_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_consolidation_source_pivot_workbook());
    assert!(xls_cfb_has_stream(&excel_bytes, "/_SX_DB_CUR/0001"));

    let cache_records = xls_record_payloads(&xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001"));
    let sxdb = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C6).then_some(payload))
        .expect("SXDB after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxdb[16..18].try_into().unwrap()),
        0x0004,
        "Excel should preserve the BIFF8 consolidation source kind"
    );

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ConsolidatedPivot")
        .expect("ConsolidatedPivot pivot after Excel re-save");
    let PivotSource::Consolidation { ranges } = &pivot.source else {
        panic!("expected consolidation source after Excel re-save");
    };
    assert_eq!(ranges.len(), 2);
    assert_eq!(ranges[0].sheet.as_deref(), Some("Sheet1"));
    assert_eq!(ranges[0].range, Some(CellRange::parse("A1:B3").unwrap()));
    assert_eq!(ranges[1].sheet.as_deref(), Some("WestData"));
    assert_eq!(ranges[1].range, Some(CellRange::parse("A1:B2").unwrap()));
    assert_eq!(pivot.rows[0].field.name, "Row");
    assert_eq!(pivot.columns[0].field.name, "Column");
    assert_eq!(pivot.measures[0].field.name, "Value");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_consolidation_page_items() {
    let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(
        &xls_consolidation_source_pivot_workbook_with_page_items(),
    );
    assert!(xls_cfb_has_stream(&excel_bytes, "/_SX_DB_CUR/0001"));

    let workbook_records = xls_record_payloads(&xls_cfb_stream(&excel_bytes, "/Workbook"));
    assert!(
        workbook_records
            .iter()
            .any(|(record_type, payload)| *record_type == 0x00D1 && payload == &[2, 0]),
        "Excel should preserve SXTBRGIITM page item count"
    );

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ConsolidatedPagePivot")
        .expect("ConsolidatedPagePivot pivot after Excel re-save");
    let PivotSource::Consolidation { ranges } = &pivot.source else {
        panic!("expected consolidation source after Excel re-save");
    };
    assert_eq!(ranges.len(), 2);
    assert_eq!(ranges[0].sheet.as_deref(), Some("Sheet1"));
    assert_eq!(ranges[0].range, Some(CellRange::parse("A1:B3").unwrap()));
    assert_eq!(ranges[0].page_items, vec!["Retail".to_string()]);
    assert_eq!(ranges[1].sheet.as_deref(), Some("WestData"));
    assert_eq!(ranges[1].range, Some(CellRange::parse("A1:B2").unwrap()));
    assert_eq!(ranges[1].page_items, vec!["Wholesale".to_string()]);
    assert_eq!(pivot.rows[0].field.name, "Row");
    assert_eq!(pivot.columns[0].field.name, "Column");
    assert_eq!(pivot.measures[0].field.name, "Value");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_named_consolidation_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_named_consolidation_source_pivot_workbook());
    assert!(xls_cfb_has_stream(&excel_bytes, "/_SX_DB_CUR/0001"));

    let cache_records = xls_record_payloads(&xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001"));
    let sxdb = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C6).then_some(payload))
        .expect("SXDB after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxdb[16..18].try_into().unwrap()),
        0x0004,
        "Excel should preserve the BIFF8 consolidation source kind"
    );

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("NamedConsolidatedPivot")
        .expect("NamedConsolidatedPivot pivot after Excel re-save");
    let PivotSource::Consolidation { ranges } = &pivot.source else {
        panic!("expected consolidation source after Excel re-save");
    };
    assert_eq!(ranges.len(), 1);
    assert_eq!(ranges[0].name.as_deref(), Some("NamedSalesSource"));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_external_consolidation_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_external_consolidation_source_pivot_workbook());
    assert!(xls_cfb_has_stream(&excel_bytes, "/_SX_DB_CUR/0001"));

    let workbook_records = xls_record_payloads(&xls_cfb_stream(&excel_bytes, "/Workbook"));
    let dconref = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0051).then_some(payload))
        .expect("DCONREF after Excel re-save");
    let expected_source = b"[external_sales.xlsx]ExternalData";
    assert_eq!(
        dconref[8], 0x00,
        "Excel should preserve compressed DCONREF text"
    );
    assert_eq!(
        dconref[9], 0x01,
        "Excel should preserve the external DCONREF marker"
    );
    assert_eq!(&dconref[10..10 + expected_source.len()], expected_source);

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ExternalConsolidatedPivot")
        .expect("ExternalConsolidatedPivot pivot after Excel re-save");
    let PivotSource::Consolidation { ranges } = &pivot.source else {
        panic!("expected consolidation source after Excel re-save");
    };
    assert_eq!(ranges.len(), 1);
    assert_eq!(ranges[0].sheet.as_deref(), Some("ExternalData"));
    assert_eq!(ranges[0].range, Some(CellRange::parse("A1:B3").unwrap()));
    assert_eq!(
        ranges[0].external_relationship_target.as_deref(),
        Some("external_sales.xlsx")
    );
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_unnamed_scenario_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_scenario_source_pivot_workbook());
    assert!(xls_cfb_has_stream(&excel_bytes, "/_SX_DB_CUR/0001"));

    let workbook_records = xls_record_payloads(&xls_cfb_stream(&excel_bytes, "/Workbook"));
    let sxvs = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00E3).then_some(payload))
        .expect("SXVS after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxvs[0..2].try_into().unwrap()),
        0x0010,
        "Excel should preserve the BIFF8 scenario workbook cache kind"
    );

    let cache_records = xls_record_payloads(&xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001"));
    let sxdb = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C6).then_some(payload))
        .expect("SXDB after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxdb[16..18].try_into().unwrap()),
        0x0008,
        "Excel should preserve the BIFF8 scenario source kind"
    );

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ScenarioPivot")
        .expect("ScenarioPivot pivot after Excel re-save");
    assert!(matches!(
        &pivot.source,
        PivotSource::Scenario { name } if name.is_empty()
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(PivotCacheSourceKind::Scenario)
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_average_pivot_data_field() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_average_pivot_workbook());
    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let sxdi = xls_record_payloads(&workbook)
        .into_iter()
        .find_map(|(record_type, payload)| (record_type == 0x00C5).then_some(payload))
        .expect("SXDI data field record");

    assert_eq!(
        u16::from_le_bytes(sxdi[2..4].try_into().unwrap()),
        2,
        "Excel should preserve Average as the data-field function"
    );
    assert!(
        String::from_utf8_lossy(&sxdi).contains("Average Revenue"),
        "Excel should preserve the data-field caption"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_column_axis() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_column_pivot_workbook());
    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let sxvd_axes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxvd_axes, vec![0x0001, 0x0002, 0x0008]);
    assert_eq!(
        records
            .iter()
            .filter(|(record_type, _)| *record_type == 0x00B4)
            .count(),
        2,
        "Excel should preserve row and column axis declarations"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_page_axis() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_page_pivot_workbook());
    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let sxvd_axes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxvd_axes, vec![0x0001, 0x0004, 0x0008]);
    assert!(
        records
            .iter()
            .any(|(record_type, _)| *record_type == 0x00B6),
        "Excel should preserve the page-field declaration"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_multi_item_page_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_multi_select_page_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("MultiPage")
        .expect("MultiPage pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );
    assert!(!pivot.rows[0].include_new_items_in_filter);

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let page_field = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B6).then_some(payload))
        .expect("page-field SXPI after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(page_field[2..4].try_into().unwrap()),
        0x7FFD,
        "Excel should preserve the BIFF8 multiple-items page-filter sentinel"
    );

    let page_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0004
        })
        .expect("page SXVD after Excel re-save");
    let page_items = sxvi_payloads_for_xls_sxvd(&records, page_sxvd_position);
    assert_eq!(
        page_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Excel should keep only Central hidden for the multi-select page filter"
    );
    assert_eq!(
        i16::from_le_bytes(page_items[2][4..6].try_into().unwrap()),
        2
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_top_n_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_top_n_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TopRegions")
        .expect("TopRegions pivot after Excel re-save");

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
        other => panic!("unexpected pivot filter after Excel re-save: {other:?}"),
    }

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let sxvdex = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX after Excel re-save");
    let flags = u32::from_le_bytes(sxvdex[0..4].try_into().unwrap());
    assert_ne!(flags & 0x0800, 0, "Excel should preserve fAutoShow");
    let sxaddl_autoshow = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 10
                && payload[4] == 0x17
                && payload[5] == 0x37)
                .then_some(payload)
        })
        .expect("SXADDL sxdAutoshow after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(sxaddl_autoshow[6..10].try_into().unwrap()),
        2,
        "Excel should preserve the top count"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_percent_top_n_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_percent_top_n_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TopPercentRegions")
        .expect("TopPercentRegions pivot after Excel re-save");

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
            assert_eq!(*n, 50);
            assert!(*top);
            assert!(*percent);
        }
        other => panic!("unexpected pivot filter after Excel re-save: {other:?}"),
    }

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let filter_extension = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 40
                && payload[4] == 0x1D
                && payload[5] == 0x3C)
                .then_some(payload)
        })
        .expect("SXADDL pivot top-percent filter after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_extension[12..16].try_into().unwrap()),
        3,
        "Excel should preserve the top-percent filter type"
    );
    assert_eq!(
        &filter_extension[32..40],
        &[0x49, 0x40, 0, 0, 0, 0, 0, 0],
        "Excel should preserve the top-percent threshold"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_label_contains_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_label_contains_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("LabelSegments")
        .expect("LabelSegments pivot after Excel re-save");

    assert_eq!(pivot.filters.len(), 1);
    match &pivot.filters[0] {
        PivotFilter::Label {
            field,
            operator,
            value,
        } => {
            assert_eq!(field.name, "Segment");
            assert_eq!(*operator, PivotFilterOperator::Contains);
            assert_eq!(value, "Ed");
        }
        other => panic!("unexpected pivot filter after Excel re-save: {other:?}"),
    }

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let filter_collection = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 24
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL pivot label filter collection after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
        1,
        "Excel should preserve the zero-based Segment source field index"
    );
    assert_eq!(
        u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
        10,
        "Excel should preserve the caption-contains filter type"
    );

    let wildcard = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 12
                && payload[4] == 0x1D
                && payload[5] == 0x3D)
                .then_some(payload)
        })
        .expect("SXADDL wildcard label criterion after Excel re-save");
    assert_eq!(xls_unicode_string_at(wildcard, 12), "*Ed*");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_label_equals_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_label_equals_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("LabelEquals")
        .expect("LabelEquals pivot after Excel re-save");

    assert_eq!(pivot.filters.len(), 1);
    match &pivot.filters[0] {
        PivotFilter::Label {
            field,
            operator,
            value,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(*operator, PivotFilterOperator::Equals);
            assert_eq!(value, "East");
        }
        other => panic!("unexpected pivot filter after Excel re-save: {other:?}"),
    }

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let filter_collection = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 24
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL pivot label filter collection after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
        0,
        "Excel should preserve the zero-based Region source field index"
    );
    assert_eq!(
        u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
        4,
        "Excel should preserve the caption-equals filter type"
    );

    let custom_filter = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 24
                && payload[4] == 0x1D
                && payload[5] == 0x3C)
                .then_some(payload)
        })
        .expect("SXADDL custom label filter record after Excel re-save");
    assert_eq!(
        &custom_filter[20..24],
        &[0x06, 0x02, 0x01, 0x00],
        "Excel should preserve the equals discriminator"
    );

    let criterion = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 12
                && payload[4] == 0x1D
                && payload[5] == 0x3D)
                .then_some(payload)
        })
        .expect("SXADDL label criterion after Excel re-save");
    assert_eq!(xls_unicode_string_at(criterion, 12), "East");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_label_prefix_suffix_filters() {
    for (operator, value, filter_type, custom_value, pivot_name) in [
        (
            PivotFilterOperator::BeginsWith,
            "Ea",
            6u32,
            "Ea*",
            "LabelBegins",
        ),
        (
            PivotFilterOperator::EndsWith,
            "st",
            8u32,
            "*st",
            "LabelEnds",
        ),
    ] {
        let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(
            &xls_label_prefix_suffix_pivot_workbook(pivot_name, operator, value),
        );
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after Excel re-save");

        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator,
                value: value.to_string(),
            }]
        );

        let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
        let records = xls_record_payloads(&workbook);
        let filter_collection = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot label filter collection after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Excel should preserve the zero-based Region source field index"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom label filter record after Excel re-save");
        assert_eq!(
            &custom_filter[20..24],
            &[0x06, 0x02, 0x00, 0x00],
            "Excel should preserve wildcard equality metadata"
        );

        let criterion = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3D)
                    .then_some(payload)
            })
            .expect("SXADDL label criterion after Excel re-save");
        assert_eq!(xls_unicode_string_at(criterion, 12), custom_value);
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_negative_label_filters() {
    for (operator, value, filter_type, discriminator, custom_value, pivot_name) in [
        (
            PivotFilterOperator::NotEquals,
            "East",
            5u32,
            1u8,
            "East",
            "LabelNotEquals",
        ),
        (
            PivotFilterOperator::DoesNotBeginWith,
            "Ea",
            7u32,
            0u8,
            "Ea*",
            "LabelNotBegins",
        ),
        (
            PivotFilterOperator::DoesNotEndWith,
            "st",
            9u32,
            0u8,
            "*st",
            "LabelNotEnds",
        ),
        (
            PivotFilterOperator::DoesNotContain,
            "Ea",
            11u32,
            0u8,
            "*Ea*",
            "LabelNotContains",
        ),
    ] {
        let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(
            &xls_label_prefix_suffix_pivot_workbook(pivot_name, operator, value),
        );
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after Excel re-save");

        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator,
                value: value.to_string(),
            }]
        );

        let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
        let records = xls_record_payloads(&workbook);
        let filter_collection = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot label filter collection after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Excel should preserve the zero-based Region source field index"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom label filter record after Excel re-save");
        assert_eq!(
            &custom_filter[20..24],
            &[0x06, 0x05, discriminator, 0x00],
            "Excel should preserve notEqual metadata"
        );

        let criterion = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3D)
                    .then_some(payload)
            })
            .expect("SXADDL label criterion after Excel re-save");
        assert_eq!(xls_unicode_string_at(criterion, 12), custom_value);
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_label_comparison_filters() {
    for (operator, filter_type, custom_operator, pivot_name) in [
        (
            PivotFilterOperator::GreaterThan,
            12u32,
            0x04u8,
            "LabelGreater",
        ),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            13u32,
            0x06u8,
            "LabelGreaterEqual",
        ),
        (PivotFilterOperator::LessThan, 14u32, 0x01u8, "LabelLess"),
        (
            PivotFilterOperator::LessThanOrEqual,
            15u32,
            0x03u8,
            "LabelLessEqual",
        ),
    ] {
        let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(
            &xls_label_prefix_suffix_pivot_workbook(pivot_name, operator, "M"),
        );
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after Excel re-save");

        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator,
                value: "M".to_string(),
            }]
        );

        let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
        let records = xls_record_payloads(&workbook);
        let filter_collection = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot label filter collection after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Excel should preserve the zero-based Region source field index"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom label filter record after Excel re-save");
        assert_eq!(
            &custom_filter[20..24],
            &[0x06, custom_operator, 0x01, 0x00],
            "Excel should preserve comparison metadata"
        );

        let criterion = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3D)
                    .then_some(payload)
            })
            .expect("SXADDL label criterion after Excel re-save");
        assert_eq!(xls_unicode_string_at(criterion, 12), "M");
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_label_between_filters() {
    for (not_between, filter_type, and_flag, first_operator, second_operator, pivot_name) in [
        (false, 16u32, 1u32, 0x06u8, 0x03u8, "LabelBetween"),
        (true, 17u32, 2u32, 0x01u8, 0x04u8, "LabelNotBetween"),
    ] {
        let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(
            &xls_label_between_pivot_workbook(pivot_name, not_between),
        );
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after Excel re-save");

        assert_eq!(
            pivot.filters,
            vec![PivotFilter::LabelBetween {
                field: PivotFieldRef::new("Region"),
                start: "East".to_string(),
                end: "West".to_string(),
                not_between,
            }]
        );

        let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
        let records = xls_record_payloads(&workbook);
        let filter_collection = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot label filter collection after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Excel should preserve the zero-based Region source field index"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 44
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom label filter record after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            2,
            "Excel should preserve two label-range criteria"
        );
        assert_eq!(&custom_filter[20..24], &[0x06, first_operator, 0x01, 0x00]);
        assert_eq!(&custom_filter[30..34], &[0x06, second_operator, 0x01, 0x00]);
        assert_eq!(
            u32::from_le_bytes(custom_filter[40..44].try_into().unwrap()),
            and_flag
        );

        let start = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3D)
                    .then_some(payload)
            })
            .expect("SXADDL lower label criterion after Excel re-save");
        assert_eq!(xls_unicode_string_at(start, 12), "East");

        let end = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3E)
                    .then_some(payload)
            })
            .expect("SXADDL upper label criterion after Excel re-save");
        assert_eq!(xls_unicode_string_at(end, 12), "West");
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_value_comparison_filters() {
    for (operator, filter_type, custom_operator, threshold, pivot_name) in [
        (
            PivotFilterOperator::Equals,
            18u32,
            0x02u8,
            20.0,
            "ValueEquals",
        ),
        (
            PivotFilterOperator::NotEquals,
            19u32,
            0x05u8,
            20.0,
            "ValueNotEquals",
        ),
        (
            PivotFilterOperator::GreaterThan,
            20u32,
            0x04u8,
            20.0,
            "ValueGreater",
        ),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            21u32,
            0x06u8,
            20.0,
            "ValueGreaterEqual",
        ),
        (
            PivotFilterOperator::LessThan,
            22u32,
            0x01u8,
            30.0,
            "ValueLess",
        ),
        (
            PivotFilterOperator::LessThanOrEqual,
            23u32,
            0x03u8,
            30.0,
            "ValueLessEqual",
        ),
    ] {
        let filter_measure =
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
        let (result, _writer_bytes, excel_bytes) =
            roundtrip_through_excel_xls_bytes(&xls_value_filter_pivot_workbook(
                pivot_name,
                PivotFilter::Value {
                    field: PivotFieldRef::new("Region"),
                    measure: filter_measure.clone(),
                    operator,
                    value: threshold,
                },
            ));
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after Excel re-save");

        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Value {
                field: PivotFieldRef::new("Region"),
                measure: filter_measure,
                operator,
                value: threshold,
            }]
        );

        let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
        let records = xls_record_payloads(&workbook);
        let filter_collection = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 36
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL value filter collection after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Excel should preserve the zero-based Region source field index"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
            0,
            "Excel should preserve the value-filter data field index"
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[32..36].try_into().unwrap()),
            -1
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 30
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom value filter after Excel re-save");
        assert_eq!(
            custom_filter[20], 0x04,
            "custom filter should store a number"
        );
        assert_eq!(custom_filter[21], custom_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[22..30].try_into().unwrap()),
            threshold
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_column_value_filter() {
    let filter_measure =
        PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_column_value_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ColumnValueFilter")
        .expect("pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::Value {
            field: PivotFieldRef::new("Region"),
            measure: filter_measure,
            operator: PivotFilterOperator::GreaterThan,
            value: 25.0,
        }]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let filter_collection = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 36
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL value filter collection after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
        0,
        "Excel should preserve the column-axis source field index"
    );
    assert_eq!(
        i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
        0,
        "Excel should preserve the column value-filter data field index"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_second_measure_value_filter() {
    let filter_measure = PivotMeasure::new("Cost", PivotAggregate::Sum).with_name("Total Cost");
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_second_measure_value_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SecondMeasureValueFilter")
        .expect("pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::Value {
            field: PivotFieldRef::new("Region"),
            measure: filter_measure,
            operator: PivotFilterOperator::GreaterThan,
            value: 10.0,
        }]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let filter_collection = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 36
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL value filter collection after Excel re-save");
    assert_eq!(
        i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
        1,
        "Excel should preserve the second data field as the value-filter target"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_value_between_filters() {
    for (not_between, filter_type, first_operator, second_operator, flag, pivot_name) in [
        (false, 24u32, 0x06u8, 0x03u8, 1u32, "ValueBetween"),
        (true, 25u32, 0x01u8, 0x04u8, 2u32, "ValueNotBetween"),
    ] {
        let filter_measure =
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
        let (result, _writer_bytes, excel_bytes) =
            roundtrip_through_excel_xls_bytes(&xls_value_filter_pivot_workbook(
                pivot_name,
                PivotFilter::ValueBetween {
                    field: PivotFieldRef::new("Region"),
                    measure: filter_measure.clone(),
                    start: 15.0,
                    end: 35.0,
                    not_between,
                },
            ));
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after Excel re-save");

        assert_eq!(
            pivot.filters,
            vec![PivotFilter::ValueBetween {
                field: PivotFieldRef::new("Region"),
                measure: filter_measure,
                start: 15.0,
                end: 35.0,
                not_between,
            }]
        );

        let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
        let records = xls_record_payloads(&workbook);
        let filter_collection = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 36
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL value range filter collection after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
            0,
            "Excel should preserve the value-filter data field index"
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[32..36].try_into().unwrap()),
            -1
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 44
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom value range filter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            2,
            "Excel should preserve the two range criteria"
        );
        assert_eq!(custom_filter[20], 0x04);
        assert_eq!(custom_filter[21], first_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[22..30].try_into().unwrap()),
            15.0
        );
        assert_eq!(custom_filter[30], 0x04);
        assert_eq!(custom_filter[31], second_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[32..40].try_into().unwrap()),
            35.0
        );
        assert_eq!(
            u32::from_le_bytes(custom_filter[40..44].try_into().unwrap()),
            flag
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_date_comparison_filters() {
    for (operator, filter_type, custom_operator, custom_type, threshold, pivot_name) in [
        (
            PivotFilterOperator::Equals,
            26u32,
            0x02u8,
            4u32,
            44958.0,
            "DateEquals",
        ),
        (
            PivotFilterOperator::NotEquals,
            62u32,
            0x05u8,
            40u32,
            44958.0,
            "DateNotEquals",
        ),
        (
            PivotFilterOperator::GreaterThan,
            28u32,
            0x04u8,
            6u32,
            44958.0,
            "DateGreater",
        ),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            64u32,
            0x06u8,
            42u32,
            44958.0,
            "DateGreaterEqual",
        ),
        (
            PivotFilterOperator::LessThan,
            27u32,
            0x01u8,
            5u32,
            44986.0,
            "DateLess",
        ),
        (
            PivotFilterOperator::LessThanOrEqual,
            63u32,
            0x03u8,
            41u32,
            44986.0,
            "DateLessEqual",
        ),
    ] {
        let (result, _writer_bytes, excel_bytes) =
            roundtrip_through_excel_xls_bytes(&xls_date_filter_pivot_workbook(
                pivot_name,
                PivotFilter::Date {
                    field: PivotFieldRef::new("Order Date"),
                    operator,
                    value: threshold,
                },
            ));
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after Excel re-save");

        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Date {
                field: PivotFieldRef::new("Order Date"),
                operator,
                value: threshold,
            }]
        );

        let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
        let records = xls_record_payloads(&workbook);
        let filter_collection = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 36
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL date filter collection after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Excel should preserve the date source field index"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
            0,
            "Excel should preserve the label-style date filter target"
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[32..36].try_into().unwrap()),
            0
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 48
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom date filter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(custom_filter[12..16].try_into().unwrap()),
            custom_type
        );
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            1
        );
        assert_eq!(custom_filter[20], 0x04);
        assert_eq!(custom_filter[21], custom_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[22..30].try_into().unwrap()),
            threshold
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_date_between_filters() {
    for (
        not_between,
        filter_type,
        first_operator,
        second_operator,
        custom_type,
        flag,
        pivot_name,
    ) in [
        (false, 29u32, 0x06u8, 0x03u8, 7u32, 1u32, "DateBetween"),
        (true, 65u32, 0x01u8, 0x04u8, 43u32, 2u32, "DateNotBetween"),
    ] {
        let (result, _writer_bytes, excel_bytes) =
            roundtrip_through_excel_xls_bytes(&xls_date_filter_pivot_workbook(
                pivot_name,
                PivotFilter::DateBetween {
                    field: PivotFieldRef::new("Order Date"),
                    start: 44927.0,
                    end: 45016.0,
                    not_between,
                },
            ));
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after Excel re-save");

        assert_eq!(
            pivot.filters,
            vec![PivotFilter::DateBetween {
                field: PivotFieldRef::new("Order Date"),
                start: 44927.0,
                end: 45016.0,
                not_between,
            }]
        );

        let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
        let records = xls_record_payloads(&workbook);
        let filter_collection = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 36
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL date range filter collection after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
            0
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[32..36].try_into().unwrap()),
            0
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 48
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom date range filter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(custom_filter[12..16].try_into().unwrap()),
            custom_type
        );
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            2
        );
        assert_eq!(custom_filter[20], 0x04);
        assert_eq!(custom_filter[21], first_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[22..30].try_into().unwrap()),
            44927.0
        );
        assert_eq!(custom_filter[30], 0x04);
        assert_eq!(custom_filter[31], second_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[32..40].try_into().unwrap()),
            45016.0
        );
        assert_eq!(
            u32::from_le_bytes(custom_filter[40..44].try_into().unwrap()),
            flag
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_date_period_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_date_filter_pivot_workbook(
            "DatePeriod",
            PivotFilter::DatePeriod {
                field: PivotFieldRef::new("Order Date"),
                period: PivotDatePeriod::ThisMonth,
            },
        ));
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("DatePeriod")
        .expect("pivot after Excel re-save");

    assert_eq!(
        pivot.filters,
        vec![PivotFilter::DatePeriod {
            field: PivotFieldRef::new("Order Date"),
            period: PivotDatePeriod::ThisMonth,
        }]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let filter_collection = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 36
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL date-period filter collection after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
        37,
        "Excel should preserve the ThisMonth pivot filter type"
    );

    let custom_filter = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 20
                && payload[4] == 0x1D
                && payload[5] == 0x3C)
                .then_some(payload)
        })
        .expect("SXADDL date-period filter after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(custom_filter[12..16].try_into().unwrap()),
        0x0F
    );
    assert_eq!(
        u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
        2
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_row_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_row_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RowItemFilter")
        .expect("RowItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let row_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD after Excel re-save");
    let row_items = sxvi_payloads_for_xls_sxvd(&records, row_sxvd_position);
    assert_eq!(
        row_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Excel should keep only Central hidden for the row item filter"
    );
    assert_eq!(
        i16::from_le_bytes(row_items[2][4..6].try_into().unwrap()),
        2
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_row_item_filter_with_default_include_new_items() {
    let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(
        &xls_row_item_filter_default_include_new_items_pivot_workbook(),
    );
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RowItemFilterDefault")
        .expect("RowItemFilterDefault pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let row_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD after Excel re-save");
    let row_sxvdex = records[row_sxvd_position + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(row_sxvdex[0..4].try_into().unwrap()) & 0x8000,
        0,
        "Excel should keep fHideNewItems clear for this BIFF8 row filter"
    );
    let row_items = sxvi_payloads_for_xls_sxvd(&records, row_sxvd_position);
    assert_eq!(
        row_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Excel should keep only Central hidden for the row item filter"
    );
    assert_eq!(
        sxaddl_field_include_new_items_flags(&records, "Region"),
        Some(0x28),
        "Excel should preserve the SXADDL include-new-items=false field flag"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_column_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_column_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ColumnItemFilter")
        .expect("ColumnItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let column_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0002
        })
        .expect("column SXVD after Excel re-save");
    let column_items = sxvi_payloads_for_xls_sxvd(&records, column_sxvd_position);
    assert_eq!(
        column_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Excel should keep only Central hidden for the column item filter"
    );
    assert_eq!(
        i16::from_le_bytes(column_items[2][4..6].try_into().unwrap()),
        2
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_source_only_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_source_only_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SourceOnlyItemFilter")
        .expect("SourceOnlyItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Channel", ["Online"])]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let hidden_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0000
                && u16::from_le_bytes(payload[6..8].try_into().unwrap()) > 0
        })
        .expect("hidden source field SXVD after Excel re-save");
    let hidden_items = sxvi_payloads_for_xls_sxvd(&records, hidden_sxvd_position);
    assert_eq!(
        hidden_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 1, 0],
        "Excel should keep only Store hidden for the source-only Channel filter"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_grouped_row_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_grouped_row_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedRowItemFilter")
        .expect("GroupedRowItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal"])]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    assert!(
        sxvi_hidden_flag_groups_for_xls_sxvd_axis(&records, 0x0001)
            .iter()
            .any(|flags| flags == &[1, 0, 0]),
        "Excel should preserve the hidden ungrouped item on the derived row field"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_grouped_column_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_grouped_column_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedColumnItemFilter")
        .expect("GroupedColumnItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal"])]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    assert!(
        sxvi_hidden_flag_groups_for_xls_sxvd_axis(&records, 0x0002)
            .iter()
            .any(|flags| flags == &[1, 0, 0]),
        "Excel should preserve the hidden ungrouped item on the derived column field"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_axis_field_options() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_axis_options_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("AxisOptions")
        .expect("AxisOptions pivot after Excel re-save");
    let field = &pivot.rows[0];

    assert_eq!(field.caption.as_deref(), Some("Market"));
    assert_eq!(field.sort, PivotSort::Descending);
    assert_eq!(field.subtotal, PivotSubtotal::Sum);
    assert_eq!(
        field.subtotals,
        vec![
            PivotSubtotal::Sum,
            PivotSubtotal::Count,
            PivotSubtotal::Average,
            PivotSubtotal::Max,
            PivotSubtotal::Min,
            PivotSubtotal::Product,
            PivotSubtotal::CountNumbers,
            PivotSubtotal::StdDev,
            PivotSubtotal::StdDevP,
            PivotSubtotal::Var,
            PivotSubtotal::VarP
        ]
    );
    assert!(field.show_empty_items);
    assert!(!field.show_drop_downs);
    assert_eq!(field.subtotal_caption.as_deref(), Some("Regional subtotal"));
    assert!(field.insert_blank_row);
    assert!(!field.subtotal_top);
    assert!(field.insert_page_break);
    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let row_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD record after Excel re-save");
    let row_sxvd = &records[row_sxvd_position].1;
    assert_eq!(
        u16::from_le_bytes(row_sxvd[4..6].try_into().unwrap()),
        0x0FFE,
        "Excel should preserve custom subtotal function bits"
    );
    let row_sxvdex = records[row_sxvd_position + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX after Excel re-save");
    let grbit1 = u32::from_le_bytes(row_sxvdex[0..4].try_into().unwrap());
    assert_ne!(
        grbit1 & 0x0040_0000,
        0,
        "Excel should preserve the blank-row SXVDEX grbit1 flag"
    );
    assert_eq!(
        grbit1 & 0x0080_0000,
        0,
        "Excel should preserve the bottom-subtotal SXVDEX grbit1 flag"
    );
    assert_eq!(row_sxvdex[4], 0xFF);
    assert_eq!(row_sxvdex[5], 0xFF);
    assert_eq!(
        u16::from_le_bytes(row_sxvdex[10..12].try_into().unwrap()),
        17,
        "Excel should preserve custom subtotal caption length"
    );
    assert_eq!(&row_sxvdex[20..], b"\0Regional subtotal");
    let field_ver10 = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 12
                && payload[4] == 0x01
                && payload[5] == 0x02)
                .then_some(payload)
        })
        .expect("field SXADDL version-10 options after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(field_ver10[6..10].try_into().unwrap()) & 0x0000_0001,
        1,
        "Excel should preserve the hidden field dropdown"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_legacy_autosort_without_measure_target() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_sort_by_measure_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ValueSortedPivot")
        .expect("ValueSortedPivot pivot after Excel re-save");
    let field = &pivot.rows[0];

    assert_eq!(field.sort, PivotSort::Descending);
    assert!(field.sort_by_measure.is_none());

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let row_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD record after Excel re-save");
    let row_sxvdex = records[row_sxvd_position + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX after Excel re-save");
    assert_eq!(
        i16::from_le_bytes(row_sxvdex[6..8].try_into().unwrap()),
        -1,
        "Excel should preserve the legacy implicit sort field"
    );
}

// features: Pivot table styles
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_style() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_styled_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("StyledPivot")
        .expect("StyledPivot after Excel re-save");

    assert_eq!(pivot.style, xls_styled_pivot_style());

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let style_record = records
        .iter()
        .find_map(|(record_type, payload)| {
            if *record_type == 0x0864
                && payload.len() >= 16
                && matches!(
                    u16::from_le_bytes(payload[4..6].try_into().unwrap()),
                    0x001E | 0x1E00
                )
            {
                Some(payload)
            } else {
                None
            }
        })
        .expect("pivot style FRT record after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(style_record[12..14].try_into().unwrap()),
        0x002E,
        "Excel should preserve the pivot style flags"
    );
    assert_eq!(utf16le_string(&style_record[16..]), "PivotStyleLight16");
}

// features: Calculated fields
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_calculated_field() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_calculated_field_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRevenue")
        .expect("CalculatedRevenue after Excel re-save");

    assert_eq!(pivot.calculated_fields.len(), 1);
    assert_eq!(pivot.calculated_fields[0].name, "Revenue");
    assert_eq!(pivot.calculated_fields[0].formula, "Units*Price");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let records = xls_record_payloads(&cache);
    let calculated_field = records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Revenue")
        .expect("calculated-field SXFDB after Excel re-save");
    assert_ne!(
        u16::from_le_bytes(calculated_field[0..2].try_into().unwrap()) & 0x8000,
        0,
        "Excel should preserve the calculated-field formula flag"
    );
    assert!(
        records
            .iter()
            .any(|(record_type, _)| *record_type == 0x00F9),
        "Excel should preserve the SxFmla calculated-field formula"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_calculated_item() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_calculated_item_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRegion")
        .expect("CalculatedRegion after Excel re-save");

    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Region");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("Combined".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "East+West");

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let cache_records = xls_record_payloads(&cache);
    assert!(
        cache_records
            .iter()
            .any(|(record_type, _)| *record_type == 0x00F9),
        "Excel should preserve SxFmla for the calculated item"
    );
    assert!(
        cache_records
            .iter()
            .any(|(record_type, _)| *record_type == 0x00F8),
        "Excel should preserve SXPAIR item references"
    );
    assert!(
        cache_records
            .iter()
            .any(|(record_type, _)| *record_type == 0x0103),
        "Excel should preserve SXFORMULA"
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let row_sxvd_position = records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD record after Excel re-save");
    let row_items = sxvi_payloads_for_xls_sxvd(&records, row_sxvd_position);
    assert!(
        row_items
            .iter()
            .any(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()) & 0x0008 != 0),
        "Excel should preserve the calculated row item flag"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_calculated_item_cell_like_refs() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_calculated_item_cell_like_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedQuarter")
        .expect("CalculatedQuarter after Excel re-save");

    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Quarter");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("H1".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "Q1+Q2");

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let records = xls_record_payloads(&cache);
    let sxfmla = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F9).then_some(payload))
        .expect("SxFmla calculated-item cell-like refs after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxfmla[0..2].try_into().unwrap()),
        13,
        "Excel should preserve the calculated-item cell-like reference token stream"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_calculated_item_string_refs() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_calculated_item_string_ref_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRegionStringRef")
        .expect("CalculatedRegionStringRef after Excel re-save");

    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Region");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("Combined".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "East+West");

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let records = xls_record_payloads(&cache);
    let formula_lengths = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00F9).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        formula_lengths,
        vec![13],
        "Excel should preserve string item references as PtgSxName token streams"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_calculated_field_function() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_calculated_field_function_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRevenueFunction")
        .expect("CalculatedRevenueFunction after Excel re-save");

    assert_eq!(pivot.calculated_fields.len(), 1);
    assert_eq!(pivot.calculated_fields[0].name, "Revenue");
    assert_eq!(pivot.calculated_fields[0].formula, "SUM(Units,Price)");

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let records = xls_record_payloads(&cache);
    let sxfmla = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F9).then_some(payload))
        .expect("SxFmla calculated-field function after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxfmla[0..2].try_into().unwrap()),
        16,
        "Excel should preserve the calculated-field function token stream"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_calculated_item_function() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_calculated_item_function_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRegionFunction")
        .expect("CalculatedRegionFunction after Excel re-save");

    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Region");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("Combined".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "MAX(East,West)");

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let records = xls_record_payloads(&cache);
    let sxfmla = records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F9).then_some(payload))
        .expect("SxFmla calculated-item function after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(sxfmla[0..2].try_into().unwrap()),
        16,
        "Excel should preserve the calculated-item function token stream"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_multi_measure_values_axis() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_multi_measure_pivot_workbook());
    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let sxvd_axes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxvd_axes, vec![0x0001, 0x0008, 0x0008]);
    assert_eq!(
        records
            .iter()
            .filter(|(record_type, _)| *record_type == 0x00C5)
            .count(),
        2,
        "Excel should preserve both data fields"
    );
    assert!(
        records
            .iter()
            .any(|(record_type, payload)| *record_type == 0x00B4 && payload == &[0xFE, 0xFF]),
        "Excel should preserve the synthetic Values column axis"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_all_aggregate_data_fields() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_all_aggregate_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("AllAggregatePivot")
        .unwrap();
    assert_eq!(pivot.layout.values_axis, PivotValuesAxis::Columns);
    assert_eq!(pivot.layout.values_axis_position, Some(0));
    assert_eq!(pivot.measures.len(), xls_all_aggregate_cases().len());
    for (measure, (aggregate, caption, _)) in pivot.measures.iter().zip(xls_all_aggregate_cases()) {
        assert_eq!(measure.field.name, "Revenue");
        assert_eq!(measure.aggregate, aggregate);
        assert_eq!(measure.name.as_deref(), Some(caption));
    }

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let data_field_codes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00C5).then(|| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    let expected_codes = xls_all_aggregate_cases()
        .iter()
        .map(|(_, _, code)| *code)
        .collect::<Vec<_>>();
    assert_eq!(
        data_field_codes, expected_codes,
        "Excel should preserve every BIFF8 pivot aggregate function code"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_custom_measure_number_format() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_custom_measure_format_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("FormattedRevenue")
        .expect("FormattedRevenue after Excel re-save");

    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].number_format.as_deref(), Some("#,##0.0"));

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let custom_formats = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x041E).then(|| {
                let id = u16::from_le_bytes(payload[0..2].try_into().unwrap());
                let code = xls_unicode_string_at(payload, 2);
                (id, code)
            })
        })
        .collect::<std::collections::HashMap<_, _>>();
    let num_fmt_id = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x00C5)
                .then(|| u16::from_le_bytes(payload[10..12].try_into().unwrap()))
        })
        .expect("pivot data-field number format after Excel re-save");
    assert_eq!(
        custom_formats.get(&num_fmt_id).map(String::as_str),
        Some("#,##0.0")
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_show_as_data_fields() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_show_as_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ShowAsPivot")
        .expect("ShowAsPivot after Excel re-save");

    assert_eq!(pivot.measures.len(), 4);
    assert_eq!(pivot.measures[0].show_as, PivotShowAs::PercentOfRowTotal);
    assert_eq!(
        pivot.measures[1].show_as,
        PivotShowAs::RunningTotal {
            base_field: PivotFieldRef::new("Region")
        }
    );
    assert_eq!(
        pivot.measures[2].show_as,
        PivotShowAs::DifferenceFrom {
            base_field: PivotFieldRef::new("Region"),
            base_item: PivotValue::String("East".to_string())
        }
    );
    assert_eq!(
        pivot.measures[3].show_as,
        PivotShowAs::PercentDifferenceFrom {
            base_field: PivotFieldRef::new("Region"),
            base_item: PivotValue::String("East".to_string())
        }
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let show_as_codes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00C5).then(|| u16::from_le_bytes(payload[4..6].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(show_as_codes, vec![5, 4, 1, 3]);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_numeric_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_numeric_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("AgeBands")
        .unwrap();
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Number {
            field,
            start,
            end,
            interval,
        } => {
            assert_eq!(field.name, "Age");
            assert_eq!(*start, Some(0.0));
            assert_eq!(*end, Some(60.0));
            assert_eq!(*interval, 10.0);
        }
        other => panic!("expected numeric grouping, got {other:?}"),
    }

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let records = xls_record_payloads(&cache);
    let sxrng_position = records
        .iter()
        .position(|(record_type, _)| *record_type == 0x00D8)
        .expect("Excel should preserve an SXRng grouping record");
    assert!(
        records[sxrng_position + 1..]
            .iter()
            .filter(|(record_type, _)| *record_type == 0x00C9)
            .take(3)
            .count()
            == 3,
        "Excel should preserve SXNum start/end/interval records after SXRng"
    );
}

// features: Grouping (dates, numbers, items)
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_date_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_date_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("MonthlyRevenue")
        .unwrap();
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "Date");
            assert_eq!(*units, vec![PivotDateGroupUnit::Months]);
        }
        other => panic!("expected date grouping, got {other:?}"),
    }

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let sxrng = xls_record_payloads(&cache)
        .into_iter()
        .find_map(|(record_type, payload)| (record_type == 0x00D8).then_some(payload))
        .expect("Excel should preserve an SXRng grouping record");
    assert_eq!(
        u16::from_le_bytes(sxrng[0..2].try_into().unwrap()),
        0x0017,
        "Excel should preserve month date grouping flags"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_multi_unit_date_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_multi_unit_date_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedDates")
        .unwrap();
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "SaleDate");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "SaleDate");
            assert_eq!(
                *units,
                vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months]
            );
        }
        other => panic!("expected multi-unit date grouping, got {other:?}"),
    }

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    assert!(
        records
            .iter()
            .any(|(record_type, payload)| *record_type == 0x00B4 && payload == &[3, 0, 2, 0]),
        "Excel should preserve the year-plus-month row axis declaration"
    );

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let range_group_flags = xls_record_payloads(&cache)
        .into_iter()
        .filter_map(|(record_type, payload)| {
            (record_type == 0x00D8).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        range_group_flags,
        vec![0x0017, 0x001F],
        "Excel should preserve month and year date grouping flags"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_manual_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_manual_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualGroupedRegions")
        .unwrap();
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Manual { field, groups } => {
            assert_eq!(field.name, "Region");
            assert_eq!(groups.len(), 2);
            assert_eq!(groups[0].name, "Coastal");
            assert_eq!(
                groups[0].members,
                vec![
                    PivotValue::String("East".to_string()),
                    PivotValue::String("West".to_string()),
                    PivotValue::String("South".to_string())
                ]
            );
            assert_eq!(groups[1].name, "Interior");
            assert_eq!(
                groups[1].members,
                vec![
                    PivotValue::String("Central".to_string()),
                    PivotValue::String("North".to_string())
                ]
            );
        }
        other => panic!("expected manual grouping, got {other:?}"),
    }

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    assert!(
        records
            .iter()
            .any(|(record_type, payload)| *record_type == 0x00B4 && payload == &[2, 0, 0, 0]),
        "Excel should preserve the derived-plus-source row axis declaration"
    );

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let cache_records = xls_record_payloads(&cache);
    let derived = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| String::from_utf8_lossy(payload).contains("Region2"))
        .expect("Excel should preserve the derived manual grouped field");
    assert_eq!(
        u16::from_le_bytes(derived[0..2].try_into().unwrap()),
        0x0001,
        "Excel should preserve the derived manual field flags"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_numeric_manual_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_numeric_manual_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualAgeGroups")
        .unwrap();
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Age");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Manual { field, groups } => {
            assert_eq!(field.name, "Age");
            assert_eq!(groups.len(), 2);
            assert_eq!(groups[0].name, "Young");
            assert_eq!(
                groups[0].members,
                vec![PivotValue::Number(7.0), PivotValue::Number(13.0)]
            );
            assert_eq!(groups[1].name, "Adult");
            assert_eq!(
                groups[1].members,
                vec![PivotValue::Number(21.0), PivotValue::Number(34.0)]
            );
        }
        other => panic!("expected numeric manual grouping, got {other:?}"),
    }

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let cache_records = xls_record_payloads(&cache);
    let age_index = cache_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00C7 && String::from_utf8_lossy(payload).contains("Age")
        })
        .expect("Excel should preserve Age cache field");
    let revenue_index = cache_records[age_index + 1..]
        .iter()
        .position(|(record_type, _)| *record_type == 0x00C7)
        .map(|offset| age_index + 1 + offset)
        .expect("Excel should preserve following cache field");
    let age_numbers = cache_records[age_index + 1..revenue_index]
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00C9).then(|| f64::from_le_bytes(payload[0..8].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        age_numbers,
        vec![7.0, 13.0, 21.0, 34.0, 55.0],
        "Excel should preserve numeric manual source items"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_bool_error_manual_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_bool_error_manual_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualBoolErrorGroups")
        .unwrap();
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Flag");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Manual { field, groups } => {
            assert_eq!(field.name, "Flag");
            assert_eq!(groups.len(), 2);
            assert_eq!(groups[0].name, "Booleans");
            assert_eq!(
                groups[0].members,
                vec![PivotValue::Boolean(true), PivotValue::Boolean(false)]
            );
            assert_eq!(groups[1].name, "Errors");
            assert_eq!(
                groups[1].members,
                vec![
                    PivotValue::Error(CellError::Na),
                    PivotValue::Error(CellError::Div0)
                ]
            );
        }
        other => panic!("expected bool/error manual grouping, got {other:?}"),
    }

    let cache = xls_cfb_stream(&excel_bytes, "/_SX_DB_CUR/0001");
    let cache_records = xls_record_payloads(&cache);
    let flag_index = cache_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00C7 && String::from_utf8_lossy(payload).contains("Flag")
        })
        .expect("Excel should preserve Flag cache field");
    let revenue_index = cache_records[flag_index + 1..]
        .iter()
        .position(|(record_type, _)| *record_type == 0x00C7)
        .map(|offset| flag_index + 1 + offset)
        .expect("Excel should preserve following cache field");
    let flag_items = cache_records[flag_index + 1..revenue_index]
        .iter()
        .filter_map(|(record_type, payload)| match *record_type {
            0x00CA | 0x00CB => Some((*record_type, payload.clone())),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(
        flag_items,
        vec![
            (0x00CA, vec![1, 0]),
            (0x00CA, vec![0, 0]),
            (0x00CB, vec![0x2A, 0]),
            (0x00CB, vec![0x07, 0]),
        ],
        "Excel should preserve SXBool/SXErr source items"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_manual_column_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_manual_column_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualColumnGroups")
        .unwrap();
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Quarter");
    assert_eq!(pivot.columns.len(), 1);
    assert_eq!(pivot.columns[0].field.name, "Region");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Manual { field, groups } => {
            assert_eq!(field.name, "Region");
            assert_eq!(groups.len(), 2);
            assert_eq!(groups[0].name, "Coastal");
            assert_eq!(
                groups[0].members,
                vec![
                    PivotValue::String("East".to_string()),
                    PivotValue::String("West".to_string()),
                    PivotValue::String("South".to_string())
                ]
            );
            assert_eq!(groups[1].name, "Interior");
            assert_eq!(
                groups[1].members,
                vec![
                    PivotValue::String("Central".to_string()),
                    PivotValue::String("North".to_string())
                ]
            );
        }
        other => panic!("expected manual column grouping, got {other:?}"),
    }

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let axis_fields = xls_record_payloads(&workbook)
        .into_iter()
        .filter_map(|(record_type, payload)| (record_type == 0x00B4).then_some(payload))
        .collect::<Vec<_>>();
    assert!(
        axis_fields.iter().any(|payload| payload == &[3, 0, 0, 0]),
        "Excel should preserve the expanded manual column axis declaration"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_manual_page_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xls_bytes(&xls_manual_page_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualPageRegions")
        .unwrap();
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Salesperson");
    assert_eq!(pivot.page_fields.len(), 1);
    assert_eq!(pivot.page_fields[0].field.name, "Region");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Manual { field, groups } => {
            assert_eq!(field.name, "Region");
            assert_eq!(groups.len(), 2);
            assert_eq!(groups[0].name, "Coastal");
            assert_eq!(
                groups[0].members,
                vec![
                    PivotValue::String("East".to_string()),
                    PivotValue::String("West".to_string()),
                    PivotValue::String("South".to_string())
                ]
            );
            assert_eq!(groups[1].name, "Interior");
            assert_eq!(
                groups[1].members,
                vec![
                    PivotValue::String("Central".to_string()),
                    PivotValue::String("North".to_string())
                ]
            );
        }
        other => panic!("expected manual page grouping, got {other:?}"),
    }
    assert!(pivot.filters.iter().any(|filter| matches!(
        filter,
        PivotFilter::FieldItems {
            field,
            allowed_items,
        } if field.name == "Region"
            && allowed_items == &vec![PivotValue::String("Coastal".to_string())]
    )));

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    let sxvd_axes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        sxvd_axes,
        vec![0x0000, 0x0001, 0x0008, 0x0004],
        "Excel should preserve source-hidden, derived-page manual grouping axes"
    );
    assert!(
        records.iter().any(|(record_type, payload)| {
            *record_type == 0x00B6 && payload == &[3, 0, 1, 0, 1, 0]
        }),
        "Excel should preserve SXPI on the derived manual group field"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xls_pivot_manual_page_grouping_multi_item_filter() {
    let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(
        &xls_manual_page_grouped_pivot_workbook_with_filter_items(&["Coastal", "Interior"]),
    );
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualPageRegions")
        .unwrap();
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal", "Interior"])]
    );

    let workbook = xls_cfb_stream(&excel_bytes, "/Workbook");
    let records = xls_record_payloads(&workbook);
    assert!(
        records.iter().any(|(record_type, payload)| {
            *record_type == 0x00B6 && payload == &[3, 0, 0xFD, 0x7F, 1, 0]
        }),
        "Excel should preserve the BIFF8 multi-item sentinel on the derived manual group field"
    );
    assert!(
        sxvi_hidden_flag_groups_for_xls_sxvd_axis(&records, 0x0004)
            .iter()
            .any(|flags| flags == &[1, 0, 0, 0]),
        "Excel should preserve hidden flags on the derived manual page field"
    );
}

fn xls_cfb_has_stream(bytes: &[u8], path: &str) -> bool {
    let reader = std::io::Cursor::new(bytes);
    let Ok(cfb) = duke_sheets_xls::cfb::CompoundFile::open(reader) else {
        return false;
    };
    cfb.exists(path)
}

fn xls_cfb_stream(bytes: &[u8], path: &str) -> Vec<u8> {
    let reader = std::io::Cursor::new(bytes);
    let cfb = duke_sheets_xls::cfb::CompoundFile::open(reader).expect("open xls cfb");
    cfb.read_stream(path)
        .unwrap_or_else(|e| panic!("read {path}: {e}"))
}

fn xls_record_payloads(stream: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    while pos + 4 <= stream.len() {
        let typ = u16::from_le_bytes([stream[pos], stream[pos + 1]]);
        let len = u16::from_le_bytes([stream[pos + 2], stream[pos + 3]]) as usize;
        let start = pos + 4;
        let end = start + len;
        out.push((typ, stream[start..end].to_vec()));
        pos = end;
    }
    out
}

fn sxvi_payloads_for_xls_sxvd(records: &[(u16, Vec<u8>)], sxvd_index: usize) -> Vec<Vec<u8>> {
    records[sxvd_index + 1..]
        .iter()
        .take_while(|(record_type, _)| *record_type == 0x00B2)
        .map(|(_, payload)| payload.clone())
        .collect()
}

fn sxaddl_field_include_new_items_flags(
    records: &[(u16, Vec<u8>)],
    field_name: &str,
) -> Option<u8> {
    let mut pending_field_name = None;
    for (record_type, payload) in records {
        if *record_type != 0x0864 || payload.len() < 12 {
            continue;
        }
        match (payload[4], payload[5]) {
            (0x17, 0x00) => pending_field_name = Some(xls_unicode_string_at(payload, 12)),
            (0x17, 0x19)
                if pending_field_name
                    .as_deref()
                    .is_some_and(|name| name.eq_ignore_ascii_case(field_name)) =>
            {
                return Some(payload[6]);
            }
            (0x17, 0xFF) => pending_field_name = None,
            _ => {}
        }
    }
    None
}

fn sxvi_hidden_flag_groups_for_xls_sxvd_axis(
    records: &[(u16, Vec<u8>)],
    axis: u16,
) -> Vec<Vec<u16>> {
    let mut in_target_field = false;
    let mut current: Option<Vec<u16>> = None;
    let mut groups = Vec::new();
    for (record_type, payload) in records {
        match *record_type {
            0x00B1 => {
                in_target_field = u16::from_le_bytes(payload[0..2].try_into().unwrap()) == axis;
                current = in_target_field.then(Vec::new);
            }
            0x00B2 if in_target_field => {
                if let Some(current) = &mut current {
                    current.push(u16::from_le_bytes(payload[2..4].try_into().unwrap()));
                }
            }
            0x0100 => {
                if let Some(current) = current.take() {
                    groups.push(current);
                }
                in_target_field = false;
            }
            _ => {}
        }
    }
    groups
}

fn xls_unicode_string_at(bytes: &[u8], offset: usize) -> String {
    if offset + 3 > bytes.len() {
        return String::new();
    }
    let char_count = u16::from_le_bytes([bytes[offset], bytes[offset + 1]]) as usize;
    let flags = bytes[offset + 2];
    let start = offset + 3;
    if flags & 0x01 != 0 {
        let byte_len = char_count * 2;
        if start + byte_len > bytes.len() {
            return String::new();
        }
        utf16le_string(&bytes[start..start + byte_len])
    } else {
        if start + char_count > bytes.len() {
            return String::new();
        }
        bytes[start..start + char_count]
            .iter()
            .map(|&byte| byte as char)
            .collect()
    }
}

fn utf16le_string(bytes: &[u8]) -> String {
    let units = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16_lossy(&units)
}
