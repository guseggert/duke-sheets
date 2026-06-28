#![allow(clippy::approx_constant)]
//! Round-trip tests for the XLS skeleton writer.
//!
//! Build an empty `Workbook`, write it to BIFF8 bytes, read it back
//! through `XlsReader`, and confirm the structure (sheet count, sheet
//! names) round-trips. The skeleton writer doesn't yet emit cells,
//! formatting, or formulas — those land in subsequent slices.

use std::io::Cursor;

use duke_sheets_core::table::{Table, TableColumn};
use duke_sheets_core::{
    CellError, CellRange, PivotAggregate, PivotCacheSourceKind, PivotDateGroupUnit,
    PivotDatePeriod, PivotField, PivotFieldRef, PivotFilter, PivotFilterOperator, PivotGrouping,
    PivotManualGroup, PivotMeasure, PivotShowAs, PivotSort, PivotSource, PivotSourceRange,
    PivotStyle, PivotSubtotal, PivotTable, PivotValue, PivotValuesAxis, Workbook,
    WorkbookConnection,
};
use duke_sheets_xls::{
    cfb::{CompoundFile, CompoundFileBuilder},
    XlsReader, XlsWriter,
};

fn add_test_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_table_source_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_consolidation_source_pivot(wb: &mut Workbook) {
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
}

fn add_consolidation_source_pivot_with_page_items(wb: &mut Workbook) {
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
}

fn add_named_consolidation_source_pivot(wb: &mut Workbook) {
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
}

fn add_external_consolidation_source_pivot(wb: &mut Workbook) {
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
}

fn add_unnamed_scenario_source_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();

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
}

const EXTERNAL_PIVOT_COMMAND: &str = "select * from [pivot_external_sales.csv]";
const EXTERNAL_PIVOT_CONNECTION: &str =
    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\temp;Extended Properties=\"text;HDR=Yes;FMT=Delimited\";";

fn add_external_source_pivot(wb: &mut Workbook) {
    wb.add_data_connection(
        WorkbookConnection::database(7, "SalesConnection", EXTERNAL_PIVOT_CONNECTION)
            .with_command(EXTERNAL_PIVOT_COMMAND)
            .with_command_type(2),
    )
    .unwrap();

    let pivot = PivotTable::builder("ExternalPivot")
        .source(PivotSource::External {
            connection_name: "SalesConnection".to_string(),
            command_text: Some(EXTERNAL_PIVOT_COMMAND.to_string()),
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn include_new_items_field(name: &str) -> PivotField {
    let mut field = PivotField::new(name);
    field.include_new_items_in_filter = true;
    field
}

fn add_average_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_column_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_page_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_top_n_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_percent_top_n_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_axis_options_pivot(wb: &mut Workbook) {
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
    region.item_page_count = 25;

    let pivot = PivotTable::builder("AxisOptions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row(region)
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_sort_by_measure_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_multi_measure_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
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

fn add_all_aggregate_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_show_as_pivot(wb: &mut Workbook) {
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
                .with_name("Pct Col")
                .with_show_as(PivotShowAs::PercentOfColumnTotal),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Pct Total")
                .with_show_as(PivotShowAs::PercentOfGrandTotal),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Index")
                .with_show_as(PivotShowAs::Index),
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_custom_measure_format_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_page_column_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Segment").unwrap();
    ws.set_cell_value("B1", "Region").unwrap();
    ws.set_cell_value("C1", "Quarter").unwrap();
    ws.set_cell_value("D1", "Revenue").unwrap();
    ws.set_cell_value("A2", "Online").unwrap();
    ws.set_cell_value("B2", "East").unwrap();
    ws.set_cell_value("C2", "Q1").unwrap();
    ws.set_cell_value("D2", 10.0).unwrap();
    ws.set_cell_value("A3", "Retail").unwrap();
    ws.set_cell_value("B3", "East").unwrap();
    ws.set_cell_value("C3", "Q2").unwrap();
    ws.set_cell_value("D3", 20.0).unwrap();
    ws.set_cell_value("A4", "Online").unwrap();
    ws.set_cell_value("B4", "West").unwrap();
    ws.set_cell_value("C4", "Q1").unwrap();
    ws.set_cell_value("D4", 30.0).unwrap();

    let pivot = PivotTable::builder("ChannelPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F2")
        .unwrap()
        .page(include_new_items_field("Segment"))
        .row("Region")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::field_items("Segment", ["Online"]))
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
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

fn add_styled_pivot(wb: &mut Workbook) {
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
        .style(styled_pivot_style())
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_calculated_field_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_calculated_item_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_calculated_item_cell_like_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_calculated_item_string_ref_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_calculated_item_dependent_string_ref_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", 5.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRegionStringRef")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_item("Region", "All Regions", "\"Combined\"+Central")
        .calculated_item("Region", "Combined", "East+West")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_calculated_field_function_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_calculated_item_function_pivot(wb: &mut Workbook) {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_numeric_grouped_pivot(wb: &mut Workbook) {
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
            field: PivotFieldRef::new("Age"),
            start: Some(0.0),
            end: Some(60.0),
            interval: 10.0,
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_date_grouped_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Date").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 43831.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 43862.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", 43891.0).unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", 43922.0).unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder("MonthlyRevenue")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Date")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: PivotFieldRef::new("Date"),
            units: vec![PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_multi_unit_date_grouped_pivot(wb: &mut Workbook) {
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
            field: PivotFieldRef::new("SaleDate"),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_duplicate_date_grouped_pivot(wb: &mut Workbook) {
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

    let pivot = PivotTable::builder("DuplicateMonthlyRevenue")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Date")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: PivotFieldRef::new("Date"),
            units: vec![PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_page_date_grouped_pivot(wb: &mut Workbook) {
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Date").unwrap();
    ws.set_cell_value("B1", "Region").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", 43831.0).unwrap();
    ws.set_cell_value("B2", "East").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", 43862.0).unwrap();
    ws.set_cell_value("B3", "West").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", 43891.0).unwrap();
    ws.set_cell_value("B4", "East").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("MonthlyPageFilter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E2")
        .unwrap()
        .page(include_new_items_field("Date"))
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::field_items("Date", ["Jan"]))
        .grouping(PivotGrouping::Date {
            field: PivotFieldRef::new("Date"),
            units: vec![PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_manual_grouped_pivot(wb: &mut Workbook) {
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
            field: PivotFieldRef::new("Region"),
            groups: vec![
                PivotManualGroup::new("Coastal", ["East", "West", "South"]),
                PivotManualGroup::new("Interior", ["Central", "North"]),
            ],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_manual_numeric_grouped_pivot(wb: &mut Workbook) {
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
            field: PivotFieldRef::new("Age"),
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_manual_bool_error_grouped_pivot(wb: &mut Workbook) {
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
            field: PivotFieldRef::new("Flag"),
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_manual_column_grouped_pivot(wb: &mut Workbook) {
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
            field: PivotFieldRef::new("Region"),
            groups: vec![
                PivotManualGroup::new("Coastal", ["East", "West", "South"]),
                PivotManualGroup::new("Interior", ["Central", "North"]),
            ],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

fn add_manual_page_grouped_pivot(wb: &mut Workbook) {
    add_manual_page_grouped_pivot_with_filter_items(wb, Some(&["Coastal"]));
}

fn add_unfiltered_manual_page_grouped_pivot(wb: &mut Workbook) {
    add_manual_page_grouped_pivot_with_filter_items(wb, None);
}

fn add_manual_page_grouped_pivot_with_filter_items(
    wb: &mut Workbook,
    filter_items: Option<&[&str]>,
) {
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

    let mut builder = PivotTable::builder("ManualPageRegions")
        .source_range(CellRange::parse("A1:C7").unwrap())
        .target_address("E2")
        .unwrap()
        .row("Salesperson")
        .page(include_new_items_field("Region"))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: PivotFieldRef::new("Region"),
            groups: vec![
                PivotManualGroup::new("Coastal", ["East", "West", "South"]),
                PivotManualGroup::new("Interior", ["Central", "North"]),
            ],
        });
    if let Some(filter_items) = filter_items {
        builder = builder.filter(PivotFilter::field_items(
            "Region",
            filter_items.iter().copied(),
        ));
    }
    let pivot = builder.build().unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
}

#[test]
fn empty_default_workbook_round_trips_via_reader() {
    let wb = Workbook::new();
    let original_count = wb.sheet_count();
    let original_name = wb.worksheet(0).unwrap().name().to_string();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize empty workbook");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back via XlsReader");

    assert_eq!(parsed.sheet_count(), original_count);
    assert_eq!(parsed.worksheet(0).unwrap().name(), original_name);
}

#[test]
fn semantic_pivot_tables_emit_native_biff8_streams() {
    let mut wb = Workbook::new();
    add_test_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    assert!(cfb.exists("/_SX_DB_CUR/0001"));

    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = record_types(&workbook);
    assert!(workbook_records.contains(&0x00D5));
    assert!(workbook_records.contains(&0x00E3));
    assert!(workbook_records.contains(&0x0051));
    assert!(workbook_records.contains(&0x0160));
    assert!(workbook_records.contains(&0x089A));
    assert!(workbook_records.contains(&0x0802));
    assert!(workbook_records.contains(&0x0864));
    assert!(
        record_position(&workbook_records, 0x00D5) < record_position(&workbook_records, 0x0085),
        "SXIDSTM must be emitted before BoundSheet8 so Excel can bind the cache stream"
    );

    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = record_types(&cache);
    assert!(cache_records.contains(&0x00C6));
    assert_eq!(
        cache_records.iter().filter(|&&typ| typ == 0x00C7).count(),
        2
    );
    assert!(cache_records.contains(&0x00CD));
    assert!(cache_records.contains(&0x00C9));

    let sxdb = records_with_payload(&cache)
        .into_iter()
        .find_map(|(record_type, payload)| (record_type == 0x00C6).then_some(payload))
        .expect("SXDB payload");
    assert_eq!(
        u16::from_le_bytes(sxdb[16..18].try_into().unwrap()),
        0x0001,
        "SXDB.vsType should identify worksheet sources"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_table_source_dconname() {
    let mut wb = Workbook::new();
    add_table_source_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize table-source pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    assert!(workbook_records.iter().any(|(typ, _)| *typ == 0x00D5));
    assert!(!workbook_records.iter().any(|(typ, _)| *typ == 0x0051));
    let dconname = workbook_records
        .iter()
        .find_map(|(typ, payload)| (*typ == 0x0052).then_some(payload))
        .expect("DCONNAME table source record");
    let char_count = u16::from_le_bytes(dconname[0..2].try_into().unwrap()) as usize;
    assert_eq!(dconname[2], 0x00);
    let name_start = 3;
    let name_end = name_start + char_count;
    assert_eq!(&dconname[name_start..name_end], b"SalesData");
    assert_eq!(
        u16::from_le_bytes(dconname[name_end..name_end + 2].try_into().unwrap()),
        0,
        "table source should be workbook-scoped"
    );
}

#[test]
fn reads_writer_xls_table_source_semantics() {
    let mut wb = Workbook::new();
    add_table_source_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize table-source pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TablePivot")
        .unwrap();

    assert!(matches!(
        &pivot.source,
        PivotSource::Table { name } if name == "SalesData"
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

#[test]
fn semantic_pivot_tables_emit_xls_consolidation_source_dconrefs() {
    let mut wb = Workbook::new();
    add_consolidation_source_pivot(&mut wb);

    let bytes =
        XlsWriter::write_to_bytes(&wb).expect("serialize consolidation-source pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    assert_eq!(
        workbook_records
            .iter()
            .filter(|(typ, _)| *typ == 0x0051)
            .count(),
        2,
        "one DCONREF should be emitted for each consolidation range"
    );
    let sxtbl = workbook_records
        .iter()
        .find_map(|(typ, payload)| (*typ == 0x00D0).then_some(payload))
        .expect("SXTBL multiple-consolidation source record");
    assert_eq!(
        sxtbl,
        &[2, 0, 2, 0, 0, 0],
        "SXTBL should declare two sources and two empty page-index records"
    );
    assert_eq!(
        workbook_records
            .iter()
            .filter(|(typ, _)| *typ == 0x00D2)
            .count(),
        2,
        "one SXTBPG page-index record should follow each DCONREF"
    );

    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let sxdb = records_with_payload(&cache)
        .into_iter()
        .find_map(|(record_type, payload)| (record_type == 0x00C6).then_some(payload))
        .expect("SXDB payload");
    assert_eq!(
        u16::from_le_bytes(sxdb[16..18].try_into().unwrap()),
        0x0004,
        "SXDB.vsType should identify consolidation sources"
    );
}

#[test]
fn reads_writer_xls_consolidation_source_semantics() {
    let mut wb = Workbook::new();
    add_consolidation_source_pivot(&mut wb);

    let bytes =
        XlsWriter::write_to_bytes(&wb).expect("serialize consolidation-source pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ConsolidatedPivot")
        .unwrap();

    match &pivot.source {
        PivotSource::Consolidation { ranges } => {
            assert_eq!(ranges.len(), 2);
            assert_eq!(ranges[0].sheet.as_deref(), Some("Sheet1"));
            assert_eq!(ranges[0].range, Some(CellRange::parse("A1:B3").unwrap()));
            assert_eq!(ranges[1].sheet.as_deref(), Some("WestData"));
            assert_eq!(ranges[1].range, Some(CellRange::parse("A1:B2").unwrap()));
        }
        other => panic!("expected consolidation source, got {other:?}"),
    }
    assert_eq!(pivot.rows[0].field.name, "Row");
    assert_eq!(pivot.columns[0].field.name, "Column");
    assert_eq!(pivot.measures[0].field.name, "Value");
    assert_eq!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(PivotCacheSourceKind::Consolidation)
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_consolidation_page_item_records() {
    let mut wb = Workbook::new();
    add_consolidation_source_pivot_with_page_items(&mut wb);

    let bytes =
        XlsWriter::write_to_bytes(&wb).expect("serialize consolidation-source pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let sxtbl = workbook_records
        .iter()
        .find_map(|(typ, payload)| (*typ == 0x00D0).then_some(payload))
        .expect("SXTBL multiple-consolidation source record");
    assert_eq!(
        sxtbl,
        &[2, 0, 2, 0, 1, 0],
        "SXTBL should declare two sources and one page field"
    );

    let page_item_refs = workbook_records
        .iter()
        .filter_map(|(typ, payload)| (*typ == 0x00D2).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(page_item_refs, vec![vec![0, 0], vec![1, 0]]);

    let page_item_count = workbook_records
        .iter()
        .find_map(|(typ, payload)| (*typ == 0x00D1).then_some(payload))
        .expect("SXTBRGIITM page item count record");
    assert_eq!(page_item_count, &[2, 0]);

    let page_item_names = workbook_records
        .iter()
        .filter_map(|(typ, payload)| (*typ == 0x00CD).then(|| xls_unicode_string_at(payload, 0)))
        .collect::<Vec<_>>();
    assert!(
        page_item_names
            .windows(2)
            .any(|names| names == ["Retail", "Wholesale"]),
        "SXTBRGIITM should be followed by Retail/Wholesale SXSTRING page labels, got {page_item_names:?}"
    );
}

#[test]
fn reads_writer_xls_consolidation_page_item_semantics() {
    let mut wb = Workbook::new();
    add_consolidation_source_pivot_with_page_items(&mut wb);

    let bytes =
        XlsWriter::write_to_bytes(&wb).expect("serialize consolidation-source pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ConsolidatedPagePivot")
        .unwrap();

    let PivotSource::Consolidation { ranges } = &pivot.source else {
        panic!("expected consolidation source");
    };
    assert_eq!(ranges.len(), 2);
    assert_eq!(ranges[0].page_items, vec!["Retail"]);
    assert_eq!(ranges[1].page_items, vec!["Wholesale"]);
    assert_eq!(pivot.rows[0].field.name, "Row");
    assert_eq!(pivot.columns[0].field.name, "Column");
    assert_eq!(pivot.measures[0].field.name, "Value");
}

#[test]
fn reads_writer_xls_named_consolidation_source_semantics() {
    let mut wb = Workbook::new();
    add_named_consolidation_source_pivot(&mut wb);

    let bytes =
        XlsWriter::write_to_bytes(&wb).expect("serialize named consolidation-source workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    assert_eq!(
        workbook_records
            .iter()
            .filter(|(typ, _)| *typ == 0x0052)
            .count(),
        1,
        "named consolidation source should be emitted as DCONNAME"
    );
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let sxdb = records_with_payload(&cache)
        .into_iter()
        .find_map(|(record_type, payload)| (record_type == 0x00C6).then_some(payload))
        .expect("SXDB payload");
    assert_eq!(u16::from_le_bytes(sxdb[16..18].try_into().unwrap()), 0x0004);

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("NamedConsolidatedPivot")
        .unwrap();
    match &pivot.source {
        PivotSource::Consolidation { ranges } => {
            assert_eq!(ranges.len(), 1);
            assert_eq!(ranges[0].name.as_deref(), Some("NamedSalesSource"));
            assert_eq!(ranges[0].range, None);
        }
        other => panic!("expected consolidation source, got {other:?}"),
    }
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(PivotCacheSourceKind::Consolidation)
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_external_consolidation_source_dconref() {
    let mut wb = Workbook::new();
    add_external_consolidation_source_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb)
        .expect("serialize external consolidation-source pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let dconref = records_with_payload(&workbook)
        .into_iter()
        .find_map(|(record_type, payload)| (record_type == 0x0051).then_some(payload))
        .expect("external DCONREF payload");

    let expected_source = b"[external_sales.xlsx]ExternalData";
    assert_eq!(
        u16::from_le_bytes(dconref[6..8].try_into().unwrap()) as usize,
        expected_source.len() + 1,
        "DCONREF length includes the hidden source marker"
    );
    assert_eq!(dconref[8], 0x00, "ASCII DCONREF source name");
    assert_eq!(dconref[9], 0x01, "0x01 marks an external DCONREF source");
    assert_eq!(&dconref[10..10 + expected_source.len()], expected_source);

    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let sxdb = records_with_payload(&cache)
        .into_iter()
        .find_map(|(record_type, payload)| (record_type == 0x00C6).then_some(payload))
        .expect("SXDB payload");
    assert_eq!(u16::from_le_bytes(sxdb[16..18].try_into().unwrap()), 0x0004);
}

#[test]
fn reads_writer_xls_external_consolidation_source_semantics() {
    let mut wb = Workbook::new();
    add_external_consolidation_source_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb)
        .expect("serialize external consolidation-source pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ExternalConsolidatedPivot")
        .unwrap();

    match &pivot.source {
        PivotSource::Consolidation { ranges } => {
            assert_eq!(ranges.len(), 1);
            assert_eq!(ranges[0].sheet.as_deref(), Some("ExternalData"));
            assert_eq!(ranges[0].range, Some(CellRange::parse("A1:B3").unwrap()));
            assert_eq!(
                ranges[0].external_relationship_target.as_deref(),
                Some("external_sales.xlsx")
            );
        }
        other => panic!("expected consolidation source, got {other:?}"),
    }
    assert_eq!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(PivotCacheSourceKind::Consolidation)
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_unnamed_scenario_source_metadata() {
    let mut wb = Workbook::new();
    add_unnamed_scenario_source_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize scenario-source pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxvs = workbook_records
        .iter()
        .find_map(|(typ, payload)| (*typ == 0x00E3).then_some(payload))
        .expect("SXVS pivot cache record");
    assert_eq!(
        u16::from_le_bytes(sxvs[0..2].try_into().unwrap()),
        0x0010,
        "scenario pivot caches should use the scenario-source SXVS kind"
    );
    assert!(
        !workbook_records
            .iter()
            .any(|(typ, _)| *typ == 0x0051 || *typ == 0x0052),
        "unnamed scenario sources should not emit worksheet DCON source records"
    );

    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let sxdb = records_with_payload(&cache)
        .into_iter()
        .find_map(|(record_type, payload)| (record_type == 0x00C6).then_some(payload))
        .expect("SXDB payload");
    assert_eq!(
        u16::from_le_bytes(sxdb[16..18].try_into().unwrap()),
        0x0008,
        "SXDB.vsType should identify scenario sources"
    );
}

#[test]
fn reads_writer_xls_unnamed_scenario_source_semantics() {
    let mut wb = Workbook::new();
    add_unnamed_scenario_source_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize scenario-source pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ScenarioPivot")
        .unwrap();

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
fn xls_writer_rejects_named_scenario_pivot_source_authoring() {
    let mut wb = Workbook::new();
    let pivot = PivotTable::builder("ScenarioPivot")
        .source(PivotSource::Scenario {
            name: "BestCase".to_string(),
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let err = XlsWriter::write_to_bytes(&wb)
        .expect_err("named scenario pivot sources should fail explicitly");
    assert!(
        err.to_string().contains(
            "XLS named scenario pivot source authoring requires Scenario Manager records"
        ),
        "{err}"
    );
}

#[test]
fn xls_writer_rejects_olap_pivot_source_authoring() {
    let mut wb = Workbook::new();
    let pivot = PivotTable::builder("OlapPivot")
        .source(PivotSource::Olap {
            connection_name: "CubeSales".to_string(),
            cube: Some("SalesCube".to_string()),
            command_text: Some("SalesCube".to_string()),
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let err =
        XlsWriter::write_to_bytes(&wb).expect_err("XLS OLAP pivot sources should fail explicitly");
    assert!(
        err.to_string().contains("XLS OLAP pivot source authoring"),
        "{err}"
    );
}

#[test]
fn xls_writer_rejects_external_database_pivot_source_authoring() {
    let mut wb = Workbook::new();
    add_external_source_pivot(&mut wb);

    let err = XlsWriter::write_to_bytes(&wb)
        .expect_err("XLS external database pivot sources should fail explicitly");
    assert!(
        err.to_string()
            .contains("XLS external database pivot source authoring"),
        "{err}"
    );
}

#[test]
fn reads_xls_external_pivot_cache_source_kind() {
    let read = read_test_pivot_with_sxdb_source_type(0x0002);
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .unwrap();

    assert!(matches!(
        &pivot.source,
        PivotSource::External {
            connection_name,
            command_text: None,
        } if connection_name.is_empty()
    ));
    assert_eq!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(PivotCacheSourceKind::External)
    );
}

#[test]
fn reads_xls_consolidation_pivot_cache_source_kind() {
    let read = read_test_pivot_with_sxdb_source_type(0x0004);
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .unwrap();

    match &pivot.source {
        PivotSource::Consolidation { ranges } => {
            assert_eq!(ranges.len(), 1);
            assert_eq!(ranges[0].sheet.as_deref(), Some("Sheet1"));
            assert_eq!(ranges[0].range, Some(CellRange::parse("A1:B3").unwrap()));
        }
        other => panic!("expected consolidation source, got {other:?}"),
    }
    assert_eq!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(PivotCacheSourceKind::Consolidation)
    );
}

#[test]
fn reads_xls_scenario_pivot_cache_source_kind() {
    let read = read_test_pivot_with_sxdb_source_type(0x0008);
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .unwrap();

    assert!(matches!(
        &pivot.source,
        PivotSource::Scenario { name } if name.is_empty()
    ));
    assert_eq!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(PivotCacheSourceKind::Scenario)
    );
}

#[test]
fn xls_pivot_fields_keep_default_include_new_items_false_without_filters() {
    let mut wb = Workbook::new();
    add_test_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let row_sxvd_position = workbook_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD record");
    let row_sxvdex = workbook_records[row_sxvd_position + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX record");
    let grbit1 = u32::from_le_bytes(row_sxvdex[0..4].try_into().unwrap());
    assert_eq!(
        grbit1 & 0x8000,
        0,
        "the XLS writer must not emit fHideNewItems without Excel-compatible legacy context"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .expect("BasicPivot pivot");
    assert!(!pivot.rows[0].include_new_items_in_filter);
}

#[test]
fn xls_writer_accepts_filtered_field_include_new_items_default_false() {
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

    let pivot = PivotTable::builder("FilteredDefault")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let row_sxvd_position = workbook_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD record");
    let row_sxvdex = workbook_records[row_sxvd_position + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX record");
    assert_eq!(
        u32::from_le_bytes(row_sxvdex[0..4].try_into().unwrap()) & 0x8000,
        0,
        "Excel's BIFF8 output keeps fHideNewItems clear for filtered default-false fields"
    );
    assert_eq!(
        sxvi_payloads_for_sxvd(&workbook_records, 0)
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Central should be hidden while East, West, and the automatic subtotal item stay visible"
    );
    assert_eq!(
        sxaddl_field_include_new_items_flags(&workbook_records, "Region"),
        Some(0x28),
        "SXADDL field flags should preserve include-new-items=false separately from SXVDEX"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("FilteredDefault")
        .expect("FilteredDefault pivot");
    assert!(!pivot.rows[0].include_new_items_in_filter);
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_average_data_field() {
    let mut wb = Workbook::new();
    add_average_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxdi = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C5).then_some(payload))
        .expect("SXDI data field record");

    assert_eq!(
        u16::from_le_bytes(sxdi[2..4].try_into().unwrap()),
        2,
        "BIFF8 data function 2 is Average"
    );
    assert!(
        String::from_utf8_lossy(sxdi).contains("Average Revenue"),
        "SXDI caption should come from the semantic pivot measure"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_column_axis_records() {
    let mut wb = Workbook::new();
    add_column_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let sxview = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW record");
    assert_eq!(u16::from_le_bytes(sxview[22..24].try_into().unwrap()), 3);
    assert_eq!(u16::from_le_bytes(sxview[24..26].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[26..28].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[34..36].try_into().unwrap()), 3);

    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxvd_axes, vec![0x0001, 0x0002, 0x0008]);

    let axis_declarations = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B4).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        axis_declarations,
        vec![vec![0, 0], vec![1, 0]],
        "row and column axes should declare Region then Quarter"
    );

    let sxli_lengths = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload.len()))
        .collect::<Vec<_>>();
    assert_eq!(sxli_lengths, vec![30, 30]);
}

#[test]
fn semantic_pivot_tables_emit_xls_page_axis_records() {
    let mut wb = Workbook::new();
    add_page_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let sxview = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW record");
    assert_eq!(u16::from_le_bytes(sxview[22..24].try_into().unwrap()), 3);
    assert_eq!(u16::from_le_bytes(sxview[24..26].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[26..28].try_into().unwrap()), 0);
    assert_eq!(u16::from_le_bytes(sxview[28..30].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[30..32].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[32..34].try_into().unwrap()), 2);

    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxvd_axes, vec![0x0001, 0x0004, 0x0008]);

    let page_fields = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B6).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(page_fields, vec![vec![1, 0, 0, 0, 1, 0]]);

    let line_items = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        line_items.first(),
        Some(&vec![
            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 13, 0, 0, 0, 0, 10, 0, 0,
        ])
    );
}

#[test]
fn semantic_pivot_tables_round_trip_xls_data_caption() {
    let mut wb = Workbook::new();
    add_test_pivot(&mut wb);
    wb.worksheet_mut(0).unwrap().pivot_tables_mut()[0]
        .layout
        .data_caption = "Measures".to_string();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxview = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW record");
    assert_eq!(
        u16::from_le_bytes(sxview[42..44].try_into().unwrap()),
        8,
        "SXVIEW data-caption length should be written"
    );
    assert!(
        String::from_utf8_lossy(sxview).contains("Measures"),
        "SXVIEW should carry the custom data caption"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .expect("pivot after read");
    assert_eq!(pivot.layout.data_caption, "Measures");
}

#[test]
fn semantic_pivot_tables_round_trip_xls_table_layout_flags() {
    let mut wb = Workbook::new();
    add_page_pivot(&mut wb);
    {
        let pivot = &mut wb.worksheet_mut(0).unwrap().pivot_tables_mut()[0];
        pivot.layout.show_row_grand_totals = false;
        pivot.layout.show_column_grand_totals = false;
        pivot.layout.page_wrap = 2;
        pivot.layout.page_over_then_down = true;
        pivot.layout.merge_item_labels = true;
        pivot.layout.error_caption = Some("ERR".to_string());
        pivot.layout.show_error = true;
        pivot.layout.missing_caption = Some("MISS".to_string());
        pivot.layout.show_missing = true;
        pivot.layout.edit_data = true;
        pivot.layout.disable_field_list = true;
        pivot.layout.enable_wizard = false;
        pivot.layout.enable_drill = false;
        pivot.layout.enable_field_properties = false;
        pivot.layout.field_print_titles = true;
        pivot.layout.item_print_titles = true;
        pivot.layout.grand_total_caption = Some("Grand".to_string());
        pivot.refresh_policy.preserve_formatting = false;
    }

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxview = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW record");
    let sxview_flags = u16::from_le_bytes(sxview[36..38].try_into().unwrap());
    assert_eq!(
        sxview_flags & 0x0003,
        0,
        "SXVIEW should clear row and column grand totals"
    );

    let sxex = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F1).then_some(payload))
        .expect("SXEX record");
    assert_eq!(u16::from_le_bytes(sxex[2..4].try_into().unwrap()), 3);
    assert_eq!(u16::from_le_bytes(sxex[4..6].try_into().unwrap()), 4);
    assert_eq!(u16::from_le_bytes(sxex[10..12].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxex[12..14].try_into().unwrap()), 1);
    assert_eq!(
        u16::from_le_bytes(sxex[14..16].try_into().unwrap()),
        0x0205,
        "SXEX grbit1 should encode across-page layout and page wrap"
    );
    assert_eq!(
        u16::from_le_bytes(sxex[16..18].try_into().unwrap()),
        0x0670,
        "SXEX grbit2 should encode table-level layout flags"
    );
    assert!(String::from_utf8_lossy(sxex).contains("ERR"));
    assert!(String::from_utf8_lossy(sxex).contains("MISS"));

    let sxviewex9 = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0810).then_some(payload))
        .expect("SXVIEWEX9 record");
    let sxviewex9_flags = u32::from_le_bytes(sxviewex9[8..12].try_into().unwrap());
    assert_eq!(sxviewex9_flags & 0x0002, 0x0002, "print titles");
    assert_eq!(
        sxviewex9_flags & 0x0020,
        0x0020,
        "repeat printed item labels"
    );
    assert!(String::from_utf8_lossy(sxviewex9).contains("Grand"));

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RevenueByRep")
        .expect("pivot after read");
    assert!(!pivot.layout.show_row_grand_totals);
    assert!(!pivot.layout.show_column_grand_totals);
    assert_eq!(pivot.layout.page_wrap, 2);
    assert!(pivot.layout.page_over_then_down);
    assert!(pivot.layout.merge_item_labels);
    assert_eq!(pivot.layout.error_caption.as_deref(), Some("ERR"));
    assert!(pivot.layout.show_error);
    assert_eq!(pivot.layout.missing_caption.as_deref(), Some("MISS"));
    assert!(pivot.layout.show_missing);
    assert!(pivot.layout.edit_data);
    assert!(pivot.layout.disable_field_list);
    assert!(!pivot.layout.enable_wizard);
    assert!(!pivot.layout.enable_drill);
    assert!(!pivot.layout.enable_field_properties);
    assert!(pivot.layout.field_print_titles);
    assert!(pivot.layout.item_print_titles);
    assert_eq!(pivot.layout.grand_total_caption.as_deref(), Some("Grand"));
    assert!(!pivot.refresh_policy.preserve_formatting);
}

#[test]
fn xls_writer_rejects_subtotal_hidden_items_layout_flag() {
    let mut wb = Workbook::new();
    add_test_pivot(&mut wb);
    wb.worksheet_mut(0).unwrap().pivot_tables_mut()[0]
        .layout
        .subtotal_hidden_items = true;

    let err = XlsWriter::write_to_bytes(&wb)
        .expect_err("XLS subtotal-hidden-items layout flag should be rejected");
    assert!(err.to_string().contains("hidden-item subtotals"), "{err}");
}

#[test]
fn xls_writer_round_trips_multi_item_page_filters() {
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

    let pivot = PivotTable::builder("RevenueByRep")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Salesperson")
        .page(include_new_items_field("Region"))
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize multi-item page filter");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let page_field = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B6).then_some(payload))
        .expect("SXPI page field");
    assert_eq!(
        u16::from_le_bytes(page_field[2..4].try_into().unwrap()),
        0x7FFD,
        "multi-select page filters should use Excel's multiple-items sentinel"
    );

    let region_items = sxvi_payloads_for_sxvd(&workbook_records, 0);
    assert_eq!(
        region_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Central should be hidden while East, West, and Auto remain visible"
    );
    assert_eq!(
        i16::from_le_bytes(region_items[2][4..6].try_into().unwrap()),
        2
    );
    let row_sxli = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload))
        .expect("row SXLI record");
    assert_eq!(
        sxli_line_item_indexes(row_sxli),
        vec![vec![0], vec![1]],
        "row SXLI tuples should omit hidden Central"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RevenueByRep")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );
    assert!(pivot.page_fields[0].include_new_items_in_filter);
}

#[test]
fn xls_writer_round_trips_top_n_pivot_filter() {
    let mut wb = Workbook::new();
    add_top_n_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxvdex = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX record");
    let flags = u32::from_le_bytes(sxvdex[0..4].try_into().unwrap());
    assert_ne!(flags & 0x0800, 0, "fAutoShow should be set");
    assert_ne!(flags & 0x1000, 0, "fTopAutoShow should be set");
    assert_eq!(
        flags >> 24,
        10,
        "SXVDEx should keep Excel's default AutoShow count"
    );
    assert_eq!(
        i16::from_le_bytes(sxvdex[6..8].try_into().unwrap()),
        0,
        "isxdiAutoShow should target the first data field"
    );
    let sxaddl_autoshow = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 10
                && payload[4] == 0x17
                && payload[5] == 0x37)
                .then_some(payload)
        })
        .expect("SXADDL sxdAutoshow record");
    assert_eq!(
        u32::from_le_bytes(sxaddl_autoshow[6..10].try_into().unwrap()),
        2,
        "SXADDL sxdAutoshow should store the top count"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TopRegions")
        .expect("pivot after read");
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
        other => panic!("unexpected filter: {other:?}"),
    }
}

#[test]
fn xls_writer_round_trips_percent_top_n_pivot_filter() {
    let mut wb = Workbook::new();
    add_percent_top_n_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxvdex = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX record");
    let flags = u32::from_le_bytes(sxvdex[0..4].try_into().unwrap());
    assert_eq!(
        flags & 0x0800,
        0,
        "percent top-N should use the SXADDL filter extension rather than legacy AutoShow"
    );

    let filter_extension = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 40
                && payload[4] == 0x1D
                && payload[5] == 0x3C)
                .then_some(payload)
        })
        .expect("SXADDL pivot filter extension record");
    assert_eq!(
        u32::from_le_bytes(filter_extension[12..16].try_into().unwrap()),
        3,
        "filter type should be top-percent"
    );
    assert_eq!(
        u32::from_le_bytes(filter_extension[16..20].try_into().unwrap()),
        1,
        "filtered Region source field should be one-based"
    );
    assert_eq!(
        u32::from_le_bytes(filter_extension[20..24].try_into().unwrap()),
        2,
        "Revenue measure source field should be one-based"
    );
    assert_eq!(&filter_extension[32..40], &[0x49, 0x40, 0, 0, 0, 0, 0, 0]);

    let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TopPercentRegions")
        .expect("pivot after read");
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
        other => panic!("unexpected filter: {other:?}"),
    }
}

#[test]
fn xls_writer_round_trips_label_contains_pivot_filter() {
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let filter_collection = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 24
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL pivot filter collection record");
    assert_eq!(
        u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
        1,
        "Segment should be the second, zero-based source field"
    );
    assert_eq!(
        u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
        10,
        "label contains should use Excel's caption-contains filter type"
    );

    let raw_criterion = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 12
                && payload[4] == 0x1D
                && payload[5] == 0x3A)
                .then_some(payload)
        })
        .expect("SXADDL raw label criterion record");
    assert_eq!(
        u16::from_le_bytes(raw_criterion[6..8].try_into().unwrap()),
        2
    );
    assert_eq!(xls_unicode_string_at(raw_criterion, 12), "Ed");

    let custom_filter = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 24
                && payload[4] == 0x1D
                && payload[5] == 0x3C)
                .then_some(payload)
        })
        .expect("SXADDL custom label filter record");
    assert_eq!(
        u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
        1,
        "first extension filter should have id 1"
    );
    assert_eq!(
        &custom_filter[20..24],
        &[0x06, 0x02, 0x00, 0x00],
        "contains is represented as equality over wildcard criteria"
    );

    let wildcard_criterion = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 12
                && payload[4] == 0x1D
                && payload[5] == 0x3D)
                .then_some(payload)
        })
        .expect("SXADDL wildcard label criterion record");
    assert_eq!(
        u16::from_le_bytes(wildcard_criterion[6..8].try_into().unwrap()),
        4
    );
    assert_eq!(xls_unicode_string_at(wildcard_criterion, 12), "*Ed*");

    let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("LabelSegments")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::Label {
            field: PivotFieldRef::new("Segment"),
            operator: PivotFilterOperator::Contains,
            value: "Ed".to_string(),
        }]
    );
}

#[test]
fn xls_writer_round_trips_label_equals_pivot_filter() {
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

    let pivot = PivotTable::builder("LabelEqualsRows")
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let filter_collection = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 24
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL pivot filter collection record");
    assert_eq!(
        u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
        0,
        "Region should be the first, zero-based source field"
    );
    assert_eq!(
        u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
        4,
        "label equals should use Excel's caption-equals filter type"
    );

    let raw_criterion = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 12
                && payload[4] == 0x1D
                && payload[5] == 0x3A)
                .then_some(payload)
        })
        .expect("SXADDL raw label criterion record");
    assert_eq!(
        u16::from_le_bytes(raw_criterion[6..8].try_into().unwrap()),
        4
    );
    assert_eq!(xls_unicode_string_at(raw_criterion, 12), "East");

    let custom_filter = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 24
                && payload[4] == 0x1D
                && payload[5] == 0x3C)
                .then_some(payload)
        })
        .expect("SXADDL custom label filter record");
    assert_eq!(
        u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
        1,
        "first extension filter should have id 1"
    );
    assert_eq!(
        &custom_filter[20..24],
        &[0x06, 0x02, 0x01, 0x00],
        "equals is represented as a raw equality criterion"
    );

    let stored_criterion = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 12
                && payload[4] == 0x1D
                && payload[5] == 0x3D)
                .then_some(payload)
        })
        .expect("SXADDL stored label criterion record");
    assert_eq!(
        u16::from_le_bytes(stored_criterion[6..8].try_into().unwrap()),
        4
    );
    assert_eq!(xls_unicode_string_at(stored_criterion, 12), "East");

    let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("LabelEqualsRows")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::Label {
            field: PivotFieldRef::new("Region"),
            operator: PivotFilterOperator::Equals,
            value: "East".to_string(),
        }]
    );
}

#[test]
fn xls_writer_round_trips_label_prefix_suffix_pivot_filters() {
    for (operator, value, filter_type, custom_value, pivot_name) in [
        (
            PivotFilterOperator::BeginsWith,
            "Ea",
            6u32,
            "Ea*",
            "LabelBeginsRows",
        ),
        (
            PivotFilterOperator::EndsWith,
            "st",
            8u32,
            "*st",
            "LabelEndsRows",
        ),
    ] {
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

        let pivot = PivotTable::builder(pivot_name)
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
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
        let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
        let workbook_records = records_with_payload(&workbook);

        let filter_collection = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot filter collection record");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Region should be the first, zero-based source field"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );

        let raw_criterion = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3A)
                    .then_some(payload)
            })
            .expect("SXADDL raw label criterion record");
        assert_eq!(xls_unicode_string_at(raw_criterion, 12), value);

        let custom_filter = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom label filter record");
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            1,
            "first extension filter should have id 1"
        );
        assert_eq!(
            &custom_filter[20..24],
            &[0x06, 0x02, 0x00, 0x00],
            "begins/ends filters are represented as equality over wildcard criteria"
        );

        let stored_criterion = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3D)
                    .then_some(payload)
            })
            .expect("SXADDL stored label criterion record");
        assert_eq!(xls_unicode_string_at(stored_criterion, 12), custom_value);

        let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator,
                value: value.to_string(),
            }]
        );
    }
}

#[test]
fn xls_writer_round_trips_negative_label_pivot_filters() {
    for (operator, value, filter_type, discriminator, custom_value, pivot_name) in [
        (
            PivotFilterOperator::NotEquals,
            "East",
            5u32,
            1u8,
            "East",
            "LabelNotEqualsRows",
        ),
        (
            PivotFilterOperator::DoesNotBeginWith,
            "Ea",
            7u32,
            0u8,
            "Ea*",
            "LabelNotBeginsRows",
        ),
        (
            PivotFilterOperator::DoesNotEndWith,
            "st",
            9u32,
            0u8,
            "*st",
            "LabelNotEndsRows",
        ),
        (
            PivotFilterOperator::DoesNotContain,
            "Ea",
            11u32,
            0u8,
            "*Ea*",
            "LabelNotContainsRows",
        ),
    ] {
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

        let pivot = PivotTable::builder(pivot_name)
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
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
        let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
        let workbook_records = records_with_payload(&workbook);

        let filter_collection = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot filter collection record");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Region should be the first, zero-based source field"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );

        let raw_criterion = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3A)
                    .then_some(payload)
            })
            .expect("SXADDL raw label criterion record");
        assert_eq!(xls_unicode_string_at(raw_criterion, 12), value);

        let custom_filter = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom label filter record");
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            1,
            "first extension filter should have id 1"
        );
        assert_eq!(
            &custom_filter[20..24],
            &[0x06, 0x05, discriminator, 0x00],
            "negative label filters are represented as notEqual criteria"
        );

        let stored_criterion = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3D)
                    .then_some(payload)
            })
            .expect("SXADDL stored label criterion record");
        assert_eq!(xls_unicode_string_at(stored_criterion, 12), custom_value);

        let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator,
                value: value.to_string(),
            }]
        );
    }
}

#[test]
fn xls_writer_round_trips_label_comparison_pivot_filters() {
    for (operator, filter_type, custom_operator, pivot_name) in [
        (
            PivotFilterOperator::GreaterThan,
            12u32,
            0x04u8,
            "LabelGreaterRows",
        ),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            13u32,
            0x06u8,
            "LabelGreaterEqualRows",
        ),
        (
            PivotFilterOperator::LessThan,
            14u32,
            0x01u8,
            "LabelLessRows",
        ),
        (
            PivotFilterOperator::LessThanOrEqual,
            15u32,
            0x03u8,
            "LabelLessEqualRows",
        ),
    ] {
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

        let pivot = PivotTable::builder(pivot_name)
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .filter(PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator,
                value: "M".to_string(),
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
        let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
        let workbook_records = records_with_payload(&workbook);

        let filter_collection = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot filter collection record");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Region should be the first, zero-based source field"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );

        let raw_criterion = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3A)
                    .then_some(payload)
            })
            .expect("SXADDL raw label criterion record");
        assert_eq!(xls_unicode_string_at(raw_criterion, 12), "M");

        let custom_filter = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom label filter record");
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            1,
            "first extension filter should have id 1"
        );
        assert_eq!(
            &custom_filter[20..24],
            &[0x06, custom_operator, 0x01, 0x00],
            "comparison label filters are represented as raw criteria"
        );

        let stored_criterion = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3D)
                    .then_some(payload)
            })
            .expect("SXADDL stored label criterion record");
        assert_eq!(xls_unicode_string_at(stored_criterion, 12), "M");

        let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator,
                value: "M".to_string(),
            }]
        );
    }
}

#[test]
fn xls_writer_round_trips_label_between_pivot_filters() {
    for (not_between, filter_type, and_flag, first_operator, second_operator, pivot_name) in [
        (false, 16u32, 1u32, 0x06u8, 0x03u8, "LabelBetweenRows"),
        (true, 17u32, 2u32, 0x01u8, 0x04u8, "LabelNotBetweenRows"),
    ] {
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

        let pivot = PivotTable::builder(pivot_name)
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
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
        let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
        let workbook_records = records_with_payload(&workbook);

        let filter_collection = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 24
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot filter collection record");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Region should be the first, zero-based source field"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );

        let raw_start = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3A)
                    .then_some(payload)
            })
            .expect("SXADDL raw lower label criterion record");
        assert_eq!(xls_unicode_string_at(raw_start, 12), "East");

        let raw_end = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3B)
                    .then_some(payload)
            })
            .expect("SXADDL raw upper label criterion record");
        assert_eq!(xls_unicode_string_at(raw_end, 12), "West");

        let custom_filter = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 44
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom label filter record");
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            2,
            "label range filters should store two custom criteria"
        );
        assert_eq!(&custom_filter[20..24], &[0x06, first_operator, 0x01, 0x00]);
        assert_eq!(&custom_filter[30..34], &[0x06, second_operator, 0x01, 0x00]);
        assert_eq!(
            u32::from_le_bytes(custom_filter[40..44].try_into().unwrap()),
            and_flag
        );

        let stored_start = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3D)
                    .then_some(payload)
            })
            .expect("SXADDL stored lower label criterion record");
        assert_eq!(xls_unicode_string_at(stored_start, 12), "East");

        let stored_end = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 12
                    && payload[4] == 0x1D
                    && payload[5] == 0x3E)
                    .then_some(payload)
            })
            .expect("SXADDL stored upper label criterion record");
        assert_eq!(xls_unicode_string_at(stored_end, 12), "West");

        let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::LabelBetween {
                field: PivotFieldRef::new("Region"),
                start: "East".to_string(),
                end: "West".to_string(),
                not_between,
            }]
        );
    }
}

#[test]
fn xls_writer_round_trips_value_comparison_pivot_filters() {
    for (operator, filter_type, custom_operator, threshold, pivot_name) in [
        (
            PivotFilterOperator::Equals,
            18u32,
            0x02u8,
            20.0,
            "ValueEqualsRows",
        ),
        (
            PivotFilterOperator::NotEquals,
            19u32,
            0x05u8,
            20.0,
            "ValueNotEqualsRows",
        ),
        (
            PivotFilterOperator::GreaterThan,
            20u32,
            0x04u8,
            20.0,
            "ValueGreaterRows",
        ),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            21u32,
            0x06u8,
            20.0,
            "ValueGreaterEqualRows",
        ),
        (
            PivotFilterOperator::LessThan,
            22u32,
            0x01u8,
            30.0,
            "ValueLessRows",
        ),
        (
            PivotFilterOperator::LessThanOrEqual,
            23u32,
            0x03u8,
            30.0,
            "ValueLessEqualRows",
        ),
    ] {
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

        let filter_measure =
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
        let pivot = PivotTable::builder(pivot_name)
            .source_range(CellRange::parse("A1:B5").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .filter(PivotFilter::Value {
                field: PivotFieldRef::new("Region"),
                measure: filter_measure.clone(),
                operator,
                value: threshold,
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
        let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
        let workbook_records = records_with_payload(&workbook);

        let filter_collection = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 36
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot filter collection record");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Region should be the first, zero-based source field"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
            0,
            "value filter should target the first data field"
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[32..36].try_into().unwrap()),
            -1
        );

        let custom_filter = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 30
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom value filter record");
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            1,
            "first extension filter should have id 1"
        );
        assert_eq!(custom_filter[20], 0x04, "numeric filter value type");
        assert_eq!(custom_filter[21], custom_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[22..30].try_into().unwrap()),
            threshold
        );

        let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Value {
                field: PivotFieldRef::new("Region"),
                measure: filter_measure,
                operator,
                value: threshold,
            }]
        );
    }
}

#[test]
fn xls_writer_round_trips_value_between_pivot_filters() {
    for (not_between, filter_type, first_operator, second_operator, flag, pivot_name) in [
        (false, 24u32, 0x06u8, 0x03u8, 1u32, "ValueBetweenRows"),
        (true, 25u32, 0x01u8, 0x04u8, 2u32, "ValueNotBetweenRows"),
    ] {
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

        let filter_measure =
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
        let pivot = PivotTable::builder(pivot_name)
            .source_range(CellRange::parse("A1:B5").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .filter(PivotFilter::ValueBetween {
                field: PivotFieldRef::new("Region"),
                measure: filter_measure.clone(),
                start: 15.0,
                end: 35.0,
                not_between,
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
        let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
        let workbook_records = records_with_payload(&workbook);

        let filter_collection = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 36
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot filter collection record");
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
            0,
            "value range filter should target the first data field"
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[32..36].try_into().unwrap()),
            -1
        );

        let custom_filter = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 48
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom value range filter record");
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            2
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

        let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after read");
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
    }
}

#[test]
fn xls_writer_round_trips_date_comparison_pivot_filters() {
    for (operator, filter_type, custom_operator, custom_type, pivot_name) in [
        (
            PivotFilterOperator::Equals,
            26u32,
            0x02u8,
            4u32,
            "DateEqualsRows",
        ),
        (
            PivotFilterOperator::NotEquals,
            62u32,
            0x05u8,
            40u32,
            "DateNotEqualsRows",
        ),
        (
            PivotFilterOperator::LessThan,
            27u32,
            0x01u8,
            5u32,
            "DateLessRows",
        ),
        (
            PivotFilterOperator::LessThanOrEqual,
            63u32,
            0x03u8,
            41u32,
            "DateLessEqualRows",
        ),
        (
            PivotFilterOperator::GreaterThan,
            28u32,
            0x04u8,
            6u32,
            "DateGreaterRows",
        ),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            64u32,
            0x06u8,
            42u32,
            "DateGreaterEqualRows",
        ),
    ] {
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

        let pivot = PivotTable::builder(pivot_name)
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Order Date")
            .measure("Revenue", PivotAggregate::Sum)
            .filter(PivotFilter::Date {
                field: PivotFieldRef::new("Order Date"),
                operator,
                value: 44958.0,
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
        let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
        let workbook_records = records_with_payload(&workbook);

        let filter_collection = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 36
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot date filter collection record");
        assert_eq!(
            u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
            0,
            "Order Date should be the first, zero-based source field"
        );
        assert_eq!(
            u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
            0,
            "date filters use Excel's label-style collection target"
        );
        assert_eq!(
            i32::from_le_bytes(filter_collection[32..36].try_into().unwrap()),
            0
        );

        let custom_filter = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 48
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom date filter record");
        assert_eq!(
            u32::from_le_bytes(custom_filter[12..16].try_into().unwrap()),
            custom_type
        );
        assert_eq!(
            u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
            1,
            "comparison date filters should have one criterion"
        );
        assert_eq!(custom_filter[20], 0x04, "date criterion stores a number");
        assert_eq!(custom_filter[21], custom_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[22..30].try_into().unwrap()),
            44958.0
        );

        let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Date {
                field: PivotFieldRef::new("Order Date"),
                operator,
                value: 44958.0,
            }]
        );
    }
}

#[test]
fn xls_writer_round_trips_date_between_pivot_filters() {
    for (
        not_between,
        filter_type,
        first_operator,
        second_operator,
        custom_type,
        flag,
        pivot_name,
    ) in [
        (false, 29u32, 0x06u8, 0x03u8, 7u32, 1u32, "DateBetweenRows"),
        (
            true,
            65u32,
            0x01u8,
            0x04u8,
            43u32,
            2u32,
            "DateNotBetweenRows",
        ),
    ] {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Order Date").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        ws.set_cell_value("A2", 44927.0).unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();
        ws.set_cell_value("A3", 44958.0).unwrap();
        ws.set_cell_value("B3", 20.0).unwrap();
        ws.set_cell_value("A4", 45016.0).unwrap();
        ws.set_cell_value("B4", 30.0).unwrap();

        let pivot = PivotTable::builder(pivot_name)
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Order Date")
            .measure("Revenue", PivotAggregate::Sum)
            .filter(PivotFilter::DateBetween {
                field: PivotFieldRef::new("Order Date"),
                start: 44927.0,
                end: 45016.0,
                not_between,
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
        let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
        let workbook_records = records_with_payload(&workbook);

        let filter_collection = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 36
                    && payload[4] == 0x1D
                    && payload[5] == 0x38)
                    .then_some(payload)
            })
            .expect("SXADDL pivot date range filter collection record");
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

        let custom_filter = workbook_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == 0x0864
                    && payload.len() >= 48
                    && payload[4] == 0x1D
                    && payload[5] == 0x3C)
                    .then_some(payload)
            })
            .expect("SXADDL custom date range filter record");
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

        let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name(pivot_name)
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::DateBetween {
                field: PivotFieldRef::new("Order Date"),
                start: 44927.0,
                end: 45016.0,
                not_between,
            }]
        );
    }
}

#[test]
fn xls_writer_round_trips_date_period_pivot_filter() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Order Date").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 44927.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 44958.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("DatePeriodRows")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Order Date")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::DatePeriod {
            field: PivotFieldRef::new("Order Date"),
            period: PivotDatePeriod::ThisMonth,
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let filter_collection = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 36
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL pivot date-period filter collection record");
    assert_eq!(
        u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
        37,
        "ThisMonth should use Excel's dynamic-date pivot filter type"
    );

    let custom_filter = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 20
                && payload[4] == 0x1D
                && payload[5] == 0x3C)
                .then_some(payload)
        })
        .expect("SXADDL date-period filter record");
    assert_eq!(
        u32::from_le_bytes(custom_filter[12..16].try_into().unwrap()),
        0x0F,
        "ThisMonth should use the worksheet dynamic-filter discriminator"
    );
    assert_eq!(
        u32::from_le_bytes(custom_filter[16..20].try_into().unwrap()),
        2,
        "relative date periods store current-period bounds for Excel"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("DatePeriodRows")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::DatePeriod {
            field: PivotFieldRef::new("Order Date"),
            period: PivotDatePeriod::ThisMonth,
        }]
    );
}

#[test]
fn xls_writer_round_trips_column_value_pivot_filter() {
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

    let filter_measure =
        PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
    let pivot = PivotTable::builder("ColumnValueFilter")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Segment")
        .column("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::Value {
            field: PivotFieldRef::new("Region"),
            measure: filter_measure.clone(),
            operator: PivotFilterOperator::GreaterThan,
            value: 25.0,
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let filter_collection = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 36
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL pivot filter collection record");
    assert_eq!(
        u32::from_le_bytes(filter_collection[12..16].try_into().unwrap()),
        0,
        "Region should be the first source field even when it is on the column axis"
    );
    assert_eq!(
        u32::from_le_bytes(filter_collection[20..24].try_into().unwrap()),
        20,
        "greater-than value filters should use Excel's value-greater filter type"
    );
    assert_eq!(
        i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
        0,
        "single-measure column value filter should target the first data field"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ColumnValueFilter")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::Value {
            field: PivotFieldRef::new("Region"),
            measure: filter_measure,
            operator: PivotFilterOperator::GreaterThan,
            value: 25.0,
        }]
    );
}

#[test]
fn xls_writer_round_trips_second_measure_value_pivot_filter() {
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

    let filter_measure = PivotMeasure::new("Cost", PivotAggregate::Sum).with_name("Total Cost");
    let mut pivot = PivotTable::builder("SecondMeasureValueFilter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .named_measure("Cost", PivotAggregate::Sum, "Total Cost")
        .filter(PivotFilter::Value {
            field: PivotFieldRef::new("Region"),
            measure: filter_measure.clone(),
            operator: PivotFilterOperator::GreaterThan,
            value: 10.0,
        })
        .build()
        .unwrap();
    pivot.layout.values_axis = PivotValuesAxis::Columns;
    pivot.layout.values_axis_position = Some(0);
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let filter_collection = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 36
                && payload[4] == 0x1D
                && payload[5] == 0x38)
                .then_some(payload)
        })
        .expect("SXADDL pivot filter collection record");
    assert_eq!(
        i32::from_le_bytes(filter_collection[28..32].try_into().unwrap()),
        1,
        "value filter should target the second data field"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SecondMeasureValueFilter")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::Value {
            field: PivotFieldRef::new("Region"),
            measure: filter_measure,
            operator: PivotFilterOperator::GreaterThan,
            value: 10.0,
        }]
    );
}

#[test]
fn xls_writer_rejects_unknown_page_filter_item() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Salesperson").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Ada").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();

    let pivot = PivotTable::builder("RevenueByRep")
        .source_range(CellRange::parse("A1:C2").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .page(include_new_items_field("Salesperson"))
        .filter(PivotFilter::field_items("Salesperson", ["Ben"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let err = XlsWriter::write_to_bytes(&wb).expect_err("unknown page item should fail for XLS");
    assert!(
        err.to_string().contains("selected item is not present"),
        "{err}"
    );
}

#[test]
fn xls_writer_round_trips_row_field_item_filters() {
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

    let pivot = PivotTable::builder("FilteredRows")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row(include_new_items_field("Region"))
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize row field item filter");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let region_items = sxvi_payloads_for_sxvd(&workbook_records, 0);
    assert_eq!(
        region_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Central should be hidden while East, West, and Auto remain visible"
    );
    assert_eq!(
        i16::from_le_bytes(region_items[2][4..6].try_into().unwrap()),
        2
    );
    assert_eq!(
        sxaddl_field_include_new_items_flags(&workbook_records, "Region"),
        Some(0x08),
        "SXADDL field flags should preserve include-new-items=true separately from SXVDEX"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("FilteredRows")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );
    assert!(pivot.rows[0].include_new_items_in_filter);
}

#[test]
fn xls_writer_round_trips_column_field_item_filters() {
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

    let pivot = PivotTable::builder("FilteredColumns")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Salesperson")
        .column(include_new_items_field("Region"))
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize column field item filter");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let column_sxvd_ordinal = workbook_records
        .iter()
        .filter(|(record_type, _)| *record_type == 0x00B1)
        .enumerate()
        .find_map(|(ordinal, (_, payload))| {
            let axis = u16::from_le_bytes(payload[0..2].try_into().unwrap());
            let item_count = u16::from_le_bytes(payload[6..8].try_into().unwrap());
            (axis == 0x0002 && item_count > 0).then_some(ordinal)
        })
        .expect("column SXVD ordinal");
    let region_items = sxvi_payloads_for_sxvd(&workbook_records, column_sxvd_ordinal);
    assert!(
        workbook_records.iter().any(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0002
        }),
        "column SXVD should be present"
    );
    assert_eq!(
        region_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Central should be hidden while East, West, and Auto remain visible"
    );
    let column_sxli = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload))
        .nth(1)
        .expect("column SXLI record");
    assert_eq!(
        sxli_line_item_indexes(column_sxli),
        vec![vec![0], vec![1]],
        "column SXLI tuples should omit hidden Central"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("FilteredColumns")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );
    assert!(pivot.columns[0].include_new_items_in_filter);
}

#[test]
fn xls_writer_round_trips_source_only_field_item_filters() {
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

    let pivot = PivotTable::builder("SourceOnlyFilter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .filter(PivotFilter::field_items("Channel", ["Online"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize source-only item filter");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert!(sxvd_axes
        .windows(3)
        .any(|axes| axes == [0x0001, 0x0000, 0x0008]));

    let channel_items = sxvi_payloads_for_sxvd(&workbook_records, 1);
    assert_eq!(
        channel_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 1, 0],
        "Store should be hidden while Online and Auto remain visible"
    );
    let row_sxli = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload))
        .next()
        .expect("row SXLI record");
    assert_eq!(
        sxli_line_item_indexes(row_sxli),
        vec![vec![0], vec![2]],
        "row SXLI tuples should omit rows hidden by the source-only Channel filter"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SourceOnlyFilter")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Channel", ["Online"])]
    );
}

#[test]
fn xls_writer_rejects_empty_label_filter_values() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("LabelFilteredRows")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .filter(PivotFilter::Label {
            field: PivotFieldRef::new("Region"),
            operator: PivotFilterOperator::Equals,
            value: String::new(),
        })
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let err =
        XlsWriter::write_to_bytes(&wb).expect_err("empty label filter values should fail for XLS");
    assert!(
        err.to_string().contains("empty label filter value"),
        "{err}"
    );
}

#[test]
fn xls_writer_round_trips_grouped_row_field_item_filters() {
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

    let pivot = PivotTable::builder("GroupedFilter")
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize grouped row item filter");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let derived_group_items = sxvi_payloads_for_sxvd(&workbook_records, 2);
    assert_eq!(
        derived_group_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![1, 0, 0],
        "Central should be hidden while Coastal and Auto remain visible"
    );
    assert_eq!(
        sxli_line_item_indexes(
            workbook_records
                .iter()
                .find_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload))
                .expect("row SXLI record")
        ),
        vec![vec![1, 0], vec![1, 1]],
        "row SXLI tuples should only include East and West under Coastal"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedFilter")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal"])]
    );
}

#[test]
fn xls_writer_round_trips_grouped_column_field_item_filters() {
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

    let pivot = PivotTable::builder("GroupedColumnFilter")
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
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize grouped column item filter");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let derived_group_items = sxvi_payloads_for_sxvd(&workbook_records, 3);
    assert_eq!(
        derived_group_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![1, 0, 0],
        "Central should be hidden while Coastal and Auto remain visible"
    );
    assert_eq!(
        sxli_line_item_indexes(
            workbook_records
                .iter()
                .filter_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload))
                .nth(1)
                .expect("column SXLI record")
        ),
        vec![vec![1, 0], vec![1, 1]],
        "column SXLI tuples should only include East and West under Coastal"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedColumnFilter")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal"])]
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_axis_field_options() {
    let mut wb = Workbook::new();
    add_axis_options_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let row_sxvd = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B1).then_some(payload))
        .find(|payload| u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001)
        .expect("row SXVD record");
    assert_eq!(u16::from_le_bytes(row_sxvd[2..4].try_into().unwrap()), 11);
    assert_eq!(
        u16::from_le_bytes(row_sxvd[4..6].try_into().unwrap()),
        0x0FFE,
        "all custom subtotal bits should be set"
    );
    assert_eq!(u16::from_le_bytes(row_sxvd[8..10].try_into().unwrap()), 6);
    assert_eq!(xls_unicode_string_no_cch_at(row_sxvd, 10, 6), "Market");

    let row_sxvd_position = workbook_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .unwrap();
    let subtotal_item_types = workbook_records[row_sxvd_position + 1..]
        .iter()
        .take_while(|(record_type, _)| *record_type == 0x00B2)
        .skip(3)
        .map(|(_, payload)| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(
        subtotal_item_types,
        vec![
            0x0002, 0x0003, 0x0004, 0x0005, 0x0006, 0x0007, 0x0008, 0x0009, 0x000A, 0x000B, 0x000C
        ]
    );
    let row_sxvdex = workbook_records[row_sxvd_position + 4..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX record");
    let grbit1 = u32::from_le_bytes(row_sxvdex[0..4].try_into().unwrap());
    assert_eq!(grbit1 >> 24, 25, "item-page count should be encoded");
    assert_ne!(grbit1 & 0x0001, 0, "show-all-items flag should be set");
    assert_ne!(grbit1 & 0x0200, 0, "autosort should be enabled");
    assert_eq!(grbit1 & 0x0400, 0, "autosort should be descending");
    assert_ne!(grbit1 & 0x4000, 0, "page breaks should be enabled");
    assert_eq!(
        grbit1 & 0x8000,
        0,
        "fHideNewItems should stay clear for Excel-compatible XLS output"
    );
    assert_ne!(
        grbit1 & 0x0040_0000,
        0,
        "SXVDEX grbit1 should request blank rows"
    );
    assert_eq!(
        grbit1 & 0x0080_0000,
        0,
        "SXVDEX grbit1 should request bottom subtotals"
    );
    assert_eq!(
        row_sxvdex[4], 0xFF,
        "SXVDEX grbit2 should keep Excel's sentinel"
    );
    assert_eq!(row_sxvdex[5], 0xFF);
    assert_eq!(i16::from_le_bytes(row_sxvdex[6..8].try_into().unwrap()), -1);
    assert_eq!(
        u16::from_le_bytes(row_sxvdex[10..12].try_into().unwrap()),
        17
    );
    assert_eq!(
        xls_unicode_string_no_cch_at(row_sxvdex, 20, 17),
        "Regional subtotal"
    );
    let field_ver10 = workbook_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == 0x0864
                && payload.len() >= 12
                && payload[4] == 0x01
                && payload[5] == 0x02)
                .then_some(payload)
        })
        .expect("field SXADDL version-10 options");
    assert_eq!(
        u32::from_le_bytes(field_ver10[6..10].try_into().unwrap()) & 0x0000_0001,
        1,
        "field SXADDL should hide the dropdown"
    );
}

#[test]
fn reads_writer_xls_axis_field_options() {
    let mut wb = Workbook::new();
    add_axis_options_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("AxisOptions")
        .expect("AxisOptions pivot");
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
    assert_eq!(field.item_page_count, 25);
}

#[test]
fn semantic_pivot_tables_emit_xls_measure_sorted_display_order() {
    let mut wb = Workbook::new();
    add_sort_by_measure_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let row_sxvd_position = workbook_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD record");
    let row_sxvdex = workbook_records[row_sxvd_position + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("row SXVDEX record");
    let grbit1 = u32::from_le_bytes(row_sxvdex[0..4].try_into().unwrap());

    assert_ne!(grbit1 & 0x0200, 0, "autosort should be enabled");
    assert_eq!(grbit1 & 0x0400, 0, "autosort should be descending");
    assert_eq!(
        i16::from_le_bytes(row_sxvdex[6..8].try_into().unwrap()),
        -1,
        "BIFF8 value sort keeps the data field implicit"
    );

    let row_sxli = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload))
        .expect("row SXLI record");
    assert_eq!(
        u16::from_le_bytes(row_sxli[8..10].try_into().unwrap()),
        1,
        "highest-revenue row item should be emitted first"
    );
}

#[test]
fn reads_writer_xls_sort_by_measure_as_legacy_autosort() {
    let mut wb = Workbook::new();
    add_sort_by_measure_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ValueSortedPivot")
        .expect("ValueSortedPivot pivot");
    let field = &pivot.rows[0];

    assert_eq!(field.sort, PivotSort::Descending);
    assert!(field.sort_by_measure.is_none());
}

#[test]
fn semantic_pivot_tables_emit_xls_multi_measure_values_axis_records() {
    let mut wb = Workbook::new();
    add_multi_measure_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let sxview = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW record");
    assert_eq!(u16::from_le_bytes(sxview[22..24].try_into().unwrap()), 3);
    assert_eq!(u16::from_le_bytes(sxview[24..26].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[26..28].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxview[30..32].try_into().unwrap()), 2);
    assert_eq!(u16::from_le_bytes(sxview[34..36].try_into().unwrap()), 2);

    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxvd_axes, vec![0x0001, 0x0008, 0x0008]);

    let axis_declarations = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B4).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(axis_declarations, vec![vec![0, 0], vec![0xFE, 0xFF]]);

    let data_field_aggregates = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00C5).then(|| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(data_field_aggregates, vec![0, 2]);

    let sxli_lengths = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B5).then_some(payload.len()))
        .collect::<Vec<_>>();
    assert_eq!(sxli_lengths, vec![30, 20]);
}

#[test]
fn semantic_pivot_tables_emit_xls_all_aggregate_data_field_records() {
    let mut wb = Workbook::new();
    add_all_aggregate_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let data_field_codes = workbook_records
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
        "BIFF8 SXDI records should preserve every pivot aggregate code"
    );

    let sxview = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B0).then_some(payload))
        .expect("SXVIEW record");
    assert_eq!(
        u16::from_le_bytes(sxview[30..32].try_into().unwrap()),
        xls_all_aggregate_cases().len() as u16,
        "SXVIEW should declare every aggregate as a data field"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_custom_measure_number_format() {
    let mut wb = Workbook::new();
    add_custom_measure_format_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let format_record = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x041E).then_some(payload))
        .expect("custom FORMAT record");
    assert_eq!(
        u16::from_le_bytes(format_record[0..2].try_into().unwrap()),
        164
    );
    assert_eq!(xls_unicode_string_at(format_record, 2), "#,##0.0");

    let sxdi = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C5).then_some(payload))
        .expect("SXDI record");
    assert_eq!(u16::from_le_bytes(sxdi[10..12].try_into().unwrap()), 164);
}

#[test]
fn reads_writer_xls_custom_measure_number_format_semantics() {
    let mut wb = Workbook::new();
    add_custom_measure_format_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = &read.worksheet(0).unwrap().pivot_tables()[0];

    assert_eq!(pivot.name, "FormattedRevenue");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].number_format.as_deref(), Some("#,##0.0"));
}

#[test]
fn semantic_pivot_tables_emit_xls_show_as_data_field_records() {
    let mut wb = Workbook::new();
    add_show_as_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let data_fields = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C5).then_some(payload))
        .collect::<Vec<_>>();
    let show_as_codes = data_fields
        .iter()
        .map(|payload| u16::from_le_bytes(payload[4..6].try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(show_as_codes, vec![5, 6, 7, 8, 4, 1, 3]);

    let base_field_indexes = data_fields
        .iter()
        .map(|payload| u16::from_le_bytes(payload[6..8].try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(base_field_indexes, vec![0, 0, 0, 0, 0, 0, 0]);

    let base_item_indexes = data_fields
        .iter()
        .map(|payload| u16::from_le_bytes(payload[8..10].try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(base_item_indexes, vec![0, 0, 0, 0, 0, 0, 0]);
}

#[test]
fn reads_writer_xls_show_as_data_field_semantics() {
    let mut wb = Workbook::new();
    add_show_as_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ShowAsPivot")
        .expect("ShowAsPivot pivot");

    assert_eq!(pivot.measures.len(), 7);
    assert_eq!(pivot.measures[0].show_as, PivotShowAs::PercentOfRowTotal);
    assert_eq!(pivot.measures[1].show_as, PivotShowAs::PercentOfColumnTotal);
    assert_eq!(pivot.measures[2].show_as, PivotShowAs::PercentOfGrandTotal);
    assert_eq!(pivot.measures[3].show_as, PivotShowAs::Index);
    assert_eq!(
        pivot.measures[4].show_as,
        PivotShowAs::RunningTotal {
            base_field: PivotFieldRef::new("Region")
        }
    );
    assert_eq!(
        pivot.measures[5].show_as,
        PivotShowAs::DifferenceFrom {
            base_field: PivotFieldRef::new("Region"),
            base_item: PivotValue::String("East".to_string())
        }
    );
    assert_eq!(
        pivot.measures[6].show_as,
        PivotShowAs::PercentDifferenceFrom {
            base_field: PivotFieldRef::new("Region"),
            base_item: PivotValue::String("East".to_string())
        }
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_style_record() {
    let mut wb = Workbook::new();
    add_styled_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let style_record = workbook_records
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
        .expect("pivot style FRT record");

    assert_eq!(
        u16::from_le_bytes(style_record[12..14].try_into().unwrap()),
        0x002E,
        "style flags should encode last-column, row/column stripes, and column headers"
    );
    assert_eq!(utf16le_string(&style_record[16..]), "PivotStyleLight16");
}

#[test]
fn semantic_pivot_tables_emit_xls_calculated_field_records() {
    let mut wb = Workbook::new();
    add_calculated_field_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let calculated_field = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Revenue")
        .expect("calculated-field SXFDB record");
    assert_ne!(
        u16::from_le_bytes(calculated_field[0..2].try_into().unwrap()) & 0x8000,
        0,
        "SXFDB should mark Revenue as a calculated field"
    );
    assert_eq!(
        u16::from_le_bytes(calculated_field[0..2].try_into().unwrap()),
        0x8425,
        "calculated field SXFDB flags should match Excel's BIFF8 encoding"
    );

    let sxfmla = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F9).then_some(payload))
        .expect("SxFmla record");
    assert_eq!(u16::from_le_bytes(sxfmla[0..2].try_into().unwrap()), 13);
    assert_eq!(u16::from_le_bytes(sxfmla[2..4].try_into().unwrap()), 2);
    assert_eq!(
        &sxfmla[4..],
        &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x05],
        "SxFmla should encode Units * Price via two PtgSxName refs"
    );

    let sxname_field_indexes = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00F6).then(|| i16::from_le_bytes(payload[2..4].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxname_field_indexes, vec![1, 2]);

    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let data_sxvd_position = workbook_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0008
        })
        .expect("calculated data-field SXVD record");
    let data_sxvd = &workbook_records[data_sxvd_position].1;
    assert_eq!(u16::from_le_bytes(data_sxvd[2..4].try_into().unwrap()), 0);
    assert_eq!(u16::from_le_bytes(data_sxvd[4..6].try_into().unwrap()), 0);
    let data_sxvdex = workbook_records[data_sxvd_position + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0100).then_some(payload))
        .expect("calculated data-field SXVDEX record");
    assert_eq!(
        u32::from_le_bytes(data_sxvdex[0..4].try_into().unwrap()),
        0x0AA0_3410
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_calculated_field_function_records() {
    let mut wb = Workbook::new();
    add_calculated_field_function_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let sxfmla = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F9).then_some(payload))
        .expect("SxFmla record");
    assert_eq!(u16::from_le_bytes(sxfmla[0..2].try_into().unwrap()), 16);
    assert_eq!(u16::from_le_bytes(sxfmla[2..4].try_into().unwrap()), 2);
    assert_eq!(
        &sxfmla[4..],
        &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x42, 0x02, 0x04, 0x00],
        "SxFmla should encode SUM(Units,Price)"
    );

    let sxname_field_indexes = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00F6).then(|| i16::from_le_bytes(payload[2..4].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxname_field_indexes, vec![1, 2]);
}

#[test]
fn reads_writer_xls_calculated_field_function_semantics() {
    let mut wb = Workbook::new();
    add_calculated_field_function_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let wb2 = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

    assert_eq!(pivot.name, "CalculatedRevenueFunction");
    assert_eq!(pivot.calculated_fields.len(), 1);
    assert_eq!(pivot.calculated_fields[0].name, "Revenue");
    assert_eq!(pivot.calculated_fields[0].formula, "SUM(Units,Price)");
}

#[test]
fn semantic_pivot_tables_emit_xls_calculated_item_records() {
    let mut wb = Workbook::new();
    add_calculated_item_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let sxfmla = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F9).then_some(payload))
        .expect("SxFmla calculated-item record");
    assert_eq!(u16::from_le_bytes(sxfmla[0..2].try_into().unwrap()), 13);
    assert_eq!(u16::from_le_bytes(sxfmla[2..4].try_into().unwrap()), 2);
    assert_eq!(
        &sxfmla[4..],
        &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x03],
        "SxFmla should encode East + West via two PtgSxName refs"
    );

    let sxformula = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x0103).then_some(payload))
        .expect("SXFORMULA calculated-item record");
    assert_eq!(&sxformula[0..4], &[0x00, 0x00, 0xFF, 0xFF]);

    let sxnames = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00F6).then_some(payload))
        .collect::<Vec<_>>();
    assert_eq!(sxnames.len(), 2);
    assert_eq!(
        &sxnames[0][0..8],
        &[0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00]
    );
    assert_eq!(
        &sxnames[1][0..8],
        &[0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00]
    );

    let sxpairs = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00F8).then_some(payload))
        .map(|payload| {
            (
                u16::from_le_bytes(payload[0..2].try_into().unwrap()),
                u16::from_le_bytes(payload[2..4].try_into().unwrap()),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(sxpairs, vec![(0, 0), (0, 1)]);

    let region = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Region")
        .expect("Region SXFDB");
    assert_eq!(u16::from_le_bytes(region[6..8].try_into().unwrap()), 3);
    let items = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00CD).then_some(payload))
        .map(|payload| xls_unicode_string_at(payload, 0))
        .take(3)
        .collect::<Vec<_>>();
    assert_eq!(items, vec!["East", "West", "Combined"]);

    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let row_field_record_index = workbook_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00B1
                && u16::from_le_bytes(payload[0..2].try_into().unwrap()) == 0x0001
        })
        .expect("row SXVD record");
    let row_field_index = workbook_records[..=row_field_record_index]
        .iter()
        .filter(|(record_type, _)| *record_type == 0x00B1)
        .count()
        - 1;
    let row_items = sxvi_payloads_for_sxvd(&workbook_records, row_field_index);
    assert!(
        row_items
            .iter()
            .any(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()) & 0x0008 != 0),
        "row pivot item should be marked as calculated"
    );
}

#[test]
fn reads_writer_xls_calculated_item_semantics() {
    let mut wb = Workbook::new();
    add_calculated_item_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let wb2 = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

    assert_eq!(pivot.name, "CalculatedRegion");
    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Region");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("Combined".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "East+West");
}

#[test]
fn semantic_pivot_tables_emit_xls_calculated_item_cell_like_records() {
    let mut wb = Workbook::new();
    add_calculated_item_cell_like_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let sxfmla = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F9).then_some(payload))
        .expect("SxFmla calculated-item record");
    assert_eq!(u16::from_le_bytes(sxfmla[0..2].try_into().unwrap()), 13);
    assert_eq!(u16::from_le_bytes(sxfmla[2..4].try_into().unwrap()), 2);
    assert_eq!(
        &sxfmla[4..],
        &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x03],
        "cell-like calculated item names should encode as PtgSxName refs"
    );

    let sxpairs = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00F8).then_some(payload))
        .map(|payload| {
            (
                u16::from_le_bytes(payload[0..2].try_into().unwrap()),
                u16::from_le_bytes(payload[2..4].try_into().unwrap()),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(sxpairs, vec![(0, 0), (0, 1)]);
}

#[test]
fn reads_writer_xls_calculated_item_cell_like_semantics() {
    let mut wb = Workbook::new();
    add_calculated_item_cell_like_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let wb2 = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

    assert_eq!(pivot.name, "CalculatedQuarter");
    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Quarter");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("H1".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "Q1+Q2");
}

#[test]
fn semantic_pivot_tables_emit_xls_calculated_item_string_ref_records() {
    let mut wb = Workbook::new();
    add_calculated_item_string_ref_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let formula_lengths = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00F9).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        formula_lengths,
        vec![13],
        "string item references should compile to the same PtgSxName shape as bare item refs"
    );

    let sxpairs = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00F8).then_some(payload))
        .map(|payload| {
            (
                u16::from_le_bytes(payload[0..2].try_into().unwrap()),
                u16::from_le_bytes(payload[2..4].try_into().unwrap()),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(sxpairs.len(), 2);
}

#[test]
fn reads_writer_xls_calculated_item_string_ref_semantics() {
    let mut wb = Workbook::new();
    add_calculated_item_string_ref_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let wb2 = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

    assert_eq!(pivot.name, "CalculatedRegionStringRef");
    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Region");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("Combined".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "East+West");
}

#[test]
fn xls_writer_rejects_calculated_item_dependencies() {
    let mut wb = Workbook::new();
    add_calculated_item_dependent_string_ref_pivot(&mut wb);

    let err = XlsWriter::write_to_bytes(&wb)
        .expect_err("dependent XLS calculated items should be rejected");
    assert!(
        err.to_string()
            .contains("references another calculated item"),
        "{err}"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_calculated_item_function_records() {
    let mut wb = Workbook::new();
    add_calculated_item_function_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let sxfmla = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F9).then_some(payload))
        .expect("SxFmla calculated-item record");
    assert_eq!(u16::from_le_bytes(sxfmla[0..2].try_into().unwrap()), 16);
    assert_eq!(u16::from_le_bytes(sxfmla[2..4].try_into().unwrap()), 2);
    assert_eq!(
        &sxfmla[4..],
        &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x42, 0x02, 0x07, 0x00],
        "SxFmla should encode MAX(East,West)"
    );
    let sxpairs = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00F8).then_some(payload))
        .map(|payload| {
            (
                u16::from_le_bytes(payload[0..2].try_into().unwrap()),
                u16::from_le_bytes(payload[2..4].try_into().unwrap()),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(sxpairs, vec![(0, 0), (0, 1)]);
}

#[test]
fn reads_writer_xls_calculated_item_function_semantics() {
    let mut wb = Workbook::new();
    add_calculated_item_function_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let wb2 = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

    assert_eq!(pivot.name, "CalculatedRegionFunction");
    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Region");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("Combined".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "MAX(East,West)");
}

#[test]
fn semantic_pivot_tables_emit_xls_numeric_grouping_records() {
    let mut wb = Workbook::new();
    add_numeric_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let age_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Age")
        .expect("Age SXFDB record");
    let flags = u16::from_le_bytes(age_sxfdb[0..2].try_into().unwrap());
    assert_ne!(
        flags & 0x0010,
        0,
        "SXFDB should mark the field as a range group"
    );
    assert_ne!(
        flags & 0x0020,
        0,
        "SXFDB should mark the grouped field as numeric"
    );
    assert_eq!(
        u16::from_le_bytes(age_sxfdb[8..10].try_into().unwrap()),
        4,
        "SXFDB csxoper should count the group labels"
    );

    let sxrng = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00D8).then_some(payload))
        .expect("SXRng record");
    assert_eq!(
        u16::from_le_bytes(sxrng[0..2].try_into().unwrap()),
        0x0040,
        "SXRng should use explicit numeric start/end and range grouping"
    );

    let sxnums = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00C9).then(|| f64::from_le_bytes(payload[0..8].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert!(
        sxnums.windows(3).any(|values| values == [0.0, 60.0, 10.0]),
        "SXRng should be followed by SXNum start/end/interval records"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_date_grouping_records() {
    let mut wb = Workbook::new();
    add_date_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let date_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Date")
        .expect("Date SXFDB record");
    let flags = u16::from_le_bytes(date_sxfdb[0..2].try_into().unwrap());
    assert_eq!(
        flags, 0x0909,
        "source date SXFDB should link to the derived grouped field"
    );
    assert_eq!(
        i16::from_le_bytes(date_sxfdb[2..4].try_into().unwrap()),
        2,
        "source date SXFDB should point at the derived month field"
    );
    assert_eq!(
        u16::from_le_bytes(date_sxfdb[12..14].try_into().unwrap()),
        4,
        "source date SXFDB should count source date atoms"
    );

    let month_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Months (Date)")
        .expect("derived month SXFDB record");
    assert_eq!(
        u16::from_le_bytes(month_sxfdb[0..2].try_into().unwrap()),
        0x0011,
        "derived date SXFDB should use Excel's grouped date flags"
    );
    assert_eq!(
        i16::from_le_bytes(month_sxfdb[4..6].try_into().unwrap()),
        0,
        "derived date SXFDB should link back to the source Date field"
    );
    assert_eq!(
        u16::from_le_bytes(month_sxfdb[6..8].try_into().unwrap()),
        14,
        "derived month field should include boundary items plus twelve months"
    );

    let sxrng_position = cache_records
        .iter()
        .position(|(record_type, _)| *record_type == 0x00D8)
        .expect("SXRng record");
    let sxrng = &cache_records[sxrng_position].1;
    assert_eq!(
        u16::from_le_bytes(sxrng[0..2].try_into().unwrap()),
        0x0017,
        "SXRng should encode automatic month date grouping"
    );

    let date_source_atoms = cache_records[..sxrng_position]
        .iter()
        .filter(|(record_type, _)| *record_type == 0x00CE)
        .collect::<Vec<_>>();
    assert_eq!(
        date_source_atoms.len(),
        4,
        "source date field should carry the four date atoms as SXDTR records"
    );

    let group_items = cache_records[..sxrng_position]
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00CD).then(|| xls_unicode_string_at(payload, 0))
        })
        .collect::<Vec<_>>();
    assert!(
        group_items
            .windows(4)
            .any(|values| values == ["Jan", "Feb", "Mar", "Apr"]),
        "derived month field should carry Excel-style month labels"
    );

    let date_group_tail = cache_records[sxrng_position + 1..]
        .iter()
        .take(3)
        .collect::<Vec<_>>();
    assert_eq!(
        date_group_tail
            .iter()
            .map(|(record_type, _)| *record_type)
            .collect::<Vec<_>>(),
        vec![0x00CE, 0x00CE, 0x00CC],
        "derived date grouping should emit start/end/interval after SXRng"
    );
    assert_eq!(
        date_group_tail[0].1,
        vec![0xE4, 0x07, 0x01, 0x00, 0x01, 0, 0, 0],
        "date grouping start should be 2020-01-01"
    );
    assert_eq!(
        date_group_tail[1].1,
        vec![0xE4, 0x07, 0x04, 0x00, 0x01, 0, 0, 0],
        "date grouping end should preserve the max source date"
    );
    assert_eq!(
        date_group_tail[2].1,
        vec![0x01, 0x00],
        "date grouping interval should be one unit"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_multi_unit_date_grouping_records() {
    let mut wb = Workbook::new();
    add_multi_unit_date_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let sxdb = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C6).then_some(payload))
        .expect("SXDB record");
    assert_eq!(
        u16::from_le_bytes(sxdb[10..12].try_into().unwrap()),
        2,
        "multi-unit date grouping should keep transformed snapshot fields out of the BIFF8 base cache"
    );
    assert_eq!(
        u16::from_le_bytes(sxdb[12..14].try_into().unwrap()),
        4,
        "multi-unit date grouping should append one BIFF8 derived field per unit"
    );

    let sale_date_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "SaleDate")
        .expect("SaleDate SXFDB record");
    assert_eq!(
        u16::from_le_bytes(sale_date_sxfdb[0..2].try_into().unwrap()),
        0x0909,
        "base date field should use grouped date source flags"
    );
    assert_eq!(
        i16::from_le_bytes(sale_date_sxfdb[2..4].try_into().unwrap()),
        3,
        "base date field should point at the outer derived year field"
    );
    assert_eq!(
        u16::from_le_bytes(sale_date_sxfdb[12..14].try_into().unwrap()),
        3,
        "base date field should count source date atoms"
    );

    let years_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Years (SaleDate)")
        .expect("derived year SXFDB record");
    assert_eq!(
        i16::from_le_bytes(years_sxfdb[4..6].try_into().unwrap()),
        0,
        "derived year field should link back to SaleDate"
    );
    assert_eq!(
        u16::from_le_bytes(years_sxfdb[6..8].try_into().unwrap()),
        4,
        "derived year field should include boundary items plus 2024 and 2025"
    );

    let months_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Months (SaleDate)")
        .expect("derived month SXFDB record");
    assert_eq!(
        i16::from_le_bytes(months_sxfdb[4..6].try_into().unwrap()),
        0,
        "derived month field should link back to SaleDate"
    );
    assert_eq!(
        u16::from_le_bytes(months_sxfdb[6..8].try_into().unwrap()),
        14,
        "derived month field should include boundary items plus twelve months"
    );

    let range_group_flags = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00D8).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        range_group_flags,
        vec![0x0017, 0x001F],
        "derived fields should encode the BIFF8 cache from inner to outer date unit"
    );

    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        sxvd_axes,
        vec![0x0000, 0x0008, 0x0001, 0x0001],
        "multi-unit date grouping should hide the base date field and put derived units on the row axis"
    );
    let row_fields = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B4).then_some(payload))
        .expect("row SXIVD record");
    assert_eq!(
        row_fields,
        &vec![3, 0, 2, 0],
        "multi-unit date grouping should expand the row axis to year plus month fields"
    );
}

#[test]
fn xls_multi_unit_date_grouping_keeps_source_name_collisions() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "SaleDate").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("C1", "SaleDate (Years)").unwrap();
    ws.set_cell_value("D1", "SaleDate (Months)").unwrap();
    ws.set_cell_value("A2", 45292.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("C2", "source years").unwrap();
    ws.set_cell_value("D2", "source months").unwrap();
    ws.set_cell_value("A3", 45323.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("C3", "source years").unwrap();
    ws.set_cell_value("D3", "source months").unwrap();
    ws.set_cell_value("A4", 45658.0).unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("C4", "source years").unwrap();
    ws.set_cell_value("D4", "source months").unwrap();

    let pivot = PivotTable::builder("GroupedDates")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("SaleDate")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: PivotFieldRef::new("SaleDate"),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);
    let sxdb = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C6).then_some(payload))
        .expect("SXDB record");
    assert_eq!(
        u16::from_le_bytes(sxdb[10..12].try_into().unwrap()),
        4,
        "real source columns named like transformed date fields must remain base fields"
    );
    assert_eq!(
        u16::from_le_bytes(sxdb[12..14].try_into().unwrap()),
        6,
        "writer should skip only transformed snapshot date fields and then append BIFF8 date unit fields"
    );

    let sxfdb_names = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00C7).then(|| xls_unicode_string_at(payload, 14))
        })
        .collect::<Vec<_>>();
    assert!(sxfdb_names.iter().any(|name| name == "SaleDate (Years)"));
    assert!(sxfdb_names.iter().any(|name| name == "SaleDate (Months)"));
    assert!(sxfdb_names.iter().any(|name| name == "Years (SaleDate)"));
    assert!(sxfdb_names.iter().any(|name| name == "Months (SaleDate)"));
}

#[test]
fn xls_multi_unit_date_grouping_keeps_calculated_name_collisions() {
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
        .calculated_field("SaleDate (Years)", "=Revenue*1")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: PivotFieldRef::new("SaleDate"),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);
    let sxdb = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00C6).then_some(payload))
        .expect("SXDB record");
    assert_eq!(
        u16::from_le_bytes(sxdb[10..12].try_into().unwrap()),
        2,
        "base field count should include worksheet source fields, not calculated or derived fields"
    );
    assert_eq!(
        u16::from_le_bytes(sxdb[12..14].try_into().unwrap()),
        5,
        "writer should skip transformed date fields without dropping calculated fields"
    );

    let calculated_years = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "SaleDate (Years)")
        .expect("calculated field with colliding name");
    assert_ne!(
        u16::from_le_bytes(calculated_years[0..2].try_into().unwrap()) & 0x8000,
        0,
        "calculated field should retain its formula flag"
    );
    assert!(cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .any(|payload| xls_unicode_string_at(payload, 14) == "Years (SaleDate)"));
    assert!(cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .any(|payload| xls_unicode_string_at(payload, 14) == "Months (SaleDate)"));
}

#[test]
fn semantic_pivot_tables_emit_xls_manual_grouping_records() {
    let mut wb = Workbook::new();
    add_manual_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let region_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Region")
        .expect("Region SXFDB record");
    assert_eq!(
        u16::from_le_bytes(region_sxfdb[0..2].try_into().unwrap()),
        0x0489,
        "source manual field should use Excel's grouped item flags"
    );
    assert_eq!(
        i16::from_le_bytes(region_sxfdb[2..4].try_into().unwrap()),
        2,
        "source manual field should point at the derived grouped field"
    );
    assert_eq!(
        u16::from_le_bytes(region_sxfdb[12..14].try_into().unwrap()),
        6,
        "source manual field should count source atoms"
    );

    let derived_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Region2")
        .expect("derived manual SXFDB record");
    assert_eq!(
        u16::from_le_bytes(derived_sxfdb[0..2].try_into().unwrap()),
        0x0001,
        "derived manual field should use Excel's grouped item flags"
    );
    assert_eq!(
        i16::from_le_bytes(derived_sxfdb[4..6].try_into().unwrap()),
        0,
        "derived manual field should link back to the source field"
    );
    assert_eq!(
        u16::from_le_bytes(derived_sxfdb[6..8].try_into().unwrap()),
        3,
        "derived manual field should carry the ungrouped item plus group names"
    );
    assert_eq!(
        u16::from_le_bytes(derived_sxfdb[10..12].try_into().unwrap()),
        6,
        "derived manual field should count source atoms"
    );

    let group_items = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00CD).then(|| xls_unicode_string_at(payload, 0))
        })
        .collect::<Vec<_>>();
    assert!(
        group_items
            .windows(3)
            .any(|values| values == ["International", "Coastal", "Interior"]),
        "derived manual field should carry ungrouped items followed by all group names"
    );

    let sxidstm = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00D9).then_some(payload))
        .expect("manual grouping item map record");
    let item_map = sxidstm
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes(chunk.try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(
        item_map,
        vec![2, 1, 1, 2, 1, 0],
        "manual grouping should map source atoms to ungrouped/group item indexes"
    );

    let row_markers = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C8).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        row_markers,
        vec![vec![0], vec![1], vec![2], vec![3], vec![4], vec![5]],
        "SXDBB rows should store source item ids, not derived group ids"
    );

    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        sxvd_axes,
        vec![0x0001, 0x0008, 0x0001],
        "manual grouping should put source and derived fields on the row axis"
    );
    let row_fields = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B4).then_some(payload))
        .expect("row SXIVD record");
    assert_eq!(
        row_fields,
        &vec![2, 0, 0, 0],
        "manual grouping should expand the row axis to derived plus source"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_numeric_manual_grouping_records() {
    let mut wb = Workbook::new();
    add_manual_numeric_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let age_index = cache_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00C7 && xls_unicode_string_at(payload, 14) == "Age"
        })
        .expect("Age SXFDB record");
    let revenue_index = cache_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00C7 && xls_unicode_string_at(payload, 14) == "Revenue"
        })
        .expect("Revenue SXFDB record");
    let derived_index = cache_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00C7 && xls_unicode_string_at(payload, 14) == "Age2"
        })
        .expect("Age2 SXFDB record");

    let age_sxfdb = &cache_records[age_index].1;
    assert_eq!(
        u16::from_le_bytes(age_sxfdb[0..2].try_into().unwrap()),
        0x0569,
        "numeric manual source field should use Excel's numeric grouped item flags"
    );
    assert_eq!(
        i16::from_le_bytes(age_sxfdb[2..4].try_into().unwrap()),
        2,
        "numeric manual source field should point at the derived grouped field"
    );

    let age_numbers = cache_records[age_index + 1..revenue_index]
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00C9).then(|| f64::from_le_bytes(payload[0..8].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        age_numbers,
        vec![7.0, 13.0, 21.0, 34.0, 55.0],
        "numeric manual source items should be emitted as SXNUM records"
    );

    let derived_sxfdb = &cache_records[derived_index].1;
    assert_eq!(
        i16::from_le_bytes(derived_sxfdb[4..6].try_into().unwrap()),
        0,
        "numeric manual derived field should link back to Age"
    );
    assert_eq!(
        u16::from_le_bytes(derived_sxfdb[6..8].try_into().unwrap()),
        3,
        "derived numeric manual field should contain the ungrouped item plus group names"
    );

    let derived_until = cache_records[derived_index + 1..]
        .iter()
        .position(|(record_type, _)| *record_type == 0x00D9)
        .map(|offset| derived_index + 1 + offset)
        .expect("numeric manual SXIDSTM record");
    let derived_numbers = cache_records[derived_index + 1..derived_until]
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00C9).then(|| f64::from_le_bytes(payload[0..8].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        derived_numbers,
        vec![55.0],
        "ungrouped numeric manual items should stay numeric in the derived field"
    );

    let sxidstm = &cache_records[derived_until].1;
    let item_map = sxidstm
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes(chunk.try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(
        item_map,
        vec![1, 1, 2, 2, 0],
        "numeric manual grouping should map source items to group/ungrouped item indexes"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_bool_error_manual_grouping_records() {
    let mut wb = Workbook::new();
    add_manual_bool_error_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let flag_index = cache_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00C7 && xls_unicode_string_at(payload, 14) == "Flag"
        })
        .expect("Flag SXFDB record");
    let revenue_index = cache_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00C7 && xls_unicode_string_at(payload, 14) == "Revenue"
        })
        .expect("Revenue SXFDB record");
    let derived_index = cache_records
        .iter()
        .position(|(record_type, payload)| {
            *record_type == 0x00C7 && xls_unicode_string_at(payload, 14) == "Flag2"
        })
        .expect("Flag2 SXFDB record");

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
        "boolean and error manual source items should use Excel's SXBool/SXErr records"
    );

    let sxidstm = cache_records[derived_index + 1..]
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00D9).then_some(payload))
        .expect("bool/error manual SXIDSTM record");
    let item_map = sxidstm
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes(chunk.try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(
        item_map,
        vec![0, 0, 1, 1],
        "bool/error manual grouping should map booleans and errors to group item indexes"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_manual_column_grouping_records() {
    let mut wb = Workbook::new();
    add_manual_column_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let region_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Region")
        .expect("Region SXFDB record");
    assert_eq!(
        u16::from_le_bytes(region_sxfdb[0..2].try_into().unwrap()),
        0x0489,
        "source manual column field should use grouped item flags"
    );
    assert_eq!(
        i16::from_le_bytes(region_sxfdb[2..4].try_into().unwrap()),
        3,
        "source manual column field should point at the derived grouped field"
    );

    let quarter_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Quarter")
        .expect("Quarter SXFDB record");
    assert_eq!(
        u16::from_le_bytes(quarter_sxfdb[12..14].try_into().unwrap()),
        3,
        "regular text fields should count all emitted source atoms"
    );

    let derived_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Region2")
        .expect("derived manual column SXFDB record");
    assert_eq!(
        i16::from_le_bytes(derived_sxfdb[4..6].try_into().unwrap()),
        0,
        "derived manual column field should link back to Region"
    );

    let row_markers = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C8).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        row_markers,
        vec![
            vec![0, 0],
            vec![1, 0],
            vec![2, 1],
            vec![3, 1],
            vec![4, 2],
            vec![5, 2],
        ],
        "manual column grouping cache rows should include every text base field item id"
    );

    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        sxvd_axes,
        vec![0x0002, 0x0001, 0x0008, 0x0002],
        "manual column grouping should put source and derived fields on the column axis"
    );

    let axis_fields = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B4).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        axis_fields,
        vec![vec![1, 0], vec![3, 0, 0, 0]],
        "manual column grouping should expand the column axis to derived plus source"
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_manual_page_grouping_records() {
    let mut wb = Workbook::new();
    add_manual_page_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let region_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Region")
        .expect("Region SXFDB record");
    assert_eq!(
        i16::from_le_bytes(region_sxfdb[2..4].try_into().unwrap()),
        3,
        "source manual page field should point at the derived grouped field"
    );

    let derived_sxfdb = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C7).then_some(payload))
        .find(|payload| xls_unicode_string_at(payload, 14) == "Region2")
        .expect("derived manual page SXFDB record");
    assert_eq!(
        i16::from_le_bytes(derived_sxfdb[4..6].try_into().unwrap()),
        0,
        "derived manual page field should link back to Region"
    );

    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let sxvd_axes = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00B1).then(|| u16::from_le_bytes(payload[0..2].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        sxvd_axes,
        vec![0x0000, 0x0001, 0x0008, 0x0004],
        "manual page grouping should hide the source field and put the derived field on the page axis"
    );

    let axis_fields = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B4).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        axis_fields,
        vec![vec![1, 0]],
        "manual page grouping should leave only Salesperson on the row axis"
    );

    let page_fields = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B6).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        page_fields,
        vec![vec![3, 0, 1, 0, 1, 0]],
        "manual page grouping should point SXPI at the derived group field and select Coastal"
    );
}

#[test]
fn semantic_pivot_tables_emit_unfiltered_xls_manual_page_grouping_records() {
    let mut wb = Workbook::new();
    add_unfiltered_manual_page_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);
    let page_fields = workbook_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00B6).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        page_fields,
        vec![vec![3, 0, 0xFD, 0x7F, 1, 0]],
        "unfiltered manual page grouping should use Excel's grouped-field all-items marker"
    );
}

#[test]
fn xls_writer_round_trips_manual_grouped_page_multi_item_filters() {
    let mut wb = Workbook::new();
    add_manual_page_grouped_pivot_with_filter_items(&mut wb, Some(&["Coastal", "Interior"]));

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let workbook_records = records_with_payload(&workbook);

    let page_field = workbook_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00B6).then_some(payload))
        .expect("SXPI page field");
    assert_eq!(
        page_field,
        &[3, 0, 0xFD, 0x7F, 1, 0],
        "multi-select manual page grouping should target the derived group field"
    );

    let derived_group_items = sxvi_payloads_for_sxvd(&workbook_records, 3);
    assert_eq!(
        derived_group_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[2..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![1, 0, 0, 0],
        "International should be hidden while Coastal, Interior, and the automatic subtotal item remain visible"
    );

    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = read
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualPageRegions")
        .expect("pivot after read");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal", "Interior"])]
    );
}

#[test]
fn semantic_pivot_tables_emit_xls_duplicate_date_cache_item_ids() {
    let mut wb = Workbook::new();
    add_duplicate_date_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let cache_records = records_with_payload(&cache);

    let row_markers = cache_records
        .iter()
        .filter_map(|(record_type, payload)| (*record_type == 0x00C8).then_some(payload.clone()))
        .collect::<Vec<_>>();
    assert_eq!(
        row_markers,
        vec![vec![0], vec![0], vec![1], vec![1]],
        "SXDBB rows should store source item ids, not source row ordinals"
    );
}

#[test]
fn reads_writer_xls_numeric_grouping_semantics() {
    let mut wb = Workbook::new();
    add_numeric_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "AgeBands");
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
}

#[test]
fn reads_writer_xls_date_grouping_semantics() {
    let mut wb = Workbook::new();
    add_date_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "MonthlyRevenue");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "Date");
            assert_eq!(*units, vec![PivotDateGroupUnit::Months]);
        }
        other => panic!("expected date grouping, got {other:?}"),
    }
}

#[test]
fn reads_writer_xls_multi_unit_date_grouping_semantics() {
    let mut wb = Workbook::new();
    add_multi_unit_date_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "GroupedDates");
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
}

#[test]
fn reads_writer_xls_mixed_groupings_in_source_order() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "SaleDate").unwrap();
    ws.set_cell_value("B1", "Bucket").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", 45292.0).unwrap();
    ws.set_cell_value("B2", 5.0).unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", 45323.0).unwrap();
    ws.set_cell_value("B3", 15.0).unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", 45658.0).unwrap();
    ws.set_cell_value("B4", 25.0).unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("MixedGroups")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("SaleDate")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: PivotFieldRef::new("SaleDate"),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .grouping(PivotGrouping::Number {
            field: PivotFieldRef::new("Bucket"),
            start: Some(0.0),
            end: Some(30.0),
            interval: 10.0,
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let pivot = &read.worksheet(0).unwrap().pivot_tables()[0];

    assert_eq!(pivot.groupings.len(), 2);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "SaleDate");
            assert_eq!(
                *units,
                vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months]
            );
        }
        other => panic!("expected first grouping to be date, got {other:?}"),
    }
    match &pivot.groupings[1] {
        PivotGrouping::Number {
            field,
            start,
            end,
            interval,
        } => {
            assert_eq!(field.name, "Bucket");
            assert_eq!(*start, Some(0.0));
            assert_eq!(*end, Some(30.0));
            assert_eq!(*interval, 10.0);
        }
        other => panic!("expected second grouping to be numeric, got {other:?}"),
    }
}

#[test]
fn reads_writer_xls_page_date_grouping_semantics() {
    let mut wb = Workbook::new();
    add_page_date_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "MonthlyPageFilter");
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.page_fields[0].field.name, "Date");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "Date");
            assert_eq!(*units, vec![PivotDateGroupUnit::Months]);
        }
        other => panic!("expected date grouping, got {other:?}"),
    }
    assert!(pivot.filters.iter().any(|filter| matches!(
        filter,
        PivotFilter::FieldItems {
            field,
            allowed_items,
        } if field.name == "Date" && allowed_items == &vec![PivotValue::String("Jan".to_string())]
    )));
}

#[test]
fn reads_writer_xls_manual_grouping_semantics() {
    let mut wb = Workbook::new();
    add_manual_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "ManualGroupedRegions");
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
}

#[test]
fn reads_writer_xls_manual_column_grouping_semantics() {
    let mut wb = Workbook::new();
    add_manual_column_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "ManualColumnGroups");
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
}

#[test]
fn reads_writer_xls_numeric_manual_grouping_semantics() {
    let mut wb = Workbook::new();
    add_manual_numeric_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "ManualAgeGroups");
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
}

#[test]
fn reads_writer_xls_bool_error_manual_grouping_semantics() {
    let mut wb = Workbook::new();
    add_manual_bool_error_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "ManualBoolErrorGroups");
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
}

#[test]
fn xls_manual_grouping_rejects_page_plus_row_axis_fields() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("ManualDuplicateAxisRegions")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .page("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: PivotFieldRef::new("Region"),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let err =
        XlsWriter::write_to_bytes(&wb).expect_err("duplicate-axis manual grouping should fail");
    assert!(
        err.to_string()
            .contains("on the page axis and another axis"),
        "{err}"
    );
}

#[test]
fn reads_writer_xls_manual_page_grouping_semantics() {
    let mut wb = Workbook::new();
    add_manual_page_grouped_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "ManualPageRegions");
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
}

#[test]
fn reads_writer_xls_pivot_table_semantics() {
    let mut wb = Workbook::new();
    add_page_column_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    assert_eq!(ws.pivot_tables().len(), 1);

    let pivot = &ws.pivot_tables()[0];
    assert_eq!(pivot.name, "ChannelPivot");
    assert_eq!(pivot.target.to_a1_string(), "F2");
    assert_eq!(
        pivot.source,
        PivotSource::range_on_sheet("Sheet1", CellRange::parse("A1:D4").unwrap())
    );
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.columns[0].field.name, "Quarter");
    assert_eq!(pivot.page_fields[0].field.name, "Segment");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
    assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Revenue"));
    assert!(pivot.filters.iter().any(|filter| matches!(
        filter,
        PivotFilter::FieldItems {
            field,
            allowed_items,
        } if field.name == "Segment"
            && allowed_items == &vec![PivotValue::String("Online".to_string())]
    )));
    let cache_info = pivot.cache_info().expect("cache diagnostics");
    assert_eq!(
        cache_info.source_kind,
        duke_sheets_core::PivotCacheSourceKind::Worksheet
    );
    assert_eq!(cache_info.record_count, Some(3));
}

#[test]
fn reads_writer_xls_multi_measure_values_axis_semantics() {
    let mut wb = Workbook::new();
    add_multi_measure_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    assert_eq!(ws.pivot_tables().len(), 1);

    let pivot = &ws.pivot_tables()[0];
    assert_eq!(pivot.name, "RevenueAndUnits");
    assert_eq!(pivot.target.to_a1_string(), "E1");
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert!(pivot.columns.is_empty());
    assert_eq!(pivot.layout.values_axis, PivotValuesAxis::Columns);
    assert_eq!(pivot.layout.values_axis_position, Some(0));
    assert_eq!(pivot.measures.len(), 2);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
    assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Revenue"));
    assert_eq!(pivot.measures[1].field.name, "Units");
    assert_eq!(pivot.measures[1].aggregate, PivotAggregate::Average);
    assert_eq!(pivot.measures[1].name.as_deref(), Some("Average Units"));
}

#[test]
fn reads_writer_xls_all_aggregate_data_field_semantics() {
    let mut wb = Workbook::new();
    add_all_aggregate_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "AllAggregatePivot");
    assert_eq!(pivot.layout.values_axis, PivotValuesAxis::Columns);
    assert_eq!(pivot.layout.values_axis_position, Some(0));
    assert_eq!(pivot.measures.len(), xls_all_aggregate_cases().len());
    for (measure, (aggregate, caption, _)) in pivot.measures.iter().zip(xls_all_aggregate_cases()) {
        assert_eq!(measure.field.name, "Revenue");
        assert_eq!(measure.aggregate, aggregate);
        assert_eq!(measure.name.as_deref(), Some(caption));
    }
}

#[test]
fn reads_writer_xls_pivot_style_semantics() {
    let mut wb = Workbook::new();
    add_styled_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.style, styled_pivot_style());
}

#[test]
fn reads_writer_xls_calculated_field_semantics() {
    let mut wb = Workbook::new();
    add_calculated_field_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let read = XlsReader::read(Cursor::new(bytes)).expect("read pivot workbook");
    let ws = read.worksheet(0).unwrap();
    let pivot = &ws.pivot_tables()[0];

    assert_eq!(pivot.name, "CalculatedRevenue");
    assert_eq!(pivot.calculated_fields.len(), 1);
    assert_eq!(pivot.calculated_fields[0].name, "Revenue");
    assert_eq!(pivot.calculated_fields[0].formula, "Units*Price");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
}

fn read_test_pivot_with_sxdb_source_type(source_type: u16) -> Workbook {
    let mut wb = Workbook::new();
    add_test_pivot(&mut wb);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize pivot workbook");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open cfb");
    let workbook = cfb.read_stream("/Workbook").expect("read workbook stream");
    let cache = cfb
        .read_stream("/_SX_DB_CUR/0001")
        .expect("read pivot cache stream");
    let patched_cache = patch_sxdb_source_type(&cache, source_type);

    let mut builder = CompoundFileBuilder::new();
    builder
        .add_stream("/Workbook", workbook)
        .expect("add workbook stream");
    builder
        .add_storage("/_SX_DB_CUR")
        .expect("add pivot cache storage");
    builder
        .add_stream("/_SX_DB_CUR/0001", patched_cache)
        .expect("add patched pivot cache stream");
    let patched = builder.build().expect("build patched XLS file");

    XlsReader::read(Cursor::new(patched)).expect("read patched pivot workbook")
}

fn patch_sxdb_source_type(stream: &[u8], source_type: u16) -> Vec<u8> {
    let mut saw_sxdb = false;
    let records = records_with_payload(stream)
        .into_iter()
        .map(|(record_type, mut payload)| {
            if record_type == 0x00C6 {
                assert!(payload.len() >= 18, "SXDB payload is too short");
                payload[16..18].copy_from_slice(&source_type.to_le_bytes());
                saw_sxdb = true;
            }
            (record_type, payload)
        })
        .collect::<Vec<_>>();
    assert!(saw_sxdb, "pivot cache stream did not contain SXDB");
    serialize_biff_records(&records)
}

fn record_types(stream: &[u8]) -> Vec<u16> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    while pos + 4 <= stream.len() {
        let typ = u16::from_le_bytes([stream[pos], stream[pos + 1]]);
        let len = u16::from_le_bytes([stream[pos + 2], stream[pos + 3]]) as usize;
        out.push(typ);
        pos += 4 + len;
    }
    out
}

fn records_with_payload(stream: &[u8]) -> Vec<(u16, Vec<u8>)> {
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

fn serialize_biff_records(records: &[(u16, Vec<u8>)]) -> Vec<u8> {
    let mut out = Vec::new();
    for (record_type, payload) in records {
        let len = u16::try_from(payload.len()).expect("BIFF record payload too large");
        out.extend_from_slice(&record_type.to_le_bytes());
        out.extend_from_slice(&len.to_le_bytes());
        out.extend_from_slice(payload);
    }
    out
}

fn sxvi_payloads_for_sxvd(records: &[(u16, Vec<u8>)], sxvd_index: usize) -> Vec<Vec<u8>> {
    let mut current_sxvd = None;
    let mut next_sxvd = 0usize;
    let mut payloads = Vec::new();
    for (record_type, payload) in records {
        match *record_type {
            0x00B1 => {
                current_sxvd = Some(next_sxvd);
                next_sxvd += 1;
            }
            0x00B2 if current_sxvd == Some(sxvd_index) => payloads.push(payload.clone()),
            0x0100 => current_sxvd = None,
            _ => {}
        }
    }
    payloads
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

fn sxli_line_item_indexes(payload: &[u8]) -> Vec<Vec<u16>> {
    let mut offset = 0usize;
    let mut lines = Vec::new();
    while offset + 8 <= payload.len() {
        let item_type = u16::from_le_bytes(payload[offset + 2..offset + 4].try_into().unwrap());
        let item_count =
            u16::from_le_bytes(payload[offset + 4..offset + 6].try_into().unwrap()) as usize + 1;
        let item_start = offset + 8;
        let item_end = item_start + item_count * 2;
        if item_end > payload.len() {
            break;
        }
        if item_type != 0x000D {
            lines.push(
                payload[item_start..item_end]
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes(chunk.try_into().unwrap()))
                    .collect(),
            );
        }
        offset = item_end;
    }
    lines
}

fn record_position(records: &[u16], typ: u16) -> usize {
    records
        .iter()
        .position(|&record| record == typ)
        .unwrap_or(usize::MAX)
}

fn utf16le_string(bytes: &[u8]) -> String {
    let units = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16_lossy(&units)
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

fn xls_unicode_string_no_cch_at(bytes: &[u8], offset: usize, len: usize) -> String {
    if offset >= bytes.len() {
        return String::new();
    }
    let flags = bytes[offset];
    let start = offset + 1;
    if flags & 0x01 != 0 {
        let end = start + len.saturating_mul(2);
        if end > bytes.len() {
            return String::new();
        }
        utf16le_string(&bytes[start..end])
    } else {
        let end = start + len;
        if end > bytes.len() {
            return String::new();
        }
        String::from_utf8_lossy(&bytes[start..end]).into_owned()
    }
}

#[test]
fn renamed_sheet_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "CustomName").expect("rename");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back");

    assert_eq!(parsed.sheet_count(), 1);
    assert_eq!(parsed.worksheet(0).unwrap().name(), "CustomName");
}

#[test]
fn multi_sheet_round_trips_with_all_names() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Alpha").expect("rename Sheet1");
    wb.add_worksheet_with_name("Beta").expect("add Beta");
    wb.add_worksheet_with_name("Gamma").expect("add Gamma");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back");

    assert_eq!(parsed.sheet_count(), 3);
    let names: Vec<_> = parsed.worksheets().map(|s| s.name().to_string()).collect();
    assert_eq!(names, vec!["Alpha", "Beta", "Gamma"]);
}

#[test]
fn writes_cfb_v3_envelope() {
    let wb = Workbook::new();
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    assert_eq!(
        &bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "CFB magic header (MS-CFB §2.2)"
    );
    assert_eq!(
        u16::from_le_bytes([bytes[26], bytes[27]]),
        0x0003,
        "major version must be 3 (512-byte sectors) for .xls"
    );
}

#[test]
fn special_and_unicode_sheet_names_round_trip() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First & Last")
        .expect("rename Sheet1");
    wb.add_worksheet_with_name("with 'apostrophe'")
        .expect("apostrophe sheet");
    wb.add_worksheet_with_name("日本語データ")
        .expect("unicode sheet");
    wb.add_worksheet_with_name("dash-dot.dot")
        .expect("dash-dot sheet");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read back");

    let names: Vec<_> = parsed.worksheets().map(|s| s.name().to_string()).collect();
    assert_eq!(
        names,
        vec![
            "First & Last",
            "with 'apostrophe'",
            "日本語データ",
            "dash-dot.dot",
        ]
    );
}

#[test]
fn write_to_bytes_then_read_file_round_trips() {
    let wb = Workbook::new();
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    let temp = tempfile::NamedTempFile::new().expect("temp file");
    std::fs::write(temp.path(), &bytes).expect("write to temp");
    let parsed = XlsReader::read_file(temp.path()).expect("read back from disk");

    assert_eq!(parsed.sheet_count(), 1);
}

/// Probe whether LibreOffice's loadenv accepts our skeleton output.
/// Useful for empirical viability checks during writer development.
/// `#[ignore]`-gated because it needs a running LO container.
#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_skeleton_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "ProbeSheet").expect("rename");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_skeleton_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the skeleton workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_numeric_grouped_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_numeric_grouped_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_grouped_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the numeric-grouped pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_date_grouped_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_date_grouped_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_date_grouped_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the date-grouped pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_multi_unit_date_grouped_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_multi_unit_date_grouped_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_multi_unit_date_grouped_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the multi-unit date-grouped pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_manual_grouped_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_manual_grouped_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_manual_grouped_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the manual-grouped pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_numeric_manual_grouped_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_manual_numeric_grouped_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_numeric_manual_grouped_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the numeric manual-grouped pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_all_aggregate_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_all_aggregate_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_all_aggregate_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the all-aggregate pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_bool_error_manual_grouped_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_manual_bool_error_grouped_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_bool_error_manual_grouped_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the bool/error manual-grouped pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_manual_column_grouped_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_manual_column_grouped_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_manual_column_grouped_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the manual column-grouped pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_manual_page_grouped_pivot_workbook() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    add_manual_page_grouped_pivot(&mut wb);
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_manual_page_grouped_pivot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<i32, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let count = wb
            .sheet_count()
            .await
            .map_err(|e| format!("sheet_count: {e}"))?;
        Ok(count)
    });
    let _ = std::fs::remove_file(&path);
    let count = outcome.expect("LO must open the manual page-grouped pivot workbook");
    assert_eq!(count, 1);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_read_cell_values_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Probe").expect("rename");
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("set A1");
    ws.set_cell_value("B2", -3.14).expect("set B2");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_cells_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write to shared dir");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<(f64, f64), String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let a1 = wb
            .get_cell_value("A1")
            .await
            .map_err(|e| format!("read A1: {e}"))?;
        let b2 = wb
            .get_cell_value("B2")
            .await
            .map_err(|e| format!("read B2: {e}"))?;
        Ok((a1, b2))
    });
    let _ = std::fs::remove_file(&path);
    let (a1, b2) = outcome.expect("LO must read cells we wrote");
    assert!((a1 - 42.0).abs() < 1e-9, "A1 = {a1}");
    assert!((b2 - -3.14).abs() < 1e-9, "B2 = {b2}");
}
