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

use duke_sheets_chart::{CellMarker, DrawingAnchor};
use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, Top10Filter};
use duke_sheets_core::conditional_format::ConditionalFormatRule;
use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::Color;
use duke_sheets_core::table::{Table, TableColumn};
use duke_sheets_core::validation::{DataValidation, ValidationOperator, ValidationType};
use duke_sheets_core::worksheet::{PageOrientation, SheetProtection, SheetVisibility};
use duke_sheets_core::PivotValuesAxis;
use duke_sheets_core::{
    CellAddress, CellError, CellRange, CellValue, Hyperlink, PivotAggregate, PivotCacheSourceKind,
    PivotDateGroupUnit, PivotDatePeriod, PivotField, PivotFieldRef, PivotFilter,
    PivotFilterOperator, PivotGrouping, PivotManualGroup, PivotMeasure, PivotShowAs, PivotSort,
    PivotSource, PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable, PivotValue, Workbook,
};
use duke_sheets_excel_com::{ChainStep, SheetRef};
use excel_com_protocol::ResponseData;
use serde_json::json;

use crate::{
    atp_all_formulas, cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm,
    push_file_to_vm, roundtrip_through_excel_xls, roundtrip_through_excel_xls_bytes,
    temp_fixture_xls, xls_externname_record_bodies, xls_formula_ptg_streams_for_compare,
};

fn range(start: &str, end: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start).unwrap(),
        CellAddress::parse(end).unwrap(),
    )
}

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

fn named_formula_workbook() -> Workbook {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();
    let data_values = [[5.0, 2.0, 3.0], [10.0, 20.0, 30.0], [15.0, 30.0, 45.0]];
    let data = wb.worksheet_mut(1).unwrap();
    for (row, values) in data_values.iter().enumerate() {
        for (col, value) in values.iter().enumerate() {
            data.set_cell_value_at(row as u32, col as u16, *value)
                .unwrap();
        }
    }

    wb.define_name("Numbers", "Data!$A$1:$A$3").unwrap();
    wb.define_name("TaxRate", "0.1").unwrap();
    wb.define_name("LeftBlock", "Data!$A$1:$B$3").unwrap();
    wb.define_name("RightBlock", "Data!$B$2:$C$3").unwrap();
    wb.define_name("TopCells", "Data!$A$1:$A$2").unwrap();
    wb.define_name("RightCells", "Data!$C$2:$C$3").unwrap();
    wb.define_name("StartCell", "Data!$A$1").unwrap();
    wb.define_name("EndCell", "Data!$A$3").unwrap();

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=SUM(Numbers)").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(30.0))
        .unwrap();
    calc.set_cell_formula("B2", "=B1*TaxRate").unwrap();
    calc.set_formula_result(1, 1, CellValue::Number(3.0))
        .unwrap();
    calc.set_cell_formula("B3", "=TaxRate*2").unwrap();
    calc.set_formula_result(2, 1, CellValue::Number(0.2))
        .unwrap();
    calc.set_cell_formula("B4", "=SUM(LeftBlock RightBlock)")
        .unwrap();
    calc.set_formula_result(3, 1, CellValue::Number(50.0))
        .unwrap();
    calc.set_cell_formula("B5", "=SUM((TopCells,RightCells))")
        .unwrap();
    calc.set_formula_result(4, 1, CellValue::Number(90.0))
        .unwrap();
    calc.set_cell_formula("B6", "=SUM(StartCell:EndCell)")
        .unwrap();
    calc.set_formula_result(5, 1, CellValue::Number(30.0))
        .unwrap();

    wb
}

fn rename_excel_sheet(
    excel: &duke_sheets_excel_com::ExcelBridge,
    workbook_handle: u64,
    index: u32,
    name: &str,
) {
    excel
        .set(
            workbook_handle,
            vec![SheetRef::Index(index).to_chain_step()],
            "Name",
            serde_json::Value::from(name),
        )
        .expect("rename Excel worksheet");
}

fn add_excel_worksheet_after(
    excel: &duke_sheets_excel_com::ExcelBridge,
    workbook_handle: u64,
    after_index: u32,
    name: &str,
) {
    let after_handle = excel
        .navigate(
            workbook_handle,
            vec![SheetRef::Index(after_index).to_chain_step()],
        )
        .expect("navigate worksheet for insert");
    let response = excel
        .invoke(
            workbook_handle,
            vec![ChainStep::Property("Worksheets".to_string())],
            "Add",
            vec![serde_json::Value::Null, json!({"$ref": after_handle})],
        )
        .expect("add Excel worksheet");
    let _ = excel.release(after_handle);
    let sheet_handle = match response {
        Some(ResponseData::Handle { handle }) => handle,
        other => panic!("expected worksheet handle from Add, got {other:?}"),
    };
    excel
        .set(sheet_handle, vec![], "Name", serde_json::Value::from(name))
        .expect("rename added worksheet");
    excel
        .release(sheet_handle)
        .expect("release added worksheet");
}

fn define_excel_name(
    excel: &duke_sheets_excel_com::ExcelBridge,
    workbook_handle: u64,
    name: &str,
    refers_to: &str,
) {
    let response = excel
        .invoke(
            workbook_handle,
            vec![ChainStep::Property("Names".to_string())],
            "Add",
            vec![
                serde_json::Value::from(name),
                serde_json::Value::from(refers_to),
            ],
        )
        .expect("define Excel name");
    if let Some(ResponseData::Handle { handle }) = response {
        excel.release(handle).expect("release name handle");
    }
}

fn excel_authored_named_formula_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let mut wb = excel.create_workbook().expect("create Excel workbook");
        rename_excel_sheet(&excel, wb.handle(), 0, "Calc");
        add_excel_worksheet_after(&excel, wb.handle(), 0, "Data");
        for (row, values) in [[5.0, 2.0, 3.0], [10.0, 20.0, 30.0], [15.0, 30.0, 45.0]]
            .iter()
            .enumerate()
        {
            for (col, value) in values.iter().enumerate() {
                let cell = format!("{}{}", (b'A' + col as u8) as char, row + 1);
                wb.set_cell_value_on_sheet(SheetRef::Name("Data".into()), &cell, *value)
                    .expect("set Excel data cell");
            }
        }

        for (name, refers_to) in [
            ("Numbers", "=Data!$A$1:$A$3"),
            ("TaxRate", "=0.1"),
            ("LeftBlock", "=Data!$A$1:$B$3"),
            ("RightBlock", "=Data!$B$2:$C$3"),
            ("TopCells", "=Data!$A$1:$A$2"),
            ("RightCells", "=Data!$C$2:$C$3"),
            ("StartCell", "=Data!$A$1"),
            ("EndCell", "=Data!$A$3"),
        ] {
            define_excel_name(&excel, wb.handle(), name, refers_to);
        }

        wb.set_active_sheet_name("Calc");
        for (cell, formula) in [
            ("B1", "=SUM(Numbers)"),
            ("B2", "=B1*TaxRate"),
            ("B3", "=TaxRate*2"),
            ("B4", "=SUM(LeftBlock RightBlock)"),
            ("B5", "=SUM((TopCells,RightCells))"),
            ("B6", "=SUM(StartCell:EndCell)"),
        ] {
            wb.set_cell_formula(cell, formula)
                .expect("set Excel formula");
        }
        excel.recalculate().expect("Excel recalculate");
        wb.save_as(&fixture.vm_path, 56).expect("Excel SaveAs xls");
        wb.close().expect("close Excel-authored workbook");
    }

    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

/// Formulas exercised by the function-arity byte parity test. Listed in
/// (cell, formula, cached_result) order; the same cell/order is used both by
/// our writer and the Excel-authored fixture so PTG streams line up
/// positionally.
const FUNCTION_ARITY_FORMULAS: &[(&str, &str, f64)] = &[
    // Fixed-arity → PtgFunc
    ("B1", "=SQRT(A1)", 2.0),                              // iftab=20
    ("B2", "=ABS(A2)", 3.0),                               // iftab=24
    ("B3", "=LEN(A3)", 5.0),                               // iftab=32
    ("B4", "=ROUND(A1,1)", 4.0),                           // iftab=27
    ("B5", "=PI()", std::f64::consts::PI),                 // iftab=19
    ("B7", "=SIN(A1)", -0.7568024953079282),               // iftab=15
    ("B8", "=COS(A1)", -0.6536436208636119),               // iftab=16
    ("B9", "=TAN(A1)", 1.1578212823495777),                // iftab=17
    ("B10", "=INT(A1)", 4.0),                              // iftab=25
    ("B11", "=SIGN(A2)", -1.0),                            // iftab=26
    ("B12", "=NOT(A4)", 1.0),                              // iftab=38
    ("B13", "=MOD(A1,3)", 1.0),                            // iftab=39
    ("B14", "=ATAN2(A1,A1)", std::f64::consts::FRAC_PI_4), // iftab=97
    ("B15", "=ISNA(A1)", 0.0),                             // iftab=2
    ("B16", "=ISERROR(A1)", 0.0),                          // iftab=3
    // Additional verified fixed-arity functions (see commit "Expand XLS
    // fixed-arity coverage" for the Excel-COM probe that confirmed each
    // one emits PtgFunc rather than PtgFuncVar).
    ("B21", "=POWER(A1,2)", 16.0),   // iftab=337
    ("B22", "=CHAR(65)", 0.0),       // iftab=111, cached not asserted exactly
    ("B23", "=YEAR(A5)", 2020.0),    // iftab=69
    ("B24", "=MONTH(A5)", 1.0),      // iftab=68
    ("B25", "=DAY(A5)", 1.0),        // iftab=67
    ("B26", "=DATE(2020,1,1)", 0.0), // iftab=65
    ("B27", "=DEGREES(A1)", 0.0),    // iftab=343
    ("B28", "=RADIANS(A1)", 0.0),    // iftab=342
    ("B29", "=ISBLANK(A1)", 0.0),    // iftab=129
    ("B30", "=ISTEXT(A3)", 1.0),     // iftab=127
    ("B31", "=ISNUMBER(A1)", 1.0),   // iftab=128
    ("B32", "=ISLOGICAL(A4)", 1.0),  // iftab=198
    ("B33", "=ISNONTEXT(A1)", 1.0),  // iftab=190
    ("B34", "=ISREF(A1)", 1.0),      // iftab=105 (R-class arg 0)
    ("B35", "=ISERR(A1)", 0.0),      // iftab=126
    ("B36", "=T(A3)", 0.0),          // iftab=130 (R-class arg 0)
    ("B37", "=N(A1)", 4.0),          // iftab=131 (R-class arg 0)
    ("B38", "=TYPE(A1)", 1.0),       // iftab=86
    ("B39", "=ERROR.TYPE(A1)", 0.0), // iftab=261
    ("B40", "=COUNTBLANK(A1)", 0.0), // iftab=347 (R-class arg 0)
    ("B41", "=FACT(5)", 120.0),      // iftab=184
    ("B42", "=CODE(A3)", 104.0),     // iftab=121
    ("B43", "=HOUR(A5)", 0.0),       // iftab=71
    ("B44", "=MINUTE(A5)", 0.0),     // iftab=72
    ("B45", "=SECOND(A5)", 0.0),     // iftab=73
    ("B46", "=TIME(1,2,3)", 0.0),    // iftab=66 (fixed 3-arg → PtgFunc)
    // Variable-arity → PtgFuncVar
    ("B6", "=ROW(A1)", 1.0),         // iftab=8 (var, ref arg)
    ("B17", "=SUM(A1:A2)", 1.0),     // iftab=4 (PtgAttrSum form, R-class operand)
    ("B18", "=AVERAGE(A1:A2)", 0.5), // iftab=5 (PtgFuncVar, R-class operand)
    ("B19", "=MIN(A1:A2)", -3.0),    // iftab=6
    ("B20", "=MAX(A1:A2)", 4.0),     // iftab=7
];

fn function_arity_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).unwrap();
    ws.set_cell_value("A2", -3.0).unwrap();
    ws.set_cell_value("A3", "hello").unwrap();
    ws.set_cell_value_at(3, 0, false).unwrap(); // A4 = FALSE for NOT(A4)
    ws.set_cell_value("A5", 43831.0).unwrap(); // A5 = serial for 2020-01-01
    for (cell, formula, expected) in FUNCTION_ARITY_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_function_arity_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 4.0).expect("set A1");
        wb.set_cell_value("A2", -3.0).expect("set A2");
        wb.set_cell_value("A3", "hello").expect("set A3");
        wb.set_cell_value("A4", false).expect("set A4");
        wb.set_cell_value("A5", 43831.0).expect("set A5");
        for (cell, formula, _) in FUNCTION_ARITY_FORMULAS {
            wb.set_cell_formula(cell, formula)
                .expect("set Excel formula");
        }
        excel.recalculate().expect("Excel recalculate");
        wb.save_as(&fixture.vm_path, 56).expect("Excel SaveAs xls");
        wb.close().expect("close Excel-authored workbook");
    }

    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

/// Volatile functions: Excel prefixes any FORMULA whose token stream
/// references one of these with a PtgAttrVolatile (0x19, flags=0x01) so the
/// recalc engine knows to re-evaluate on every change. Without the prefix
/// Excel still recalculates the cell but the saved bytes differ from
/// Excel's canonical output.
const VOLATILE_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=NOW()", 45000.0),   // iftab=74, 0 args
    ("B2", "=RAND()", 0.5),      // iftab=63, 0 args
    ("B3", "=TODAY()", 45000.0), // iftab=221, 0 args
    // OFFSET is volatile; result value isn't asserted byte-for-byte.
    ("B4", "=OFFSET(A1,0,0)", 4.0), // iftab=78, variable args, ref-class arg 0
    // INDIRECT is volatile.
    ("B5", "=INDIRECT(\"A1\")", 4.0), // iftab=148, variable args
                                      //
                                      // RANDBETWEEN (iftab=464) and INFO (iftab=244) — both fixed for runtime/
                                      // writer volatile-flag drift in the FunctionDef unification — are NOT in
                                      // this batch because Excel re-save behaviour is environment-dependent
                                      // (RANDBETWEEN was Analysis ToolPak before Excel 2010 and depending on
                                      // the bridge's Excel version may resolve to #N/A on recalc; INFO returns
                                      // host-OS-specific strings). Regression coverage for those bugs lives in
                                      // duke-sheets-xls's xls_formula_round_trip.rs as token-level unit tests.
];

fn volatile_function_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).unwrap();
    for (cell, formula, expected) in VOLATILE_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

/// Nested-function formulas exercising operand-class propagation. A
/// reference-class function (IF, CHOOSE) emitted in a reference argument
/// position takes R-class on its PtgFunc/PtgFuncVar token; in a value
/// position (top level, or inside a pure value function like ABS) it takes
/// V-class. Verified against Excel-authored bytes.
const NESTED_FUNCTION_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=IF(A1>0,IF(A1>10,1,2),3)", 2.0), // inner IF R-class (t-branch)
    ("B2", "=IF(A1>0,3,IF(A1>10,1,2))", 3.0), // inner IF R-class (f-branch)
    ("B3", "=SUM(ABS(A1))", 4.0),             // ABS stays V-class in SUM
    ("B4", "=SUM(IF(A1>0,A1,A2))", 4.0),      // IF R-class in SUM
    ("B5", "=CHOOSE(1,IF(A2>0,1,2),3)", 2.0), // IF R-class as CHOOSE choice (sel=1 → IF; A2<0 → 2)
    ("B6", "=IF(A1>0,CHOOSE(2,10,20),3)", 20.0), // CHOOSE R-class in IF t-branch
    ("B7", "=ABS(IF(A1>0,A1,A2))", 4.0),      // IF V-class in ABS
    ("B8", "=SUM(OFFSET(A1,0,0))", 4.0),      // OFFSET R-class in SUM (volatile)
    ("B9", "=SUM(INDIRECT(\"A1\"))", 4.0),    // INDIRECT R-class in SUM (volatile)
    ("B10", "=OFFSET(A1,0,0)", 4.0),          // OFFSET V-class at top level
    ("B11", "=SUM(INDEX(A1:A3,1))", 4.0),     // INDEX R-class in SUM, arg0 R-class
    ("B12", "=INDEX(A1:A3,1)", 4.0),          // INDEX V-class top level, arg0 R-class
];

fn nested_function_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).unwrap();
    ws.set_cell_value("A2", -3.0).unwrap();
    ws.set_cell_value("A3", 6.0).unwrap();
    for (cell, formula, expected) in NESTED_FUNCTION_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_nested_function_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 4.0).expect("set A1");
        wb.set_cell_value("A2", -3.0).expect("set A2");
        wb.set_cell_value("A3", 6.0).expect("set A3");
        for (cell, formula, _) in NESTED_FUNCTION_FORMULAS {
            wb.set_cell_formula(cell, formula)
                .expect("set Excel formula");
        }
        excel.recalculate().expect("Excel recalculate");
        wb.save_as(&fixture.vm_path, 56).expect("Excel SaveAs xls");
        wb.close().expect("close Excel-authored workbook");
    }

    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

/// CHOOSE-optimization formulas: Excel emits CHOOSE using PtgAttrChoose
/// (MS-XLS §2.5.198.40) with a u16 jump table — one entry per choice plus
/// a final exit entry — followed by each choice and a PtgAttrSkip. Matching
/// Excel's emission byte-for-byte requires the same layout.
const CHOOSE_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=CHOOSE(A1,10,20,30)", 20.0),
    ("B2", "=CHOOSE(2,100,200)", 200.0),
    ("B3", "=CHOOSE(A1,A1,A1*2,A1*3)", 8.0),
    ("B4", "=CHOOSE(A1,1,2,3,4,5)", 2.0),
];

fn choose_optimization_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    for (cell, formula, expected) in CHOOSE_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_choose_optimization_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 2.0).expect("set A1");
        for (cell, formula, _) in CHOOSE_FORMULAS {
            wb.set_cell_formula(cell, formula)
                .expect("set Excel formula");
        }
        excel.recalculate().expect("Excel recalculate");
        wb.save_as(&fixture.vm_path, 56).expect("Excel SaveAs xls");
        wb.close().expect("close Excel-authored workbook");
    }

    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

/// IF-optimization formulas: Excel emits `IF` using PtgAttrIf + PtgAttrGoto
/// (MS-XLS §2.5.198.39 / §2.5.198.37) so only one branch is evaluated at
/// runtime. Covers both the 3-arg and 2-arg forms across literal-only,
/// reference-bearing, and arithmetic branches.
const IF_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=IF(A1>0,1,2)", 1.0),
    ("B2", "=IF(A1>0,A1,0)", 4.0),
    ("B3", "=IF(A1>0,A1*2,A1)", 8.0),
    ("B4", "=IF(A1<0,A1)", 0.0),
    ("B5", "=IF(A1>0,A1*2)", 8.0),
];

fn if_optimization_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).unwrap();
    for (cell, formula, expected) in IF_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_if_optimization_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 4.0).expect("set A1");
        for (cell, formula, _) in IF_FORMULAS {
            wb.set_cell_formula(cell, formula)
                .expect("set Excel formula");
        }
        excel.recalculate().expect("Excel recalculate");
        wb.save_as(&fixture.vm_path, 56).expect("Excel SaveAs xls");
        wb.close().expect("close Excel-authored workbook");
    }

    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

fn excel_authored_volatile_function_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 4.0).expect("set A1");
        for (cell, formula, _) in VOLATILE_FORMULAS {
            wb.set_cell_formula(cell, formula)
                .expect("set Excel formula");
        }
        excel.recalculate().expect("Excel recalculate");
        wb.save_as(&fixture.vm_path, 56).expect("Excel SaveAs xls");
        wb.close().expect("close Excel-authored workbook");
    }

    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

/// Unary-plus and parenthesis preservation: Excel keeps `=+A1` as PtgUplus
/// (0x12) and every paren pair as a postfix PtgParen (0x15), including
/// redundant and nested parens. Verified byte-for-byte.
const UPLUS_PAREN_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=+A1", 2.0),
    ("B2", "=(A1+A2)*2", 10.0),
    ("B3", "=(A1)", 2.0),
    ("B4", "=A1+(A2)", 5.0),
    ("B5", "=-(A1)", -2.0),
    ("B6", "=2*(A1+A2)", 10.0),
    ("B7", "=((A1))", 2.0),
    // Unary operand classes: refs under unary value operators inside
    // an R-forced argument position.
    ("B8", "=SUM(-A1)", -2.0),
    ("B9", "=-A1+A2", 1.0),
    ("B10", "=SUM(A1%)", 0.02),
];

fn uplus_paren_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_value("A2", 3.0).unwrap();
    for (cell, formula, expected) in UPLUS_PAREN_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_uplus_paren_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 2.0).unwrap();
        wb.set_cell_value("A2", 3.0).unwrap();
        for (cell, formula, _) in UPLUS_PAREN_FORMULAS {
            wb.set_cell_formula(cell, formula).unwrap();
        }
        excel.recalculate().unwrap();
        wb.save_as(&fixture.vm_path, 56).unwrap();
        wb.close().unwrap();
    }
    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

#[test]
fn excel_byte_parity_for_uplus_paren_we_emit() {
    let wb = uplus_paren_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        UPLUS_PAREN_FORMULAS.len(),
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let resave_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLS unary-plus/paren token streams on re-save"
    );
    let authored_ptgs =
        xls_formula_ptg_streams_for_compare(&excel_authored_uplus_paren_xls_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLS unary-plus/paren token streams differ from Excel-authored output"
    );
}

/// Array constants: Excel emits {...} as a PtgArray token (rgce) plus the
/// element data in the rgcb. Verified byte-for-byte (the PtgArray reserved
/// bytes and PtgAttrSum unused bytes are normalized).
const ARRAY_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=SUM({1,2,3})", 6.0),
    ("B2", "=SUM({1,2;3,4})", 10.0),
    ("B3", "=SUM({10,20})", 30.0),
];

fn array_constant_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (cell, formula, expected) in ARRAY_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_array_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        for (cell, formula, _) in ARRAY_FORMULAS {
            wb.set_cell_formula(cell, formula).unwrap();
        }
        excel.recalculate().unwrap();
        wb.save_as(&fixture.vm_path, 56).unwrap();
        wb.close().unwrap();
    }
    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

#[test]
fn excel_byte_parity_for_array_constants_we_emit() {
    let wb = array_constant_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        ARRAY_FORMULAS.len(),
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let resave_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLS array-constant token streams on re-save"
    );
    let authored_ptgs = xls_formula_ptg_streams_for_compare(&excel_authored_array_xls_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLS array-constant token streams differ from Excel-authored output"
    );
}

#[test]
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

    let (result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    assert_eq!(
        xls_formula_ptg_streams_for_compare(&writer_bytes),
        xls_formula_ptg_streams_for_compare(&excel_bytes),
        "Excel canonicalized our XLS cross-sheet formula token streams"
    );
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
    // Verify formula structure survives — without this, Excel could
    // have inlined the cross-sheet ref (e.g. "=10") and the cached
    // value would still match.
    let f1 = s.get_formula_at(0, 1).expect("B1 still a formula");
    assert!(
        f1.contains("Data") && f1.contains("A1"),
        "cross-sheet ref lost from B1 formula: {f1:?}"
    );
    let f2 = s.get_formula_at(1, 1).expect("B2 still a formula");
    assert!(
        f2.contains("SUM") && f2.contains("Data") && f2.contains("A1:A3"),
        "cross-sheet SUM lost from B2 formula: {f2:?}"
    );
}

#[test]
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
fn excel_can_evaluate_named_range_formulas_we_emit() {
    let wb = named_formula_workbook();

    let (result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        6, // named_formula_workbook formulas B1..B6
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let excel_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, excel_ptgs,
        "Excel canonicalized our XLS formula token streams"
    );
    assert_eq!(
        writer_ptgs,
        xls_formula_ptg_streams_for_compare(&excel_authored_named_formula_xls_bytes()),
        "our XLS formula token streams differ from Excel-authored output"
    );

    let s = result.worksheet_by_name("Calc").unwrap();
    for (row, expected) in [
        (0, 30.0),
        (1, 3.0),
        (2, 0.2),
        (3, 50.0),
        (4, 90.0),
        (5, 30.0),
    ] {
        match s.get_value_at(row, 1).effective_value() {
            CellValue::Number(n) => assert!((n - expected).abs() < 1e-9, "B{} = {n}", row + 1),
            other => panic!("B{} expected Number({expected}), got {other:?}", row + 1),
        }
    }
    // Verify named ranges survive in the formula text. The
    // workbook-level `workbook.named_ranges()` map is documented as
    // NOT being repopulated by the XLS reader (FEATURES.md rows
    // 202-205 with R●/W● for XLS), so we only check the formula
    // text here, pinned by xls_formula_round_trip::
    // named_range_in_formula_text_survives_xls_roundtrip in-process.
    let f1 = s.get_formula_at(0, 1).expect("B1 still a formula");
    assert!(
        f1.contains("Numbers"),
        "named range Numbers lost from B1: {f1:?}"
    );
    let f2 = s.get_formula_at(1, 1).expect("B2 still a formula");
    assert!(
        f2.contains("TaxRate"),
        "named range TaxRate lost from B2: {f2:?}"
    );
    for (row, expected_names) in [
        (2, &["TaxRate"][..]),
        (3, &["LeftBlock", "RightBlock"][..]),
        (4, &["TopCells", "RightCells"][..]),
        (5, &["StartCell", "EndCell"][..]),
    ] {
        let formula = s
            .get_formula_at(row, 1)
            .unwrap_or_else(|| panic!("B{} still a formula", row + 1));
        for expected_name in expected_names {
            assert!(
                formula.contains(expected_name),
                "named range {expected_name} lost from B{}: {formula:?}",
                row + 1
            );
        }
    }
}

#[test]
fn excel_byte_parity_for_function_arity_we_emit() {
    let wb = function_arity_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        FUNCTION_ARITY_FORMULAS.len(),
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let resave_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLS function formula token streams on re-save"
    );
    let authored_ptgs =
        xls_formula_ptg_streams_for_compare(&excel_authored_function_arity_xls_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLS function formula token streams differ from Excel-authored output"
    );
}

#[test]
fn excel_byte_parity_for_nested_functions_we_emit() {
    let wb = nested_function_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        NESTED_FUNCTION_FORMULAS.len(),
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let resave_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLS nested-function token streams on re-save"
    );
    let authored_ptgs =
        xls_formula_ptg_streams_for_compare(&excel_authored_nested_function_xls_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLS nested-function token streams differ from Excel-authored output"
    );
}

#[test]
fn excel_byte_parity_for_choose_optimization_we_emit() {
    let wb = choose_optimization_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        CHOOSE_FORMULAS.len(),
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let resave_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLS CHOOSE formula token streams on re-save"
    );
    let authored_ptgs =
        xls_formula_ptg_streams_for_compare(&excel_authored_choose_optimization_xls_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLS CHOOSE formula token streams differ from Excel-authored output"
    );
}

#[test]
fn excel_byte_parity_for_if_optimization_we_emit() {
    let wb = if_optimization_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        IF_FORMULAS.len(),
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let resave_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLS IF formula token streams on re-save"
    );
    let authored_ptgs =
        xls_formula_ptg_streams_for_compare(&excel_authored_if_optimization_xls_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLS IF formula token streams differ from Excel-authored output"
    );
}

#[test]
fn excel_byte_parity_for_volatile_functions_we_emit() {
    let wb = volatile_function_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);
    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        VOLATILE_FORMULAS.len(),
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let resave_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLS volatile formula token streams on re-save"
    );
    let authored_ptgs =
        xls_formula_ptg_streams_for_compare(&excel_authored_volatile_function_xls_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLS volatile formula token streams differ from Excel-authored output"
    );
}

#[test]
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

    // repeat_rows must survive — set on input but previously not
    // asserted. A regression dropping repeat_rows would have gone
    // unnoticed.
    let repeat_rows = s
        .page_setup()
        .repeat_rows
        .expect("repeat_rows must survive Excel round-trip");
    assert_eq!(
        repeat_rows,
        (0, 0),
        "repeat_rows lost or mangled: {:?}",
        repeat_rows
    );
}

#[test]
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

    // Page setup fields were set on input but previously not asserted.
    // A regression dropping landscape orientation, gridline printing,
    // header text, or left_margin would have gone unnoticed.
    let ps = s.page_setup();
    assert_eq!(
        ps.orientation,
        PageOrientation::Landscape,
        "landscape orientation lost"
    );
    assert!(ps.print_gridlines, "print_gridlines flag lost");
    let header = ps
        .odd_header
        .as_deref()
        .expect("odd_header text lost after round-trip");
    assert!(
        header.contains("Hdr"),
        "odd_header content mangled: {header:?}"
    );
    assert!(
        (ps.left_margin - 0.5).abs() < 0.05,
        "left_margin lost or mangled: {} (expected ~0.5)",
        ps.left_margin
    );
}

#[test]
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

    // Excel rewrites the password hash with its own algorithm. We
    // can't require byte-identity but we can require something
    // non-zero — a writer that dropped the password entirely would
    // produce password_hash=None or Some(0).
    let hash = prot
        .password_hash
        .expect("password_hash lost after round-trip");
    assert_ne!(hash, 0, "password_hash dropped to 0 — writer regression");
}

#[test]
fn excel_preserves_workbook_protection_and_protected_ranges_we_emit() {
    use duke_sheets_core::{ProtectedRange, WorkbookProtection};

    let mut wb = Workbook::new();
    wb.set_workbook_protection(Some(WorkbookProtection {
        structure: true,
        windows: true,
        password_hash: Some(0xCAFE),
    }));

    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "editable").unwrap();
    ws.set_protected_ranges(vec![ProtectedRange {
        name: "Editable".to_string(),
        ranges: vec![range("A1", "B2"), range("D4", "D5")],
        password_hash: Some(0xCAFE),
        security_descriptor: None,
    }]);

    let result = roundtrip_through_excel_xls(&wb);
    let protection = result
        .workbook_protection()
        .expect("workbook protection must survive Excel round-trip");
    assert!(protection.structure, "workbook structure protection lost");
    assert!(protection.windows, "workbook window protection lost");
    assert!(
        protection.password_hash.unwrap_or_default() != 0,
        "workbook password hash lost"
    );

    let ranges = result.worksheet(0).unwrap().protected_ranges();
    assert_eq!(ranges.len(), 1, "protected range count changed");
    assert_eq!(ranges[0].name, "Editable");
    assert_eq!(ranges[0].ranges[0].to_string(), "A1:B2");
    assert_eq!(ranges[0].ranges[1].to_string(), "D4:D5");
    assert!(
        ranges[0].password_hash.unwrap_or_default() != 0,
        "protected range password hash lost"
    );
}

#[test]
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
    let formula = s.get_formula_at(0, 4).expect("E1 must still be a formula");
    assert!(
        formula.contains("A1:B3") && formula.contains("B2:C3"),
        "intersection ranges lost from formula: {formula:?}"
    );
}

#[test]
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
    let formula = s.get_formula_at(0, 4).expect("E1 must still be a formula");
    assert!(
        formula.contains("A1:A2") && formula.contains("C2:C3"),
        "union ranges lost from formula: {formula:?}"
    );
}

/// Excel must accept our XLS comment-shape emit (MSODRAWINGGROUP +
/// MSODRAWING + OBJ + TXO + NOTE) and round-trip the comment text +
/// author + anchor cell through SaveAs. The roundtrip helper asserts
/// no `Repaired` warning fires, which is the canonical signal that
/// our Escher record output is spec-compliant.
#[test]
fn excel_can_read_comment_we_emit() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.set_comment_at(0, 0, CellComment::new("Alice", "Hello from duke-sheets")).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    let c = s
        .comment_at(0, 0)
        .expect("comment must survive Excel re-save");
    assert!(
        c.plain_text().contains("Hello from duke-sheets"),
        "comment text lost after Excel round-trip: {:?}",
        c.text
    );
    assert!(
        c.author.contains("Alice"),
        "comment author lost after Excel round-trip: {:?}",
        c.author
    );
}

/// Excel parity for rich comment text: our TXO formatting runs must
/// survive Excel's re-save with their run boundary and bold flag.
// features: Rich text in comments
#[test]
fn excel_preserves_rich_comment_runs_we_emit_xls() {
    use duke_sheets_core::rich_text::{RichTextRun, RunFont};
    use duke_sheets_core::{CellComment, DrawingText};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.set_comment_at(
        0,
        0,
        CellComment {
            author: "Reviewer".to_string(),
            text: DrawingText {
                runs: vec![
                    RichTextRun {
                        text: "Bold lead".to_string(),
                        font: Some(RunFont {
                            bold: Some(true),
                            ..RunFont::default()
                        }),
                    },
                    RichTextRun {
                        text: " then plain".to_string(),
                        font: None,
                    },
                ],
                ..DrawingText::default()
            },
        },
    )
    .unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let comment = result
        .worksheet(0)
        .unwrap()
        .comment_at(0, 0)
        .expect("comment must survive Excel re-save");
    assert_eq!(comment.plain_text(), "Bold lead then plain");
    let bold_run = comment
        .text
        .runs
        .iter()
        .find(|run| run.text.contains("Bold lead"))
        .expect("bold run boundary survives Excel re-save");
    assert_eq!(
        bold_run.font.as_ref().and_then(|font| font.bold),
        Some(true),
        "bold formatting lost: {:?}",
        comment.text.runs
    );
}

/// Multi-comment scenario: Excel must accept multiple SP_CONTAINERs
/// inside one DG_CONTAINER and preserve each comment's text +
/// author + anchor cell. Catches off-by-one bugs in shape ID
/// allocation, OBJ.ftCmo.id, and NOTE.objId linking.
#[test]
fn excel_can_read_multiple_comments_we_emit() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "first").unwrap();
    ws.set_comment_at(0, 0, CellComment::new("Alice", "Comment one")).unwrap();
    ws.set_cell_value("C3", "second").unwrap();
    ws.set_comment_at(2, 2, CellComment::new("Bob", "Comment two body")).unwrap();
    ws.set_cell_value("E5", "third").unwrap();
    ws.set_comment_at(4, 4, CellComment::new("Carol", "Comment three")).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.comment_count(), 3, "all three comments must survive");

    // Verify each comment's text AND author land on the correct cell.
    // Without per-author assertion, the test would silently pass if
    // Excel scrambled the author–text mapping during NOTE re-save.
    let c1 = s.comment_at(0, 0).unwrap();
    assert!(c1.plain_text().contains("Comment one"), "A1 text: {:?}", c1.plain_text());
    assert!(c1.author.contains("Alice"), "A1 author: {:?}", c1.author);

    let c2 = s.comment_at(2, 2).unwrap();
    assert!(
        c2.plain_text().contains("Comment two body"),
        "C3 text: {:?}",
        c2.text
    );
    assert!(c2.author.contains("Bob"), "C3 author: {:?}", c2.author);

    let c3 = s.comment_at(4, 4).unwrap();
    assert!(c3.plain_text().contains("Comment three"), "E5 text: {:?}", c3.plain_text());
    assert!(c3.author.contains("Carol"), "E5 author: {:?}", c3.author);
}

/// Unicode comment text + author — drives the writer onto the
/// UTF-16LE TXO CONTINUE and NOTE author paths. Confirms Excel
/// preserves non-Latin glyphs through both the text and the author
/// field during the round-trip.
#[test]
fn excel_can_read_unicode_comment_we_emit() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "jp").unwrap();
    ws.set_comment_at(0, 0, CellComment::new("作者", "こんにちは")).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let c = result.worksheet(0).unwrap().comment_at(0, 0).unwrap();
    assert!(
        c.plain_text().contains("こんにちは"),
        "Japanese text lost: {:?}",
        c.text
    );
    assert!(
        c.author.contains("作者"),
        "Japanese author lost: {:?}",
        c.author
    );
}

/// Mixed picture + comment on the same sheet. The XLS writer
/// places ALL picture SP_CONTAINERs into the first MSODRAWING then
/// starts the first comment's SP_CONTAINER in the same record (up
/// through ClientData), then interleaves comment OBJ / TXO /
/// CONTINUE with continuation MSODRAWING records carrying the
/// ClientTextbox markers. This is an unusual layout — Excel might
/// reject the BIFF stream ordering. The parity test catches that.
#[test]
fn excel_can_read_xls_picture_and_comment_on_same_sheet_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
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
        },
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 1,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 4,
                col_offset_emu: 0,
                row: 5,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();
    ws.set_comment_at(
        7,
        5,
        CellComment::new("Alice", "Mixed-with-picture comment"),
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.image_count(), 1, "picture must survive Excel re-save");
    assert_eq!(s.comment_count(), 1, "comment must survive Excel re-save");
    let img = s.images().next().unwrap().payload;
    assert_eq!(img.format, ImageFormat::Png);
    assert_eq!(img.data, TEST_PNG_1X1, "PNG bytes preserved");
    let c = s.comment_at(7, 5).expect("comment at G8 must exist");
    assert!(
        c.plain_text().contains("Mixed-with-picture comment"),
        "comment text lost: {:?}",
        c.text
    );
}

/// Multi-sheet pictures: exercises the per-drawing 1024-aligned
/// shape-ID cluster allocation when pictures (not comments) live on
/// non-contiguous sheets. Two separate FDGG cluster entries are
/// emitted; Excel must accept the multi-cluster layout and both
/// pictures must survive Excel's SaveAs.
#[test]
fn excel_can_read_xls_pictures_across_multiple_sheets_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
    use duke_sheets_core::DrawingObject;

    fn pic(name: &str, col: u16, row: u32) -> DrawingObject {
        DrawingObject::image(EmbeddedImage {
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
        })
        .with_anchor(DrawingAnchor::TwoCell {
            from: CellMarker {
                col,
                col_offset_emu: 0,
                row,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: col + 2,
                col_offset_emu: 0,
                row: row + 2,
                row_offset_emu: 0,
            },
            edit_as: None,
        })
        .with_name(name)
    }

    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Alpha").unwrap();
    wb.add_worksheet_with_name("Beta").unwrap();
    wb.add_worksheet_with_name("Gamma").unwrap();

    // Picture on Alpha and Gamma; Beta is empty. Forces two FDGG
    // cluster entries with non-contiguous drawing IDs.
    wb.worksheet_mut(0)
        .unwrap()
        .add_drawing(pic("Pic on Alpha", 1, 1)).unwrap();
    wb.worksheet_mut(2)
        .unwrap()
        .add_drawing(pic("Pic on Gamma", 3, 3)).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    assert_eq!(
        result.worksheet(0).unwrap().image_count(),
        1,
        "Alpha's picture must survive"
    );
    assert_eq!(
        result.worksheet(1).unwrap().image_count(),
        0,
        "Beta must stay empty"
    );
    assert_eq!(
        result.worksheet(2).unwrap().image_count(),
        1,
        "Gamma's picture must survive"
    );
    assert_eq!(
        result
            .worksheet(0)
            .unwrap()
            .images()
            .next()
            .unwrap()
            .payload
            .data,
        TEST_PNG_1X1,
        "Alpha PNG bytes preserved"
    );
    assert_eq!(
        result
            .worksheet(2)
            .unwrap()
            .images()
            .next()
            .unwrap()
            .payload
            .data,
        TEST_PNG_1X1,
        "Gamma PNG bytes preserved"
    );
}

/// Multi-sheet comments: exercises the per-drawing 1024-aligned
/// shape-ID cluster allocation. A workbook with comments on
/// non-contiguous sheets must produce one FDGG cluster entry per
/// drawing and unique shape IDs across the workbook.
#[test]
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
        .set_comment_at(0, 0, CellComment::new("Alice", "Alpha A1")).unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_comment_at(3, 3, CellComment::new("Alice", "Alpha D4")).unwrap();
    wb.worksheet_mut(2)
        .unwrap()
        .set_comment_at(5, 5, CellComment::new("Carol", "Gamma F6")).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    assert_eq!(result.worksheet(0).unwrap().comment_count(), 2);
    assert_eq!(result.worksheet(1).unwrap().comment_count(), 0);
    assert_eq!(result.worksheet(2).unwrap().comment_count(), 1);

    // Verify text + author per comment so a mis-mapping across
    // sheets would be caught.
    let alpha_a1 = result.worksheet(0).unwrap().comment_at(0, 0).unwrap();
    assert!(
        alpha_a1.plain_text().contains("Alpha A1"),
        "Alpha A1 text: {:?}",
        alpha_a1.text
    );
    assert!(
        alpha_a1.author.contains("Alice"),
        "Alpha A1 author: {:?}",
        alpha_a1.author
    );

    let alpha_d4 = result.worksheet(0).unwrap().comment_at(3, 3).unwrap();
    assert!(
        alpha_d4.plain_text().contains("Alpha D4"),
        "Alpha D4 text: {:?}",
        alpha_d4.text
    );
    assert!(
        alpha_d4.author.contains("Alice"),
        "Alpha D4 author: {:?}",
        alpha_d4.author
    );

    let gamma_f6 = result.worksheet(2).unwrap().comment_at(5, 5).unwrap();
    assert!(
        gamma_f6.plain_text().contains("Gamma F6"),
        "Gamma F6 text: {:?}",
        gamma_f6.text
    );
    assert!(
        gamma_f6.author.contains("Carol"),
        "Gamma F6 author: {:?}",
        gamma_f6.author
    );
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
fn excel_can_read_xls_png_image_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
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
        },
        DrawingAnchor::TwoCell {
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
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 1, "image must survive Excel re-save");
    let img = images[0].payload;
    assert_eq!(img.format, ImageFormat::Png);
    assert_eq!(
        img.data, TEST_PNG_1X1,
        "PNG bytes must round-trip through Excel verbatim"
    );
    match &images[0].object.unwrap().anchor {
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

    // Width/height in EMU should be synthesised on read from the
    // anchor cell range × default cell sizes. We don't pin exact
    // values (Excel may adjust the anchor), only that they are
    // non-zero — proving the writer-side anchor encoding survives
    // and the reader correctly computes dimensions.
    assert!(
        img.width_emu > 0,
        "width_emu must be non-zero after Excel round-trip; got {}",
        img.width_emu
    );
    assert!(
        img.height_emu > 0,
        "height_emu must be non-zero after Excel round-trip; got {}",
        img.height_emu
    );
}

/// 632-byte 1x1 RGB JPEG at quality 50 (PIL-generated). Same
/// payload as the in-process JPEG round-trip test.
const TEST_JPEG_1X1: &[u8] = &[
    0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01,
    0x00, 0x01, 0x00, 0x00, 0xFF, 0xDB, 0x00, 0x43, 0x00, 0x10, 0x0B, 0x0C, 0x0E, 0x0C, 0x0A, 0x10,
    0x0E, 0x0D, 0x0E, 0x12, 0x11, 0x10, 0x13, 0x18, 0x28, 0x1A, 0x18, 0x16, 0x16, 0x18, 0x31, 0x23,
    0x25, 0x1D, 0x28, 0x3A, 0x33, 0x3D, 0x3C, 0x39, 0x33, 0x38, 0x37, 0x40, 0x48, 0x5C, 0x4E, 0x40,
    0x44, 0x57, 0x45, 0x37, 0x38, 0x50, 0x6D, 0x51, 0x57, 0x5F, 0x62, 0x67, 0x68, 0x67, 0x3E, 0x4D,
    0x71, 0x79, 0x70, 0x64, 0x78, 0x5C, 0x65, 0x67, 0x63, 0xFF, 0xDB, 0x00, 0x43, 0x01, 0x11, 0x12,
    0x12, 0x18, 0x15, 0x18, 0x2F, 0x1A, 0x1A, 0x2F, 0x63, 0x42, 0x38, 0x42, 0x63, 0x63, 0x63, 0x63,
    0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63,
    0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63,
    0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0x63, 0xFF, 0xC0,
    0x00, 0x11, 0x08, 0x00, 0x01, 0x00, 0x01, 0x03, 0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11,
    0x01, 0xFF, 0xC4, 0x00, 0x1F, 0x00, 0x00, 0x01, 0x05, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09,
    0x0A, 0x0B, 0xFF, 0xC4, 0x00, 0xB5, 0x10, 0x00, 0x02, 0x01, 0x03, 0x03, 0x02, 0x04, 0x03, 0x05,
    0x05, 0x04, 0x04, 0x00, 0x00, 0x01, 0x7D, 0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21,
    0x31, 0x41, 0x06, 0x13, 0x51, 0x61, 0x07, 0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xA1, 0x08, 0x23,
    0x42, 0xB1, 0xC1, 0x15, 0x52, 0xD1, 0xF0, 0x24, 0x33, 0x62, 0x72, 0x82, 0x09, 0x0A, 0x16, 0x17,
    0x18, 0x19, 0x1A, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2A, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A,
    0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A,
    0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A,
    0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99,
    0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7,
    0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5,
    0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE1, 0xE2, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF1,
    0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xF9, 0xFA, 0xFF, 0xC4, 0x00, 0x1F, 0x01, 0x00, 0x03,
    0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
    0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0xFF, 0xC4, 0x00, 0xB5, 0x11, 0x00,
    0x02, 0x01, 0x02, 0x04, 0x04, 0x03, 0x04, 0x07, 0x05, 0x04, 0x04, 0x00, 0x01, 0x02, 0x77, 0x00,
    0x01, 0x02, 0x03, 0x11, 0x04, 0x05, 0x21, 0x31, 0x06, 0x12, 0x41, 0x51, 0x07, 0x61, 0x71, 0x13,
    0x22, 0x32, 0x81, 0x08, 0x14, 0x42, 0x91, 0xA1, 0xB1, 0xC1, 0x09, 0x23, 0x33, 0x52, 0xF0, 0x15,
    0x62, 0x72, 0xD1, 0x0A, 0x16, 0x24, 0x34, 0xE1, 0x25, 0xF1, 0x17, 0x18, 0x19, 0x1A, 0x26, 0x27,
    0x28, 0x29, 0x2A, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49,
    0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69,
    0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88,
    0x89, 0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6,
    0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4,
    0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE2,
    0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xF9,
    0xFA, 0xFF, 0xDA, 0x00, 0x0C, 0x03, 0x01, 0x00, 0x02, 0x11, 0x03, 0x11, 0x00, 0x3F, 0x00, 0xC5,
    0xA2, 0x8A, 0x2B, 0xCB, 0x3E, 0xF0, 0xFF, 0xD9,
];

/// Picture rotation and flip flags: writer must encode rotation in
/// the FOPT `0x0004` property (FixedPoint 16.16 degrees) and flip H/V
/// in the FSP grfPersistence flag bits. Excel keeps non-quarter-turn
/// rotations (and their flips) as live transform properties; the
/// rotation and flip must round-trip.
///
/// Excel rasterizes exact 90-degree-multiple rotations and pure flips
/// into the bitmap when saving XLS (the transform properties are
/// dropped and the blip is re-encoded), so a second workbook pins
/// that boundary: the picture survives, its transform reads back as
/// identity.
#[test]
fn excel_can_read_xls_picture_rotation_and_flip_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let image = |rotation: Option<i32>, flip_h: bool| EmbeddedImage {
        format: ImageFormat::Png,
        media_path: String::new(),
        svg_media_path: None,
        width_emu: 1_000_000,
        height_emu: 1_000_000,
        rotation,
        flip_h,
        flip_v: false,
        data: TEST_PNG_1X1.to_vec(),
        svg_data: None,
    };
    let anchor = || DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            col_offset_emu: 0,
            row: 1,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: 4,
            col_offset_emu: 0,
            row: 5,
            row_offset_emu: 0,
        },
        edit_as: None,
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "rotated").unwrap();
    // 45 degrees clockwise: outside Excel's rasterization cases.
    ws.add_image(image(Some(2_700_000), true), anchor()).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 1, "rotated picture must survive Excel");
    let img = images[0].payload;
    assert_eq!(
        img.rotation,
        Some(2_700_000),
        "rotation must round-trip through Excel"
    );
    assert!(img.flip_h, "flip_h must round-trip through Excel");
    assert!(!img.flip_v, "flip_v=false must round-trip through Excel");

    // Boundary: a quarter-turn rotation gets baked into the pixels.
    // The picture must still read back (Excel emits a header-only
    // placeholder FBSE, which must keep its blip-id slot).
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "quarter-turn").unwrap();
    ws.add_image(image(Some(5_400_000), false), anchor()).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 1, "rasterized picture must survive Excel");
    let img = images[0].payload;
    assert_eq!(
        img.rotation, None,
        "Excel bakes quarter-turn rotations into the bitmap on XLS save"
    );
    assert!(!img.flip_h);
}

/// OneCell anchor variant: input has only a `from` cell + width/height
/// in EMU. Writer encodes via ClientAnchor flag=2 (move only). Excel
/// must accept the file and the picture's visual area must survive
/// the round-trip.
#[test]
fn excel_can_read_xls_onecell_image_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    const COL_EMU: i64 = 609_600;
    const ROW_EMU: i64 = 190_500;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "onecell-anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 2 * COL_EMU,
            height_emu: 3 * ROW_EMU,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        },
        DrawingAnchor::OneCell {
            from: CellMarker {
                col: 2,
                col_offset_emu: 0,
                row: 3,
                row_offset_emu: 0,
            },
            width_emu: 2 * COL_EMU,
            height_emu: 3 * ROW_EMU,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(
        images.len(),
        1,
        "OneCell picture must survive Excel re-save"
    );
    let img = images[0].payload;
    assert_eq!(img.format, ImageFormat::Png);
    match &images[0].object.unwrap().anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            // OneCell at (col=2, row=3) + 2 cols × 3 rows of default
            // cells means the picture spans columns 2..4 and rows
            // 3..6 inclusive at default sizes.
            assert_eq!(from.col, 2, "from col preserved");
            assert_eq!(from.row, 3, "from row preserved");
            assert_eq!(to.col, 4, "OneCell width must extend by 2 cols");
            assert_eq!(to.row, 6, "OneCell height must extend by 3 rows");
        }
        other => panic!("expected TwoCell after round-trip, got {other:?}"),
    }
}

/// Absolute anchor variant: input has explicit x/y position +
/// width/height. Writer encodes via ClientAnchor flag=3 (no move,
/// no resize). Excel must accept the file and the visual area
/// must survive.
#[test]
fn excel_can_read_xls_absolute_image_we_emit() {
    use duke_sheets_chart::{DrawingAnchor, EmbeddedImage, ImageFormat};

    const COL_EMU: i64 = 609_600;
    const ROW_EMU: i64 = 190_500;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "absolute-anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 2 * COL_EMU,
            height_emu: 4 * ROW_EMU,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        },
        DrawingAnchor::Absolute {
            x_emu: 3 * COL_EMU,
            y_emu: 2 * ROW_EMU,
            width_emu: 2 * COL_EMU,
            height_emu: 4 * ROW_EMU,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(
        images.len(),
        1,
        "Absolute picture must survive Excel re-save"
    );
    let img = images[0].payload;
    assert_eq!(img.format, ImageFormat::Png);
    match &images[0].object.unwrap().anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            // Absolute (x=3 cols, y=2 rows) + (2 cols × 4 rows) at
            // default cell sizes lands the picture starting at col=3
            // row=2 and ending at col=5 row=6.
            assert_eq!(from.col, 3, "Absolute x maps to col 3");
            assert_eq!(from.row, 2, "Absolute y maps to row 2");
            assert_eq!(to.col, 5, "width extends by 2 cols");
            assert_eq!(to.row, 6, "height extends by 4 rows");
        }
        other => panic!("expected TwoCell after round-trip, got {other:?}"),
    }
}

/// 58-byte 1x1 24-bit BMP (white pixel).
const TEST_BMP_1X1: &[u8] = &[
    0x42, 0x4D, 0x3A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x36, 0x00, 0x00, 0x00, 0x28, 0x00,
    0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x18, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x13, 0x0B, 0x00, 0x00, 0x13, 0x0B, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0x00,
];

/// BMP parity: Excel converts BMP input to PNG internally on
/// SaveAs (verified via probe). The test asserts Excel accepts our
/// DIB blip emit (no Repaired warning) and a single image survives
/// the round-trip. The format on the way back will be PNG (Excel's
/// re-encoding), so we assert the count + acceptance, not byte
/// equality.
#[test]
fn excel_can_read_xls_bmp_image_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "bmp-anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
            format: ImageFormat::Bmp,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_000_000,
            height_emu: 1_000_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_BMP_1X1.to_vec(),
            svg_data: None,
        },
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 1,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 4,
                col_offset_emu: 0,
                row: 6,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(
        images.len(),
        1,
        "BMP picture must survive Excel re-save (possibly as PNG)"
    );
    // Excel re-encodes BMP as PNG on SaveAs; format may flip to PNG.
    let img = images[0].payload;
    assert!(
        matches!(img.format, ImageFormat::Bmp | ImageFormat::Png),
        "expected BMP or PNG after Excel round-trip, got {:?}",
        img.format
    );
}

/// JPEG variant of the picture parity test. Confirms our writer
/// dispatches OfficeArtBlipJPEG correctly and Excel preserves the
/// JPEG bytes verbatim through its SaveAs.
#[test]
fn excel_can_read_xls_jpeg_image_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "jpeg-anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
            format: ImageFormat::Jpeg,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_000_000,
            height_emu: 1_000_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_JPEG_1X1.to_vec(),
            svg_data: None,
        },
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 1,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 4,
                col_offset_emu: 0,
                row: 6,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 1, "JPEG must survive Excel re-save");
    let img = images[0].payload;
    assert_eq!(img.format, ImageFormat::Jpeg, "format must stay JPEG");
    assert_eq!(
        img.data, TEST_JPEG_1X1,
        "JPEG bytes must round-trip through Excel verbatim"
    );
}

/// Analysis-ToolPak add-in functions (Ftab 384..=476). Excel does not emit a
/// native PtgFunc/PtgFuncVar for these; instead it writes an AddIn SUPBOOK
/// with one EXTERNNAME per distinct function and references it from each
/// formula via PtgNameX + PtgFuncVar(iftab=0x00FF). Sorted alphabetically the
/// distinct names here are EDATE(1), EOMONTH(2), GCD(3), NETWORKDAYS(4),
/// WORKDAY(5) — the 1-based EXTERNNAME nameindex.
const ATP_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=EDATE(A1,12)", 0.0),       // iftab=449, fixed 2-arg
    ("B2", "=EOMONTH(A1,1)", 0.0),      // iftab=450, fixed 2-arg
    ("B3", "=GCD(A1,A2)", 0.0),         // iftab=473, variable-arity
    ("B4", "=NETWORKDAYS(A1,A2)", 0.0), // iftab=472, variable-arity
    ("B5", "=WORKDAY(A1,5)", 0.0),      // iftab=471, variable-arity
];

fn atp_function_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 43831.0).unwrap(); // 2020-01-01 serial
    ws.set_cell_value("A2", 43862.0).unwrap(); // 2020-02-01 serial
    for (cell, formula, expected) in ATP_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_atp_xls_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 43831.0).expect("set A1");
        wb.set_cell_value("A2", 43862.0).expect("set A2");
        for (cell, formula, _) in ATP_FORMULAS {
            wb.set_cell_formula(cell, formula)
                .expect("set Excel formula");
        }
        excel.recalculate().expect("Excel recalculate");
        wb.save_as(&fixture.vm_path, 56).expect("Excel SaveAs xls");
        wb.close().expect("close Excel-authored workbook");
    }

    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

#[test]
fn excel_byte_parity_for_atp_functions_we_emit() {
    let wb = atp_function_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xls_bytes(&wb);

    let writer_ptgs = xls_formula_ptg_streams_for_compare(&writer_bytes);
    let resave_ptgs = xls_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLS ATP formula token streams on re-save"
    );

    assert_eq!(
        writer_ptgs.len(),
        ATP_FORMULAS.len(),
        "expected one formula stream per ATP formula"
    );

    let authored_bytes = excel_authored_atp_xls_bytes();
    let authored_ptgs = xls_formula_ptg_streams_for_compare(&authored_bytes);
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLS ATP (PtgNameX) token streams differ from Excel-authored output"
    );

    // The formula streams only pin the PtgNameX nameindex; byte-compare the
    // EXTERNNAME records themselves (grbit / reserved / cch / name / #REF!
    // placeholder formula) against Excel's native output.
    let writer_externnames = xls_externname_record_bodies(&writer_bytes);
    let authored_externnames = xls_externname_record_bodies(&authored_bytes);
    // Guard against a vacuous [] == [] pass: there are 5 distinct add-in
    // functions, so both sides must carry 5 EXTERNNAME records.
    assert_eq!(
        writer_externnames.len(),
        5,
        "writer must emit one EXTERNNAME per distinct add-in function"
    );
    assert_eq!(
        writer_externnames, authored_externnames,
        "our EXTERNNAME records differ from Excel-authored output"
    );
}

#[test]
fn excel_preserves_external_udf_xls_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 7.0).unwrap();
    ws.set_cell_formula("B1", r#"=[1]!TBLink("acct",A1)"#)
        .unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(42.0))
        .unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let ws = result.worksheet(0).unwrap();
    let formula = ws
        .get_formula_at(0, 1)
        .expect("external UDF formula must survive Excel re-save");
    assert!(
        formula.to_ascii_uppercase().contains("TBLINK"),
        "external UDF formula lost through Excel: {formula:?}"
    );
}

/// Comprehensive: every Analysis-ToolPak function (Ftab 384..=476) in one
/// workbook, byte-compared against Excel's native XLS emission — the add-in
/// form (PtgNameX + EXTERNNAME, R-class args, PtgFuncVar iftab=255). Emitted
/// in Ftab order (not alphabetical), so this also pins that our alphabetical
/// EXTERNNAME ordering / nameindex assignment matches Excel's: a divergent
/// order would shift every nameindex.
///
/// Verifies three things: (1) Excel opens our file with no "Repaired"
/// warning (via the round-trip helper) so the full 93-EXTERNNAME add-in
/// structure is valid at scale; (2) byte-parity of every formula token
/// stream; (3) byte-parity of all 93 EXTERNNAME record bodies, including
/// their alphabetical order.
#[test]
fn excel_byte_parity_for_all_xls_atp_functions_we_emit() {
    let formulas = atp_all_formulas();
    // Lock the coverage count: the add-in range is Ftab 384..=476 (93 fns).
    assert_eq!(formulas.len(), 93, "must exercise all 93 ATP functions");

    // Author each in Excel (XLS); keep the accepted set + Excel's bytes.
    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();
    let mut rejected = Vec::new();
    let mut accepted: Vec<(String, String)> = Vec::new();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        for i in 1..=10u32 {
            wb.set_cell_value(&format!("A{i}"), 30000.0 + i as f64)
                .unwrap();
        }
        for (cell, formula) in &formulas {
            match wb.set_cell_formula(cell, formula) {
                Ok(_) => accepted.push((cell.clone(), formula.clone())),
                Err(e) => rejected.push(format!("{formula} ({e})")),
            }
        }
        excel.recalculate().unwrap();
        wb.save_as(&fixture.vm_path, 56).expect("Excel SaveAs xls");
        wb.close().unwrap();
    }
    pull_file_from_vm(&fixture);
    let excel_bytes = std::fs::read(&fixture.host_path).unwrap();
    cleanup_fixture(&fixture);

    // Build our writer's workbook from the SAME accepted formulas. A cached
    // result is required: the XLS writer only emits FORMULA records for cells
    // that carry a value.
    let mut wb = Workbook::new();
    {
        let ws = wb.worksheet_mut(0).unwrap();
        for i in 1..=10u32 {
            ws.set_cell_value_at(i - 1, 0, 30000.0 + i as f64).unwrap();
        }
        for (cell, formula) in &accepted {
            ws.set_cell_formula(cell, formula).unwrap();
            let addr = CellAddress::parse(cell).unwrap();
            ws.set_formula_result(addr.row, addr.col, CellValue::Number(0.0))
                .unwrap();
        }
    }
    // Round-trip our file through Excel: this writes our bytes, opens them in
    // Excel asserting no "Repaired" warning (so the 93-EXTERNNAME add-in
    // structure is proven valid at scale, not just openable at 5 functions),
    // and returns our writer's bytes for the parity comparison.
    let (_result, our_bytes, _resave) = roundtrip_through_excel_xls_bytes(&wb);

    // The full range is expected to be valid — any rejection is a real
    // metadata bug (wrong min_args) to investigate, not a function to skip.
    assert!(
        rejected.is_empty(),
        "Excel rejected {} ATP call(s): {:?}",
        rejected.len(),
        rejected
    );
    assert_eq!(
        accepted.len(),
        formulas.len(),
        "all ATP formulas should be authored"
    );

    // Formula token streams: PtgNameX(nameindex) + R-class args + PtgFuncVar.
    let mut ours = xls_formula_ptg_streams_for_compare(&our_bytes);
    let mut excel = xls_formula_ptg_streams_for_compare(&excel_bytes);
    ours.sort_by_key(|s| (s.row, s.col));
    excel.sort_by_key(|s| (s.row, s.col));
    // Absolute anchor: both sides come from the same extractor, so a
    // broken extractor would otherwise compare empty-vs-empty.
    assert_eq!(
        ours.len(),
        accepted.len(),
        "formula-stream extraction came back short ({} of {})",
        ours.len(),
        accepted.len()
    );
    assert_eq!(
        ours.len(),
        excel.len(),
        "formula record count differs: ours={} excel={}",
        ours.len(),
        excel.len()
    );
    let mut diffs = Vec::new();
    for (w, a) in ours.iter().zip(excel.iter()) {
        if (w.row, w.col) != (a.row, a.col) || w.tokens != a.tokens || w.rgcb != a.rgcb {
            diffs.push(format!(
                "r{}c{} vs r{}c{}: ours tok={:02X?} rgcb={:02X?}\n        excel tok={:02X?} rgcb={:02X?}",
                w.row, w.col, a.row, a.col, w.tokens, w.rgcb, a.tokens, a.rgcb
            ));
        }
    }
    assert!(
        diffs.is_empty(),
        "XLS ATP token-stream mismatches ({} of {}):\n{}",
        diffs.len(),
        accepted.len(),
        diffs.join("\n")
    );

    // EXTERNNAME records: one per distinct function, our alphabetical order
    // must match Excel's byte-for-byte (count + bodies).
    let our_en = xls_externname_record_bodies(&our_bytes);
    let excel_en = xls_externname_record_bodies(&excel_bytes);
    assert_eq!(
        our_en.len(),
        accepted.len(),
        "one EXTERNNAME per accepted add-in function"
    );
    assert_eq!(
        our_en, excel_en,
        "our EXTERNNAME records (order/bodies) differ from Excel-authored output"
    );
}

/// Anchor helper for form-control parity tests.
fn control_anchor(from_col: u16, from_row: u32, to_col: u16, to_row: u32) -> DrawingAnchor {
    DrawingAnchor::TwoCell {
        from: CellMarker {
            col: from_col,
            col_offset_emu: 0,
            row: from_row,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: to_col,
            col_offset_emu: 0,
            row: to_row,
            row_offset_emu: 0,
        },
        edit_as: None,
    }
}

/// One of every Forms control kind survives the Excel round-trip
/// with its kind-specific data intact: captions, checked states,
/// cell links, input ranges, selections, and scroll parameters.
///
/// `no_3d` is deliberately not asserted: Excel rewrites the 3D flag
/// unconditionally on re-save, so only in-process round trips can
/// pin it.
#[test]
fn excel_can_read_form_controls_we_emit() {
    use duke_sheets_core::{CheckState, FormControl, FormControlKind, ListSelection};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");
    for (i, item) in ["Alpha", "Beta", "Gamma", "Delta"].iter().enumerate() {
        ws.set_cell_value_at(i as u32, 7, *item).expect("list item");
    }
    let kinds: Vec<FormControlKind> = vec![
        FormControlKind::Button {
            caption: "Run Report".into(),
        },
        FormControlKind::Checkbox {
            caption: "Enable audit".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$2".to_string()),
            no_3d: true,
        },
        FormControlKind::Checkbox {
            caption: "Tri state".into(),
            state: CheckState::Mixed,
            cell_link: None,
            no_3d: true,
        },
        FormControlKind::OptionButton {
            caption: "Opt A".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$3".to_string()),
            first_in_group: true,
            no_3d: true,
        },
        FormControlKind::OptionButton {
            caption: "Opt B".into(),
            state: CheckState::Unchecked,
            cell_link: None,
            first_in_group: false,
            no_3d: true,
        },
        FormControlKind::Label {
            caption: "Status label".into(),
        },
        FormControlKind::GroupBox {
            caption: "Choices".into(),
            no_3d: true,
        },
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: Some("$D$5".to_string()),
            selection: ListSelection::Single,
            selected: vec![3],
            no_3d: true,
        },
        FormControlKind::Dropdown {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: Some("$D$4".to_string()),
            selected: Some(2),
            lines: 6,
            no_3d: true,
        },
        FormControlKind::Scrollbar {
            value: 40,
            min: 5,
            max: 95,
            increment: 2,
            page: 10,
            horizontal: false,
            cell_link: Some("$D$6".to_string()),
        },
        FormControlKind::Spinner {
            value: 12,
            min: 0,
            max: 30,
            increment: 3,
            cell_link: Some("$D$7".to_string()),
        },
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![1, 3],
            no_3d: true,
        },
    ];
    let count = kinds.len();
    for (i, kind) in kinds.into_iter().enumerate() {
        let row = 1 + 2 * i as u32;
        ws.add_form_control(FormControl::new(kind), control_anchor(1, row, 3, row + 1)).unwrap();
    }
    assert_eq!(wb.sync_form_control_links(), 6);

    let result = roundtrip_through_excel_xls(&wb);
    let sheet = result.worksheet(0).unwrap();
    assert_eq!(sheet.get_value("D2").unwrap(), CellValue::Boolean(true));
    assert_eq!(sheet.get_value("D3").unwrap(), CellValue::Number(1.0));
    assert_eq!(sheet.get_value("D4").unwrap(), CellValue::Number(3.0));
    assert_eq!(sheet.get_value("D5").unwrap(), CellValue::Number(4.0));
    assert_eq!(sheet.get_value("D6").unwrap(), CellValue::Number(40.0));
    assert_eq!(sheet.get_value("D7").unwrap(), CellValue::Number(12.0));
    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(
        controls.len(),
        count,
        "every control must survive the Excel round-trip"
    );

    match &controls[0].payload.kind {
        FormControlKind::Button { caption } => assert_eq!(caption.plain_text(), "Run Report"),
        other => panic!("control 0: expected Button, got {other:?}"),
    }
    match &controls[1].payload.kind {
        FormControlKind::Checkbox {
            caption,
            state,
            cell_link,
            ..
        } => {
            assert_eq!(caption.plain_text(), "Enable audit");
            assert_eq!(*state, CheckState::Checked);
            assert_eq!(cell_link.as_deref(), Some("$D$2"));
        }
        other => panic!("control 1: expected Checkbox, got {other:?}"),
    }
    match &controls[2].payload.kind {
        FormControlKind::Checkbox { caption, state, .. } => {
            assert_eq!(caption.plain_text(), "Tri state");
            assert_eq!(*state, CheckState::Mixed, "mixed state must survive");
        }
        other => panic!("control 2: expected Checkbox, got {other:?}"),
    }
    match &controls[3].payload.kind {
        FormControlKind::OptionButton {
            caption,
            state,
            cell_link,
            first_in_group,
            ..
        } => {
            assert_eq!(caption.plain_text(), "Opt A");
            assert_eq!(*state, CheckState::Checked);
            assert_eq!(cell_link.as_deref(), Some("$D$3"));
            assert!(*first_in_group, "first radio keeps fFirstBtn");
        }
        other => panic!("control 3: expected OptionButton, got {other:?}"),
    }
    match &controls[4].payload.kind {
        FormControlKind::OptionButton { caption, state, .. } => {
            assert_eq!(caption.plain_text(), "Opt B");
            assert_eq!(*state, CheckState::Unchecked);
        }
        other => panic!("control 4: expected OptionButton, got {other:?}"),
    }
    match &controls[5].payload.kind {
        FormControlKind::Label { caption } => assert_eq!(caption.plain_text(), "Status label"),
        other => panic!("control 5: expected Label, got {other:?}"),
    }
    match &controls[6].payload.kind {
        FormControlKind::GroupBox { caption, .. } => {
            assert_eq!(caption.plain_text(), "Choices")
        }
        other => panic!("control 6: expected GroupBox, got {other:?}"),
    }
    match &controls[7].payload.kind {
        FormControlKind::ListBox {
            input_range,
            cell_link,
            selection,
            selected,
            ..
        } => {
            assert_eq!(input_range.as_deref(), Some("$H$1:$H$4"));
            assert_eq!(cell_link.as_deref(), Some("$D$5"));
            assert_eq!(*selection, ListSelection::Single);
            assert_eq!(selected, &vec![3]);
        }
        other => panic!("control 7: expected ListBox, got {other:?}"),
    }
    match &controls[8].payload.kind {
        FormControlKind::Dropdown {
            input_range,
            cell_link,
            selected,
            lines,
            ..
        } => {
            assert_eq!(input_range.as_deref(), Some("$H$1:$H$4"));
            assert_eq!(cell_link.as_deref(), Some("$D$4"));
            assert_eq!(*selected, Some(2));
            assert_eq!(*lines, 6);
        }
        other => panic!("control 8: expected Dropdown, got {other:?}"),
    }
    match &controls[9].payload.kind {
        FormControlKind::Scrollbar {
            value,
            min,
            max,
            increment,
            page,
            horizontal,
            cell_link,
        } => {
            assert_eq!(*value, 40);
            assert_eq!(*min, 5);
            assert_eq!(*max, 95);
            assert_eq!(*increment, 2);
            assert_eq!(*page, 10);
            assert!(!*horizontal);
            assert_eq!(cell_link.as_deref(), Some("$D$6"));
        }
        other => panic!("control 9: expected Scrollbar, got {other:?}"),
    }
    match &controls[10].payload.kind {
        FormControlKind::Spinner {
            value,
            min,
            max,
            increment,
            cell_link,
        } => {
            assert_eq!(*value, 12);
            assert_eq!(*min, 0);
            assert_eq!(*max, 30);
            assert_eq!(*increment, 3);
            assert_eq!(cell_link.as_deref(), Some("$D$7"));
        }
        other => panic!("control 10: expected Spinner, got {other:?}"),
    }
    match &controls[11].payload.kind {
        FormControlKind::ListBox {
            selection,
            selected,
            ..
        } => {
            assert_eq!(*selection, ListSelection::Multi);
            assert_eq!(selected, &vec![1, 3], "multi-selection must survive");
        }
        other => panic!("control 11: expected ListBox, got {other:?}"),
    }

    // Anchors survive (cell coordinates; offsets are requantised by
    // Excel and asserted in the in-process layer instead).
    match &controls[0].object.unwrap().anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            assert_eq!((from.col, from.row), (1, 1), "button anchor from");
            assert_eq!((to.col, to.row), (3, 2), "button anchor to");
        }
        other => panic!("expected TwoCell anchor, got {other:?}"),
    }
}

#[test]
fn excel_preserves_xls_control_visual_metadata_we_emit() {
    use duke_sheets_core::style::{HorizontalAlignment, Underline, VerticalAlignment};
    use duke_sheets_core::{CheckState, ControlText, DrawingObject, FormControl, FormControlKind};

    let text = ControlText {
        runs: vec![
            RichTextRun::with_font(
                "Red ",
                RunFont {
                    name: Some("Segoe UI".into()),
                    size: Some(9.0),
                    color: Some(Color::rgb(255, 0, 0)),
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
            RichTextRun::with_font(
                "Blue",
                RunFont {
                    name: Some("Arial".into()),
                    size: Some(12.0),
                    color: Some(Color::rgb(0, 0, 255)),
                    italic: Some(true),
                    underline: Some(Underline::Single),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: Some(HorizontalAlignment::Right),
        vertical_alignment: Some(VerticalAlignment::Bottom),
    };
    let control = FormControl::new(FormControlKind::Checkbox {
        caption: text,
        state: CheckState::Checked,
        cell_link: None,
        no_3d: false,
    })
    .with_macro_name("RunProbe");
    let mut object = DrawingObject::form_control(control).with_anchor(control_anchor(1, 1, 4, 3));
    object.meta.name = Some("Visual Probe".into());
    object.meta.alt_text = Some("Visual probe alternative".into());
    object.meta.title = Some("Not carried by XLS".into());
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();

    let result = roundtrip_through_excel_xls(&workbook);
    let drawn = result.worksheet(0).unwrap().form_controls().next().unwrap();
    assert_eq!(drawn.object.unwrap().meta.name.as_deref(), Some("Visual Probe"));
    assert_eq!(
        drawn.object.unwrap().meta.alt_text.as_deref(),
        Some("Visual probe alternative")
    );
    assert_eq!(drawn.payload.caption_text().as_deref(), Some("Red Blue"));
    assert_eq!(drawn.payload.macro_name.as_deref(), Some("RunProbe"));
    let caption = drawn.payload.caption().unwrap();
    assert_eq!(
        caption.horizontal_alignment,
        Some(HorizontalAlignment::Right)
    );
    assert_eq!(caption.vertical_alignment, Some(VerticalAlignment::Bottom));
    assert_eq!(caption.runs.len(), 2);
    let red = caption.runs[0].font.as_ref().unwrap();
    assert_eq!(red.name.as_deref(), Some("Segoe UI"));
    assert_eq!(red.size, Some(9.0));
    assert_eq!(red.color, Some(Color::rgb(255, 0, 0)));
    assert_eq!(red.bold, Some(true));
    let blue = caption.runs[1].font.as_ref().unwrap();
    assert_eq!(blue.name.as_deref(), Some("Arial"));
    assert_eq!(blue.size, Some(12.0));
    assert_eq!(blue.color, Some(Color::rgb(0, 0, 255)));
    assert_eq!(blue.italic, Some(true));
    assert_eq!(blue.underline, Some(Underline::Single));
}

#[test]
fn excel_preserves_xls_custom_metric_control_anchor_we_emit() {
    use duke_sheets_core::{CheckState, FormControl, FormControlKind};

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_column_width(0, 20.0);
    sheet.set_row_height(0, 30.0);
    sheet.add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "metric anchor".into(),
            state: CheckState::Unchecked,
            cell_link: None,
            no_3d: false,
        }),
        DrawingAnchor::OneCell {
            from: CellMarker::default(),
            width_emu: 609_600,
            height_emu: 190_500,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xls(&workbook);
    let drawn = result.worksheet(0).unwrap().form_controls().next().unwrap();
    match &drawn.object.unwrap().anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            assert_eq!((from.col, from.col_offset_emu), (0, 0));
            assert_eq!((from.row, from.row_offset_emu), (0, 0));
            assert_eq!((to.col, to.row), (0, 0));
            assert!(
                (to.col_offset_emu - 609_600).abs() <= 2_000,
                "XLS 1/1024-column anchor offset drifted: {}",
                to.col_offset_emu
            );
            assert!(
                (to.row_offset_emu - 190_500).abs() <= 2_000,
                "XLS 1/256-row anchor offset drifted: {}",
                to.row_offset_emu
            );
        }
        other => panic!("expected Excel-resaved TwoCell control anchor, got {other:?}"),
    }
}

/// Excel's own interpretation of persisted selections pins the
/// zero-based model ↔ one-based file conversion. A symmetric
/// writer/reader offset bug survives a model round-trip, but not
/// this: `ControlFormat.ListIndex` is one-based, so model index N
/// must surface as N+1 in Excel.
#[test]
fn excel_interprets_xls_list_selections_one_based() {
    use duke_sheets_core::{FormControl, FormControlKind, ListSelection};
    use duke_sheets_xls::XlsWriter;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (i, item) in ["Alpha", "Beta", "Gamma", "Delta"].iter().enumerate() {
        ws.set_cell_value_at(i as u32, 7, *item).expect("item");
    }
    ws.add_form_control(
        FormControl::new(FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selection: ListSelection::Single,
            selected: vec![2],
            no_3d: false,
        }),
        control_anchor(1, 1, 3, 3),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(FormControlKind::Dropdown {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selected: Some(1),
            lines: 8,
            no_3d: false,
        }),
        control_anchor(1, 5, 3, 6),
    ).unwrap();

    let fixture = temp_fixture_xls();
    let bytes = XlsWriter::write_to_bytes(&wb).expect("write xls");
    std::fs::write(&fixture.host_path, &bytes).expect("write fixture");
    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&fixture.vm_path)
        .expect("Excel should open our XLS without error");
    let name = opened.name().expect("workbook name");
    assert!(
        !name.contains("Repaired"),
        "Excel repaired the file: {name}"
    );

    // `Shapes.Item` is a method in Excel's type library, so it is
    // unreachable through chain steps (GetProperty binding); invoke
    // it explicitly.
    let shapes_handle = excel
        .navigate(
            opened.handle(),
            vec![
                SheetRef::Index(0).to_chain_step(),
                ChainStep::Property("Shapes".into()),
            ],
        )
        .expect("navigate shapes");
    let list_index = |shape: u32| -> f64 {
        let shape_handle = match excel.invoke(
            shapes_handle,
            vec![],
            "Item",
            vec![serde_json::Value::from(shape)],
        ) {
            Ok(Some(ResponseData::Handle { handle })) => handle,
            other => panic!("expected shape handle, got {other:?}"),
        };
        let index = match excel.get(
            shape_handle,
            vec![ChainStep::Property("ControlFormat".into())],
            "ListIndex",
        ) {
            Ok(Some(ResponseData::Value { value })) => value.as_f64().expect("numeric ListIndex"),
            other => panic!("expected ListIndex value, got {other:?}"),
        };
        excel.release(shape_handle).expect("release shape");
        index
    };
    assert_eq!(
        list_index(1),
        3.0,
        "list box: model index 2 is Excel item 3"
    );
    assert_eq!(
        list_index(2),
        2.0,
        "dropdown: model index 1 is Excel item 2"
    );

    excel.release(shapes_handle).expect("release shapes");
    opened.close().expect("close workbook");
    cleanup_fixture(&fixture);

    // Read side: Excel authors the controls (ListIndex is one-based),
    // saves XLS, and our reader must surface zero-based indices.
    let authored = excel.create_workbook().expect("create workbook");
    for (i, item) in ["Alpha", "Beta", "Gamma", "Delta"].iter().enumerate() {
        authored
            .set_cell_value(&format!("H{}", i + 1), *item)
            .expect("item");
    }
    let sheet_shapes = excel
        .navigate(
            authored.handle(),
            vec![
                SheetRef::Index(0).to_chain_step(),
                ChainStep::Property("Shapes".into()),
            ],
        )
        .expect("navigate authored shapes");
    // xlListBox = 6, xlDropDown = 2
    for (control_type, list_index) in [(6, 3), (2, 2)] {
        let shape = match excel.invoke(
            sheet_shapes,
            vec![],
            "AddFormControl",
            vec![
                serde_json::Value::from(control_type),
                serde_json::Value::from(10 * control_type),
                serde_json::Value::from(10),
                serde_json::Value::from(60),
                serde_json::Value::from(60),
            ],
        ) {
            Ok(Some(ResponseData::Handle { handle })) => handle,
            other => panic!("expected AddFormControl handle, got {other:?}"),
        };
        excel
            .set(
                shape,
                vec![ChainStep::Property("ControlFormat".into())],
                "ListFillRange",
                serde_json::Value::from("$H$1:$H$4"),
            )
            .expect("set ListFillRange");
        excel
            .set(
                shape,
                vec![ChainStep::Property("ControlFormat".into())],
                "ListIndex",
                serde_json::Value::from(list_index),
            )
            .expect("set ListIndex");
        excel.release(shape).ok();
    }
    let out_xls = temp_fixture_xls();
    authored.save_as(&out_xls.vm_path, 56).expect("save xls");
    authored.close().expect("close authored");
    excel.release(sheet_shapes).ok();
    pull_file_from_vm(&out_xls);
    let authored_model = duke_sheets_xls::XlsReader::read_file(&out_xls.host_path).expect("read");
    let authored_controls: Vec<_> = authored_model
        .worksheet(0)
        .unwrap()
        .form_controls()
        .collect();
    match &authored_controls[0].payload.kind {
        FormControlKind::ListBox { selected, .. } => {
            assert_eq!(selected, &vec![2], "Excel item 3 is model index 2");
        }
        other => panic!("expected ListBox, got {other:?}"),
    }
    match &authored_controls[1].payload.kind {
        FormControlKind::Dropdown { selected, .. } => {
            assert_eq!(*selected, Some(1), "Excel item 2 is model index 1");
        }
        other => panic!("expected Dropdown, got {other:?}"),
    }
    cleanup_fixture(&out_xls);
}

/// Excel-authored controls pin how each form-control state projects into its
/// linked cell, including the cases where the cell stays blank or becomes an
/// error. The saved XLS assertions also verify that those values and a true
/// multi-selection survive BIFF8 serialization.
#[test]
fn excel_authored_xls_form_control_linked_cell_semantics() {
    use duke_sheets_core::{CellError, FormControlKind, ListSelection};
    use excel_com_protocol::CellValue as ExcelCellValue;

    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let workbook = excel.create_workbook().expect("create Excel workbook");
    rename_excel_sheet(&excel, workbook.handle(), 0, "Controls");
    add_excel_worksheet_after(&excel, workbook.handle(), 0, "Linked Data");
    for (i, item) in ["Alpha", "Beta", "Gamma", "Delta"].iter().enumerate() {
        workbook
            .set_cell_value(&format!("H{}", i + 1), *item)
            .expect("set list item");
    }

    let shapes = excel
        .navigate(
            workbook.handle(),
            vec![
                SheetRef::Index(0).to_chain_step(),
                ChainStep::Property("Shapes".into()),
            ],
        )
        .expect("navigate shapes");
    let add_control = |control_type: i32, left: i32, top: i32| -> u64 {
        match excel.invoke(
            shapes,
            vec![],
            "AddFormControl",
            vec![
                serde_json::Value::from(control_type),
                serde_json::Value::from(left),
                serde_json::Value::from(top),
                serde_json::Value::from(100),
                serde_json::Value::from(30),
            ],
        ) {
            Ok(Some(ResponseData::Handle { handle })) => handle,
            other => panic!("expected AddFormControl handle, got {other:?}"),
        }
    };
    let set_control = |shape: u64, property: &str, value: serde_json::Value| {
        excel
            .set(
                shape,
                vec![ChainStep::Property("ControlFormat".into())],
                property,
                value,
            )
            .unwrap_or_else(|e| panic!("set ControlFormat.{property}: {e}"));
    };
    let get_control = |shape: u64, property: &str| -> serde_json::Value {
        match excel
            .get(
                shape,
                vec![ChainStep::Property("ControlFormat".into())],
                property,
            )
            .unwrap_or_else(|e| panic!("get ControlFormat.{property}: {e}"))
        {
            Some(ResponseData::Value { value }) => value,
            other => panic!("expected ControlFormat.{property} value, got {other:?}"),
        }
    };

    // xlCheckBox = 1; xlOff = -4146, xlOn = 1, xlMixed = 2.
    let checkbox = add_control(1, 10, 10);
    set_control(checkbox, "LinkedCell", json!("$A$1"));
    assert_eq!(workbook.get_cell_value("A1").unwrap(), ExcelCellValue::Null);
    set_control(checkbox, "Value", json!(-4146));
    assert_eq!(workbook.get_cell_value("A1").unwrap(), ExcelCellValue::Null);
    set_control(checkbox, "Value", json!(1));
    assert_eq!(
        workbook.get_cell_value("A1").unwrap(),
        ExcelCellValue::Bool(true)
    );
    set_control(checkbox, "Value", json!(-4146));
    assert_eq!(
        workbook.get_cell_value("A1").unwrap(),
        ExcelCellValue::Bool(false)
    );
    set_control(checkbox, "Value", json!(2));
    assert_eq!(
        workbook.get_cell_value("A1").unwrap(),
        ExcelCellValue::Number(-2146826246.0),
        "the bridge exposes Excel's #N/A CVErr as its signed numeric value"
    );

    // xlListBox = 6; xlNone = -4142, xlSimple = -4154.
    let single_list = add_control(6, 10, 60);
    set_control(single_list, "ListFillRange", json!("$H$1:$H$4"));
    set_control(single_list, "LinkedCell", json!("$A$2"));
    assert_eq!(workbook.get_cell_value("A2").unwrap(), ExcelCellValue::Null);
    set_control(single_list, "ListIndex", json!(3));
    assert_eq!(
        workbook.get_cell_value("A2").unwrap(),
        ExcelCellValue::Number(3.0)
    );

    let multi_list = add_control(6, 10, 110);
    set_control(multi_list, "ListFillRange", json!("$H$1:$H$4"));
    set_control(multi_list, "MultiSelect", json!(-4154));
    assert_eq!(get_control(multi_list, "MultiSelect"), json!(-4154));
    assert_eq!(get_control(multi_list, "LinkedCell"), json!(""));
    assert_eq!(workbook.get_cell_value("A3").unwrap(), ExcelCellValue::Null);

    // Selected is an indexed property, which the generic bridge cannot put.
    // A temporary module drives the same Excel object and is removed pre-save.
    let vb_components = excel
        .navigate(
            workbook.handle(),
            vec![
                ChainStep::Property("VBProject".into()),
                ChainStep::Property("VBComponents".into()),
            ],
        )
        .expect("navigate VB components");
    let probe_module = match excel.invoke(vb_components, vec![], "Add", vec![json!(1)]) {
        Ok(Some(ResponseData::Handle { handle })) => handle,
        other => panic!("expected VB module handle, got {other:?}"),
    };
    excel
        .invoke(
            probe_module,
            vec![ChainStep::Property("CodeModule".into())],
            "AddFromString",
            vec![json!(
                "Sub SelectProbeItems()\nWith Worksheets(\"Controls\").ListBoxes(2)\n.Selected(1) = True\n.Selected(3) = True\n.Selected(4) = True\nEnd With\nEnd Sub"
            )],
        )
        .expect("add selection probe macro");
    excel
        .invoke(0, vec![], "Run", vec![json!("SelectProbeItems")])
        .expect("run selection probe macro");
    assert_eq!(workbook.get_cell_value("A3").unwrap(), ExcelCellValue::Null);
    assert_eq!(get_control(multi_list, "MultiSelect"), json!(-4154));
    excel
        .invoke(
            vb_components,
            vec![],
            "Remove",
            vec![json!({"$ref": probe_module})],
        )
        .expect("remove selection probe macro");
    workbook
        .set_cell_value("A3", 99.0)
        .expect("preseed multi-select linked cell");
    set_control(multi_list, "LinkedCell", json!("$A$3"));
    assert_eq!(get_control(multi_list, "LinkedCell"), json!("$A$3"));
    assert_eq!(get_control(multi_list, "MultiSelect"), json!(-4154));
    assert_eq!(
        workbook.get_cell_value("A3").unwrap(),
        ExcelCellValue::Number(0.0)
    );

    // xlDropDown = 2.
    let dropdown = add_control(2, 10, 160);
    set_control(dropdown, "ListFillRange", json!("$H$1:$H$4"));
    set_control(dropdown, "LinkedCell", json!("$A$4"));
    assert_eq!(workbook.get_cell_value("A4").unwrap(), ExcelCellValue::Null);
    set_control(dropdown, "ListIndex", json!(2));
    assert_eq!(
        workbook.get_cell_value("A4").unwrap(),
        ExcelCellValue::Number(2.0)
    );

    // xlScrollBar = 8; xlSpinner = 9.
    let scrollbar = add_control(8, 150, 10);
    set_control(scrollbar, "Min", json!(5));
    set_control(scrollbar, "Max", json!(95));
    set_control(scrollbar, "Value", json!(40));
    set_control(scrollbar, "LinkedCell", json!("$A$5"));
    assert_eq!(
        workbook.get_cell_value("A5").unwrap(),
        ExcelCellValue::Number(40.0)
    );
    set_control(scrollbar, "Value", json!(55));
    assert_eq!(
        workbook.get_cell_value("A5").unwrap(),
        ExcelCellValue::Number(55.0)
    );

    let spinner = add_control(9, 150, 60);
    set_control(spinner, "Min", json!(0));
    set_control(spinner, "Max", json!(30));
    set_control(spinner, "Value", json!(12));
    set_control(spinner, "LinkedCell", json!("$A$6"));
    assert_eq!(
        workbook.get_cell_value("A6").unwrap(),
        ExcelCellValue::Number(12.0)
    );
    set_control(spinner, "Value", json!(18));
    assert_eq!(
        workbook.get_cell_value("A6").unwrap(),
        ExcelCellValue::Number(18.0)
    );

    // xlGroupBox = 4; xlOptionButton = 7.
    let group_box = add_control(4, 300, 10);
    excel
        .set(group_box, vec![], "Height", json!(130))
        .expect("size option group box");
    let option_1 = add_control(7, 310, 35);
    let option_2 = add_control(7, 310, 65);
    let option_3 = add_control(7, 310, 95);
    for option in [option_1, option_2, option_3] {
        set_control(option, "LinkedCell", json!("$A$7"));
    }
    assert_eq!(workbook.get_cell_value("A7").unwrap(), ExcelCellValue::Null);
    set_control(option_2, "Value", json!(1));
    assert_eq!(
        workbook.get_cell_value("A7").unwrap(),
        ExcelCellValue::Number(2.0)
    );

    let formula_checkbox = add_control(1, 150, 110);
    set_control(formula_checkbox, "LinkedCell", json!("$A$8"));
    workbook
        .set_cell_formula("A8", "=FALSE")
        .expect("set linked-cell formula");
    excel.recalculate().expect("recalculate formula");
    assert_eq!(
        workbook.get_cell_value("A8").unwrap(),
        ExcelCellValue::Bool(false)
    );
    assert_eq!(workbook.get_cell_formula("A8").unwrap(), "=FALSE");
    assert_eq!(get_control(formula_checkbox, "Value"), json!(-4146));
    set_control(formula_checkbox, "Value", json!(1));
    assert_eq!(workbook.get_cell_formula("A8").unwrap(), "TRUE");
    assert_eq!(
        workbook.get_cell_value("A8").unwrap(),
        ExcelCellValue::Bool(true)
    );
    assert_eq!(get_control(formula_checkbox, "Value"), json!(1));

    let cross_sheet_checkbox = add_control(1, 150, 160);
    set_control(
        cross_sheet_checkbox,
        "LinkedCell",
        json!("'Linked Data'!$A$1"),
    );
    set_control(cross_sheet_checkbox, "Value", json!(1));
    assert_eq!(
        workbook
            .get_cell_value_on_sheet(SheetRef::Name("Linked Data".into()), "A1")
            .unwrap(),
        ExcelCellValue::Bool(true)
    );

    // Excel also accepts a defined name as the cell link.
    define_excel_name(&excel, workbook.handle(), "LinkedTarget", "=Controls!$B$1");
    let named_checkbox = add_control(1, 150, 210);
    set_control(named_checkbox, "LinkedCell", json!("LinkedTarget"));
    assert_eq!(
        get_control(named_checkbox, "LinkedCell"),
        json!("LinkedTarget")
    );
    set_control(named_checkbox, "Value", json!(1));
    assert_eq!(
        workbook.get_cell_value("B1").unwrap(),
        ExcelCellValue::Bool(true)
    );

    workbook.save_as(&fixture.vm_path, 56).expect("save xls");
    for handle in [
        checkbox,
        single_list,
        multi_list,
        dropdown,
        scrollbar,
        spinner,
        option_1,
        option_2,
        option_3,
        group_box,
        formula_checkbox,
        cross_sheet_checkbox,
        named_checkbox,
        probe_module,
        vb_components,
        shapes,
    ] {
        excel.release(handle).expect("release COM handle");
    }
    workbook.close().expect("close workbook");
    pull_file_from_vm(&fixture);
    let model = duke_sheets_xls::XlsReader::read_file(&fixture.host_path).expect("read xls");
    let controls_sheet = model.worksheet(0).unwrap();
    assert_eq!(
        controls_sheet.get_value("A1").unwrap(),
        CellValue::Error(CellError::Na)
    );
    assert_eq!(
        controls_sheet.get_value("A2").unwrap(),
        CellValue::Number(3.0)
    );
    assert_eq!(
        controls_sheet.get_value("A3").unwrap(),
        CellValue::Number(0.0)
    );
    assert_eq!(
        controls_sheet.get_value("A4").unwrap(),
        CellValue::Number(2.0)
    );
    assert_eq!(
        controls_sheet.get_value("A5").unwrap(),
        CellValue::Number(55.0)
    );
    assert_eq!(
        controls_sheet.get_value("A6").unwrap(),
        CellValue::Number(18.0)
    );
    assert_eq!(
        controls_sheet.get_value("A7").unwrap(),
        CellValue::Number(2.0)
    );
    assert_eq!(
        controls_sheet.get_value("A8").unwrap(),
        CellValue::Boolean(true)
    );
    assert_eq!(
        model.worksheet(1).unwrap().get_value("A1").unwrap(),
        CellValue::Boolean(true)
    );
    let controls: Vec<_> = controls_sheet.form_controls().collect();
    match &controls[2].payload.kind {
        FormControlKind::ListBox {
            cell_link,
            selection,
            selected,
            ..
        } => {
            assert_eq!(cell_link.as_deref(), Some("$A$3"));
            assert_eq!(*selection, ListSelection::Multi);
            assert_eq!(selected, &vec![0, 2, 3]);
        }
        other => panic!("expected multi-select ListBox, got {other:?}"),
    }
    // The name-linked checkbox persists its link as the defined name,
    // and the name's target carries the pushed value.
    assert_eq!(
        controls_sheet.get_value("B1").unwrap(),
        CellValue::Boolean(true)
    );
    let named_link = controls
        .iter()
        .filter_map(|control| control.payload.cell_link())
        .find(|link| link.contains("LinkedTarget"));
    assert_eq!(named_link, Some("LinkedTarget"));
    cleanup_fixture(&fixture);
}

/// Excel-authored probe pinning how a FORMULA in a linked cell drives
/// each control kind when it recalculates: the checkbox on/off/mixed
/// mapping, scrollbar clamping (and whether Excel rewrites the formula
/// cell to the clamp), list box out-of-range handling, and radio-group
/// index selection. The saved XLS must persist the driven control
/// states while the linked cells keep their formulas.
#[test]
fn excel_authored_xls_linked_formulas_drive_controls() {
    use duke_sheets_core::{CellError, CheckState, FormControlKind};
    use excel_com_protocol::CellValue as ExcelCellValue;

    let fixture = temp_fixture_xls();
    ensure_vm_temp_dir();

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let workbook = excel.create_workbook().expect("create Excel workbook");
    for (i, item) in ["Alpha", "Beta", "Gamma", "Delta"].iter().enumerate() {
        workbook
            .set_cell_value(&format!("H{}", i + 1), *item)
            .expect("set list item");
    }

    let shapes = excel
        .navigate(
            workbook.handle(),
            vec![
                SheetRef::Index(0).to_chain_step(),
                ChainStep::Property("Shapes".into()),
            ],
        )
        .expect("navigate shapes");
    let add_control = |control_type: i32, left: i32, top: i32| -> u64 {
        match excel.invoke(
            shapes,
            vec![],
            "AddFormControl",
            vec![
                serde_json::Value::from(control_type),
                serde_json::Value::from(left),
                serde_json::Value::from(top),
                serde_json::Value::from(100),
                serde_json::Value::from(30),
            ],
        ) {
            Ok(Some(ResponseData::Handle { handle })) => handle,
            other => panic!("expected AddFormControl handle, got {other:?}"),
        }
    };
    let set_control = |shape: u64, property: &str, value: serde_json::Value| {
        excel
            .set(
                shape,
                vec![ChainStep::Property("ControlFormat".into())],
                property,
                value,
            )
            .unwrap_or_else(|e| panic!("set ControlFormat.{property}: {e}"));
    };
    let get_control = |shape: u64, property: &str| -> serde_json::Value {
        match excel
            .get(
                shape,
                vec![ChainStep::Property("ControlFormat".into())],
                property,
            )
            .unwrap_or_else(|e| panic!("get ControlFormat.{property}: {e}"))
        {
            Some(ResponseData::Value { value }) => value,
            other => panic!("expected ControlFormat.{property} value, got {other:?}"),
        }
    };
    let drive = |cell: &str, formula: &str| {
        workbook
            .set_cell_formula(cell, formula)
            .unwrap_or_else(|e| panic!("set {cell} formula {formula}: {e}"));
        excel
            .recalculate()
            .unwrap_or_else(|e| panic!("recalculate after {cell} = {formula}: {e}"));
    };

    // xlCheckBox = 1 linked to $A$1; xlOn = 1, xlOff = -4146, xlMixed = 2.
    let checkbox = add_control(1, 10, 10);
    set_control(checkbox, "LinkedCell", json!("$A$1"));
    let checkbox_values: Vec<serde_json::Value> = ["=5", "=0", "=TRUE", "=\"text\"", "=NA()"]
        .into_iter()
        .map(|formula| {
            drive("A1", formula);
            let value = get_control(checkbox, "Value");
            eprintln!("checkbox: A1 {formula} -> Value {value}");
            value
        })
        .collect();

    // xlScrollBar = 8 (Min=5, Max=95) linked to $A$2. Alongside the
    // control's Value, capture what the cell holds afterwards to pin
    // whether Excel rewrites the formula cell to the clamp.
    let scrollbar = add_control(8, 150, 10);
    set_control(scrollbar, "Min", json!(5));
    set_control(scrollbar, "Max", json!(95));
    set_control(scrollbar, "LinkedCell", json!("$A$2"));
    let scrollbar_probe: Vec<(serde_json::Value, ExcelCellValue, String)> = ["=150", "=2", "=40"]
        .into_iter()
        .map(|formula| {
            drive("A2", formula);
            let value = get_control(scrollbar, "Value");
            let cell_value = workbook.get_cell_value("A2").expect("read A2 value");
            let cell_formula = workbook.get_cell_formula("A2").expect("read A2 formula");
            eprintln!(
                "scrollbar: A2 {formula} -> Value {value}, cell {cell_value:?} ({cell_formula})"
            );
            (value, cell_value, cell_formula)
        })
        .collect();

    // xlListBox = 6, single-select over $H$1:$H$4, linked to $A$3.
    let list_box = add_control(6, 10, 60);
    set_control(list_box, "ListFillRange", json!("$H$1:$H$4"));
    set_control(list_box, "LinkedCell", json!("$A$3"));
    let list_indexes: Vec<serde_json::Value> = ["=2", "=99", "=0"]
        .into_iter()
        .map(|formula| {
            drive("A3", formula);
            let index = get_control(list_box, "ListIndex");
            eprintln!("list box: A3 {formula} -> ListIndex {index}");
            index
        })
        .collect();

    // xlOptionButton = 7 pair sharing $A$4. They are the only radios on
    // the sheet, so they form the sheet-level group by themselves.
    let option_1 = add_control(7, 300, 10);
    let option_2 = add_control(7, 300, 50);
    for option in [option_1, option_2] {
        set_control(option, "LinkedCell", json!("$A$4"));
    }
    let radio_probe: Vec<(serde_json::Value, serde_json::Value)> = ["=2", "=0"]
        .into_iter()
        .map(|formula| {
            drive("A4", formula);
            let first = get_control(option_1, "Value");
            let second = get_control(option_2, "Value");
            eprintln!("radios: A4 {formula} -> ({first}, {second})");
            (first, second)
        })
        .collect();

    assert_eq!(
        checkbox_values,
        vec![json!(1), json!(-4146), json!(1), json!(1), json!(2)],
        "checkbox: nonzero, TRUE and text all check; 0 unchecks; only #N/A mixes"
    );
    assert_eq!(
        scrollbar_probe,
        vec![
            (json!(95), ExcelCellValue::Number(150.0), "=150".to_string()),
            (json!(5), ExcelCellValue::Number(2.0), "=2".to_string()),
            (json!(40), ExcelCellValue::Number(40.0), "=40".to_string()),
        ],
        "scrollbar Value clamps to Min/Max; the formula cell is not rewritten"
    );
    assert_eq!(
        list_indexes,
        vec![json!(2), json!(4), json!(0)],
        "list box: in-range selects; out-of-range clamps to the last item; 0 deselects"
    );
    assert_eq!(
        radio_probe,
        vec![(json!(-4146), json!(1)), (json!(-4146), json!(-4146))],
        "radio pair (first, second) Value per formula =2, =0"
    );

    workbook.save_as(&fixture.vm_path, 56).expect("save xls");
    for handle in [checkbox, scrollbar, list_box, option_1, option_2, shapes] {
        excel.release(handle).expect("release COM handle");
    }
    workbook.close().expect("close workbook");
    pull_file_from_vm(&fixture);
    let model = duke_sheets_xls::XlsReader::read_file(&fixture.host_path).expect("read xls");
    let sheet = model.worksheet(0).unwrap();

    // The linked cells persist as formulas with recalculated caches,
    // not as constants pushed back by the controls.
    assert_eq!(sheet.get_formula_at(0, 0), Some("=NA()"));
    assert_eq!(
        sheet.get_value("A1").unwrap(),
        CellValue::Error(CellError::Na)
    );
    assert_eq!(sheet.get_formula_at(1, 0), Some("=40"));
    assert_eq!(sheet.get_value("A2").unwrap(), CellValue::Number(40.0));
    assert_eq!(sheet.get_formula_at(2, 0), Some("=0"));
    assert_eq!(sheet.get_value("A3").unwrap(), CellValue::Number(0.0));
    assert_eq!(sheet.get_formula_at(3, 0), Some("=0"));
    assert_eq!(sheet.get_value("A4").unwrap(), CellValue::Number(0.0));

    // The persisted control states match the last formula-driven states.
    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), 5, "all controls survive the save");
    match &controls[0].payload.kind {
        FormControlKind::Checkbox {
            state, cell_link, ..
        } => {
            assert_eq!(
                *state,
                CheckState::Mixed,
                "checkbox persists the #N/A-driven mixed state"
            );
            assert_eq!(cell_link.as_deref(), Some("$A$1"));
        }
        other => panic!("expected Checkbox, got {other:?}"),
    }
    match &controls[1].payload.kind {
        FormControlKind::Scrollbar {
            value,
            min,
            max,
            cell_link,
            ..
        } => {
            assert_eq!((*min, *max, *value), (5, 95, 40));
            assert_eq!(cell_link.as_deref(), Some("$A$2"));
        }
        other => panic!("expected Scrollbar, got {other:?}"),
    }
    match &controls[2].payload.kind {
        FormControlKind::ListBox {
            selected,
            cell_link,
            ..
        } => {
            assert_eq!(
                selected,
                &Vec::<u16>::new(),
                "ListIndex 0 persists as no selection"
            );
            assert_eq!(cell_link.as_deref(), Some("$A$3"));
        }
        other => panic!("expected ListBox, got {other:?}"),
    }
    let radios: Vec<(CheckState, Option<&str>, bool)> = controls[3..]
        .iter()
        .map(|control| match &control.payload.kind {
            FormControlKind::OptionButton {
                state,
                cell_link,
                first_in_group,
                ..
            } => (*state, cell_link.as_deref(), *first_in_group),
            other => panic!("expected OptionButton, got {other:?}"),
        })
        .collect();
    assert_eq!(
        radios,
        vec![
            (CheckState::Unchecked, Some("$A$4"), true),
            (CheckState::Unchecked, None, false),
        ],
        "=0 leaves both radios unchecked; Excel stores the group link on the first radio only"
    );
    cleanup_fixture(&fixture);
}

/// An unselected radio group whose linked cell held a stale value
/// must survive Excel: synchronization resets the cell to 0, which
/// Excel reads back as "no radio checked". Without the reset, Excel
/// would check the radio at the stale one-based index on open.
#[test]
fn excel_preserves_unselected_radio_group_over_stale_link() {
    use duke_sheets_core::{CheckState, FormControl, FormControlKind};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("D1", 2.0).expect("stale link value");
    for (caption, row) in [("First", 1), ("Second", 3)] {
        ws.add_form_control(
            FormControl::new(FormControlKind::OptionButton {
                caption: caption.into(),
                state: CheckState::Unchecked,
                cell_link: Some("$D$1".to_string()),
                first_in_group: false,
                no_3d: true,
            }),
            control_anchor(1, row, 2, row + 1),
        ).unwrap();
    }
    assert_eq!(wb.sync_form_control_links(), 1);
    assert_eq!(
        wb.worksheet(0).unwrap().get_value("D1").unwrap(),
        CellValue::Number(0.0)
    );

    let result = roundtrip_through_excel_xls(&wb);
    let sheet = result.worksheet(0).unwrap();
    assert_eq!(sheet.get_value("D1").unwrap(), CellValue::Number(0.0));
    let states: Vec<CheckState> = sheet
        .form_controls()
        .map(|control| match &control.payload.kind {
            FormControlKind::OptionButton { state, .. } => *state,
            other => panic!("expected OptionButton, got {other:?}"),
        })
        .collect();
    assert_eq!(
        states,
        vec![CheckState::Unchecked, CheckState::Unchecked],
        "no radio may become checked from the stale link"
    );
}

/// Radio grouping by enclosing group box survives the Excel
/// round-trip: two boxes with two radios each plus a loose radio
/// come back as three groups (fFirstBtn on each group's first
/// radio), with per-group checked states intact.
#[test]
fn excel_can_read_radio_groups_we_emit() {
    use duke_sheets_core::{CheckState, FormControl, FormControlKind};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");

    let radio = |caption: &str, state: CheckState| FormControlKind::OptionButton {
        caption: caption.into(),
        state,
        cell_link: None,
        first_in_group: false,
        no_3d: true,
    };
    let group_box = |caption: &str| FormControlKind::GroupBox {
        caption: caption.into(),
        no_3d: true,
    };
    ws.add_form_control(
        FormControl::new(group_box("Box A")),
        control_anchor(0, 0, 2, 6),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(group_box("Box B")),
        control_anchor(4, 0, 6, 6),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(radio("A1", CheckState::Checked)),
        control_anchor(1, 1, 2, 2),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(radio("B1", CheckState::Unchecked)),
        control_anchor(5, 1, 6, 2),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(radio("A2", CheckState::Unchecked)),
        control_anchor(1, 3, 2, 4),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(radio("B2", CheckState::Checked)),
        control_anchor(5, 3, 6, 4),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(radio("Loose", CheckState::Unchecked)),
        control_anchor(8, 1, 9, 2),
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let sheet = result.worksheet(0).unwrap();
    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), 7, "all controls survive");

    let radios: Vec<(String, CheckState, bool)> = controls
        .iter()
        .filter_map(|c| match &c.payload.kind {
            FormControlKind::OptionButton {
                caption,
                state,
                first_in_group,
                ..
            } => Some((caption.plain_text(), *state, *first_in_group)),
            _ => None,
        })
        .collect();
    assert_eq!(
        radios,
        vec![
            ("A1".to_string(), CheckState::Checked, true),
            ("B1".to_string(), CheckState::Unchecked, true),
            ("A2".to_string(), CheckState::Unchecked, false),
            ("B2".to_string(), CheckState::Checked, false),
            ("Loose".to_string(), CheckState::Unchecked, true),
        ],
        "three radio groups with per-group states must survive Excel"
    );
}

/// A form control coexisting with a comment and a picture keeps the
/// OBJ↔shape pairing straight through Excel's re-save.
#[test]
fn excel_can_read_mixed_control_comment_picture_we_emit() {
    use duke_sheets_core::{CheckState, FormControl, FormControlKind};

    const TEST_PNG_1X1: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60,
        0x00, 0x02, 0x00, 0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00,
        0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 7.0).expect("A1");
    ws.set_cell_value("D2", true).expect("D2");
    ws.add_image(
        duke_sheets_chart::EmbeddedImage {
            format: duke_sheets_chart::ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_000_000,
            height_emu: 1_000_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        },
        control_anchor(6, 1, 8, 4),
    ).unwrap();
    ws.set_comment_at(0, 0, duke_sheets_core::CellComment::new("Author", "a note")).unwrap();
    ws.add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "mixed sheet".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$2".to_string()),
            no_3d: true,
        }),
        control_anchor(1, 1, 3, 2),
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let sheet = result.worksheet(0).unwrap();
    assert_eq!(sheet.image_count(), 1, "picture survives");
    assert_eq!(sheet.comment_count(), 1, "comment survives");
    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), 1, "control survives");
    match &controls[0].payload.kind {
        FormControlKind::Checkbox {
            caption,
            state,
            cell_link,
            ..
        } => {
            assert_eq!(caption.plain_text(), "mixed sheet");
            assert_eq!(*state, CheckState::Checked);
            assert_eq!(cell_link.as_deref(), Some("$D$2"));
        }
        other => panic!("expected Checkbox, got {other:?}"),
    }
}

/// A multi-select list whose bsels array exceeds one OBJ body uses
/// CONTINUE records and survives Excel without repair.
#[test]
fn excel_can_read_large_xls_list_control_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::{FormControl, FormControlKind, ListSelection};

    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_form_control(
        FormControl::new(FormControlKind::ListBox {
            input_range: Some("$H$1:$H$10000".to_string()),
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![0, 4_999, 9_999],
            no_3d: false,
        }),
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 0,
                col_offset_emu: 0,
                row: 0,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 2,
                col_offset_emu: 0,
                row: 10,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let sheet = result.worksheet(0).unwrap();
    let control = sheet.form_controls().next().expect("control survives");
    match &control.payload.kind {
        FormControlKind::ListBox {
            selection,
            selected,
            ..
        } => {
            assert_eq!(*selection, ListSelection::Multi);
            assert_eq!(selected, &vec![0, 4_999, 9_999]);
        }
        other => panic!("expected ListBox, got {other:?}"),
    }
}

/// The drawings-list z-order (picture, control, picture) survives
/// Excel's XLS re-save: BIFF8 keeps every shape in one OfficeArt
/// container whose order is z-order, so the interleave must read
/// back exactly.
#[test]
fn excel_preserves_xls_drawing_z_order_we_emit() {
    use duke_sheets_chart::{EmbeddedImage, ImageFormat};
    use duke_sheets_core::{CheckState, DrawingKind, DrawingObject, FormControl, FormControlKind};

    let two_cell = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };
    let png = |name: &str| {
        DrawingObject::image(EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 300_000,
            height_emu: 300_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        })
        .with_name(name)
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    ws.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "Middle".into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }))
        .with_anchor(two_cell(1, 1, 3, 3)),
    ).unwrap();
    ws.add_drawing(png("Above").with_anchor(two_cell(2, 2, 4, 4))).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let sheet = result.worksheet(0).unwrap();
    let tags: Vec<&str> = sheet
        .drawings()
        .iter()
        .map(|object| match &object.kind {
            DrawingKind::Image(_) => "image",
            DrawingKind::FormControl(_) => "control",
            other => panic!("unexpected drawing kind after Excel round-trip: {other:?}"),
        })
        .collect();
    assert_eq!(
        tags,
        vec!["image", "control", "image"],
        "z-order must survive Excel XLS re-save"
    );
    assert_eq!(
        sheet
            .form_controls()
            .next()
            .unwrap()
            .payload
            .caption_text()
            .as_deref(),
        Some("Middle")
    );
}

/// Drawing-object hidden flags survive Excel's XLS re-save: both the
/// hidden image and the hidden control ride the Escher FOPT Group
/// Shape Boolean Properties entry (0x03BF fHidden, MS-ODRAW
/// §2.3.4.44), and the visible siblings stay visible.
#[test]
fn excel_preserves_hidden_drawing_flags_we_emit() {
    use duke_sheets_chart::{EmbeddedImage, ImageFormat};
    use duke_sheets_core::{CheckState, DrawingObject, FormControl, FormControlKind};

    let two_cell = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };
    let png = |name: &str| {
        DrawingObject::image(EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 300_000,
            height_emu: 300_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        })
        .with_name(name)
    };
    let checkbox = |caption: &str| {
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: caption.into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }))
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_drawing(png("Shown").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    ws.add_drawing(
        png("Ghost")
            .with_anchor(two_cell(2, 2, 4, 4))
            .with_hidden(true),
    ).unwrap();
    ws.add_drawing(checkbox("Visible box").with_anchor(two_cell(4, 4, 6, 6))).unwrap();
    ws.add_drawing(
        checkbox("Cloaked box")
            .with_anchor(two_cell(6, 6, 8, 8))
            .with_hidden(true),
    ).unwrap();

    let result = roundtrip_through_excel_xls(&wb);
    let sheet = result.worksheet(0).unwrap();

    let images: Vec<_> = sheet.images().collect();
    assert_eq!(images.len(), 2, "both images survive Excel re-save");
    let image_hidden = |name: &str| {
        images
            .iter()
            .find(|i| i.object.unwrap().meta.name.as_deref() == Some(name))
            .unwrap_or_else(|| panic!("image {name:?} lost in Excel re-save"))
            .meta
            .hidden
    };
    assert!(!image_hidden("Shown"), "visible image must stay visible");
    assert!(
        image_hidden("Ghost"),
        "hidden image must survive Excel re-save with hidden intact"
    );

    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), 2, "both controls survive Excel re-save");
    let control_hidden = |caption: &str| {
        controls
            .iter()
            .find(|c| c.payload.caption_text().as_deref() == Some(caption))
            .unwrap_or_else(|| panic!("control {caption:?} lost in Excel re-save"))
            .meta
            .hidden
    };
    assert!(
        !control_hidden("Visible box"),
        "visible control must stay visible"
    );
    assert!(
        control_hidden("Cloaked box"),
        "hidden control must survive Excel re-save with hidden intact"
    );
}

#[test]
fn excel_preserves_xls_basic_shape_we_emit() {
    use duke_sheets_core::style::{HorizontalAlignment, VerticalAlignment};
    use duke_sheets_core::{
        DrawingObject, DrawingText, Shape, ShapeFill, ShapeGeometry, ShapeLine,
    };

    let text = DrawingText {
        runs: vec![
            RichTextRun::with_font(
                "Bold ",
                RunFont {
                    name: Some("Segoe UI".into()),
                    size: Some(10.0),
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
            RichTextRun::with_font(
                "Italic",
                RunFont {
                    name: Some("Arial".into()),
                    size: Some(12.0),
                    italic: Some(true),
                    color: Some(Color::rgb(0, 0, 255)),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: Some(HorizontalAlignment::Center),
        vertical_alignment: Some(VerticalAlignment::Center),
    };
    let shape = Shape::rectangle()
        .with_fill(ShapeFill::Solid(Color::rgb(255, 0, 0)))
        .with_line(ShapeLine {
            color: Some(Color::rgb(0, 0, 255)),
            width_emu: Some(25_400),
            dash_style: Some("dash".into()),
            no_fill: false,
        })
        .with_text(text)
        .with_rotation(900_000)
        .with_flip_h(true);
    let mut object = DrawingObject::shape(shape).with_anchor(DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            row: 2,
            ..CellMarker::default()
        },
        to: CellMarker {
            col: 5,
            row: 8,
            ..CellMarker::default()
        },
        edit_as: None,
    });
    object.meta.name = Some("Status panel".into());
    object.meta.alt_text = Some("red status rectangle".into());
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();

    let result = roundtrip_through_excel_xls(&workbook);
    let drawn = result.worksheet(0).unwrap().shapes().next().expect("shape");
    assert_eq!(drawn.object.unwrap().meta.name.as_deref(), Some("Status panel"));
    assert_eq!(
        drawn.object.unwrap().meta.alt_text.as_deref(),
        Some("red status rectangle")
    );
    assert_eq!(drawn.payload.geometry, ShapeGeometry::Preset("rect".into()));
    assert_eq!(drawn.payload.fill, ShapeFill::Solid(Color::rgb(255, 0, 0)));
    assert_eq!(drawn.payload.line.color, Some(Color::rgb(0, 0, 255)));
    assert_eq!(drawn.payload.line.width_emu, Some(25_400));
    assert_eq!(drawn.payload.line.dash_style.as_deref(), Some("dash"));
    assert_eq!(drawn.payload.rotation, 900_000);
    assert!(drawn.payload.flip_h);
    let text = drawn.payload.text.as_ref().expect("shape text");
    assert_eq!(text.plain_text(), "Bold Italic");
    assert_eq!(text.horizontal_alignment, Some(HorizontalAlignment::Center));
    assert_eq!(text.vertical_alignment, Some(VerticalAlignment::Center));
    assert_eq!(
        text.runs[0].font.as_ref().unwrap().name.as_deref(),
        Some("Segoe UI")
    );
    assert_eq!(text.runs[0].font.as_ref().unwrap().bold, Some(true));
    assert_eq!(text.runs[0].font.as_ref().unwrap().size, Some(10.0));
    assert_eq!(
        text.runs[1].font.as_ref().unwrap().name.as_deref(),
        Some("Arial")
    );
    assert_eq!(text.runs[1].font.as_ref().unwrap().italic, Some(true));
    assert_eq!(text.runs[1].font.as_ref().unwrap().size, Some(12.0));
    assert_eq!(
        text.runs[1].font.as_ref().unwrap().color,
        Some(Color::rgb(0, 0, 255))
    );
}
