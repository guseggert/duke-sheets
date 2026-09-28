//! Excel COM parity tests for XLSB pivot table writing.

use crate::{
    ensure_vm_temp_dir, excel_bridge, roundtrip_through_excel_xlsb_bytes,
};
use crate::writing_xlsb::xlsb_olap_connection_workbook;
use duke_sheets_core::table::{Table, TableColumn};
use duke_sheets_core::{
    CellRange, PivotAggregate, PivotCacheSourceKind, PivotDateGroupUnit, PivotDatePeriod,
    PivotField, PivotFieldRef, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayout,
    PivotLayoutKind, PivotManualGroup, PivotMeasure, PivotRefreshPolicy, PivotShowAs, PivotSort,
    PivotSource, PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable, PivotValue,
    PivotValuesAxis, Workbook, WorkbookConnection, WorkbookConnectionKind,
};

const XLSB_EXTERNAL_CONSOLIDATION_SOURCE_TARGET: &str = "pivot_external_source.xlsx";
const XLSB_EXTERNAL_CONSOLIDATION_SOURCE_VM_PATH: &str = r"C:\temp\pivot_external_source.xlsx";
const XLSB_EXTERNAL_PIVOT_CSV_VM_PATH: &str = r"C:\temp\pivot_external_sales.csv";

fn xlsb_basic_pivot_workbook() -> Workbook {
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

fn xlsb_table_source_pivot_workbook() -> Workbook {
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

fn xlsb_consolidation_source_pivot_workbook() -> Workbook {
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
                PivotSourceRange::new("Sheet1", CellRange::parse("A1:B3").unwrap())
                    .with_page_items(["Retail"]),
                PivotSourceRange::new("WestData", CellRange::parse("A1:B2").unwrap())
                    .with_page_items(["Wholesale"]),
            ],
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

fn xlsb_named_consolidation_source_pivot_workbook() -> Workbook {
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

fn xlsb_external_consolidation_source_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let pivot = PivotTable::builder("ExternalConsolidatedPivot")
        .source(PivotSource::Consolidation {
            ranges: vec![
                PivotSourceRange::new("Sheet1", CellRange::parse("A1:B3").unwrap())
                    .with_external_relationship_id("rIdExternalSource")
                    .with_external_relationship_target(XLSB_EXTERNAL_CONSOLIDATION_SOURCE_TARGET),
            ],
        })
        .target_address("D1")
        .unwrap()
        .row("Row")
        .column("Column")
        .named_measure("Value", PivotAggregate::Sum, "Total Value")
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn seed_xlsb_external_consolidation_source_workbook() {
    ensure_vm_temp_dir();
    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let _ = excel.delete_file(XLSB_EXTERNAL_CONSOLIDATION_SOURCE_VM_PATH);
    let source = excel
        .create_workbook()
        .expect("create external source workbook");
    source.set_cell_value("A1", "Region").unwrap();
    source.set_cell_value("B1", "Revenue").unwrap();
    source.set_cell_value("A2", "East").unwrap();
    source.set_cell_value("B2", 10.0).unwrap();
    source.set_cell_value("A3", "West").unwrap();
    source.set_cell_value("B3", 20.0).unwrap();
    source
        .save_as(XLSB_EXTERNAL_CONSOLIDATION_SOURCE_VM_PATH, 51)
        .expect("save external source workbook");
    source.close().expect("close external source workbook");
}

fn xlsb_external_connection_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let connection = WorkbookConnection::database(
        1,
        "ExternalSalesText",
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\temp;Extended Properties=\"text;HDR=Yes;FMT=Delimited\";",
    )
    .with_command("select * from [pivot_external_sales.csv]")
    .with_command_type(2)
    .with_connection_type(5)
    .with_save_data(false);
    wb.add_data_connection(connection).unwrap();

    let pivot = PivotTable::builder("ExternalSales")
        .source(PivotSource::External {
            connection_name: "ExternalSalesText".to_string(),
            command_text: Some("select * from [pivot_external_sales.csv]".to_string()),
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

fn xlsb_external_connection_background_query_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let connection = WorkbookConnection::database(
        1,
        "ExternalSalesText",
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\temp;Extended Properties=\"text;HDR=Yes;FMT=Delimited\";",
    )
    .with_command("select * from [pivot_external_sales.csv]")
    .with_command_type(2)
    .with_connection_type(5)
    .with_save_data(false);
    wb.add_data_connection(connection).unwrap();

    let refresh_policy = PivotRefreshPolicy {
        refresh_on_open: false,
        preserve_formatting: true,
        background_query: true,
        missing_items_limit: None,
    };
    let pivot = PivotTable::builder("ExternalSalesBackground")
        .source(PivotSource::External {
            connection_name: "ExternalSalesText".to_string(),
            command_text: Some("select * from [pivot_external_sales.csv]".to_string()),
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .refresh_policy(refresh_policy)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_olap_pivot_source_workbook() -> Workbook {
    let mut wb = xlsb_olap_connection_workbook();
    let pivot = PivotTable::builder("OlapSales")
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
    wb
}

fn xlsb_scenario_source_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let pivot = PivotTable::builder("ScenarioSales")
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

fn xlsb_refresh_policy_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let refresh_policy = PivotRefreshPolicy {
        refresh_on_open: true,
        preserve_formatting: true,
        background_query: false,
        missing_items_limit: Some(5),
    };
    let pivot = PivotTable::builder("RefreshPolicyPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .refresh_policy(refresh_policy)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_no_preserve_formatting_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let refresh_policy = PivotRefreshPolicy {
        refresh_on_open: false,
        preserve_formatting: false,
        background_query: false,
        missing_items_limit: None,
    };
    let pivot = PivotTable::builder("NoPreserveFormatting")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .refresh_policy(refresh_policy)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_layout_flags_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Segment").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Retail").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "Wholesale").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Outline;
    layout.show_row_grand_totals = false;
    layout.show_column_grand_totals = false;
    layout.show_field_headers = false;
    layout.show_expand_collapse = false;
    layout.print_drill_indicators = true;
    layout.item_print_titles = true;
    layout.field_print_titles = true;
    layout.page_wrap = 2;
    layout.page_over_then_down = true;
    layout.merge_item_labels = true;
    layout.grand_total_caption = Some("Grand".to_string());
    layout.error_caption = Some("ERR".to_string());
    layout.show_error = true;
    layout.missing_caption = Some("MISSING".to_string());
    layout.show_missing = true;
    layout.show_items = false;
    layout.edit_data = true;
    layout.disable_field_list = true;
    layout.show_calculated_members = false;
    layout.visual_totals = false;
    layout.show_multiple_label = false;
    layout.show_data_drop_down = false;
    layout.show_member_property_tips = false;
    layout.show_data_tips = false;
    layout.enable_wizard = false;
    layout.enable_drill = false;
    layout.enable_field_properties = false;
    layout.indent = 3;
    layout.show_empty_rows = true;
    layout.show_empty_columns = true;

    let pivot = PivotTable::builder("LayoutFlags")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E2")
        .unwrap()
        .row("Region")
        .page("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .layout(layout)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

fn seed_xlsb_external_pivot_csv() {
    ensure_vm_temp_dir();
    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    excel
        .put_file(
            XLSB_EXTERNAL_PIVOT_CSV_VM_PATH,
            b"Region,Revenue\r\nEast,10\r\nWest,20\r\n",
        )
        .expect("seed external pivot CSV");
}

fn xlsb_column_pivot_workbook() -> Workbook {
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

fn xlsb_page_pivot_workbook() -> Workbook {
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
        .page("Salesperson")
        .filter(PivotFilter::field_items("Salesperson", ["Ada"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_multi_select_page_pivot_workbook() -> Workbook {
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
        .page("Region")
        .row("Salesperson")
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_top_n_pivot_workbook() -> Workbook {
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

fn xlsb_label_contains_pivot_workbook() -> Workbook {
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

    let pivot = PivotTable::builder("LabelRegions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::Label {
            field: PivotFieldRef::new("Region"),
            operator: PivotFilterOperator::Contains,
            value: "Ea".to_string(),
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_label_equals_pivot_workbook() -> Workbook {
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

fn xlsb_label_prefix_suffix_pivot_workbook(
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

fn xlsb_label_between_pivot_workbook(name: &str, not_between: bool) -> Workbook {
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

fn xlsb_value_filter_pivot_workbook(name: &str, filter: PivotFilter) -> Workbook {
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

fn xlsb_date_filter_pivot_workbook(name: &str, filter: PivotFilter) -> Workbook {
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

fn xlsb_column_value_filter_pivot_workbook() -> Workbook {
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

fn xlsb_second_measure_value_filter_pivot_workbook() -> Workbook {
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

fn xlsb_row_item_filter_pivot_workbook() -> Workbook {
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
        .row("Region")
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_column_item_filter_pivot_workbook() -> Workbook {
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

    let pivot = PivotTable::builder("ColumnItemFilter")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .column("Region")
        .filter(PivotFilter::field_items("Region", ["East", "West"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_source_only_item_filter_pivot_workbook() -> Workbook {
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

fn xlsb_grouped_row_item_filter_pivot_workbook() -> Workbook {
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
        .row("Region")
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

fn xlsb_grouped_column_item_filter_pivot_workbook() -> Workbook {
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
        .column("Region")
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

fn xlsb_axis_options_pivot_workbook() -> Workbook {
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
        .with_subtotal_caption("? subtotal")
        .with_subtotals([
            PivotSubtotal::Sum,
            PivotSubtotal::Count,
            PivotSubtotal::Average,
        ]);
    region.sort = PivotSort::Descending;
    region.show_empty_items = true;
    region.show_drop_downs = false;
    region.subtotal_top = false;
    region.insert_blank_row = true;
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

fn xlsb_sort_by_measure_pivot_workbook() -> Workbook {
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
    let pivot = PivotTable::builder("ValueSortedPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Sum of Revenue")
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_styled_pivot_style() -> PivotStyle {
    PivotStyle {
        name: Some("PivotStyleLight16".to_string()),
        show_row_headers: false,
        show_column_headers: true,
        show_row_stripes: true,
        show_column_stripes: true,
        show_last_column: true,
    }
}

fn xlsb_styled_pivot_workbook() -> Workbook {
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
        .style(xlsb_styled_pivot_style())
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_calculated_field_pivot_workbook() -> Workbook {
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

fn xlsb_calculated_item_pivot_workbook() -> Workbook {
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

fn xlsb_calculated_item_cell_like_pivot_workbook() -> Workbook {
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

fn xlsb_calculated_item_string_ref_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
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
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_calculated_field_function_pivot_workbook() -> Workbook {
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

fn xlsb_calculated_item_function_pivot_workbook() -> Workbook {
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

fn xlsb_multi_measure_pivot_workbook() -> Workbook {
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

fn xlsb_all_aggregate_cases() -> [(PivotAggregate, &'static str, u32); 11] {
    [
        (PivotAggregate::Sum, "Sum Revenue", 0x00),
        (PivotAggregate::Count, "Count Revenue", 0x01),
        (PivotAggregate::Average, "Average Revenue", 0x02),
        (PivotAggregate::Max, "Max Revenue", 0x03),
        (PivotAggregate::Min, "Min Revenue", 0x04),
        (PivotAggregate::Product, "Product Revenue", 0x05),
        (PivotAggregate::CountNumbers, "Count Numbers Revenue", 0x06),
        (PivotAggregate::StdDev, "StdDev Revenue", 0x07),
        (PivotAggregate::StdDevP, "StdDevP Revenue", 0x08),
        (PivotAggregate::Var, "Var Revenue", 0x09),
        (PivotAggregate::VarP, "VarP Revenue", 0x0A),
    ]
}

fn xlsb_all_aggregate_pivot_workbook() -> Workbook {
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
    for (aggregate, caption, _) in xlsb_all_aggregate_cases() {
        builder = builder.named_measure("Revenue", aggregate, caption);
    }
    let mut pivot = builder.build().unwrap();
    pivot.layout.values_axis = PivotValuesAxis::Columns;
    pivot.layout.values_axis_position = Some(0);
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_show_as_pivot_workbook() -> Workbook {
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
        .column("Quarter")
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

fn xlsb_custom_measure_format_pivot_workbook() -> Workbook {
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

fn xlsb_x14_show_as_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Rep").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "Ada").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "East").unwrap();
    ws.set_cell_value("B3", "Ben").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();
    ws.set_cell_value("A4", "West").unwrap();
    ws.set_cell_value("B4", "Ada").unwrap();
    ws.set_cell_value("C4", 30.0).unwrap();
    ws.set_cell_value("A5", "West").unwrap();
    ws.set_cell_value("B5", "Ben").unwrap();
    ws.set_cell_value("C5", 40.0).unwrap();

    let mut pivot = PivotTable::builder("X14ShowAsPivot")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Rep")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Pct Parent Row")
                .with_show_as(PivotShowAs::PercentOfParentRowTotal),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Pct Parent Col")
                .with_show_as(PivotShowAs::PercentOfParentColumnTotal),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Pct Parent Rep")
                .with_show_as(PivotShowAs::PercentOfParentTotal {
                    base_field: PivotFieldRef::new("Rep"),
                }),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Rank Rep Asc")
                .with_show_as(PivotShowAs::RankAscending {
                    base_field: PivotFieldRef::new("Rep"),
                }),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Rank Region Desc")
                .with_show_as(PivotShowAs::RankDescending {
                    base_field: PivotFieldRef::new("Region"),
                }),
        )
        .build()
        .unwrap();
    pivot.layout.values_axis = PivotValuesAxis::Columns;
    pivot.layout.values_axis_position = Some(0);
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_numeric_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Age").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 5.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 12.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", 23.0).unwrap();
    ws.set_cell_value("B4", 30.0).unwrap();
    ws.set_cell_value("A5", 41.0).unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder("GroupedAges")
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

fn xlsb_date_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
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
            field: "Date".into(),
            units: vec![PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_page_date_grouped_pivot_workbook(allowed_items: &[f64]) -> Workbook {
    let mut wb = Workbook::new();
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

    let mut pivot = PivotTable::builder("MonthlyPageFilter")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E2")
        .unwrap()
        .page("Date")
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: "Date".into(),
            units: vec![PivotDateGroupUnit::Months],
        });
    if !allowed_items.is_empty() {
        pivot = pivot.filter(PivotFilter::field_items(
            "Date",
            allowed_items.iter().copied(),
        ));
    }
    let pivot = pivot.build().unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_multi_unit_date_grouped_pivot_workbook() -> Workbook {
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

fn xlsb_multi_unit_date_grouped_page_pivot_workbook() -> Workbook {
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

    let pivot = PivotTable::builder("GroupedDatePage")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .page("SaleDate")
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

fn xlsb_multi_unit_date_grouped_page_filter_pivot_workbook() -> Workbook {
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
    ws.set_cell_value("A5", 45717.0).unwrap();
    ws.set_cell_value("B5", 40.0).unwrap();

    let pivot = PivotTable::builder("GroupedDatePage")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .page("SaleDate")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Date {
            field: "SaleDate".into(),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .filter(PivotFilter::field_items("SaleDate (Years)", [2024.0]))
        .filter(PivotFilter::field_items("SaleDate (Months)", [1.0, 2.0]))
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_manual_grouped_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    ws.set_cell_value("A4", "Central").unwrap();
    ws.set_cell_value("B4", 5.0).unwrap();

    let pivot = PivotTable::builder("ManualGroupedRegions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

fn xlsb_manual_page_grouped_pivot_workbook(allowed_items: &[&str]) -> Workbook {
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
        .page("Region");
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
fn excel_opens_xlsb_with_native_pivot_table() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_basic_pivot_workbook());

    assert!(zip_has_entry(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin"
    ));
    assert!(zip_has_entry(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin"
    ));
    assert!(zip_has_entry(
        &excel_bytes,
        "xl/pivotCache/pivotCacheRecords1.bin"
    ));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_table_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_table_source_pivot_workbook());

    assert!(zip_has_entry(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin"
    ));
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TablePivot")
        .unwrap();
    assert!(matches!(
        &pivot.source,
        PivotSource::Table { name } if name == "SalesData"
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert!(pivot.columns.is_empty());
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_consolidation_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_consolidation_source_pivot_workbook());

    let cache_records = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    assert!(cache_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCDS_CONSOL));
    assert!(cache_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCDSC_SET));

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ConsolidatedPivot")
        .unwrap();
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
fn excel_preserves_xlsb_pivot_named_consolidation_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_named_consolidation_source_pivot_workbook());

    let cache_records = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    assert!(cache_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCDS_CONSOL));
    assert!(cache_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCDSC_SET));

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("NamedConsolidatedPivot")
        .unwrap();
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
fn excel_preserves_xlsb_pivot_external_consolidation_source() {
    seed_xlsb_external_consolidation_source_workbook();
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_external_consolidation_source_pivot_workbook());

    let cache_records = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    assert!(cache_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCDS_CONSOL));
    assert!(cache_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCDSC_SET));

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ExternalConsolidatedPivot")
        .unwrap();
    let PivotSource::Consolidation { ranges } = &pivot.source else {
        panic!("expected consolidation source after Excel re-save");
    };
    assert_eq!(ranges.len(), 1);
    assert!(ranges[0].external_relationship_id.is_some());
    assert_eq!(
        ranges[0].external_relationship_target.as_deref(),
        Some(XLSB_EXTERNAL_CONSOLIDATION_SOURCE_TARGET)
    );
    assert_eq!(pivot.rows[0].field.name, "Row");
    assert_eq!(pivot.columns[0].field.name, "Column");
    assert_eq!(pivot.measures[0].field.name, "Value");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_external_connection_source() {
    seed_xlsb_external_pivot_csv();
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_external_connection_pivot_workbook());

    assert!(zip_has_entry(&excel_bytes, "xl/connections.bin"));
    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let source_payload = cache_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCD_SOURCE)
                .then_some(payload)
        })
        .expect("external pivot cache source header after Excel re-save");
    assert_eq!(&source_payload[..4], &1u32.to_le_bytes());
    assert_eq!(&source_payload[4..8], &1u32.to_le_bytes());

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ExternalSales")
        .unwrap();
    assert!(matches!(
        &pivot.source,
        PivotSource::External {
            connection_name,
            command_text,
        } if connection_name == "ExternalSalesText"
            && command_text.as_deref() == Some("select * from [pivot_external_sales.csv]")
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Revenue"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_external_pivot_cache_background_query() {
    seed_xlsb_external_pivot_csv();
    let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xlsb_bytes(
        &xlsb_external_connection_background_query_pivot_workbook(),
    );

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let cache_def = cache_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PIVOT_CACHE_DEF)
                .then_some(payload)
        })
        .expect("pivot cache definition after Excel re-save");
    assert_ne!(cache_def[3] & 0x20, 0);

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ExternalSalesBackground")
        .unwrap();
    assert!(pivot.refresh_policy.background_query);
    assert!(matches!(
        &pivot.source,
        PivotSource::External {
            connection_name,
            command_text,
        } if connection_name == "ExternalSalesText"
            && command_text.as_deref() == Some("select * from [pivot_external_sales.csv]")
    ));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_olap_source_metadata() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_olap_pivot_source_workbook());

    assert!(zip_has_entry(&excel_bytes, "xl/connections.bin"));
    let connection = result
        .data_connection_by_name("CubeSales")
        .expect("CubeSales connection after Excel re-save");
    assert!(matches!(
        &connection.kind,
        WorkbookConnectionKind::Olap { .. }
    ));

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let source_payload = cache_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCD_SOURCE)
                .then_some(payload)
        })
        .expect("OLAP pivot cache source header after Excel re-save");
    assert_eq!(&source_payload[..4], &1u32.to_le_bytes());
    assert_eq!(&source_payload[4..8], &connection.id.to_le_bytes());

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("OlapSales")
        .unwrap();
    assert!(matches!(
        &pivot.source,
        PivotSource::Olap {
            connection_name,
            cube,
            command_text,
        } if connection_name == "CubeSales"
            && cube.as_deref() == Some("SalesCube")
            && command_text.as_deref() == Some("SalesCube")
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Sales");
    assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Sales"));
    assert!(matches!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(PivotCacheSourceKind::Olap)
    ));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_unnamed_scenario_source() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_scenario_source_pivot_workbook());

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let source_payload = cache_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCD_SOURCE)
                .then_some(payload)
        })
        .expect("scenario pivot cache source header after Excel re-save");
    assert_eq!(&source_payload[..4], &3u32.to_le_bytes());
    assert_eq!(&source_payload[4..8], &0u32.to_le_bytes());

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ScenarioSales")
        .unwrap();
    assert!(matches!(&pivot.source, PivotSource::Scenario { name } if name.is_empty()));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Revenue"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_cache_refresh_policy() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_refresh_policy_pivot_workbook());

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let cache_def = cache_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PIVOT_CACHE_DEF)
                .then_some(payload)
        })
        .expect("pivot cache definition after Excel re-save");
    assert_ne!(cache_def[3] & 0x04, 0);
    assert_eq!(i32::from_le_bytes(cache_def[4..8].try_into().unwrap()), 5);

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RefreshPolicyPivot")
        .unwrap();
    assert!(pivot.refresh_policy.refresh_on_open);
    assert_eq!(pivot.refresh_policy.missing_items_limit, Some(5));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_preserve_formatting_policy() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_no_preserve_formatting_pivot_workbook());

    let pivot_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let sx_view = pivot_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXVIEW).then_some(payload)
        })
        .expect("pivot view after Excel re-save");
    assert_eq!(sx_view[4] & 0x80, 0);

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("NoPreserveFormatting")
        .unwrap();
    assert!(!pivot.refresh_policy.preserve_formatting);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_table_layout_flags() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_layout_flags_pivot_workbook());

    let pivot_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let sx_view = pivot_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXVIEW).then_some(payload)
        })
        .expect("pivot view after Excel re-save");
    assert_eq!(sx_view[13], 2, "pageWrap should survive Excel re-save");
    assert!(!xlsb_sxview_flag(sx_view, 13), "rowGrandTotals");
    assert!(!xlsb_sxview_flag(sx_view, 14), "colGrandTotals");
    assert!(xlsb_sxview_flag(sx_view, 15), "fieldPrintTitles");
    assert!(xlsb_sxview_flag(sx_view, 17), "itemPrintTitles");
    assert!(xlsb_sxview_flag(sx_view, 18), "mergeItem");
    assert!(xlsb_sxview_flag(sx_view, 20), "grandTotalCaption present");
    assert!(!xlsb_sxview_flag(sx_view, 38), "errorCaption present");
    assert!(!xlsb_sxview_flag(sx_view, 39), "missingCaption present");

    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("LayoutFlags")
        .unwrap();
    assert_eq!(pivot.layout.kind, PivotLayoutKind::Outline);
    assert!(!pivot.layout.show_row_grand_totals);
    assert!(!pivot.layout.show_column_grand_totals);
    assert!(!pivot.layout.show_field_headers);
    assert!(!pivot.layout.show_expand_collapse);
    assert!(pivot.layout.print_drill_indicators);
    assert!(pivot.layout.item_print_titles);
    assert!(pivot.layout.field_print_titles);
    assert_eq!(pivot.layout.page_wrap, 2);
    assert!(pivot.layout.page_over_then_down);
    assert!(pivot.layout.merge_item_labels);
    assert_eq!(pivot.layout.grand_total_caption.as_deref(), Some("Grand"));
    assert_eq!(pivot.layout.error_caption.as_deref(), Some("ERR"));
    assert!(pivot.layout.show_error);
    assert_eq!(pivot.layout.missing_caption.as_deref(), Some("MISSING"));
    assert!(pivot.layout.show_missing);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_column_axis() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_column_pivot_workbook());

    let pivot_records = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    assert!(pivot_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_ISXVD_COLS));
    assert!(pivot_records.contains(&duke_sheets_xlsb::biff12::records::BRT_END_ISXVD_COLS));
    assert!(pivot_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_COL_ITEMS));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_page_axis() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_page_pivot_workbook());

    let pivot_records = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    assert!(pivot_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPIS));
    assert!(pivot_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPI));
    assert!(pivot_records.contains(&duke_sheets_xlsb::biff12::records::BRT_END_SXPIS));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_multi_item_page_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_multi_select_page_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("MultiPage")
        .expect("MultiPage pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let page_field = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
        })
        .expect("BrtBeginSXPI after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(page_field[4..8].try_into().unwrap()),
        0x0010_00FE,
        "Excel should preserve the BIFF12 multiple-items page-filter sentinel"
    );

    let page_items = sxvi_payloads_for_xlsb_sxvd_axis(&records, 0x04);
    assert_eq!(
        page_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Excel should keep only Central hidden for the multi-select page filter"
    );
    assert_eq!(
        i32::from_le_bytes(page_items[2][3..7].try_into().unwrap()),
        2
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_top_n_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_top_n_pivot_workbook());
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

    let record_types = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    assert!(record_types.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTERS));
    assert!(record_types.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER));
    assert!(record_types.contains(&duke_sheets_xlsb::biff12::records::BRT_TOP10_FILTER));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_label_contains_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_label_contains_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("LabelRegions")
        .expect("LabelRegions pivot after Excel re-save");

    assert_eq!(pivot.filters.len(), 1);
    match &pivot.filters[0] {
        PivotFilter::Label {
            field,
            operator,
            value,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(*operator, PivotFilterOperator::Contains);
            assert_eq!(value, "Ea");
        }
        other => panic!("unexpected pivot filter after Excel re-save: {other:?}"),
    }

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    assert!(records.iter().any(|(record_type, _)| {
        *record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTERS
    }));

    let filter_header = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                .then_some(payload)
        })
        .expect("BrtBeginSXFilter after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
        0
    );
    assert_eq!(
        u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
        10
    );
    assert_eq!(
        u16::from_le_bytes(filter_header[28..30].try_into().unwrap()),
        4
    );
    assert_eq!(wide_string_at(filter_header, 30).0, "Ea");

    let custom_filter = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                .then_some(payload)
        })
        .expect("BrtCustomFilter after Excel re-save");
    assert_eq!(custom_filter[0], 6, "custom filter value type");
    assert_eq!(custom_filter[1], 2, "contains uses equality over wildcards");
    assert_eq!(wide_string_at(custom_filter, 10).0, "*Ea*");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_label_equals_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_label_equals_pivot_workbook());
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

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let filter_header = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                .then_some(payload)
        })
        .expect("BrtBeginSXFilter after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
        0
    );
    assert_eq!(
        u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
        4
    );
    assert_eq!(
        u16::from_le_bytes(filter_header[28..30].try_into().unwrap()),
        4
    );
    assert_eq!(wide_string_at(filter_header, 30).0, "East");
    assert!(
        !records.iter().any(|(record_type, _)| {
            *record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER
        }),
        "caption-equals filters should survive Excel without a wildcard custom filter"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_label_prefix_suffix_filters() {
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
        let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xlsb_bytes(
            &xlsb_label_prefix_suffix_pivot_workbook(pivot_name, operator, value),
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

        let records = xlsb_record_payloads(&zip_entry_bytes(
            &excel_bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                    .then_some(payload)
            })
            .expect("BrtBeginSXFilter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
            0
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            u16::from_le_bytes(filter_header[28..30].try_into().unwrap()),
            4
        );
        assert_eq!(wide_string_at(filter_header, 30).0, value);

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                    .then_some(payload)
            })
            .expect("BrtCustomFilter after Excel re-save");
        assert_eq!(custom_filter[0], 6, "custom filter value type");
        assert_eq!(custom_filter[1], 2, "wildcard label filters use equality");
        assert_eq!(wide_string_at(custom_filter, 10).0, custom_value);
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_negative_label_filters() {
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
        let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xlsb_bytes(
            &xlsb_label_prefix_suffix_pivot_workbook(pivot_name, operator, value),
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

        let records = xlsb_record_payloads(&zip_entry_bytes(
            &excel_bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                    .then_some(payload)
            })
            .expect("BrtBeginSXFilter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
            0
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            filter_type
        );
        assert_eq!(wide_string_at(filter_header, 30).0, value);

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                    .then_some(payload)
            })
            .expect("BrtCustomFilter after Excel re-save");
        assert_eq!(custom_filter[0], 6, "custom filter value type");
        assert_eq!(custom_filter[1], 5, "negative label filters use notEqual");
        assert_eq!(custom_filter[2], discriminator);
        assert_eq!(wide_string_at(custom_filter, 10).0, custom_value);
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_label_comparison_filters() {
    for (operator, filter_type, custom_operator, pivot_name) in [
        (PivotFilterOperator::GreaterThan, 12u32, 4u8, "LabelGreater"),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            13u32,
            6u8,
            "LabelGreaterEqual",
        ),
        (PivotFilterOperator::LessThan, 14u32, 1u8, "LabelLess"),
        (
            PivotFilterOperator::LessThanOrEqual,
            15u32,
            3u8,
            "LabelLessEqual",
        ),
    ] {
        let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xlsb_bytes(
            &xlsb_label_prefix_suffix_pivot_workbook(pivot_name, operator, "M"),
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

        let records = xlsb_record_payloads(&zip_entry_bytes(
            &excel_bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                    .then_some(payload)
            })
            .expect("BrtBeginSXFilter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
            0
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            filter_type
        );
        assert_eq!(wide_string_at(filter_header, 30).0, "M");

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                    .then_some(payload)
            })
            .expect("BrtCustomFilter after Excel re-save");
        assert_eq!(custom_filter[0], 6, "custom filter value type");
        assert_eq!(custom_filter[1], custom_operator);
        assert_eq!(custom_filter[2], 1, "comparison filters use raw criteria");
        assert_eq!(wide_string_at(custom_filter, 10).0, "M");
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_label_between_filters() {
    for (not_between, filter_type, and_flag, first_operator, second_operator, pivot_name) in [
        (false, 16u32, 1i32, 6u8, 3u8, "LabelBetween"),
        (true, 17u32, 0i32, 1u8, 4u8, "LabelNotBetween"),
    ] {
        let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xlsb_bytes(
            &xlsb_label_between_pivot_workbook(pivot_name, not_between),
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

        let records = xlsb_record_payloads(&zip_entry_bytes(
            &excel_bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                    .then_some(payload)
            })
            .expect("BrtBeginSXFilter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
            0
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            filter_type
        );
        assert_eq!(wide_string_at(filter_header, 30).0, "East");

        let custom_filters_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_CUSTOM_FILTERS)
                    .then_some(payload)
            })
            .expect("BrtBeginCustomFilters after Excel re-save");
        assert_eq!(
            i32::from_le_bytes(custom_filters_header[0..4].try_into().unwrap()),
            and_flag
        );

        let custom_filters = records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                    .then_some(payload)
            })
            .collect::<Vec<_>>();
        assert_eq!(custom_filters.len(), 2);
        assert_eq!(custom_filters[0][0], 6, "label filter value type");
        assert_eq!(custom_filters[0][1], first_operator);
        assert_eq!(custom_filters[0][2], 1);
        assert_eq!(wide_string_at(custom_filters[0], 10).0, "East");
        assert_eq!(custom_filters[1][0], 6, "label filter value type");
        assert_eq!(custom_filters[1][1], second_operator);
        assert_eq!(custom_filters[1][2], 1);
        assert_eq!(wide_string_at(custom_filters[1], 10).0, "West");
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_value_comparison_filters() {
    for (operator, filter_type, custom_operator, threshold, pivot_name) in [
        (PivotFilterOperator::Equals, 18u32, 2u8, 20.0, "ValueEquals"),
        (
            PivotFilterOperator::NotEquals,
            19u32,
            5u8,
            20.0,
            "ValueNotEquals",
        ),
        (
            PivotFilterOperator::GreaterThan,
            20u32,
            4u8,
            20.0,
            "ValueGreater",
        ),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            21u32,
            6u8,
            20.0,
            "ValueGreaterEqual",
        ),
        (PivotFilterOperator::LessThan, 22u32, 1u8, 30.0, "ValueLess"),
        (
            PivotFilterOperator::LessThanOrEqual,
            23u32,
            3u8,
            30.0,
            "ValueLessEqual",
        ),
    ] {
        let filter_measure =
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
        let (result, _writer_bytes, excel_bytes) =
            roundtrip_through_excel_xlsb_bytes(&xlsb_value_filter_pivot_workbook(
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

        let records = xlsb_record_payloads(&zip_entry_bytes(
            &excel_bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                    .then_some(payload)
            })
            .expect("BrtBeginSXFilter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
            0
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
            0,
            "Excel should preserve the value-filter data field index"
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                    .then_some(payload)
            })
            .expect("BrtCustomFilter after Excel re-save");
        assert_eq!(custom_filter[0], 4, "custom filter should store a number");
        assert_eq!(custom_filter[1], custom_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[2..10].try_into().unwrap()),
            threshold
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_column_value_filter() {
    let filter_measure =
        PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_column_value_filter_pivot_workbook());
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

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let filter_header = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                .then_some(payload)
        })
        .expect("BrtBeginSXFilter after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
        0,
        "Excel should preserve the column-axis source field index"
    );
    assert_eq!(
        u32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
        0,
        "Excel should preserve the column value-filter data field index"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_second_measure_value_filter() {
    let filter_measure = PivotMeasure::new("Cost", PivotAggregate::Sum).with_name("Total Cost");
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_second_measure_value_filter_pivot_workbook());
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

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let filter_header = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                .then_some(payload)
        })
        .expect("BrtBeginSXFilter after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
        1,
        "Excel should preserve the second data field as the value-filter target"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_value_between_filters() {
    for (not_between, filter_type, and_flag, first_operator, second_operator, pivot_name) in [
        (false, 24u32, 1i32, 6u8, 3u8, "ValueBetween"),
        (true, 25u32, 0i32, 1u8, 4u8, "ValueNotBetween"),
    ] {
        let filter_measure =
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
        let (result, _writer_bytes, excel_bytes) =
            roundtrip_through_excel_xlsb_bytes(&xlsb_value_filter_pivot_workbook(
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

        let records = xlsb_record_payloads(&zip_entry_bytes(
            &excel_bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                    .then_some(payload)
            })
            .expect("BrtBeginSXFilter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
            0,
            "Excel should preserve the value-filter data field index"
        );

        let custom_filters_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_CUSTOM_FILTERS)
                    .then_some(payload)
            })
            .expect("BrtBeginCustomFilters after Excel re-save");
        assert_eq!(
            i32::from_le_bytes(custom_filters_header[0..4].try_into().unwrap()),
            and_flag
        );

        let custom_filters = records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                    .then_some(payload)
            })
            .collect::<Vec<_>>();
        assert_eq!(custom_filters.len(), 2);
        assert_eq!(custom_filters[0][0], 4);
        assert_eq!(custom_filters[0][1], first_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filters[0][2..10].try_into().unwrap()),
            15.0
        );
        assert_eq!(custom_filters[1][0], 4);
        assert_eq!(custom_filters[1][1], second_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filters[1][2..10].try_into().unwrap()),
            35.0
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_date_comparison_filters() {
    for (operator, filter_type, custom_operator, threshold, pivot_name) in [
        (
            PivotFilterOperator::Equals,
            26u32,
            0x02u8,
            44958.0,
            "DateEquals",
        ),
        (
            PivotFilterOperator::NotEquals,
            62u32,
            0x05u8,
            44958.0,
            "DateNotEquals",
        ),
        (
            PivotFilterOperator::GreaterThan,
            28u32,
            0x04u8,
            44958.0,
            "DateGreater",
        ),
        (
            PivotFilterOperator::GreaterThanOrEqual,
            64u32,
            0x06u8,
            44958.0,
            "DateGreaterEqual",
        ),
        (
            PivotFilterOperator::LessThan,
            27u32,
            0x01u8,
            44986.0,
            "DateLess",
        ),
        (
            PivotFilterOperator::LessThanOrEqual,
            63u32,
            0x03u8,
            44986.0,
            "DateLessEqual",
        ),
    ] {
        let (result, _writer_bytes, excel_bytes) =
            roundtrip_through_excel_xlsb_bytes(&xlsb_date_filter_pivot_workbook(
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

        let records = xlsb_record_payloads(&zip_entry_bytes(
            &excel_bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                    .then_some(payload)
            })
            .expect("BrtBeginSXFilter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
            0,
            "Excel should preserve the date source field index"
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
            -1,
            "date filters should not target a data field"
        );

        let custom_filter = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                    .then_some(payload)
            })
            .expect("BrtCustomFilter after Excel re-save");
        assert_eq!(
            custom_filter[0], 4,
            "custom date filter should store a number"
        );
        assert_eq!(custom_filter[1], custom_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filter[2..10].try_into().unwrap()),
            threshold
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_date_between_filters() {
    for (not_between, filter_type, and_flag, first_operator, second_operator, pivot_name) in [
        (false, 29u32, 1i32, 0x06u8, 0x03u8, "DateBetween"),
        (true, 65u32, 0i32, 0x01u8, 0x04u8, "DateNotBetween"),
    ] {
        let (result, _writer_bytes, excel_bytes) =
            roundtrip_through_excel_xlsb_bytes(&xlsb_date_filter_pivot_workbook(
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

        let records = xlsb_record_payloads(&zip_entry_bytes(
            &excel_bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                    .then_some(payload)
            })
            .expect("BrtBeginSXFilter after Excel re-save");
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            filter_type
        );
        assert_eq!(
            i32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
            -1
        );

        let custom_filters_header = records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_CUSTOM_FILTERS)
                    .then_some(payload)
            })
            .expect("BrtBeginCustomFilters after Excel re-save");
        assert_eq!(
            i32::from_le_bytes(custom_filters_header[0..4].try_into().unwrap()),
            and_flag
        );

        let custom_filters = records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == duke_sheets_xlsb::biff12::records::BRT_CUSTOM_FILTER)
                    .then_some(payload)
            })
            .collect::<Vec<_>>();
        assert_eq!(custom_filters.len(), 2);
        assert_eq!(custom_filters[0][0], 4);
        assert_eq!(custom_filters[0][1], first_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filters[0][2..10].try_into().unwrap()),
            44927.0
        );
        assert_eq!(custom_filters[1][0], 4);
        assert_eq!(custom_filters[1][1], second_operator);
        assert_eq!(
            f64::from_le_bytes(custom_filters[1][2..10].try_into().unwrap()),
            45016.0
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_date_period_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_date_filter_pivot_workbook(
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

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let filter_header = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SX_FILTER)
                .then_some(payload)
        })
        .expect("BrtBeginSXFilter after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
        37,
        "Excel should preserve the ThisMonth pivot filter type"
    );

    let dynamic_filter = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_DYNAMIC_FILTER)
                .then_some(payload)
        })
        .expect("BrtDynamicFilter after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(dynamic_filter[0..4].try_into().unwrap()),
        0x0F
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_row_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_row_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RowItemFilter")
        .expect("RowItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let row_items = sxvi_payloads_for_xlsb_sxvd_axis(&records, 0x01);
    assert_eq!(
        row_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Excel should keep only Central hidden for the row item filter"
    );
    assert_eq!(
        i32::from_le_bytes(row_items[2][3..7].try_into().unwrap()),
        2
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_column_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_column_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ColumnItemFilter")
        .expect("ColumnItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["East", "West"])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let column_items = sxvi_payloads_for_xlsb_sxvd_axis(&records, 0x02);
    assert_eq!(
        column_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 0, 1, 0],
        "Excel should keep only Central hidden for the column item filter"
    );
    assert_eq!(
        i32::from_le_bytes(column_items[2][3..7].try_into().unwrap()),
        2
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_source_only_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_source_only_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SourceOnlyItemFilter")
        .expect("SourceOnlyItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Channel", ["Online"])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let hidden_field_items = sxvi_payloads_for_xlsb_sxvd_axis(&records, 0x00);
    assert_eq!(
        hidden_field_items
            .iter()
            .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![0, 1, 0],
        "Excel should keep only Store hidden for the source-only Channel filter"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_grouped_row_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_grouped_row_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedRowItemFilter")
        .expect("GroupedRowItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal"])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    assert!(
        sxvi_hidden_flag_groups_for_xlsb_sxvd_axis(&records, 0x01)
            .iter()
            .any(|flags| flags == &[1, 0, 0]),
        "Excel should preserve the hidden ungrouped item on the derived row field"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_grouped_column_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_grouped_column_item_filter_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedColumnItemFilter")
        .expect("GroupedColumnItemFilter pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal"])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    assert!(
        sxvi_hidden_flag_groups_for_xlsb_sxvd_axis(&records, 0x02)
            .iter()
            .any(|flags| flags == &[1, 0, 0]),
        "Excel should preserve the hidden ungrouped item on the derived column field"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_axis_field_options() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_axis_options_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("AxisOptions")
        .expect("AxisOptions pivot after Excel re-save");
    let field = &pivot.rows[0];

    assert_eq!(field.caption.as_deref(), Some("Market"));
    assert_eq!(field.sort, PivotSort::Descending);
    assert_eq!(field.subtotal, PivotSubtotal::Sum);
    assert_eq!(field.subtotal_caption.as_deref(), Some("? subtotal"));
    assert_eq!(
        field.subtotals,
        vec![
            PivotSubtotal::Sum,
            PivotSubtotal::Count,
            PivotSubtotal::Average
        ]
    );
    assert!(field.show_empty_items);
    assert!(!field.show_drop_downs);
    assert!(!field.subtotal_top);
    assert!(field.insert_blank_row);
    assert!(field.insert_page_break);
    assert!(field.include_new_items_in_filter);

    let pivot_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let row_field_payload = pivot_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXVD
                && payload[0] & 0x01 != 0)
                .then_some(payload)
        })
        .expect("row BrtBeginSXVD payload after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(row_field_payload[1..3].try_into().unwrap()),
        0x000E,
        "Excel should preserve custom subtotal function bits"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_sort_by_measure() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_sort_by_measure_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ValueSortedPivot")
        .expect("ValueSortedPivot after Excel re-save");

    assert_eq!(pivot.rows[0].sort, PivotSort::Descending);
    let measure = pivot.rows[0]
        .sort_by_measure
        .as_ref()
        .expect("sort-by-measure after Excel re-save");
    assert_eq!(measure.field.name, "Revenue");
    assert_eq!(measure.aggregate, PivotAggregate::Sum);
    assert_eq!(measure.name.as_deref(), Some("Sum of Revenue"));

    let pivot_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let record_types: Vec<_> = pivot_records
        .iter()
        .map(|(record_type, _)| *record_type)
        .collect();
    assert!(record_types.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXVD14));
    assert!(record_types.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PIVOT_AREA));
    assert!(record_types.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PIVOT_AREA_REF));

    let item_payload = pivot_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_PIVOT_AREA_REF_ITEM)
                .then_some(payload)
        })
        .expect("data-field auto-sort item after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(item_payload[..4].try_into().unwrap()),
        0,
        "auto-sort scope should target the first pivot data field"
    );
}

// features: Pivot table styles
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_style() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_styled_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("StyledPivot")
        .expect("StyledPivot after Excel re-save");

    assert_eq!(pivot.style, xlsb_styled_pivot_style());

    let pivot_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let style_payload = pivot_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_SX_VIEW_STYLE)
                .then_some(payload)
        })
        .expect("pivot style record after Excel re-save");
    assert_eq!(
        u16::from_le_bytes(style_payload[0..2].try_into().unwrap()),
        0x002E,
        "Excel should preserve the pivot style flags"
    );
    assert_eq!(wide_string_at(style_payload, 2).0, "PivotStyleLight16");
}

// features: Calculated fields
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_calculated_field() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_calculated_field_pivot_workbook());
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

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let calculated_field = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCD_FIELD)
                .then_some(payload)
        })
        .find(|payload| wide_string_at(payload, 20).0 == "Revenue")
        .expect("calculated-field BrtBeginPCDField after Excel re-save");
    assert_ne!(
        u16::from_le_bytes(calculated_field[0..2].try_into().unwrap()) & 0x0100,
        0,
        "Excel should preserve the calculated-field formula flag"
    );
    assert!(
        cache_records
            .iter()
            .any(|(record_type, _)| *record_type
                == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PNAMES),
        "Excel should preserve pivot formula source-field names"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_calculated_item() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_calculated_item_pivot_workbook());
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

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    assert!(
        cache_records
            .iter()
            .any(|(record_type, _)| *record_type == 0x00F3),
        "Excel should preserve BrtBeginPCDCalcItems"
    );
    assert!(
        cache_records
            .iter()
            .any(|(record_type, _)| *record_type == 0x00F5),
        "Excel should preserve BrtBeginPCDCalcItem"
    );
    assert!(
        cache_records.iter().any(|(record_type, payload)| {
            *record_type == duke_sheets_xlsb::biff12::records::BRT_PCDIA_STRING
                && wide_string_at(payload, 0).0 == "Combined"
                && u32::from_le_bytes(
                    payload[payload.len() - 6..payload.len() - 2]
                        .try_into()
                        .unwrap(),
                ) & 0x0000_0002
                    != 0
        }),
        "Excel should preserve the calculated shared item flag"
    );

    let pivot_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let row_items = sxvi_payloads_for_xlsb_sxvd_axis(&pivot_records, 0x01);
    assert!(
        row_items
            .iter()
            .any(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()) & 0x0008 != 0),
        "Excel should preserve the calculated row item flag"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_calculated_item_cell_like_refs() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_calculated_item_cell_like_pivot_workbook());
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

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let calculated_item = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F5).then_some(payload))
        .expect("BrtBeginPCDCalcItem after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(calculated_item[4..8].try_into().unwrap()),
        13,
        "Excel should preserve the calculated-item cell-like reference token stream"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_calculated_item_string_refs() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_calculated_item_string_ref_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRegionStringRef")
        .expect("CalculatedRegionStringRef after Excel re-save");

    assert_eq!(pivot.calculated_items.len(), 2);
    let all_regions = pivot
        .calculated_items
        .iter()
        .find(|item| item.item == PivotValue::String("All Regions".into()))
        .expect("All Regions calculated item after Excel re-save");
    assert_eq!(all_regions.field.name, "Region");
    assert_eq!(all_regions.formula, "Combined+Central");
    let combined = pivot
        .calculated_items
        .iter()
        .find(|item| item.item == PivotValue::String("Combined".into()))
        .expect("Combined calculated item after Excel re-save");
    assert_eq!(combined.formula, "East+West");

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let formula_lengths = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == 0x00F5).then(|| u32::from_le_bytes(payload[4..8].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(
        formula_lengths,
        vec![13, 13],
        "Excel should preserve string item references as PtgSxName token streams"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_calculated_field_function() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_calculated_field_function_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRevenueFunction")
        .expect("CalculatedRevenueFunction after Excel re-save");

    assert_eq!(pivot.calculated_fields.len(), 1);
    assert_eq!(pivot.calculated_fields[0].name, "Revenue");
    assert_eq!(pivot.calculated_fields[0].formula, "SUM(Units,Price)");

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let calculated_field = cache_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCD_FIELD)
                .then_some(payload)
        })
        .find(|payload| wide_string_at(payload, 20).0 == "Revenue")
        .expect("calculated-field function BrtBeginPCDField after Excel re-save");
    let (_, name_len) = wide_string_at(calculated_field, 20);
    let formula_offset = 20 + name_len;
    assert_eq!(
        u32::from_le_bytes(
            calculated_field[formula_offset..formula_offset + 4]
                .try_into()
                .unwrap()
        ),
        16,
        "Excel should preserve the calculated-field function token stream"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_calculated_item_function() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_calculated_item_function_pivot_workbook());
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

    let cache_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    let calculated_item = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F5).then_some(payload))
        .expect("BrtBeginPCDCalcItem after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(calculated_item[4..8].try_into().unwrap()),
        16,
        "Excel should preserve the calculated-item function token stream"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_multi_measure_values_axis() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_multi_measure_pivot_workbook());

    let pivot_records = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    assert!(pivot_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_ISXVD_COLS));
    assert!(pivot_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXDIS));
    assert_eq!(
        pivot_records
            .iter()
            .filter(|&&record| record == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXDI)
            .count(),
        2
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_all_aggregate_data_fields() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_all_aggregate_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("AllAggregatePivot")
        .unwrap();
    assert_eq!(pivot.layout.values_axis, PivotValuesAxis::Columns);
    assert_eq!(pivot.layout.values_axis_position, Some(0));
    assert_eq!(pivot.measures.len(), xlsb_all_aggregate_cases().len());
    for (measure, (aggregate, caption, _)) in pivot.measures.iter().zip(xlsb_all_aggregate_cases())
    {
        assert_eq!(measure.field.name, "Revenue");
        assert_eq!(measure.aggregate, aggregate);
        assert_eq!(measure.name.as_deref(), Some(caption));
    }

    let pivot_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let data_field_codes = pivot_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXDI)
                .then(|| u32::from_le_bytes(payload[4..8].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    let expected_codes = xlsb_all_aggregate_cases()
        .iter()
        .map(|(_, _, code)| *code)
        .collect::<Vec<_>>();
    assert_eq!(
        data_field_codes, expected_codes,
        "Excel should preserve every BIFF12 pivot aggregate function code"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_custom_measure_number_format() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_custom_measure_format_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("FormattedRevenue")
        .expect("FormattedRevenue after Excel re-save");

    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].number_format.as_deref(), Some("#,##0.0"));

    let style_records = xlsb_record_payloads(&zip_entry_bytes(&excel_bytes, "xl/styles.bin"));
    let custom_formats = style_records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_FMT).then(|| {
                let id = u16::from_le_bytes(payload[0..2].try_into().unwrap()) as u32;
                let (code, _) = wide_string_at(payload, 2);
                (id, code)
            })
        })
        .collect::<std::collections::HashMap<_, _>>();

    let pivot_records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let num_fmt_id = pivot_records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXDI)
                .then(|| u32::from_le_bytes(payload[20..24].try_into().unwrap()))
        })
        .expect("pivot data-field number format after Excel re-save");
    assert_eq!(
        custom_formats.get(&num_fmt_id).map(String::as_str),
        Some("#,##0.0")
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_show_as_data_fields() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_show_as_pivot_workbook());
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

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let show_as_codes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXDI)
                .then(|| u32::from_le_bytes(payload[8..12].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(show_as_codes, vec![5, 4, 1, 3]);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_x14_show_as_data_fields() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_x14_show_as_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("X14ShowAsPivot")
        .expect("X14ShowAsPivot after Excel re-save");

    assert_eq!(pivot.measures.len(), 5);
    assert_eq!(
        pivot.measures[0].show_as,
        PivotShowAs::PercentOfParentRowTotal
    );
    assert_eq!(
        pivot.measures[1].show_as,
        PivotShowAs::PercentOfParentColumnTotal
    );
    assert_eq!(
        pivot.measures[2].show_as,
        PivotShowAs::PercentOfParentTotal {
            base_field: PivotFieldRef::new("Rep")
        }
    );
    assert_eq!(
        pivot.measures[3].show_as,
        PivotShowAs::RankAscending {
            base_field: PivotFieldRef::new("Rep")
        }
    );
    assert_eq!(
        pivot.measures[4].show_as,
        PivotShowAs::RankDescending {
            base_field: PivotFieldRef::new("Region")
        }
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let base_show_as_codes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXDI)
                .then(|| u32::from_le_bytes(payload[8..12].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(base_show_as_codes, vec![0, 0, 0, 0, 0]);

    let base_field_indexes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXDI)
                .then(|| u32::from_le_bytes(payload[12..16].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(base_field_indexes, vec![0, 0, 1, 1, 0]);

    let x14_codes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_SXDI14)
                .then(|| u32::from_le_bytes(payload[4..8].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(x14_codes, vec![9, 10, 11, 13, 14]);
}

// features: Grouping (dates, numbers, items)
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_numeric_grouping() {
    let (result, _writer_bytes, _excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_numeric_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedAges")
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
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_date_grouping() {
    let (result, _writer_bytes, _excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_date_grouped_pivot_workbook());
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
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_date_grouped_page_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_page_date_grouped_pivot_workbook(&[1.0]));
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("MonthlyPageFilter")
        .expect("MonthlyPageFilter pivot after Excel re-save");
    assert_eq!(pivot.page_fields[0].field.name, "Date");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "Date");
            assert_eq!(*units, vec![PivotDateGroupUnit::Months]);
        }
        other => panic!("expected date grouping, got {other:?}"),
    }
    assert_eq!(pivot.filters, vec![PivotFilter::field_items("Date", [1.0])]);

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let page_field = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
        })
        .expect("BrtBeginSXPI after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(page_field[0..4].try_into().unwrap()),
        0,
        "Excel should preserve the date grouped source field index"
    );
    assert_eq!(
        u32::from_le_bytes(page_field[4..8].try_into().unwrap()),
        0,
        "Excel should preserve the selected grouped month item"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_date_grouped_page_multi_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_page_date_grouped_pivot_workbook(&[1.0, 2.0]));
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("MonthlyPageFilter")
        .expect("MonthlyPageFilter pivot after Excel re-save");
    assert_eq!(pivot.page_fields[0].field.name, "Date");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "Date");
            assert_eq!(*units, vec![PivotDateGroupUnit::Months]);
        }
        other => panic!("expected date grouping, got {other:?}"),
    }
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Date", [1.0, 2.0])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let page_field = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
        })
        .expect("BrtBeginSXPI after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(page_field[0..4].try_into().unwrap()),
        0,
        "Excel should preserve the date grouped source field index"
    );
    assert_eq!(
        u32::from_le_bytes(page_field[4..8].try_into().unwrap()),
        0x0010_00FE,
        "Excel should preserve the BIFF12 multiple-items page-filter sentinel"
    );
    assert!(
        sxvi_hidden_flag_groups_for_xlsb_sxvd_axis(&records, 0x04)
            .iter()
            .any(|flags| flags == &[0, 0, 1, 0]),
        "Excel should preserve hidden flags on the date grouped page field"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_multi_unit_date_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_multi_unit_date_grouped_pivot_workbook());
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
        other => panic!("expected date grouping, got {other:?}"),
    }

    let cache_records = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    assert_eq!(
        cache_records
            .iter()
            .filter(|&&record| record == duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCDFG_RANGE)
            .count(),
        2,
        "Excel should preserve both derived date grouping ranges"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_multi_unit_date_grouped_page_field() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_multi_unit_date_grouped_page_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedDatePage")
        .expect("GroupedDatePage after Excel re-save");

    assert_eq!(pivot.page_fields.len(), 1);
    assert_eq!(pivot.page_fields[0].field.name, "SaleDate");
    assert!(pivot.filters.is_empty());
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "SaleDate");
            assert_eq!(
                *units,
                vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months]
            );
        }
        other => panic!("expected date grouping, got {other:?}"),
    }

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let page_fields_payload = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPIS).then_some(payload)
        })
        .expect("page field collection after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(page_fields_payload[0..4].try_into().unwrap()),
        2,
        "Excel should preserve one native page field for each date grouping unit"
    );
    let page_field_indexes = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPI)
                .then(|| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(page_field_indexes, vec![2, 3]);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_multi_unit_date_grouped_page_filters() {
    let (result, _writer_bytes, excel_bytes) = roundtrip_through_excel_xlsb_bytes(
        &xlsb_multi_unit_date_grouped_page_filter_pivot_workbook(),
    );
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedDatePage")
        .expect("GroupedDatePage after Excel re-save");

    assert_eq!(pivot.page_fields.len(), 1);
    assert_eq!(pivot.page_fields[0].field.name, "SaleDate");
    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Date { field, units } => {
            assert_eq!(field.name, "SaleDate");
            assert_eq!(
                *units,
                vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months]
            );
        }
        other => panic!("expected date grouping, got {other:?}"),
    }
    assert_eq!(
        pivot.filters,
        vec![
            PivotFilter::field_items("SaleDate (Years)", [2024.0]),
            PivotFilter::field_items("SaleDate (Months)", [1.0, 2.0]),
        ]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let page_fields = records
        .iter()
        .filter_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
        })
        .collect::<Vec<_>>();
    assert_eq!(page_fields.len(), 2);
    assert_eq!(
        page_fields
            .iter()
            .map(|payload| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
            .collect::<Vec<_>>(),
        vec![2, 3],
        "Excel should preserve Years and Months as separate page fields"
    );
    assert_eq!(
        u32::from_le_bytes(page_fields[0][4..8].try_into().unwrap()),
        0,
        "Excel should preserve the selected 2024 year item"
    );
    assert_eq!(
        u32::from_le_bytes(page_fields[1][4..8].try_into().unwrap()),
        0x0010_00FE,
        "Excel should preserve the multi-selected month sentinel"
    );
    let page_hidden_flags = sxvi_hidden_flag_groups_for_xlsb_sxvd_axis(&records, 0x04);
    assert!(
        page_hidden_flags.iter().any(|flags| flags == &[0, 0, 1, 0]),
        "Excel should preserve hidden flags on the date-derived month page field"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_manual_grouping() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_manual_grouped_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualGroupedRegions")
        .unwrap();
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
        other => panic!("expected manual grouping, got {other:?}"),
    }

    let cache_records = xlsb_record_types(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin",
    ));
    assert!(
        cache_records.contains(&duke_sheets_xlsb::biff12::records::BRT_BEGIN_PCDFG_DISCRETE),
        "Excel should preserve a manual grouping discrete map"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_manual_page_grouping_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_manual_page_grouped_pivot_workbook(&["Coastal"]));
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualPageRegions")
        .expect("ManualPageRegions pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal"])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let page_field = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
        })
        .expect("BrtBeginSXPI after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(page_field[0..4].try_into().unwrap()),
        3,
        "Excel should preserve the derived manual group field index"
    );
    assert_eq!(
        u32::from_le_bytes(page_field[4..8].try_into().unwrap()),
        1,
        "Excel should preserve the selected manual group item"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_pivot_manual_page_grouping_multi_item_filter() {
    let (result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_manual_page_grouped_pivot_workbook(&[
            "Coastal", "Interior",
        ]));
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualPageRegions")
        .expect("ManualPageRegions pivot after Excel re-save");
    assert_eq!(
        pivot.filters,
        vec![PivotFilter::field_items("Region", ["Coastal", "Interior"])]
    );

    let records = xlsb_record_payloads(&zip_entry_bytes(
        &excel_bytes,
        "xl/pivotTables/pivotTable1.bin",
    ));
    let page_field = records
        .iter()
        .find_map(|(record_type, payload)| {
            (*record_type == duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
        })
        .expect("BrtBeginSXPI after Excel re-save");
    assert_eq!(
        u32::from_le_bytes(page_field[0..4].try_into().unwrap()),
        3,
        "Excel should preserve the derived manual group field index"
    );
    assert_eq!(
        u32::from_le_bytes(page_field[4..8].try_into().unwrap()),
        0x0010_00FE,
        "Excel should preserve the BIFF12 multiple-items page-filter sentinel"
    );
    assert!(
        sxvi_hidden_flag_groups_for_xlsb_sxvd_axis(&records, 0x04)
            .iter()
            .any(|flags| flags == &[1, 0, 0, 0]),
        "Excel should preserve hidden flags on the derived manual page field"
    );
}

fn zip_has_entry(bytes: &[u8], name: &str) -> bool {
    let reader = std::io::Cursor::new(bytes);
    let Ok(mut archive) = zip::ZipArchive::new(reader) else {
        return false;
    };
    let found = archive.by_name(name).is_ok();
    found
}

fn zip_entry_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    use std::io::Read;

    let reader = std::io::Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(reader).expect("open xlsb zip");
    let mut file = archive
        .by_name(name)
        .unwrap_or_else(|e| panic!("{name}: {e}"));
    let mut out = Vec::new();
    file.read_to_end(&mut out).unwrap();
    out
}

fn xlsb_record_types(data: &[u8]) -> Vec<u16> {
    let mut iter = duke_sheets_xlsb::biff12::RecordIter::new(std::io::Cursor::new(data));
    let mut out = Vec::new();
    let mut buf = Vec::new();
    while let Ok((record_type, len)) = iter.next_record(&mut buf) {
        out.push(record_type);
        buf.truncate(len);
    }
    out
}

fn xlsb_record_payloads(data: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut iter = duke_sheets_xlsb::biff12::RecordIter::new(std::io::Cursor::new(data));
    let mut out = Vec::new();
    let mut buf = Vec::new();
    while let Ok((record_type, len)) = iter.next_record(&mut buf) {
        buf.truncate(len);
        out.push((record_type, buf.clone()));
    }
    out
}

fn xlsb_sxview_flag(payload: &[u8], bit: usize) -> bool {
    payload[4 + bit / 8] & (1u8 << (bit % 8)) != 0
}

fn sxvi_payloads_for_xlsb_sxvd_axis(records: &[(u16, Vec<u8>)], axis: u8) -> Vec<Vec<u8>> {
    let mut in_target_field = false;
    let mut payloads = Vec::new();
    for (record_type, payload) in records {
        match *record_type {
            duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXVD => {
                in_target_field = payload.first().copied() == Some(axis);
            }
            duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXVI if in_target_field => {
                payloads.push(payload.clone());
            }
            duke_sheets_xlsb::biff12::records::BRT_END_SXVD => {
                in_target_field = false;
            }
            _ => {}
        }
    }
    payloads
}

fn sxvi_hidden_flag_groups_for_xlsb_sxvd_axis(
    records: &[(u16, Vec<u8>)],
    axis: u8,
) -> Vec<Vec<u16>> {
    let mut in_target_field = false;
    let mut current: Option<Vec<u16>> = None;
    let mut groups = Vec::new();
    for (record_type, payload) in records {
        match *record_type {
            duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXVD => {
                in_target_field = payload.first().copied() == Some(axis);
                current = in_target_field.then(Vec::new);
            }
            duke_sheets_xlsb::biff12::records::BRT_BEGIN_SXVI if in_target_field => {
                if let Some(current) = &mut current {
                    current.push(u16::from_le_bytes(payload[1..3].try_into().unwrap()));
                }
            }
            duke_sheets_xlsb::biff12::records::BRT_END_SXVD => {
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

fn wide_string_at(data: &[u8], offset: usize) -> (String, usize) {
    let count = u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap()) as usize;
    let byte_len = count * 2;
    let start = offset + 4;
    let end = start + byte_len;
    let units = data[start..end]
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    (String::from_utf16_lossy(&units), 4 + byte_len)
}
