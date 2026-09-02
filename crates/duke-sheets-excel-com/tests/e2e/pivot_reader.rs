use crate::{
    cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture,
    temp_fixture_xls, temp_fixture_xlsb, TempFixture,
};
use duke_sheets_core::{
    CellAddress, CellRange, PivotAggregate, PivotDateGroupUnit, PivotGrouping, PivotSource,
    Workbook,
};
use duke_sheets_excel_com::{ChainStep, ExcelBridge, SheetRef};
use duke_sheets_xls::XlsReader;
use duke_sheets_xlsb::XlsbReader;
use duke_sheets_xlsx::XlsxReader;
use excel_com_protocol::ResponseData;
use serde_json::{json, Value};

const XL_DATABASE: i32 = 1;
const XL_ROW_FIELD: i32 = 1;
const XL_COLUMN_FIELD: i32 = 2;
const XL_PAGE_FIELD: i32 = 3;
const XL_SUM: i32 = -4157;

#[derive(Clone, Copy)]
enum PivotFormat {
    Xlsx,
    Xlsb,
    Xls,
}

impl PivotFormat {
    fn fixture(self) -> TempFixture {
        match self {
            Self::Xlsx => temp_fixture(),
            Self::Xlsb => temp_fixture_xlsb(),
            Self::Xls => temp_fixture_xls(),
        }
    }

    fn save_format(self) -> i32 {
        match self {
            Self::Xlsx => 51,
            Self::Xlsb => 50,
            Self::Xls => 56,
        }
    }

    fn read(self, fixture: &TempFixture) -> Workbook {
        match self {
            Self::Xlsx => {
                XlsxReader::read_file(&fixture.host_path).expect("read native XLSX pivot")
            }
            Self::Xlsb => {
                XlsbReader::read_file(&fixture.host_path).expect("read native XLSB pivot")
            }
            Self::Xls => XlsReader::read_file(&fixture.host_path).expect("read native XLS pivot"),
        }
    }
}

fn sheet_chain() -> ChainStep {
    SheetRef::Index(0).to_chain_step()
}

fn indexed(name: &str, index: impl Into<Value>) -> ChainStep {
    ChainStep::Indexed(name.to_string(), index.into())
}

fn handle_from(response: Option<ResponseData>, operation: &str) -> u64 {
    match response {
        Some(ResponseData::Handle { handle }) => handle,
        other => panic!("{operation} did not return a COM handle: {other:?}"),
    }
}

fn set_field_orientation(excel: &ExcelBridge, pivot: u64, field: &str, orientation: i32) {
    excel
        .set(
            pivot,
            vec![indexed("PivotFields", field)],
            "Orientation",
            json!(orientation),
        )
        .unwrap_or_else(|e| panic!("set {field} pivot orientation: {e}"));
}

fn populate_basic_source(wb: &duke_sheets_excel_com::Workbook<'_>) {
    let rows: [[Value; 4]; 9] = [
        [
            json!("Region"),
            json!("Product"),
            json!("Channel"),
            json!("Amount"),
        ],
        [json!("East"), json!("Hardware"), json!("Direct"), json!(10)],
        [json!("East"), json!("Hardware"), json!("Retail"), json!(20)],
        [json!("East"), json!("Software"), json!("Direct"), json!(30)],
        [json!("East"), json!("Software"), json!("Retail"), json!(40)],
        [json!("West"), json!("Hardware"), json!("Direct"), json!(50)],
        [json!("West"), json!("Hardware"), json!("Retail"), json!(60)],
        [json!("West"), json!("Software"), json!("Direct"), json!(70)],
        [json!("West"), json!("Software"), json!("Retail"), json!(80)],
    ];

    for (row, values) in rows.iter().enumerate() {
        for (column, value) in values.iter().enumerate() {
            let cell = format!("{}{}", (b'A' + column as u8) as char, row + 1);
            match value {
                Value::String(value) => wb
                    .set_cell_value(&cell, value.as_str())
                    .expect("write source text"),
                Value::Number(value) => wb
                    .set_cell_value(&cell, value.as_i64().expect("integer source value") as f64)
                    .expect("write source number"),
                _ => unreachable!("source values are strings or numbers"),
            }
        }
    }
}

fn add_native_pivot(excel: &ExcelBridge, workbook: u64, source_range: &str, name: &str) -> u64 {
    let source = excel
        .navigate(
            workbook,
            vec![sheet_chain(), indexed("Range", source_range)],
        )
        .expect("navigate to pivot source range");
    let destination = excel
        .navigate(workbook, vec![sheet_chain(), indexed("Range", "G3")])
        .expect("navigate to pivot destination");

    let caches = handle_from(
        excel
            .invoke(workbook, vec![], "PivotCaches", vec![])
            .expect("Workbook.PivotCaches"),
        "Workbook.PivotCaches",
    );
    let cache = handle_from(
        excel
            .invoke(
                caches,
                vec![],
                "Create",
                vec![json!(XL_DATABASE), json!({"$ref": source})],
            )
            .expect("PivotCaches.Create"),
        "PivotCaches.Create",
    );
    let pivot = handle_from(
        excel
            .invoke(
                cache,
                vec![],
                "CreatePivotTable",
                vec![json!({"$ref": destination}), json!(name)],
            )
            .expect("PivotCache.CreatePivotTable"),
        "PivotCache.CreatePivotTable",
    );

    set_field_orientation(excel, pivot, "Region", XL_ROW_FIELD);
    set_field_orientation(excel, pivot, "Product", XL_COLUMN_FIELD);
    set_field_orientation(excel, pivot, "Channel", XL_PAGE_FIELD);

    let amount = excel
        .navigate(pivot, vec![indexed("PivotFields", "Amount")])
        .expect("navigate to Amount pivot field");
    let data_field = excel
        .invoke(
            pivot,
            vec![],
            "AddDataField",
            vec![
                json!({"$ref": amount}),
                json!("Total Amount"),
                json!(XL_SUM),
            ],
        )
        .expect("PivotTable.AddDataField");
    if let Some(ResponseData::Handle { handle }) = data_field {
        excel.release(handle).expect("release data field");
    }

    excel.release(amount).expect("release Amount field");
    excel.release(source).expect("release source range");
    excel
        .release(destination)
        .expect("release destination range");
    excel.release(cache).expect("release pivot cache");
    excel.release(caches).expect("release pivot caches");
    pivot
}

fn author_basic_pivot(format: PivotFormat, name: &str) -> (TempFixture, Workbook) {
    let fixture = format.fixture();
    ensure_vm_temp_dir();
    {
        let excel = excel_bridge().lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        excel
            .set(wb.handle(), vec![sheet_chain()], "Name", json!("Data"))
            .expect("rename source worksheet");
        populate_basic_source(&wb);

        let pivot = add_native_pivot(excel, wb.handle(), "A1:D9", name);
        excel.release(pivot).expect("release pivot table");
        wb.save_as(&fixture.vm_path, format.save_format())
            .expect("Excel SaveAs native pivot workbook");
        let saved_name = wb.name().expect("read saved workbook name");
        assert!(
            !saved_name.contains("Repaired"),
            "Excel marked its native pivot workbook as repaired: {saved_name}"
        );
        wb.close().expect("close native pivot workbook");
    }

    pull_file_from_vm(&fixture);
    let workbook = format.read(&fixture);
    (fixture, workbook)
}

fn assert_basic_pivot(workbook: &Workbook, format: PivotFormat, name: &str) {
    let sheet = workbook.worksheet(0).expect("Data worksheet");
    let pivot = sheet
        .pivot_table_by_name(name)
        .unwrap_or_else(|| panic!("reader did not return native pivot {name}"));

    match &pivot.source {
        PivotSource::WorksheetRange { sheet, range } => {
            assert_eq!(sheet.as_deref(), Some("Data"), "pivot source sheet");
            assert_eq!(
                *range,
                CellRange::parse("A1:D9").unwrap(),
                "pivot source range"
            );
        }
        other => panic!("unexpected native pivot source: {other:?}"),
    }
    // BIFF SXVIEW/SXLOCATION stores the body at G3 plus the page-area height,
    // so the binary readers recover the semantic full-output origin at G1.
    let expected_target = match format {
        PivotFormat::Xlsx => "G3",
        PivotFormat::Xlsb | PivotFormat::Xls => "G1",
    };
    assert_eq!(
        pivot.target,
        CellAddress::parse(expected_target).unwrap(),
        "semantic pivot target"
    );
    assert_eq!(
        pivot
            .rendered_range
            .expect("native pivot rendered range")
            .start,
        CellAddress::parse("G3").unwrap(),
        "stored pivot body destination"
    );
    assert_eq!(pivot.rows.len(), 1, "row fields: {:#?}", pivot.rows);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(
        pivot.columns.len(),
        1,
        "column fields: {:#?}",
        pivot.columns
    );
    assert_eq!(pivot.columns[0].field.name, "Product");
    assert_eq!(
        pivot.page_fields.len(),
        1,
        "page fields: {:#?}",
        pivot.page_fields
    );
    assert_eq!(pivot.page_fields[0].field.name, "Channel");
    assert_eq!(pivot.measures.len(), 1, "measures: {:#?}", pivot.measures);
    assert_eq!(pivot.measures[0].field.name, "Amount");
    assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Amount"));
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
}

// features: Pivot cache (source data); Pivot table definition; Row / column / value fields; Filter (page) fields; Aggregate functions (Sum/Count/Avg/...)
#[test]
fn excel_authored_xlsx_pivot_is_parsed() {
    let (fixture, workbook) = author_basic_pivot(PivotFormat::Xlsx, "NativeXlsxPivot");
    assert_basic_pivot(&workbook, PivotFormat::Xlsx, "NativeXlsxPivot");
    cleanup_fixture(&fixture);
}

// features: Pivot cache (source data); Pivot table definition; Row / column / value fields; Filter (page) fields; Aggregate functions (Sum/Count/Avg/...)
#[test]
fn excel_authored_xlsb_pivot_is_parsed() {
    let (fixture, workbook) = author_basic_pivot(PivotFormat::Xlsb, "NativeXlsbPivot");
    assert_basic_pivot(&workbook, PivotFormat::Xlsb, "NativeXlsbPivot");
    cleanup_fixture(&fixture);
}

// features: Pivot cache (source data); Pivot table definition; Row / column / value fields; Filter (page) fields; Aggregate functions (Sum/Count/Avg/...)
#[test]
fn excel_authored_xls_pivot_is_parsed() {
    let (fixture, workbook) = author_basic_pivot(PivotFormat::Xls, "NativeXlsPivot");
    assert_basic_pivot(&workbook, PivotFormat::Xls, "NativeXlsPivot");
    cleanup_fixture(&fixture);
}

fn author_grouped_xlsx_pivot() -> (TempFixture, Workbook) {
    let fixture = temp_fixture();
    ensure_vm_temp_dir();
    {
        let excel = excel_bridge().lock().unwrap();
        let wb = excel
            .create_workbook()
            .expect("create grouped pivot workbook");
        excel
            .set(wb.handle(), vec![sheet_chain()], "Name", json!("Data"))
            .expect("rename grouped pivot worksheet");

        for (cell, value) in [
            ("A1", "Region"),
            ("B1", "Product"),
            ("C1", "Channel"),
            ("D1", "Amount"),
            ("E1", "SaleDate"),
        ] {
            wb.set_cell_value(cell, value)
                .expect("write grouped pivot header");
        }
        let dates = [
            45292.0, 45323.0, 45352.0, 45383.0, 45413.0, 45444.0, 45474.0, 45505.0,
        ];
        for (index, date) in dates.iter().enumerate() {
            let row = index + 2;
            wb.set_cell_value(&format!("A{row}"), if index < 4 { "East" } else { "West" })
                .expect("write region");
            wb.set_cell_value(
                &format!("B{row}"),
                if index % 2 == 0 {
                    "Hardware"
                } else {
                    "Software"
                },
            )
            .expect("write product");
            wb.set_cell_value(
                &format!("C{row}"),
                if index % 2 == 0 { "Direct" } else { "Retail" },
            )
            .expect("write channel");
            wb.set_cell_value(&format!("D{row}"), ((index + 1) * 10) as f64)
                .expect("write amount");
            wb.set_cell_value(&format!("E{row}"), *date)
                .expect("write date");
        }
        wb.set_number_format("E2:E9", "m/d/yyyy")
            .expect("format source dates");

        let pivot = add_native_pivot(excel, wb.handle(), "A1:E9", "NativeGroupedPivot");
        set_field_orientation(excel, pivot, "SaleDate", XL_ROW_FIELD);
        excel
            .invoke(
                pivot,
                vec![
                    indexed("PivotFields", "SaleDate"),
                    indexed("PivotItems", 1),
                    ChainStep::Property("LabelRange".to_string()),
                ],
                "Group",
                vec![
                    json!(true),
                    json!(true),
                    json!(1),
                    json!([false, false, false, false, true, false, true]),
                ],
            )
            .expect("Range.Group dates by months and years");

        excel.release(pivot).expect("release grouped pivot");
        wb.save_as(&fixture.vm_path, 51)
            .expect("save grouped XLSX pivot");
        let saved_name = wb.name().expect("read grouped workbook name");
        assert!(
            !saved_name.contains("Repaired"),
            "Excel marked grouped workbook as repaired: {saved_name}"
        );
        wb.close().expect("close grouped pivot workbook");
    }

    pull_file_from_vm(&fixture);
    let workbook =
        XlsxReader::read_file(&fixture.host_path).expect("read native grouped XLSX pivot");
    (fixture, workbook)
}

// features: Pivot cache (source data); Pivot table definition; Row / column / value fields; Filter (page) fields; Aggregate functions (Sum/Count/Avg/...); Grouping (dates, numbers, items)
#[test]
fn excel_authored_xlsx_date_grouping_is_parsed() {
    let (fixture, workbook) = author_grouped_xlsx_pivot();
    let pivot = workbook
        .worksheet(0)
        .expect("Data worksheet")
        .pivot_table_by_name("NativeGroupedPivot")
        .expect("reader returns grouped native pivot");
    assert!(
        pivot.groupings.iter().any(|grouping| matches!(
            grouping,
            PivotGrouping::Date { field, units }
                if field.name == "SaleDate"
                    && units.contains(&PivotDateGroupUnit::Months)
                    && units.contains(&PivotDateGroupUnit::Years)
        )),
        "expected SaleDate grouped by months and years, parsed: {:#?}",
        pivot.groupings
    );
    cleanup_fixture(&fixture);
}
