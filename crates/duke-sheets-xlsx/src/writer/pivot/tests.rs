use super::super::*;
use crate::reader::XlsxReader;
use duke_sheets_core::{
    CellRange, PivotAggregate, PivotDateGroupUnit, PivotDatePeriod, PivotExtension, PivotField,
    PivotFieldRef, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayout, PivotLayoutKind,
    PivotManualGroup, PivotMeasure, PivotRefreshPolicy, PivotShowAs, PivotSort, PivotSource,
    PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable, PivotValue, PivotValuesAxis,
    WorkbookConnection, WorkbookConnectionCredentials, WorkbookConnectionKind,
    WorkbookConnectionParameter, WorkbookConnectionParameterValue, WorkbookExtensionPart,
};
use duke_sheets_pivot::WorkbookPivotExt;
use ssfmt::{date_serial::date_to_serial, DateSystem};
use std::io::Read;

fn read_zip_entry(bytes: Vec<u8>, path: &str) -> String {
    let cursor = Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(path).expect("zip entry exists");
    let mut s = String::new();
    file.read_to_string(&mut s).expect("read zip entry utf8");
    s
}

// features: Pivot cache (source data); Pivot table definition; Row / column / value fields; Filter (page) fields; Aggregate functions (Sum/Count/Avg/...)
#[test]
fn test_writer_emits_pivot_table_and_cache_parts() {
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
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "Q2").unwrap();
    sheet.set_cell_value("C4", 15.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Region", ["East"]))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let ct = read_zip_entry(bytes.clone(), "[Content_Types].xml");
    assert!(ct.contains("/xl/pivotTables/pivotTable1.xml"));
    assert!(ct.contains("/xl/pivotCache/pivotCacheDefinition1.xml"));
    assert!(ct.contains("/xl/pivotCache/pivotCacheRecords1.xml"));

    let workbook_xml = read_zip_entry(bytes.clone(), "xl/workbook.xml");
    assert!(workbook_xml.contains("<pivotCaches>"));
    assert!(workbook_xml.contains(r#"cacheId="1""#));
    assert!(workbook_xml.contains(r#"r:id="rIdPivotCache1""#));

    let workbook_rels = read_zip_entry(bytes.clone(), "xl/_rels/workbook.xml.rels");
    assert!(workbook_rels.contains(RT_PIVOT_CACHE_DEFINITION));
    assert!(workbook_rels.contains("pivotCache/pivotCacheDefinition1.xml"));

    let sheet_xml = read_zip_entry(bytes.clone(), "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains("<pivotTableDefinitions>"));
    assert!(sheet_xml.contains("<pivotTableDefinition r:id="));

    let sheet_rels = read_zip_entry(bytes.clone(), "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(sheet_rels.contains(RT_PIVOT_TABLE));
    assert!(sheet_rels.contains("../pivotTables/pivotTable1.xml"));

    let pivot_rels = read_zip_entry(bytes.clone(), "xl/pivotTables/_rels/pivotTable1.xml.rels");
    assert!(pivot_rels.contains(RT_PIVOT_CACHE_DEFINITION));
    assert!(pivot_rels.contains("../pivotCache/pivotCacheDefinition1.xml"));

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"name="SalesPivot""#));
    assert!(pivot_xml.contains(r#"cacheId="1""#));
    assert!(pivot_xml.contains(r#"<rowFields count="1"><field x="0"/>"#));
    assert!(pivot_xml.contains(r#"<colFields count="1"><field x="1"/>"#));
    assert!(pivot_xml.contains(r#"<dataField name="Revenue" fld="2" subtotal="sum"/>"#));
    assert!(
        pivot_xml.contains(r#"<item x="1" h="1"/>"#),
        "West should be hidden by the Region item filter"
    );

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<worksheetSource ref="A1:C4" sheet="Sheet1"/>"#));
    assert!(cache_def.contains(r#"<cacheFields count="3">"#));
    assert!(cache_def.contains(r#"<cacheField name="Region">"#));

    let cache_records = read_zip_entry(bytes, "xl/pivotCache/pivotCacheRecords1.xml");
    assert!(cache_records.contains(r#"<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">"#));
    assert_eq!(cache_records.matches("<r>").count(), 3);

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("rewrite workbook");
    let roundtrip = XlsxReader::read(Cursor::new(out.into_inner())).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SalesPivot")
        .unwrap();
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.columns.len(), 1);
    assert_eq!(pivot.columns[0].field.name, "Quarter");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
    assert!(matches!(
        &pivot.source,
        PivotSource::WorksheetRange {
            sheet: Some(sheet),
            range
        } if sheet == "Sheet1" && range.to_a1_string() == "A1:C4"
    ));
    assert_eq!(pivot.filters.len(), 1);
    match &pivot.filters[0] {
        PivotFilter::FieldItems {
            field,
            allowed_items,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(allowed_items.len(), 1);
            assert_eq!(allowed_items[0].to_string(), "East");
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
}
#[test]
fn test_writer_round_trips_table_source_pivot() {
    use duke_sheets_core::table::{Table, TableColumn};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
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

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<worksheetSource name="SalesData"/>"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SalesPivot")
        .unwrap();
    assert!(matches!(
        &pivot.source,
        PivotSource::Table { name } if name == "SalesData"
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

#[test]
fn test_writer_round_trips_external_pivot_source_definition() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let pivot = PivotTable::builder("ExternalSales")
        .source(PivotSource::External {
            connection_name: "7".to_string(),
            command_text: None,
        })
        .target_address("A1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<cacheSource type="external" connectionId="7"/>"#));
    assert!(cache_def.contains(r#"saveData="0""#));
    assert!(cache_def.contains(r#"recordCount="0""#));
    assert!(cache_def.contains(r#"<cacheField name="Region">"#));
    assert!(cache_def.contains(r#"<cacheField name="Revenue">"#));

    let cache_records = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheRecords1.xml");
    assert!(cache_records.contains(r#"<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0">"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ExternalSales")
        .unwrap();
    assert!(matches!(
        &pivot.source,
        PivotSource::External {
            connection_name,
            command_text: None
        } if connection_name == "7"
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert!(matches!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(duke_sheets_core::PivotCacheSourceKind::External)
    ));
}

#[test]
fn test_writer_round_trips_external_pivot_database_connection() {
    let command = "select Region, Revenue from Sales";
    let mut wb = Workbook::new();
    let mut connection =
        WorkbookConnection::database(7, "SalesConnection", "Provider=MSDASQL;DSN=Sales;")
            .with_command(command)
            .with_source_file("connections/sales.dsn")
            .with_odc_file("connections/sales.odc")
            .with_description("Sales warehouse")
            .with_connection_type(5)
            .with_keep_alive(true)
            .with_interval(30)
            .with_reconnection_method(2)
            .with_refresh_on_load(true)
            .with_save_password(true)
            .with_only_use_connection_file(true)
            .with_credentials(WorkbookConnectionCredentials::Stored)
            .with_single_sign_on_id("sales-sso")
            .with_parameter({
                let mut parameter = WorkbookConnectionParameter::value(
                    "RegionParam",
                    WorkbookConnectionParameterValue::String("East".to_string()),
                );
                parameter.sql_type = 12;
                parameter
            })
            .with_parameter({
                let mut parameter =
                    WorkbookConnectionParameter::cell("MinRevenue", "Sheet1!$A$1");
                parameter.sql_type = 8;
                parameter.refresh_on_change = true;
                parameter
            });
    connection.min_refreshable_version = 3;
    connection.new_connection = true;
    wb.add_data_connection(connection).unwrap();
    let sheet = wb.worksheet_mut(0).unwrap();
    let pivot = PivotTable::builder("ExternalSales")
        .source(PivotSource::External {
            connection_name: "SalesConnection".to_string(),
            command_text: Some(command.to_string()),
        })
        .target_address("A1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let content_types = read_zip_entry(bytes.clone(), "[Content_Types].xml");
    assert!(content_types.contains("/xl/connections.xml"));
    assert!(content_types.contains(CT_CONNECTIONS));

    let workbook_rels = read_zip_entry(bytes.clone(), "xl/_rels/workbook.xml.rels");
    assert!(workbook_rels.contains(RT_CONNECTIONS));
    assert!(workbook_rels.contains("connections.xml"));

    let connections = read_zip_entry(bytes.clone(), "xl/connections.xml");
    assert!(connections.contains(r#"<connection id="7" name="SalesConnection""#));
    assert!(connections.contains(r#"sourceFile="connections/sales.dsn""#));
    assert!(connections.contains(r#"odcFile="connections/sales.odc""#));
    assert!(connections.contains(r#"keepAlive="1""#));
    assert!(connections.contains(r#"interval="30""#));
    assert!(connections.contains(r#"description="Sales warehouse""#));
    assert!(connections.contains(r#"type="5""#));
    assert!(connections.contains(r#"reconnectionMethod="2""#));
    assert!(connections.contains(r#"minRefreshableVersion="3""#));
    assert!(connections.contains(r#"savePassword="1""#));
    assert!(connections.contains(r#"new="1""#));
    assert!(connections.contains(r#"onlyUseConnectionFile="1""#));
    assert!(connections.contains(r#"refreshOnLoad="1""#));
    assert!(connections.contains(r#"credentials="stored""#));
    assert!(connections.contains(r#"singleSignOnId="sales-sso""#));
    assert!(connections.contains(r#"<dbPr connection="Provider=MSDASQL;DSN=Sales;" command="select Region, Revenue from Sales" commandType="2"/>"#));
    assert!(connections.contains(r#"<parameters count="2">"#));
    assert!(connections.contains(r#"<parameter name="RegionParam" sqlType="12" parameterType="value" refreshOnChange="0" string="East"/>"#));
    assert!(connections.contains(r#"<parameter name="MinRevenue" sqlType="8" parameterType="cell" refreshOnChange="1" cell="Sheet1!$A$1"/>"#));

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<cacheSource type="external" connectionId="7"/>"#));
    assert!(cache_def.contains(r#"saveData="0""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let connection = &roundtrip.data_connections()[0];
    assert_eq!(connection.id, 7);
    assert_eq!(connection.name, "SalesConnection");
    assert_eq!(
        connection.source_file.as_deref(),
        Some("connections/sales.dsn")
    );
    assert_eq!(
        connection.odc_file.as_deref(),
        Some("connections/sales.odc")
    );
    assert_eq!(connection.description.as_deref(), Some("Sales warehouse"));
    assert_eq!(connection.connection_type, Some(5));
    assert_eq!(connection.min_refreshable_version, 3);
    assert!(connection.keep_alive);
    assert_eq!(connection.interval, 30);
    assert_eq!(connection.reconnection_method, 2);
    assert!(connection.save_password);
    assert!(connection.new_connection);
    assert!(connection.only_use_connection_file);
    assert_eq!(
        connection.credentials,
        Some(WorkbookConnectionCredentials::Stored)
    );
    assert_eq!(connection.single_sign_on_id.as_deref(), Some("sales-sso"));
    assert_eq!(connection.parameters.len(), 2);
    assert_eq!(
        connection.parameters[0].name.as_deref(),
        Some("RegionParam")
    );
    assert_eq!(connection.parameters[0].sql_type, 12);
    assert_eq!(
        connection.parameters[0].value,
        WorkbookConnectionParameterValue::String("East".to_string())
    );
    assert_eq!(connection.parameters[1].name.as_deref(), Some("MinRevenue"));
    assert_eq!(connection.parameters[1].sql_type, 8);
    assert!(connection.parameters[1].refresh_on_change);
    assert_eq!(
        connection.parameters[1].value,
        WorkbookConnectionParameterValue::Cell("Sheet1!$A$1".to_string())
    );
    match &connection.kind {
        WorkbookConnectionKind::Database {
            connection,
            command: roundtrip_command,
            command_type,
        } => {
            assert_eq!(connection, "Provider=MSDASQL;DSN=Sales;");
            assert_eq!(roundtrip_command.as_deref(), Some(command));
            assert_eq!(*command_type, Some(2));
        }
        other => panic!("unexpected connection kind: {other:?}"),
    }

    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ExternalSales")
        .unwrap();
    assert!(matches!(
        &pivot.source,
        PivotSource::External {
            connection_name,
            command_text: Some(roundtrip_command)
        } if connection_name == "SalesConnection" && roundtrip_command == command
    ));
}

#[test]
fn test_writer_round_trips_olap_pivot_source_definition() {
    let mut wb = Workbook::new();
    let mut connection = WorkbookConnection::olap(10, "CubeSales").with_connection_type(5);
    connection.kind = WorkbookConnectionKind::Olap {
        connection: Some("Provider=MSOLAP;Data Source=olapserver;".to_string()),
        command: Some("SalesCube".to_string()),
        command_type: Some(1),
        local: false,
        local_connection: None,
        local_refresh: true,
        send_locale: true,
        row_drill_count: None,
    };
    wb.add_data_connection(connection).unwrap();
    let sheet = wb.worksheet_mut(0).unwrap();
    let pivot = PivotTable::builder("OlapSales")
        .source(PivotSource::Olap {
            connection_name: "CubeSales".to_string(),
            cube: Some("SalesCube".to_string()),
            command_text: Some("SalesCube".to_string()),
        })
        .target_address("A1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<cacheSource type="olap" connectionId="10"/>"#));
    assert!(cache_def.contains(r#"saveData="0""#));
    assert!(cache_def.contains(r#"recordCount="0""#));
    assert!(cache_def.contains(r#"<cacheField name="Region">"#));
    assert!(cache_def.contains(r#"<cacheField name="Revenue">"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("OlapSales")
        .unwrap();
    assert!(matches!(
        &pivot.source,
        PivotSource::Olap {
            connection_name,
            cube: Some(cube),
            command_text: Some(command_text)
        } if connection_name == "CubeSales"
            && cube == "SalesCube"
            && command_text == "SalesCube"
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert!(matches!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(duke_sheets_core::PivotCacheSourceKind::Olap)
    ));
}

#[test]
fn test_writer_round_trips_scenario_pivot_source_definition() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let pivot = PivotTable::builder("ScenarioSales")
        .source(PivotSource::Scenario {
            name: "BestCase".to_string(),
        })
        .target_address("A1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<cacheSource type="scenario">"#));
    assert!(cache_def.contains(r#"<worksheetSource name="BestCase"/>"#));
    assert!(cache_def.contains(r#"saveData="0""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ScenarioSales")
        .unwrap();
    assert!(matches!(
        &pivot.source,
        PivotSource::Scenario { name } if name == "BestCase"
    ));
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert!(matches!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(duke_sheets_core::PivotCacheSourceKind::Scenario)
    ));
}

#[test]
fn test_writer_round_trips_consolidation_pivot_source_ranges() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let pivot = PivotTable::builder("ConsolidatedSales")
        .source(PivotSource::Consolidation {
            ranges: vec![
                PivotSourceRange::new("North", CellRange::parse("A1:B4").unwrap())
                    .with_name("NorthPlan")
                    .with_page_items(["FY2025", "Plan"]),
                PivotSourceRange::new("South", CellRange::parse("C1:D4").unwrap())
                    .with_name("SouthActual")
                    .with_page_items(["FY2025", "Actual"]),
                PivotSourceRange::named("GlobalNamedSource")
                    .with_page_items(["FY2025", "Plan"]),
                PivotSourceRange::new("ExternalData", CellRange::parse("E1:F4").unwrap())
                    .with_external_relationship_target("file:///C:/data/source.xlsx")
                    .with_page_items(["FY2025", "Actual"]),
            ],
        })
        .target_address("A1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<cacheSource type="consolidation">"#));
    assert!(cache_def.contains(r#"<pages count="2">"#));
    assert!(cache_def.contains(r#"<page count="1"><pageItem name="FY2025"/></page>"#));
    assert!(cache_def.contains(
        r#"<page count="2"><pageItem name="Plan"/><pageItem name="Actual"/></page>"#
    ));
    assert!(cache_def.contains(r#"<rangeSets count="4">"#));
    assert!(cache_def
        .contains(r#"<rangeSet ref="A1:B4" sheet="North" name="NorthPlan" i1="0" i2="0"/>"#));
    assert!(cache_def
        .contains(r#"<rangeSet ref="C1:D4" sheet="South" name="SouthActual" i1="0" i2="1"/>"#));
    assert!(cache_def.contains(r#"<rangeSet name="GlobalNamedSource" i1="0" i2="0"/>"#));
    assert!(cache_def.contains(
        r#"<rangeSet ref="E1:F4" sheet="ExternalData" r:id="rIdExternal1" i1="0" i2="1"/>"#
    ));
    assert!(cache_def.contains(r#"saveData="0""#));
    let cache_rels = read_zip_entry(
        bytes.clone(),
        "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
    );
    assert!(cache_rels.contains(r#"Id="rIdExternal1""#));
    assert!(cache_rels.contains(
        r#"Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath""#
    ));
    assert!(cache_rels.contains(r#"Target="file:///C:/data/source.xlsx""#));
    assert!(cache_rels.contains(r#"TargetMode="External""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ConsolidatedSales")
        .unwrap();
    match &pivot.source {
        PivotSource::Consolidation { ranges } => {
            assert_eq!(ranges.len(), 4);
            assert_eq!(ranges[0].sheet.as_deref(), Some("North"));
            assert_eq!(
                ranges[0].range.map(|range| range.to_a1_string()).as_deref(),
                Some("A1:B4")
            );
            assert_eq!(ranges[0].name.as_deref(), Some("NorthPlan"));
            assert_eq!(ranges[0].page_items, ["FY2025", "Plan"]);
            assert_eq!(ranges[1].sheet.as_deref(), Some("South"));
            assert_eq!(
                ranges[1].range.map(|range| range.to_a1_string()).as_deref(),
                Some("C1:D4")
            );
            assert_eq!(ranges[1].name.as_deref(), Some("SouthActual"));
            assert_eq!(ranges[1].page_items, ["FY2025", "Actual"]);
            assert_eq!(ranges[2].sheet, None);
            assert_eq!(ranges[2].range, None);
            assert_eq!(ranges[2].name.as_deref(), Some("GlobalNamedSource"));
            assert_eq!(ranges[2].page_items, ["FY2025", "Plan"]);
            assert_eq!(ranges[3].sheet.as_deref(), Some("ExternalData"));
            assert_eq!(
                ranges[3].range.map(|range| range.to_a1_string()).as_deref(),
                Some("E1:F4")
            );
            assert_eq!(
                ranges[3].external_relationship_id.as_deref(),
                Some("rIdExternal1")
            );
            assert_eq!(
                ranges[3].external_relationship_target.as_deref(),
                Some("file:///C:/data/source.xlsx")
            );
            assert_eq!(ranges[3].page_items, ["FY2025", "Actual"]);
        }
        other => panic!("unexpected pivot source: {other:?}"),
    }
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert!(matches!(
        pivot.cache_info().map(|info| info.source_kind),
        Some(duke_sheets_core::PivotCacheSourceKind::Consolidation)
    ));
}

#[test]
fn test_writer_round_trips_pivot_field_sort() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();

    let mut region = PivotField::new("Region");
    region.sort = PivotSort::Descending;
    let mut quarter = PivotField::new("Quarter");
    quarter.sort = PivotSort::None;
    let pivot = PivotTable::builder("SortedPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .column(quarter)
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"axis="axisRow" sortType="descending""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SortedPivot")
        .unwrap();
    assert_eq!(pivot.rows[0].sort, PivotSort::Descending);
    assert_eq!(pivot.columns[0].sort, PivotSort::None);
}

#[test]
fn test_writer_round_trips_pivot_field_sort_by_measure() {
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

    let mut region = PivotField::new("Region")
        .with_sort_by(PivotFieldRef::new("Revenue"), PivotAggregate::Sum);
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
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"axis="axisRow" sortType="descending""#));
    assert!(pivot_xml.contains("<autoSortScope>"));
    assert!(pivot_xml.contains(r#"<reference field="4294967294" count="1" selected="0">"#));
    assert!(pivot_xml.contains(r#"<x v="0"/>"#));
    assert!(pivot_xml.contains(
        r#"<rowItems count="3"><i><x v="1"/></i><i><x/></i><i t="grand"><x/></i></rowItems>"#
    ));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ValueSortedPivot")
        .unwrap();
    assert_eq!(pivot.rows[0].sort, PivotSort::Descending);
    let measure = pivot.rows[0]
        .sort_by_measure
        .as_ref()
        .expect("sort measure");
    assert_eq!(measure.field.name, "Revenue");
    assert_eq!(measure.aggregate, PivotAggregate::Sum);
    assert_eq!(measure.name.as_deref(), Some("Sum of Revenue"));
}

#[test]
fn test_writer_round_trips_pivot_axis_field_options() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();

    let mut region = PivotField::new("Region");
    region.caption = Some("Market".to_string());
    region.subtotal = PivotSubtotal::Sum;
    region.subtotal_caption = Some("Subtotal".to_string());
    region.show_empty_items = true;
    region.show_drop_downs = false;
    region.subtotal_top = false;
    region.insert_blank_row = true;
    region.insert_page_break = true;
    region.include_new_items_in_filter = true;
    region.item_page_count = 25;
    region = region.with_collapsed_items(["East"]);
    let mut quarter = PivotField::new("Quarter");
    quarter.subtotal = PivotSubtotal::None;
    let pivot = PivotTable::builder("AxisFieldOptions")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .column(quarter)
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"name="Market""#));
    assert!(pivot_xml.contains(r#"showAll="1""#));
    assert!(pivot_xml.contains(r#"showAll="0""#));
    assert!(pivot_xml.contains(r#"sumSubtotal="1""#));
    assert!(pivot_xml.contains(r#"subtotalCaption="Subtotal""#));
    assert!(pivot_xml.contains(r#"defaultSubtotal="0""#));
    assert!(pivot_xml.contains(r#"showDropDowns="0""#));
    assert!(pivot_xml.contains(r#"subtotalTop="0""#));
    assert!(pivot_xml.contains(r#"insertBlankRow="1""#));
    assert!(pivot_xml.contains(r#"insertPageBreak="1""#));
    assert!(pivot_xml.contains(r#"includeNewItemsInFilter="1""#));
    assert!(pivot_xml.contains(r#"itemPageCount="25""#));
    assert!(pivot_xml.contains(r#"sd="0""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("AxisFieldOptions")
        .unwrap();
    assert_eq!(pivot.rows[0].subtotal, PivotSubtotal::Sum);
    assert_eq!(pivot.rows[0].caption.as_deref(), Some("Market"));
    assert_eq!(pivot.rows[0].subtotal_caption.as_deref(), Some("Subtotal"));
    assert!(pivot.rows[0].show_empty_items);
    assert!(!pivot.rows[0].show_drop_downs);
    assert!(!pivot.rows[0].subtotal_top);
    assert!(pivot.rows[0].insert_blank_row);
    assert!(pivot.rows[0].insert_page_break);
    assert!(pivot.rows[0].include_new_items_in_filter);
    assert_eq!(pivot.rows[0].item_page_count, 25);
    assert_eq!(
        pivot.rows[0].collapsed_items,
        vec![PivotValue::String("East".into())]
    );
    assert_eq!(pivot.columns[0].subtotal, PivotSubtotal::None);
    assert_eq!(pivot.columns[0].caption, None);
    assert_eq!(pivot.columns[0].subtotal_caption, None);
    assert!(!pivot.columns[0].show_empty_items);
    assert!(pivot.columns[0].show_drop_downs);
    assert!(pivot.columns[0].subtotal_top);
    assert!(!pivot.columns[0].insert_blank_row);
    assert!(!pivot.columns[0].insert_page_break);
    assert!(!pivot.columns[0].include_new_items_in_filter);
    assert_eq!(pivot.columns[0].item_page_count, 10);
}

#[test]
fn test_writer_round_trips_pivot_subtotal_functions() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let headers = [
        "Region", "Segment", "Channel", "Bucket", "Group", "Class", "Tier", "Market", "Revenue",
    ];
    for (index, header) in headers.iter().enumerate() {
        sheet.set_cell_value_at(0, index as u16, *header).unwrap();
        sheet
            .set_cell_value_at(1, index as u16, format!("{header} A"))
            .unwrap();
    }
    sheet.set_cell_value("I2", 10.0).unwrap();

    let subtotal_fields = [
        ("Region", PivotSubtotal::Count),
        ("Segment", PivotSubtotal::CountNumbers),
        ("Channel", PivotSubtotal::Product),
        ("Bucket", PivotSubtotal::StdDev),
        ("Group", PivotSubtotal::StdDevP),
        ("Class", PivotSubtotal::Var),
        ("Tier", PivotSubtotal::VarP),
    ];

    let mut builder = PivotTable::builder("SubtotalFunctions")
        .source_range(CellRange::parse("A1:I2").unwrap())
        .target_address("K1")
        .unwrap()
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue");
    let multi_subtotal_field = PivotField::new("Market").with_subtotals([
        PivotSubtotal::Sum,
        PivotSubtotal::Average,
        PivotSubtotal::Max,
    ]);
    builder = builder.row(multi_subtotal_field);
    for (field_name, subtotal) in subtotal_fields {
        let mut field = PivotField::new(field_name);
        field.subtotal = subtotal;
        builder = builder.row(field);
    }
    sheet.add_pivot_table(builder.build().unwrap()).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"sumSubtotal="1""#));
    assert!(pivot_xml.contains(r#"avgSubtotal="1""#));
    assert!(pivot_xml.contains(r#"maxSubtotal="1""#));
    assert!(pivot_xml.contains(r#"countASubtotal="1""#));
    assert!(pivot_xml.contains(r#"countSubtotal="1""#));
    assert!(pivot_xml.contains(r#"productSubtotal="1""#));
    assert!(pivot_xml.contains(r#"stdDevSubtotal="1""#));
    assert!(pivot_xml.contains(r#"stdDevPSubtotal="1""#));
    assert!(pivot_xml.contains(r#"varSubtotal="1""#));
    assert!(pivot_xml.contains(r#"varPSubtotal="1""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SubtotalFunctions")
        .unwrap();
    assert_eq!(pivot.rows[0].subtotal, PivotSubtotal::Sum);
    assert_eq!(
        pivot.rows[0].subtotals,
        vec![
            PivotSubtotal::Sum,
            PivotSubtotal::Average,
            PivotSubtotal::Max
        ]
    );
    assert_eq!(pivot.rows[1].subtotal, PivotSubtotal::Count);
    assert_eq!(pivot.rows[2].subtotal, PivotSubtotal::CountNumbers);
    assert_eq!(pivot.rows[3].subtotal, PivotSubtotal::Product);
    assert_eq!(pivot.rows[4].subtotal, PivotSubtotal::StdDev);
    assert_eq!(pivot.rows[5].subtotal, PivotSubtotal::StdDevP);
    assert_eq!(pivot.rows[6].subtotal, PivotSubtotal::Var);
    assert_eq!(pivot.rows[7].subtotal, PivotSubtotal::VarP);
}

#[test]
fn test_writer_round_trips_pivot_layout_flags() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Tabular;
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
    layout.data_caption = "Metrics".into();
    layout.values_axis = PivotValuesAxis::Rows;
    layout.values_axis_position = Some(1);
    layout.grand_total_caption = Some("Overall".into());
    layout.error_caption = Some("ERR".into());
    layout.show_error = true;
    layout.missing_caption = Some("N/A".into());
    layout.show_missing = false;
    layout.asterisk_totals = true;
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
    layout.subtotal_hidden_items = true;
    layout.show_drop_zones = false;
    layout.indent = 3;
    layout.show_empty_rows = true;
    layout.show_empty_columns = true;
    let pivot = PivotTable::builder("LayoutPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"rowGrandTotals="0""#));
    assert!(pivot_xml.contains(r#"colGrandTotals="0""#));
    assert!(pivot_xml.contains(r#"showHeaders="0""#));
    assert!(pivot_xml.contains(r#"showDrill="0""#));
    assert!(pivot_xml.contains(r#"printDrill="1""#));
    assert!(pivot_xml.contains(r#"itemPrintTitles="1""#));
    assert!(pivot_xml.contains(r#"fieldPrintTitles="1""#));
    assert!(pivot_xml.contains(r#"pageWrap="2""#));
    assert!(pivot_xml.contains(r#"pageOverThenDown="1""#));
    assert!(pivot_xml.contains(r#"mergeItem="1""#));
    assert!(pivot_xml.contains(r#"dataCaption="Metrics""#));
    assert!(pivot_xml.contains(r#"dataOnRows="1""#));
    assert!(pivot_xml.contains(r#"dataPosition="1""#));
    assert!(pivot_xml.contains(r#"grandTotalCaption="Overall""#));
    assert!(pivot_xml.contains(r#"errorCaption="ERR""#));
    assert!(pivot_xml.contains(r#"showError="1""#));
    assert!(pivot_xml.contains(r#"missingCaption="N/A""#));
    assert!(pivot_xml.contains(r#"showMissing="0""#));
    assert!(pivot_xml.contains(r#"asteriskTotals="1""#));
    assert!(pivot_xml.contains(r#"showItems="0""#));
    assert!(pivot_xml.contains(r#"editData="1""#));
    assert!(pivot_xml.contains(r#"disableFieldList="1""#));
    assert!(pivot_xml.contains(r#"showCalcMbrs="0""#));
    assert!(pivot_xml.contains(r#"visualTotals="0""#));
    assert!(pivot_xml.contains(r#"showMultipleLabel="0""#));
    assert!(pivot_xml.contains(r#"showDataDropDown="0""#));
    assert!(pivot_xml.contains(r#"showMemberPropertyTips="0""#));
    assert!(pivot_xml.contains(r#"showDataTips="0""#));
    assert!(pivot_xml.contains(r#"enableWizard="0""#));
    assert!(pivot_xml.contains(r#"enableDrill="0""#));
    assert!(pivot_xml.contains(r#"enableFieldProperties="0""#));
    assert!(pivot_xml.contains(r#"subtotalHiddenItems="1""#));
    assert!(pivot_xml.contains(r#"showDropZones="0""#));
    assert!(pivot_xml.contains(r#"indent="3""#));
    assert!(pivot_xml.contains(r#"showEmptyRow="1""#));
    assert!(pivot_xml.contains(r#"showEmptyCol="1""#));
    assert!(pivot_xml.contains(r#"compact="0""#));
    assert!(pivot_xml.contains(r#"outline="0""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("LayoutPivot")
        .unwrap();
    assert_eq!(pivot.layout.kind, PivotLayoutKind::Tabular);
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
    assert_eq!(pivot.layout.data_caption, "Metrics");
    assert_eq!(pivot.layout.values_axis, PivotValuesAxis::Rows);
    assert_eq!(pivot.layout.values_axis_position, Some(1));
    assert_eq!(pivot.layout.grand_total_caption.as_deref(), Some("Overall"));
    assert_eq!(pivot.layout.error_caption.as_deref(), Some("ERR"));
    assert!(pivot.layout.show_error);
    assert_eq!(pivot.layout.missing_caption.as_deref(), Some("N/A"));
    assert!(!pivot.layout.show_missing);
    assert!(pivot.layout.asterisk_totals);
    assert!(!pivot.layout.show_items);
    assert!(pivot.layout.edit_data);
    assert!(pivot.layout.disable_field_list);
    assert!(!pivot.layout.show_calculated_members);
    assert!(!pivot.layout.visual_totals);
    assert!(!pivot.layout.show_multiple_label);
    assert!(!pivot.layout.show_data_drop_down);
    assert!(!pivot.layout.show_member_property_tips);
    assert!(!pivot.layout.show_data_tips);
    assert!(!pivot.layout.enable_wizard);
    assert!(!pivot.layout.enable_drill);
    assert!(!pivot.layout.enable_field_properties);
    assert!(pivot.layout.subtotal_hidden_items);
    assert!(!pivot.layout.show_drop_zones);
    assert_eq!(pivot.layout.indent, 3);
    assert!(pivot.layout.show_empty_rows);
    assert!(pivot.layout.show_empty_columns);
}

// features: Pivot table styles
#[test]
fn test_writer_round_trips_pivot_style_flags() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();

    let style = PivotStyle {
        name: Some("PivotStyleLight16".to_string()),
        show_row_headers: false,
        show_column_headers: false,
        show_row_stripes: true,
        show_column_stripes: true,
        show_last_column: true,
    };
    let pivot = PivotTable::builder("StylePivot")
        .source_range(CellRange::parse("A1:B2").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .style(style)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"name="PivotStyleLight16""#));
    assert!(pivot_xml.contains(r#"showRowHeaders="0""#));
    assert!(pivot_xml.contains(r#"showColHeaders="0""#));
    assert!(pivot_xml.contains(r#"showRowStripes="1""#));
    assert!(pivot_xml.contains(r#"showColStripes="1""#));
    assert!(pivot_xml.contains(r#"showLastColumn="1""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("StylePivot")
        .unwrap();
    assert_eq!(pivot.style.name.as_deref(), Some("PivotStyleLight16"));
    assert!(!pivot.style.show_row_headers);
    assert!(!pivot.style.show_column_headers);
    assert!(pivot.style.show_row_stripes);
    assert!(pivot.style.show_column_stripes);
    assert!(pivot.style.show_last_column);
}

#[test]
fn test_writer_round_trips_pivot_refresh_policy() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let refresh_policy = PivotRefreshPolicy {
        refresh_on_open: true,
        preserve_formatting: false,
        background_query: true,
        missing_items_limit: Some(5),
    };
    let pivot = PivotTable::builder("RefreshPolicyPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .refresh_policy(refresh_policy.clone())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"preserveFormatting="0""#));
    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"refreshOnLoad="1""#));
    assert!(cache_def.contains(r#"backgroundQuery="1""#));
    assert!(cache_def.contains(r#"missingItemsLimit="5""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RefreshPolicyPivot")
        .unwrap();
    assert_eq!(pivot.refresh_policy, refresh_policy);
}

#[test]
fn test_writer_round_trips_pivot_table_extensions() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();

    let mut pivot = PivotTable::builder("ExtensionPivot")
        .source_range(CellRange::parse("A1:B2").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    pivot.extensions.push(PivotExtension {
        uri: "{pivot-ext}".to_string(),
        payload: br#"<ext uri="{pivot-ext}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:pivotTableDefinition customListSort="1"/></ext>"#.to_vec(),
    });
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"<extLst><ext uri="{pivot-ext}""#));
    assert!(pivot_xml.contains(r#"x15:pivotTableDefinition customListSort="1""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ExtensionPivot")
        .unwrap();
    assert_eq!(pivot.extensions.len(), 1);
    assert_eq!(pivot.extensions[0].uri, "{pivot-ext}");
    let payload = std::str::from_utf8(&pivot.extensions[0].payload).unwrap();
    assert!(payload.contains(r#"uri="{pivot-ext}""#));
    assert!(payload.contains(
        r#"xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main""#
    ));
    assert!(payload.contains(r#"customListSort="1""#));
}

#[test]
fn test_writer_separates_pivot_caches_for_refresh_policy() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let source = CellRange::parse("A1:B3").unwrap();
    let default_policy = PivotTable::builder("DefaultPolicyPivot")
        .source_range(source)
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    let mut refresh_policy = PivotRefreshPolicy::default();
    refresh_policy.missing_items_limit = Some(2);
    let custom_policy = PivotTable::builder("CustomPolicyPivot")
        .source_range(source)
        .target_address("G1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .refresh_policy(refresh_policy)
        .build()
        .unwrap();
    sheet.add_pivot_table(default_policy).unwrap();
    sheet.add_pivot_table(custom_policy).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let first_cache = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    let second_cache = read_zip_entry(bytes, "xl/pivotCache/pivotCacheDefinition2.xml");
    assert!(!first_cache.contains("missingItemsLimit"));
    assert!(second_cache.contains(r#"missingItemsLimit="2""#));
}

#[test]
fn test_writer_round_trips_pivot_page_fields() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Enterprise").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 15.0).unwrap();

    let pivot = PivotTable::builder("SalesBySegment")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .page("Segment")
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Segment", ["Retail"]))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"axis="axisPage""#));
    assert!(pivot_xml.contains(r#"<pageFields count="1"><pageField fld="1" item="0"/>"#));
    assert!(pivot_xml
        .contains(r#"<rowItems count="2"><i><x/></i><i t="grand"><x/></i></rowItems>"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SalesBySegment")
        .unwrap();
    assert_eq!(pivot.page_fields.len(), 1);
    assert_eq!(pivot.page_fields[0].field.name, "Segment");
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.filters.len(), 1);
    match &pivot.filters[0] {
        PivotFilter::FieldItems {
            field,
            allowed_items,
        } => {
            assert_eq!(field.name, "Segment");
            assert_eq!(allowed_items.len(), 1);
            assert_eq!(allowed_items[0].to_string(), "Retail");
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
}

#[test]
fn test_writer_round_trips_pivot_page_field_multi_select() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Enterprise").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", "Public").unwrap();
    sheet.set_cell_value("C4", 15.0).unwrap();

    let pivot = PivotTable::builder("SalesBySegments")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .page("Segment")
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Segment", ["Retail", "Public"]))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"axis="axisPage""#));
    assert!(pivot_xml.contains(r#"multipleItemSelectionAllowed="1""#));
    assert!(pivot_xml.contains(r#"<pageFields count="1"><pageField fld="1"/>"#));
    assert!(pivot_xml.contains(r#"<item x="1" h="1"/>"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SalesBySegments")
        .unwrap();
    assert_eq!(pivot.page_fields.len(), 1);
    assert_eq!(pivot.page_fields[0].field.name, "Segment");
    assert_eq!(pivot.filters.len(), 1);
    match &pivot.filters[0] {
        PivotFilter::FieldItems {
            field,
            allowed_items,
        } => {
            assert_eq!(field.name, "Segment");
            assert_eq!(
                allowed_items
                    .iter()
                    .map(ToString::to_string)
                    .collect::<Vec<_>>(),
                vec!["Retail".to_string(), "Public".to_string()]
            );
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
}

#[test]
fn test_writer_round_trips_pivot_advanced_filters() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
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
    sheet.set_cell_value("B3", "West").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet
        .set_cell_value("A4", date_to_serial(2024, 2, 1, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B4", "North").unwrap();
    sheet.set_cell_value("C4", 30.0).unwrap();

    let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Revenue");
    let date_start = date_to_serial(2024, 1, 1, DateSystem::Date1900);
    let date_end = date_to_serial(2024, 1, 31, DateSystem::Date1900);
    let pivot = PivotTable::builder("FilteredPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .pivot_measure(measure.clone())
        .filter(PivotFilter::Label {
            field: "Region".into(),
            operator: PivotFilterOperator::Contains,
            value: "e".into(),
        })
        .filter(PivotFilter::LabelBetween {
            field: "Region".into(),
            start: "East".into(),
            end: "North".into(),
            not_between: false,
        })
        .filter(PivotFilter::Value {
            field: "Region".into(),
            measure: measure.clone(),
            operator: PivotFilterOperator::GreaterThanOrEqual,
            value: 20.0,
        })
        .filter(PivotFilter::ValueBetween {
            field: "Region".into(),
            measure: measure.clone(),
            start: 10.0,
            end: 30.0,
            not_between: false,
        })
        .filter(PivotFilter::DateBetween {
            field: "Date".into(),
            start: date_start,
            end: date_end,
            not_between: false,
        })
        .filter(PivotFilter::DatePeriod {
            field: "Date".into(),
            period: PivotDatePeriod::Month(1),
        })
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

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"<filters count="7">"#));
    assert!(pivot_xml.contains(r#"type="captionContains""#));
    assert!(pivot_xml.contains(r#"stringValue1="e""#));
    assert!(pivot_xml.contains(r#"type="captionBetween""#));
    assert!(pivot_xml.contains(r#"stringValue1="East""#));
    assert!(pivot_xml.contains(r#"stringValue2="North""#));
    assert!(pivot_xml.contains(r#"type="valueGreaterThanOrEqual""#));
    assert!(pivot_xml.contains(r#"iMeasureFld="0""#));
    assert!(pivot_xml.contains(r#"<customFilter operator="greaterThanOrEqual" val="20"/>"#));
    assert!(pivot_xml.contains(r#"type="valueBetween""#));
    assert!(pivot_xml.contains(r#"stringValue1="10""#));
    assert!(pivot_xml.contains(r#"stringValue2="30""#));
    assert!(pivot_xml.contains(r#"type="dateBetween""#));
    assert!(pivot_xml.contains(&format!(r#"stringValue1="{date_start}""#)));
    assert!(pivot_xml.contains(&format!(r#"stringValue2="{date_end}""#)));
    assert!(pivot_xml.contains(r#"<customFilters and="1">"#));
    assert!(pivot_xml.contains(&format!(
        r#"<customFilter operator="greaterThanOrEqual" val="{date_start}"/>"#
    )));
    assert!(pivot_xml.contains(&format!(
        r#"<customFilter operator="lessThanOrEqual" val="{date_end}"/>"#
    )));
    assert!(pivot_xml.contains(r#"type="M1""#));
    assert!(pivot_xml.contains(r#"type="count""#));
    assert!(pivot_xml.contains(r#"<top10 top="1" percent="0" val="2"/>"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("FilteredPivot")
        .unwrap();
    assert_eq!(pivot.filters.len(), 7);
    match &pivot.filters[0] {
        PivotFilter::Label {
            field,
            operator,
            value,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(*operator, PivotFilterOperator::Contains);
            assert_eq!(value, "e");
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
    match &pivot.filters[1] {
        PivotFilter::LabelBetween {
            field,
            start,
            end,
            not_between,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(start, "East");
            assert_eq!(end, "North");
            assert!(!*not_between);
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
    match &pivot.filters[2] {
        PivotFilter::Value {
            field,
            measure,
            operator,
            value,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(measure.field.name, "Revenue");
            assert_eq!(measure.aggregate, PivotAggregate::Sum);
            assert_eq!(measure.name.as_deref(), Some("Revenue"));
            assert_eq!(*operator, PivotFilterOperator::GreaterThanOrEqual);
            assert_eq!(*value, 20.0);
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
    match &pivot.filters[3] {
        PivotFilter::ValueBetween {
            field,
            measure,
            start,
            end,
            not_between,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(measure.field.name, "Revenue");
            assert_eq!(measure.aggregate, PivotAggregate::Sum);
            assert_eq!(measure.name.as_deref(), Some("Revenue"));
            assert_eq!(*start, 10.0);
            assert_eq!(*end, 30.0);
            assert!(!*not_between);
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
    match &pivot.filters[4] {
        PivotFilter::DateBetween {
            field,
            start,
            end,
            not_between,
        } => {
            assert_eq!(field.name, "Date");
            assert_eq!(*start, date_start);
            assert_eq!(*end, date_end);
            assert!(!*not_between);
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
    match &pivot.filters[5] {
        PivotFilter::DatePeriod { field, period } => {
            assert_eq!(field.name, "Date");
            assert_eq!(*period, PivotDatePeriod::Month(1));
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
    match &pivot.filters[6] {
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
            assert_eq!(measure.name.as_deref(), Some("Revenue"));
            assert_eq!(*n, 2);
            assert!(*top);
            assert!(!*percent);
        }
        other => panic!("unexpected pivot filter: {other:?}"),
    }
}

#[test]
fn test_writer_round_trips_pivot_show_as_percentages() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", 30.0).unwrap();

    let pivot = PivotTable::builder("SalesPercent")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("% of Row")
                .with_show_as(PivotShowAs::PercentOfRowTotal),
        )
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"showDataAs="percentOfRow""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SalesPercent")
        .unwrap();
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].show_as, PivotShowAs::PercentOfRowTotal);
}

#[test]
fn test_writer_round_trips_pivot_measure_number_format() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Rate").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 0.25).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 0.125).unwrap();

    let pivot = PivotTable::builder("RatePivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .pivot_measure(
            PivotMeasure::new("Rate", PivotAggregate::Sum).with_number_format("0.0%"),
        )
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let styles_xml = read_zip_entry(bytes.clone(), "xl/styles.xml");
    assert!(styles_xml.contains(r#"numFmtId="164""#));
    assert!(styles_xml.contains(r#"formatCode="0.0%""#));
    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"numFmtId="164""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("RatePivot")
        .unwrap();
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].number_format.as_deref(), Some("0.0%"));
}

#[test]
fn test_writer_round_trips_pivot_index_show_as() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
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

    let pivot = PivotTable::builder("SalesIndex")
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

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"showDataAs="index""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SalesIndex")
        .unwrap();
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].show_as, PivotShowAs::Index);
}

#[test]
fn test_writer_round_trips_pivot_show_as_base_field() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Period").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", 1.0).unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", 2.0).unwrap();
    sheet.set_cell_value("B3", 15.0).unwrap();

    let pivot = PivotTable::builder("SalesDifference")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
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
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"showDataAs="difference""#));
    assert!(pivot_xml.contains(r#"baseField="0""#));
    assert!(pivot_xml.contains(r#"baseItem="0""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SalesDifference")
        .unwrap();
    assert_eq!(pivot.measures.len(), 1);
    match &pivot.measures[0].show_as {
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        } => {
            assert_eq!(base_field.name, "Period");
            assert_eq!(*base_item, PivotValue::Number(1.0));
        }
        other => panic!("unexpected show-as mode: {other:?}"),
    }
}

#[test]
fn test_writer_round_trips_pivot_rank_show_as() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Period").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", 1.0).unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", 2.0).unwrap();
    sheet.set_cell_value("B3", 15.0).unwrap();
    sheet.set_cell_value("A4", 3.0).unwrap();
    sheet.set_cell_value("B4", 20.0).unwrap();

    let pivot = PivotTable::builder("SalesRank")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
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
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"baseField="0""#));
    assert!(pivot_xml.contains(r#"<x14:dataField "#));
    assert!(pivot_xml.contains(r#"pivotShowAs="rankDescending""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("SalesRank")
        .unwrap();
    assert_eq!(pivot.measures.len(), 1);
    match &pivot.measures[0].show_as {
        PivotShowAs::RankDescending { base_field } => {
            assert_eq!(base_field.name, "Period");
        }
        other => panic!("unexpected show-as mode: {other:?}"),
    }
}

#[test]
fn test_writer_round_trips_pivot_parent_show_as() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Year").unwrap();
    sheet.set_cell_value("D1", "Quarter").unwrap();
    sheet.set_cell_value("E1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", "2024").unwrap();
    sheet.set_cell_value("D2", "Q1").unwrap();
    sheet.set_cell_value("E2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", "2024").unwrap();
    sheet.set_cell_value("D3", "Q2").unwrap();
    sheet.set_cell_value("E3", 30.0).unwrap();

    let pivot = PivotTable::builder("ParentShowAs")
        .source_range(CellRange::parse("A1:E3").unwrap())
        .target_address("G1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .column("Year")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("% Parent Row")
                .with_show_as(PivotShowAs::PercentOfParentRowTotal),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("% Parent Column")
                .with_show_as(PivotShowAs::PercentOfParentColumnTotal),
        )
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("% Parent Region")
                .with_show_as(PivotShowAs::PercentOfParentTotal {
                    base_field: "Region".into(),
                }),
        )
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"pivotShowAs="percentOfParentRow""#));
    assert!(pivot_xml.contains(r#"pivotShowAs="percentOfParentCol""#));
    assert!(pivot_xml.contains(r#"pivotShowAs="percentOfParent""#));
    assert!(pivot_xml.contains(r#"sourceField="0""#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ParentShowAs")
        .unwrap();
    assert_eq!(pivot.measures.len(), 3);
    assert_eq!(
        pivot.measures[0].show_as,
        PivotShowAs::PercentOfParentRowTotal
    );
    assert_eq!(
        pivot.measures[1].show_as,
        PivotShowAs::PercentOfParentColumnTotal
    );
    match &pivot.measures[2].show_as {
        PivotShowAs::PercentOfParentTotal { base_field } => {
            assert_eq!(base_field.name, "Region");
        }
        other => panic!("unexpected show-as mode: {other:?}"),
    }
}

// features: Calculated fields
#[test]
fn test_writer_round_trips_pivot_calculated_fields() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Units").unwrap();
    sheet.set_cell_value("C1", "Price").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 2.0).unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 7.0).unwrap();
    sheet.set_cell_value("C3", 3.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRevenue")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .calculated_field("Revenue", "=Units*Price")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def
        .contains(r#"<cacheField name="Revenue" formula="Units*Price" databaseField="0">"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRevenue")
        .unwrap();
    assert_eq!(pivot.calculated_fields.len(), 1);
    assert_eq!(pivot.calculated_fields[0].name, "Revenue");
    assert_eq!(pivot.calculated_fields[0].formula, "Units*Price");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

// features: Calculated fields
#[test]
fn test_writer_round_trips_pivot_calculated_field_workbook_reference() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Units").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 2.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 7.0).unwrap();
    sheet.set_cell_value("J1", 3.0).unwrap();
    let pivot = PivotTable::builder("WorkbookRef")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_field("Adjusted", "=Units*$J$1")
        .measure("Adjusted", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb.refresh_pivots().unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).unwrap();
    let mut roundtrip = XlsxReader::read(Cursor::new(out.into_inner())).unwrap();
    let formula =
        &roundtrip.worksheet(0).unwrap().pivot_tables()[0].calculated_fields[0].formula;
    assert_eq!(formula, "Units*$J$1");
    roundtrip.refresh_pivots().unwrap();
    assert_eq!(
        roundtrip
            .worksheet(0)
            .unwrap()
            .get_value("E2")
            .unwrap()
            .as_number(),
        Some(6.0)
    );
    assert_eq!(
        roundtrip
            .worksheet(0)
            .unwrap()
            .get_value("E3")
            .unwrap()
            .as_number(),
        Some(21.0)
    );
}

#[test]
fn test_writer_serializes_refreshable_consolidation_records() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().set_name("North");
    let south = wb.add_worksheet().unwrap();
    wb.worksheet_mut(south).unwrap().set_name("South");
    for sheet_index in [0, south] {
        let sheet = wb.worksheet_mut(sheet_index).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet
            .set_cell_value("A2", if sheet_index == 0 { "East" } else { "West" })
            .unwrap();
        sheet
            .set_cell_value("B2", if sheet_index == 0 { 10.0 } else { 20.0 })
            .unwrap();
    }
    let pivot = PivotTable::builder("Consolidated")
        .source(PivotSource::Consolidation {
            ranges: vec![
                PivotSourceRange::new("North", CellRange::parse("A1:B2").unwrap()),
                PivotSourceRange::new("South", CellRange::parse("A1:B2").unwrap()),
            ],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).unwrap();
    let bytes = out.into_inner();
    let definition = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(definition.contains(r#"recordCount="2" saveData="1""#));
    let records = read_zip_entry(bytes, "xl/pivotCache/pivotCacheRecords1.xml");
    assert!(records.contains(r#"<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">"#));
    assert_eq!(records.matches("<r>").count(), 2);
}

#[test]
fn test_writer_avoids_preserved_pivot_relationship_part_collision() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet
        .add_pivot_table(
            PivotTable::builder("Sales")
                .source_range(CellRange::parse("A1:B2").unwrap())
                .target_address("D1")
                .unwrap()
                .row("Region")
                .measure("Revenue", PivotAggregate::Sum)
                .build()
                .unwrap(),
        )
        .unwrap();
    wb.workbook_extension_parts_mut().push(WorkbookExtensionPart::new(
        "xl/pivotTables/_rels/pivotTable1.xml.rels",
        "application/vnd.openxmlformats-package.relationships+xml",
        "http://example.com/preserved-pivot-rel",
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#.to_vec(),
    ));

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).unwrap();
    let bytes = out.into_inner();
    assert!(read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable2.xml").contains("Sales"));
    let rels = read_zip_entry(bytes, "xl/pivotTables/_rels/pivotTable2.xml.rels");
    assert!(rels.contains("pivotCacheDefinition1.xml"));
}

#[test]
fn test_writer_round_trips_table_qualified_pivot_calculated_fields() {
    use duke_sheets_core::table::{Table, TableColumn};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Units").unwrap();
    sheet.set_cell_value("C1", "Price").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 2.0).unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 7.0).unwrap();
    sheet.set_cell_value("C3", 3.0).unwrap();

    let mut table = Table::new(1, "SalesData", CellRange::parse("A1:C3").unwrap());
    table.columns = vec![
        TableColumn::new(1, "Region"),
        TableColumn::new(2, "Units"),
        TableColumn::new(3, "Price"),
    ];
    sheet.add_table(table);

    let pivot = PivotTable::builder("CalculatedTableRevenue")
        .table_source("SalesData")
        .target_address("E1")
        .unwrap()
        .row("Region")
        .calculated_field("Revenue", "=SalesData[@Units]*SalesData[@Price]")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(
        r#"<cacheField name="Revenue" formula="SalesData[@Units]*SalesData[@Price]" databaseField="0">"#
    ));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedTableRevenue")
        .unwrap();
    assert_eq!(pivot.calculated_fields.len(), 1);
    assert_eq!(pivot.calculated_fields[0].name, "Revenue");
    assert_eq!(
        pivot.calculated_fields[0].formula,
        "SalesData[@Units]*SalesData[@Price]"
    );
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

#[test]
fn test_writer_round_trips_escaped_table_qualified_pivot_calculated_fields() {
    use duke_sheets_core::table::{Table, TableColumn};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Gross Sales").unwrap();
    sheet.set_cell_value("C1", "Rate").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 100.0).unwrap();
    sheet.set_cell_value("C2", 0.1).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 80.0).unwrap();
    sheet.set_cell_value("C3", 0.25).unwrap();

    let mut table = Table::new(1, "SalesData", CellRange::parse("A1:C3").unwrap());
    table.columns = vec![
        TableColumn::new(1, "Region"),
        TableColumn::new(2, "Gross Sales"),
        TableColumn::new(3, "Rate"),
    ];
    sheet.add_table(table);

    let pivot = PivotTable::builder("CalculatedTableCommission")
        .table_source("SalesData")
        .target_address("E1")
        .unwrap()
        .row("Region")
        .calculated_field("Commission", "=SalesData[@[Gross Sales]]*SalesData[@Rate]")
        .named_measure("Commission", PivotAggregate::Sum, "Commission")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(
        r#"<cacheField name="Commission" formula="SalesData[@[Gross Sales]]*SalesData[@Rate]" databaseField="0">"#
    ));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedTableCommission")
        .unwrap();
    assert_eq!(pivot.calculated_fields.len(), 1);
    assert_eq!(pivot.calculated_fields[0].name, "Commission");
    assert_eq!(
        pivot.calculated_fields[0].formula,
        "SalesData[@[Gross Sales]]*SalesData[@Rate]"
    );
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Commission");
}

#[test]
fn test_writer_round_trips_pivot_calculated_items() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRegion")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_item("Region", "Combined", "East+West")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<s v="Combined"/>"#));
    assert!(cache_def.contains(r#"<calculatedItems count="1">"#));
    assert!(cache_def.contains(r#"<calculatedItem field="0" formula="East+West">"#));
    assert!(cache_def.contains(r#"<pivotArea field="0" cacheIndex="1">"#));
    assert!(cache_def.contains(r#"<x v="2"/>"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("CalculatedRegion")
        .unwrap();
    assert_eq!(pivot.calculated_items.len(), 1);
    assert_eq!(pivot.calculated_items[0].field.name, "Region");
    assert_eq!(
        pivot.calculated_items[0].item,
        PivotValue::String("Combined".into())
    );
    assert_eq!(pivot.calculated_items[0].formula, "East+West");
}

// features: Grouping (dates, numbers, items)
#[test]
fn test_writer_round_trips_pivot_grouping() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Amount").unwrap();
    sheet.set_cell_value("B1", "SaleDate").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", 2.0).unwrap();
    sheet.set_cell_value("B2", 45292.0).unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", 12.0).unwrap();
    sheet.set_cell_value("B3", 45323.0).unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", 22.0).unwrap();
    sheet.set_cell_value("B4", 45352.0).unwrap();
    sheet.set_cell_value("C4", 30.0).unwrap();

    let pivot = PivotTable::builder("GroupedSales")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Amount")
        .column("SaleDate")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
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

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(
        r#"<fieldGroup base="0"><rangePr autoStart="0" autoEnd="0" startNum="0" endNum="30" groupInterval="10"/><groupItems count="5"><s v="&lt;0"/><s v="0-9"/><s v="10-19"/><s v="20-30"/><s v="&gt;30"/>"#
    ));
    assert!(cache_def.contains(r#"<fieldGroup par="3"/>"#));
    assert!(cache_def.contains(
        r#"<cacheField name="Months (SaleDate)" numFmtId="0" databaseField="0"><fieldGroup base="1"><rangePr groupBy="months" startDate="2024-01-01T00:00:00" endDate="2024-03-02T00:00:00"/><groupItems count="14"><s v="&lt;1/1/2024"/><s v="Jan"/><s v="Feb"/><s v="Mar"/>"#
    ));
    for source_value in ["2", "12", "22"] {
        assert!(
            cache_def.contains(&format!(r#"<n v="{source_value}"/>"#)),
            "grouped cache lost source item {source_value}: {cache_def}"
        );
    }
    for source_date in ["2024-01-01", "2024-02-01", "2024-03-01"] {
        assert!(cache_def.contains(&format!(r#"<d v="{source_date}T00:00:00"/>"#)));
    }
    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml
        .contains(r#"<rowItems count="4"><i><x v="1"/></i><i><x v="2"/></i><i><x v="3"/></i>"#));
    assert!(pivot_xml
        .contains(r#"<colItems count="4"><i><x v="1"/></i><i><x v="2"/></i><i><x v="3"/></i>"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("GroupedSales")
        .unwrap();
    assert_eq!(pivot.groupings.len(), 2);

    let amount_grouping = pivot
        .groupings
        .iter()
        .find(|grouping| {
            matches!(
                grouping,
                PivotGrouping::Number { field, .. } if field.name == "Amount"
            )
        })
        .expect("amount grouping");
    match amount_grouping {
        PivotGrouping::Number {
            start,
            end,
            interval,
            ..
        } => {
            assert_eq!(*start, Some(0.0));
            assert_eq!(*end, Some(30.0));
            assert_eq!(*interval, 10.0);
        }
        other => panic!("unexpected grouping: {other:?}"),
    }

    let date_grouping = pivot
        .groupings
        .iter()
        .find(|grouping| {
            matches!(
                grouping,
                PivotGrouping::Date { field, .. } if field.name == "SaleDate"
            )
        })
        .expect("date grouping");
    match date_grouping {
        PivotGrouping::Date { units, .. } => {
            assert_eq!(*units, vec![PivotDateGroupUnit::Months]);
        }
        other => panic!("unexpected grouping: {other:?}"),
    }
}

#[test]
fn test_writer_round_trips_manual_pivot_grouping() {
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

    let pivot = PivotTable::builder("ManualGroupedRegions")
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

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<fieldGroup par="2"/></cacheField>"#));
    assert!(cache_def.contains(
        r#"<cacheField name="Region2" numFmtId="0" databaseField="0"><fieldGroup base="0"><discretePr count="3"><x v="1"/><x v="1"/><x v="0"/></discretePr><groupItems count="2"><s v="Central"/><s v="Coastal"/></groupItems></fieldGroup></cacheField>"#
    ));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
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
        other => panic!("unexpected grouping: {other:?}"),
    }
}

#[test]
fn test_writer_round_trips_multi_unit_date_grouping() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "SaleDate").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", 45292.0).unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", 45323.0).unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", 45658.0).unwrap();
    sheet.set_cell_value("B4", 30.0).unwrap();

    let pivot = PivotTable::builder("GroupedDates")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("SaleDate")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .grouping(PivotGrouping::Date {
            field: "SaleDate".into(),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).expect("write workbook");
    let bytes = out.into_inner();

    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<cacheFields count="4">"#));
    assert!(cache_def
        .contains(r#"<cacheField name="Years (SaleDate)" numFmtId="0" databaseField="0">"#));
    assert!(cache_def
        .contains(r#"<cacheField name="Months (SaleDate)" numFmtId="0" databaseField="0">"#));
    assert!(
        cache_def.contains(r#"<fieldGroup base="0"><rangePr groupBy="years" startDate="2024-01-01T00:00:00" endDate="2025-01-02T00:00:00"/>"#)
    );
    assert!(cache_def.contains(r#"<fieldGroup base="0"><rangePr groupBy="months" startDate="2024-01-01T00:00:00" endDate="2025-01-02T00:00:00"/>"#));
    assert!(cache_def.contains(r#"<groupItems count="4"><s v="&lt;1/1/2024"/><s v="2024"/><s v="2025"/><s v="&gt;1/2/2025"/>"#));

    let pivot_xml = read_zip_entry(bytes.clone(), "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(r#"<rowFields count="2">"#));
    assert!(pivot_xml.contains(r#"<field x="2"/><field x="3"/>"#));

    let roundtrip = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let pivot = roundtrip
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
        other => panic!("unexpected grouping: {other:?}"),
    }
}

#[test]
fn test_writer_emits_fractional_numeric_group_boundaries_and_extremes() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Amount").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    for (row, amount) in [-1.25, 0.25, 1.75, 3.0].into_iter().enumerate() {
        let row = row + 2;
        sheet.set_cell_value(&format!("A{row}"), amount).unwrap();
        sheet.set_cell_value(&format!("B{row}"), 1.0).unwrap();
    }
    let pivot = PivotTable::builder("FractionGroups")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Amount")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Number {
            field: "Amount".into(),
            start: Some(0.0),
            end: Some(2.0),
            interval: 0.5,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).unwrap();
    let bytes = out.into_inner();
    let cache_def = read_zip_entry(bytes.clone(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(
        r#"<groupItems count="6"><s v="&lt;0"/><s v="0-0.5"/><s v="0.5-1"/><s v="1-1.5"/><s v="1.5-2"/><s v="&gt;2"/>"#
    ));
    let pivot_xml = read_zip_entry(bytes, "xl/pivotTables/pivotTable1.xml");
    assert!(pivot_xml.contains(
        r#"<rowItems count="5"><i><x/></i><i><x v="1"/></i><i><x v="4"/></i><i><x v="5"/></i>"#
    ));
}

#[test]
fn test_writer_emits_1904_date_times_as_typed_cache_items() {
    let mut wb = Workbook::new();
    wb.settings_mut().date_1904 = true;
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "OccurredAt").unwrap();
    sheet.set_cell_value("B1", "Count").unwrap();
    sheet.set_cell_value("A2", 0.25).unwrap();
    sheet.set_cell_value("B2", 1.0).unwrap();
    sheet.set_cell_value("A3", 0.75).unwrap();
    sheet.set_cell_value("B3", 1.0).unwrap();
    let pivot = PivotTable::builder("HourlyGroups")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("OccurredAt")
        .measure("Count", PivotAggregate::Sum)
        .grouping(PivotGrouping::Date {
            field: "OccurredAt".into(),
            units: vec![PivotDateGroupUnit::Hours],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let mut out = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut out).unwrap();
    let cache_def = read_zip_entry(out.into_inner(), "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(cache_def.contains(r#"<d v="1904-01-01T06:00:00"/>"#));
    assert!(cache_def.contains(r#"<d v="1904-01-01T18:00:00"/>"#));
    assert!(cache_def.contains(
        r#"<rangePr groupBy="hours" startDate="1904-01-01T06:00:00" endDate="1904-01-02T00:00:00"/>"#
    ));
}
