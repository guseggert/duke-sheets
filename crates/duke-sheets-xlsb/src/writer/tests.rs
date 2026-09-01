#[cfg(test)]
mod tests {
    use std::io::{Cursor, Read, Write};

    use duke_sheets_core::style::{
        BorderEdge, BorderLineStyle, Color, FillStyle, NumberFormat, Style,
    };
    use duke_sheets_core::table::{Table, TableColumn};
    use duke_sheets_core::{
        hash_legacy_protection_password, CellRange, CellValue, PivotAggregate, PivotDateGroupUnit,
        PivotDatePeriod, PivotField, PivotFieldRef, PivotFilter, PivotFilterOperator,
        PivotGrouping, PivotLayout, PivotLayoutKind, PivotManualGroup, PivotMeasure,
        PivotRefreshPolicy, PivotShowAs, PivotSort, PivotSource, PivotSourceRange, PivotStyle,
        PivotSubtotal, PivotTable, PivotValue, PivotValuesAxis, ProtectedRange, Workbook,
        WorkbookConnection, WorkbookConnectionCredentials, WorkbookConnectionKind,
        WorkbookConnectionParameter, WorkbookConnectionParameterValue, WorkbookProtection,
    };

    use crate::reader::XlsbReader;
    use crate::writer::XlsbWriter;

    fn round_trip(wb: &Workbook) -> Workbook {
        let mut buf = Vec::new();
        XlsbWriter::write(wb, Cursor::new(&mut buf)).unwrap();
        XlsbReader::read(Cursor::new(&buf)).unwrap()
    }

    fn sxview_flag(payload: &[u8], bit: usize) -> bool {
        payload[4 + bit / 8] & (1u8 << (bit % 8)) != 0
    }

    fn add_test_pivot(wb: &mut Workbook) {
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Sales").unwrap();
        ws.set_cell_value("A2", "East").unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();
        ws.set_cell_value("A3", "West").unwrap();
        ws.set_cell_value("B3", 20.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B3").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_table_source_pivot(wb: &mut Workbook) {
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Sales").unwrap();
        ws.set_cell_value("A2", "East").unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();
        ws.set_cell_value("A3", "West").unwrap();
        ws.set_cell_value("B3", 20.0).unwrap();

        let mut table = Table::new(1, "SalesData", CellRange::parse("A1:B3").unwrap());
        table.columns = vec![TableColumn::new(1, "Region"), TableColumn::new(2, "Sales")];
        ws.add_table(table);

        let pivot = PivotTable::builder("SalesTablePivot")
            .table_source("SalesData")
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_consolidation_source_pivot(wb: &mut Workbook) {
        wb.add_worksheet_with_name("WestData").unwrap();
        {
            let ws = wb.worksheet_mut(0).unwrap();
            ws.set_cell_value("A1", "Region").unwrap();
            ws.set_cell_value("B1", "Sales").unwrap();
            ws.set_cell_value("A2", "East").unwrap();
            ws.set_cell_value("B2", 10.0).unwrap();
            ws.set_cell_value("A3", "East").unwrap();
            ws.set_cell_value("B3", 15.0).unwrap();
        }
        {
            let ws = wb.worksheet_mut(1).unwrap();
            ws.set_cell_value("A1", "Region").unwrap();
            ws.set_cell_value("B1", "Sales").unwrap();
            ws.set_cell_value("A2", "West").unwrap();
            ws.set_cell_value("B2", 20.0).unwrap();
        }

        let pivot = PivotTable::builder("ConsolidatedSales")
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
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_named_consolidation_source_pivot(wb: &mut Workbook) {
        let pivot = PivotTable::builder("NamedConsolidatedSales")
            .source(PivotSource::Consolidation {
                ranges: vec![PivotSourceRange::named("NamedSalesSource")],
            })
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_external_consolidation_source_pivot(wb: &mut Workbook) {
        let pivot = PivotTable::builder("ExternalConsolidatedSales")
            .source(PivotSource::Consolidation {
                ranges: vec![PivotSourceRange::new(
                    "ExternalData",
                    CellRange::parse("A1:B3").unwrap(),
                )
                .with_external_relationship_id("rIdExternalSource")
                .with_external_relationship_target("file:///C:/data/source.xlsx")],
            })
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_external_connection_pivot(wb: &mut Workbook) {
        wb.add_data_connection(
            WorkbookConnection::database(7, "SalesConnection", "Provider=MSDASQL;DSN=Sales;")
                .with_command("select Region, Sales from Sales")
                .with_command_type(2),
        )
        .unwrap();

        let pivot = PivotTable::builder("ExternalSales")
            .source(PivotSource::External {
                connection_name: "SalesConnection".to_string(),
                command_text: Some("select Region, Sales from Sales".to_string()),
            })
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_olap_connection_pivot(wb: &mut Workbook) {
        let mut connection = WorkbookConnection::olap(10, "CubeSales").with_connection_type(5);
        connection.kind = WorkbookConnectionKind::Olap {
            connection: Some("Provider=MSOLAP;Data Source=olapserver;".to_string()),
            command: Some("SalesCube".to_string()),
            command_type: Some(1),
            local: false,
            local_connection: None,
            local_refresh: true,
            send_locale: true,
            row_drill_count: Some(1000),
        };
        wb.add_data_connection(connection).unwrap();

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
    }

    fn add_column_test_pivot(wb: &mut Workbook) {
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

    fn add_page_test_pivot(wb: &mut Workbook) {
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
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_numeric_row_filter_pivot(wb: &mut Workbook) {
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Bucket").unwrap();
        ws.set_cell_value("B1", "Region").unwrap();
        ws.set_cell_value("C1", "Revenue").unwrap();
        ws.set_cell_value("A2", 1.0).unwrap();
        ws.set_cell_value("B2", "East").unwrap();
        ws.set_cell_value("C2", 10.0).unwrap();
        ws.set_cell_value("A3", 2.0).unwrap();
        ws.set_cell_value("B3", "West").unwrap();
        ws.set_cell_value("C3", 20.0).unwrap();
        ws.set_cell_value("A4", 1.0).unwrap();
        ws.set_cell_value("B4", "East").unwrap();
        ws.set_cell_value("C4", 30.0).unwrap();

        let pivot = PivotTable::builder("NumericRowFilter")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Bucket")
            .filter(PivotFilter::field_items(
                "Bucket",
                [PivotValue::Number(2.0)],
            ))
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
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
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_multi_measure_column_values_pivot(wb: &mut Workbook) {
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

    fn add_all_aggregate_column_values_pivot(wb: &mut Workbook) {
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
            .column("Quarter")
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

    fn add_x14_show_as_pivot(wb: &mut Workbook) {
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

    fn add_page_date_grouped_pivot(wb: &mut Workbook, allowed_items: &[f64]) {
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
                field: PivotFieldRef::new("Date"),
                units: vec![PivotDateGroupUnit::Months],
            });
        if !allowed_items.is_empty() {
            pivot = pivot.filter(PivotFilter::field_items(
                "Date",
                allowed_items.iter().copied(),
            ));
        }
        let pivot = pivot.build().unwrap();
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

    fn add_multi_unit_date_grouped_page_pivot(wb: &mut Workbook) {
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
                field: PivotFieldRef::new("SaleDate"),
                units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_multi_unit_date_grouped_page_filter_pivot(wb: &mut Workbook) {
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
                field: PivotFieldRef::new("SaleDate"),
                units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
            })
            .filter(PivotFilter::field_items("SaleDate (Years)", [2024.0]))
            .filter(PivotFilter::field_items("SaleDate (Months)", [1.0, 2.0]))
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_manual_grouped_pivot(wb: &mut Workbook) {
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
                field: PivotFieldRef::new("Region"),
                groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    }

    fn add_raw_drawing(ws: &mut duke_sheets_core::Worksheet, bytes: Vec<u8>) {
        use duke_sheets_core::{DrawingObject, RawDrawing};
        ws.add_drawing(DrawingObject::raw(RawDrawing {
            bytes,
            rels: vec![],
        })).unwrap();
    }

    fn raw_drawing_bytes(ws: &duke_sheets_core::Worksheet) -> Vec<&Vec<u8>> {
        ws.drawings()
            .iter()
            .filter_map(|o| match &o.kind {
                duke_sheets_core::DrawingKind::Raw(raw) => Some(&raw.bytes),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn empty_workbook() {
        let wb = Workbook::new();
        let wb2 = round_trip(&wb);
        assert_eq!(wb2.sheet_count(), 1);
    }

    #[test]
    fn semantic_pivot_tables_emit_native_parts() {
        let mut wb = Workbook::new();
        add_test_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let ct = read_zip_entry(&bytes, "[Content_Types].xml");
        assert!(ct.contains("/xl/pivotTables/pivotTable1.bin"));
        assert!(ct.contains("/xl/pivotCache/pivotCacheDefinition1.bin"));
        assert!(ct.contains("/xl/pivotCache/pivotCacheRecords1.bin"));
        assert!(ct.contains("/xl/worksheets/binaryIndex1.bin"));

        let workbook_rels = read_zip_entry(&bytes, "xl/_rels/workbook.bin.rels");
        assert!(workbook_rels.contains("pivotCache/pivotCacheDefinition1.bin"));

        let sheet_rels = read_zip_entry(&bytes, "xl/worksheets/_rels/sheet1.bin.rels");
        assert!(sheet_rels.contains("xlBinaryIndex"));
        assert!(sheet_rels.contains("../pivotTables/pivotTable1.bin"));

        let pivot_rels = read_zip_entry(&bytes, "xl/pivotTables/_rels/pivotTable1.bin.rels");
        assert!(pivot_rels.contains("../pivotCache/pivotCacheDefinition1.bin"));

        let cache_rels =
            read_zip_entry(&bytes, "xl/pivotCache/_rels/pivotCacheDefinition1.bin.rels");
        assert!(cache_rels.contains("pivotCacheRecords1.bin"));

        let workbook_records = record_types(read_zip_entry_bytes(&bytes, "xl/workbook.bin"));
        assert!(workbook_records.contains(&crate::biff12::records::BRT_BEGIN_PIVOT_CACHE_IDS));
        assert!(workbook_records.contains(&crate::biff12::records::BRT_PIVOT_CACHE_ID));

        let cache_def_records = record_types(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        assert!(cache_def_records.contains(&crate::biff12::records::BRT_BEGIN_PIVOT_CACHE_DEF));
        assert!(cache_def_records.contains(&crate::biff12::records::BRT_BEGIN_PCD_FIELDS));
        assert!(cache_def_records.contains(&crate::biff12::records::BRT_PCDI_STRING));

        let cache_records = record_types(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheRecords1.bin",
        ));
        assert!(cache_records.contains(&crate::biff12::records::BRT_BEGIN_PCD_RECORDS));
        assert_eq!(
            cache_records
                .iter()
                .filter(|&&typ| typ == crate::biff12::records::BRT_PCD_RECORD)
                .count(),
            2
        );

        let pivot_records = record_types(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        assert!(pivot_records.contains(&crate::biff12::records::BRT_BEGIN_SXVIEW));
        assert!(pivot_records.contains(&crate::biff12::records::BRT_BEGIN_SXVDS));
        assert!(pivot_records.contains(&crate::biff12::records::BRT_BEGIN_SX_ROW_ITEMS));
        assert!(pivot_records.contains(&crate::biff12::records::BRT_BEGIN_SXDIS));
    }

    #[test]
    fn reads_xlsb_pivot_cache_records_for_metadata_only_numeric_items() {
        let mut wb = Workbook::new();
        add_numeric_row_filter_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_def = metadata_light_first_numeric_cache_field(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let cache_records = inline_first_numeric_cache_record_field(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheRecords1.bin",
        ));
        let bytes = replace_zip_entries(
            &bytes,
            &[
                ("xl/pivotCache/pivotCacheDefinition1.bin", cache_def),
                ("xl/pivotCache/pivotCacheRecords1.bin", cache_records),
            ],
        );

        let wb2 = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];
        assert_eq!(pivot.name, "NumericRowFilter");
        assert_eq!(pivot.rows[0].field.name, "Bucket");
        assert_eq!(
            pivot.cache_info().and_then(|info| info.record_count),
            Some(3)
        );
    }

    #[test]
    fn reads_xlsb_external_pivot_cache_source_header() {
        let mut wb = Workbook::new();
        add_test_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_def = patch_pcd_source_header(
            read_zip_entry_bytes(&bytes, "xl/pivotCache/pivotCacheDefinition1.bin"),
            1,
            7,
        );
        let content_types =
            add_connections_content_type(read_zip_entry(&bytes, "[Content_Types].xml"));
        let workbook_rels =
            add_connections_workbook_rel(read_zip_entry(&bytes, "xl/_rels/workbook.bin.rels"));
        let bytes = replace_zip_entries(
            &bytes,
            &[
                ("[Content_Types].xml", content_types),
                ("xl/_rels/workbook.bin.rels", workbook_rels),
                ("xl/pivotCache/pivotCacheDefinition1.bin", cache_def),
                ("xl/connections.bin", xlsb_database_connections_bin()),
            ],
        );

        let wb2 = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let connection = wb2.data_connection_by_id(7).expect("XLSB connection");
        assert_eq!(connection.name, "SalesConnection");
        match &connection.kind {
            duke_sheets_core::WorkbookConnectionKind::Database {
                connection,
                command,
                command_type,
            } => {
                assert_eq!(connection, "Provider=MSDASQL;DSN=Sales;");
                assert_eq!(command.as_deref(), Some("select * from Sales"));
                assert_eq!(*command_type, Some(2));
            }
            other => panic!("unexpected connection kind: {other:?}"),
        }
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];
        assert!(matches!(
            &pivot.source,
            PivotSource::External {
                connection_name,
                command_text: Some(command_text),
            } if connection_name == "SalesConnection" && command_text == "select * from Sales"
        ));
        assert_eq!(pivot.rows[0].field.name, "Region");
        assert_eq!(pivot.measures[0].field.name, "Sales");
        assert!(matches!(
            pivot.cache_info().map(|info| info.source_kind),
            Some(duke_sheets_core::PivotCacheSourceKind::External)
        ));
    }

    #[test]
    fn xlsb_writer_round_trips_external_pivot_source_authoring() {
        let mut wb = Workbook::new();
        add_external_connection_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        assert!(read_zip_entry(&bytes, "[Content_Types].xml").contains("/xl/connections.bin"));
        assert!(read_zip_entry(&bytes, "xl/_rels/workbook.bin.rels").contains("connections.bin"));
        let connection_records =
            records_with_payload(read_zip_entry_bytes(&bytes, "xl/connections.bin"));
        let connection_payload = connection_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_EXT_CONNECTION)
                    .then_some(payload)
            })
            .expect("external connection header payload");
        assert_eq!(
            connection_payload[2], 0x02,
            "XLSB external connections must encode pc=2 when no password is saved"
        );
        assert_eq!(
            u16::from_le_bytes([connection_payload[8], connection_payload[9]]) & 0x0008,
            0x0008,
            "XLSB external connections must set reserved flag bit K"
        );

        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let source = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_SOURCE).then_some(payload)
            })
            .expect("PCD source payload");
        assert_eq!(
            source,
            &[1, 0, 0, 0, 7, 0, 0, 0],
            "external XLSB pivot sources should reference the workbook connection id"
        );

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = &read.worksheet(0).unwrap().pivot_tables()[0];
        assert!(matches!(
            &pivot.source,
            PivotSource::External {
                connection_name,
                command_text: Some(command_text),
            } if connection_name == "SalesConnection" && command_text == "select Region, Sales from Sales"
        ));
    }

    #[test]
    fn xlsb_writer_rejects_cell_bound_connection_parameter_authoring() {
        let mut wb = Workbook::new();
        let mut parameter = WorkbookConnectionParameter::cell("MinRevenue", "Sheet1!$A$1");
        parameter.sql_type = 8;
        parameter.refresh_on_change = true;
        wb.add_data_connection(
            WorkbookConnection::database(7, "SalesConnection", "Provider=MSDASQL;DSN=Sales;")
                .with_command("select Region, Sales from Sales")
                .with_command_type(2)
                .with_parameter(parameter),
        )
        .unwrap();

        let mut out = Cursor::new(Vec::new());
        let err = XlsbWriter::write(&wb, &mut out)
            .expect_err("cell-bound XLSB connection parameters should be rejected");
        assert!(
            err.to_string()
                .contains("cell-bound parameter authoring is not supported"),
            "{err}"
        );
    }

    #[test]
    fn xlsb_writer_round_trips_external_pivot_cache_background_query() {
        let mut wb = Workbook::new();
        wb.add_data_connection(
            WorkbookConnection::database(7, "SalesConnection", "Provider=MSDASQL;DSN=Sales;")
                .with_command("select Region, Sales from Sales")
                .with_command_type(2),
        )
        .unwrap();

        let refresh_policy = PivotRefreshPolicy {
            refresh_on_open: false,
            preserve_formatting: true,
            background_query: true,
            missing_items_limit: None,
        };
        let pivot = PivotTable::builder("ExternalSales")
            .source(PivotSource::External {
                connection_name: "SalesConnection".to_string(),
                command_text: Some("select Region, Sales from Sales".to_string()),
            })
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .refresh_policy(refresh_policy)
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let cache_def = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PIVOT_CACHE_DEF)
                    .then_some(payload)
            })
            .expect("pivot cache definition payload");
        assert_ne!(
            cache_def[3] & 0x20,
            0,
            "external pivot cache background query should set fBackgroundQuery"
        );

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = &read.worksheet(0).unwrap().pivot_tables()[0];
        assert!(pivot.refresh_policy.background_query);
        assert!(matches!(
            pivot.cache_info().map(|info| info.source_kind),
            Some(duke_sheets_core::PivotCacheSourceKind::External)
        ));
    }

    #[test]
    fn xlsb_writer_rejects_background_query_for_non_external_pivot_cache() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Sales").unwrap();
        ws.set_cell_value("A2", "East").unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();

        let refresh_policy = PivotRefreshPolicy {
            refresh_on_open: false,
            preserve_formatting: true,
            background_query: true,
            missing_items_limit: None,
        };
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B2").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .refresh_policy(refresh_policy)
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let err = XlsbWriter::write(&wb, Cursor::new(Vec::new()))
            .expect_err("background query is only valid for external XLSB pivot caches");
        assert!(
            err.to_string().contains(
                "XLSB pivot background query is only valid for external pivot cache sources"
            ),
            "{err}"
        );
    }

    #[test]
    fn xlsb_writer_rejects_external_pivot_source_without_matching_connection() {
        let mut wb = Workbook::new();
        let pivot = PivotTable::builder("ExternalSales")
            .source(PivotSource::External {
                connection_name: "MissingConnection".to_string(),
                command_text: None,
            })
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let err = XlsbWriter::write(&wb, Cursor::new(Vec::new()))
            .expect_err("external pivot sources should require a matching workbook connection");
        assert!(
            err.to_string()
                .contains("XLSB external pivot source references unknown connection"),
            "{err}"
        );
    }

    #[test]
    fn xlsb_writer_round_trips_olap_pivot_source_authoring() {
        let mut wb = Workbook::new();
        add_olap_connection_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let content_types = read_zip_entry(&bytes, "[Content_Types].xml");
        assert!(content_types.contains("/xl/connections.bin"));
        assert!(
            !content_types.contains("/xl/pivotCache/pivotCacheRecords1.bin"),
            "OLAP pivot caches must not advertise cache records"
        );
        assert!(
            !zip_entry_exists(&bytes, "xl/pivotCache/pivotCacheRecords1.bin"),
            "OLAP pivot caches must not write cache records"
        );
        assert!(read_zip_entry(&bytes, "xl/_rels/workbook.bin.rels").contains("connections.bin"));
        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let cache_def_record_types = cache_def_records
            .iter()
            .map(|(record_type, _)| *record_type)
            .collect::<Vec<_>>();
        assert!(cache_def_record_types.contains(&crate::biff12::records::BRT_BEGIN_PCD_HIERARCHIES));
        assert!(cache_def_record_types.contains(&crate::biff12::records::BRT_BEGIN_PCD_HIERARCHY));
        assert!(
            cache_def_record_types.contains(&crate::biff12::records::BRT_BEGIN_PCDH_FIELDS_USAGE)
        );
        assert!(cache_def_record_types.contains(&crate::biff12::records::BRT_BEGIN_DIMS));
        assert!(cache_def_record_types.contains(&crate::biff12::records::BRT_BEGIN_DIM));
        let source = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_SOURCE).then_some(payload)
            })
            .expect("PCD source payload");
        assert_eq!(
            source,
            &[1, 0, 0, 0, 10, 0, 0, 0],
            "OLAP XLSB pivot sources should reference the OLAP workbook connection id"
        );

        let pivot_records = record_types(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        assert!(pivot_records.contains(&crate::biff12::records::BRT_BEGIN_SXTHS));
        assert!(pivot_records.contains(&crate::biff12::records::BRT_BEGIN_SXTH));
        assert!(pivot_records.contains(&crate::biff12::records::BRT_BEGIN_ISXTH_RWS));

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = &read.worksheet(0).unwrap().pivot_tables()[0];
        assert!(
            matches!(
                &pivot.source,
                PivotSource::Olap {
                    connection_name,
                    cube: Some(cube),
                    command_text: Some(command_text),
                } if connection_name == "CubeSales"
                    && cube == "SalesCube"
                    && command_text == "SalesCube"
            ),
            "unexpected pivot source: {:?}",
            pivot.source
        );
        assert_eq!(pivot.rows[0].field.name, "Region");
        assert_eq!(pivot.measures[0].field.name, "Sales");
        assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Sales"));
        assert!(matches!(
            pivot.cache_info().map(|info| info.source_kind),
            Some(duke_sheets_core::PivotCacheSourceKind::Olap)
        ));
    }

    #[test]
    fn xlsb_writer_rejects_olap_pivot_source_without_matching_connection() {
        let mut wb = Workbook::new();
        let pivot = PivotTable::builder("OlapSales")
            .source(PivotSource::Olap {
                connection_name: "MissingConnection".to_string(),
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

        let err = XlsbWriter::write(&wb, Cursor::new(Vec::new()))
            .expect_err("OLAP pivot sources should require a matching workbook connection");
        assert!(
            err.to_string()
                .contains("XLSB OLAP pivot source references unknown connection"),
            "{err}"
        );
    }

    #[test]
    fn xlsb_writer_rejects_olap_pivot_source_with_non_olap_connection() {
        let mut wb = Workbook::new();
        wb.add_data_connection(
            WorkbookConnection::database(10, "CubeSales", "Provider=MSDASQL;DSN=Sales;")
                .with_command("select Region, Sales from Sales")
                .with_command_type(2),
        )
        .unwrap();
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

        let err = XlsbWriter::write(&wb, Cursor::new(Vec::new()))
            .expect_err("OLAP pivot sources should require OLAP workbook connections");
        assert!(
            err.to_string()
                .contains("XLSB OLAP pivot source requires an OLAP connection"),
            "{err}"
        );
    }

    #[test]
    fn xlsb_writer_round_trips_unnamed_scenario_pivot_source_authoring() {
        let mut wb = Workbook::new();
        let pivot = PivotTable::builder("ScenarioSales")
            .source(PivotSource::Scenario {
                name: String::new(),
            })
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let source = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_SOURCE).then_some(payload)
            })
            .expect("PCD source payload");
        assert_eq!(
            source,
            &[3, 0, 0, 0, 0, 0, 0, 0],
            "scenario XLSB pivot sources should use the BIFF12 scenario source type"
        );

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = &read.worksheet(0).unwrap().pivot_tables()[0];
        assert!(matches!(
            &pivot.source,
            PivotSource::Scenario { name } if name.is_empty()
        ));
        assert_eq!(pivot.rows[0].field.name, "Region");
        assert_eq!(pivot.measures[0].field.name, "Sales");
        assert!(
            matches!(
                pivot.cache_info().map(|info| info.source_kind),
                Some(duke_sheets_core::PivotCacheSourceKind::Scenario)
            ),
            "scenario pivot cache source kind should survive round-trip"
        );
    }

    #[test]
    fn xlsb_writer_round_trips_pivot_cache_refresh_policy() {
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
            .refresh_policy(refresh_policy.clone())
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let cache_def = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PIVOT_CACHE_DEF)
                    .then_some(payload)
            })
            .expect("pivot cache definition payload");
        assert_ne!(
            cache_def[3] & 0x04,
            0,
            "refresh-on-open should set the BIFF12 cache refresh flag"
        );
        assert_eq!(
            i32::from_le_bytes(cache_def[4..8].try_into().unwrap()),
            5,
            "missing-items limit should be written to the BIFF12 ghost-items field"
        );

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("RefreshPolicyPivot")
            .expect("RefreshPolicyPivot after round-trip");
        assert_eq!(pivot.refresh_policy.refresh_on_open, true);
        assert_eq!(pivot.refresh_policy.missing_items_limit, Some(5));
    }

    #[test]
    fn xlsb_writer_round_trips_pivot_preserve_formatting_policy() {
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

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let sx_view = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVIEW).then_some(payload)
            })
            .expect("pivot view payload");
        assert_eq!(
            sx_view[4] & 0x80,
            0,
            "preserve-formatting=false should clear fPreserveFormatting"
        );

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("NoPreserveFormatting")
            .expect("NoPreserveFormatting after round-trip");
        assert!(!pivot.refresh_policy.preserve_formatting);
    }

    #[test]
    fn xlsb_writer_round_trips_pivot_table_layout_flags() {
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

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let sx_view = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVIEW).then_some(payload)
            })
            .expect("pivot view payload");

        assert_eq!(sx_view[1], 0x36);
        assert_eq!(
            u16::from_le_bytes(sx_view[2..4].try_into().unwrap()),
            0x83B1
        );
        assert_eq!(sx_view[13], 2, "pageWrap should be written");
        assert!(sxview_flag(sx_view, 2), "showEmptyRows");
        assert!(sxview_flag(sx_view, 3), "showEmptyColumns");
        assert!(!sxview_flag(sx_view, 4), "enableWizard");
        assert!(!sxview_flag(sx_view, 5), "enableDrill");
        assert!(!sxview_flag(sx_view, 6), "enableFieldProperties");
        assert!(sxview_flag(sx_view, 9), "showError");
        assert!(sxview_flag(sx_view, 10), "showMissing");
        assert!(sxview_flag(sx_view, 11), "pageOverThenDown");
        assert!(!sxview_flag(sx_view, 13), "rowGrandTotals");
        assert!(!sxview_flag(sx_view, 14), "colGrandTotals");
        assert!(sxview_flag(sx_view, 15), "fieldPrintTitles");
        assert!(sxview_flag(sx_view, 17), "itemPrintTitles");
        assert!(sxview_flag(sx_view, 18), "mergeItem");
        assert!(sxview_flag(sx_view, 20), "grandTotalCaption present");
        assert!(sxview_flag(sx_view, 33), "outline default");
        assert!(sxview_flag(sx_view, 34), "outline data");
        assert!(!sxview_flag(sx_view, 38), "errorCaption present");
        assert!(!sxview_flag(sx_view, 39), "missingCaption present");

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("LayoutFlags")
            .expect("LayoutFlags after round-trip");
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
        assert_eq!(pivot.layout.indent, 3);
        assert!(pivot.layout.show_empty_rows);
        assert!(pivot.layout.show_empty_columns);
    }

    #[test]
    fn xlsb_writer_rejects_named_scenario_pivot_source_authoring() {
        let mut wb = Workbook::new();
        let pivot = PivotTable::builder("ScenarioSales")
            .source(PivotSource::Scenario {
                name: "BestCase".to_string(),
            })
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Sales", PivotAggregate::Sum, "Total Sales")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let err = XlsbWriter::write(&wb, Cursor::new(Vec::new()))
            .expect_err("named scenario pivot sources should fail explicitly");
        assert!(
            err.to_string().contains(
                "XLSB named scenario pivot source authoring requires Scenario Manager records"
            ),
            "{err}"
        );
    }

    #[test]
    fn xlsb_writer_round_trips_workbook_data_connections() {
        let mut wb = Workbook::new();
        let mut region_param = WorkbookConnectionParameter::value(
            "RegionParam",
            WorkbookConnectionParameterValue::String("East".to_string()),
        );
        region_param.sql_type = 12;
        let database =
            WorkbookConnection::database(7, "SalesConnection", "Provider=MSDASQL;DSN=Sales;")
                .with_command("select Region, Revenue from Sales")
                .with_command_type(2)
                .with_source_file("connections/sales.dsn")
                .with_odc_file("connections/sales.odc")
                .with_description("Sales warehouse")
                .with_keep_alive(true)
                .with_interval(30)
                .with_reconnection_method(2)
                .with_refresh_on_load(true)
                .with_background(true)
                .with_save_data(true)
                .with_save_password(true)
                .with_credentials(WorkbookConnectionCredentials::Stored)
                .with_single_sign_on_id("sales-sso")
                .with_parameter(region_param);
        wb.add_data_connection(database).unwrap();

        let mut web = WorkbookConnection::web(8, "WebSales", "http://127.0.0.1/duke-sheets/sales");
        web.kind = WorkbookConnectionKind::Web {
            url: Some("http://127.0.0.1/duke-sheets/sales".to_string()),
            xml: false,
            source_data: true,
            html_tables: true,
            html_format: Some("all".to_string()),
            post: Some("region=all".to_string()),
            edit_page: Some("http://127.0.0.1/duke-sheets/edit".to_string()),
        };
        wb.add_data_connection(web.with_parameter(WorkbookConnectionParameter::prompt(
            "Region",
            "Choose region",
        )))
        .unwrap();

        let mut text = WorkbookConnection::text(9, "CsvSales", "/data/sales.csv");
        text.kind = WorkbookConnectionKind::Text {
            source_file: Some("/data/sales.csv".to_string()),
            delimiter: Some("|".to_string()),
            first_row: 2,
            delimited: true,
            decimal: Some(".".to_string()),
            thousands: Some(",".to_string()),
        };
        wb.add_data_connection(text).unwrap();

        let mut olap = WorkbookConnection::olap(10, "CubeSales").with_connection_type(5);
        olap.kind = WorkbookConnectionKind::Olap {
            connection: Some("Provider=MSOLAP;Data Source=olapserver;".to_string()),
            command: Some("SalesCube".to_string()),
            command_type: Some(1),
            local: false,
            local_connection: Some("CubeFile=cube.cub".to_string()),
            local_refresh: false,
            send_locale: true,
            row_drill_count: Some(1000),
        };
        wb.add_data_connection(olap).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        assert!(read_zip_entry(&bytes, "[Content_Types].xml").contains("/xl/connections.bin"));
        assert!(read_zip_entry(&bytes, "xl/_rels/workbook.bin.rels").contains("connections.bin"));

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        assert_eq!(read.data_connections().len(), 4);

        let database = read
            .data_connection_by_name("SalesConnection")
            .expect("database connection");
        assert_eq!(
            database.source_file.as_deref(),
            Some("connections/sales.dsn")
        );
        assert_eq!(database.odc_file.as_deref(), Some("connections/sales.odc"));
        assert_eq!(database.description.as_deref(), Some("Sales warehouse"));
        assert_eq!(
            database.credentials,
            Some(WorkbookConnectionCredentials::Stored)
        );
        assert!(database.save_password);
        assert_eq!(database.single_sign_on_id.as_deref(), Some("sales-sso"));
        assert_eq!(database.parameters.len(), 1);
        assert_eq!(database.parameters[0].name.as_deref(), Some("RegionParam"));
        assert_eq!(
            database.parameters[0].value,
            WorkbookConnectionParameterValue::String("East".to_string())
        );
        match &database.kind {
            WorkbookConnectionKind::Database {
                connection,
                command,
                command_type,
            } => {
                assert_eq!(connection, "Provider=MSDASQL;DSN=Sales;");
                assert_eq!(
                    command.as_deref(),
                    Some("select Region, Revenue from Sales")
                );
                assert_eq!(*command_type, Some(2));
            }
            other => panic!("unexpected database connection kind: {other:?}"),
        }

        match &read.data_connection_by_name("WebSales").unwrap().kind {
            WorkbookConnectionKind::Web {
                url,
                source_data,
                html_tables,
                html_format,
                post,
                edit_page,
                ..
            } => {
                assert_eq!(url.as_deref(), Some("http://127.0.0.1/duke-sheets/sales"));
                assert!(*source_data);
                assert!(*html_tables);
                assert_eq!(html_format.as_deref(), Some("all"));
                assert_eq!(post.as_deref(), Some("region=all"));
                assert_eq!(
                    edit_page.as_deref(),
                    Some("http://127.0.0.1/duke-sheets/edit")
                );
            }
            other => panic!("unexpected web connection kind: {other:?}"),
        }
        assert_eq!(
            read.data_connection_by_name("WebSales").unwrap().parameters[0]
                .prompt
                .as_deref(),
            Some("Choose region")
        );

        match &read.data_connection_by_name("CsvSales").unwrap().kind {
            WorkbookConnectionKind::Text {
                source_file,
                delimiter,
                first_row,
                decimal,
                thousands,
                ..
            } => {
                assert_eq!(source_file.as_deref(), Some("/data/sales.csv"));
                assert_eq!(delimiter.as_deref(), Some("|"));
                assert_eq!(*first_row, 2);
                assert_eq!(decimal.as_deref(), Some("."));
                assert_eq!(thousands.as_deref(), Some(","));
            }
            other => panic!("unexpected text connection kind: {other:?}"),
        }

        match &read.data_connection_by_name("CubeSales").unwrap().kind {
            WorkbookConnectionKind::Olap {
                connection,
                command,
                command_type,
                local_connection,
                local_refresh,
                send_locale,
                row_drill_count,
                ..
            } => {
                assert_eq!(
                    connection.as_deref(),
                    Some("Provider=MSOLAP;Data Source=olapserver;")
                );
                assert_eq!(command.as_deref(), Some("SalesCube"));
                assert_eq!(*command_type, Some(1));
                assert_eq!(local_connection.as_deref(), Some("CubeFile=cube.cub"));
                assert!(!*local_refresh);
                assert!(*send_locale);
                assert_eq!(*row_drill_count, Some(1000));
            }
            other => panic!("unexpected OLAP connection kind: {other:?}"),
        }
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_table_source_cache_records() {
        let mut wb = Workbook::new();
        add_table_source_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let source_payload = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_SOURCE).then_some(payload)
            })
            .expect("pivot cache source record");
        assert_eq!(&source_payload[0..4], &0u32.to_le_bytes());
        assert_eq!(&source_payload[4..8], &0u32.to_le_bytes());

        let range_payload = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDS_SHEET).then_some(payload)
            })
            .expect("pivot cache range source record");
        assert_eq!(range_payload[0] & 0x01, 0x01);
        assert_eq!(range_payload[1], 0x00);
        assert_eq!(range_payload[2], 0x00);
        let (source_name, consumed) = wide_string_at(range_payload, 3);
        assert_eq!(source_name, "SalesData");
        assert_eq!(range_payload.len(), 3 + consumed);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_table_source() {
        let mut wb = Workbook::new();
        add_table_source_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = wb2
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("SalesTablePivot")
            .unwrap();
        assert!(matches!(
            &pivot.source,
            PivotSource::Table { name } if name == "SalesData"
        ));
        assert_eq!(pivot.rows[0].field.name, "Region");
        assert_eq!(pivot.measures[0].field.name, "Sales");
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_consolidation_source_records() {
        let mut wb = Workbook::new();
        add_consolidation_source_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let record_types = cache_def_records
            .iter()
            .map(|(record_type, _)| *record_type)
            .collect::<Vec<_>>();

        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_PCDS_CONSOL));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_PCDSC_PAGES));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_PCDSC_SETS));

        let source_payload = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_SOURCE).then_some(payload)
            })
            .expect("pivot cache source record");
        assert_eq!(&source_payload[0..4], &2u32.to_le_bytes());
        assert_eq!(&source_payload[4..8], &0u32.to_le_bytes());

        let page_items = cache_def_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDSC_PITEM)
                    .then(|| wide_string_at(payload, 0).0)
            })
            .collect::<Vec<_>>();
        assert_eq!(page_items, vec!["Retail", "Wholesale"]);

        let sets = cache_def_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDSC_SET).then_some(payload)
            })
            .collect::<Vec<_>>();
        assert_eq!(sets.len(), 2);
        assert_eq!(&sets[0][0..4], &0u32.to_le_bytes());
        assert_eq!(sets[0][16], 0);
        assert_eq!(sets[0][18] & 0x02, 0x02);
        let (sheet_name, consumed) = wide_string_at(sets[0], 19);
        assert_eq!(sheet_name, "Sheet1");
        let rfx = 19 + consumed;
        assert_eq!(&sets[0][rfx..rfx + 4], &0u32.to_le_bytes());
        assert_eq!(&sets[0][rfx + 4..rfx + 8], &2u32.to_le_bytes());
        assert_eq!(&sets[0][rfx + 8..rfx + 12], &0u32.to_le_bytes());
        assert_eq!(&sets[0][rfx + 12..rfx + 16], &1u32.to_le_bytes());
        assert_eq!(&sets[1][0..4], &1u32.to_le_bytes());
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_consolidation_source() {
        let mut wb = Workbook::new();
        add_consolidation_source_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = wb2
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("ConsolidatedSales")
            .unwrap();
        let PivotSource::Consolidation { ranges } = &pivot.source else {
            panic!("expected consolidation source");
        };
        assert_eq!(ranges.len(), 2);
        assert_eq!(ranges[0].sheet.as_deref(), Some("Sheet1"));
        assert_eq!(ranges[0].range, Some(CellRange::parse("A1:B3").unwrap()));
        assert_eq!(ranges[0].page_items, vec!["Retail".to_string()]);
        assert_eq!(ranges[1].sheet.as_deref(), Some("WestData"));
        assert_eq!(ranges[1].range, Some(CellRange::parse("A1:B2").unwrap()));
        assert_eq!(ranges[1].page_items, vec!["Wholesale".to_string()]);
        assert_eq!(pivot.rows[0].field.name, "Row");
        assert_eq!(pivot.measures[0].field.name, "Value");
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_named_consolidation_source() {
        let mut wb = Workbook::new();
        add_named_consolidation_source_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let source_payload = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_SOURCE).then_some(payload)
            })
            .expect("pivot cache source record");
        assert_eq!(&source_payload[0..4], &2u32.to_le_bytes());
        let set = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDSC_SET).then_some(payload)
            })
            .expect("consolidation source set");
        assert_eq!(set[16], 1, "fName should mark a named source");
        let (name, _) = wide_string_at(set, 19);
        assert_eq!(name, "NamedSalesSource");

        let read = XlsbReader::read(Cursor::new(bytes)).unwrap();
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("NamedConsolidatedSales")
            .unwrap();
        let PivotSource::Consolidation { ranges } = &pivot.source else {
            panic!("expected consolidation source");
        };
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0].name.as_deref(), Some("NamedSalesSource"));
        assert_eq!(ranges[0].range, None);
        assert_eq!(pivot.rows[0].field.name, "Region");
        assert_eq!(pivot.measures[0].field.name, "Sales");
        assert_eq!(
            pivot.cache_info().and_then(|info| info.record_count),
            Some(0)
        );
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_external_consolidation_source() {
        let mut wb = Workbook::new();
        add_external_consolidation_source_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_rels =
            read_zip_entry(&bytes, "xl/pivotCache/_rels/pivotCacheDefinition1.bin.rels");
        assert!(cache_rels.contains("externalLinkPath"));
        assert!(cache_rels.contains("Id=\"rIdExternalSource\""));
        assert!(cache_rels.contains("Target=\"file:///C:/data/source.xlsx\""));
        assert!(cache_rels.contains("TargetMode=\"External\""));

        let cache_def_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let set = cache_def_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDSC_SET).then_some(payload)
            })
            .expect("external consolidation source set");
        assert_eq!(set[16], 0, "external range source should still use rfx");
        assert_eq!(set[18] & 0x03, 0x03, "fLoadRelId + fLoadSheet");
        let (sheet, consumed) = wide_string_at(set, 19);
        assert_eq!(sheet, "ExternalData");
        let (relationship_id, consumed_rel) = wide_string_at(set, 19 + consumed);
        assert_eq!(relationship_id, "rIdExternalSource");
        let rfx = 19 + consumed + consumed_rel;
        assert_eq!(&set[rfx..rfx + 4], &0u32.to_le_bytes());
        assert_eq!(&set[rfx + 4..rfx + 8], &2u32.to_le_bytes());

        let read = XlsbReader::read(Cursor::new(bytes)).unwrap();
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("ExternalConsolidatedSales")
            .unwrap();
        let PivotSource::Consolidation { ranges } = &pivot.source else {
            panic!("expected consolidation source");
        };
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0].sheet.as_deref(), Some("ExternalData"));
        assert_eq!(ranges[0].range, Some(CellRange::parse("A1:B3").unwrap()));
        assert_eq!(
            ranges[0].external_relationship_id.as_deref(),
            Some("rIdExternalSource")
        );
        assert_eq!(
            ranges[0].external_relationship_target.as_deref(),
            Some("file:///C:/data/source.xlsx")
        );
        assert_eq!(pivot.rows[0].field.name, "Region");
        assert_eq!(pivot.measures[0].field.name, "Sales");
        assert_eq!(
            pivot.cache_info().and_then(|info| info.record_count),
            Some(0)
        );
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_column_axis_records() {
        let mut wb = Workbook::new();
        add_column_test_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let record_types = pivot_records
            .iter()
            .map(|(record_type, _)| *record_type)
            .collect::<Vec<_>>();

        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_ISXVD_RWS));
        assert!(record_types.contains(&crate::biff12::records::BRT_END_ISXVD_RWS));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_ISXVD_COLS));
        assert!(record_types.contains(&crate::biff12::records::BRT_END_ISXVD_COLS));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_SX_COL_ITEMS));

        let col_fields_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_ISXVD_COLS).then_some(payload)
            })
            .expect("column axis field declaration");
        assert_eq!(
            col_fields_payload,
            &vec![1, 0, 0, 0, 1, 0, 0, 0],
            "one column field, cache field index 1 (Quarter)"
        );

        let sxvd_axes = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD).then(|| payload[0])
            })
            .collect::<Vec<_>>();
        assert_eq!(sxvd_axes, vec![0x01, 0x02, 0x08]);
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_page_axis_records() {
        let mut wb = Workbook::new();
        add_page_test_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let record_types = pivot_records
            .iter()
            .map(|(record_type, _)| *record_type)
            .collect::<Vec<_>>();

        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_SXPIS));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_SXPI));
        assert!(record_types.contains(&crate::biff12::records::BRT_END_SXPI));
        assert!(record_types.contains(&crate::biff12::records::BRT_END_SXPIS));

        let page_fields_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPIS).then_some(payload)
            })
            .expect("page axis field collection");
        assert_eq!(page_fields_payload, &vec![1, 0, 0, 0]);

        let page_field_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
            })
            .expect("page axis field declaration");
        assert_eq!(
            page_field_payload,
            &vec![
                1, 0, 0, 0, // field index 1 (Salesperson)
                0, 0, 0, 0, // selected item index 0 (Ada)
                0xFF, 0xFF, 0xFF, 0xFF, // non-OLAP hierarchy sentinel
                0,
            ]
        );

        let sxvd_axes = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD).then(|| payload[0])
            })
            .collect::<Vec<_>>();
        assert_eq!(sxvd_axes, vec![0x01, 0x04, 0x08]);
    }

    #[test]
    fn xlsb_writer_round_trips_multi_item_page_filters() {
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
            .page("Region")
            .filter(PivotFilter::field_items("Region", ["East", "West"]))
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let sxvd_axes = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD).then(|| payload[0])
            })
            .collect::<Vec<_>>();
        assert_eq!(
            sxvd_axes,
            vec![0x04, 0x01, 0x08],
            "multi-item page filters should keep the source page, row, and data fields in cache order"
        );
        let page_field_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
            })
            .expect("page axis field declaration");
        assert_eq!(
            &page_field_payload[4..8],
            &0x0010_00FEu32.to_le_bytes(),
            "multi-select page filters should use Excel's multiple-items sentinel"
        );

        let region_items = sxvi_payloads_for_sxvd(&pivot_records, 0);
        assert_eq!(
            region_items
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![0, 0, 1, 0],
            "Central should be hidden while East, West, and Auto remain visible"
        );
        assert_eq!(
            i32::from_le_bytes(region_items[2][3..7].try_into().unwrap()),
            2
        );
        assert_eq!(
            sxli_item_values_for_axis(
                &pivot_records,
                crate::biff12::records::BRT_BEGIN_SX_ROW_ITEMS,
                crate::biff12::records::BRT_END_SX_ROW_ITEMS,
            ),
            vec![vec![0], vec![1]],
            "row SXLI tuples should omit hidden Central"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("RevenueByRep")
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::field_items("Region", ["East", "West"])]
        );
    }

    #[test]
    fn xlsb_writer_rejects_unknown_page_filter_item() {
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
            .page("Salesperson")
            .filter(PivotFilter::field_items("Salesperson", ["Ben"]))
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let mut buf = Vec::new();
        let err = XlsbWriter::write(&wb, Cursor::new(&mut buf))
            .expect_err("unknown page item should fail for XLSB");
        assert!(
            err.to_string().contains("selected item is not present"),
            "{err}"
        );
    }

    #[test]
    fn xlsb_writer_round_trips_row_field_item_filters() {
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
            .row("Region")
            .filter(PivotFilter::field_items("Region", ["East", "West"]))
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let region_items = sxvi_payloads_for_sxvd(&pivot_records, 0);
        assert_eq!(
            region_items
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![0, 0, 1, 0],
            "Central should be hidden while East, West, and Auto remain visible"
        );
        assert_eq!(
            i32::from_le_bytes(region_items[2][3..7].try_into().unwrap()),
            2
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("FilteredRows")
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::field_items("Region", ["East", "West"])]
        );
    }

    #[test]
    fn xlsb_writer_round_trips_column_field_item_filters() {
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

        let pivot = PivotTable::builder("FilteredColumns")
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .column("Region")
            .filter(PivotFilter::field_items("Region", ["East", "West"]))
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let region_items = sxvi_payloads_for_sxvd(&pivot_records, 0);
        assert_eq!(
            region_items
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![0, 0, 1, 0],
            "Central should be hidden while East, West, and Auto remain visible"
        );
        assert_eq!(
            sxli_item_values_for_axis(
                &pivot_records,
                crate::biff12::records::BRT_BEGIN_SX_COL_ITEMS,
                crate::biff12::records::BRT_END_SX_COL_ITEMS,
            ),
            vec![vec![0], vec![1]],
            "column SXLI tuples should omit hidden Central"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("FilteredColumns")
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::field_items("Region", ["East", "West"])]
        );
    }

    #[test]
    fn xlsb_writer_round_trips_source_only_field_item_filters() {
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

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let sxvd_axes = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD).then(|| payload[0])
            })
            .collect::<Vec<_>>();
        assert_eq!(sxvd_axes, vec![0x01, 0x00, 0x08]);

        let channel_items = sxvi_payloads_for_sxvd(&pivot_records, 1);
        assert_eq!(
            channel_items
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![0, 1, 0],
            "Store should be hidden while Online and Auto remain visible"
        );
        assert_eq!(
            sxli_item_values_for_axis(
                &pivot_records,
                crate::biff12::records::BRT_BEGIN_SX_ROW_ITEMS,
                crate::biff12::records::BRT_END_SX_ROW_ITEMS,
            ),
            vec![vec![0], vec![2]],
            "row SXLI tuples should omit rows hidden by the source-only Channel filter"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
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
    fn xlsb_writer_rejects_empty_label_filter_values() {
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

        let mut buf = Vec::new();
        let err = XlsbWriter::write(&wb, Cursor::new(&mut buf))
            .expect_err("empty label filter values should fail for XLSB");
        assert!(err.to_string().contains("empty label filter"), "{err}");
    }

    #[test]
    fn xlsb_writer_round_trips_label_contains_pivot_filter() {
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

        let pivot = PivotTable::builder("LabelFilteredRows")
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .filter(PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator: PivotFilterOperator::Contains,
                value: "Ea".to_string(),
            })
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let row_field = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD
                    && payload.first() == Some(&0x01))
                .then_some(payload)
            })
            .expect("row-field SXVD");
        let behavior_flags = u32::from_le_bytes(row_field[8..12].try_into().unwrap());
        assert_ne!(
            behavior_flags & (1 << 17),
            0,
            "label filters should set fHasAdvFilter on the filtered field"
        );

        let filter_header = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
            })
            .expect("BrtBeginSXFilter");
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
        let (label_value, _) = crate::biff12::parser::wide_str(filter_header, 30).unwrap();
        assert_eq!(label_value, "Ea");

        let custom_filter = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
            })
            .expect("BrtCustomFilter");
        assert_eq!(custom_filter[0], 6, "custom filter value type");
        assert_eq!(custom_filter[1], 2, "contains uses equality over wildcards");
        let (custom_value, _) = crate::biff12::parser::wide_str(custom_filter, 10).unwrap();
        assert_eq!(custom_value, "*Ea*");

        let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("LabelFilteredRows")
            .expect("LabelFilteredRows round-trip");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator: PivotFilterOperator::Contains,
                value: "Ea".to_string(),
            }]
        );
    }

    #[test]
    fn xlsb_writer_round_trips_label_equals_pivot_filter() {
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
            .filter(PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator: PivotFilterOperator::Equals,
                value: "East".to_string(),
            })
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let row_field = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD
                    && payload.first() == Some(&0x01))
                .then_some(payload)
            })
            .expect("row-field SXVD");
        let behavior_flags = u32::from_le_bytes(row_field[8..12].try_into().unwrap());
        assert_ne!(
            behavior_flags & (1 << 17),
            0,
            "label filters should set fHasAdvFilter on the filtered field"
        );

        let filter_header = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
            })
            .expect("BrtBeginSXFilter");
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
        let (label_value, _) = crate::biff12::parser::wide_str(filter_header, 30).unwrap();
        assert_eq!(label_value, "East");
        assert!(
            !pivot_records
                .iter()
                .any(|(record_type, _)| *record_type == crate::biff12::records::BRT_CUSTOM_FILTER),
            "caption-equals filters should not emit a custom wildcard filter"
        );

        let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("LabelEqualsRows")
            .expect("LabelEqualsRows round-trip");
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
    fn xlsb_writer_round_trips_label_prefix_suffix_pivot_filters() {
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
                .filter(PivotFilter::Label {
                    field: PivotFieldRef::new("Region"),
                    operator,
                    value: value.to_string(),
                })
                .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
                .build()
                .unwrap();
            wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

            let bytes = write_xlsb_bytes(&wb);
            let pivot_records = records_with_payload(read_zip_entry_bytes(
                &bytes,
                "xl/pivotTables/pivotTable1.bin",
            ));
            let row_field = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SXVD
                        && payload.first() == Some(&0x01))
                    .then_some(payload)
                })
                .expect("row-field SXVD");
            let behavior_flags = u32::from_le_bytes(row_field[8..12].try_into().unwrap());
            assert_ne!(
                behavior_flags & (1 << 17),
                0,
                "label filters should set fHasAdvFilter on the filtered field"
            );

            let filter_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
                })
                .expect("BrtBeginSXFilter");
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
            let (label_value, _) = crate::biff12::parser::wide_str(filter_header, 30).unwrap();
            assert_eq!(label_value, value);

            let custom_filter = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
                })
                .expect("BrtCustomFilter");
            assert_eq!(custom_filter[0], 6, "custom filter value type");
            assert_eq!(custom_filter[1], 2, "wildcard label filters use equality");
            let (stored_value, _) = crate::biff12::parser::wide_str(custom_filter, 10).unwrap();
            assert_eq!(stored_value, custom_value);

            let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
            let pivot = result
                .worksheet(0)
                .unwrap()
                .pivot_table_by_name(pivot_name)
                .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_negative_label_pivot_filters() {
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
                .filter(PivotFilter::Label {
                    field: PivotFieldRef::new("Region"),
                    operator,
                    value: value.to_string(),
                })
                .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
                .build()
                .unwrap();
            wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

            let bytes = write_xlsb_bytes(&wb);
            let pivot_records = records_with_payload(read_zip_entry_bytes(
                &bytes,
                "xl/pivotTables/pivotTable1.bin",
            ));
            let filter_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
                })
                .expect("BrtBeginSXFilter");
            assert_eq!(
                u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
                0
            );
            assert_eq!(
                u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
                filter_type
            );
            let (label_value, _) = crate::biff12::parser::wide_str(filter_header, 30).unwrap();
            assert_eq!(label_value, value);

            let custom_filter = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
                })
                .expect("BrtCustomFilter");
            assert_eq!(custom_filter[0], 6, "custom filter value type");
            assert_eq!(custom_filter[1], 5, "negative label filters use notEqual");
            assert_eq!(custom_filter[2], discriminator);
            let (stored_value, _) = crate::biff12::parser::wide_str(custom_filter, 10).unwrap();
            assert_eq!(stored_value, custom_value);

            let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
            let pivot = result
                .worksheet(0)
                .unwrap()
                .pivot_table_by_name(pivot_name)
                .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_label_comparison_pivot_filters() {
        for (operator, filter_type, custom_operator, pivot_name) in [
            (
                PivotFilterOperator::GreaterThan,
                12u32,
                4u8,
                "LabelGreaterRows",
            ),
            (
                PivotFilterOperator::GreaterThanOrEqual,
                13u32,
                6u8,
                "LabelGreaterEqualRows",
            ),
            (PivotFilterOperator::LessThan, 14u32, 1u8, "LabelLessRows"),
            (
                PivotFilterOperator::LessThanOrEqual,
                15u32,
                3u8,
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
                .filter(PivotFilter::Label {
                    field: PivotFieldRef::new("Region"),
                    operator,
                    value: "M".to_string(),
                })
                .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
                .build()
                .unwrap();
            wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

            let bytes = write_xlsb_bytes(&wb);
            let pivot_records = records_with_payload(read_zip_entry_bytes(
                &bytes,
                "xl/pivotTables/pivotTable1.bin",
            ));
            let filter_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
                })
                .expect("BrtBeginSXFilter");
            assert_eq!(
                u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
                0
            );
            assert_eq!(
                u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
                filter_type
            );
            let (label_value, _) = crate::biff12::parser::wide_str(filter_header, 30).unwrap();
            assert_eq!(label_value, "M");

            let custom_filter = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
                })
                .expect("BrtCustomFilter");
            assert_eq!(custom_filter[0], 6, "custom filter value type");
            assert_eq!(custom_filter[1], custom_operator);
            assert_eq!(custom_filter[2], 1, "comparison filters use raw criteria");
            let (stored_value, _) = crate::biff12::parser::wide_str(custom_filter, 10).unwrap();
            assert_eq!(stored_value, "M");

            let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
            let pivot = result
                .worksheet(0)
                .unwrap()
                .pivot_table_by_name(pivot_name)
                .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_label_between_pivot_filters() {
        for (not_between, filter_type, and_flag, first_operator, second_operator, pivot_name) in [
            (false, 16u32, 1i32, 6u8, 3u8, "LabelBetweenRows"),
            (true, 17u32, 0i32, 1u8, 4u8, "LabelNotBetweenRows"),
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
                .filter(PivotFilter::LabelBetween {
                    field: PivotFieldRef::new("Region"),
                    start: "East".to_string(),
                    end: "West".to_string(),
                    not_between,
                })
                .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
                .build()
                .unwrap();
            wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

            let bytes = write_xlsb_bytes(&wb);
            let pivot_records = records_with_payload(read_zip_entry_bytes(
                &bytes,
                "xl/pivotTables/pivotTable1.bin",
            ));
            let filter_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
                })
                .expect("BrtBeginSXFilter");
            assert_eq!(
                u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
                0
            );
            assert_eq!(
                u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
                filter_type
            );
            let (label_value, _) = crate::biff12::parser::wide_str(filter_header, 30).unwrap();
            assert_eq!(label_value, "East");

            let custom_filters_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_CUSTOM_FILTERS)
                        .then_some(payload)
                })
                .expect("BrtBeginCustomFilters");
            assert_eq!(
                i32::from_le_bytes(custom_filters_header[0..4].try_into().unwrap()),
                and_flag
            );

            let custom_filters = pivot_records
                .iter()
                .filter_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
                })
                .collect::<Vec<_>>();
            assert_eq!(custom_filters.len(), 2);
            assert_eq!(custom_filters[0][0], 6, "label filter value type");
            assert_eq!(custom_filters[0][1], first_operator);
            assert_eq!(custom_filters[0][2], 1);
            assert_eq!(custom_filters[1][0], 6, "label filter value type");
            assert_eq!(custom_filters[1][1], second_operator);
            assert_eq!(custom_filters[1][2], 1);
            assert_eq!(
                crate::biff12::parser::wide_str(custom_filters[0], 10)
                    .unwrap()
                    .0,
                "East"
            );
            assert_eq!(
                crate::biff12::parser::wide_str(custom_filters[1], 10)
                    .unwrap()
                    .0,
                "West"
            );

            let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
            let pivot = result
                .worksheet(0)
                .unwrap()
                .pivot_table_by_name(pivot_name)
                .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_value_comparison_pivot_filters() {
        for (operator, filter_type, custom_operator, threshold, pivot_name) in [
            (
                PivotFilterOperator::Equals,
                18u32,
                2u8,
                20.0,
                "ValueEqualsRows",
            ),
            (
                PivotFilterOperator::NotEquals,
                19u32,
                5u8,
                20.0,
                "ValueNotEqualsRows",
            ),
            (
                PivotFilterOperator::GreaterThan,
                20u32,
                4u8,
                20.0,
                "ValueGreaterRows",
            ),
            (
                PivotFilterOperator::GreaterThanOrEqual,
                21u32,
                6u8,
                20.0,
                "ValueGreaterEqualRows",
            ),
            (
                PivotFilterOperator::LessThan,
                22u32,
                1u8,
                30.0,
                "ValueLessRows",
            ),
            (
                PivotFilterOperator::LessThanOrEqual,
                23u32,
                3u8,
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

            let bytes = write_xlsb_bytes(&wb);
            let pivot_records = records_with_payload(read_zip_entry_bytes(
                &bytes,
                "xl/pivotTables/pivotTable1.bin",
            ));
            let filter_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
                })
                .expect("BrtBeginSXFilter");
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
                "value filter should target the first data field"
            );

            let custom_filters_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_CUSTOM_FILTERS)
                        .then_some(payload)
                })
                .expect("BrtBeginCustomFilters");
            assert_eq!(
                i32::from_le_bytes(custom_filters_header[0..4].try_into().unwrap()),
                0
            );

            let custom_filter = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
                })
                .expect("BrtCustomFilter");
            assert_eq!(custom_filter[0], 4, "numeric filter value type");
            assert_eq!(custom_filter[1], custom_operator);
            assert_eq!(
                f64::from_le_bytes(custom_filter[2..10].try_into().unwrap()),
                threshold
            );

            let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
            let pivot = result
                .worksheet(0)
                .unwrap()
                .pivot_table_by_name(pivot_name)
                .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_value_between_pivot_filters() {
        for (not_between, filter_type, and_flag, first_operator, second_operator, pivot_name) in [
            (false, 24u32, 1i32, 6u8, 3u8, "ValueBetweenRows"),
            (true, 25u32, 0i32, 1u8, 4u8, "ValueNotBetweenRows"),
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

            let bytes = write_xlsb_bytes(&wb);
            let pivot_records = records_with_payload(read_zip_entry_bytes(
                &bytes,
                "xl/pivotTables/pivotTable1.bin",
            ));
            let filter_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
                })
                .expect("BrtBeginSXFilter");
            assert_eq!(
                u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
                filter_type
            );
            assert_eq!(
                u32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
                0,
                "value range filter should target the first data field"
            );

            let custom_filters_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_CUSTOM_FILTERS)
                        .then_some(payload)
                })
                .expect("BrtBeginCustomFilters");
            assert_eq!(
                i32::from_le_bytes(custom_filters_header[0..4].try_into().unwrap()),
                and_flag
            );

            let custom_filters = pivot_records
                .iter()
                .filter_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
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

            let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
            let pivot = result
                .worksheet(0)
                .unwrap()
                .pivot_table_by_name(pivot_name)
                .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_date_comparison_pivot_filters() {
        for (operator, filter_type, custom_operator, pivot_name) in [
            (PivotFilterOperator::Equals, 26u32, 2u8, "DateEqualsRows"),
            (
                PivotFilterOperator::NotEquals,
                62u32,
                5u8,
                "DateNotEqualsRows",
            ),
            (PivotFilterOperator::LessThan, 27u32, 1u8, "DateLessRows"),
            (
                PivotFilterOperator::LessThanOrEqual,
                63u32,
                3u8,
                "DateLessEqualRows",
            ),
            (
                PivotFilterOperator::GreaterThan,
                28u32,
                4u8,
                "DateGreaterRows",
            ),
            (
                PivotFilterOperator::GreaterThanOrEqual,
                64u32,
                6u8,
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

            let bytes = write_xlsb_bytes(&wb);
            let pivot_records = records_with_payload(read_zip_entry_bytes(
                &bytes,
                "xl/pivotTables/pivotTable1.bin",
            ));
            let filter_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
                })
                .expect("BrtBeginSXFilter");
            assert_eq!(
                u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
                filter_type
            );
            assert_eq!(
                i32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
                -1,
                "date filters should not target a data field"
            );

            let custom_filter = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
                })
                .expect("BrtCustomFilter");
            assert_eq!(custom_filter[0], 4);
            assert_eq!(custom_filter[1], custom_operator);
            assert_eq!(
                f64::from_le_bytes(custom_filter[2..10].try_into().unwrap()),
                44958.0
            );

            let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
            let pivot = result
                .worksheet(0)
                .unwrap()
                .pivot_table_by_name(pivot_name)
                .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_date_between_pivot_filters() {
        for (not_between, filter_type, and_flag, first_operator, second_operator, pivot_name) in [
            (false, 29u32, 1i32, 6u8, 3u8, "DateBetweenRows"),
            (true, 65u32, 0i32, 1u8, 4u8, "DateNotBetweenRows"),
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

            let bytes = write_xlsb_bytes(&wb);
            let pivot_records = records_with_payload(read_zip_entry_bytes(
                &bytes,
                "xl/pivotTables/pivotTable1.bin",
            ));
            let filter_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
                })
                .expect("BrtBeginSXFilter");
            assert_eq!(
                u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
                filter_type
            );

            let custom_filters_header = pivot_records
                .iter()
                .find_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_BEGIN_CUSTOM_FILTERS)
                        .then_some(payload)
                })
                .expect("BrtBeginCustomFilters");
            assert_eq!(
                i32::from_le_bytes(custom_filters_header[0..4].try_into().unwrap()),
                and_flag
            );

            let custom_filters = pivot_records
                .iter()
                .filter_map(|(record_type, payload)| {
                    (*record_type == crate::biff12::records::BRT_CUSTOM_FILTER).then_some(payload)
                })
                .collect::<Vec<_>>();
            assert_eq!(custom_filters.len(), 2);
            assert_eq!(custom_filters[0][1], first_operator);
            assert_eq!(
                f64::from_le_bytes(custom_filters[0][2..10].try_into().unwrap()),
                44927.0
            );
            assert_eq!(custom_filters[1][1], second_operator);
            assert_eq!(
                f64::from_le_bytes(custom_filters[1][2..10].try_into().unwrap()),
                45016.0
            );

            let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
            let pivot = result
                .worksheet(0)
                .unwrap()
                .pivot_table_by_name(pivot_name)
                .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_date_period_pivot_filter() {
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

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
            })
            .expect("BrtBeginSXFilter");
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            37,
            "ThisMonth should use the dynamic-date pivot filter type"
        );
        let dynamic_filter = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_DYNAMIC_FILTER).then_some(payload)
            })
            .expect("BrtDynamicFilter");
        assert_eq!(
            u32::from_le_bytes(dynamic_filter[0..4].try_into().unwrap()),
            0x0F
        );

        let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("DatePeriodRows")
            .expect("round-tripped pivot");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::DatePeriod {
                field: PivotFieldRef::new("Order Date"),
                period: PivotDatePeriod::ThisMonth,
            }]
        );
    }

    #[test]
    fn xlsb_writer_round_trips_column_value_pivot_filter() {
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

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
            })
            .expect("BrtBeginSXFilter");
        assert_eq!(
            u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
            0,
            "Region should be the first source field even when it is on the column axis"
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            20,
            "greater-than value filters should use Excel's value-greater filter type"
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
            0,
            "single-measure column value filter should target the first data field"
        );

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("ColumnValueFilter")
            .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_second_measure_value_pivot_filter() {
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

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let filter_header = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
            })
            .expect("BrtBeginSXFilter");
        assert_eq!(
            u32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
            1,
            "value filter should target the second data field"
        );

        let read = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("SecondMeasureValueFilter")
            .expect("round-tripped pivot");
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
    fn xlsb_writer_round_trips_top_n_pivot_filter() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        for (row, (region, revenue)) in [
            ("East", 10.0),
            ("West", 20.0),
            ("North", 30.0),
            ("South", 40.0),
        ]
        .iter()
        .enumerate()
        {
            let row = row + 2;
            ws.set_cell_value(&format!("A{row}"), *region).unwrap();
            ws.set_cell_value(&format!("B{row}"), *revenue).unwrap();
        }

        let filter_measure =
            PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue");
        let pivot = PivotTable::builder("TopRegions")
            .source_range(CellRange::parse("A1:B5").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .filter(PivotFilter::TopN {
                field: PivotFieldRef::new("Region"),
                measure: filter_measure.clone(),
                n: 50,
                top: false,
                percent: true,
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let row_field = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD
                    && payload.first() == Some(&0x01))
                .then_some(payload)
            })
            .expect("row-field SXVD");
        let behavior_flags = u32::from_le_bytes(row_field[8..12].try_into().unwrap());
        assert_ne!(
            behavior_flags & (1 << 17),
            0,
            "top-N filters should set fHasAdvFilter on the filtered field"
        );

        let filter_collection = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTERS).then_some(payload)
            })
            .expect("BrtBeginSXFilters");
        assert_eq!(filter_collection, &1u32.to_le_bytes());
        let filter_header = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SX_FILTER).then_some(payload)
            })
            .expect("BrtBeginSXFilter");
        assert_eq!(
            u32::from_le_bytes(filter_header[0..4].try_into().unwrap()),
            0
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[8..12].try_into().unwrap()),
            2
        );
        assert_eq!(
            u32::from_le_bytes(filter_header[20..24].try_into().unwrap()),
            0
        );
        let top10 = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_TOP10_FILTER).then_some(payload)
            })
            .expect("BrtTop10Filter");
        assert_eq!(top10[0], 0x06, "bottom percent + applied flags");
        assert_eq!(f64::from_le_bytes(top10[1..9].try_into().unwrap()), 50.0);

        let result = XlsbReader::read(Cursor::new(&bytes)).unwrap();
        let pivot = result
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("TopRegions")
            .expect("TopRegions round-trip");
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
                assert_eq!(measure, &filter_measure);
                assert_eq!(*n, 50);
                assert!(!*top);
                assert!(*percent);
            }
            other => panic!("unexpected filter: {other:?}"),
        }
    }

    #[test]
    fn xlsb_writer_round_trips_grouped_row_field_item_filters() {
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
            .row("Region")
            .filter(PivotFilter::field_items("Region", ["Coastal"]))
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .grouping(PivotGrouping::Manual {
                field: PivotFieldRef::new("Region"),
                groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let sxvd_axes = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD).then(|| payload[0])
            })
            .collect::<Vec<_>>();
        assert_eq!(
            sxvd_axes,
            vec![0x01, 0x08, 0x01],
            "manual grouped row pivots should keep both the source and derived group fields on the row axis"
        );
        let derived_group_items = sxvi_payloads_for_sxvd(&pivot_records, 2);
        assert_eq!(
            derived_group_items
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![1, 0, 0],
            "Central should be hidden while Coastal and Auto remain visible"
        );
        assert_eq!(
            sxli_item_values_for_axis(
                &pivot_records,
                crate::biff12::records::BRT_BEGIN_SX_ROW_ITEMS,
                crate::biff12::records::BRT_END_SX_ROW_ITEMS,
            ),
            vec![vec![1, 0], vec![1, 1]],
            "row SXLI tuples should only include East and West under Coastal"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
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
    fn xlsb_writer_round_trips_grouped_column_field_item_filters() {
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
            .column("Region")
            .filter(PivotFilter::field_items("Region", ["Coastal"]))
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .grouping(PivotGrouping::Manual {
                field: PivotFieldRef::new("Region"),
                groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
            })
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let sxvd_axes = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD).then(|| payload[0])
            })
            .collect::<Vec<_>>();
        assert_eq!(
            sxvd_axes,
            vec![0x02, 0x01, 0x08, 0x02],
            "manual grouped column pivots should keep both the source and derived group fields on the column axis"
        );
        let derived_group_items = sxvi_payloads_for_sxvd(&pivot_records, 3);
        assert_eq!(
            derived_group_items
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![1, 0, 0],
            "Central should be hidden while Coastal and Auto remain visible"
        );
        assert_eq!(
            sxli_item_values_for_axis(
                &pivot_records,
                crate::biff12::records::BRT_BEGIN_SX_COL_ITEMS,
                crate::biff12::records::BRT_END_SX_COL_ITEMS,
            ),
            vec![vec![1, 0], vec![1, 1]],
            "column SXLI tuples should only include East and West under Coastal"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
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

    fn manual_page_grouped_pivot_workbook(allowed_items: &[&str]) -> Workbook {
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

        let mut builder = PivotTable::builder("ManualPageRegions")
            .source_range(CellRange::parse("A1:C7").unwrap())
            .target_address("E2")
            .unwrap()
            .row("Salesperson")
            .page("Region")
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .grouping(PivotGrouping::Manual {
                field: PivotFieldRef::new("Region"),
                groups: vec![
                    PivotManualGroup::new("Coastal", ["East", "West", "South"]),
                    PivotManualGroup::new("Interior", ["Central", "North"]),
                ],
            });
        if !allowed_items.is_empty() {
            builder = builder.filter(PivotFilter::field_items(
                "Region",
                allowed_items.iter().copied(),
            ));
        }
        let pivot = builder.build().unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
        wb
    }

    #[test]
    fn xlsb_writer_round_trips_manual_grouped_page_item_filter() {
        let wb = manual_page_grouped_pivot_workbook(&["Coastal"]);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let page_field_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
            })
            .expect("page axis field declaration");
        assert_eq!(
            u32::from_le_bytes(page_field_payload[0..4].try_into().unwrap()),
            3,
            "manual grouped page field should target the derived group field"
        );
        assert_eq!(
            u32::from_le_bytes(page_field_payload[4..8].try_into().unwrap()),
            1,
            "Coastal should be selected from the derived group item list"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("ManualPageRegions")
            .expect("pivot after read");
        assert_eq!(
            pivot.filters,
            vec![PivotFilter::field_items("Region", ["Coastal"])]
        );
    }

    #[test]
    fn xlsb_writer_round_trips_manual_grouped_page_multi_item_filter() {
        let wb = manual_page_grouped_pivot_workbook(&["Coastal", "Interior"]);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let page_field_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
            })
            .expect("page axis field declaration");
        assert_eq!(
            u32::from_le_bytes(page_field_payload[0..4].try_into().unwrap()),
            3,
            "manual grouped page field should target the derived group field"
        );
        assert_eq!(
            u32::from_le_bytes(page_field_payload[4..8].try_into().unwrap()),
            0x0010_00FE,
            "multi-select page filters should use Excel's multiple-items sentinel"
        );

        let derived_group_items = sxvi_payloads_for_sxvd(&pivot_records, 3);
        assert_eq!(
            derived_group_items
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![1, 0, 0, 0],
            "International should be hidden while Coastal, Interior, and the automatic subtotal item remain visible"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
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
    fn semantic_pivot_tables_emit_xlsb_axis_field_options() {
        let mut wb = Workbook::new();
        add_axis_options_pivot(&mut wb);
        let bytes = write_xlsb_bytes(&wb);

        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let row_field_payload = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVD && payload[0] & 0x01 != 0)
                    .then_some(payload)
            })
            .next()
            .expect("row BrtBeginSXVD payload");

        assert_eq!(row_field_payload[0], 0x01);
        assert_eq!(
            u16::from_le_bytes(row_field_payload[1..3].try_into().unwrap()),
            0x000E
        );
        assert_eq!(row_field_payload[3] & 0x72, 0x72);

        let behavior_flags = u32::from_le_bytes(row_field_payload[8..12].try_into().unwrap());
        assert_ne!(behavior_flags & (1 << 5), 0, "show-all-items bit");
        assert_ne!(behavior_flags & (1 << 7), 0, "insert-blank-row bit");
        assert_eq!(behavior_flags & (1 << 8), 0, "subtotal-at-top bit");
        assert_ne!(behavior_flags & (1 << 11), 0, "page-break bit");
        assert_ne!(behavior_flags & (1 << 12), 0, "autosort bit");
        assert_eq!(behavior_flags & (1 << 13), 0, "descending sort bit");
        assert_eq!(
            behavior_flags & (1 << 18),
            0,
            "new items are included in existing filters"
        );
        assert_eq!(
            u32::from_le_bytes(row_field_payload[12..16].try_into().unwrap()),
            25,
            "item-page count"
        );

        let (caption, caption_len) = wide_string_at(row_field_payload, 20);
        let (subtotal_caption, _) = wide_string_at(row_field_payload, 20 + caption_len);
        assert_eq!(caption, "Market");
        assert_eq!(subtotal_caption, "? subtotal");

        let item_types = sxvi_item_types_for_sxvd(&pivot_records, 0);
        assert!(item_types.ends_with(&[0x02, 0x03, 0x04]));
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_axis_field_options() {
        let mut wb = Workbook::new();
        add_axis_options_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];
        let field = &pivot.rows[0];

        assert_eq!(pivot.name, "AxisOptions");
        assert_eq!(field.field.name, "Region");
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
        assert_eq!(field.item_page_count, 25);
    }

    #[test]
    fn xlsb_writer_rejects_item_page_count_with_advanced_filter() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        ws.set_cell_value("A2", "East").unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();

        let mut region = PivotField::new("Region");
        region.item_page_count = 25;
        let pivot = PivotTable::builder("ItemPageCount")
            .source_range(CellRange::parse("A1:B2").unwrap())
            .target_address("D1")
            .unwrap()
            .row(region)
            .filter(PivotFilter::Label {
                field: PivotFieldRef::new("Region"),
                operator: PivotFilterOperator::Contains,
                value: "E".to_string(),
            })
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let mut buf = Vec::new();
        let err = XlsbWriter::write(&wb, Cursor::new(&mut buf))
            .expect_err("item page count with advanced filter should fail for XLSB");
        assert!(err.to_string().contains("item page count"), "{err}");
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_sort_by_measure() {
        let mut wb = Workbook::new();
        add_sort_by_measure_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let record_types = pivot_records
            .iter()
            .map(|(record_type, _)| *record_type)
            .collect::<Vec<_>>();
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_SXVD14));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_PIVOT_AREA));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_PIVOT_AREA_REF));

        let measure_ref = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_PIVOT_AREA_REF_ITEM).then_some(payload)
            })
            .expect("sort-by-measure pivot-area item");
        assert_eq!(u32::from_le_bytes(measure_ref[0..4].try_into().unwrap()), 0);

        let roundtrip = XlsbReader::read(Cursor::new(bytes)).unwrap();
        let pivot = &roundtrip.worksheet(0).unwrap().pivot_tables()[0];
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
    fn semantic_pivot_tables_emit_xlsb_multi_measure_values_axis_records() {
        let mut wb = Workbook::new();
        add_multi_measure_column_values_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));

        let col_fields_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_ISXVD_COLS).then_some(payload)
            })
            .expect("column axis field declaration");
        assert_eq!(
            col_fields_payload,
            &vec![1, 0, 0, 0, 0xFE, 0xFF, 0xFF, 0xFF],
            "multiple measures are represented by the synthetic Values field (-2)"
        );

        let mut in_column_items = false;
        let mut column_line_data_items = Vec::new();
        let mut column_item_values = Vec::new();
        for (record_type, payload) in &pivot_records {
            match *record_type {
                crate::biff12::records::BRT_BEGIN_SX_COL_ITEMS => {
                    in_column_items = true;
                    assert_eq!(payload, &vec![2, 0, 0, 0], "one line per data item");
                }
                crate::biff12::records::BRT_END_SX_COL_ITEMS => in_column_items = false,
                crate::biff12::records::BRT_BEGIN_SXLI if in_column_items => {
                    column_line_data_items
                        .push(u32::from_le_bytes(payload[8..12].try_into().unwrap()));
                }
                crate::biff12::records::BRT_SXLI_ITEM if in_column_items => {
                    column_item_values.push(u32::from_le_bytes(payload[..4].try_into().unwrap()));
                }
                _ => {}
            }
        }
        assert_eq!(column_line_data_items, vec![0, 1]);
        assert_eq!(column_item_values, vec![0, 1]);

        let sx_view_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVIEW).then_some(payload)
            })
            .expect("pivot view");
        assert_eq!(sx_view_payload[12], 0x02, "Values field is on columns");
        assert_eq!(
            i32::from_le_bytes(sx_view_payload[16..20].try_into().unwrap()),
            0,
            "Values field is inserted at the requested axis position"
        );

        let data_field_count = pivot_records
            .iter()
            .filter(|(record_type, _)| *record_type == crate::biff12::records::BRT_BEGIN_SXDI)
            .count();
        assert_eq!(data_field_count, 2);

        let data_field_aggregates = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXDI)
                    .then(|| u32::from_le_bytes(payload[4..8].try_into().unwrap()))
            })
            .collect::<Vec<_>>();
        assert_eq!(data_field_aggregates, vec![0x00, 0x02]);
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_all_aggregate_data_field_records() {
        let mut wb = Workbook::new();
        add_all_aggregate_column_values_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));

        let data_field_aggregates = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXDI)
                    .then(|| u32::from_le_bytes(payload[4..8].try_into().unwrap()))
            })
            .collect::<Vec<_>>();
        let expected_codes = xlsb_all_aggregate_cases()
            .iter()
            .map(|(_, _, code)| *code)
            .collect::<Vec<_>>();
        assert_eq!(
            data_field_aggregates, expected_codes,
            "BIFF12 data-field records should preserve every pivot aggregate code"
        );

        let data_field_count = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXDIS)
                    .then(|| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
            })
            .expect("data-field collection record");
        assert_eq!(data_field_count, xlsb_all_aggregate_cases().len() as u32);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_all_aggregate_data_fields() {
        let mut wb = Workbook::new();
        add_all_aggregate_column_values_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "AllAggregatePivot");
        assert_eq!(pivot.layout.values_axis, PivotValuesAxis::Columns);
        assert_eq!(pivot.layout.values_axis_position, Some(0));
        assert_eq!(pivot.measures.len(), xlsb_all_aggregate_cases().len());
        for (measure, (aggregate, caption, _)) in
            pivot.measures.iter().zip(xlsb_all_aggregate_cases())
        {
            assert_eq!(measure.field.name, "Revenue");
            assert_eq!(measure.aggregate, aggregate);
            assert_eq!(measure.name.as_deref(), Some(caption));
        }
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_custom_measure_number_format() {
        let mut wb = Workbook::new();
        add_custom_measure_format_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let style_records = records_with_payload(read_zip_entry_bytes(&bytes, "xl/styles.bin"));
        let custom_format = style_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_FMT).then_some(payload)
            })
            .expect("custom number format record");
        assert_eq!(
            u16::from_le_bytes(custom_format[0..2].try_into().unwrap()),
            164
        );
        let (format_code, _) = crate::biff12::parser::wide_str(custom_format, 2).unwrap();
        assert_eq!(format_code, "#,##0.0");

        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let sxdi = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXDI).then_some(payload)
            })
            .expect("data field record");
        assert_eq!(u32::from_le_bytes(sxdi[20..24].try_into().unwrap()), 164);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_custom_measure_number_format() {
        let mut wb = Workbook::new();
        add_custom_measure_format_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "FormattedRevenue");
        assert_eq!(pivot.measures.len(), 1);
        assert_eq!(pivot.measures[0].number_format.as_deref(), Some("#,##0.0"));
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_show_as_data_field_records() {
        let mut wb = Workbook::new();
        add_show_as_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));

        let data_fields = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXDI).then_some(payload)
            })
            .collect::<Vec<_>>();
        let show_as_codes = data_fields
            .iter()
            .map(|payload| u32::from_le_bytes(payload[8..12].try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(show_as_codes, vec![5, 6, 7, 8, 4, 1, 3]);

        let base_field_indexes = data_fields
            .iter()
            .map(|payload| u32::from_le_bytes(payload[12..16].try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(base_field_indexes, vec![0, 0, 0, 0, 0, 0, 0]);

        let base_item_indexes = data_fields
            .iter()
            .map(|payload| u32::from_le_bytes(payload[16..20].try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(base_item_indexes, vec![0, 0, 0, 0, 0, 0, 0]);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_show_as_data_fields() {
        let mut wb = Workbook::new();
        add_show_as_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "ShowAsPivot");
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
    fn semantic_pivot_tables_emit_xlsb_x14_show_as_data_field_records() {
        let mut wb = Workbook::new();
        add_x14_show_as_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));

        let data_fields = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXDI).then_some(payload)
            })
            .collect::<Vec<_>>();
        let show_as_codes = data_fields
            .iter()
            .map(|payload| u32::from_le_bytes(payload[8..12].try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(show_as_codes, vec![0, 0, 0, 0, 0]);

        let base_field_indexes = data_fields
            .iter()
            .map(|payload| u32::from_le_bytes(payload[12..16].try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(base_field_indexes, vec![0, 0, 1, 1, 0]);

        let x14_payloads = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_SXDI14).then_some(payload)
            })
            .collect::<Vec<_>>();
        let x14_codes = x14_payloads
            .iter()
            .map(|payload| u32::from_le_bytes(payload[4..8].try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(x14_codes, vec![9, 10, 11, 13, 14]);
        for payload in x14_payloads {
            assert_eq!(payload.len(), 13);
            assert_eq!(u32::from_le_bytes(payload[0..4].try_into().unwrap()), 0);
            assert_eq!(i32::from_le_bytes(payload[8..12].try_into().unwrap()), -1);
            assert_eq!(payload[12], 0);
        }
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_x14_show_as_data_fields() {
        let mut wb = Workbook::new();
        add_x14_show_as_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "X14ShowAsPivot");
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
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_style_record() {
        let mut wb = Workbook::new();
        add_styled_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let style_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_SX_VIEW_STYLE).then_some(payload)
            })
            .expect("pivot style record");

        assert_eq!(
            u16::from_le_bytes(style_payload[0..2].try_into().unwrap()),
            0x002E,
            "style flags should encode last-column, row/column stripes, and column headers"
        );
        assert_eq!(wide_string(&style_payload[2..]), "PivotStyleLight16");
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_style() {
        let mut wb = Workbook::new();
        add_styled_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.style, styled_pivot_style());
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_calculated_field_records() {
        let mut wb = Workbook::new();
        add_calculated_field_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));

        let calculated_field = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_FIELD).then_some(payload)
            })
            .find(|payload| wide_string_at(payload, 20).0 == "Revenue")
            .expect("calculated-field BrtBeginPCDField");
        assert_ne!(
            u16::from_le_bytes(calculated_field[0..2].try_into().unwrap()) & 0x0100,
            0,
            "BrtBeginPCDField should set fLoadFmla for calculated fields"
        );
        assert_eq!(
            u16::from_le_bytes(calculated_field[0..2].try_into().unwrap()) & 0x0004,
            0,
            "calculated fields should not be marked as source database fields"
        );

        let (_, name_len) = wide_string_at(calculated_field, 20);
        let formula_offset = 20 + name_len;
        assert_eq!(
            u32::from_le_bytes(
                calculated_field[formula_offset..formula_offset + 4]
                    .try_into()
                    .unwrap()
            ),
            13
        );
        assert_eq!(
            &calculated_field[formula_offset + 4..formula_offset + 17],
            &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x05],
            "PivotParsedFormula should encode Units * Price via two PtgSxName refs"
        );
        assert_eq!(
            u32::from_le_bytes(
                calculated_field[formula_offset + 17..formula_offset + 21]
                    .try_into()
                    .unwrap()
            ),
            0,
            "PivotParsedFormula should not carry extra rgcb data for this expression"
        );

        let pnames_count = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PNAMES)
                    .then(|| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
            })
            .expect("BrtBeginPNames");
        assert_eq!(pnames_count, 2);
        let pname_field_indexes = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PNAME)
                    .then(|| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
            })
            .collect::<Vec<_>>();
        assert_eq!(pname_field_indexes, vec![1, 2]);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_calculated_field() {
        let mut wb = Workbook::new();
        add_calculated_field_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "CalculatedRevenue");
        assert_eq!(pivot.calculated_fields.len(), 1);
        assert_eq!(pivot.calculated_fields[0].name, "Revenue");
        assert_eq!(pivot.calculated_fields[0].formula, "Units*Price");
        assert_eq!(pivot.measures.len(), 1);
        assert_eq!(pivot.measures[0].field.name, "Revenue");
        assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_calculated_field_function_records() {
        let mut wb = Workbook::new();
        add_calculated_field_function_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let calculated_field = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_FIELD).then_some(payload)
            })
            .find(|payload| wide_string_at(payload, 20).0 == "Revenue")
            .expect("calculated-field BrtBeginPCDField");
        let (_, name_len) = wide_string_at(calculated_field, 20);
        let formula_offset = 20 + name_len;
        assert_eq!(
            u32::from_le_bytes(
                calculated_field[formula_offset..formula_offset + 4]
                    .try_into()
                    .unwrap()
            ),
            16
        );
        assert_eq!(
            &calculated_field[formula_offset + 4..formula_offset + 20],
            &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x42, 0x02, 0x04, 0x00],
            "PivotParsedFormula should encode SUM(Units,Price)"
        );
        let pname_field_indexes = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PNAME)
                    .then(|| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
            })
            .collect::<Vec<_>>();
        assert_eq!(pname_field_indexes, vec![1, 2]);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_calculated_field_function() {
        let mut wb = Workbook::new();
        add_calculated_field_function_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "CalculatedRevenueFunction");
        assert_eq!(pivot.calculated_fields.len(), 1);
        assert_eq!(pivot.calculated_fields[0].name, "Revenue");
        assert_eq!(pivot.calculated_fields[0].formula, "SUM(Units,Price)");
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_calculated_item_records() {
        let mut wb = Workbook::new();
        add_calculated_item_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let record_types = cache_records
            .iter()
            .map(|(record_type, _)| *record_type)
            .collect::<Vec<_>>();

        assert!(record_types.contains(&0x00F3), "BrtBeginPCDCalcItems");
        assert!(record_types.contains(&0x00F5), "BrtBeginPCDCalcItem");
        assert!(record_types.contains(&0x017E), "BrtBeginPRFItem");
        assert!(
            cache_records
                .iter()
                .any(|(record_type, payload)| *record_type == 0x001F
                    && wide_string_at(payload, 0).0 == "Combined"
                    && payload[payload.len() - 6..payload.len()] == [0x02, 0, 0, 0, 0, 0]),
            "calculated item should be stored as a PCDIAString with fFmla set"
        );

        let formula = cache_records
            .iter()
            .find_map(|(record_type, payload)| (*record_type == 0x00F5).then_some(payload))
            .expect("BrtBeginPCDCalcItem");
        assert_eq!(&formula[0..4], &[0xFF, 0xFF, 0xFF, 0xFF]);
        assert_eq!(u32::from_le_bytes(formula[4..8].try_into().unwrap()), 13);
        assert_eq!(
            &formula[8..21],
            &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x03],
            "calculated item formula should encode East + West via two PtgSxName refs"
        );
        assert_eq!(u32::from_le_bytes(formula[21..25].try_into().unwrap()), 0);

        let pr_filter = cache_records
            .iter()
            .find_map(|(record_type, payload)| (*record_type == 0x00FB).then_some(payload))
            .expect("BrtBeginPRFilter");
        assert_eq!(u32::from_le_bytes(pr_filter[0..4].try_into().unwrap()), 0);
        assert_eq!(u32::from_le_bytes(pr_filter[4..8].try_into().unwrap()), 1);
        let pr_item = cache_records
            .iter()
            .find_map(|(record_type, payload)| (*record_type == 0x017E).then_some(payload))
            .expect("BrtBeginPRFItem");
        assert_eq!(u32::from_le_bytes(pr_item[0..4].try_into().unwrap()), 2);

        let pnpairs = cache_records
            .iter()
            .filter_map(|(record_type, payload)| (*record_type == 0x0103).then_some(payload))
            .map(|payload| {
                (
                    u32::from_le_bytes(payload[1..5].try_into().unwrap()),
                    i32::from_le_bytes(payload[5..9].try_into().unwrap()),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(pnpairs, vec![(0, 0), (0, 1)]);

        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let row_sxvis = sxvi_payloads_for_sxvd(&pivot_records, 0);
        assert!(
            row_sxvis.iter().any(|payload| {
                u16::from_le_bytes(payload[1..3].try_into().unwrap()) & 0x0008 != 0
            }),
            "pivot item should be marked as calculated"
        );
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_calculated_item() {
        let mut wb = Workbook::new();
        add_calculated_item_pivot(&mut wb);

        let wb2 = round_trip(&wb);
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
    fn semantic_pivot_tables_emit_xlsb_calculated_item_cell_like_records() {
        let mut wb = Workbook::new();
        add_calculated_item_cell_like_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let formula = cache_records
            .iter()
            .find_map(|(record_type, payload)| (*record_type == 0x00F5).then_some(payload))
            .expect("BrtBeginPCDCalcItem");
        assert_eq!(u32::from_le_bytes(formula[4..8].try_into().unwrap()), 13);
        assert_eq!(
            &formula[8..21],
            &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x03],
            "cell-like calculated item names should encode as PtgSxName refs"
        );

        let pnpairs = cache_records
            .iter()
            .filter_map(|(record_type, payload)| (*record_type == 0x0103).then_some(payload))
            .map(|payload| {
                (
                    u32::from_le_bytes(payload[1..5].try_into().unwrap()),
                    i32::from_le_bytes(payload[5..9].try_into().unwrap()),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(pnpairs, vec![(0, 0), (0, 1)]);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_calculated_item_cell_like() {
        let mut wb = Workbook::new();
        add_calculated_item_cell_like_pivot(&mut wb);

        let wb2 = round_trip(&wb);
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
    fn semantic_pivot_tables_emit_xlsb_calculated_item_string_ref_records() {
        let mut wb = Workbook::new();
        add_calculated_item_string_ref_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let formula_lengths = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == 0x00F5)
                    .then(|| u32::from_le_bytes(payload[4..8].try_into().unwrap()))
            })
            .collect::<Vec<_>>();
        assert_eq!(
            formula_lengths,
            vec![13, 13],
            "string item references should compile to the same PtgSxName shape as bare item refs"
        );

        let pnpairs = cache_records
            .iter()
            .filter_map(|(record_type, payload)| (*record_type == 0x0103).then_some(payload))
            .map(|payload| {
                (
                    u32::from_le_bytes(payload[1..5].try_into().unwrap()),
                    i32::from_le_bytes(payload[5..9].try_into().unwrap()),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(pnpairs.len(), 4);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_calculated_item_string_ref() {
        let mut wb = Workbook::new();
        add_calculated_item_string_ref_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "CalculatedRegionStringRef");
        assert_eq!(pivot.calculated_items.len(), 2);
        assert_eq!(pivot.calculated_items[0].field.name, "Region");
        assert_eq!(
            pivot.calculated_items[0].item,
            PivotValue::String("All Regions".into())
        );
        assert_eq!(pivot.calculated_items[0].formula, "Combined+Central");
        assert_eq!(
            pivot.calculated_items[1].item,
            PivotValue::String("Combined".into())
        );
        assert_eq!(pivot.calculated_items[1].formula, "East+West");
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_calculated_item_function_records() {
        let mut wb = Workbook::new();
        add_calculated_item_function_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let formula = cache_records
            .iter()
            .find_map(|(record_type, payload)| (*record_type == 0x00F5).then_some(payload))
            .expect("BrtBeginPCDCalcItem");
        assert_eq!(u32::from_le_bytes(formula[4..8].try_into().unwrap()), 16);
        assert_eq!(
            &formula[8..24],
            &[0x18, 0x1D, 0, 0, 0, 0, 0x18, 0x1D, 1, 0, 0, 0, 0x42, 0x02, 0x07, 0x00],
            "calculated item formula should encode MAX(East,West)"
        );
        let pnpairs = cache_records
            .iter()
            .filter_map(|(record_type, payload)| (*record_type == 0x0103).then_some(payload))
            .map(|payload| {
                (
                    u32::from_le_bytes(payload[1..5].try_into().unwrap()),
                    i32::from_le_bytes(payload[5..9].try_into().unwrap()),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(pnpairs, vec![(0, 0), (0, 1)]);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_calculated_item_function() {
        let mut wb = Workbook::new();
        add_calculated_item_function_pivot(&mut wb);

        let wb2 = round_trip(&wb);
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
    fn semantic_pivot_tables_emit_xlsb_numeric_grouping_records() {
        let mut wb = Workbook::new();
        add_numeric_grouped_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let record_types = cache_records
            .iter()
            .map(|(record_type, _)| *record_type)
            .collect::<Vec<_>>();

        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_PCDF_GROUP));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_PCDFG_RANGE));
        assert!(record_types.contains(&crate::biff12::records::BRT_BEGIN_PCDFG_ITEMS));
        assert!(record_types.contains(&crate::biff12::records::BRT_END_PCDF_GROUP));

        let shared_items_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_SHARED_ITEMS)
                    .then_some(payload)
            })
            .expect("BrtBeginPCDFAtbl payload");
        let shared_items_flags = u16::from_le_bytes(shared_items_payload[0..2].try_into().unwrap());
        assert_eq!(
            shared_items_flags & 0x0001,
            0,
            "number-only shared items must not set fTextEtcField"
        );
        assert_eq!(
            u32::from_le_bytes(shared_items_payload[2..6].try_into().unwrap()),
            4,
            "numeric grouped base field stores raw source items"
        );
        assert_eq!(
            f64::from_le_bytes(shared_items_payload[6..14].try_into().unwrap()),
            5.0
        );
        assert_eq!(
            f64::from_le_bytes(shared_items_payload[14..22].try_into().unwrap()),
            41.0
        );

        let group_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDF_GROUP).then_some(payload)
            })
            .expect("BrtBeginPCDFGroup payload");
        assert_eq!(
            i32::from_le_bytes(group_payload[0..4].try_into().unwrap()),
            -1,
            "numeric grouping has no parent grouped field"
        );
        assert_eq!(
            i32::from_le_bytes(group_payload[4..8].try_into().unwrap()),
            0,
            "numeric grouping base field is Age at cache field index 0"
        );

        let range_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDFG_RANGE).then_some(payload)
            })
            .expect("BrtBeginPCDFGRange payload");
        assert_eq!(range_payload[0], 0, "range grouping type");
        assert_eq!(range_payload[1], 0, "explicit numeric start/end flags");
        assert_eq!(
            f64::from_le_bytes(range_payload[2..10].try_into().unwrap()),
            0.0
        );
        assert_eq!(
            f64::from_le_bytes(range_payload[10..18].try_into().unwrap()),
            60.0
        );
        assert_eq!(
            f64::from_le_bytes(range_payload[18..26].try_into().unwrap()),
            10.0
        );

        let items_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDFG_ITEMS).then_some(payload)
            })
            .expect("BrtBeginPCDFGItems payload");
        assert_eq!(
            u32::from_le_bytes(items_payload[0..4].try_into().unwrap()),
            8,
            "numeric grouping emits underflow, interval, and overflow group items"
        );
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_numeric_grouping() {
        let mut wb = Workbook::new();
        add_numeric_grouped_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "GroupedAges");
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
            other => panic!("unexpected grouping: {other:?}"),
        }
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_date_grouping_records() {
        let mut wb = Workbook::new();
        add_date_grouped_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        assert!(
            cache_records
                .iter()
                .any(|(record_type, _)| *record_type == crate::biff12::records::BRT_PCDI_DATETIME),
            "date grouped cache field should store shared items as datetime records"
        );
        let range_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDFG_RANGE).then_some(payload)
            })
            .expect("BrtBeginPCDFGRange payload");

        assert_eq!(range_payload[0], 0x05, "month date grouping type");
        assert_eq!(
            range_payload[1] & 0x03,
            0x03,
            "date grouping uses automatic range bounds"
        );
        assert_eq!(
            range_payload[1] & 0x04,
            0x04,
            "date grouping marks the range payload as dates"
        );
        assert_eq!(
            f64::from_le_bytes(range_payload[18..26].try_into().unwrap()),
            1.0
        );
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_date_grouping() {
        let mut wb = Workbook::new();
        add_date_grouped_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "MonthlyRevenue");
        assert_eq!(pivot.groupings.len(), 1);
        match &pivot.groupings[0] {
            PivotGrouping::Date { field, units } => {
                assert_eq!(field.name, "Date");
                assert_eq!(*units, vec![PivotDateGroupUnit::Months]);
            }
            other => panic!("unexpected grouping: {other:?}"),
        }
    }

    #[test]
    fn xlsb_writer_round_trips_multi_unit_date_grouped_page_field() {
        let mut wb = Workbook::new();
        add_multi_unit_date_grouped_page_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let page_fields_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPIS).then_some(payload)
            })
            .expect("page field collection payload");
        assert_eq!(
            u32::from_le_bytes(page_fields_payload[0..4].try_into().unwrap()),
            2,
            "multi-unit date grouped page fields should expand to one native page field per date unit"
        );
        let page_field_payloads = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
            })
            .collect::<Vec<_>>();
        assert_eq!(page_field_payloads.len(), 2);
        let page_field_indexes = page_field_payloads
            .iter()
            .map(|payload| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(
            page_field_indexes,
            vec![2, 3],
            "native page fields should point at the Years and Months derived cache fields"
        );
        assert!(
            page_field_payloads.iter().all(|payload| {
                u32::from_le_bytes(payload[4..8].try_into().unwrap()) == 0x0010_00FE
            }),
            "unfiltered expanded page fields should use Excel's all-items sentinel"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("GroupedDatePage")
            .expect("pivot after read");
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
            other => panic!("expected multi-unit date grouping, got {other:?}"),
        }
    }

    #[test]
    fn xlsb_writer_rejects_multi_unit_date_grouped_page_item_filter() {
        let mut wb = Workbook::new();
        add_multi_unit_date_grouped_page_pivot(&mut wb);
        wb.worksheet_mut(0)
            .unwrap()
            .pivot_tables_mut()
            .get_mut(0)
            .unwrap()
            .filters
            .push(PivotFilter::field_items("SaleDate", [2024.0]));

        let mut buf = Vec::new();
        let err = XlsbWriter::write(&wb, Cursor::new(&mut buf))
            .expect_err("unit-less multi-unit date page filters should fail for XLSB");
        assert!(
            err.to_string().contains("filters date-grouped page field"),
            "{err}"
        );
    }

    #[test]
    fn xlsb_writer_round_trips_multi_unit_date_grouped_page_item_filters() {
        let mut wb = Workbook::new();
        add_multi_unit_date_grouped_page_filter_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let page_field_payloads = pivot_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
            })
            .collect::<Vec<_>>();
        assert_eq!(page_field_payloads.len(), 2);
        assert_eq!(
            page_field_payloads
                .iter()
                .map(|payload| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![2, 3],
            "page filters should target the Years and Months derived cache fields"
        );
        assert_eq!(
            u32::from_le_bytes(page_field_payloads[0][4..8].try_into().unwrap()),
            0,
            "2024 should be the first grouped year item"
        );
        assert_eq!(
            u32::from_le_bytes(page_field_payloads[1][4..8].try_into().unwrap()),
            0x0010_00FE,
            "multi-selected months should use Excel's multiple-items sentinel"
        );
        assert_eq!(
            sxvi_payloads_for_sxvd(&pivot_records, 2)
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![0, 0, 0],
            "single selected Years page item should be carried by SXPI, not hidden flags"
        );
        assert_eq!(
            sxvi_payloads_for_sxvd(&pivot_records, 3)
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![0, 0, 1, 0],
            "March should be hidden while January, February, and the automatic subtotal item stay visible"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("GroupedDatePage")
            .expect("pivot after read");
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
            other => panic!("expected multi-unit date grouping, got {other:?}"),
        }
        assert_eq!(
            pivot.filters,
            vec![
                PivotFilter::field_items("SaleDate (Years)", [2024.0]),
                PivotFilter::field_items("SaleDate (Months)", [1.0, 2.0]),
            ]
        );
    }

    #[test]
    fn xlsb_writer_round_trips_date_grouped_page_item_filter() {
        let mut wb = Workbook::new();
        add_page_date_grouped_pivot(&mut wb, &[1.0]);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let page_field_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
            })
            .expect("page axis field declaration");
        assert_eq!(
            u32::from_le_bytes(page_field_payload[0..4].try_into().unwrap()),
            0,
            "single-unit date grouped page fields should select on the grouped source field"
        );
        assert_eq!(
            u32::from_le_bytes(page_field_payload[4..8].try_into().unwrap()),
            0,
            "January should be the first grouped page item"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("MonthlyPageFilter")
            .expect("pivot after read");
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
    }

    #[test]
    fn xlsb_writer_round_trips_date_grouped_page_multi_item_filter() {
        let mut wb = Workbook::new();
        add_page_date_grouped_pivot(&mut wb, &[1.0, 2.0]);

        let bytes = write_xlsb_bytes(&wb);
        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let page_field_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXPI).then_some(payload)
            })
            .expect("page axis field declaration");
        assert_eq!(
            u32::from_le_bytes(page_field_payload[0..4].try_into().unwrap()),
            0,
            "single-unit date grouped page fields should select on the grouped source field"
        );
        assert_eq!(
            u32::from_le_bytes(page_field_payload[4..8].try_into().unwrap()),
            0x0010_00FE,
            "multi-select grouped page filters should use Excel's multiple-items sentinel"
        );
        let page_items = sxvi_payloads_for_sxvd(&pivot_records, 0);
        assert_eq!(
            page_items
                .iter()
                .map(|payload| u16::from_le_bytes(payload[1..3].try_into().unwrap()))
                .collect::<Vec<_>>(),
            vec![0, 0, 1, 0],
            "March should be hidden while January, February, and the automatic subtotal item stay visible"
        );

        let read = XlsbReader::read(Cursor::new(bytes)).expect("read xlsb");
        let pivot = read
            .worksheet(0)
            .unwrap()
            .pivot_table_by_name("MonthlyPageFilter")
            .expect("pivot after read");
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
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_multi_unit_date_grouping_records() {
        let mut wb = Workbook::new();
        add_multi_unit_date_grouped_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let field_count_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_FIELDS).then_some(payload)
            })
            .expect("BrtBeginPCDFields payload");
        assert_eq!(
            u32::from_le_bytes(field_count_payload[0..4].try_into().unwrap()),
            4,
            "multi-unit date grouping should add one derived cache field per unit"
        );

        let field_groups = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDF_GROUP).then(|| {
                    (
                        i32::from_le_bytes(payload[0..4].try_into().unwrap()),
                        i32::from_le_bytes(payload[4..8].try_into().unwrap()),
                    )
                })
            })
            .collect::<Vec<_>>();
        assert_eq!(
            field_groups,
            vec![(2, -1), (-1, 0), (-1, 0)],
            "base date field should point at the outer derived unit and derived date unit fields should point back to the base date field"
        );

        let range_group_bys = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDFG_RANGE).then(|| payload[0])
            })
            .collect::<Vec<_>>();
        assert_eq!(
            range_group_bys,
            vec![0x07, 0x05],
            "derived fields should encode years then months"
        );

        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let row_fields_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_ISXVD_RWS).then_some(payload)
            })
            .expect("BrtBeginISXVDRws payload");
        assert_eq!(
            u32::from_le_bytes(row_fields_payload[0..4].try_into().unwrap()),
            2,
            "multi-unit date grouping should expand the row axis to the unit fields"
        );
        let row_field_indexes = row_fields_payload[4..]
            .chunks_exact(4)
            .map(|chunk| i32::from_le_bytes(chunk.try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(row_field_indexes, vec![2, 3]);

        let sxvd_item_list_counts = sxvd_item_list_counts(&pivot_records);
        assert_eq!(
            sxvd_item_list_counts.len(),
            4,
            "pivot view should include source fields plus one derived field per date unit"
        );
        assert_eq!(
            sxvd_item_list_counts[0], 0,
            "the hidden base date field should not emit view item records"
        );
        assert_eq!(
            &sxvd_item_list_counts[2..],
            &[1, 1],
            "derived date unit fields should carry their visible item lists"
        );

        let cache_rows = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheRecords1.bin",
        ))
        .into_iter()
        .filter_map(|(record_type, payload)| {
            (record_type == crate::biff12::records::BRT_PCD_RECORD).then_some(payload)
        })
        .collect::<Vec<_>>();
        assert_eq!(cache_rows.len(), 3);
        assert!(
            cache_rows.iter().all(|payload| payload.len() == 12),
            "BIFF12 date-unit derived fields should not widen cache record rows"
        );
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_multi_unit_date_grouping() {
        let mut wb = Workbook::new();
        add_multi_unit_date_grouped_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

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
            other => panic!("unexpected grouping: {other:?}"),
        }
    }

    #[test]
    fn semantic_pivot_tables_emit_xlsb_manual_grouping_records() {
        let mut wb = Workbook::new();
        add_manual_grouped_pivot(&mut wb);

        let bytes = write_xlsb_bytes(&wb);
        let cache_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotCache/pivotCacheDefinition1.bin",
        ));
        let field_count_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCD_FIELDS).then_some(payload)
            })
            .expect("BrtBeginPCDFields payload");
        assert_eq!(
            u32::from_le_bytes(field_count_payload[0..4].try_into().unwrap()),
            3,
            "manual grouping should add a derived cache field"
        );

        let field_groups = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDF_GROUP).then(|| {
                    (
                        i32::from_le_bytes(payload[0..4].try_into().unwrap()),
                        i32::from_le_bytes(payload[4..8].try_into().unwrap()),
                    )
                })
            })
            .collect::<Vec<_>>();
        assert_eq!(
            field_groups,
            vec![(2, -1), (-1, 0)],
            "base field should point at derived field and derived field should point back to base"
        );

        let discrete_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDFG_DISCRETE)
                    .then_some(payload)
            })
            .expect("BrtBeginPCDFGDiscrete payload");
        assert_eq!(
            u32::from_le_bytes(discrete_payload[0..4].try_into().unwrap()),
            3,
            "manual grouping should map each source shared item"
        );
        assert_eq!(
            discrete_payload.len(),
            4,
            "discrete group begin record carries only the item count"
        );
        let discrete_indexes = cache_records
            .iter()
            .filter_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_PCDI_INDEX)
                    .then(|| u32::from_le_bytes(payload[0..4].try_into().unwrap()))
            })
            .collect::<Vec<_>>();
        assert_eq!(
            discrete_indexes,
            vec![1, 1, 0],
            "East and West should map to Coastal while Central remains ungrouped"
        );

        let group_items_payload = cache_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_PCDFG_ITEMS).then_some(payload)
            })
            .expect("BrtBeginPCDFGItems payload");
        assert_eq!(
            u32::from_le_bytes(group_items_payload[0..4].try_into().unwrap()),
            2,
            "manual grouping should emit Central plus the Coastal group item"
        );

        let pivot_records = records_with_payload(read_zip_entry_bytes(
            &bytes,
            "xl/pivotTables/pivotTable1.bin",
        ));
        let view_field_count_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_SXVDS).then_some(payload)
            })
            .expect("BrtBeginSXVDs payload");
        assert_eq!(
            u32::from_le_bytes(view_field_count_payload[0..4].try_into().unwrap()),
            3,
            "pivot view should include the derived cache field"
        );
        let row_fields_payload = pivot_records
            .iter()
            .find_map(|(record_type, payload)| {
                (*record_type == crate::biff12::records::BRT_BEGIN_ISXVD_RWS).then_some(payload)
            })
            .expect("BrtBeginISXVDRws payload");
        assert_eq!(
            u32::from_le_bytes(row_fields_payload[0..4].try_into().unwrap()),
            2,
            "manual grouping should expand the row axis to derived plus base"
        );
        let row_field_indexes = row_fields_payload[4..]
            .chunks_exact(4)
            .map(|chunk| i32::from_le_bytes(chunk.try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(row_field_indexes, vec![2, 0]);
    }

    #[test]
    fn semantic_pivot_tables_round_trip_xlsb_manual_grouping() {
        let mut wb = Workbook::new();
        add_manual_grouped_pivot(&mut wb);

        let wb2 = round_trip(&wb);
        let pivot = &wb2.worksheet(0).unwrap().pivot_tables()[0];

        assert_eq!(pivot.name, "ManualGroupedRegions");
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
    fn string_values() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "hello").unwrap();
        ws.set_cell_value_at(0, 1, "world").unwrap();
        ws.set_cell_value_at(1, 0, "hello").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("hello"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::string("world"));
        assert_eq!(ws2.get_value_at(1, 0), CellValue::string("hello"));
    }

    #[test]
    fn numeric_values() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 42.0).unwrap();
        ws.set_cell_value_at(0, 1, -3.14).unwrap();
        ws.set_cell_value_at(0, 2, 0.0).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(42.0));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Number(-3.14));
        assert_eq!(ws2.get_value_at(0, 2), CellValue::Number(0.0));
    }

    #[test]
    fn boolean_values() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, true).unwrap();
        ws.set_cell_value_at(0, 1, false).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Boolean(true));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Boolean(false));
    }

    #[test]
    fn error_values() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, CellValue::Error(duke_sheets_core::CellError::Div0))
            .unwrap();
        ws.set_cell_value_at(0, 1, CellValue::Error(duke_sheets_core::CellError::Na))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(
            ws2.get_value_at(0, 0),
            CellValue::Error(duke_sheets_core::CellError::Div0)
        );
        assert_eq!(
            ws2.get_value_at(0, 1),
            CellValue::Error(duke_sheets_core::CellError::Na)
        );
    }

    #[test]
    fn multiple_sheets() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0).unwrap().set_name("First");
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "sheet1")
            .unwrap();
        wb.add_worksheet().unwrap();
        wb.worksheet_mut(1).unwrap().set_name("Second");
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, 99.0)
            .unwrap();

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.sheet_count(), 2);
        assert_eq!(wb2.worksheet(0).unwrap().name(), "First");
        assert_eq!(wb2.worksheet(1).unwrap().name(), "Second");
        assert_eq!(
            wb2.worksheet(0).unwrap().get_value_at(0, 0),
            CellValue::string("sheet1")
        );
        assert_eq!(
            wb2.worksheet(1).unwrap().get_value_at(0, 0),
            CellValue::Number(99.0)
        );
    }

    #[test]
    fn mixed_cell_types() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "text").unwrap();
        ws.set_cell_value_at(0, 1, 123.456).unwrap();
        ws.set_cell_value_at(0, 2, true).unwrap();
        ws.set_cell_value_at(1, 0, CellValue::Error(duke_sheets_core::CellError::Value))
            .unwrap();
        ws.set_cell_value_at(2, 5, "sparse").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("text"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Number(123.456));
        assert_eq!(ws2.get_value_at(0, 2), CellValue::Boolean(true));
        assert_eq!(
            ws2.get_value_at(1, 0),
            CellValue::Error(duke_sheets_core::CellError::Value)
        );
        assert_eq!(ws2.get_value_at(2, 5), CellValue::string("sparse"));
        assert_eq!(ws2.get_value_at(1, 1), CellValue::Empty);
    }

    #[test]
    fn unicode_strings() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "café").unwrap();
        ws.set_cell_value_at(0, 1, "日本語").unwrap();
        ws.set_cell_value_at(0, 2, "😀").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("café"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::string("日本語"));
        assert_eq!(ws2.get_value_at(0, 2), CellValue::string("😀"));
    }

    #[test]
    fn write_file_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "file test").unwrap();
        ws.set_cell_value_at(0, 1, 42.0).unwrap();

        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsb");
        XlsbWriter::write_file(&wb, &path).unwrap();

        let wb2 = XlsbReader::read_file(&path).unwrap();
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("file test"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Number(42.0));
    }

    #[test]
    fn style_bold_font_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "bold").unwrap();
        let style = Style::new()
            .bold(true)
            .font_size(14.0)
            .font_color(Color::RED);
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("bold"));
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert!(s.font.bold);
        assert_eq!(s.font.size, 14.0);
        assert_eq!(s.font.color, Color::Rgb { r: 255, g: 0, b: 0 });
    }

    #[test]
    fn style_number_format_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 3.14159).unwrap();
        let style = Style::new().number_format("0.00");
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(3.14159));
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(s.number_format, NumberFormat::Custom("0.00".to_string()));
    }

    #[test]
    fn style_fill_color_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "colored").unwrap();
        let style = Style::new().fill_color(Color::rgb(0, 128, 255));
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(
            s.fill,
            FillStyle::Solid {
                color: Color::Rgb {
                    r: 0,
                    g: 128,
                    b: 255
                }
            }
        );
    }

    #[test]
    fn style_border_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "bordered").unwrap();
        let mut style = Style::new();
        style.border = style
            .border
            .clone()
            .with_bottom(BorderLineStyle::Thin, Color::BLACK)
            .with_top(BorderLineStyle::Medium, Color::RED);
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(
            s.border.bottom,
            Some(BorderEdge::new(BorderLineStyle::Thin, Color::BLACK))
        );
        assert_eq!(
            s.border.top,
            Some(BorderEdge::new(BorderLineStyle::Medium, Color::RED))
        );
    }

    #[test]
    fn formula_cached_number_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(0, 1, 20.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 2, "=A1+B1", CellValue::Number(30.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 2), CellValue::Number(30.0));
    }

    #[test]
    fn formula_cached_string_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(
            0,
            0,
            "=CONCATENATE(\"hello\",\" \",\"world\")",
            CellValue::string("hello world"),
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("hello world"));
    }

    #[test]
    fn formula_cached_bool_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=TRUE", CellValue::Boolean(true))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Boolean(true));
    }

    #[test]
    fn formula_cached_error_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(
            0,
            0,
            "=1/0",
            CellValue::Error(duke_sheets_core::CellError::Div0),
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(
            ws2.get_value_at(0, 0),
            CellValue::Error(duke_sheets_core::CellError::Div0)
        );
    }

    #[test]
    fn multiple_styles_across_cells() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        ws.set_cell_value_at(0, 0, "bold").unwrap();
        ws.set_cell_style_at(0, 0, &Style::new().bold(true))
            .unwrap();

        ws.set_cell_value_at(0, 1, "italic").unwrap();
        ws.set_cell_style_at(0, 1, &Style::new().italic(true))
            .unwrap();

        ws.set_cell_value_at(0, 2, "red bg").unwrap();
        ws.set_cell_style_at(0, 2, &Style::new().fill_color(Color::RED))
            .unwrap();

        ws.set_cell_value_at(1, 0, "default").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();

        let s0 = ws2.cell_style_at(0, 0).unwrap();
        assert!(s0.font.bold);
        assert!(!s0.font.italic);

        let s1 = ws2.cell_style_at(0, 1).unwrap();
        assert!(!s1.font.bold);
        assert!(s1.font.italic);

        let s2 = ws2.cell_style_at(0, 2).unwrap();
        assert_eq!(
            s2.fill,
            FillStyle::Solid {
                color: Color::Rgb { r: 255, g: 0, b: 0 }
            }
        );

        let s3 = ws2.cell_style_at(1, 0);
        assert!(s3.is_none() || *s3.unwrap() == Style::default());
    }

    #[test]
    fn style_deduplication_across_sheets() {
        let mut wb = Workbook::new();
        let bold = Style::new().bold(true);

        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "sheet1").unwrap();
        ws.set_cell_style_at(0, 0, &bold).unwrap();

        wb.add_worksheet().unwrap();
        let ws2 = wb.worksheet_mut(1).unwrap();
        ws2.set_cell_value_at(0, 0, "sheet2").unwrap();
        ws2.set_cell_style_at(0, 0, &bold).unwrap();

        let wb2 = round_trip(&wb);
        let s1 = wb2.worksheet(0).unwrap().cell_style_at(0, 0).unwrap();
        let s2 = wb2.worksheet(1).unwrap().cell_style_at(0, 0).unwrap();
        assert!(s1.font.bold);
        assert!(s2.font.bold);
    }

    #[test]
    fn formula_with_style_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=1+2", CellValue::Number(3.0))
            .unwrap();
        ws.set_cell_style_at(0, 0, &Style::new().bold(true))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(3.0));
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert!(s.font.bold);
    }

    #[test]
    fn builtin_number_format_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 0.5).unwrap();
        let mut style = Style::new();
        style.number_format = NumberFormat::BuiltIn(10);
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(s.number_format, NumberFormat::BuiltIn(10));
    }

    #[test]
    fn merged_cells_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "merged").unwrap();
        ws.merge_cells(&duke_sheets_core::CellRange::parse("A1:C3").unwrap())
            .unwrap();
        ws.merge_cells(&duke_sheets_core::CellRange::parse("E5:F6").unwrap())
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let regions = ws2.merged_regions();
        assert_eq!(regions.len(), 2);
        assert_eq!(regions[0].to_string(), "A1:C3");
        assert_eq!(regions[1].to_string(), "E5:F6");
    }

    #[test]
    fn row_height_and_hidden_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "row0").unwrap();
        ws.set_row_height(0, 30.0);
        ws.set_cell_value_at(1, 0, "row1").unwrap();
        ws.set_row_hidden(1, true);
        ws.set_cell_value_at(2, 0, "row2").unwrap();
        ws.set_row_height(2, 45.5);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert!((ws2.row_height(0) - 30.0).abs() < 0.1);
        assert!(ws2.is_row_hidden(1));
        assert!((ws2.row_height(2) - 45.5).abs() < 0.1);
    }

    #[test]
    fn column_width_and_hidden_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "col0").unwrap();
        ws.set_column_width(0, 20.0);
        ws.set_column_hidden(1, true);
        ws.set_column_width(2, 5.5);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert!((ws2.column_width(0) - 20.0).abs() < 0.1);
        assert!(ws2.is_column_hidden(1));
        assert!((ws2.column_width(2) - 5.5).abs() < 0.1);
    }

    #[test]
    fn freeze_pane_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "frozen").unwrap();
        ws.set_freeze_panes(2, 1);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fp = ws2.freeze_panes().expect("freeze panes should exist");
        assert_eq!(fp.row, 2);
        assert_eq!(fp.col, 1);
    }

    #[test]
    fn page_margins_roundtrip() {
        use duke_sheets_core::worksheet::PageSetup;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "margins").unwrap();
        let mut ps = PageSetup::default();
        ps.left_margin = 1.5;
        ps.right_margin = 1.5;
        ps.top_margin = 2.0;
        ps.bottom_margin = 2.0;
        ps.header_margin = 0.5;
        ps.footer_margin = 0.5;
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let ps2 = ws2.page_setup();
        assert!((ps2.left_margin - 1.5).abs() < 0.001);
        assert!((ps2.right_margin - 1.5).abs() < 0.001);
        assert!((ps2.top_margin - 2.0).abs() < 0.001);
        assert!((ps2.bottom_margin - 2.0).abs() < 0.001);
        assert!((ps2.header_margin - 0.5).abs() < 0.001);
        assert!((ps2.footer_margin - 0.5).abs() < 0.001);
    }

    #[test]
    fn page_setup_landscape_roundtrip() {
        use duke_sheets_core::worksheet::{PageOrientation, PageSetup};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "landscape").unwrap();
        let mut ps = PageSetup::default();
        ps.orientation = PageOrientation::Landscape;
        ps.paper_size = 9;
        ps.scale = 75;
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let ps2 = ws2.page_setup();
        assert_eq!(ps2.orientation, PageOrientation::Landscape);
        assert_eq!(ps2.paper_size, 9);
        assert_eq!(ps2.scale, 75);
    }

    #[test]
    fn autofilter_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "Name").unwrap();
        ws.set_cell_value_at(0, 1, "Value").unwrap();
        ws.set_cell_value_at(1, 0, "A").unwrap();
        ws.set_cell_value_at(1, 1, 1.0).unwrap();
        let range = duke_sheets_core::CellRange::parse("A1:B2").unwrap();
        ws.set_auto_filter(Some(duke_sheets_core::AutoFilter::new(range)));

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let af = ws2.auto_filter().expect("autofilter should exist");
        assert_eq!(af.range.to_string(), "A1:B2");
    }

    #[test]
    fn header_footer_roundtrip() {
        use duke_sheets_core::worksheet::PageSetup;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "hf").unwrap();
        let mut ps = PageSetup::default();
        ps.odd_header = Some("&CPage &P".to_string());
        ps.odd_footer = Some("&LLeft&RRight".to_string());
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let ps2 = ws2.page_setup();
        assert_eq!(ps2.odd_header.as_deref(), Some("&CPage &P"));
        assert_eq!(ps2.odd_footer.as_deref(), Some("&LLeft&RRight"));
    }

    #[test]
    fn sparse_header_footer_roundtrip() {
        use duke_sheets_core::worksheet::PageSetup;

        // Absent strings are written as XLNullableWideString null
        // markers (0xFFFFFFFF). The reader must skip over a null
        // marker and keep parsing: a footer-only setup means slot 1
        // (header) is null and slot 2 (footer) carries the data.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "hf").unwrap();
        let mut ps = PageSetup::default();
        ps.odd_footer = Some("&CPage &P".to_string());
        ps.first_footer = Some("&Cfirst".to_string());
        ps.different_first = true;
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let ps2 = ws2.page_setup();
        assert_eq!(ps2.odd_header, None);
        assert_eq!(ps2.odd_footer.as_deref(), Some("&CPage &P"));
        assert_eq!(ps2.first_footer.as_deref(), Some("&Cfirst"));
        assert!(ps2.different_first);
    }

    #[test]
    fn comments_roundtrip() {
        use duke_sheets_core::comment::CellComment;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "commented").unwrap();
        ws.set_comment_at(0, 0, CellComment::new("Author1", "First comment")).unwrap();
        ws.set_cell_value_at(1, 1, "also commented").unwrap();
        ws.set_comment_at(1, 1, CellComment::new("Author2", "Second comment")).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let c1 = ws2.comment_at(0, 0).expect("comment at A1");
        assert_eq!(c1.author, "Author1");
        assert_eq!(c1.plain_text(), "First comment");
        let c2 = ws2.comment_at(1, 1).expect("comment at B2");
        assert_eq!(c2.author, "Author2");
        assert_eq!(c2.plain_text(), "Second comment");
    }

    #[test]
    fn external_hyperlink_roundtrip() {
        use duke_sheets_core::Hyperlink;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "click me").unwrap();
        ws.set_hyperlink(
            "A1",
            Hyperlink {
                target: "https://example.com".to_string(),
                display: Some("Example".to_string()),
                tooltip: Some("Go to example".to_string()),
                location: None,
            },
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let link = ws2.hyperlink_at(0, 0).expect("hyperlink at A1");
        assert_eq!(link.target, "https://example.com");
        assert_eq!(link.display.as_deref(), Some("Example"));
        assert_eq!(link.tooltip.as_deref(), Some("Go to example"));
    }

    #[test]
    fn internal_hyperlink_roundtrip() {
        use duke_sheets_core::Hyperlink;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "go to B5").unwrap();
        ws.set_hyperlink(
            "A1",
            Hyperlink {
                target: "#Sheet1!B5".to_string(),
                display: None,
                tooltip: None,
                location: None,
            },
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let link = ws2.hyperlink_at(0, 0).expect("hyperlink at A1");
        assert_eq!(link.target, "#Sheet1!B5");
    }

    #[test]
    fn hidden_row_no_data_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "visible").unwrap();
        ws.set_row_hidden(5, true);
        ws.set_row_height(5, 25.0);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert!(ws2.is_row_hidden(5));
        assert!((ws2.row_height(5) - 25.0).abs() < 0.1);
    }

    #[test]
    fn data_validation_roundtrip() {
        use duke_sheets_core::validation::{DataValidation, ValidationOperator, ValidationType};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "validated").unwrap();
        let dv = DataValidation {
            validation_type: ValidationType::Whole {
                operator: ValidationOperator::GreaterThan,
                value1: "0".to_string(),
                value2: None,
            },
            ranges: vec![duke_sheets_core::CellRange::parse("A1:A10").unwrap()],
            allow_blank: true,
            show_dropdown: true,
            show_input_message: true,
            input_title: Some("Input".to_string()),
            input_message: Some("Enter a number".to_string()),
            show_error_alert: true,
            error_style: duke_sheets_core::validation::ValidationErrorStyle::Stop,
            error_title: Some("Error".to_string()),
            error_message: Some("Must be positive".to_string()),
        };
        ws.add_data_validation(dv);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let dvs = ws2.data_validations();
        assert_eq!(dvs.len(), 1);
        assert_eq!(dvs[0].ranges.len(), 1);
        assert_eq!(dvs[0].ranges[0].to_string(), "A1:A10");
        assert_eq!(dvs[0].error_title.as_deref(), Some("Error"));
        assert_eq!(dvs[0].error_message.as_deref(), Some("Must be positive"));
        assert_eq!(dvs[0].input_title.as_deref(), Some("Input"));
        assert_eq!(dvs[0].input_message.as_deref(), Some("Enter a number"));
        assert!(dvs[0].allow_blank);
        assert!(dvs[0].show_input_message);
        assert!(dvs[0].show_error_alert);
    }

    /// Locate the first BrtDVal record and split out (header, formula1
    /// rgce) for byte-level assertions.
    fn first_dval_header_and_formula1(wb: &Workbook) -> (u32, Vec<u8>) {
        let recs = sheet1_records(wb);
        let (_, payload) = recs
            .iter()
            .find(|(t, _)| *t == 0x0040)
            .expect("BrtDVal record present");
        let header = u32::from_le_bytes(payload[0..4].try_into().unwrap());
        let mut pos = 4;
        let range_count = u32::from_le_bytes(payload[pos..pos + 4].try_into().unwrap()) as usize;
        pos += 4 + range_count * 16;
        // 4 DValStrings (XLNullableWideString)
        for _ in 0..4 {
            let cch = u32::from_le_bytes(payload[pos..pos + 4].try_into().unwrap());
            pos += 4;
            if cch != 0xFFFFFFFF {
                pos += cch as usize * 2;
            }
        }
        let cce = u32::from_le_bytes(payload[pos..pos + 4].try_into().unwrap()) as usize;
        pos += 4;
        (header, payload[pos..pos + cce].to_vec())
    }

    fn list_dv(source: &str) -> Workbook {
        use duke_sheets_core::validation::DataValidation;
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut dv = DataValidation::list(source);
        dv.ranges = vec![duke_sheets_core::CellRange::parse("A1:A5").unwrap()];
        ws.add_data_validation(dv);
        wb
    }

    #[test]
    fn dv_list_literal_sets_str_lookup_bit() {
        // A literal list source is stored as a tStr formula with the
        // fStrLookup header bit (bit 7) set so Excel splits the string
        // into dropdown entries ([MS-XLSB] BrtDVal).
        let (header, rgce) = first_dval_header_and_formula1(&list_dv("Red,Green,Blue"));
        assert_ne!(
            header & (1 << 7),
            0,
            "fStrLookup bit must be set: {header:08X}"
        );
        assert_eq!(rgce.first(), Some(&0x17u8), "formula1 must be a tStr token");
    }

    #[test]
    fn dv_list_range_source_compiles_as_reference() {
        // A range source must compile to a reference token, not a
        // quoted string literal, and must not set fStrLookup.
        let (header, rgce) = first_dval_header_and_formula1(&list_dv("$A$1:$A$3"));
        assert_eq!(
            header & (1 << 7),
            0,
            "fStrLookup must be clear: {header:08X}"
        );
        assert_eq!(
            rgce.first().map(|b| b & 0x1F),
            Some(0x05),
            "formula1 must start with a PtgArea-class token; rgce={rgce:02X?}"
        );
    }

    #[test]
    fn dv_custom_array_constant_formula_is_dropped() {
        use duke_sheets_core::validation::{DataValidation, ValidationType};
        // DVParsedFormula MUST NOT contain PtgArray (or union /
        // intersection tokens); emit an empty formula instead of
        // spec-invalid bytes Excel would repair away.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let dv = DataValidation {
            validation_type: ValidationType::Custom {
                formula: "=SUM({1,2,3})>0".to_string(),
            },
            ranges: vec![duke_sheets_core::CellRange::parse("A1:A5").unwrap()],
            ..DataValidation::list("x")
        };
        ws.add_data_validation(dv);
        let (_, rgce) = first_dval_header_and_formula1(&wb);
        assert!(
            rgce.is_empty(),
            "array-constant DV formula must be dropped, not emitted: {rgce:02X?}"
        );
    }

    #[test]
    fn dv_list_quoted_values_roundtrip() {
        // Embedded quotes are stored doubled inside the tStr; the
        // reader must unescape them (and strip only the outer pair).
        let source = "Say \"Hi\",Plain";
        let wb2 = round_trip(&list_dv(source));
        let dvs = wb2.worksheet(0).unwrap().data_validations();
        assert_eq!(dvs.len(), 1);
        match &dvs[0].validation_type {
            duke_sheets_core::validation::ValidationType::List { source: s } => {
                assert_eq!(s, source);
            }
            other => panic!("expected List, got {other:?}"),
        }
    }

    #[test]
    fn conditional_format_roundtrip() {
        use duke_sheets_core::conditional_format::{CfRuleType, ConditionalFormatRule};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 100.0).unwrap();
        let rule = ConditionalFormatRule::new(CfRuleType::DuplicateValues)
            .with_range(duke_sheets_core::CellRange::parse("A1:A10").unwrap())
            .with_priority(1);
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let cfs = ws2.conditional_formats();
        assert_eq!(cfs.len(), 1);
        assert!(matches!(cfs[0].rule_type, CfRuleType::DuplicateValues));
        assert_eq!(cfs[0].ranges.len(), 1);
        assert_eq!(cfs[0].ranges[0].to_string(), "A1:A10");
    }

    #[test]
    fn combined_features_roundtrip() {
        use duke_sheets_core::comment::CellComment;
        use duke_sheets_core::worksheet::{PageOrientation, PageSetup};
        use duke_sheets_core::{AutoFilter, Hyperlink};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        ws.set_cell_value_at(0, 0, "Header").unwrap();
        ws.set_cell_value_at(0, 1, "Value").unwrap();
        ws.set_cell_value_at(1, 0, "A").unwrap();
        ws.set_cell_value_at(1, 1, 42.0).unwrap();

        ws.merge_cells(&duke_sheets_core::CellRange::parse("D1:E2").unwrap())
            .unwrap();
        ws.set_freeze_panes(1, 0);
        ws.set_row_height(0, 20.0);
        ws.set_column_width(0, 15.0);
        ws.set_auto_filter(Some(AutoFilter::new(
            duke_sheets_core::CellRange::parse("A1:B2").unwrap(),
        )));
        ws.set_comment_at(0, 0, CellComment::new("Test", "A note")).unwrap();
        ws.set_hyperlink(
            "B1",
            Hyperlink {
                target: "https://example.com".to_string(),
                display: None,
                tooltip: None,
                location: None,
            },
        )
        .unwrap();

        let mut ps = PageSetup::default();
        ps.orientation = PageOrientation::Landscape;
        ps.odd_header = Some("&CTest Header".to_string());
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();

        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("Header"));
        assert_eq!(ws2.get_value_at(1, 1), CellValue::Number(42.0));
        assert_eq!(ws2.merged_regions().len(), 1);
        assert_eq!(ws2.merged_regions()[0].to_string(), "D1:E2");
        let fp = ws2.freeze_panes().unwrap();
        assert_eq!(fp.row, 1);
        assert_eq!(fp.col, 0);
        assert!((ws2.row_height(0) - 20.0).abs() < 0.1);
        assert!((ws2.column_width(0) - 15.0).abs() < 0.1);
        assert!(ws2.auto_filter().is_some());
        assert!(ws2.comment_at(0, 0).is_some());
        assert!(ws2.hyperlink_at(0, 1).is_some());
        assert_eq!(ws2.page_setup().orientation, PageOrientation::Landscape);
        assert_eq!(
            ws2.page_setup().odd_header.as_deref(),
            Some("&CTest Header")
        );
    }

    #[test]
    fn formula_text_roundtrip_simple_arithmetic() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=1+2", CellValue::Number(3.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(3.0));
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=1+2");
    }

    #[test]
    fn formula_text_roundtrip_sum() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(1, 0, 20.0).unwrap();
        ws.set_formula_with_cached_value_at(2, 0, "=SUM(A1:A2)", CellValue::Number(30.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(2, 0), CellValue::Number(30.0));
        let fd = ws2.formula_data_at(2, 0).expect("formula should exist");
        assert_eq!(fd.text, "=SUM(A1:A2)");
    }

    #[test]
    fn formula_text_roundtrip_external_udf() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(
            0,
            0,
            r#"=[1]!TBLink("acct")"#,
            CellValue::Number(42.0),
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(42.0));
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert!(
            fd.text.contains("[1]!TBLink") && fd.text.contains("acct"),
            "external UDF formula lost: {:?}",
            fd.text
        );
    }

    #[test]
    fn formula_text_roundtrip_string_result() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=\"hello\"", CellValue::string("hello"))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("hello"));
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=\"hello\"");
    }

    #[test]
    fn formula_text_roundtrip_boolean_result() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=TRUE", CellValue::Boolean(true))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=TRUE");
    }

    #[test]
    fn formula_text_roundtrip_cross_sheet() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0).unwrap().set_name("Sheet1");
        wb.add_worksheet().unwrap();
        wb.worksheet_mut(1).unwrap().set_name("Sheet2");
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, 42.0)
            .unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_formula_with_cached_value_at(0, 0, "=Sheet2!A1", CellValue::Number(42.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(42.0));
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=Sheet2!A1");
    }

    #[test]
    fn formula_text_roundtrip_if_function() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 5.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 1, "=IF(A1>0,TRUE,FALSE)", CellValue::Boolean(true))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 1).expect("formula should exist");
        assert_eq!(fd.text, "=IF(A1>0,TRUE,FALSE)");
    }

    #[test]
    fn formula_text_roundtrip_cell_ref() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 1, "=A1", CellValue::Number(10.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 1).expect("formula should exist");
        assert_eq!(fd.text, "=A1");
    }

    #[test]
    fn formula_text_roundtrip_concat() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=\"a\"&\"b\"", CellValue::string("ab"))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=\"a\"&\"b\"");
    }

    #[test]
    fn formula_array_constant_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=SUM({1,2,3})", CellValue::Number(6.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=SUM({1,2,3})");
    }

    #[test]
    fn formula_array_constant_mixed_types_roundtrip() {
        // SerAr element encodings per [MS-XLSB] (cross-checked against
        // LO importArrayToken): string = type 1 + u16 cch + UTF-16;
        // bool = type 2 + 1 byte (no padding); error = type 4 + 1 byte
        // + 3 reserved bytes.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(
            0,
            0,
            "=COUNTA({\"ab\",\"cde\"})",
            CellValue::Number(2.0),
        )
        .unwrap();
        ws.set_formula_with_cached_value_at(0, 1, "=OR({TRUE,FALSE})", CellValue::Boolean(true))
            .unwrap();
        ws.set_formula_with_cached_value_at(0, 2, "=COUNT({1,#N/A,3})", CellValue::Number(2.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let texts: Vec<String> = (0..3)
            .map(|c| {
                ws2.formula_data_at(0, c)
                    .expect("formula should exist")
                    .text
                    .clone()
            })
            .collect();
        assert_eq!(texts[0], "=COUNTA({\"ab\",\"cde\"})");
        assert_eq!(texts[1], "=OR({TRUE,FALSE})");
        assert_eq!(texts[2], "=COUNT({1,#N/A,3})");
    }

    #[test]
    fn many_arg_function_roundtrip() {
        // BIFF12 PtgFuncVar cparams is a full unsigned byte (LO reads
        // it unmasked; BIFF8's fPrompt bit does not exist here), so a
        // 200-arg call must survive. Beyond 255 args the writer must
        // fall back to the cached value rather than emit a wrapped
        // argc byte.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let f200 = format!(
            "=SUM({})",
            (1..=200)
                .map(|i| i.to_string())
                .collect::<Vec<_>>()
                .join(",")
        );
        ws.set_formula_with_cached_value_at(0, 0, &f200, CellValue::Number(20100.0))
            .unwrap();
        let f300 = format!(
            "=SUM({})",
            (1..=300)
                .map(|i| i.to_string())
                .collect::<Vec<_>>()
                .join(",")
        );
        ws.set_formula_with_cached_value_at(0, 1, &f300, CellValue::Number(45150.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(
            ws2.formula_data_at(0, 0).expect("200-arg formula").text,
            f200,
            "200-arg PtgFuncVar must round-trip"
        );
        // 300 args cannot be represented: cached value only, never a
        // corrupt token stream.
        assert!(
            ws2.formula_data_at(0, 1).is_none(),
            "300-arg formula must fall back to cached value, got {:?}",
            ws2.formula_data_at(0, 1).map(|f| f.text.clone())
        );
        match ws2.get_value_at(0, 1).effective_value() {
            CellValue::Number(n) => assert!((n - 45150.0).abs() < 1e-9),
            other => panic!("cached value lost: {other:?}"),
        }
    }

    #[test]
    fn formula_uplus_paren_roundtrip() {
        // Redundant parens only survive if the compiler emits PtgParen;
        // the leading plus only survives via PtgUplus.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 2.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 1, "=+A1", CellValue::Number(2.0))
            .unwrap();
        ws.set_formula_with_cached_value_at(0, 2, "=(A1+1)", CellValue::Number(3.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let uplus = ws2.formula_data_at(0, 1).expect("formula should exist");
        assert_eq!(uplus.text, "=+A1");
        let paren = ws2.formula_data_at(0, 2).expect("formula should exist");
        assert_eq!(paren.text, "=(A1+1)");
    }

    fn read_zip_entry(data: &[u8], name: &str) -> String {
        let cursor = Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut file = archive.by_name(name).unwrap();
        let mut buf = String::new();
        std::io::Read::read_to_string(&mut file, &mut buf).unwrap();
        buf
    }

    fn zip_entry_exists(data: &[u8], name: &str) -> bool {
        let cursor = Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let exists = archive.by_name(name).is_ok();
        exists
    }

    fn write_xlsb_bytes(wb: &Workbook) -> Vec<u8> {
        let mut buf = Vec::new();
        XlsbWriter::write(wb, Cursor::new(&mut buf)).unwrap();
        buf
    }

    #[test]
    fn writer_emits_theme_xml() {
        let wb = Workbook::new();
        let bytes = write_xlsb_bytes(&wb);
        let theme = read_zip_entry(&bytes, "xl/theme/theme1.xml");
        assert!(theme.contains("<a:theme"));
        assert!(theme.contains("<a:clrScheme"));
    }

    #[test]
    fn writer_emits_theme_content_type_and_relationship() {
        let wb = Workbook::new();
        let bytes = write_xlsb_bytes(&wb);

        let ct = read_zip_entry(&bytes, "[Content_Types].xml");
        assert!(ct.contains("/xl/theme/theme1.xml"));

        let rels = read_zip_entry(&bytes, "xl/_rels/workbook.bin.rels");
        assert!(rels.contains("Target=\"theme/theme1.xml\""));
        assert!(rels.contains("/theme"));
    }

    #[test]
    fn round_trip_preserves_rgb_colors() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "test").unwrap();
        let mut style = Style::new().font_color(Color::Rgb {
            r: 79,
            g: 129,
            b: 189,
        });
        style.fill = FillStyle::Solid {
            color: Color::Rgb {
                r: 192,
                g: 80,
                b: 77,
            },
        };
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let s2 = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(
            s2.font.color,
            Color::Rgb {
                r: 79,
                g: 129,
                b: 189
            }
        );
        match s2.fill {
            FillStyle::Solid { color } => {
                assert_eq!(
                    color,
                    Color::Rgb {
                        r: 192,
                        g: 80,
                        b: 77
                    }
                );
            }
            ref other => panic!("expected Solid fill, got {:?}", other),
        }
    }

    #[test]
    fn rich_text_round_trip() {
        use duke_sheets_core::rich_text::{RichTextRun, RunFont};
        use duke_sheets_core::style::Color;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        let runs = vec![
            RichTextRun::plain("Hello "),
            RichTextRun::with_font(
                "World",
                RunFont {
                    bold: Some(true),
                    size: Some(14.0),
                    color: Some(Color::Rgb { r: 255, g: 0, b: 0 }),
                    name: Some("Arial".to_string()),
                    ..Default::default()
                },
            ),
        ];
        ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let val = ws2.get_value_at(0, 0);

        match val {
            CellValue::RichText(ref rt_runs) => {
                assert_eq!(rt_runs.len(), 2);
                assert_eq!(rt_runs[0].text, "Hello ");
                assert_eq!(rt_runs[1].text, "World");
                let font = rt_runs[1].font.as_ref().expect("should have font");
                assert_eq!(font.bold, Some(true));
                assert_eq!(font.size, Some(14.0));
                assert_eq!(font.name, Some("Arial".to_string()));
                assert!(matches!(
                    font.color,
                    Some(Color::Rgb { r: 255, g: 0, b: 0 })
                ));
            }
            other => panic!("Expected RichText, got {:?}", other),
        }
    }

    #[test]
    fn rich_text_with_plain_text_coexist() {
        use duke_sheets_core::rich_text::{RichTextRun, RunFont};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        ws.set_cell_value_at(0, 0, "plain text").unwrap();

        let runs = vec![
            RichTextRun::with_font(
                "Bold",
                RunFont {
                    bold: Some(true),
                    ..Default::default()
                },
            ),
            RichTextRun::plain(" normal"),
        ];
        ws.set_cell_value_at(0, 1, CellValue::rich_text(runs))
            .unwrap();

        ws.set_cell_value_at(0, 2, "another plain").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();

        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("plain text"));
        assert!(ws2.get_value_at(0, 1).is_rich_text());
        assert_eq!(ws2.get_value_at(0, 2), CellValue::string("another plain"));

        if let CellValue::RichText(runs) = ws2.get_value_at(0, 1) {
            assert_eq!(runs.len(), 2);
            assert_eq!(runs[0].text, "Bold");
            assert_eq!(runs[0].font.as_ref().unwrap().bold, Some(true));
            assert_eq!(runs[1].text, " normal");
        }
    }

    #[test]
    fn drawing_round_trip() {
        let anchor_xml = br#"<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id="7" name="Hello Connector"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>"#;

        let mut wb = Workbook::new();
        add_raw_drawing(wb.worksheet_mut(0).unwrap(), anchor_xml.to_vec());

        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

        let mut archive = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let entry_names: Vec<String> = (0..archive.len())
            .map(|i| archive.by_index(i).unwrap().name().to_string())
            .collect();
        assert!(
            entry_names.iter().any(|n| n == "xl/drawings/drawing1.xml"),
            "drawing XML missing from ZIP: {:?}",
            entry_names
        );
        drop(archive);

        let ct = {
            let mut a = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
            let mut f = a.by_name("[Content_Types].xml").unwrap();
            let mut s = String::new();
            std::io::Read::read_to_string(&mut f, &mut s).unwrap();
            s
        };
        assert!(
            ct.contains("drawing+xml"),
            "content types missing drawing override: {}",
            ct
        );

        let rels = {
            let mut a = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
            let mut f = a.by_name("xl/worksheets/_rels/sheet1.bin.rels").unwrap();
            let mut s = String::new();
            std::io::Read::read_to_string(&mut f, &mut s).unwrap();
            s
        };
        assert!(
            rels.contains("/drawing"),
            "sheet rels missing drawing relationship: {}",
            rels
        );

        let wb2 = XlsbReader::read(Cursor::new(&buf)).unwrap();
        let ws2 = wb2.worksheet(0).unwrap();
        let raw2 = raw_drawing_bytes(ws2);
        assert_eq!(raw2.len(), 1, "unmodeled anchor read back as raw");
        let text = std::str::from_utf8(raw2[0]).unwrap();
        assert!(text.contains("Hello"), "anchor content preserved: {text}");

        // The raw anchor survives a second write unchanged.
        let wb3 = round_trip(&wb2);
        let ws3 = wb3.worksheet(0).unwrap();
        assert_eq!(raw_drawing_bytes(ws3), raw_drawing_bytes(ws2));
    }

    #[test]
    fn drawing_with_chart_round_trip() {
        use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, DrawingAnchor};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        for (i, v) in [3.0, 1.0, 4.0].iter().enumerate() {
            ws.set_cell_value_at(i as u32, 0, *v).unwrap();
        }
        let mut chart = Chart::new(ChartType::ColumnClustered);
        chart.title = Some("Chart Title".to_string());
        chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
        ws.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

        let mut archive = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let names: Vec<String> = (0..archive.len())
            .map(|i| archive.by_index(i).unwrap().name().to_string())
            .collect();
        assert!(names.iter().any(|n| n == "xl/charts/chart1.xml"));
        assert!(names.iter().any(|n| n == "xl/drawings/drawing1.xml"));
        drop(archive);

        let ct = {
            let mut a = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
            let mut f = a.by_name("[Content_Types].xml").unwrap();
            let mut s = String::new();
            std::io::Read::read_to_string(&mut f, &mut s).unwrap();
            s
        };
        assert!(ct.contains("drawingml.chart+xml"));

        let wb2 = XlsbReader::read(Cursor::new(&buf)).unwrap();
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.chart_count(), 1);
        let chart2 = ws2.charts().next().unwrap().payload;
        assert_eq!(chart2.chart_type, ChartType::ColumnClustered);
        assert_eq!(chart2.title.as_deref(), Some("Chart Title"));
        assert_eq!(chart2.series.len(), 1);
    }

    #[test]
    fn no_drawing_produces_no_drawing_parts() {
        let wb = Workbook::new();
        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

        let mut archive = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let names: Vec<String> = (0..archive.len())
            .map(|i| archive.by_index(i).unwrap().name().to_string())
            .collect();
        assert!(
            !names.iter().any(|n| n.contains("drawing")),
            "unexpected drawing entry: {:?}",
            names
        );
    }

    #[test]
    fn anchor_xml_fallback_round_trip() {
        let anchor = b"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/><xdr:clientData/></xdr:twoCellAnchor>";

        let mut wb = Workbook::new();
        add_raw_drawing(wb.worksheet_mut(0).unwrap(), anchor.to_vec());

        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

        let drawing_xml = {
            let mut a = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
            let mut f = a.by_name("xl/drawings/drawing1.xml").unwrap();
            let mut s = String::new();
            std::io::Read::read_to_string(&mut f, &mut s).unwrap();
            s
        };
        assert!(drawing_xml.contains("<xdr:wsDr"));
        assert!(drawing_xml.contains("<xdr:twoCellAnchor>"));
    }

    #[test]
    fn rich_text_deduplication() {
        use duke_sheets_core::rich_text::{RichTextRun, RunFont};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        let make_runs = || {
            vec![RichTextRun::with_font(
                "Italic",
                RunFont {
                    italic: Some(true),
                    ..Default::default()
                },
            )]
        };

        ws.set_cell_value_at(0, 0, CellValue::rich_text(make_runs()))
            .unwrap();
        ws.set_cell_value_at(0, 1, CellValue::rich_text(make_runs()))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();

        assert!(ws2.get_value_at(0, 0).is_rich_text());
        assert!(ws2.get_value_at(0, 1).is_rich_text());

        if let CellValue::RichText(runs) = ws2.get_value_at(0, 0) {
            assert_eq!(runs[0].text, "Italic");
            assert_eq!(runs[0].font.as_ref().unwrap().italic, Some(true));
        }
    }

    #[test]
    fn diagnostic_style_mapping_multi_sheet() {
        use crate::writer::styles::build_style_table;
        use duke_sheets_core::style::{Color, Style};

        let mut wb = Workbook::new();

        // Sheet 0: several styled cells
        let ws0 = wb.worksheet_mut(0).unwrap();
        ws0.set_name("Styled");
        for i in 0..10 {
            ws0.set_cell_value_at(i, 0, format!("row{}", i)).unwrap();
            let style = Style::new()
                .bold(i % 2 == 0)
                .font_size(10.0 + i as f64)
                .font_color(Color::rgb(i as u8 * 25, 0, 0));
            ws0.set_cell_style_at(i, 0, &style).unwrap();
        }

        // Sheets 1-4: plain data
        for s in 1..5 {
            wb.add_worksheet().unwrap();
            let ws = wb.worksheet_mut(s).unwrap();
            ws.set_name(&format!("Sheet{}", s));
            ws.set_cell_value_at(0, 0, format!("sheet{}", s)).unwrap();
        }

        let (_table, mapping, _) = build_style_table(&wb, &[]);
        let xf_count = mapping.xf_count();
        let max_mapped = mapping.max_mapped_xf();

        eprintln!("XF count: {}", xf_count);
        eprintln!("Max mapped XF: {}", max_mapped);
        for i in 0..5 {
            if let Some(map) = mapping.sheet_map(i) {
                let max_val = map.values().copied().max().unwrap_or(0);
                eprintln!(
                    "Sheet {} map: {} entries, max XF = {}",
                    i,
                    map.len(),
                    max_val
                );
            }
        }

        assert!(
            max_mapped < xf_count,
            "OUT OF RANGE: max mapped XF {} >= xf_count {}",
            max_mapped,
            xf_count
        );

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.sheet_count(), 5);
        for i in 0..10 {
            let s = wb2.worksheet(0).unwrap().cell_style_at(i, 0);
            assert!(s.is_some(), "Missing style on sheet0 row {}", i);
        }
    }

    #[test]
    fn active_sheet_roundtrip() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Sheet2").unwrap();
        wb.add_worksheet_with_name("Sheet3").unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "s1")
            .unwrap();
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, "s2")
            .unwrap();
        wb.worksheet_mut(2)
            .unwrap()
            .set_cell_value_at(0, 0, "s3")
            .unwrap();
        wb.set_active_sheet(2).unwrap();

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.active_sheet(), 2);
    }

    #[test]
    fn active_sheet_zero_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.active_sheet(), 0);
    }

    #[test]
    fn sheet_visibility_roundtrip() {
        use duke_sheets_core::worksheet::SheetVisibility;

        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Hidden").unwrap();
        wb.add_worksheet_with_name("VeryHidden").unwrap();

        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "visible")
            .unwrap();
        wb.worksheet_mut(1)
            .unwrap()
            .set_visibility(SheetVisibility::Hidden);
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, "hidden")
            .unwrap();
        wb.worksheet_mut(2)
            .unwrap()
            .set_visibility(SheetVisibility::VeryHidden);
        wb.worksheet_mut(2)
            .unwrap()
            .set_cell_value_at(0, 0, "very hidden")
            .unwrap();

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.sheet_count(), 3);
        assert_eq!(
            wb2.worksheet(0).unwrap().visibility(),
            SheetVisibility::Visible
        );
        assert_eq!(
            wb2.worksheet(1).unwrap().visibility(),
            SheetVisibility::Hidden
        );
        assert_eq!(
            wb2.worksheet(2).unwrap().visibility(),
            SheetVisibility::VeryHidden
        );
        assert_eq!(
            wb2.worksheet(0).unwrap().get_value_at(0, 0),
            CellValue::string("visible")
        );
        assert_eq!(
            wb2.worksheet(1).unwrap().get_value_at(0, 0),
            CellValue::string("hidden")
        );
        assert_eq!(
            wb2.worksheet(2).unwrap().get_value_at(0, 0),
            CellValue::string("very hidden")
        );
    }

    #[test]
    fn table_roundtrip() {
        use duke_sheets_core::table::{Table, TableColumn, TableStyleInfo};
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "Name").unwrap();
        ws.set_cell_value_at(0, 1, "Value").unwrap();
        ws.set_cell_value_at(0, 2, "Score").unwrap();
        ws.set_cell_value_at(1, 0, "Alice").unwrap();
        ws.set_cell_value_at(1, 1, 100.0).unwrap();
        ws.set_cell_value_at(1, 2, 95.0).unwrap();
        ws.set_cell_value_at(2, 0, "Bob").unwrap();
        ws.set_cell_value_at(2, 1, 200.0).unwrap();
        ws.set_cell_value_at(2, 2, 88.0).unwrap();

        let table = Table {
            id: 1,
            name: "SalesTable".to_string(),
            display_name: "SalesTable".to_string(),
            reference: CellRange::parse("A1:C3").unwrap(),
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Name".to_string(),
                    totals_row_function: None,
                    totals_row_formula: None,
                    totals_row_label: None,
                    calculated_column_formula: None,
                },
                TableColumn {
                    id: 2,
                    name: "Value".to_string(),
                    totals_row_function: None,
                    totals_row_formula: None,
                    totals_row_label: None,
                    calculated_column_formula: None,
                },
                TableColumn {
                    id: 3,
                    name: "Score".to_string(),
                    totals_row_function: None,
                    totals_row_formula: None,
                    totals_row_label: None,
                    calculated_column_formula: None,
                },
            ],
            style_info: Some(TableStyleInfo {
                name: Some("TableStyleMedium2".to_string()),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: true,
        };
        ws.add_table(table);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let tables = ws2.tables();
        assert_eq!(tables.len(), 1, "should have 1 table");
        let t = &tables[0];
        assert_eq!(t.name, "SalesTable");
        assert_eq!(t.display_name, "SalesTable");
        assert_eq!(t.columns.len(), 3);
        assert_eq!(t.columns[0].name, "Name");
        assert_eq!(t.columns[1].name, "Value");
        assert_eq!(t.columns[2].name, "Score");
        assert!(t.style_info.is_some());
        let si = t.style_info.as_ref().unwrap();
        assert_eq!(si.name.as_deref(), Some("TableStyleMedium2"));
        assert!(si.show_row_stripes);
        assert!(!si.show_column_stripes);
    }

    #[test]
    fn named_range_roundtrip() {
        use duke_sheets_core::named_range::{NameScope, NamedRange};
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Data").unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "header")
            .unwrap();
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, 100.0)
            .unwrap();

        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("SalesData", "Data!$A$1:$D$10"));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::sheet_scope("LocalRate", "Sheet1!$B$1", 0));

        let wb2 = round_trip(&wb);
        let nr = wb2.named_ranges();
        assert!(nr.get("SalesData", 0).is_some(), "SalesData should exist");
        let sd = nr.get("SalesData", 0).unwrap();
        assert_eq!(sd.scope, NameScope::Workbook);
        assert!(
            sd.refers_to.contains("Data!"),
            "SalesData refers_to should reference Data sheet"
        );

        assert!(
            nr.get("LocalRate", 0).is_some(),
            "LocalRate should exist for sheet 0"
        );
        let lr = nr.get("LocalRate", 0).unwrap();
        assert!(matches!(lr.scope, NameScope::Sheet(0)));
    }

    #[test]
    fn tab_color_rgb_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_tab_color(Some(Color::rgb(255, 0, 0)));

        let wb2 = round_trip(&wb);
        assert_eq!(
            wb2.worksheet(0).unwrap().tab_color(),
            Some(Color::rgb(255, 0, 0)),
        );
    }

    #[test]
    fn tab_color_none_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0).unwrap().set_tab_color(None);

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.worksheet(0).unwrap().tab_color(), None);
    }

    #[test]
    fn cf_color_scale_roundtrip() {
        use duke_sheets_core::conditional_format::{
            CfColorValue, CfRuleType, CfValueType, ConditionalFormatRule,
        };
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(1, 0, 50.0).unwrap();
        ws.set_cell_value_at(2, 0, 90.0).unwrap();

        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::ColorScale {
                colors: vec![
                    CfColorValue::new(CfValueType::Min, None, Color::rgb(255, 0, 0)),
                    CfColorValue::new(CfValueType::Max, None, Color::rgb(0, 255, 0)),
                ],
            },
            ranges: vec![CellRange::parse("A1:A3").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1, "should have 1 CF rule");
        match &rules[0].rule_type {
            CfRuleType::ColorScale { colors } => {
                assert_eq!(colors.len(), 2);
                assert_eq!(colors[0].value_type, CfValueType::Min);
                assert_eq!(colors[1].value_type, CfValueType::Max);
            }
            other => panic!("expected ColorScale, got {:?}", other),
        }
    }

    #[test]
    fn cf_data_bar_roundtrip() {
        use duke_sheets_core::conditional_format::{
            CfRuleType, CfValue, CfValueType, ConditionalFormatRule,
        };
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 25.0).unwrap();
        ws.set_cell_value_at(1, 0, 75.0).unwrap();

        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::DataBar {
                min_value: CfValue::min(),
                max_value: CfValue::max(),
                color: Color::rgb(99, 142, 198),
                show_value: true,
                gradient: true,
                border_color: None,
                negative_color: None,
            },
            ranges: vec![CellRange::parse("A1:A2").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1);
        match &rules[0].rule_type {
            CfRuleType::DataBar {
                min_value,
                max_value,
                show_value,
                ..
            } => {
                assert_eq!(min_value.value_type, CfValueType::Min);
                assert_eq!(max_value.value_type, CfValueType::Max);
                assert!(*show_value);
            }
            other => panic!("expected DataBar, got {:?}", other),
        }
    }

    #[test]
    fn cf_icon_set_roundtrip() {
        use duke_sheets_core::conditional_format::{
            CfRuleType, CfValue, CfValueType, ConditionalFormatRule, IconSetStyle,
        };
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(1, 0, 50.0).unwrap();
        ws.set_cell_value_at(2, 0, 90.0).unwrap();

        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::IconSet {
                icon_style: IconSetStyle::Arrows3,
                values: vec![
                    CfValue::new(CfValueType::Percent, Some("0".to_string())),
                    CfValue::new(CfValueType::Percent, Some("33".to_string())),
                    CfValue::new(CfValueType::Percent, Some("67".to_string())),
                ],
                reverse: false,
                show_value: true,
            },
            ranges: vec![CellRange::parse("A1:A3").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1);
        match &rules[0].rule_type {
            CfRuleType::IconSet {
                icon_style,
                values,
                reverse,
                show_value,
            } => {
                assert_eq!(*icon_style, IconSetStyle::Arrows3);
                assert_eq!(values.len(), 3);
                assert!(!*reverse);
                assert!(*show_value);
            }
            other => panic!("expected IconSet, got {:?}", other),
        }
    }

    #[test]
    fn advanced_cf_records_match_biff12_layout() {
        use crate::biff12::{records, RecordIter};
        use duke_sheets_core::conditional_format::{ConditionalFormatRule, IconSetStyle};
        use std::io::Read;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut scale = ConditionalFormatRule::color_scale_3(
            Color::rgb(255, 0, 0),
            Color::rgb(255, 255, 0),
            Color::rgb(0, 255, 0),
        );
        scale.ranges = vec![CellRange::parse("A1:A5").unwrap()];
        ws.add_conditional_format(scale);
        let mut bar = ConditionalFormatRule::data_bar(Color::rgb(99, 142, 198));
        bar.ranges = vec![CellRange::parse("B1:B5").unwrap()];
        ws.add_conditional_format(bar);
        let mut icons = ConditionalFormatRule::icon_set(IconSetStyle::Arrows3);
        icons.ranges = vec![CellRange::parse("C1:C5").unwrap()];
        ws.add_conditional_format(icons);

        let mut xlsb = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut xlsb)).unwrap();
        let mut zip = zip::ZipArchive::new(Cursor::new(xlsb)).unwrap();
        let mut sheet = Vec::new();
        zip.by_name("xl/worksheets/sheet1.bin")
            .unwrap()
            .read_to_end(&mut sheet)
            .unwrap();
        let mut iter = RecordIter::new(Cursor::new(sheet));
        let mut buf = Vec::new();
        let mut cf_records = Vec::new();
        while let Ok((record_type, len)) = iter.next_record(&mut buf) {
            if matches!(record_type, 461..=471 | 564) {
                cf_records.push((record_type, buf[..len].to_vec()));
            }
        }

        // Pinned against Excel output and [MS-XLSB] 2.4.23, 2.4.43, 2.4.91, 2.4.334, and 2.4.337.
        let types: Vec<u16> = cf_records.iter().map(|record| record.0).collect();
        assert_eq!(
            types,
            vec![
                461, 463, 469, 471, 471, 471, 564, 564, 564, 470, 464, 462, 461, 463, 467, 471,
                471, 564, 468, 464, 462, 461, 463, 465, 471, 471, 471, 466, 464, 462,
            ]
        );

        let begin_formats: Vec<&Vec<u8>> = cf_records
            .iter()
            .filter(|record| record.0 == records::BRT_BEGIN_COND_FMT)
            .map(|record| &record.1)
            .collect();
        assert_eq!(begin_formats.len(), 3);
        assert!(begin_formats
            .iter()
            .all(|payload| &payload[..4] == 1u32.to_le_bytes()));
        assert_eq!(
            u32::from_le_bytes(begin_formats[0][20..24].try_into().unwrap()),
            0
        );
        assert_eq!(
            u32::from_le_bytes(begin_formats[1][20..24].try_into().unwrap()),
            1
        );
        assert_eq!(
            u32::from_le_bytes(begin_formats[2][20..24].try_into().unwrap()),
            2
        );

        let rules: Vec<&Vec<u8>> = cf_records
            .iter()
            .filter(|record| record.0 == records::BRT_BEGIN_CF_RULE)
            .map(|record| &record.1)
            .collect();
        assert_eq!((read_u32(rules[0], 0), read_u32(rules[0], 4)), (3, 2));
        assert_eq!((read_u32(rules[1], 0), read_u32(rules[1], 4)), (4, 3));
        assert_eq!((read_u32(rules[2], 0), read_u32(rules[2], 4)), (6, 4));
        assert_eq!(
            rules
                .iter()
                .map(|rule| read_u32(rule, 12))
                .collect::<Vec<_>>(),
            vec![1, 2, 3]
        );

        let data_bar = cf_records
            .iter()
            .find(|record| record.0 == records::BRT_BEGIN_DATA_BAR)
            .unwrap();
        assert_eq!(data_bar.1, [10, 90, 1]);
        let icon_set = cf_records
            .iter()
            .find(|record| record.0 == records::BRT_BEGIN_ICON_SET)
            .unwrap();
        assert_eq!(icon_set.1, [0, 0, 0, 0, 0, 0]);

        let cfvos: Vec<&Vec<u8>> = cf_records
            .iter()
            .filter(|record| record.0 == records::BRT_CFVO)
            .map(|record| &record.1)
            .collect();
        assert!(cfvos.iter().all(|payload| payload.len() == 24));
        assert_eq!(
            f64::from_le_bytes(cfvos[1][4..12].try_into().unwrap()),
            50.0
        );
        for payload in &cfvos[5..8] {
            assert_eq!(read_u32(payload, 12), 1);
            assert_eq!(read_u32(payload, 16), 1);
        }
    }

    fn read_u32(data: &[u8], offset: usize) -> u32 {
        u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
    }

    #[test]
    fn cf_with_dxf_style_roundtrip() {
        use duke_sheets_core::conditional_format::{CfOperator, CfRuleType, ConditionalFormatRule};
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 100.0).unwrap();

        let format = Style::new().fill_color(Color::rgb(0, 255, 0)).bold(true);
        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::CellIs {
                operator: CfOperator::GreaterThan,
                formula1: "50".to_string(),
                formula2: None,
            },
            ranges: vec![CellRange::parse("A1").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: Some(format),
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1);
        match &rules[0].rule_type {
            CfRuleType::CellIs {
                operator, formula1, ..
            } => {
                assert_eq!(*operator, CfOperator::GreaterThan);
                assert_eq!(formula1, "50");
            }
            other => panic!("expected CellIs, got {:?}", other),
        }
    }

    #[test]
    fn cf_three_color_scale_roundtrip() {
        use duke_sheets_core::conditional_format::{
            CfColorValue, CfRuleType, CfValueType, ConditionalFormatRule,
        };
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(1, 0, 50.0).unwrap();
        ws.set_cell_value_at(2, 0, 90.0).unwrap();

        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::ColorScale {
                colors: vec![
                    CfColorValue::new(CfValueType::Min, None, Color::rgb(255, 0, 0)),
                    CfColorValue::new(
                        CfValueType::Percentile,
                        Some("50".to_string()),
                        Color::rgb(255, 255, 0),
                    ),
                    CfColorValue::new(CfValueType::Max, None, Color::rgb(0, 255, 0)),
                ],
            },
            ranges: vec![CellRange::parse("A1:A3").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1);
        match &rules[0].rule_type {
            CfRuleType::ColorScale { colors } => {
                assert_eq!(colors.len(), 3);
                assert_eq!(colors[0].value_type, CfValueType::Min);
                assert_eq!(colors[1].value_type, CfValueType::Percentile);
                assert_eq!(colors[2].value_type, CfValueType::Max);
                assert_eq!(colors[0].color, Color::rgb(255, 0, 0));
                assert_eq!(colors[1].color, Color::rgb(255, 255, 0));
                assert_eq!(colors[2].color, Color::rgb(0, 255, 0));
            }
            other => panic!("expected ColorScale, got {:?}", other),
        }
    }

    #[test]
    fn zoom_scale_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "zoom").unwrap();
        ws.set_zoom_scale(Some(150));

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.zoom_scale(), Some(150));
    }

    #[test]
    fn row_outline_level_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "outline").unwrap();
        ws.set_row_outline_level(1, 1);
        ws.set_row_outline_level(2, 1);
        ws.set_row_outline_level(3, 1);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.row_outline_level(1), 1);
        assert_eq!(ws2.row_outline_level(2), 1);
        assert_eq!(ws2.row_outline_level(3), 1);
        assert_eq!(ws2.row_outline_level(0), 0);
    }

    #[test]
    fn column_outline_level_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "outline").unwrap();
        ws.set_column_outline_level(1, 1);
        ws.set_column_outline_level(2, 1);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.column_outline_level(1), 1);
        assert_eq!(ws2.column_outline_level(2), 1);
        assert_eq!(ws2.column_outline_level(0), 0);
    }

    #[test]
    fn active_cell_selection_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0).unwrap().set_selection_active_cell(5, 3);

        let wb2 = round_trip(&wb);
        assert_eq!(
            wb2.worksheet(0).unwrap().selection_active_cell(),
            Some((5, 3))
        );
    }

    #[test]
    fn cf_time_period_roundtrip() {
        use duke_sheets_core::conditional_format::{CfRuleType, ConditionalFormatRule, TimePeriod};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "date").unwrap();
        let rule = ConditionalFormatRule::new(CfRuleType::TimePeriod {
            period: TimePeriod::ThisMonth,
        })
        .with_range(duke_sheets_core::CellRange::parse("A1:A10").unwrap())
        .with_priority(1);
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let cfs = ws2.conditional_formats();
        assert_eq!(cfs.len(), 1);
        match &cfs[0].rule_type {
            CfRuleType::TimePeriod { period } => {
                assert_eq!(*period, TimePeriod::ThisMonth);
            }
            other => panic!("expected TimePeriod, got {:?}", other),
        }
        assert_eq!(cfs[0].ranges[0].to_string(), "A1:A10");
    }

    #[test]
    fn tab_selected_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "s1")
            .unwrap();
        wb.add_worksheet().unwrap();
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, "s2")
            .unwrap();
        wb.worksheet_mut(1).unwrap().set_selected(true);

        let wb2 = round_trip(&wb);
        assert!(wb2.worksheet(1).unwrap().is_selected());
    }

    #[test]
    fn split_panes_roundtrip() {
        use duke_sheets_core::worksheet::SplitPanes;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "split").unwrap();
        ws.set_split_panes(Some(SplitPanes {
            x_split: 2000.0,
            y_split: 3000.0,
            top_left: Some((5, 3)),
            active_pane: Some("bottomRight".to_string()),
        }));

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let sp = ws2.split_panes().expect("split panes should exist");
        assert!((sp.x_split - 2000.0).abs() < 0.1);
        assert!((sp.y_split - 3000.0).abs() < 0.1);
        assert_eq!(sp.top_left, Some((5, 3)));
        assert_eq!(sp.active_pane.as_deref(), Some("bottomRight"));
        assert!(ws2.freeze_panes().is_none());
    }

    #[test]
    fn sheet_protection_roundtrip() {
        use duke_sheets_core::worksheet::SheetProtection;
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_protection(Some(SheetProtection {
                protected: true,
                password_hash: Some(hash_legacy_protection_password("password")),
                select_locked_cells: true,
                select_unlocked_cells: true,
                format_cells: true,
                format_columns: true,
                format_rows: true,
                insert_columns: true,
                insert_rows: true,
                insert_hyperlinks: true,
                delete_columns: true,
                delete_rows: true,
                sort: true,
                auto_filter: true,
                pivot_tables: true,
            }));
        let wb2 = round_trip(&wb);
        let prot = wb2.worksheet(0).unwrap().protection().expect("protection");
        assert!(prot.protected);
        assert_eq!(
            prot.password_hash,
            Some(hash_legacy_protection_password("password"))
        );
        assert!(prot.select_locked_cells);
        assert!(prot.select_unlocked_cells);
        assert!(prot.format_cells);
        assert!(prot.format_columns);
        assert!(prot.format_rows);
        assert!(prot.insert_columns);
        assert!(prot.insert_rows);
        assert!(prot.insert_hyperlinks);
        assert!(prot.delete_columns);
        assert!(prot.delete_rows);
        assert!(prot.sort);
        assert!(prot.auto_filter);
        assert!(prot.pivot_tables);
    }

    #[test]
    fn sheet_protection_raw_password_hash_roundtrip() {
        use duke_sheets_core::worksheet::SheetProtection;
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_protection(Some(SheetProtection {
                protected: true,
                password_hash: Some(0xCAFE),
                ..Default::default()
            }));

        let wb2 = round_trip(&wb);
        let prot = wb2.worksheet(0).unwrap().protection().expect("protection");
        assert_eq!(prot.password_hash, Some(0xCAFE));
    }

    #[test]
    fn workbook_protection_roundtrip() {
        let mut wb = Workbook::new();
        wb.set_workbook_protection(Some(WorkbookProtection {
            structure: true,
            windows: true,
            password_hash: Some(hash_legacy_protection_password("book")),
        }));
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();

        let wb2 = round_trip(&wb);
        let protection = wb2.workbook_protection().expect("workbook protection");
        assert!(protection.structure);
        assert!(protection.windows);
        assert_eq!(
            protection.password_hash,
            Some(hash_legacy_protection_password("book"))
        );
    }

    #[test]
    fn protected_ranges_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "editable").unwrap();
        ws.set_protected_ranges(vec![
            ProtectedRange::new(
                "MainEdit",
                vec![
                    CellRange::parse("A1:B2").unwrap(),
                    CellRange::parse("D4:D5").unwrap(),
                ],
            )
            .with_password("range"),
            ProtectedRange {
                name: "RawHash".to_string(),
                ranges: vec![CellRange::parse("F1:F3").unwrap()],
                password_hash: Some(0xCAFE),
                security_descriptor: None,
            },
        ]);

        let wb2 = round_trip(&wb);
        let ranges = wb2.worksheet(0).unwrap().protected_ranges();
        assert_eq!(ranges.len(), 2);
        assert_eq!(ranges[0].name, "MainEdit");
        assert_eq!(ranges[0].ranges.len(), 2);
        assert_eq!(ranges[0].ranges[0].to_string(), "A1:B2");
        assert_eq!(ranges[0].ranges[1].to_string(), "D4:D5");
        assert_eq!(
            ranges[0].password_hash,
            Some(hash_legacy_protection_password("range"))
        );
        assert_eq!(ranges[1].name, "RawHash");
        assert_eq!(ranges[1].ranges[0].to_string(), "F1:F3");
        assert_eq!(ranges[1].password_hash, Some(0xCAFE));
    }

    #[test]
    fn page_breaks_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_row_break(5);
        let wb2 = round_trip(&wb);
        let breaks = wb2.worksheet(0).unwrap().row_breaks();
        assert_eq!(breaks.len(), 1);
        assert_eq!(breaks[0].id, 5);
    }

    #[test]
    fn print_area_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_print_area(duke_sheets_core::CellRange::parse("A1:D10").unwrap());
        let wb2 = round_trip(&wb);
        assert!(wb2.worksheet(0).unwrap().print_area().is_some());
    }

    #[test]
    fn dynamic_filter_in_process_roundtrip() {
        use duke_sheets_core::auto_filter::{
            AutoFilter, ColumnFilter, DynamicFilter, DynamicFilterType, FilterColumn,
        };
        use duke_sheets_core::{CellAddress, CellRange};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut af = AutoFilter::new(CellRange::new(
            CellAddress::parse("A1").unwrap(),
            CellAddress::parse("A5").unwrap(),
        ));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Dynamic(DynamicFilter {
                filter_type: DynamicFilterType::Today,
                val: Some(45000.5),
                max_val: Some(45001.0),
            }),
        ));
        ws.set_auto_filter(Some(af));

        let wb2 = round_trip(&wb);
        let af = wb2.worksheet(0).unwrap().auto_filter().unwrap().clone();
        let col0 = af.filter_columns.iter().find(|fc| fc.col_id == 0).unwrap();
        match &col0.filter {
            ColumnFilter::Dynamic(d) => {
                assert_eq!(d.filter_type, DynamicFilterType::Today);
                assert_eq!(d.val, Some(45000.5));
                assert_eq!(d.max_val, Some(45001.0));
            }
            other => panic!("expected Dynamic filter, got {other:?}"),
        }
    }

    #[test]
    fn color_filter_in_process_roundtrip() {
        use duke_sheets_core::auto_filter::{AutoFilter, ColorFilter, ColumnFilter, FilterColumn};
        use duke_sheets_core::{CellAddress, CellRange};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut af = AutoFilter::new(CellRange::new(
            CellAddress::parse("A1").unwrap(),
            CellAddress::parse("A3").unwrap(),
        ));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Color(ColorFilter {
                dxf_id: Some(2),
                cell_color: false,
            }),
        ));
        ws.set_auto_filter(Some(af));

        let wb2 = round_trip(&wb);
        let af = wb2.worksheet(0).unwrap().auto_filter().unwrap().clone();
        let col0 = af.filter_columns.iter().find(|fc| fc.col_id == 0).unwrap();
        match &col0.filter {
            ColumnFilter::Color(c) => {
                // The dxfid is writer-assigned (it must reference the
                // synthesized DXF entry in styles.bin), so the model's
                // input index is not preserved — only its presence.
                assert!(c.dxf_id.is_some(), "dxf_id lost");
                assert!(!c.cell_color, "cell_color flag drifted");
            }
            other => panic!("expected Color filter, got {other:?}"),
        }
    }

    /// Raw (record_type, payload) pairs of `xl/worksheets/sheet1.bin`.
    fn sheet1_records(wb: &Workbook) -> Vec<(u16, Vec<u8>)> {
        let bytes = write_xlsb_bytes(wb);
        let bin = read_zip_entry_bytes(&bytes, "xl/worksheets/sheet1.bin");
        let mut iter = crate::biff12::RecordIter::new(Cursor::new(bin));
        let mut out = Vec::new();
        let mut buf = Vec::new();
        while let Ok((typ, len)) = iter.next_record(&mut buf) {
            out.push((typ, buf[..len].to_vec()));
            buf.truncate(len);
        }
        out
    }

    fn read_zip_entry_bytes(data: &[u8], name: &str) -> Vec<u8> {
        let cursor = Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut file = archive.by_name(name).unwrap();
        let mut buf = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut buf).unwrap();
        buf
    }

    fn record_types(data: Vec<u8>) -> Vec<u16> {
        let mut iter = crate::biff12::RecordIter::new(Cursor::new(data));
        let mut out = Vec::new();
        let mut buf = Vec::new();
        while let Ok((typ, len)) = iter.next_record(&mut buf) {
            out.push(typ);
            buf.truncate(len);
        }
        out
    }

    fn records_with_payload(data: Vec<u8>) -> Vec<(u16, Vec<u8>)> {
        let mut iter = crate::biff12::RecordIter::new(Cursor::new(data));
        let mut out = Vec::new();
        let mut buf = Vec::new();
        while let Ok((typ, len)) = iter.next_record(&mut buf) {
            out.push((typ, buf[..len].to_vec()));
            buf.truncate(len);
        }
        out
    }

    fn metadata_light_first_numeric_cache_field(data: Vec<u8>) -> Vec<u8> {
        let records = records_with_payload(data);
        let mut out = Vec::new();
        let mut rw = crate::biff12::RecordWriter::new(&mut out);
        let mut field_index = None;
        let mut in_first_shared_items = false;

        for (record_type, mut payload) in records {
            match record_type {
                crate::biff12::records::BRT_BEGIN_PIVOT_CACHE_DEF if payload.len() >= 21 => {
                    payload[17..21].copy_from_slice(&0u32.to_le_bytes());
                }
                crate::biff12::records::BRT_BEGIN_PCD_FIELD => {
                    field_index = Some(field_index.map_or(0usize, |index| index + 1));
                }
                crate::biff12::records::BRT_BEGIN_PCD_SHARED_ITEMS if field_index == Some(0) => {
                    if payload.len() >= 6 {
                        payload[2..6].copy_from_slice(&0u32.to_le_bytes());
                    }
                    rw.write_record(record_type, &payload).unwrap();
                    in_first_shared_items = true;
                    continue;
                }
                crate::biff12::records::BRT_END_PCD_SHARED_ITEMS if in_first_shared_items => {
                    rw.write_record(record_type, &payload).unwrap();
                    in_first_shared_items = false;
                    continue;
                }
                _ => {}
            }

            if in_first_shared_items {
                continue;
            }
            rw.write_record(record_type, &payload).unwrap();
        }

        drop(rw);
        out
    }

    fn inline_first_numeric_cache_record_field(data: Vec<u8>) -> Vec<u8> {
        let records = records_with_payload(data);
        let mut out = Vec::new();
        let mut rw = crate::biff12::RecordWriter::new(&mut out);

        for (record_type, payload) in records {
            if record_type == crate::biff12::records::BRT_PCD_RECORD && payload.len() >= 4 {
                let item_index = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                let value = match item_index {
                    0 => 1.0f64,
                    1 => 2.0f64,
                    _ => 0.0f64,
                };
                let mut patched = value.to_le_bytes().to_vec();
                patched.extend_from_slice(&payload[4..]);
                rw.write_record(record_type, &patched).unwrap();
            } else {
                rw.write_record(record_type, &payload).unwrap();
            }
        }

        drop(rw);
        out
    }

    fn patch_pcd_source_header(data: Vec<u8>, source_type: u32, connection_id: u32) -> Vec<u8> {
        let records = records_with_payload(data);
        let mut out = Vec::new();
        let mut rw = crate::biff12::RecordWriter::new(&mut out);

        for (record_type, payload) in records {
            if record_type == crate::biff12::records::BRT_BEGIN_PCD_SOURCE {
                let mut patched = Vec::new();
                patched.extend_from_slice(&source_type.to_le_bytes());
                patched.extend_from_slice(&connection_id.to_le_bytes());
                rw.write_record(record_type, &patched).unwrap();
            } else {
                rw.write_record(record_type, &payload).unwrap();
            }
        }

        drop(rw);
        out
    }

    fn add_connections_content_type(mut content_types: String) -> Vec<u8> {
        let override_xml = "<Override PartName=\"/xl/connections.bin\" ContentType=\"application/vnd.ms-excel.connections\"/>";
        content_types = content_types.replace("</Types>", &format!("{override_xml}</Types>"));
        content_types.into_bytes()
    }

    fn add_connections_workbook_rel(mut rels: String) -> Vec<u8> {
        let rel = "<Relationship Id=\"rIdConnections\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections\" Target=\"connections.bin\"/>";
        rels = rels.replace("</Relationships>", &format!("{rel}</Relationships>"));
        rels.into_bytes()
    }

    fn xlsb_database_connections_bin() -> Vec<u8> {
        let mut out = Vec::new();
        let mut rw = crate::biff12::RecordWriter::new(&mut out);
        rw.write_record(
            crate::biff12::records::BRT_BEGIN_EXT_CONNECTIONS,
            &0u32.to_le_bytes(),
        )
        .unwrap();
        rw.write_record(
            crate::biff12::records::BRT_BEGIN_EXT_CONNECTION,
            &xlsb_ext_connection_payload(),
        )
        .unwrap();
        rw.write_record(
            crate::biff12::records::BRT_BEGIN_EC_DB_PROPS,
            &xlsb_db_props_payload(),
        )
        .unwrap();
        out
    }

    fn xlsb_ext_connection_payload() -> Vec<u8> {
        let mut payload = vec![3, 0, 2, 0];
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&(0x0001u16 | 0x0040).to_le_bytes());
        payload.extend_from_slice(&0x0008u16.to_le_bytes());
        payload.extend_from_slice(&5u32.to_le_bytes());
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&7u32.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&crate::biff12::encode_wide_str("SalesConnection"));
        payload
    }

    fn xlsb_db_props_payload() -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&2u32.to_le_bytes());
        payload.push(0x02);
        payload.extend_from_slice(&crate::biff12::encode_wide_str(
            "Provider=MSDASQL;DSN=Sales;",
        ));
        payload.extend_from_slice(&crate::biff12::encode_wide_str("select * from Sales"));
        payload
    }

    fn replace_zip_entries(data: &[u8], replacements: &[(&str, Vec<u8>)]) -> Vec<u8> {
        let cursor = Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut out = Vec::new();
        let mut zip = zip::ZipWriter::new(Cursor::new(&mut out));
        let opts = zip::write::SimpleFileOptions::default();
        let mut seen = vec![false; replacements.len()];

        for index in 0..archive.len() {
            let mut file = archive.by_index(index).unwrap();
            let name = file.name().to_string();
            if file.is_dir() {
                zip.add_directory(name, opts).unwrap();
                continue;
            }

            let replacement =
                replacements
                    .iter()
                    .enumerate()
                    .find_map(|(replacement_index, (path, bytes))| {
                        (*path == name).then(|| {
                            seen[replacement_index] = true;
                            bytes.as_slice()
                        })
                    });
            let mut bytes = Vec::new();
            if let Some(replacement) = replacement {
                bytes.extend_from_slice(replacement);
            } else {
                file.read_to_end(&mut bytes).unwrap();
            }
            zip.start_file(name, opts).unwrap();
            zip.write_all(&bytes).unwrap();
        }
        for (replacement_index, (path, bytes)) in replacements.iter().enumerate() {
            if !seen[replacement_index] {
                zip.start_file(*path, opts).unwrap();
                zip.write_all(bytes).unwrap();
            }
        }

        zip.finish().unwrap();
        out
    }

    fn sxvd_item_list_counts(records: &[(u16, Vec<u8>)]) -> Vec<usize> {
        let mut counts = Vec::new();
        let mut current = None;
        for (record_type, _) in records {
            match *record_type {
                crate::biff12::records::BRT_BEGIN_SXVD => current = Some(0usize),
                crate::biff12::records::BRT_BEGIN_SXVIS => {
                    if let Some(count) = current.as_mut() {
                        *count += 1;
                    }
                }
                crate::biff12::records::BRT_END_SXVD => {
                    if let Some(count) = current.take() {
                        counts.push(count);
                    }
                }
                _ => {}
            }
        }
        counts
    }

    fn sxvi_item_types_for_sxvd(records: &[(u16, Vec<u8>)], sxvd_index: usize) -> Vec<u8> {
        let mut current_sxvd = None;
        let mut next_sxvd = 0usize;
        let mut item_types = Vec::new();
        for (record_type, payload) in records {
            match *record_type {
                crate::biff12::records::BRT_BEGIN_SXVD => {
                    current_sxvd = Some(next_sxvd);
                    next_sxvd += 1;
                }
                crate::biff12::records::BRT_BEGIN_SXVI if current_sxvd == Some(sxvd_index) => {
                    item_types.push(payload[0]);
                }
                crate::biff12::records::BRT_END_SXVD => current_sxvd = None,
                _ => {}
            }
        }
        item_types
    }

    fn sxvi_payloads_for_sxvd(records: &[(u16, Vec<u8>)], sxvd_index: usize) -> Vec<Vec<u8>> {
        let mut current_sxvd = None;
        let mut next_sxvd = 0usize;
        let mut payloads = Vec::new();
        for (record_type, payload) in records {
            match *record_type {
                crate::biff12::records::BRT_BEGIN_SXVD => {
                    current_sxvd = Some(next_sxvd);
                    next_sxvd += 1;
                }
                crate::biff12::records::BRT_BEGIN_SXVI if current_sxvd == Some(sxvd_index) => {
                    payloads.push(payload.clone());
                }
                crate::biff12::records::BRT_END_SXVD => current_sxvd = None,
                _ => {}
            }
        }
        payloads
    }

    fn sxli_item_values_for_axis(
        records: &[(u16, Vec<u8>)],
        begin_record: u16,
        end_record: u16,
    ) -> Vec<Vec<u32>> {
        let mut in_axis = false;
        let mut current_line: Option<Vec<u32>> = None;
        let mut lines = Vec::new();
        for (record_type, payload) in records {
            match *record_type {
                record if record == begin_record => in_axis = true,
                record if record == end_record => in_axis = false,
                crate::biff12::records::BRT_BEGIN_SXLI if in_axis => {
                    let item_type = u16::from_le_bytes(payload[2..4].try_into().unwrap());
                    current_line = (item_type != 13).then(Vec::new);
                }
                crate::biff12::records::BRT_SXLI_ITEM if in_axis => {
                    if let Some(line) = &mut current_line {
                        line.push(u32::from_le_bytes(payload[0..4].try_into().unwrap()));
                    }
                }
                crate::biff12::records::BRT_END_SXLI if in_axis => {
                    if let Some(line) = current_line.take() {
                        lines.push(line);
                    }
                }
                _ => {}
            }
        }
        lines
    }

    fn wide_string(data: &[u8]) -> String {
        let count = u32::from_le_bytes(data[0..4].try_into().unwrap()) as usize;
        let end = 4 + count * 2;
        let units = data[4..end]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16_lossy(&units)
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

    #[test]
    fn color_filter_emits_brt_color_filter_record_id() {
        use duke_sheets_core::auto_filter::{AutoFilter, ColorFilter, ColumnFilter, FilterColumn};
        use duke_sheets_core::{CellAddress, CellRange};

        // [MS-XLSB] record enumeration: 168 (0x00A8) = BrtColorFilter
        // (§2.4.339), 169 (0x00A9) = BrtIconFilter. A struct-level
        // round-trip cannot catch a swapped ID (both payloads are two
        // u32s), so pin the raw record id and payload here.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut af = AutoFilter::new(CellRange::new(
            CellAddress::parse("A1").unwrap(),
            CellAddress::parse("A3").unwrap(),
        ));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Color(ColorFilter {
                dxf_id: Some(2),
                cell_color: false,
            }),
        ));
        ws.set_auto_filter(Some(af));

        let recs = sheet1_records(&wb);
        let color: Vec<_> = recs.iter().filter(|(t, _)| *t == 0x00A8).collect();
        assert_eq!(
            color.len(),
            1,
            "expected exactly one BrtColorFilter (0x00A8) record"
        );
        // dxfid 0 = the writer-synthesized DXF entry; fCellColor = 0.
        let mut expected = 0u32.to_le_bytes().to_vec();
        expected.extend_from_slice(&0u32.to_le_bytes());
        assert_eq!(color[0].1, expected, "dxfid + fCellColor payload");
        assert!(
            !recs.iter().any(|(t, _)| *t == 0x00A9),
            "0x00A9 is BrtIconFilter; a color filter must not emit it"
        );

        // The dxfid must not dangle: styles.bin must carry a BrtDXF
        // (0x01FB) entry for it, or Excel refuses to open the file.
        let bytes = write_xlsb_bytes(&wb);
        let styles = read_zip_entry_bytes(&bytes, "xl/styles.bin");
        let mut iter = crate::biff12::RecordIter::new(Cursor::new(styles));
        let mut buf = Vec::new();
        let mut dxf_count = 0;
        while let Ok((typ, _len)) = iter.next_record(&mut buf) {
            if typ == 0x01FB {
                dxf_count += 1;
            }
        }
        assert!(
            dxf_count >= 1,
            "styles.bin must contain the DXF entry backing the color filter"
        );
    }

    fn custom_filter_workbook(and: bool) -> Workbook {
        use duke_sheets_core::auto_filter::{
            AutoFilter, ColumnFilter, CustomFilterCondition, CustomFilters, FilterColumn,
            FilterOperator,
        };
        use duke_sheets_core::{CellAddress, CellRange};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut af = AutoFilter::new(CellRange::new(
            CellAddress::parse("A1").unwrap(),
            CellAddress::parse("A5").unwrap(),
        ));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Custom(CustomFilters {
                and,
                conditions: vec![
                    CustomFilterCondition {
                        operator: FilterOperator::GreaterThan,
                        value: "5".to_string(),
                    },
                    CustomFilterCondition {
                        operator: FilterOperator::LessThan,
                        value: "10".to_string(),
                    },
                ],
            }),
        ));
        ws.set_auto_filter(Some(af));
        wb
    }

    #[test]
    fn custom_filter_emits_spec_wire_layout() {
        // BrtCustomFilter per [MS-XLSB] §2.4.348 (cross-checked against
        // LibreOffice's FilterCriterionModel::readBiffData): vts u8,
        // operator u8, xNumOrError 8 bytes, then the string (32-bit cch)
        // when vts = 6 (string). BrtBeginCustomFilters carries an i32
        // where 0 = AND, nonzero = OR (CustomFilter::importRecord).
        let recs = sheet1_records(&custom_filter_workbook(false));

        let begin: Vec<_> = recs.iter().filter(|(t, _)| *t == 0x00AC).collect();
        assert_eq!(begin.len(), 1, "one BrtBeginCustomFilters record");
        assert_eq!(
            begin[0].1,
            1i32.to_le_bytes(),
            "OR custom filters must carry i32 1 in BrtBeginCustomFilters"
        );

        let customs: Vec<_> = recs.iter().filter(|(t, _)| *t == 0x00AE).collect();
        assert_eq!(customs.len(), 2, "two BrtCustomFilter records");

        // vts=6 (string), operator=4 (greaterThan), 8 zero value bytes,
        // cch=1, UTF-16LE "5".
        let mut first = vec![6u8, 4u8];
        first.extend_from_slice(&[0u8; 8]);
        first.extend_from_slice(&1u32.to_le_bytes());
        first.extend_from_slice(&[0x35, 0x00]);
        assert_eq!(customs[0].1, first, "first BrtCustomFilter payload");

        // vts=6, operator=1 (lessThan), zeros, cch=2, UTF-16LE "10".
        let mut second = vec![6u8, 1u8];
        second.extend_from_slice(&[0u8; 8]);
        second.extend_from_slice(&2u32.to_le_bytes());
        second.extend_from_slice(&[0x31, 0x00, 0x30, 0x00]);
        assert_eq!(customs[1].1, second, "second BrtCustomFilter payload");

        // AND variant: begin record payload is 0.
        let recs_and = sheet1_records(&custom_filter_workbook(true));
        let begin_and: Vec<_> = recs_and.iter().filter(|(t, _)| *t == 0x00AC).collect();
        assert_eq!(begin_and[0].1, 0i32.to_le_bytes());
    }

    #[test]
    fn top10_without_computed_value_roundtrips_none() {
        use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, Top10Filter};
        use duke_sheets_core::{CellAddress, CellRange};

        // fApplied (flags bit 2) asserts xNumFilter is a real value
        // from the range; without a computed filter value it must stay
        // clear, and the absence must survive the round trip instead
        // of materializing as Some(val).
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut af = AutoFilter::new(CellRange::new(
            CellAddress::parse("A1").unwrap(),
            CellAddress::parse("A5").unwrap(),
        ));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Top10(Top10Filter {
                top: true,
                percent: false,
                val: 3.0,
                filter_val: None,
            }),
        ));
        ws.set_auto_filter(Some(af));

        let wb2 = round_trip(&wb);
        let af2 = wb2.worksheet(0).unwrap().auto_filter().unwrap().clone();
        match &af2.filter_columns[0].filter {
            ColumnFilter::Top10(t) => {
                assert!(t.top);
                assert_eq!(t.val, 3.0);
                assert_eq!(t.filter_val, None, "fApplied must not be claimed");
            }
            other => panic!("expected Top10, got {other:?}"),
        }
    }

    #[test]
    fn custom_filter_in_process_roundtrip() {
        use duke_sheets_core::auto_filter::{ColumnFilter, FilterOperator};

        for and in [true, false] {
            let wb2 = round_trip(&custom_filter_workbook(and));
            let af = wb2.worksheet(0).unwrap().auto_filter().unwrap().clone();
            let col0 = af.filter_columns.iter().find(|fc| fc.col_id == 0).unwrap();
            match &col0.filter {
                ColumnFilter::Custom(cf) => {
                    assert_eq!(cf.and, and, "AND/OR flag drifted");
                    assert_eq!(cf.conditions.len(), 2);
                    assert_eq!(cf.conditions[0].operator, FilterOperator::GreaterThan);
                    assert_eq!(cf.conditions[0].value, "5");
                    assert_eq!(cf.conditions[1].operator, FilterOperator::LessThan);
                    assert_eq!(cf.conditions[1].value, "10");
                }
                other => panic!("expected Custom filter, got {other:?}"),
            }
        }
    }

    #[test]
    fn union_parens_survive_outside_sum() {
        // Union parens are semantic: =COUNT((A1,B1)) has one argument.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 1.0).unwrap();
        ws.set_cell_value_at(0, 1, 2.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 3, "=COUNT((A1,B1))", CellValue::Number(2.0))
            .unwrap();
        ws.set_formula_with_cached_value_at(1, 3, "=(A1,B1)", CellValue::Number(1.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(
            ws2.formula_data_at(0, 3)
                .expect("formula should exist")
                .text,
            "=COUNT((A1,B1))"
        );
        assert_eq!(
            ws2.formula_data_at(1, 3)
                .expect("formula should exist")
                .text,
            "=(A1,B1)"
        );
    }

    #[test]
    fn ptg_name_index_accounts_for_xlfn_names() {
        use duke_sheets_core::named_range::{NameScope, NamedRange};

        // PtgName's ilbl indexes the whole BrtName record stream, and
        // _xlfn.* records are written before user names. A formula
        // using a post-2007 function (here IFS, absent from the Ftab)
        // plus a named-range reference must not bind the name to the
        // _xlfn record.
        let mut wb = Workbook::new();
        wb.named_ranges_mut().define_or_update(NamedRange::new(
            "MyRange",
            "Sheet1!$A$1:$A$3",
            NameScope::Workbook,
        ));
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 1.0).unwrap();
        ws.set_cell_value_at(1, 0, 2.0).unwrap();
        ws.set_cell_value_at(2, 0, 3.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 1, "=IFS(A1>0,1)", CellValue::Number(1.0))
            .unwrap();
        ws.set_formula_with_cached_value_at(0, 2, "=SUM(MyRange)", CellValue::Number(6.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let name_ref = ws2.formula_data_at(0, 2).expect("formula should exist");
        assert_eq!(name_ref.text, "=SUM(MyRange)");
        let xlfn = ws2.formula_data_at(0, 1).expect("formula should exist");
        assert_eq!(xlfn.text, "=IFS(A1>0,1)");
    }

    #[test]
    fn brt_name_ends_after_comment_when_not_a_macro() {
        use duke_sheets_core::named_range::{NameScope, NamedRange};

        // BrtName's four trailing strings (unusedstring1, description,
        // helpTopic, unusedstring2) MUST exist if and only if fProc is
        // set ([MS-XLSB] §2.4.718). We never write macro names, so the
        // record must end right after the comment string.
        let mut wb = Workbook::new();
        wb.named_ranges_mut().define_or_update(
            NamedRange::new("MyTax", "0.07", NameScope::Workbook).with_comment("rate"),
        );
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=MyTax*2", CellValue::Number(0.14))
            .unwrap();

        let bytes = write_xlsb_bytes(&wb);
        let bin = read_zip_entry_bytes(&bytes, "xl/workbook.bin");
        let mut iter = crate::biff12::RecordIter::new(Cursor::new(bin));
        let mut buf = Vec::new();
        let mut checked = 0;
        while let Ok((typ, len)) = iter.next_record(&mut buf) {
            if typ != 0x0027 {
                continue; // BrtName
            }
            let p = &buf[..len];
            let mut pos = 4 + 1 + 4; // flags + chKey + itab
            let read_str = |pos: &mut usize| {
                let cch = u32::from_le_bytes(p[*pos..*pos + 4].try_into().unwrap());
                *pos += 4;
                if cch != 0xFFFFFFFF {
                    *pos += cch as usize * 2;
                }
            };
            read_str(&mut pos); // name
            let cce = u32::from_le_bytes(p[pos..pos + 4].try_into().unwrap()) as usize;
            pos += 4 + cce;
            let cb = u32::from_le_bytes(p[pos..pos + 4].try_into().unwrap()) as usize;
            pos += 4 + cb;
            read_str(&mut pos); // comment
            assert_eq!(
                pos,
                len,
                "BrtName must end after the comment; {} trailing bytes remain",
                len - pos
            );
            checked += 1;
        }
        assert!(checked >= 1, "no BrtName records found");
    }

    #[test]
    fn named_range_comment_roundtrip() {
        use duke_sheets_core::named_range::{NameScope, NamedRange};

        let mut wb = Workbook::new();
        let nr =
            NamedRange::new("MyTax", "0.07", NameScope::Workbook).with_comment("Sales tax rate");
        wb.named_ranges_mut().define_or_update(nr);

        let wb2 = round_trip(&wb);
        let got = wb2
            .named_ranges()
            .iter()
            .find(|n| n.name == "MyTax")
            .expect("named range lost");
        assert_eq!(
            got.comment.as_deref(),
            Some("Sales tax rate"),
            "named range comment lost: {:?}",
            got.comment
        );
    }

    #[test]
    fn form_controls_roundtrip() {
        use duke_sheets_chart::{CellMarker, DrawingAnchor};
        use duke_sheets_core::{CheckState, FormControl, FormControlKind, ListSelection};

        let anchor = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
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
                first_in_group: false,
                no_3d: true,
            },
            FormControlKind::Label {
                caption: "Status".into(),
            },
            FormControlKind::GroupBox {
                caption: "Choices".into(),
                no_3d: true,
            },
            FormControlKind::ListBox {
                input_range: Some("$H$1:$H$5".to_string()),
                cell_link: None,
                selection: ListSelection::Multi,
                selected: vec![0, 2, 4],
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
        ];

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 42.0).unwrap();
        let count = kinds.len();
        for (i, kind) in kinds.iter().enumerate() {
            let row = 1 + 2 * i as u32;
            ws.add_form_control(FormControl::new(kind.clone()), anchor(1, row, 3, row + 1)).unwrap();
        }

        let wb2 = round_trip(&wb);
        let controls: Vec<_> = wb2.worksheet(0).unwrap().form_controls().collect();
        assert_eq!(controls.len(), count, "every control survives");
        for (i, drawn) in controls.iter().enumerate() {
            // The writer recomputes radio grouping; the single radio
            // becomes its own group head.
            let mut expected = kinds[i].clone();
            if let FormControlKind::OptionButton { first_in_group, .. } = &mut expected {
                *first_in_group = true;
            }
            assert_eq!(drawn.payload.kind, expected, "control {i} kind mismatch");
        }
        match &controls[0].object.unwrap().anchor {
            DrawingAnchor::TwoCell { from, to, .. } => {
                assert_eq!((from.col, from.row), (1, 1));
                assert_eq!((to.col, to.row), (3, 2));
            }
            other => panic!("expected TwoCell anchor, got {other:?}"),
        }
    }

    #[test]
    fn form_controls_and_comments_share_vml() {
        use duke_sheets_chart::{CellMarker, DrawingAnchor};
        use duke_sheets_core::comment::CellComment;
        use duke_sheets_core::{CheckState, FormControl, FormControlKind};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_comment_at(0, 0, CellComment::new("Author", "note")).unwrap();
        ws.add_form_control(
            FormControl::new(FormControlKind::Checkbox {
                caption: "check".into(),
                state: CheckState::Checked,
                cell_link: None,
                no_3d: false,
            }),
            DrawingAnchor::TwoCell {
                from: CellMarker {
                    col: 1,
                    col_offset_emu: 0,
                    row: 1,
                    row_offset_emu: 0,
                },
                to: CellMarker {
                    col: 3,
                    col_offset_emu: 0,
                    row: 2,
                    row_offset_emu: 0,
                },
                edit_as: None,
            },
        ).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.comment_count(), 1, "comment survives");
        assert_eq!(ws2.form_control_count(), 1, "control survives");
        assert_eq!(
            ws2.form_controls()
                .next()
                .unwrap()
                .payload
                .caption_text()
                .as_deref(),
            Some("check")
        );
    }

    #[test]
    fn form_control_table_relationship_ids_match_sheet_records() {
        use std::io::Read;

        use duke_sheets_chart::{CellMarker, DrawingAnchor};
        use duke_sheets_core::table::{Table, TableColumn};
        use duke_sheets_core::{CellRange, CheckState, FormControl, FormControlKind};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "Name").unwrap();
        ws.set_cell_value_at(1, 0, "Alice").unwrap();
        ws.add_table(Table {
            id: 1,
            name: "People".to_string(),
            display_name: "People".to_string(),
            reference: CellRange::parse("A1:A2").unwrap(),
            columns: vec![TableColumn {
                id: 1,
                name: "Name".to_string(),
                totals_row_function: None,
                totals_row_formula: None,
                totals_row_label: None,
                calculated_column_formula: None,
            }],
            style_info: None,
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: false,
        });
        ws.add_form_control(
            FormControl::new(FormControlKind::Checkbox {
                caption: "check".into(),
                state: CheckState::Checked,
                cell_link: None,
                no_3d: false,
            }),
            DrawingAnchor::TwoCell {
                from: CellMarker {
                    col: 2,
                    col_offset_emu: 0,
                    row: 0,
                    row_offset_emu: 0,
                },
                to: CellMarker {
                    col: 4,
                    col_offset_emu: 0,
                    row: 1,
                    row_offset_emu: 0,
                },
                edit_as: None,
            },
        ).unwrap();

        let mut bytes = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut bytes)).unwrap();
        let mut zip = zip::ZipArchive::new(Cursor::new(&bytes)).unwrap();

        let mut rels = String::new();
        zip.by_name("xl/worksheets/_rels/sheet1.bin.rels")
            .unwrap()
            .read_to_string(&mut rels)
            .unwrap();
        assert!(rels.contains("Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\""));
        assert!(rels.contains("Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\""));

        let mut sheet = Vec::new();
        zip.by_name("xl/worksheets/sheet1.bin")
            .unwrap()
            .read_to_end(&mut sheet)
            .unwrap();
        let mut iter = crate::biff12::RecordIter::new(Cursor::new(sheet));
        let mut payload = Vec::new();
        let mut legacy_rid = None;
        loop {
            let Ok(record_type) = iter.read_type() else {
                break;
            };
            let len = iter.fill_buffer(&mut payload).unwrap();
            if record_type == crate::biff12::records::BRT_LEGACY_DRAWING {
                legacy_rid = Some(
                    crate::biff12::parser::wide_str(&payload[..len], 0)
                        .unwrap()
                        .0,
                );
            }
        }
        assert_eq!(legacy_rid.as_deref(), Some("rId2"));
    }

    /// The chartEx content type must use Excel's lowercase spelling
    /// (`application/vnd.ms-office.chartex+xml`), matching the XLSX
    /// writer.
    #[test]
    fn chart_ex_content_type_is_lowercase() {
        use std::io::Read;

        let chart_ex = duke_sheets_chart::ChartEx {
            version: None,
            feature_list: None,
            fallback_img: None,
            title: None,
            data: Vec::new(),
            external_data: None,
            plot_area: Default::default(),
            legend: None,
            shape_properties: None,
            text_properties: None,
            color_map_override: None,
            format_overrides: Vec::new(),
            print_settings: None,
            style: None,
        preserved_rels: Vec::new(),
            color_style: None,
            extensions: None,
            raw_extensions: Default::default(),
            raw_mc_fallback: None,
        };
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .add_chart_ex(chart_ex, duke_sheets_chart::DrawingAnchor::default()).unwrap();

        let mut bytes = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut bytes)).unwrap();
        let mut zip = zip::ZipArchive::new(Cursor::new(&bytes)).unwrap();
        let mut content_types = String::new();
        zip.by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut content_types)
            .unwrap();
        assert!(
            content_types.contains("application/vnd.ms-office.chartex+xml"),
            "lowercase chartex content type: {content_types}"
        );
        assert!(
            !content_types.contains("chartEx+xml"),
            "no camel-case chartEx content type: {content_types}"
        );
    }

    /// Non-visual CF rules must encode as CF_TYPE_EXPRIS (2) with the
    /// matching CFTemp template. MS-XLSB 2.5.18 CFType defines only
    /// values 1..6; 2.4.23 BrtBeginCFRule pins the allowed
    /// iType/iTemplate pairs and the iParam semantics (CFTextOper
    /// 2.5.17 for CONTAINSTEXT, CFDateOper 2.5.12 for TIMEPERIOD*).
    #[test]
    fn cf_rule_records_use_spec_type_template_pairs() {
        use duke_sheets_core::conditional_format::{
            CfRuleType, ConditionalFormatRule, TimePeriod,
        };
        use std::io::Read;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let range = CellRange::parse("A1:A10").unwrap();
        let rules = [
            CfRuleType::Expression {
                formula: "MOD(A1,2)=0".into(),
            },
            CfRuleType::ContainsText { text: "x".into() },
            CfRuleType::BeginsWith { text: "x".into() },
            CfRuleType::EndsWith { text: "x".into() },
            CfRuleType::UniqueValues,
            CfRuleType::DuplicateValues,
            CfRuleType::ContainsBlanks,
            CfRuleType::NotContainsBlanks,
            CfRuleType::ContainsErrors,
            CfRuleType::NotContainsErrors,
            CfRuleType::TimePeriod {
                period: TimePeriod::LastWeek,
            },
            CfRuleType::AboveAverage {
                above: true,
                equal_average: false,
                std_dev: None,
            },
            CfRuleType::AboveAverage {
                above: false,
                equal_average: true,
                std_dev: None,
            },
        ];
        for (i, rule_type) in rules.into_iter().enumerate() {
            ws.add_conditional_format(
                ConditionalFormatRule::new(rule_type)
                    .with_range(range.clone())
                    .with_priority((i + 1) as u32),
            );
        }

        let mut bytes = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut bytes)).unwrap();
        let mut zip = zip::ZipArchive::new(Cursor::new(&bytes)).unwrap();
        let mut sheet = Vec::new();
        zip.by_name("xl/worksheets/sheet1.bin")
            .unwrap()
            .read_to_end(&mut sheet)
            .unwrap();

        let mut iter = crate::biff12::RecordIter::new(Cursor::new(sheet));
        let mut payload = Vec::new();
        let mut triples = Vec::new();
        while let Ok((typ, len)) = iter.next_record(&mut payload) {
            if typ == crate::biff12::records::BRT_BEGIN_CF_RULE && len >= 20 {
                triples.push((
                    crate::biff12::parser::read_u32(&payload, 0),
                    crate::biff12::parser::read_u32(&payload, 4),
                    crate::biff12::parser::read_u32(&payload, 16),
                ));
            }
        }
        assert_eq!(
            triples,
            vec![
                (2, 0x01, 0), // Expression: FMLA
                (2, 0x08, 0), // ContainsText: CONTAINSTEXT + CF_TEXTOPER_CONTAINS
                (2, 0x08, 2), // BeginsWith: CONTAINSTEXT + CF_TEXTOPER_BEGINSWITH
                (2, 0x08, 3), // EndsWith: CONTAINSTEXT + CF_TEXTOPER_ENDSWITH
                (2, 0x07, 0), // UniqueValues
                (2, 0x1B, 0), // DuplicateValues
                (2, 0x09, 0), // ContainsBlanks
                (2, 0x0A, 0), // NotContainsBlanks
                (2, 0x0B, 0), // ContainsErrors
                (2, 0x0C, 0), // NotContainsErrors
                (2, 0x17, 4), // TimePeriod LastWeek + CF_TIMEPERIOD_LASTWEEK
                (2, 0x19, 0), // AboveAverage
                (2, 0x1E, 0), // EqualBelowAverage
            ],
            "BrtBeginCFRule (iType, iTemplate, iParam) triples"
        );
    }
}
