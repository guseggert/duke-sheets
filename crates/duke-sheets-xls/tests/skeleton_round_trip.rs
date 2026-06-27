#![allow(clippy::approx_constant)]
//! Round-trip tests for the XLS skeleton writer.
//!
//! Build an empty `Workbook`, write it to BIFF8 bytes, read it back
//! through `XlsReader`, and confirm the structure (sheet count, sheet
//! names) round-trips. The skeleton writer doesn't yet emit cells,
//! formatting, or formulas — those land in subsequent slices.

use std::io::Cursor;

use duke_sheets_core::{
    CellRange, PivotAggregate, PivotDateGroupUnit, PivotFieldRef, PivotFilter, PivotGrouping,
    PivotManualGroup, PivotSource, PivotStyle, PivotTable, PivotValue, PivotValuesAxis, Workbook,
};
use duke_sheets_xls::{cfb::CompoundFile, XlsReader, XlsWriter};

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
        .page("Salesperson")
        .filter(PivotFilter::field_items("Salesperson", ["Ada"]))
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
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
        .page("Segment")
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
        .page("Date")
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

    let sxfmla = cache_records
        .iter()
        .find_map(|(record_type, payload)| (*record_type == 0x00F8).then_some(payload))
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
            (*record_type == 0x00F5).then(|| i16::from_le_bytes(payload[2..4].try_into().unwrap()))
        })
        .collect::<Vec<_>>();
    assert_eq!(sxname_field_indexes, vec![1, 2]);
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
fn xls_manual_grouping_rejects_non_text_items() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Age").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", 7.0).unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", 13.0).unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("ManualAgeBands")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Age")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: PivotFieldRef::new("Age"),
            groups: vec![PivotManualGroup::new(
                "Young",
                [PivotValue::Number(7.0), PivotValue::Number(13.0)],
            )],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let err = XlsWriter::write_to_bytes(&wb).expect_err("numeric manual grouping should fail");
    assert!(
        err.to_string()
            .contains("currently supports only text or blank source items"),
        "{err}"
    );
}

#[test]
fn xls_manual_grouping_rejects_page_axis_fields() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("ManualPageRegions")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .page("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: PivotFieldRef::new("Region"),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let err = XlsWriter::write_to_bytes(&wb).expect_err("page manual grouping should fail");
    assert!(
        err.to_string()
            .contains("currently supports only row-axis fields"),
        "{err}"
    );
}

#[test]
fn xls_manual_grouping_rejects_column_axis_fields() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Product").unwrap();
    ws.set_cell_value("C1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", "A").unwrap();
    ws.set_cell_value("C2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", "A").unwrap();
    ws.set_cell_value("C3", 20.0).unwrap();

    let pivot = PivotTable::builder("ManualColumnRegions")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Product")
        .column("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .grouping(PivotGrouping::Manual {
            field: PivotFieldRef::new("Region"),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

    let err = XlsWriter::write_to_bytes(&wb).expect_err("column manual grouping should fail");
    assert!(
        err.to_string()
            .contains("currently supports only row-axis fields"),
        "{err}"
    );
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
