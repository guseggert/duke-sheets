//! Facade-level pivot refresh and format round-trip tests.

use duke_sheets::{
    CellRange, PivotAggregate, PivotFilter, PivotTable, Workbook, WorkbookExt, WorkbookPivotExt,
};

const PIVOT_NAME: &str = "SalesByQuarter";
const CALCULATED_FORMULA: &str = "Revenue*$K$1";

fn build_workbook(calculated: bool) -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).expect("default worksheet");
    for (address, value) in [
        ("A1", "Segment"),
        ("B1", "Region"),
        ("C1", "Quarter"),
        ("D1", "Revenue"),
        ("A2", "Online"),
        ("B2", "East"),
        ("C2", "Q1"),
        ("A3", "Retail"),
        ("B3", "East"),
        ("C3", "Q2"),
        ("A4", "Online"),
        ("B4", "West"),
        ("C4", "Q1"),
        ("A5", "Online"),
        ("B5", "West"),
        ("C5", "Q2"),
    ] {
        sheet.set_cell_value(address, value).unwrap();
    }
    for (address, value) in [("D2", 10.0), ("D3", 20.0), ("D4", 30.0), ("D5", 5.0)] {
        sheet.set_cell_value(address, value).unwrap();
    }
    sheet.set_cell_value("K1", 2.0).unwrap();

    let builder = PivotTable::builder(PIVOT_NAME)
        .source_range(CellRange::parse("A1:D5").unwrap())
        .target_address("F2")
        .unwrap()
        .page("Segment")
        .row("Region")
        .column("Quarter")
        .filter(PivotFilter::field_items("Segment", ["Online"]));
    let pivot = if calculated {
        builder
            .calculated_field("Adjusted", format!("={CALCULATED_FORMULA}"))
            .named_measure("Adjusted", PivotAggregate::Sum, "Adjusted Revenue")
            .build()
            .unwrap()
    } else {
        builder
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .build()
            .unwrap()
    };
    sheet.add_pivot_table(pivot).unwrap();

    let stats = workbook.refresh_pivots().expect("refresh pivot");
    assert_refresh_stats(stats);
    assert_pivot(&workbook, calculated);
    assert_cells(&workbook, if calculated { 2.0 } else { 1.0 });
    workbook
}

fn assert_refresh_stats(stats: duke_sheets::PivotRefreshStats) {
    assert_eq!(stats.pivot_count, 1);
    assert_eq!(stats.pivots_refreshed, 1);
    assert_eq!(stats.source_rows, 4);
    assert_eq!(stats.output_cells, 24);
    assert_eq!(stats.cache_hits, 0);
    assert_eq!(stats.cache_misses, 1);
}

fn assert_pivot(workbook: &Workbook, calculated: bool) {
    let pivot = workbook
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name(PIVOT_NAME)
        .expect("semantic pivot");

    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.columns.len(), 1);
    assert_eq!(pivot.columns[0].field.name, "Quarter");
    assert_eq!(pivot.page_fields.len(), 1);
    assert_eq!(pivot.page_fields[0].field.name, "Segment");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);

    if calculated {
        assert_eq!(pivot.measures[0].field.name, "Adjusted");
        assert_eq!(pivot.measures[0].name.as_deref(), Some("Adjusted Revenue"));
        assert_eq!(pivot.calculated_fields.len(), 1);
        assert_eq!(pivot.calculated_fields[0].name, "Adjusted");
        assert_eq!(
            pivot.calculated_fields[0].formula.trim_start_matches('='),
            CALCULATED_FORMULA
        );
    } else {
        assert_eq!(pivot.measures[0].field.name, "Revenue");
        assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Revenue"));
        assert!(pivot.calculated_fields.is_empty());
    }
}

fn assert_cells(workbook: &Workbook, multiplier: f64) {
    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.get_value("A1").unwrap().as_string(), Some("Segment"));
    assert_eq!(sheet.get_value("A3").unwrap().as_string(), Some("Retail"));
    assert_eq!(sheet.get_value("D5").unwrap().as_number(), Some(5.0));
    assert_eq!(sheet.get_value("K1").unwrap().as_number(), Some(2.0));

    assert_eq!(sheet.get_value("F2").unwrap().as_string(), Some("Segment"));
    assert_eq!(sheet.get_value("G2").unwrap().as_string(), Some("Online"));
    assert_eq!(sheet.get_value("F4").unwrap().as_string(), Some("Region"));
    assert_eq!(sheet.get_value("G4").unwrap().as_string(), Some("Q1"));
    assert_eq!(sheet.get_value("H4").unwrap().as_string(), Some("Q2"));
    assert_eq!(
        sheet.get_value("I4").unwrap().as_string(),
        Some("Grand Total")
    );
    assert_eq!(sheet.get_value("F5").unwrap().as_string(), Some("East"));
    assert_eq!(
        sheet.get_value("G5").unwrap().as_number(),
        Some(10.0 * multiplier)
    );
    assert_eq!(
        sheet.get_value("I5").unwrap().as_number(),
        Some(10.0 * multiplier)
    );
    assert_eq!(sheet.get_value("F6").unwrap().as_string(), Some("West"));
    assert_eq!(
        sheet.get_value("G6").unwrap().as_number(),
        Some(30.0 * multiplier)
    );
    assert_eq!(
        sheet.get_value("H6").unwrap().as_number(),
        Some(5.0 * multiplier)
    );
    assert_eq!(
        sheet.get_value("I6").unwrap().as_number(),
        Some(35.0 * multiplier)
    );
    assert_eq!(
        sheet.get_value("F7").unwrap().as_string(),
        Some("Grand Total")
    );
    assert_eq!(
        sheet.get_value("G7").unwrap().as_number(),
        Some(40.0 * multiplier)
    );
    assert_eq!(
        sheet.get_value("H7").unwrap().as_number(),
        Some(5.0 * multiplier)
    );
    assert_eq!(
        sheet.get_value("I7").unwrap().as_number(),
        Some(45.0 * multiplier)
    );
}

fn save_and_reopen(workbook: &Workbook, extension: &str) -> Workbook {
    let file = tempfile::Builder::new()
        .suffix(&format!(".{extension}"))
        .tempfile()
        .expect("temporary workbook");
    workbook.save(file.path()).expect("save through facade");
    let bytes = std::fs::read(file.path()).expect("read saved workbook");
    Workbook::from_bytes(&bytes).expect("reopen through facade")
}

fn assert_reopened(mut workbook: Workbook, calculated: bool) {
    let multiplier = if calculated { 2.0 } else { 1.0 };
    assert_pivot(&workbook, calculated);
    assert_cells(&workbook, multiplier);
    assert_eq!(
        workbook.worksheet(0).unwrap().pivot_tables()[0]
            .target
            .to_a1_string(),
        "F2",
        "round-trip changed the semantic pivot target"
    );

    let stats = workbook.refresh_pivots().expect("refresh reopened pivot");
    assert_refresh_stats(stats);
    assert_cells(&workbook, multiplier);
}

#[cfg(feature = "xlsx")]
#[test]
fn pivot_round_trips_through_xlsx_facade() {
    let workbook = build_workbook(true);
    assert_reopened(save_and_reopen(&workbook, "xlsx"), true);
}

#[cfg(feature = "xlsb")]
#[test]
fn pivot_round_trips_through_xlsb_facade() {
    let workbook = build_workbook(false);
    assert_reopened(save_and_reopen(&workbook, "xlsb"), false);
}

#[cfg(feature = "xls")]
#[test]
fn pivot_round_trips_through_xls_facade() {
    let workbook = build_workbook(false);
    assert_reopened(save_and_reopen(&workbook, "xls"), false);
}
