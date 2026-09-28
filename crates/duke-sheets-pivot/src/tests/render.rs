use crate::test_support::*;
use pretty_assertions::assert_eq;

#[test]
fn refreshes_row_and_column_fields() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Q1").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source(PivotSource::range(CellRange::parse("A1:C4").unwrap()))
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Region");
    assert_eq!(text(&workbook, "F1"), "Q1");
    assert_eq!(text(&workbook, "G1"), "Q2");
    assert_eq!(text(&workbook, "H1"), "Grand Total");
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 10.0);
    assert_eq!(number(&workbook, "G2"), 5.0);
    assert_eq!(number(&workbook, "H2"), 15.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(number(&workbook, "F3"), 7.0);
    assert_eq!(number(&workbook, "H3"), 7.0);
    assert_eq!(text(&workbook, "E4"), "Grand Total");
    assert_eq!(number(&workbook, "F4"), 17.0);
    assert_eq!(number(&workbook, "G4"), 5.0);
    assert_eq!(number(&workbook, "H4"), 22.0);
}

#[test]
fn refreshes_values_axis_on_rows() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("D1", "Units").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("D2", 2.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("D3", 3.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Q1").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();
    sheet.set_cell_value("D4", 4.0).unwrap();

    let mut layout = tabular_layout();
    layout.data_caption = "Metric".to_string();
    layout.values_axis = PivotValuesAxis::Rows;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Revenue")
                .with_number_format("0.0"),
        )
        .pivot_measure(
            PivotMeasure::new("Units", PivotAggregate::Sum)
                .with_name("Units")
                .with_number_format("0.000"),
        )
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "Metric");
    assert_eq!(text(&workbook, "H1"), "Q1");
    assert_eq!(text(&workbook, "I1"), "Q2");
    assert_eq!(text(&workbook, "J1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(text(&workbook, "G2"), "Revenue");
    assert_eq!(number(&workbook, "H2"), 10.0);
    assert_eq!(number(&workbook, "I2"), 5.0);
    assert_eq!(number(&workbook, "J2"), 15.0);
    assert_eq!(text(&workbook, "F3"), "East");
    assert_eq!(text(&workbook, "G3"), "Units");
    assert_eq!(number(&workbook, "H3"), 2.0);
    assert_eq!(number(&workbook, "I3"), 3.0);
    assert_eq!(number(&workbook, "J3"), 5.0);
    assert_eq!(text(&workbook, "F4"), "West");
    assert_eq!(text(&workbook, "G4"), "Revenue");
    assert_eq!(number(&workbook, "H4"), 7.0);
    assert_eq!(text(&workbook, "I4"), "");
    assert_eq!(number(&workbook, "J4"), 7.0);
    assert_eq!(text(&workbook, "F5"), "West");
    assert_eq!(text(&workbook, "G5"), "Units");
    assert_eq!(number(&workbook, "H5"), 4.0);
    assert_eq!(text(&workbook, "I5"), "");
    assert_eq!(number(&workbook, "J5"), 4.0);
    assert_eq!(text(&workbook, "F6"), "Grand Total");
    assert_eq!(text(&workbook, "G6"), "Revenue");
    assert_eq!(number(&workbook, "H6"), 17.0);
    assert_eq!(number(&workbook, "I6"), 5.0);
    assert_eq!(number(&workbook, "J6"), 22.0);
    assert_eq!(text(&workbook, "F7"), "Grand Total");
    assert_eq!(text(&workbook, "G7"), "Units");
    assert_eq!(number(&workbook, "H7"), 6.0);
    assert_eq!(number(&workbook, "I7"), 3.0);
    assert_eq!(number(&workbook, "J7"), 9.0);

    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.formatted_value("H2").unwrap(), "10.0");
    assert_eq!(sheet.formatted_value("H3").unwrap(), "2.000");
    assert_eq!(sheet.formatted_value("J6").unwrap(), "22.0");
    assert_eq!(sheet.formatted_value("J7").unwrap(), "9.000");
}

#[test]
fn refreshes_values_axis_on_rows_measure_number_formats() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("C1", "Rate").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("C2", 0.25).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("C3", 0.5).unwrap();

    let mut layout = tabular_layout();
    layout.data_caption = "Metric".to_string();
    layout.values_axis = PivotValuesAxis::Rows;
    let pivot = PivotTable::builder("FormatPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .pivot_measure(
            PivotMeasure::new("Revenue", PivotAggregate::Sum)
                .with_name("Revenue")
                .with_number_format("0.0"),
        )
        .pivot_measure(
            PivotMeasure::new("Rate", PivotAggregate::Sum)
                .with_name("Rate")
                .with_number_format("0.0%"),
        )
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.formatted_value("G2").unwrap(), "10.0");
    assert_eq!(sheet.formatted_value("G3").unwrap(), "25.0%");
    assert_eq!(sheet.formatted_value("G6").unwrap(), "30.0");
    assert_eq!(sheet.formatted_value("G7").unwrap(), "75.0%");
}

#[test]
fn refreshes_values_axis_on_rows_at_position() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Product").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("D1", "Units").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "A").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("D2", 2.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "B").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("D3", 3.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "A").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();
    sheet.set_cell_value("D4", 4.0).unwrap();

    let mut layout = tabular_layout();
    layout.data_caption = "Metric".to_string();
    layout.values_axis = PivotValuesAxis::Rows;
    layout.values_axis_position = Some(1);
    let mut region = PivotField::new("Region");
    region.subtotal = PivotSubtotal::None;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row(region)
        .row("Product")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .named_measure("Units", PivotAggregate::Sum, "Units")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "Metric");
    assert_eq!(text(&workbook, "H1"), "Product");
    assert_eq!(text(&workbook, "I1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(text(&workbook, "G2"), "Revenue");
    assert_eq!(text(&workbook, "H2"), "A");
    assert_eq!(number(&workbook, "I2"), 10.0);
    assert_eq!(text(&workbook, "F3"), "East");
    assert_eq!(text(&workbook, "G3"), "Units");
    assert_eq!(text(&workbook, "H3"), "A");
    assert_eq!(number(&workbook, "I3"), 2.0);
    assert_eq!(text(&workbook, "F4"), "East");
    assert_eq!(text(&workbook, "G4"), "Revenue");
    assert_eq!(text(&workbook, "H4"), "B");
    assert_eq!(number(&workbook, "I4"), 5.0);
    assert_eq!(text(&workbook, "F5"), "East");
    assert_eq!(text(&workbook, "G5"), "Units");
    assert_eq!(text(&workbook, "H5"), "B");
    assert_eq!(number(&workbook, "I5"), 3.0);
    assert_eq!(text(&workbook, "F6"), "West");
    assert_eq!(text(&workbook, "G6"), "Revenue");
    assert_eq!(text(&workbook, "H6"), "A");
    assert_eq!(number(&workbook, "I6"), 7.0);
    assert_eq!(text(&workbook, "F7"), "West");
    assert_eq!(text(&workbook, "G7"), "Units");
    assert_eq!(text(&workbook, "H7"), "A");
    assert_eq!(number(&workbook, "I7"), 4.0);
    assert_eq!(text(&workbook, "F8"), "Grand Total");
    assert_eq!(text(&workbook, "G8"), "Revenue");
    assert_eq!(text(&workbook, "H8"), "");
    assert_eq!(number(&workbook, "I8"), 22.0);
    assert_eq!(text(&workbook, "F9"), "Grand Total");
    assert_eq!(text(&workbook, "G9"), "Units");
    assert_eq!(text(&workbook, "H9"), "");
    assert_eq!(number(&workbook, "I9"), 9.0);
}

#[test]
fn refreshes_values_axis_on_rows_with_row_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("E1", "Units").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("E2", 2.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("E3", 3.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();
    sheet.set_cell_value("E4", 4.0).unwrap();

    let mut layout = tabular_layout();
    layout.data_caption = "Metric".to_string();
    layout.values_axis = PivotValuesAxis::Rows;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:E4").unwrap())
        .target_address("G1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .named_measure("Units", PivotAggregate::Sum, "Units")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "G1"), "Region");
    assert_eq!(text(&workbook, "H1"), "Segment");
    assert_eq!(text(&workbook, "I1"), "Metric");
    assert_eq!(text(&workbook, "J1"), "Q1");
    assert_eq!(text(&workbook, "K1"), "Q2");
    assert_eq!(text(&workbook, "L1"), "Grand Total");
    assert_eq!(text(&workbook, "G6"), "East Total");
    assert_eq!(text(&workbook, "H6"), "");
    assert_eq!(text(&workbook, "I6"), "Revenue");
    assert_eq!(number(&workbook, "J6"), 10.0);
    assert_eq!(number(&workbook, "K6"), 5.0);
    assert_eq!(number(&workbook, "L6"), 15.0);
    assert_eq!(text(&workbook, "G7"), "East Total");
    assert_eq!(text(&workbook, "I7"), "Units");
    assert_eq!(number(&workbook, "J7"), 2.0);
    assert_eq!(number(&workbook, "K7"), 3.0);
    assert_eq!(number(&workbook, "L7"), 5.0);
    assert_eq!(text(&workbook, "G10"), "West Total");
    assert_eq!(text(&workbook, "I10"), "Revenue");
    assert_eq!(number(&workbook, "J10"), 7.0);
    assert_eq!(text(&workbook, "K10"), "");
    assert_eq!(number(&workbook, "L10"), 7.0);
    assert_eq!(text(&workbook, "G11"), "West Total");
    assert_eq!(text(&workbook, "I11"), "Units");
    assert_eq!(number(&workbook, "J11"), 4.0);
    assert_eq!(text(&workbook, "K11"), "");
    assert_eq!(number(&workbook, "L11"), 4.0);
    assert_eq!(text(&workbook, "G12"), "Grand Total");
    assert_eq!(text(&workbook, "I12"), "Revenue");
    assert_eq!(number(&workbook, "J12"), 17.0);
    assert_eq!(number(&workbook, "K12"), 5.0);
    assert_eq!(number(&workbook, "L12"), 22.0);
    assert_eq!(text(&workbook, "G13"), "Grand Total");
    assert_eq!(text(&workbook, "I13"), "Units");
    assert_eq!(number(&workbook, "J13"), 6.0);
    assert_eq!(number(&workbook, "K13"), 3.0);
    assert_eq!(number(&workbook, "L13"), 9.0);
}

#[test]
fn refreshes_outline_values_axis_on_rows_subtotals_at_top() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("D1", "Units").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("D2", 2.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("D3", 3.0).unwrap();

    let mut region = PivotField::new("Region");
    region.subtotal_top = true;
    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Outline;
    layout.values_axis = PivotValuesAxis::Rows;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D3").unwrap())
        .target_address("F1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .named_measure("Units", PivotAggregate::Sum, "Units")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F2"), "East Total");
    assert_eq!(text(&workbook, "G2"), "");
    assert_eq!(text(&workbook, "H2"), "Revenue");
    assert_eq!(number(&workbook, "I2"), 15.0);
    assert_eq!(text(&workbook, "F3"), "East Total");
    assert_eq!(text(&workbook, "H3"), "Units");
    assert_eq!(number(&workbook, "I3"), 5.0);
    assert_eq!(text(&workbook, "F4"), "");
    assert_eq!(text(&workbook, "G4"), "Online");
    assert_eq!(text(&workbook, "H4"), "Revenue");
    assert_eq!(number(&workbook, "I4"), 5.0);
    assert_eq!(text(&workbook, "F8"), "Grand Total");
    assert_eq!(text(&workbook, "H8"), "Revenue");
    assert_eq!(number(&workbook, "I8"), 15.0);
    assert_eq!(text(&workbook, "F9"), "Grand Total");
    assert_eq!(text(&workbook, "H9"), "Units");
    assert_eq!(number(&workbook, "I9"), 5.0);
}

#[test]
fn refreshes_compact_values_axis_on_rows_with_row_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("D1", "Units").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("D2", 2.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("D3", 3.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.values_axis = PivotValuesAxis::Rows;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D3").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .named_measure("Units", PivotAggregate::Sum, "Units")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(text(&workbook, "F3"), "Online");
    assert_eq!(text(&workbook, "G3"), "Revenue");
    assert_eq!(number(&workbook, "H3"), 5.0);
    assert_eq!(text(&workbook, "F7"), "East Total");
    assert_eq!(text(&workbook, "G7"), "Revenue");
    assert_eq!(number(&workbook, "H7"), 15.0);
    assert_eq!(text(&workbook, "F8"), "East Total");
    assert_eq!(text(&workbook, "G8"), "Units");
    assert_eq!(number(&workbook, "H8"), 5.0);
    assert_eq!(text(&workbook, "F9"), "Grand Total");
    assert_eq!(text(&workbook, "G9"), "Revenue");
    assert_eq!(number(&workbook, "H9"), 15.0);
    assert_eq!(text(&workbook, "F10"), "Grand Total");
    assert_eq!(text(&workbook, "G10"), "Units");
    assert_eq!(number(&workbook, "H10"), 5.0);
}

#[test]
fn refreshes_values_axis_on_rows_page_break_offsets() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("D1", "Units").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("D2", 2.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("D3", 3.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();
    sheet.set_cell_value("D4", 4.0).unwrap();
    sheet.add_row_break(20);

    let mut region = PivotField::new("Region");
    region.insert_page_break = true;
    let mut layout = tabular_layout();
    layout.values_axis = PivotValuesAxis::Rows;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .named_measure("Units", PivotAggregate::Sum, "Units")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();
    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    let mut pivot_breaks = sheet
        .row_breaks()
        .iter()
        .filter(|break_| break_.pt)
        .map(|break_| break_.id)
        .collect::<Vec<_>>();
    pivot_breaks.sort_unstable();
    assert_eq!(pivot_breaks, vec![6, 10]);

    let mut user_breaks = sheet
        .row_breaks()
        .iter()
        .filter(|break_| !break_.pt)
        .map(|break_| break_.id)
        .collect::<Vec<_>>();
    user_breaks.sort_unstable();
    assert_eq!(user_breaks, vec![20]);
}

#[test]
fn refresh_applies_grand_total_caption() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Q1").unwrap();
    sheet.set_cell_value("C4", 5.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.grand_total_caption = Some("Overall".to_string());
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .measure("Revenue", PivotAggregate::Sum)
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "H1"), "Overall");
    assert_eq!(text(&workbook, "E4"), "Overall");
    assert_eq!(number(&workbook, "H4"), 35.0);
}

#[test]
fn refresh_applies_axis_subtotal_caption() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Tabular;
    let region = PivotField::new("Region").with_subtotal_caption("Subtotal");
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E4"), "East Subtotal");
    assert_eq!(number(&workbook, "G4"), 15.0);
    assert_eq!(text(&workbook, "E6"), "West Subtotal");
    assert_eq!(number(&workbook, "G6"), 7.0);
}

#[test]
fn refresh_applies_total_asterisk_captions() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Tabular;
    layout.asterisk_totals = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .column("Quarter")
        .measure("Revenue", PivotAggregate::Sum)
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "J1"), "Grand Total*");
    assert_eq!(text(&workbook, "F4"), "East Total*");
    assert_eq!(number(&workbook, "J4"), 15.0);
    assert_eq!(text(&workbook, "F6"), "West Total*");
    assert_eq!(text(&workbook, "F7"), "Grand Total*");
    assert_eq!(number(&workbook, "J7"), 22.0);
}

#[test]
fn refresh_applies_missing_value_caption() {
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
    sheet.set_cell_value("C3", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.missing_caption = Some("-".to_string());
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .measure("Revenue", PivotAggregate::Sum)
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Q1");
    assert_eq!(text(&workbook, "G1"), "Q2");
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 10.0);
    assert_eq!(text(&workbook, "G2"), "-");
    assert_eq!(number(&workbook, "H2"), 10.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(text(&workbook, "F3"), "-");
    assert_eq!(number(&workbook, "G3"), 7.0);
    assert_eq!(number(&workbook, "H3"), 7.0);
}

#[test]
fn refresh_applies_error_caption_to_axis_labels() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet
        .set_cell_value("A2", CellValue::Error(CellError::Div0))
        .unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet
        .set_cell_value("B3", CellValue::Error(CellError::Value))
        .unwrap();
    sheet.set_cell_value("C3", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.show_error = true;
    layout.error_caption = Some("ERR".to_string());
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .measure("Revenue", PivotAggregate::Sum)
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Q1");
    assert_eq!(text(&workbook, "G1"), "ERR");
    assert_eq!(text(&workbook, "E2"), "West");
    assert_eq!(text(&workbook, "E3"), "ERR");
    assert_eq!(number(&workbook, "G2"), 7.0);
    assert_eq!(number(&workbook, "F3"), 10.0);
}

#[test]
fn refresh_respects_hidden_field_headers() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.show_field_headers = false;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D1"), "East");
    assert_eq!(number(&workbook, "E1"), 10.0);
    assert_eq!(text(&workbook, "D2"), "West");
    assert_eq!(number(&workbook, "E2"), 20.0);
    assert_eq!(text(&workbook, "D3"), "Grand Total");
    assert_eq!(number(&workbook, "E3"), 30.0);
}

#[test]
fn refresh_applies_axis_field_captions() {
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

    let region = PivotField::new("Region").with_caption("Market");
    let period = PivotField::new("Quarter").with_caption("Period");
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .page(period)
        .row(region)
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Period");
    assert_eq!(text(&workbook, "E3"), "Market");
    assert_eq!(number(&workbook, "F4"), 10.0);
    assert_eq!(number(&workbook, "F5"), 20.0);
}

#[test]
fn refreshes_tabular_layout_without_repeated_item_labels() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Tabular;
    layout.repeat_item_labels = false;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Region");
    assert_eq!(text(&workbook, "F1"), "Segment");
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "Online");
    assert_eq!(text(&workbook, "E3"), "");
    assert_eq!(text(&workbook, "F3"), "Retail");
    assert_eq!(text(&workbook, "E4"), "East Total");
    assert_eq!(number(&workbook, "G4"), 15.0);
    assert_eq!(text(&workbook, "E5"), "West");
    assert_eq!(text(&workbook, "F5"), "Retail");
    assert_eq!(number(&workbook, "G5"), 7.0);
}

#[test]
fn refreshes_tabular_layout_with_repeated_item_labels() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Tabular;
    layout.repeat_item_labels = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "Online");
    assert_eq!(text(&workbook, "E3"), "East");
    assert_eq!(text(&workbook, "F3"), "Retail");
    assert_eq!(text(&workbook, "E4"), "East Total");
}

#[test]
fn refreshes_outline_layout_subtotals_at_top() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Outline;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Region");
    assert_eq!(text(&workbook, "F1"), "Segment");
    assert_eq!(text(&workbook, "E2"), "East Total");
    assert_eq!(number(&workbook, "G2"), 15.0);
    assert_eq!(text(&workbook, "E3"), "");
    assert_eq!(text(&workbook, "F3"), "Online");
    assert_eq!(number(&workbook, "G3"), 5.0);
    assert_eq!(text(&workbook, "E4"), "");
    assert_eq!(text(&workbook, "F4"), "Retail");
    assert_eq!(number(&workbook, "G4"), 10.0);
    assert_eq!(text(&workbook, "E5"), "West Total");
    assert_eq!(number(&workbook, "G5"), 7.0);
    assert_eq!(text(&workbook, "E6"), "");
    assert_eq!(text(&workbook, "F6"), "Retail");
    assert_eq!(number(&workbook, "G6"), 7.0);
    assert_eq!(text(&workbook, "E7"), "Grand Total");
    assert_eq!(number(&workbook, "G7"), 22.0);
}

#[test]
fn refresh_writes_row_outline_levels_for_outline_layout() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Outline;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.row_outline_level(0), 0);
    assert_eq!(sheet.row_outline_level(1), 0);
    assert_eq!(sheet.row_outline_level(2), 1);
    assert_eq!(sheet.row_outline_level(3), 1);
    assert_eq!(sheet.row_outline_level(4), 0);
    assert_eq!(sheet.row_outline_level(5), 1);
    assert_eq!(sheet.row_outline_level(6), 0);
}

#[test]
fn refresh_writes_row_collapsed_flags_for_collapsed_items() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Outline;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row(PivotField::new("Region").with_collapsed_items(["East"]))
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    assert!(sheet.is_row_collapsed(1));
    assert!(!sheet.is_row_hidden(1));
    assert!(sheet.is_row_hidden(2));
    assert!(sheet.is_row_hidden(3));
    assert!(!sheet.is_row_hidden(4));
    assert!(!sheet.is_row_collapsed(2));
    assert!(!sheet.is_row_collapsed(3));
    assert!(!sheet.is_row_collapsed(4));
}

#[test]
fn refresh_clears_stale_pivot_row_outline_levels() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Outline;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();
    assert_eq!(workbook.worksheet(0).unwrap().row_outline_level(2), 1);

    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.pivot_tables_mut()[0].layout.show_expand_collapse = false;
    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    for row in 0..=4 {
        assert_eq!(sheet.row_outline_level(row), 0);
        assert!(!sheet.is_row_collapsed(row));
    }
}

#[test]
fn refresh_merges_and_clears_tabular_item_labels() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Tabular;
    layout.merge_item_labels = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    let merged_label = CellRange::parse("E2:E3").unwrap();
    assert!(workbook
        .worksheet(0)
        .unwrap()
        .merged_regions()
        .contains(&merged_label));
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "E3"), "");
    assert_eq!(text(&workbook, "E4"), "East Total");

    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.pivot_tables_mut()[0].layout.merge_item_labels = false;
    workbook.refresh_pivots().unwrap();

    assert!(!workbook
        .worksheet(0)
        .unwrap()
        .merged_regions()
        .contains(&merged_label));
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "E3"), "");
}

#[test]
fn refreshes_compact_layout_hierarchy_without_column_fields() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Row Labels");
    assert_eq!(text(&workbook, "F1"), "Revenue");
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "");
    assert_eq!(text(&workbook, "E3"), "Online");
    assert_eq!(number(&workbook, "F3"), 5.0);
    assert_eq!(text(&workbook, "E4"), "Retail");
    assert_eq!(number(&workbook, "F4"), 10.0);
    assert_eq!(text(&workbook, "E5"), "East Total");
    assert_eq!(number(&workbook, "F5"), 15.0);
    assert_eq!(text(&workbook, "E6"), "West");
    assert_eq!(text(&workbook, "F6"), "");
    assert_eq!(text(&workbook, "E7"), "Retail");
    assert_eq!(number(&workbook, "F7"), 7.0);
    assert_eq!(text(&workbook, "E8"), "West Total");
    assert_eq!(number(&workbook, "F8"), 7.0);
    assert_eq!(text(&workbook, "E9"), "Grand Total");
    assert_eq!(number(&workbook, "F9"), 22.0);
}

#[test]
fn refreshes_compact_layout_hierarchy_with_column_fields() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Row Labels");
    assert_eq!(text(&workbook, "G1"), "Q1");
    assert_eq!(text(&workbook, "H1"), "Q2");
    assert_eq!(text(&workbook, "I1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(text(&workbook, "G2"), "");
    assert_eq!(text(&workbook, "H2"), "");
    assert_eq!(text(&workbook, "I2"), "");
    assert_eq!(text(&workbook, "F3"), "Online");
    assert_eq!(text(&workbook, "G3"), "");
    assert_eq!(number(&workbook, "H3"), 5.0);
    assert_eq!(number(&workbook, "I3"), 5.0);
    assert_eq!(text(&workbook, "F4"), "Retail");
    assert_eq!(number(&workbook, "G4"), 10.0);
    assert_eq!(text(&workbook, "H4"), "");
    assert_eq!(number(&workbook, "I4"), 10.0);
    assert_eq!(text(&workbook, "F5"), "East Total");
    assert_eq!(number(&workbook, "G5"), 10.0);
    assert_eq!(number(&workbook, "H5"), 5.0);
    assert_eq!(number(&workbook, "I5"), 15.0);
    assert_eq!(text(&workbook, "F6"), "West");
    assert_eq!(text(&workbook, "G6"), "");
    assert_eq!(text(&workbook, "H6"), "");
    assert_eq!(text(&workbook, "I6"), "");
    assert_eq!(text(&workbook, "F7"), "Retail");
    assert_eq!(number(&workbook, "G7"), 7.0);
    assert_eq!(text(&workbook, "H7"), "");
    assert_eq!(number(&workbook, "I7"), 7.0);
    assert_eq!(text(&workbook, "F8"), "West Total");
    assert_eq!(number(&workbook, "G8"), 7.0);
    assert_eq!(text(&workbook, "H8"), "");
    assert_eq!(number(&workbook, "I8"), 7.0);
    assert_eq!(text(&workbook, "F9"), "Grand Total");
    assert_eq!(number(&workbook, "G9"), 17.0);
    assert_eq!(number(&workbook, "H9"), 5.0);
    assert_eq!(number(&workbook, "I9"), 22.0);
}

#[test]
fn refreshes_show_empty_items_on_row_fields() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 7.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 3.0).unwrap();

    let mut region = PivotField::new("Region");
    region.show_empty_items = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Segment", ["Retail"]))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Region");
    assert_eq!(text(&workbook, "F1"), "Revenue");
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 10.0);
    assert_eq!(text(&workbook, "E3"), "North");
    assert_eq!(number(&workbook, "F3"), 3.0);
    assert_eq!(text(&workbook, "E4"), "West");
    assert_eq!(text(&workbook, "F4"), "");
    assert_eq!(text(&workbook, "E5"), "Grand Total");
    assert_eq!(number(&workbook, "F5"), 13.0);
}

#[test]
fn refreshes_show_empty_items_on_column_fields() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Segment").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", "Retail").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", "Online").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();

    let mut quarter = PivotField::new("Quarter");
    quarter.show_empty_items = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D3").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column(quarter)
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Segment", ["Retail"]))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "Q1");
    assert_eq!(text(&workbook, "H1"), "Q2");
    assert_eq!(text(&workbook, "I1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(number(&workbook, "G2"), 10.0);
    assert_eq!(text(&workbook, "H2"), "");
    assert_eq!(number(&workbook, "I2"), 10.0);
    assert_eq!(text(&workbook, "F3"), "Grand Total");
    assert_eq!(number(&workbook, "G3"), 10.0);
    assert_eq!(text(&workbook, "H3"), "");
    assert_eq!(number(&workbook, "I3"), 10.0);
}

#[test]
fn refreshes_layout_show_empty_rows() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 7.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 3.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.show_empty_rows = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Segment", ["Retail"]))
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Region");
    assert_eq!(text(&workbook, "F1"), "Revenue");
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 10.0);
    assert_eq!(text(&workbook, "E3"), "North");
    assert_eq!(number(&workbook, "F3"), 3.0);
    assert_eq!(text(&workbook, "E4"), "West");
    assert_eq!(text(&workbook, "F4"), "");
    assert_eq!(text(&workbook, "E5"), "Grand Total");
    assert_eq!(number(&workbook, "F5"), 13.0);
}

#[test]
fn refreshes_layout_show_empty_columns() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Segment").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", "Retail").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Q2").unwrap();
    sheet.set_cell_value("C3", "Online").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "Q3").unwrap();
    sheet.set_cell_value("C4", "Retail").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.show_empty_columns = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Segment", ["Retail"]))
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "Q1");
    assert_eq!(text(&workbook, "H1"), "Q2");
    assert_eq!(text(&workbook, "I1"), "Q3");
    assert_eq!(text(&workbook, "J1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(number(&workbook, "G2"), 10.0);
    assert_eq!(text(&workbook, "H2"), "");
    assert_eq!(number(&workbook, "I2"), 7.0);
    assert_eq!(number(&workbook, "J2"), 17.0);
    assert_eq!(text(&workbook, "F3"), "Grand Total");
    assert_eq!(number(&workbook, "G3"), 10.0);
    assert_eq!(text(&workbook, "H3"), "");
    assert_eq!(number(&workbook, "I3"), 7.0);
    assert_eq!(number(&workbook, "J3"), 17.0);
}

#[test]
fn refreshes_row_field_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "Online");
    assert_eq!(number(&workbook, "G2"), 5.0);
    assert_eq!(text(&workbook, "E3"), "East");
    assert_eq!(text(&workbook, "F3"), "Retail");
    assert_eq!(number(&workbook, "G3"), 10.0);
    assert_eq!(text(&workbook, "E4"), "East Total");
    assert_eq!(text(&workbook, "F4"), "");
    assert_eq!(number(&workbook, "G4"), 15.0);
    assert_eq!(text(&workbook, "E5"), "West");
    assert_eq!(text(&workbook, "F5"), "Retail");
    assert_eq!(number(&workbook, "G5"), 7.0);
    assert_eq!(text(&workbook, "E6"), "West Total");
    assert_eq!(number(&workbook, "G6"), 7.0);
    assert_eq!(text(&workbook, "E7"), "Grand Total");
    assert_eq!(number(&workbook, "G7"), 22.0);
}

#[test]
fn refresh_includes_hidden_items_in_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "Online").unwrap();
    sheet.set_cell_value("C5", 11.0).unwrap();

    let mut layout = tabular_layout();
    layout.subtotal_hidden_items = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .filter(PivotFilter::field_items("Segment", ["Retail"]))
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "Retail");
    assert_eq!(number(&workbook, "G2"), 10.0);
    assert_eq!(text(&workbook, "E3"), "East Total");
    assert_eq!(number(&workbook, "G3"), 15.0);
    assert_eq!(text(&workbook, "E4"), "West");
    assert_eq!(text(&workbook, "F4"), "Retail");
    assert_eq!(number(&workbook, "G4"), 7.0);
    assert_eq!(text(&workbook, "E5"), "West Total");
    assert_eq!(number(&workbook, "G5"), 18.0);
    assert_eq!(text(&workbook, "E6"), "Grand Total");
    assert_eq!(number(&workbook, "G6"), 17.0);
}

#[test]
fn refresh_inserts_blank_rows_after_row_field_items() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut region = PivotField::new("Region");
    region.insert_blank_row = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "Online");
    assert_eq!(number(&workbook, "G2"), 5.0);
    assert_eq!(text(&workbook, "E3"), "East");
    assert_eq!(text(&workbook, "F3"), "Retail");
    assert_eq!(number(&workbook, "G3"), 10.0);
    assert_eq!(text(&workbook, "E4"), "East Total");
    assert_eq!(number(&workbook, "G4"), 15.0);
    assert_eq!(text(&workbook, "E5"), "");
    assert_eq!(text(&workbook, "F5"), "");
    assert_eq!(text(&workbook, "G5"), "");
    assert_eq!(text(&workbook, "E6"), "West");
    assert_eq!(text(&workbook, "F6"), "Retail");
    assert_eq!(number(&workbook, "G6"), 7.0);
    assert_eq!(text(&workbook, "E7"), "West Total");
    assert_eq!(number(&workbook, "G7"), 7.0);
    assert_eq!(text(&workbook, "E8"), "");
    assert_eq!(text(&workbook, "F8"), "");
    assert_eq!(text(&workbook, "G8"), "");
    assert_eq!(text(&workbook, "E9"), "Grand Total");
    assert_eq!(number(&workbook, "G9"), 22.0);
}

#[test]
fn refresh_inserts_leaf_field_blank_rows_before_parent_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut segment = PivotField::new("Segment");
    segment.insert_blank_row = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .row(segment)
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "Online");
    assert_eq!(number(&workbook, "G2"), 5.0);
    assert_eq!(text(&workbook, "E3"), "");
    assert_eq!(text(&workbook, "F3"), "");
    assert_eq!(text(&workbook, "E4"), "East");
    assert_eq!(text(&workbook, "F4"), "Retail");
    assert_eq!(number(&workbook, "G4"), 10.0);
    assert_eq!(text(&workbook, "E5"), "");
    assert_eq!(text(&workbook, "F5"), "");
    assert_eq!(text(&workbook, "E6"), "East Total");
    assert_eq!(number(&workbook, "G6"), 15.0);
    assert_eq!(text(&workbook, "E7"), "West");
    assert_eq!(text(&workbook, "F7"), "Retail");
    assert_eq!(number(&workbook, "G7"), 7.0);
    assert_eq!(text(&workbook, "E8"), "");
    assert_eq!(text(&workbook, "F8"), "");
    assert_eq!(text(&workbook, "E9"), "West Total");
    assert_eq!(number(&workbook, "G9"), 7.0);
    assert_eq!(text(&workbook, "E10"), "Grand Total");
    assert_eq!(number(&workbook, "G10"), 22.0);
}

#[test]
fn refresh_inserts_page_breaks_after_row_field_items() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();
    sheet.add_row_break(20);

    let mut region = PivotField::new("Region");
    region.insert_page_break = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();
    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    let mut pivot_breaks = sheet
        .row_breaks()
        .iter()
        .filter(|break_| break_.pt)
        .map(|break_| {
            assert!(break_.man);
            assert_eq!(break_.min, 0);
            assert_eq!(break_.max, 16383);
            break_.id
        })
        .collect::<Vec<_>>();
    pivot_breaks.sort_unstable();
    assert_eq!(pivot_breaks, vec![3, 5]);

    let mut user_breaks = sheet
        .row_breaks()
        .iter()
        .filter(|break_| !break_.pt)
        .map(|break_| break_.id)
        .collect::<Vec<_>>();
    user_breaks.sort_unstable();
    assert_eq!(user_breaks, vec![20]);
}

#[test]
fn refresh_inserts_page_breaks_after_outline_top_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 7.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Outline;
    let mut region = PivotField::new("Region");
    region.insert_page_break = true;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    let mut pivot_breaks = workbook
        .worksheet(0)
        .unwrap()
        .row_breaks()
        .iter()
        .filter(|break_| break_.pt)
        .map(|break_| break_.id)
        .collect::<Vec<_>>();
    pivot_breaks.sort_unstable();
    assert_eq!(pivot_breaks, vec![3, 5]);
}

#[test]
fn refreshes_custom_row_subtotal_function() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();

    let mut region = PivotField::new("Region");
    region.subtotal = PivotSubtotal::Average;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "Online");
    assert_eq!(number(&workbook, "G2"), 20.0);
    assert_eq!(text(&workbook, "E3"), "East");
    assert_eq!(text(&workbook, "F3"), "Retail");
    assert_eq!(number(&workbook, "G3"), 10.0);
    assert_eq!(text(&workbook, "E4"), "East Total");
    assert_eq!(number(&workbook, "G4"), 15.0);
    assert_eq!(text(&workbook, "E5"), "Grand Total");
    assert_eq!(number(&workbook, "G5"), 30.0);
}

#[test]
fn refreshes_multiple_row_subtotal_functions() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();

    let region =
        PivotField::new("Region").with_subtotals([PivotSubtotal::Average, PivotSubtotal::Max]);
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E4"), "East Average");
    assert_eq!(number(&workbook, "G4"), 15.0);
    assert_eq!(text(&workbook, "E5"), "East Max");
    assert_eq!(number(&workbook, "G5"), 20.0);
    assert_eq!(text(&workbook, "E6"), "Grand Total");
    assert_eq!(number(&workbook, "G6"), 30.0);
}

#[test]
fn refreshes_row_field_subtotals_with_column_fields() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "Segment");
    assert_eq!(text(&workbook, "H1"), "Q1");
    assert_eq!(text(&workbook, "I1"), "Q2");
    assert_eq!(text(&workbook, "J1"), "Grand Total");
    assert_eq!(text(&workbook, "F4"), "East Total");
    assert_eq!(number(&workbook, "H4"), 10.0);
    assert_eq!(number(&workbook, "I4"), 5.0);
    assert_eq!(number(&workbook, "J4"), 15.0);
    assert_eq!(text(&workbook, "F6"), "West Total");
    assert_eq!(number(&workbook, "H6"), 7.0);
    assert_eq!(text(&workbook, "I6"), "");
    assert_eq!(number(&workbook, "J6"), 7.0);
    assert_eq!(text(&workbook, "F7"), "Grand Total");
    assert_eq!(number(&workbook, "H7"), 17.0);
    assert_eq!(number(&workbook, "I7"), 5.0);
    assert_eq!(number(&workbook, "J7"), 22.0);
}

#[test]
fn refreshes_column_field_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Year").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "2024").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "2024").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "2025").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "2024").unwrap();
    sheet.set_cell_value("C5", "Q1").unwrap();
    sheet.set_cell_value("D5", 3.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D5").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column("Year")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "2024 | Q1");
    assert_eq!(text(&workbook, "H1"), "2024 | Q2");
    assert_eq!(text(&workbook, "I1"), "2024 Total");
    assert_eq!(text(&workbook, "J1"), "2025 | Q1");
    assert_eq!(text(&workbook, "K1"), "2025 Total");
    assert_eq!(text(&workbook, "L1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(number(&workbook, "G2"), 10.0);
    assert_eq!(number(&workbook, "H2"), 5.0);
    assert_eq!(number(&workbook, "I2"), 15.0);
    assert_eq!(number(&workbook, "J2"), 7.0);
    assert_eq!(number(&workbook, "K2"), 7.0);
    assert_eq!(number(&workbook, "L2"), 22.0);
    assert_eq!(text(&workbook, "F3"), "West");
    assert_eq!(number(&workbook, "G3"), 3.0);
    assert_eq!(number(&workbook, "I3"), 3.0);
    assert_eq!(text(&workbook, "J3"), "");
    assert_eq!(text(&workbook, "K3"), "");
    assert_eq!(number(&workbook, "L3"), 3.0);
    assert_eq!(text(&workbook, "F4"), "Grand Total");
    assert_eq!(number(&workbook, "G4"), 13.0);
    assert_eq!(number(&workbook, "H4"), 5.0);
    assert_eq!(number(&workbook, "I4"), 18.0);
    assert_eq!(number(&workbook, "J4"), 7.0);
    assert_eq!(number(&workbook, "K4"), 7.0);
    assert_eq!(number(&workbook, "L4"), 25.0);
}

#[test]
fn refresh_writes_column_outline_levels_for_column_hierarchy() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Year").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "2024").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "2024").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "2025").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column("Year")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "G1"), "2024 | Q1");
    assert_eq!(text(&workbook, "H1"), "2024 | Q2");
    assert_eq!(text(&workbook, "I1"), "2024 Total");
    assert_eq!(text(&workbook, "J1"), "2025 | Q1");
    assert_eq!(text(&workbook, "K1"), "2025 Total");
    assert_eq!(text(&workbook, "L1"), "Grand Total");

    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.column_outline_level(5), 0);
    assert_eq!(sheet.column_outline_level(6), 1);
    assert_eq!(sheet.column_outline_level(7), 1);
    assert_eq!(sheet.column_outline_level(8), 0);
    assert_eq!(sheet.column_outline_level(9), 1);
    assert_eq!(sheet.column_outline_level(10), 0);
    assert_eq!(sheet.column_outline_level(11), 0);
}

#[test]
fn refresh_writes_column_collapsed_flags_for_collapsed_items() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Year").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "2024").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "2024").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "2025").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column(PivotField::new("Year").with_collapsed_items(["2024"]))
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    assert!(sheet.is_column_collapsed(8));
    assert!(sheet.is_column_hidden(6));
    assert!(sheet.is_column_hidden(7));
    assert!(!sheet.is_column_hidden(8));
    assert!(!sheet.is_column_hidden(10));
    assert!(!sheet.is_column_collapsed(6));
    assert!(!sheet.is_column_collapsed(7));
    assert!(!sheet.is_column_collapsed(10));
}

#[test]
fn refresh_clears_stale_pivot_column_outline_levels() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Year").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "2024").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "2024").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D3").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column("Year")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();
    assert_eq!(workbook.worksheet(0).unwrap().column_outline_level(6), 1);

    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.pivot_tables_mut()[0].layout.show_expand_collapse = false;
    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    for col in 5..=9 {
        assert_eq!(sheet.column_outline_level(col), 0);
        assert!(!sheet.is_column_collapsed(col));
    }
}

#[test]
fn refreshes_outline_layout_column_subtotals_at_top() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Year").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "2024").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "2024").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "2025").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "2024").unwrap();
    sheet.set_cell_value("C5", "Q1").unwrap();
    sheet.set_cell_value("D5", 3.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Outline;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D5").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column("Year")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "2024 Total");
    assert_eq!(text(&workbook, "H1"), "2024 | Q1");
    assert_eq!(text(&workbook, "I1"), "2024 | Q2");
    assert_eq!(text(&workbook, "J1"), "2025 Total");
    assert_eq!(text(&workbook, "K1"), "2025 | Q1");
    assert_eq!(text(&workbook, "L1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(number(&workbook, "G2"), 15.0);
    assert_eq!(number(&workbook, "H2"), 10.0);
    assert_eq!(number(&workbook, "I2"), 5.0);
    assert_eq!(number(&workbook, "J2"), 7.0);
    assert_eq!(number(&workbook, "K2"), 7.0);
    assert_eq!(number(&workbook, "L2"), 22.0);
    assert_eq!(text(&workbook, "F3"), "West");
    assert_eq!(number(&workbook, "G3"), 3.0);
    assert_eq!(number(&workbook, "H3"), 3.0);
    assert_eq!(text(&workbook, "J3"), "");
    assert_eq!(text(&workbook, "K3"), "");
    assert_eq!(number(&workbook, "L3"), 3.0);
    assert_eq!(text(&workbook, "F4"), "Grand Total");
    assert_eq!(number(&workbook, "G4"), 18.0);
    assert_eq!(number(&workbook, "H4"), 13.0);
    assert_eq!(number(&workbook, "I4"), 5.0);
    assert_eq!(number(&workbook, "J4"), 7.0);
    assert_eq!(number(&workbook, "K4"), 7.0);
    assert_eq!(number(&workbook, "L4"), 25.0);
}

#[test]
fn refreshes_custom_column_subtotal_function() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Year").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "2024").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "2024").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 20.0).unwrap();

    let mut year = PivotField::new("Year");
    year.subtotal = PivotSubtotal::Average;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D3").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column(year)
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "2024 | Q1");
    assert_eq!(text(&workbook, "H1"), "2024 | Q2");
    assert_eq!(text(&workbook, "I1"), "2024 Total");
    assert_eq!(text(&workbook, "J1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(number(&workbook, "G2"), 10.0);
    assert_eq!(number(&workbook, "H2"), 20.0);
    assert_eq!(number(&workbook, "I2"), 15.0);
    assert_eq!(number(&workbook, "J2"), 30.0);
    assert_eq!(text(&workbook, "F3"), "Grand Total");
    assert_eq!(number(&workbook, "I3"), 15.0);
    assert_eq!(number(&workbook, "J3"), 30.0);
}

#[test]
fn refreshes_multiple_column_subtotal_functions() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Year").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "2024").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "2024").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 20.0).unwrap();

    let year = PivotField::new("Year").with_subtotals([PivotSubtotal::Average, PivotSubtotal::Max]);
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D3").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column(year)
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "2024 | Q1");
    assert_eq!(text(&workbook, "H1"), "2024 | Q2");
    assert_eq!(text(&workbook, "I1"), "2024 Average");
    assert_eq!(text(&workbook, "J1"), "2024 Max");
    assert_eq!(text(&workbook, "K1"), "Grand Total");
    assert_eq!(text(&workbook, "F2"), "East");
    assert_eq!(number(&workbook, "G2"), 10.0);
    assert_eq!(number(&workbook, "H2"), 20.0);
    assert_eq!(number(&workbook, "I2"), 15.0);
    assert_eq!(number(&workbook, "J2"), 20.0);
    assert_eq!(number(&workbook, "K2"), 30.0);
    assert_eq!(text(&workbook, "F3"), "Grand Total");
    assert_eq!(number(&workbook, "I3"), 15.0);
    assert_eq!(number(&workbook, "J3"), 20.0);
    assert_eq!(number(&workbook, "K3"), 30.0);
}

#[test]
fn refreshes_row_and_column_subtotal_intersections() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
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
    sheet.set_cell_value("E3", 5.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", "2025").unwrap();
    sheet.set_cell_value("D4", "Q1").unwrap();
    sheet.set_cell_value("E4", 7.0).unwrap();
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "Retail").unwrap();
    sheet.set_cell_value("C5", "2024").unwrap();
    sheet.set_cell_value("D5", "Q1").unwrap();
    sheet.set_cell_value("E5", 3.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:E5").unwrap())
        .target_address("G1")
        .unwrap()
        .row("Region")
        .row("Segment")
        .column("Year")
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "G1"), "Region");
    assert_eq!(text(&workbook, "H1"), "Segment");
    assert_eq!(text(&workbook, "I1"), "2024 | Q1");
    assert_eq!(text(&workbook, "J1"), "2024 | Q2");
    assert_eq!(text(&workbook, "K1"), "2024 Total");
    assert_eq!(text(&workbook, "L1"), "2025 | Q1");
    assert_eq!(text(&workbook, "M1"), "2025 Total");
    assert_eq!(text(&workbook, "N1"), "Grand Total");
    assert_eq!(text(&workbook, "G4"), "East Total");
    assert_eq!(number(&workbook, "I4"), 10.0);
    assert_eq!(number(&workbook, "J4"), 5.0);
    assert_eq!(number(&workbook, "K4"), 15.0);
    assert_eq!(number(&workbook, "L4"), 7.0);
    assert_eq!(number(&workbook, "M4"), 7.0);
    assert_eq!(number(&workbook, "N4"), 22.0);
    assert_eq!(text(&workbook, "G7"), "Grand Total");
    assert_eq!(number(&workbook, "K7"), 18.0);
    assert_eq!(number(&workbook, "M7"), 7.0);
    assert_eq!(number(&workbook, "N7"), 25.0);
}

#[test]
fn refresh_respects_disabled_row_field_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "Online").unwrap();
    sheet.set_cell_value("C3", 5.0).unwrap();

    let mut region = PivotField::new("Region");
    region.subtotal = PivotSubtotal::None;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .row("Segment")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .layout(tabular_layout())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(text(&workbook, "F2"), "Online");
    assert_eq!(text(&workbook, "E3"), "East");
    assert_eq!(text(&workbook, "F3"), "Retail");
    assert_eq!(text(&workbook, "E4"), "Grand Total");
    assert_eq!(number(&workbook, "G4"), 15.0);
}

#[test]
fn refresh_respects_disabled_column_field_subtotals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Year").unwrap();
    sheet.set_cell_value("C1", "Quarter").unwrap();
    sheet.set_cell_value("D1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "2024").unwrap();
    sheet.set_cell_value("C2", "Q1").unwrap();
    sheet.set_cell_value("D2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", "2024").unwrap();
    sheet.set_cell_value("C3", "Q2").unwrap();
    sheet.set_cell_value("D3", 5.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "2025").unwrap();
    sheet.set_cell_value("C4", "Q1").unwrap();
    sheet.set_cell_value("D4", 7.0).unwrap();

    let mut year = PivotField::new("Year");
    year.subtotal = PivotSubtotal::None;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:D4").unwrap())
        .target_address("F1")
        .unwrap()
        .row("Region")
        .column(year)
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "Region");
    assert_eq!(text(&workbook, "G1"), "2024 | Q1");
    assert_eq!(text(&workbook, "H1"), "2024 | Q2");
    assert_eq!(text(&workbook, "I1"), "2025 | Q1");
    assert_eq!(text(&workbook, "J1"), "Grand Total");
    assert_eq!(number(&workbook, "J2"), 22.0);
}

#[test]
fn refresh_applies_measure_number_format() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Rate").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 0.25).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 0.5).unwrap();

    let pivot = PivotTable::builder("RatePivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .pivot_measure(PivotMeasure::new("Rate", PivotAggregate::Sum).with_number_format("0.0%"))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.formatted_value("E2").unwrap(), "25.0%");
    assert_eq!(sheet.formatted_value("E3").unwrap(), "50.0%");
    assert_eq!(sheet.formatted_value("E4").unwrap(), "75.0%");
}

#[test]
fn refresh_renders_page_fields_above_body() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Wholesale").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();
    sheet.set_cell_value("A4", "East").unwrap();
    sheet.set_cell_value("B4", "Retail").unwrap();
    sheet.set_cell_value("C4", 15.0).unwrap();

    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .page("Segment")
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::field_items("Segment", ["Retail"]))
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Segment");
    assert_eq!(text(&workbook, "F1"), "Retail");
    assert_eq!(text(&workbook, "E3"), "Region");
    assert_eq!(text(&workbook, "E4"), "East");
    assert_eq!(number(&workbook, "F4"), 25.0);
    assert_eq!(text(&workbook, "E5"), "Grand Total");
    assert_eq!(number(&workbook, "F5"), 25.0);
}

#[test]
fn refresh_respects_multiple_page_field_label_visibility() {
    fn workbook_with_page_filter(show_multiple_label: bool) -> Workbook {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Wholesale").unwrap();
        sheet.set_cell_value("C3", 20.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Online").unwrap();
        sheet.set_cell_value("C4", 7.0).unwrap();

        let mut layout = PivotLayout::default();
        layout.show_multiple_label = show_multiple_label;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .page("Segment")
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .filter(PivotFilter::field_items("Segment", ["Retail", "Wholesale"]))
            .layout(layout)
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook
    }

    let mut visible = workbook_with_page_filter(true);
    visible.refresh_pivots().unwrap();
    assert_eq!(text(&visible, "F1"), "(Multiple Items)");
    assert_eq!(number(&visible, "F4"), 30.0);

    let mut hidden = workbook_with_page_filter(false);
    hidden.refresh_pivots().unwrap();
    assert_eq!(text(&hidden, "F1"), "(All)");
    assert_eq!(number(&hidden, "F4"), 30.0);
}

#[test]
fn refresh_applies_error_caption_to_page_field_label() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet
        .set_cell_value("B2", CellValue::Error(CellError::Div0))
        .unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Retail").unwrap();
    sheet.set_cell_value("C3", 20.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.show_error = true;
    layout.error_caption = Some("ERR".to_string());
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .page("Segment")
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .filter(PivotFilter::field_items(
            "Segment",
            [PivotValue::Error(CellError::Div0)],
        ))
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Segment");
    assert_eq!(text(&workbook, "F1"), "ERR");
    assert_eq!(text(&workbook, "E4"), "East");
    assert_eq!(number(&workbook, "F4"), 10.0);
    assert_eq!(text(&workbook, "E5"), "Grand Total");
    assert_eq!(number(&workbook, "F5"), 10.0);
}

#[test]
fn refresh_wraps_page_fields_down_then_over() {
    let mut workbook = workbook_with_wrapped_page_fields(false);

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "G1"), "Segment");
    assert_eq!(text(&workbook, "H1"), "(All)");
    assert_eq!(text(&workbook, "I1"), "Country");
    assert_eq!(text(&workbook, "J1"), "(All)");
    assert_eq!(text(&workbook, "G2"), "Channel");
    assert_eq!(text(&workbook, "H2"), "(All)");
    assert_eq!(text(&workbook, "I2"), "");
    assert_eq!(text(&workbook, "J2"), "");
    assert_eq!(text(&workbook, "G4"), "Region");
    assert_eq!(text(&workbook, "G5"), "East");
    assert_eq!(number(&workbook, "H5"), 10.0);
    assert_eq!(text(&workbook, "G6"), "West");
    assert_eq!(number(&workbook, "H6"), 20.0);
}

#[test]
fn refresh_wraps_page_fields_over_then_down() {
    let mut workbook = workbook_with_wrapped_page_fields(true);

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "G1"), "Segment");
    assert_eq!(text(&workbook, "H1"), "(All)");
    assert_eq!(text(&workbook, "I1"), "Channel");
    assert_eq!(text(&workbook, "J1"), "(All)");
    assert_eq!(text(&workbook, "G2"), "Country");
    assert_eq!(text(&workbook, "H2"), "(All)");
    assert_eq!(text(&workbook, "I2"), "");
    assert_eq!(text(&workbook, "J2"), "");
    assert_eq!(text(&workbook, "G4"), "Region");
    assert_eq!(text(&workbook, "G5"), "East");
    assert_eq!(number(&workbook, "H5"), 10.0);
    assert_eq!(text(&workbook, "G6"), "West");
    assert_eq!(number(&workbook, "H6"), 20.0);
}
