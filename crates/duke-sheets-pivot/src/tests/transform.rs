use crate::test_support::*;
use pretty_assertions::assert_eq;

#[test]
fn encoded_column_push_coalesces_signed_zero() {
    let mut column = EncodedColumn::with_capacity(2);
    column.push(PivotValue::Number(-0.0));
    column.push(PivotValue::Number(0.0));

    assert_eq!(column.dictionary.len(), 1);
    assert_eq!(column.values, vec![0, 0]);
    let PivotValue::Number(value) = column.dictionary[0] else {
        panic!("expected numeric dictionary value");
    };
    assert_eq!(value.to_bits(), 0.0f64.to_bits());
    assert_eq!(column.id_for_value(&PivotValue::Number(-0.0)), Some(0));
    assert_eq!(column.id_for_value(&PivotValue::Number(0.0)), Some(0));
}

#[test]
fn encoded_column_ensure_dictionary_value_coalesces_signed_zero() {
    let mut column = EncodedColumn::with_capacity(0);

    let negative_id = column.ensure_dictionary_value(PivotValue::Number(-0.0));
    let positive_id = column.ensure_dictionary_value(PivotValue::Number(0.0));

    assert_eq!(negative_id, positive_id);
    assert_eq!(column.dictionary.len(), 1);
    let PivotValue::Number(value) = column.dictionary[0] else {
        panic!("expected numeric dictionary value");
    };
    assert_eq!(value.to_bits(), 0.0f64.to_bits());
}

#[test]
fn remapping_dictionary_coalesces_signed_zero_output() {
    let mut column = EncodedColumn::with_capacity(2);
    column.push(PivotValue::Number(-1.0));
    column.push(PivotValue::Number(1.0));

    let grouped = column.remap_dictionary(|value| match value {
        PivotValue::Number(value) if value.is_sign_negative() => PivotValue::Number(-0.0),
        PivotValue::Number(_) => PivotValue::Number(0.0),
        value => value.clone(),
    });

    assert_eq!(grouped.dictionary.len(), 1);
    assert_eq!(grouped.values, vec![0, 0]);
    let PivotValue::Number(value) = grouped.dictionary[0] else {
        panic!("expected numeric dictionary value");
    };
    assert_eq!(value.to_bits(), 0.0f64.to_bits());
    assert_eq!(grouped.id_for_value(&PivotValue::Number(-0.0)), Some(0));
    assert_eq!(grouped.id_for_value(&PivotValue::Number(0.0)), Some(0));
}

#[test]
fn encoded_column_preserves_nan_bits() {
    let mut column = EncodedColumn::with_capacity(2);
    let first_nan = f64::from_bits(0x7ff8_0000_0000_0001);
    let second_nan = f64::from_bits(0x7ff8_0000_0000_0002);
    column.push(PivotValue::Number(first_nan));
    column.push(PivotValue::Number(second_nan));

    assert_eq!(column.dictionary.len(), 2);
    assert_eq!(column.id_for_value(&PivotValue::Number(first_nan)), Some(0));
    assert_eq!(column.id_for_value(&PivotValue::Number(second_nan)), Some(1));
}

#[test]
fn remapping_grouped_column_coalesces_dictionary_ids() {
    let mut column = EncodedColumn::with_capacity(5);
    for value in ["East", "West", "South", "East", "West"] {
        column.push(PivotValue::String(value.to_string()));
    }

    let grouped = column.remap_dictionary(|value| match value {
        PivotValue::String(region) if region == "East" || region == "West" => {
            PivotValue::String("Coastal".to_string())
        }
        value => value.clone(),
    });

    assert_eq!(
        grouped.dictionary,
        vec![
            PivotValue::String("Coastal".to_string()),
            PivotValue::String("South".to_string()),
        ]
    );
    assert_eq!(grouped.values, vec![0, 0, 1, 0, 0]);
    assert_eq!(
        grouped.id_for_value(&PivotValue::String("Coastal".to_string())),
        Some(0)
    );
}

#[test]
fn refreshes_calculated_field_measure() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Units").unwrap();
    sheet.set_cell_value("C1", "Price").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 2.0).unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", 3.0).unwrap();
    sheet.set_cell_value("C3", 10.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", 7.0).unwrap();
    sheet.set_cell_value("C4", 3.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRevenue")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .calculated_field("Revenue", "=Units*Price")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Region");
    assert_eq!(text(&workbook, "F1"), "Revenue");
    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 50.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(number(&workbook, "F3"), 21.0);
    assert_eq!(text(&workbook, "E4"), "Grand Total");
    assert_eq!(number(&workbook, "F4"), 71.0);
}

#[test]
fn refreshes_calculated_field_workbook_range_references() {
    let mut workbook = Workbook::new();
    let rates_index = workbook.add_worksheet_with_name("Rates").unwrap();
    let rates = workbook.worksheet_mut(rates_index).unwrap();
    rates.set_cell_value("A1", 1.0).unwrap();
    rates.set_cell_value("A2", 0.5).unwrap();

    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Units").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 2.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", 3.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", 7.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRevenue")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_field("Revenue", "=Units*SUM(Rates!A1:A2)")
        .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_close(number(&workbook, "E2"), 7.5);
    assert_close(number(&workbook, "E3"), 10.5);
    assert_close(number(&workbook, "E4"), 18.0);

    workbook
        .worksheet_mut(rates_index)
        .unwrap()
        .set_cell_value("A2", 1.0)
        .unwrap();
    workbook.refresh_pivots().unwrap();

    assert_close(number(&workbook, "E2"), 10.0);
    assert_close(number(&workbook, "E3"), 14.0);
    assert_close(number(&workbook, "E4"), 24.0);
}

#[test]
fn refreshes_row_calculated_items() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
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
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D1"), "Region");
    assert_eq!(text(&workbook, "E1"), "Sum of Revenue");
    assert_eq!(text(&workbook, "D2"), "Combined");
    assert_eq!(number(&workbook, "E2"), 30.0);
    assert_eq!(text(&workbook, "D3"), "East");
    assert_eq!(number(&workbook, "E3"), 10.0);
    assert_eq!(text(&workbook, "D4"), "West");
    assert_eq!(number(&workbook, "E4"), 20.0);
    assert_eq!(text(&workbook, "D5"), "Grand Total");
    assert_eq!(number(&workbook, "E5"), 60.0);
}

#[test]
fn refreshes_calculated_items_across_column_buckets() {
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
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "Q2").unwrap();
    sheet.set_cell_value("C5", 3.0).unwrap();

    let pivot = PivotTable::builder("CalculatedRegionColumns")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .calculated_item("Region", "Combined", "East+West")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E1"), "Region");
    assert_eq!(text(&workbook, "F1"), "Q1");
    assert_eq!(text(&workbook, "G1"), "Q2");
    assert_eq!(text(&workbook, "H1"), "Grand Total");
    assert_eq!(text(&workbook, "E2"), "Combined");
    assert_eq!(number(&workbook, "F2"), 17.0);
    assert_eq!(number(&workbook, "G2"), 8.0);
    assert_eq!(number(&workbook, "H2"), 25.0);
    assert_eq!(text(&workbook, "E5"), "Grand Total");
    assert_eq!(number(&workbook, "F5"), 34.0);
    assert_eq!(number(&workbook, "G5"), 16.0);
    assert_eq!(number(&workbook, "H5"), 50.0);
}

#[test]
fn refreshes_column_calculated_items_with_cell_like_item_names() {
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
    sheet.set_cell_value("A5", "West").unwrap();
    sheet.set_cell_value("B5", "Q2").unwrap();
    sheet.set_cell_value("C5", 3.0).unwrap();

    let pivot = PivotTable::builder("CalculatedQuarterColumns")
        .source_range(CellRange::parse("A1:C5").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .column("Quarter")
        .calculated_item("Quarter", "H1", "Q1+Q2")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "F1"), "H1");
    assert_eq!(text(&workbook, "G1"), "Q1");
    assert_eq!(text(&workbook, "H1"), "Q2");
    assert_eq!(text(&workbook, "I1"), "Grand Total");
    assert_eq!(number(&workbook, "F2"), 15.0);
    assert_eq!(number(&workbook, "F3"), 10.0);
    assert_eq!(number(&workbook, "F4"), 25.0);
}

#[test]
fn refreshes_dependent_calculated_items_out_of_order() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "Central").unwrap();
    sheet.set_cell_value("B4", 5.0).unwrap();

    let pivot = PivotTable::builder("DependentCalculatedItems")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_item("Region", "All Regions", "\"Combined\"+Central")
        .calculated_item("Region", "Combined", "East+West")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "All Regions");
    assert_eq!(number(&workbook, "E2"), 35.0);
    assert_eq!(text(&workbook, "D3"), "Central");
    assert_eq!(number(&workbook, "E3"), 5.0);
    assert_eq!(text(&workbook, "D4"), "Combined");
    assert_eq!(number(&workbook, "E4"), 30.0);
    assert_eq!(text(&workbook, "D5"), "East");
    assert_eq!(number(&workbook, "E5"), 10.0);
    assert_eq!(text(&workbook, "D6"), "West");
    assert_eq!(number(&workbook, "E6"), 20.0);
    assert_eq!(text(&workbook, "D7"), "Grand Total");
    assert_eq!(number(&workbook, "E7"), 100.0);
}

#[test]
fn refresh_rejects_circular_calculated_items() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();

    let pivot = PivotTable::builder("CircularCalculatedItems")
        .source_range(CellRange::parse("A1:B2").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_item("Region", "Calc One", "\"Calc Two\"+1")
        .calculated_item("Region", "Calc Two", "\"Calc One\"+1")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let error = workbook.refresh_pivots().unwrap_err().to_string();
    assert!(
        error.contains("calculated items contain a circular reference"),
        "{error}"
    );
}

#[test]
fn refresh_rejects_cell_like_calculated_item_self_reference() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();

    let pivot = PivotTable::builder("CellLikeCalculatedItemSelfReference")
        .source_range(CellRange::parse("A1:B2").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .calculated_item("Region", "H1", "H1+1")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let error = workbook.refresh_pivots().unwrap_err().to_string();
    assert!(
        error.contains("calculated item H1 references itself"),
        "{error}"
    );
}

#[test]
fn refreshes_sequential_calculated_fields_with_structured_refs() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Gross Sales").unwrap();
    sheet.set_cell_value("C1", "Rate").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 100.0).unwrap();
    sheet.set_cell_value("C2", 0.1).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", 50.0).unwrap();
    sheet.set_cell_value("C3", 0.2).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", 80.0).unwrap();
    sheet.set_cell_value("C4", 0.25).unwrap();

    let pivot = PivotTable::builder("CalculatedCommission")
        .source_range(CellRange::parse("A1:C4").unwrap())
        .target_address("E1")
        .unwrap()
        .row("Region")
        .calculated_field("Commission", "=[Gross Sales]*Rate")
        .calculated_field("CommissionWithFee", "=Commission+1")
        .named_measure("CommissionWithFee", PivotAggregate::Sum, "Commission")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 22.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(number(&workbook, "F3"), 21.0);
    assert_eq!(text(&workbook, "E4"), "Grand Total");
    assert_eq!(number(&workbook, "F4"), 43.0);
}

#[test]
fn refreshes_table_qualified_calculated_fields_with_structured_refs() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Units").unwrap();
    sheet.set_cell_value("C1", "Price").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 2.0).unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", 3.0).unwrap();
    sheet.set_cell_value("C3", 4.0).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", 7.0).unwrap();
    sheet.set_cell_value("C4", 3.0).unwrap();

    let mut table = Table::new(1, "SalesData", CellRange::parse("A1:C4").unwrap());
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

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 32.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(number(&workbook, "F3"), 21.0);
    assert_eq!(text(&workbook, "E4"), "Grand Total");
    assert_eq!(number(&workbook, "F4"), 53.0);
}

#[test]
fn refreshes_table_qualified_calculated_fields_with_escaped_structured_refs() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Gross Sales").unwrap();
    sheet.set_cell_value("C1", "Rate").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 100.0).unwrap();
    sheet.set_cell_value("C2", 0.1).unwrap();
    sheet.set_cell_value("A3", "East").unwrap();
    sheet.set_cell_value("B3", 50.0).unwrap();
    sheet.set_cell_value("C3", 0.2).unwrap();
    sheet.set_cell_value("A4", "West").unwrap();
    sheet.set_cell_value("B4", 80.0).unwrap();
    sheet.set_cell_value("C4", 0.25).unwrap();

    let mut table = Table::new(1, "SalesData", CellRange::parse("A1:C4").unwrap());
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

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "E2"), "East");
    assert_eq!(number(&workbook, "F2"), 20.0);
    assert_eq!(text(&workbook, "E3"), "West");
    assert_eq!(number(&workbook, "F3"), 20.0);
    assert_eq!(text(&workbook, "E4"), "Grand Total");
    assert_eq!(number(&workbook, "F4"), 40.0);
}

#[test]
fn refreshes_numeric_grouping() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Amount").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", 2.0).unwrap();
    sheet.set_cell_value("B2", 1.0).unwrap();
    sheet.set_cell_value("A3", 7.0).unwrap();
    sheet.set_cell_value("B3", 2.0).unwrap();
    sheet.set_cell_value("A4", 12.0).unwrap();
    sheet.set_cell_value("B4", 3.0).unwrap();
    sheet.set_cell_value("A5", 17.0).unwrap();
    sheet.set_cell_value("B5", 4.0).unwrap();
    sheet.set_cell_value("A6", 25.0).unwrap();
    sheet.set_cell_value("B6", 5.0).unwrap();

    let pivot = PivotTable::builder("GroupedAmounts")
        .source_range(CellRange::parse("A1:B6").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Amount")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Number {
            field: "Amount".into(),
            start: Some(0.0),
            end: None,
            interval: 10.0,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(number(&workbook, "D2"), 0.0);
    assert_eq!(number(&workbook, "E2"), 3.0);
    assert_eq!(number(&workbook, "D3"), 10.0);
    assert_eq!(number(&workbook, "E3"), 7.0);
    assert_eq!(number(&workbook, "D4"), 20.0);
    assert_eq!(number(&workbook, "E4"), 5.0);
    assert_eq!(text(&workbook, "D5"), "Grand Total");
    assert_eq!(number(&workbook, "E5"), 15.0);
}

#[test]
fn refreshes_manual_item_grouping() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", 7.0).unwrap();
    sheet.set_cell_value("A5", "South").unwrap();
    sheet.set_cell_value("B5", 8.0).unwrap();
    sheet.set_cell_value("A6", "Central").unwrap();
    sheet.set_cell_value("B6", 5.0).unwrap();
    sheet.set_cell_value("A7", "East").unwrap();
    sheet.set_cell_value("B7", 3.0).unwrap();

    let pivot = PivotTable::builder("GroupedRegions")
        .source_range(CellRange::parse("A1:B7").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![
                PivotManualGroup::new("Coastal", ["East", "West"]),
                PivotManualGroup::new("Inland", ["North", "South"]),
            ],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "Central");
    assert_eq!(number(&workbook, "E2"), 5.0);
    assert_eq!(text(&workbook, "D3"), "Coastal");
    assert_eq!(number(&workbook, "E3"), 33.0);
    assert_eq!(text(&workbook, "D4"), "Inland");
    assert_eq!(number(&workbook, "E4"), 15.0);
    assert_eq!(text(&workbook, "D5"), "Grand Total");
    assert_eq!(number(&workbook, "E5"), 53.0);
}

#[test]
fn refreshes_multi_unit_date_grouping_hierarchy() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Date").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet
        .set_cell_value("A2", date_to_serial(2024, 1, 15, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet
        .set_cell_value("A3", date_to_serial(2024, 1, 20, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B3", 5.0).unwrap();
    sheet
        .set_cell_value("A4", date_to_serial(2024, 2, 1, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B4", 7.0).unwrap();
    sheet
        .set_cell_value("A5", date_to_serial(2025, 1, 1, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B5", 11.0).unwrap();

    let pivot = PivotTable::builder("GroupedDates")
        .source_range(CellRange::parse("A1:B5").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Date")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Date {
            field: "Date".into(),
            units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D1"), "Row Labels");
    assert_eq!(text(&workbook, "E1"), "Sum of Revenue");
    assert_eq!(text(&workbook, "D2"), "2024");
    assert_eq!(text(&workbook, "E2"), "");
    assert_eq!(text(&workbook, "D3"), "1");
    assert_eq!(number(&workbook, "E3"), 15.0);
    assert_eq!(text(&workbook, "D4"), "2");
    assert_eq!(number(&workbook, "E4"), 7.0);
    assert_eq!(text(&workbook, "D5"), "2024 Total");
    assert_eq!(number(&workbook, "E5"), 22.0);
    assert_eq!(text(&workbook, "D6"), "2025");
    assert_eq!(text(&workbook, "D7"), "1");
    assert_eq!(number(&workbook, "E7"), 11.0);
    assert_eq!(text(&workbook, "D8"), "2025 Total");
    assert_eq!(number(&workbook, "E8"), 11.0);
    assert_eq!(text(&workbook, "D9"), "Grand Total");
    assert_eq!(number(&workbook, "E9"), 33.0);
}

#[test]
fn refreshes_single_unit_date_grouping() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Date").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet
        .set_cell_value("A2", date_to_serial(2024, 1, 15, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet
        .set_cell_value("A3", date_to_serial(2024, 1, 20, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B3", 5.0).unwrap();
    sheet
        .set_cell_value("A4", date_to_serial(2024, 2, 1, DateSystem::Date1900))
        .unwrap();
    sheet.set_cell_value("B4", 7.0).unwrap();

    let pivot = PivotTable::builder("GroupedDates")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Date")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Date {
            field: "Date".into(),
            units: vec![PivotDateGroupUnit::Months],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook.refresh_pivots().unwrap();

    assert_eq!(text(&workbook, "D2"), "1");
    assert_eq!(number(&workbook, "E2"), 15.0);
    assert_eq!(text(&workbook, "D3"), "2");
    assert_eq!(number(&workbook, "E3"), 7.0);
    assert_eq!(text(&workbook, "D4"), "Grand Total");
    assert_eq!(number(&workbook, "E4"), 22.0);
}
