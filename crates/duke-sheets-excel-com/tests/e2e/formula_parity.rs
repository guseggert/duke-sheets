use std::path::{Path, PathBuf};

use crate::{ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_excel_com::{BridgeError, ChainStep, Workbook};
use excel_com_protocol::{ResponseData, SheetRef};
use serde_json::json;

const VM_FIXTURE_PATH: &str = r"C:\temp\formula_parity.xlsx";
const HOST_FIXTURE_PATH: &str = "/tmp/duke-sheets-excel/formula_parity.xlsx";
const REPO_FIXTURE_RELATIVE_PATH: &str = "data/formula-parity.xlsx";
const EXPECTED_PARITY_CASE_COUNT: usize = 911;

#[derive(Clone, Copy)]
enum FormulaKind {
    Formula,
    Formula2,
}

struct FormulaCase {
    id: String,
    label: String,
    formula: String,
    expected_type: &'static str,
    kind: FormulaKind,
}

fn case(id: &str, label: &str, formula: &str, expected_type: &'static str) -> FormulaCase {
    FormulaCase {
        id: id.to_string(),
        label: label.to_string(),
        formula: formula.to_string(),
        expected_type,
        kind: FormulaKind::Formula,
    }
}

fn case2(id: &str, label: &str, formula: &str, expected_type: &'static str) -> FormulaCase {
    FormulaCase {
        id: id.to_string(),
        label: label.to_string(),
        formula: formula.to_string(),
        expected_type,
        kind: FormulaKind::Formula2,
    }
}

#[test]
fn generate_formula_parity_spreadsheet() {
    let bridge = excel_bridge();
    let mut fixture = temp_fixture();
    fixture.host_path = PathBuf::from(HOST_FIXTURE_PATH);
    fixture.vm_path = VM_FIXTURE_PATH.to_string();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();

        let mut wb = excel.create_workbook().expect("create workbook");
        rename_sheet(&excel, wb.handle(), 0, "Data").expect("rename Sheet1 to Data");
        add_worksheet_after(&excel, wb.handle(), 0, "Tests").expect("add Tests worksheet");

        wb.set_active_sheet_index(0);
        populate_data_sheet(&wb).expect("populate Data sheet");

        wb.set_active_sheet_index(1);
        populate_tests_sheet(&mut wb).expect("populate Tests sheet");

        excel.recalculate().expect("recalculate workbook");
        wb.save(&fixture.vm_path).expect("save workbook");
        wb.close().expect("close workbook");
    }

    pull_file_from_vm(&fixture);
    copy_fixture_into_repo(&fixture.host_path).expect("copy fixture into repo data directory");
}

fn rename_sheet(
    excel: &duke_sheets_excel_com::ExcelBridge,
    workbook_handle: u64,
    index: u32,
    name: &str,
) -> Result<(), BridgeError> {
    excel.set(
        workbook_handle,
        vec![SheetRef::Index(index).to_chain_step()],
        "Name",
        serde_json::Value::from(name),
    )
}

fn add_worksheet_after(
    excel: &duke_sheets_excel_com::ExcelBridge,
    workbook_handle: u64,
    after_index: u32,
    name: &str,
) -> Result<(), BridgeError> {
    let after_handle = excel.navigate(
        workbook_handle,
        vec![SheetRef::Index(after_index).to_chain_step()],
    )?;
    let response = excel.invoke(
        workbook_handle,
        vec![ChainStep::Property("Worksheets".to_string())],
        "Add",
        vec![serde_json::Value::Null, json!({"$ref": after_handle})],
    );
    let _ = excel.release(after_handle);

    let sheet_handle = match response? {
        Some(ResponseData::Handle { handle }) => handle,
        _ => return Err(BridgeError::ExpectedHandle),
    };

    let rename_result = excel.set(sheet_handle, vec![], "Name", serde_json::Value::from(name));
    let release_result = excel.release(sheet_handle);
    rename_result?;
    release_result
}

fn populate_data_sheet(wb: &Workbook<'_>) -> Result<(), BridgeError> {
    for (offset, month) in (1..=12).enumerate() {
        wb.set_cell_value(&cell_addr(1, offset as u32 + 1), month as f64)?;
    }

    for (offset, value) in (1..=12).map(|n| n * 10).enumerate() {
        wb.set_cell_value(&cell_addr(2, offset as u32 + 1), value as f64)?;
        wb.set_cell_value(&cell_addr(offset as u32 + 3, 1), value as f64)?;
        wb.set_cell_value(&cell_addr(offset as u32 + 2, 14), (130 - value) as f64)?;
    }

    let fruit_prices = [
        ("apple", 1.50),
        ("banana", 2.00),
        ("cherry", 0.75),
        ("date", 3.00),
        ("elderberry", 5.50),
        ("fig", 1.25),
        ("grape", 2.50),
        ("honeydew", 4.00),
        ("kiwi", 3.50),
        ("lemon", 0.50),
    ];
    for (offset, (fruit, price)) in fruit_prices.iter().enumerate() {
        let row = offset as u32 + 16;
        wb.set_cell_value(&cell_addr(row, 1), *fruit)?;
        wb.set_cell_value(&cell_addr(row, 2), *price)?;
    }

    wb.set_cell_value("D16", "")?;
    wb.set_cell_value("E16", "North")?;
    wb.set_cell_value("F16", "South")?;
    wb.set_cell_value("G16", "West")?;
    let quarter_rows = [
        (17, "Q1", [101.0, 102.0, 103.0]),
        (18, "Q2", [201.0, 202.0, 203.0]),
        (19, "Q3", [301.0, 302.0, 303.0]),
        (20, "Q4", [401.0, 402.0, 403.0]),
    ];
    for (row, quarter, values) in quarter_rows {
        wb.set_cell_value(&cell_addr(row, 4), quarter)?;
        for (index, value) in values.into_iter().enumerate() {
            wb.set_cell_value(&cell_addr(row, index as u32 + 5), value)?;
        }
    }

    wb.set_cell_value("M16", "item")?;
    wb.set_cell_value("N16", "value")?;
    let duplicate_lookup = [
        ("apple", 1.0),
        ("banana", 2.0),
        ("apple", 3.0),
        ("date", 4.0),
    ];
    for (offset, (item, value)) in duplicate_lookup.iter().enumerate() {
        let row = offset as u32 + 17;
        wb.set_cell_value(&cell_addr(row, 13), *item)?;
        wb.set_cell_value(&cell_addr(row, 14), *value)?;
    }

    wb.set_cell_value("P16", "score")?;
    wb.set_cell_value("Q16", "grade")?;
    let grades = [
        (0.0, "F"),
        (60.0, "D"),
        (70.0, "C"),
        (80.0, "B"),
        (90.0, "A"),
    ];
    for (offset, (score, grade)) in grades.iter().enumerate() {
        let row = offset as u32 + 17;
        wb.set_cell_value(&cell_addr(row, 16), *score)?;
        wb.set_cell_value(&cell_addr(row, 17), *grade)?;
    }

    wb.set_cell_value("A28", "name")?;
    wb.set_cell_value("B28", "category")?;
    wb.set_cell_value("C28", "amount")?;
    let mixed_rows = [
        ("alice", "fruit", 10.0),
        ("bob", "fruit", 20.0),
        ("carl", "veg", 15.0),
        ("dina", "fruit", 25.0),
        ("erin", "veg", 30.0),
        ("fiona", "grain", 12.0),
        ("george", "fruit", 18.0),
        ("hannah", "veg", 22.0),
        ("ian", "fruit", 35.0),
    ];
    for (offset, (name, category, amount)) in mixed_rows.iter().enumerate() {
        let row = offset as u32 + 29;
        wb.set_cell_value(&cell_addr(row, 1), *name)?;
        wb.set_cell_value(&cell_addr(row, 2), *category)?;
        wb.set_cell_value(&cell_addr(row, 3), *amount)?;
    }

    let pi_digits = [3.0, 1.0, 4.0, 1.0, 5.0, 9.0, 2.0, 6.0, 5.0, 3.0, 5.0, 8.0];
    for (offset, value) in pi_digits.iter().enumerate() {
        wb.set_cell_value(&cell_addr(offset as u32 + 1, 19), *value)?;
    }

    wb.set_cell_value("A40", "hello")?;
    wb.set_cell_value("C40", 42.0)?;
    wb.set_cell_value("E40", "world")?;

    wb.set_cell_value("A43", 200000.0)?;
    wb.set_cell_value("B43", 0.06)?;
    wb.set_cell_value("C43", 30.0)?;
    wb.set_cell_value("D43", 12.0)?;
    let cash_flows = [-1000.0, 300.0, 420.0, 680.0, 250.0];
    for (offset, value) in cash_flows.iter().enumerate() {
        wb.set_cell_value(&cell_addr(offset as u32 + 44, 1), *value)?;
    }

    wb.set_cell_value("A50", 42.0)?;
    wb.set_cell_value("A51", "hello")?;
    wb.set_cell_value("A52", true)?;
    wb.set_cell_formula("A54", "=1/0")?;

    wb.set_cell_value("A56", 0.1)?;
    wb.set_cell_value("B56", 0.2)?;
    wb.set_cell_value("A57", 1e15)?;
    wb.set_cell_value("B57", 1.0)?;

    wb.set_cell_value("A60", "Hello World")?;
    wb.set_cell_value("A61", "mississippi")?;
    wb.set_cell_value("A62", "2024-01-15")?;

    for offset in 0..20u32 {
        let row = 70 + offset;
        let id = offset + 1;
        wb.set_cell_value(&cell_addr(row, 1), id as f64)?;
        wb.set_cell_value(&cell_addr(row, 2), (id * 100) as f64)?;
        wb.set_cell_value(
            &cell_addr(row, 3),
            if id % 2 == 1 { "alpha" } else { "beta" },
        )?;
    }

    wb.set_cell_value("A92", 50000.0)?;
    wb.set_cell_value("B92", 0.045)?;
    wb.set_cell_value("C92", 15.0)?;
    wb.set_cell_value("D92", 12.0)?;
    let alt_cash_flows = [-5000.0, 1500.0, 1800.0, 2100.0, 900.0];
    for (offset, value) in alt_cash_flows.iter().enumerate() {
        wb.set_cell_value(&cell_addr(offset as u32 + 93, 1), *value)?;
    }

    let stats_values = [12.0, 15.0, 22.0, 22.0, 25.0, 28.0, 31.0, 35.0, 42.0, 50.0];
    for (offset, value) in stats_values.iter().enumerate() {
        wb.set_cell_value(&cell_addr(offset as u32 + 100, 1), *value)?;
    }

    wb.set_cell_formula("A112", "=DATE(2024,3,15)")?;
    wb.set_cell_formula("A113", "=DATE(2024,12,25)")?;

    let matrix_values = [
        (115, [1.0, 2.0, 3.0]),
        (116, [4.0, 5.0, 6.0]),
        (117, [7.0, 8.0, 9.0]),
    ];
    for (row, values) in matrix_values {
        for (index, value) in values.into_iter().enumerate() {
            wb.set_cell_value(&cell_addr(row, index as u32 + 1), value)?;
        }
    }

    wb.set_cell_value("A120", 0.0)?;
    wb.set_cell_formula("A121", "=PI()/6")?;
    wb.set_cell_formula("A122", "=PI()/4")?;
    wb.set_cell_formula("A123", "=PI()/2")?;
    wb.set_cell_formula("A124", "=PI()")?;

    wb.set_cell_value("A126", "Hello World 123")?;
    wb.set_cell_value("A127", "  spaces  ")?;
    wb.set_cell_value("A128", "ABCDEFGHIJ")?;
    wb.set_cell_value("A129", "test@email.com")?;
    wb.set_cell_value("A130", "12345.6789")?;

    for (offset, value) in [10.0, 20.0, 30.0, 40.0, 50.0].iter().enumerate() {
        wb.set_cell_value(&cell_addr(offset as u32 + 132, 1), *value)?;
    }

    let ifs_rows = [
        ("alice", 85.0, "math"),
        ("bob", 92.0, "math"),
        ("alice", 78.0, "science"),
        ("bob", 88.0, "science"),
        ("alice", 95.0, "math"),
    ];
    for (offset, (name, score, category)) in ifs_rows.iter().enumerate() {
        let row = offset as u32 + 138;
        wb.set_cell_value(&cell_addr(row, 1), *name)?;
        wb.set_cell_value(&cell_addr(row, 2), *score)?;
        wb.set_cell_value(&cell_addr(row, 3), *category)?;
    }

    wb.set_cell_value("A144", 100.0)?;
    wb.set_cell_value("A145", 200.0)?;
    wb.set_cell_formula("A146", "=1/0")?;
    wb.set_cell_value("A147", 400.0)?;
    wb.set_cell_value("A148", 500.0)?;

    wb.set_cell_value("A150", "Name")?;
    wb.set_cell_value("B150", "Category")?;
    wb.set_cell_value("C150", "Amount")?;
    wb.set_cell_value("D150", "Region")?;
    let database_rows = [
        ("Alice", "Sales", 100.0, "North"),
        ("Bob", "Engineering", 200.0, "South"),
        ("Alice", "Sales", 150.0, "South"),
        ("Carol", "Engineering", 300.0, "North"),
        ("Bob", "Sales", 120.0, "North"),
        ("Alice", "Engineering", 180.0, "South"),
        ("Carol", "Sales", 90.0, "North"),
        ("Bob", "Engineering", 250.0, "South"),
        ("Alice", "Sales", 110.0, "North"),
        ("Carol", "Engineering", 220.0, "South"),
    ];
    for (offset, (name, category, amount, region)) in database_rows.iter().enumerate() {
        let row = 151 + offset as u32;
        wb.set_cell_value(&cell_addr(row, 1), *name)?;
        wb.set_cell_value(&cell_addr(row, 2), *category)?;
        wb.set_cell_value(&cell_addr(row, 3), *amount)?;
        wb.set_cell_value(&cell_addr(row, 4), *region)?;
    }
    wb.set_cell_value("A162", "Name")?;
    wb.set_cell_value("B162", "Category")?;
    wb.set_cell_value("A163", "Alice")?;
    wb.set_cell_value("B163", "Sales")?;
    wb.set_cell_value("F162", "Name")?;
    wb.set_cell_value("G162", "Category")?;
    wb.set_cell_value("F163", "Carol")?;
    wb.set_cell_value("G163", "Sales")?;

    for i in 0u32..50 {
        let row = 170 + i;
        wb.set_cell_value(&cell_addr(row, 1), (i + 1) as f64)?;
        wb.set_cell_value(&cell_addr(row, 2), ((i + 1) * 100) as f64)?;
        wb.set_cell_value(
            &cell_addr(row, 3),
            if i % 3 == 0 {
                "alpha"
            } else if i % 3 == 1 {
                "beta"
            } else {
                "gamma"
            },
        )?;
    }

    for (offset, value) in [0.5, 1.0, 2.0, 5.0, 10.0].iter().enumerate() {
        wb.set_cell_value(&cell_addr(222 + offset as u32, 1), *value)?;
    }
    for (offset, value) in [0.0, 0.25, 0.5, 0.75, 1.0].iter().enumerate() {
        wb.set_cell_value(&cell_addr(222 + offset as u32, 2), *value)?;
    }

    wb.set_cell_value("A233", 1000.0)?;
    wb.set_cell_value("B233", 0.05)?;
    wb.set_cell_value("C233", 0.04)?;
    wb.set_cell_value("A234", 44927.0)?;
    wb.set_cell_value("B234", 46387.0)?;
    wb.set_cell_value("A235", 10000.0)?;
    wb.set_cell_value("B235", 1000.0)?;
    wb.set_cell_value("C235", 10.0)?;

    for i in 0u32..10 {
        wb.set_cell_value(&cell_addr(240 + i, 1), ((i + 1) * 7) as f64)?;
    }

    for i in 0u32..10 {
        let x = i as f64 * 0.7;
        wb.set_cell_value(&cell_addr(252 + i, 1), x)?;
    }

    for (offset, value) in [1.0, 2.0, 3.0, 4.0, 5.0].iter().enumerate() {
        wb.set_cell_value(&cell_addr(264, offset as u32 + 1), *value)?;
    }
    for (offset, value) in [6.0, 7.0, 8.0, 9.0, 10.0].iter().enumerate() {
        wb.set_cell_value(&cell_addr(265, offset as u32 + 1), *value)?;
    }
    for (offset, value) in [11.0, 12.0, 13.0, 14.0, 15.0].iter().enumerate() {
        wb.set_cell_value(&cell_addr(266, offset as u32 + 1), *value)?;
    }
    for (offset, value) in ["x", "y", "z"].iter().enumerate() {
        wb.set_cell_value(&cell_addr(267, offset as u32 + 1), *value)?;
    }
    for (offset, value) in ["a", "b", "c"].iter().enumerate() {
        wb.set_cell_value(&cell_addr(268, offset as u32 + 1), *value)?;
    }

    Ok(())
}

fn populate_tests_sheet(wb: &mut Workbook<'_>) -> Result<(), BridgeError> {
    wb.set_active_sheet_name("Tests");

    wb.set_cell_value("A1", "Test ID")?;
    wb.set_cell_value("B1", "Formula Label")?;
    wb.set_cell_value("C1", "Result (Formula)")?;
    wb.set_cell_value("D1", "Expected Type")?;

    let mut cases = static_formula_cases();
    append_tier3_cache_cases(&mut cases);
    append_tier4_validation_cases(&mut cases);

    assert_eq!(
        cases.len(),
        EXPECTED_PARITY_CASE_COUNT,
        "unexpected parity case count"
    );

    let mut row = 2u32;
    for case in &cases {
        wb.set_cell_value(&cell_addr(row, 1), case.id.as_str())?;
        wb.set_cell_value(&cell_addr(row, 2), case.label.as_str())?;
        wb.set_cell_value(&cell_addr(row, 4), case.expected_type)?;
        match case.kind {
            FormulaKind::Formula => {
                wb.set_cell_formula(&cell_addr(row, 3), case.formula.as_str())?
            }
            FormulaKind::Formula2 => {
                wb.set_cell_formula2(&cell_addr(row, 3), case.formula.as_str())?
            }
        }
        row += 1;
    }

    Ok(())
}

fn static_formula_cases() -> Vec<FormulaCase> {
    vec![
        // Existing cases
        case(
            "INDEX_2arg_single_row_3",
            "INDEX(Data!A2:L2,3)",
            "=INDEX(Data!A2:L2,3)",
            "number",
        ),
        case(
            "INDEX_2arg_single_col_4",
            "INDEX(Data!A3:A14,4)",
            "=INDEX(Data!A3:A14,4)",
            "number",
        ),
        case(
            "INDEX_3arg_matrix_q2_west",
            "INDEX(Data!E17:G20,2,3)",
            "=INDEX(Data!E17:G20,2,3)",
            "number",
        ),
        case(
            "INDEX_3arg_row_vector_col_5",
            "INDEX(Data!A2:L2,1,5)",
            "=INDEX(Data!A2:L2,1,5)",
            "number",
        ),
        case(
            "INDEX_MATCH_single_row_60",
            "INDEX(Data!A2:L2,MATCH(60,Data!A2:L2,0))",
            "=INDEX(Data!A2:L2,MATCH(60,Data!A2:L2,0))",
            "number",
        ),
        case(
            "MATCH_exact_numeric_50",
            "MATCH(50,Data!A2:L2,0)",
            "=MATCH(50,Data!A2:L2,0)",
            "number",
        ),
        case(
            "MATCH_approx_ascending_55",
            "MATCH(55,Data!A2:L2,1)",
            "=MATCH(55,Data!A2:L2,1)",
            "number",
        ),
        case(
            "MATCH_approx_descending_55",
            "MATCH(55,Data!N2:N13,-1)",
            "=MATCH(55,Data!N2:N13,-1)",
            "number",
        ),
        case(
            "MATCH_exact_string_cherry",
            "MATCH(\"cherry\",Data!A16:A25,0)",
            "=MATCH(\"cherry\",Data!A16:A25,0)",
            "number",
        ),
        case(
            "MATCH_not_found_orange",
            "MATCH(\"orange\",Data!A16:A25,0)",
            "=MATCH(\"orange\",Data!A16:A25,0)",
            "error",
        ),
        case(
            "VLOOKUP_exact_banana_price",
            "VLOOKUP(\"banana\",Data!A16:B25,2,FALSE)",
            "=VLOOKUP(\"banana\",Data!A16:B25,2,FALSE)",
            "number",
        ),
        case(
            "VLOOKUP_approx_grade_88",
            "VLOOKUP(88,Data!P17:Q21,2,TRUE)",
            "=VLOOKUP(88,Data!P17:Q21,2,TRUE)",
            "string",
        ),
        case(
            "VLOOKUP_diff_col_q3_west",
            "VLOOKUP(\"Q3\",Data!D17:G20,4,FALSE)",
            "=VLOOKUP(\"Q3\",Data!D17:G20,4,FALSE)",
            "number",
        ),
        case(
            "VLOOKUP_not_found_orange",
            "VLOOKUP(\"orange\",Data!A16:B25,2,FALSE)",
            "=VLOOKUP(\"orange\",Data!A16:B25,2,FALSE)",
            "error",
        ),
        case(
            "HLOOKUP_exact_month_4",
            "HLOOKUP(4,Data!A1:L2,2,FALSE)",
            "=HLOOKUP(4,Data!A1:L2,2,FALSE)",
            "number",
        ),
        case(
            "HLOOKUP_approx_month_5_5",
            "HLOOKUP(5.5,Data!A1:L2,2,TRUE)",
            "=HLOOKUP(5.5,Data!A1:L2,2,TRUE)",
            "number",
        ),
        case(
            "HLOOKUP_not_found_month_13",
            "HLOOKUP(13,Data!A1:L2,2,FALSE)",
            "=HLOOKUP(13,Data!A1:L2,2,FALSE)",
            "error",
        ),
        case(
            "SUMIFS_single_category_fruit",
            "SUMIFS(Data!C29:C37,Data!B29:B37,\"fruit\")",
            "=SUMIFS(Data!C29:C37,Data!B29:B37,\"fruit\")",
            "number",
        ),
        case(
            "SUMIFS_multi_category_and_name",
            "SUMIFS(Data!C29:C37,Data!B29:B37,\"fruit\",Data!A29:A37,\"*a*\")",
            "=SUMIFS(Data!C29:C37,Data!B29:B37,\"fruit\",Data!A29:A37,\"*a*\")",
            "number",
        ),
        case(
            "COUNTIFS_single_category_veg",
            "COUNTIFS(Data!B29:B37,\"veg\")",
            "=COUNTIFS(Data!B29:B37,\"veg\")",
            "number",
        ),
        case(
            "COUNTIFS_multi_wildcard_i",
            "COUNTIFS(Data!B29:B37,\"fruit\",Data!A29:A37,\"*i*\")",
            "=COUNTIFS(Data!B29:B37,\"fruit\",Data!A29:A37,\"*i*\")",
            "number",
        ),
        case(
            "IF_basic_false",
            "IF(Data!A2>50,\"big\",\"small\")",
            "=IF(Data!A2>50,\"big\",\"small\")",
            "string",
        ),
        case(
            "IF_basic_true",
            "IF(Data!L2>50,\"big\",\"small\")",
            "=IF(Data!L2>50,\"big\",\"small\")",
            "string",
        ),
        case(
            "IF_nested_medium",
            "IF(Data!A2>50,\"big\",IF(Data!A2>5,\"medium\",\"small\"))",
            "=IF(Data!A2>50,\"big\",IF(Data!A2>5,\"medium\",\"small\"))",
            "string",
        ),
        case(
            "IFS_multi_condition_medium",
            "IFS(Data!A2>50,\"big\",Data!A2>5,\"medium\",TRUE,\"small\")",
            "=IFS(Data!A2>50,\"big\",Data!A2>5,\"medium\",TRUE,\"small\")",
            "string",
        ),
        case(
            "SWITCH_numeric_two",
            "SWITCH(2,1,\"one\",2,\"two\",\"other\")",
            "=SWITCH(2,1,\"one\",2,\"two\",\"other\")",
            "string",
        ),
        case(
            "SWITCH_default_other",
            "SWITCH(\"kiwi\",\"apple\",\"A\",\"banana\",\"B\",\"other\")",
            "=SWITCH(\"kiwi\",\"apple\",\"A\",\"banana\",\"B\",\"other\")",
            "string",
        ),
        case(
            "SUM_row_values",
            "SUM(Data!A2:L2)",
            "=SUM(Data!A2:L2)",
            "number",
        ),
        case(
            "AVERAGE_row_values",
            "AVERAGE(Data!A2:L2)",
            "=AVERAGE(Data!A2:L2)",
            "number",
        ),
        case(
            "MIN_col_values",
            "MIN(Data!A3:A14)",
            "=MIN(Data!A3:A14)",
            "number",
        ),
        case(
            "MAX_col_values",
            "MAX(Data!A3:A14)",
            "=MAX(Data!A3:A14)",
            "number",
        ),
        case(
            "COUNT_col_values",
            "COUNT(Data!A3:A14)",
            "=COUNT(Data!A3:A14)",
            "number",
        ),
        case(
            "COUNTA_lookup_keys",
            "COUNTA(Data!A16:A25)",
            "=COUNTA(Data!A16:A25)",
            "number",
        ),
        case(
            "SUMPRODUCT_months_and_values",
            "SUMPRODUCT(Data!A1:L1,Data!A2:L2)",
            "=SUMPRODUCT(Data!A1:L1,Data!A2:L2)",
            "number",
        ),
        case(
            "ROUND_basic",
            "ROUND(1.2345,2)",
            "=ROUND(1.2345,2)",
            "number",
        ),
        case(
            "ROUNDUP_basic",
            "ROUNDUP(1.231,2)",
            "=ROUNDUP(1.231,2)",
            "number",
        ),
        case(
            "ROUNDDOWN_basic",
            "ROUNDDOWN(1.239,2)",
            "=ROUNDDOWN(1.239,2)",
            "number",
        ),
        case(
            "LEFT_banana_3",
            "LEFT(\"banana\",3)",
            "=LEFT(\"banana\",3)",
            "string",
        ),
        case(
            "RIGHT_banana_2",
            "RIGHT(\"banana\",2)",
            "=RIGHT(\"banana\",2)",
            "string",
        ),
        case(
            "MID_elderberry_3_5",
            "MID(\"elderberry\",3,5)",
            "=MID(\"elderberry\",3,5)",
            "string",
        ),
        case(
            "LEN_honeydew",
            "LEN(\"honeydew\")",
            "=LEN(\"honeydew\")",
            "number",
        ),
        case(
            "FIND_berry_in_elderberry",
            "FIND(\"berry\",\"elderberry\")",
            "=FIND(\"berry\",\"elderberry\")",
            "number",
        ),
        case(
            "SEARCH_err_in_elderberry",
            "SEARCH(\"ERR\",\"elderberry\")",
            "=SEARCH(\"ERR\",\"elderberry\")",
            "number",
        ),
        case(
            "CONCATENATE_apple_pie",
            "CONCATENATE(\"apple\",\"-\",\"pie\")",
            "=CONCATENATE(\"apple\",\"-\",\"pie\")",
            "string",
        ),
        case(
            "CONCAT_duke_sheets",
            "CONCAT(\"duke\",\"-\",\"sheets\")",
            "=CONCAT(\"duke\",\"-\",\"sheets\")",
            "string",
        ),
        case(
            "TEXTJOIN_skip_empty",
            "TEXTJOIN(\",\",TRUE,\"apple\",\"\",\"banana\")",
            "=TEXTJOIN(\",\",TRUE,\"apple\",\"\",\"banana\")",
            "string",
        ),
        case(
            "UPPER_kiwi",
            "UPPER(\"Kiwi\")",
            "=UPPER(\"Kiwi\")",
            "string",
        ),
        case(
            "LOWER_lemon",
            "LOWER(\"LEMON\")",
            "=LOWER(\"LEMON\")",
            "string",
        ),
        case(
            "PROPER_honeydew_melon",
            "PROPER(\"hONEYDEW MELON\")",
            "=PROPER(\"hONEYDEW MELON\")",
            "string",
        ),
        case(
            "TRIM_many_spaces",
            "TRIM(\"  too   many spaces  \")",
            "=TRIM(\"  too   many spaces  \")",
            "string",
        ),
        case(
            "DATE_2024_02_29",
            "DATE(2024,2,29)",
            "=DATE(2024,2,29)",
            "number",
        ),
        case(
            "YEAR_from_date",
            "YEAR(DATE(2024,2,29))",
            "=YEAR(DATE(2024,2,29))",
            "number",
        ),
        case(
            "MONTH_from_date",
            "MONTH(DATE(2024,2,29))",
            "=MONTH(DATE(2024,2,29))",
            "number",
        ),
        case(
            "DAY_from_date",
            "DAY(DATE(2024,2,29))",
            "=DAY(DATE(2024,2,29))",
            "number",
        ),
        case("TODAY_type_only", "TODAY()", "=TODAY()", "number"),
        case(
            "EDATE_plus_one_month",
            "EDATE(DATE(2024,1,31),1)",
            "=EDATE(DATE(2024,1,31),1)",
            "number",
        ),
        case(
            "EOMONTH_plus_one_month",
            "EOMONTH(DATE(2024,1,15),1)",
            "=EOMONTH(DATE(2024,1,15),1)",
            "number",
        ),
        case(
            "NETWORKDAYS_basic_range",
            "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10))",
            "=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10))",
            "number",
        ),
        case(
            "AND_all_true",
            "AND(TRUE,1<2,2<3)",
            "=AND(TRUE,1<2,2<3)",
            "boolean",
        ),
        case(
            "OR_one_true",
            "OR(FALSE,2<1,3=3)",
            "=OR(FALSE,2<1,3=3)",
            "boolean",
        ),
        case(
            "NOT_basic",
            "NOT(Data!A2>50)",
            "=NOT(Data!A2>50)",
            "boolean",
        ),
        case(
            "XOR_three_args",
            "XOR(TRUE,FALSE,TRUE)",
            "=XOR(TRUE,FALSE,TRUE)",
            "boolean",
        ),
        case(
            "IFERROR_div_zero_fallback",
            "IFERROR(1/0,\"fallback\")",
            "=IFERROR(1/0,\"fallback\")",
            "string",
        ),
        case(
            "IFNA_match_missing",
            "IFNA(MATCH(\"orange\",Data!A16:A25,0),\"missing\")",
            "=IFNA(MATCH(\"orange\",Data!A16:A25,0),\"missing\")",
            "string",
        ),
        case(
            "XLOOKUP_exact_kiwi",
            "XLOOKUP(\"kiwi\",Data!A16:A25,Data!B16:B25)",
            "=XLOOKUP(\"kiwi\",Data!A16:A25,Data!B16:B25)",
            "number",
        ),
        case(
            "XLOOKUP_if_not_found",
            "XLOOKUP(\"orange\",Data!A16:A25,Data!B16:B25,\"missing\")",
            "=XLOOKUP(\"orange\",Data!A16:A25,Data!B16:B25,\"missing\")",
            "string",
        ),
        case(
            "XLOOKUP_reverse_last_apple",
            "XLOOKUP(\"apple\",Data!M17:M20,Data!N17:N20,\"missing\",0,-1)",
            "=XLOOKUP(\"apple\",Data!M17:M20,Data!N17:N20,\"missing\",0,-1)",
            "number",
        ),
        case2(
            "FILTER_second_fruit_name",
            "INDEX(FILTER(Data!A29:A37,Data!B29:B37=\"fruit\"),2)",
            "=INDEX(FILTER(Data!A29:A37,Data!B29:B37=\"fruit\"),2)",
            "string",
        ),
        case2(
            "SORT_first_name",
            "INDEX(SORT(Data!A29:A37),1)",
            "=INDEX(SORT(Data!A29:A37),1)",
            "string",
        ),
        case2(
            "UNIQUE_category_count",
            "COUNTA(UNIQUE(Data!B29:B37))",
            "=COUNTA(UNIQUE(Data!B29:B37))",
            "number",
        ),
        case2(
            "SEQUENCE_sum_1_to_4",
            "SUM(SEQUENCE(4,1,1,1))",
            "=SUM(SEQUENCE(4,1,1,1))",
            "number",
        ),
        case(
            "VLOOKUP_approx_below_min",
            "VLOOKUP(-5,Data!P17:Q21,2,TRUE)",
            "=VLOOKUP(-5,Data!P17:Q21,2,TRUE)",
            "error",
        ),
        case(
            "VLOOKUP_approx_above_max",
            "VLOOKUP(99,Data!P17:Q21,2,TRUE)",
            "=VLOOKUP(99,Data!P17:Q21,2,TRUE)",
            "string",
        ),
        case(
            "VLOOKUP_approx_exact_bound",
            "VLOOKUP(90,Data!P17:Q21,2,TRUE)",
            "=VLOOKUP(90,Data!P17:Q21,2,TRUE)",
            "string",
        ),
        case(
            "VLOOKUP_approx_first_match",
            "VLOOKUP(0,Data!P17:Q21,2,TRUE)",
            "=VLOOKUP(0,Data!P17:Q21,2,TRUE)",
            "string",
        ),
        case(
            "INDEX_2arg_row_out_of_bounds",
            "INDEX(Data!A2:L2,15)",
            "=INDEX(Data!A2:L2,15)",
            "error",
        ),
        case(
            "INDEX_2arg_col_out_of_bounds",
            "INDEX(Data!A3:A14,15)",
            "=INDEX(Data!A3:A14,15)",
            "error",
        ),
        case(
            "INDEX_2arg_row_position_1",
            "INDEX(Data!A2:L2,1)",
            "=INDEX(Data!A2:L2,1)",
            "number",
        ),
        case(
            "INDEX_2arg_col_position_1",
            "INDEX(Data!A3:A14,1)",
            "=INDEX(Data!A3:A14,1)",
            "number",
        ),
        case(
            "INDEX_2arg_row_last",
            "INDEX(Data!A2:L2,12)",
            "=INDEX(Data!A2:L2,12)",
            "number",
        ),
        case(
            "MEDIAN_numbers",
            "MEDIAN(Data!A2:L2)",
            "=MEDIAN(Data!A2:L2)",
            "number",
        ),
        case(
            "STDEV_numbers",
            "STDEV(Data!A2:L2)",
            "=STDEV(Data!A2:L2)",
            "number",
        ),
        case(
            "VAR_numbers",
            "VAR(Data!A2:L2)",
            "=VAR(Data!A2:L2)",
            "number",
        ),
        case(
            "LARGE_3rd",
            "LARGE(Data!A2:L2,3)",
            "=LARGE(Data!A2:L2,3)",
            "number",
        ),
        case(
            "SMALL_2nd",
            "SMALL(Data!A2:L2,2)",
            "=SMALL(Data!A2:L2,2)",
            "number",
        ),
        case(
            "PERCENTILE_50",
            "PERCENTILE(Data!A2:L2,0.5)",
            "=PERCENTILE(Data!A2:L2,0.5)",
            "number",
        ),
        case(
            "QUARTILE_1",
            "QUARTILE(Data!A2:L2,1)",
            "=QUARTILE(Data!A2:L2,1)",
            "number",
        ),
        case(
            "RANK_30",
            "RANK(30,Data!A2:L2)",
            "=RANK(30,Data!A2:L2)",
            "number",
        ),
        case(
            "MODE_SNGL",
            "MODE.SNGL(Data!S1:S12)",
            "=MODE.SNGL(Data!S1:S12)",
            "number",
        ),
        case(
            "AVERAGEIF_fruit",
            "AVERAGEIF(Data!B29:B37,\"fruit\",Data!C29:C37)",
            "=AVERAGEIF(Data!B29:B37,\"fruit\",Data!C29:C37)",
            "number",
        ),
        case("ABS_negative", "ABS(-42.5)", "=ABS(-42.5)", "number"),
        case("MOD_basic", "MOD(17,5)", "=MOD(17,5)", "number"),
        case("POWER_basic", "POWER(2,10)", "=POWER(2,10)", "number"),
        case("SQRT_basic", "SQRT(144)", "=SQRT(144)", "number"),
        case("LOG_base10", "LOG(1000,10)", "=LOG(1000,10)", "number"),
        case("LOG10_basic", "LOG10(100)", "=LOG10(100)", "number"),
        case("LN_basic", "LN(EXP(1))", "=LN(EXP(1))", "number"),
        case(
            "CEILING_basic",
            "CEILING(2.3,1)",
            "=CEILING(2.3,1)",
            "number",
        ),
        case(
            "FLOOR_MATH_basic",
            "FLOOR.MATH(2.7,1)",
            "=FLOOR.MATH(2.7,1)",
            "number",
        ),
        case("INT_basic", "INT(3.9)", "=INT(3.9)", "number"),
        case("SIGN_negative", "SIGN(-42)", "=SIGN(-42)", "number"),
        case("PI_value", "PI()", "=PI()", "number"),
        case("RAND_type", "RAND()", "=RAND()", "number"),
        case(
            "RANDBETWEEN_type",
            "RANDBETWEEN(1,100)",
            "=RANDBETWEEN(1,100)",
            "number",
        ),
        case(
            "ISNUMBER_yes",
            "ISNUMBER(Data!A50)",
            "=ISNUMBER(Data!A50)",
            "boolean",
        ),
        case(
            "ISNUMBER_no",
            "ISNUMBER(Data!A51)",
            "=ISNUMBER(Data!A51)",
            "boolean",
        ),
        case(
            "ISTEXT_yes",
            "ISTEXT(Data!A51)",
            "=ISTEXT(Data!A51)",
            "boolean",
        ),
        case(
            "ISTEXT_no",
            "ISTEXT(Data!A50)",
            "=ISTEXT(Data!A50)",
            "boolean",
        ),
        case(
            "ISBLANK_yes",
            "ISBLANK(Data!A53)",
            "=ISBLANK(Data!A53)",
            "boolean",
        ),
        case(
            "ISBLANK_no",
            "ISBLANK(Data!A50)",
            "=ISBLANK(Data!A50)",
            "boolean",
        ),
        case("ISERROR_yes", "ISERROR(1/0)", "=ISERROR(1/0)", "boolean"),
        case("ISERROR_no", "ISERROR(42)", "=ISERROR(42)", "boolean"),
        case(
            "ISLOGICAL_yes",
            "ISLOGICAL(Data!A52)",
            "=ISLOGICAL(Data!A52)",
            "boolean",
        ),
        case(
            "ISLOGICAL_no",
            "ISLOGICAL(Data!A50)",
            "=ISLOGICAL(Data!A50)",
            "boolean",
        ),
        case("TYPE_number", "TYPE(Data!A50)", "=TYPE(Data!A50)", "number"),
        case("TYPE_text", "TYPE(Data!A51)", "=TYPE(Data!A51)", "number"),
        case(
            "TYPE_logical",
            "TYPE(Data!A52)",
            "=TYPE(Data!A52)",
            "number",
        ),
        case(
            "SUBSTITUTE_basic",
            "SUBSTITUTE(\"Hello World\",\"World\",\"Earth\")",
            "=SUBSTITUTE(\"Hello World\",\"World\",\"Earth\")",
            "string",
        ),
        case(
            "SUBSTITUTE_nth",
            "SUBSTITUTE(\"mississippi\",\"s\",\"S\",2)",
            "=SUBSTITUTE(\"mississippi\",\"s\",\"S\",2)",
            "string",
        ),
        case(
            "REPLACE_basic",
            "REPLACE(\"Hello World\",7,5,\"Earth\")",
            "=REPLACE(\"Hello World\",7,5,\"Earth\")",
            "string",
        ),
        case("REPT_basic", "REPT(\"ab\",3)", "=REPT(\"ab\",3)", "string"),
        case(
            "EXACT_match",
            "EXACT(\"hello\",\"hello\")",
            "=EXACT(\"hello\",\"hello\")",
            "boolean",
        ),
        case(
            "EXACT_no_match",
            "EXACT(\"hello\",\"Hello\")",
            "=EXACT(\"hello\",\"Hello\")",
            "boolean",
        ),
        case(
            "VALUE_numeric",
            "VALUE(\"123.45\")",
            "=VALUE(\"123.45\")",
            "number",
        ),
        case(
            "TEXT_format",
            "TEXT(0.75,\"0.0%\")",
            "=TEXT(0.75,\"0.0%\")",
            "string",
        ),
        case(
            "CLEAN_basic",
            "CLEAN(CHAR(9)&\"hello\"&CHAR(10))",
            "=CLEAN(CHAR(9)&\"hello\"&CHAR(10))",
            "string",
        ),
        case(
            "NUMBERVALUE_basic",
            "NUMBERVALUE(\"1,234.56\",\".\",\",\")",
            "=NUMBERVALUE(\"1,234.56\",\".\",\",\")",
            "number",
        ),
        case(
            "PMT_basic",
            "PMT(Data!B43/Data!D43,Data!C43*Data!D43,-Data!A43)",
            "=PMT(Data!B43/Data!D43,Data!C43*Data!D43,-Data!A43)",
            "number",
        ),
        case(
            "FV_basic",
            "FV(0.05/12,10*12,-200)",
            "=FV(0.05/12,10*12,-200)",
            "number",
        ),
        case(
            "PV_basic",
            "PV(0.08/12,20*12,-500)",
            "=PV(0.08/12,20*12,-500)",
            "number",
        ),
        case(
            "NPER_basic",
            "NPER(0.06/12,-200,10000)",
            "=NPER(0.06/12,-200,10000)",
            "number",
        ),
        case(
            "NPV_basic",
            "NPV(0.1,Data!A45:A48)+Data!A44",
            "=NPV(0.1,Data!A45:A48)+Data!A44",
            "number",
        ),
        case(
            "IRR_basic",
            "IRR(Data!A44:A48)",
            "=IRR(Data!A44:A48)",
            "number",
        ),
        case("ERROR_div_zero", "1/0", "=1/0", "error"),
        case(
            "ERROR_ref",
            "INDEX(Data!A2:L2,99)",
            "=INDEX(Data!A2:L2,99)",
            "error",
        ),
        case(
            "ERROR_value",
            "VALUE(\"not_a_number\")",
            "=VALUE(\"not_a_number\")",
            "error",
        ),
        case(
            "ERROR_nested_iferror",
            "IFERROR(IFERROR(1/0,SQRT(-1)),\"caught\")",
            "=IFERROR(IFERROR(1/0,SQRT(-1)),\"caught\")",
            "string",
        ),
        case("SUM_with_error", "SUM(1,2,1/0)", "=SUM(1,2,1/0)", "error"),
        case(
            "SUM_with_blanks",
            "SUM(Data!A40:E40)",
            "=SUM(Data!A40:E40)",
            "number",
        ),
        case(
            "COUNTA_with_blanks",
            "COUNTA(Data!A40:E40)",
            "=COUNTA(Data!A40:E40)",
            "number",
        ),
        case(
            "COUNTBLANK_range",
            "COUNTBLANK(Data!A40:E40)",
            "=COUNTBLANK(Data!A40:E40)",
            "number",
        ),
        case(
            "CONCATENATE_blank",
            "CONCATENATE(Data!A40,Data!B40,Data!C40)",
            "=CONCATENATE(Data!A40,Data!B40,Data!C40)",
            "string",
        ),
        case(
            "IF_blank",
            "IF(Data!B40=\"\",\"empty\",\"full\")",
            "=IF(Data!B40=\"\",\"empty\",\"full\")",
            "string",
        ),
        case(
            "PRECISION_sum",
            "Data!A56+Data!B56",
            "=Data!A56+Data!B56",
            "number",
        ),
        case(
            "PRECISION_round",
            "ROUND(Data!A56+Data!B56,1)",
            "=ROUND(Data!A56+Data!B56,1)",
            "number",
        ),
        case(
            "PRECISION_large",
            "Data!A57+Data!B57",
            "=Data!A57+Data!B57",
            "number",
        ),
        case(
            "SUMPRODUCT_IF",
            "SUMPRODUCT((Data!B29:B37=\"fruit\")*Data!C29:C37)",
            "=SUMPRODUCT((Data!B29:B37=\"fruit\")*Data!C29:C37)",
            "number",
        ),
        case(
            "SUMPRODUCT_MULTI",
            "SUMPRODUCT((Data!B29:B37=\"fruit\")*(Data!C29:C37>15))",
            "=SUMPRODUCT((Data!B29:B37=\"fruit\")*(Data!C29:C37>15))",
            "number",
        ),
        case(
            "CHOOSE_basic",
            "CHOOSE(2,\"apple\",\"banana\",\"cherry\")",
            "=CHOOSE(2,\"apple\",\"banana\",\"cherry\")",
            "string",
        ),
        case(
            "CHOOSE_calc",
            "CHOOSE(MATCH(70,Data!P17:P21,1),\"F\",\"D\",\"C\",\"B\",\"A\")",
            "=CHOOSE(MATCH(70,Data!P17:P21,1),\"F\",\"D\",\"C\",\"B\",\"A\")",
            "string",
        ),
        case2(
            "SORTBY_basic",
            "INDEX(SORTBY(Data!A29:A37,Data!C29:C37,-1),1)",
            "=INDEX(SORTBY(Data!A29:A37,Data!C29:C37,-1),1)",
            "string",
        ),
        case2(
            "FILTER_multi",
            "SUM(FILTER(Data!C29:C37,(Data!B29:B37=\"fruit\")*(Data!C29:C37>15)))",
            "=SUM(FILTER(Data!C29:C37,(Data!B29:B37=\"fruit\")*(Data!C29:C37>15)))",
            "number",
        ),
        case2(
            "SEQUENCE_2d_sum",
            "SUM(SEQUENCE(3,3,1,1))",
            "=SUM(SEQUENCE(3,3,1,1))",
            "number",
        ),
        case("TMPL_ref_row70", "Data!B70", "=Data!B70", "number"),
        case("TMPL_ref_row71", "Data!B71", "=Data!B71", "number"),
        case("TMPL_ref_row75", "Data!B75", "=Data!B75", "number"),
        case("TMPL_ref_row80", "Data!B80", "=Data!B80", "number"),
        case("TMPL_ref_row85", "Data!B85", "=Data!B85", "number"),
        case("TMPL_ref_row89", "Data!B89", "=Data!B89", "number"),
        case(
            "TMPL_sum_row70",
            "Data!A70+Data!B70",
            "=Data!A70+Data!B70",
            "number",
        ),
        case(
            "TMPL_sum_row75",
            "Data!A75+Data!B75",
            "=Data!A75+Data!B75",
            "number",
        ),
        case(
            "TMPL_sum_row80",
            "Data!A80+Data!B80",
            "=Data!A80+Data!B80",
            "number",
        ),
        case(
            "TMPL_sum_row89",
            "Data!A89+Data!B89",
            "=Data!A89+Data!B89",
            "number",
        ),
        case(
            "TMPL_vlookup_70",
            "VLOOKUP(1,Data!A70:B89,2,FALSE)",
            "=VLOOKUP(1,Data!A70:B89,2,FALSE)",
            "number",
        ),
        case(
            "TMPL_vlookup_89",
            "VLOOKUP(20,Data!A70:B89,2,FALSE)",
            "=VLOOKUP(20,Data!A70:B89,2,FALSE)",
            "number",
        ),
        case(
            "TMPL_if_row70",
            "IF(Data!C70=\"alpha\",\"yes\",\"no\")",
            "=IF(Data!C70=\"alpha\",\"yes\",\"no\")",
            "string",
        ),
        case(
            "TMPL_if_row71",
            "IF(Data!C71=\"alpha\",\"yes\",\"no\")",
            "=IF(Data!C71=\"alpha\",\"yes\",\"no\")",
            "string",
        ),
        case(
            "TMPL_if_row72",
            "IF(Data!C72=\"alpha\",\"yes\",\"no\")",
            "=IF(Data!C72=\"alpha\",\"yes\",\"no\")",
            "string",
        ),
        case(
            "TMPL_index_r5",
            "INDEX(Data!B70:B89,5)",
            "=INDEX(Data!B70:B89,5)",
            "number",
        ),
        case(
            "TMPL_index_r10",
            "INDEX(Data!B70:B89,10)",
            "=INDEX(Data!B70:B89,10)",
            "number",
        ),
        case(
            "TMPL_index_r20",
            "INDEX(Data!B70:B89,20)",
            "=INDEX(Data!B70:B89,20)",
            "number",
        ),
        case(
            "TMPL_countif_alpha",
            "COUNTIF(Data!C70:C89,\"alpha\")",
            "=COUNTIF(Data!C70:C89,\"alpha\")",
            "number",
        ),
        case(
            "TMPL_sumif_alpha",
            "SUMIF(Data!C70:C89,\"alpha\",Data!B70:B89)",
            "=SUMIF(Data!C70:C89,\"alpha\",Data!B70:B89)",
            "number",
        ),
        case(
            "PMT_small_loan",
            "PMT(Data!B92/Data!D92,Data!C92*Data!D92,-Data!A92)",
            "=PMT(Data!B92/Data!D92,Data!C92*Data!D92,-Data!A92)",
            "number",
        ),
        case(
            "RATE_basic",
            "RATE(60,-200,10000)*12",
            "=RATE(60,-200,10000)*12",
            "number",
        ),
        case(
            "NPV_alt",
            "NPV(0.08,Data!A94:A97)+Data!A93",
            "=NPV(0.08,Data!A94:A97)+Data!A93",
            "number",
        ),
        case(
            "IRR_alt",
            "IRR(Data!A93:A97)",
            "=IRR(Data!A93:A97)",
            "number",
        ),
        case(
            "FV_monthly",
            "FV(0.04/12,5*12,-300)",
            "=FV(0.04/12,5*12,-300)",
            "number",
        ),
        case(
            "PV_annuity",
            "PV(0.05,10,-1000)",
            "=PV(0.05,10,-1000)",
            "number",
        ),
        case(
            "NPER_double",
            "NPER(0.07/12,-500,30000)",
            "=NPER(0.07/12,-500,30000)",
            "number",
        ),
        case(
            "IPMT_first",
            "IPMT(0.06/12,1,360,-200000)",
            "=IPMT(0.06/12,1,360,-200000)",
            "number",
        ),
        case(
            "PPMT_first",
            "PPMT(0.06/12,1,360,-200000)",
            "=PPMT(0.06/12,1,360,-200000)",
            "number",
        ),
        case(
            "CUMIPMT_year1",
            "CUMIPMT(0.06/12,360,200000,1,12,0)",
            "=CUMIPMT(0.06/12,360,200000,1,12,0)",
            "number",
        ),
        case(
            "CUMPRINC_year1",
            "CUMPRINC(0.06/12,360,200000,1,12,0)",
            "=CUMPRINC(0.06/12,360,200000,1,12,0)",
            "number",
        ),
        case(
            "SLN_basic",
            "SLN(10000,1000,10)",
            "=SLN(10000,1000,10)",
            "number",
        ),
        case(
            "DB_year1",
            "DB(10000,1000,10,1)",
            "=DB(10000,1000,10,1)",
            "number",
        ),
        case(
            "EFFECT_basic",
            "EFFECT(0.06,12)",
            "=EFFECT(0.06,12)",
            "number",
        ),
        case(
            "NOMINAL_basic",
            "NOMINAL(0.0617,12)",
            "=NOMINAL(0.0617,12)",
            "number",
        ),
        case(
            "STDEV_S_set",
            "STDEV.S(Data!A100:A109)",
            "=STDEV.S(Data!A100:A109)",
            "number",
        ),
        case(
            "STDEV_P_set",
            "STDEV.P(Data!A100:A109)",
            "=STDEV.P(Data!A100:A109)",
            "number",
        ),
        case(
            "VAR_S_set",
            "VAR.S(Data!A100:A109)",
            "=VAR.S(Data!A100:A109)",
            "number",
        ),
        case(
            "VAR_P_set",
            "VAR.P(Data!A100:A109)",
            "=VAR.P(Data!A100:A109)",
            "number",
        ),
        case(
            "PERCENTILE_INC_25",
            "PERCENTILE.INC(Data!A100:A109,0.25)",
            "=PERCENTILE.INC(Data!A100:A109,0.25)",
            "number",
        ),
        case(
            "PERCENTILE_INC_75",
            "PERCENTILE.INC(Data!A100:A109,0.75)",
            "=PERCENTILE.INC(Data!A100:A109,0.75)",
            "number",
        ),
        case(
            "PERCENTILE_EXC_25",
            "PERCENTILE.EXC(Data!A100:A109,0.25)",
            "=PERCENTILE.EXC(Data!A100:A109,0.25)",
            "number",
        ),
        case(
            "QUARTILE_INC_3",
            "QUARTILE.INC(Data!A100:A109,3)",
            "=QUARTILE.INC(Data!A100:A109,3)",
            "number",
        ),
        case(
            "LARGE_1st",
            "LARGE(Data!A100:A109,1)",
            "=LARGE(Data!A100:A109,1)",
            "number",
        ),
        case(
            "SMALL_1st",
            "SMALL(Data!A100:A109,1)",
            "=SMALL(Data!A100:A109,1)",
            "number",
        ),
        case(
            "RANK_EQ_22",
            "RANK.EQ(22,Data!A100:A109)",
            "=RANK.EQ(22,Data!A100:A109)",
            "number",
        ),
        case(
            "RANK_AVG_22",
            "RANK.AVG(22,Data!A100:A109)",
            "=RANK.AVG(22,Data!A100:A109)",
            "number",
        ),
        case(
            "AVERAGE_set",
            "AVERAGE(Data!A100:A109)",
            "=AVERAGE(Data!A100:A109)",
            "number",
        ),
        case(
            "GEOMEAN_set",
            "GEOMEAN(Data!A100:A109)",
            "=GEOMEAN(Data!A100:A109)",
            "number",
        ),
        case(
            "HARMEAN_set",
            "HARMEAN(Data!A100:A109)",
            "=HARMEAN(Data!A100:A109)",
            "number",
        ),
        case(
            "TRIMMEAN_set",
            "TRIMMEAN(Data!A100:A109,0.2)",
            "=TRIMMEAN(Data!A100:A109,0.2)",
            "number",
        ),
        case(
            "DEVSQ_set",
            "DEVSQ(Data!A100:A109)",
            "=DEVSQ(Data!A100:A109)",
            "number",
        ),
        case(
            "AVEDEV_set",
            "AVEDEV(Data!A100:A109)",
            "=AVEDEV(Data!A100:A109)",
            "number",
        ),
        case(
            "KURT_set",
            "KURT(Data!A100:A109)",
            "=KURT(Data!A100:A109)",
            "number",
        ),
        case(
            "SKEW_set",
            "SKEW(Data!A100:A109)",
            "=SKEW(Data!A100:A109)",
            "number",
        ),
        case2(
            "FILTER_gt25",
            "SUM(FILTER(Data!B70:B89,Data!B70:B89>500))",
            "=SUM(FILTER(Data!B70:B89,Data!B70:B89>500))",
            "number",
        ),
        case2(
            "FILTER_alpha_sum",
            "SUM(FILTER(Data!B70:B89,Data!C70:C89=\"alpha\"))",
            "=SUM(FILTER(Data!B70:B89,Data!C70:C89=\"alpha\"))",
            "number",
        ),
        case2(
            "FILTER_no_match",
            "FILTER(Data!B70:B89,Data!B70:B89>99999,\"none\")",
            "=FILTER(Data!B70:B89,Data!B70:B89>99999,\"none\")",
            "string",
        ),
        case2(
            "SORT_desc_first",
            "INDEX(SORT(Data!B70:B89,1,-1),1)",
            "=INDEX(SORT(Data!B70:B89,1,-1),1)",
            "number",
        ),
        case2(
            "SORT_asc_first",
            "INDEX(SORT(Data!B70:B89),1)",
            "=INDEX(SORT(Data!B70:B89),1)",
            "number",
        ),
        case2(
            "SORTBY_name_first",
            "INDEX(SORTBY(Data!A70:A89,Data!B70:B89,-1),1)",
            "=INDEX(SORTBY(Data!A70:A89,Data!B70:B89,-1),1)",
            "number",
        ),
        case2(
            "UNIQUE_categories",
            "COUNTA(UNIQUE(Data!C70:C89))",
            "=COUNTA(UNIQUE(Data!C70:C89))",
            "number",
        ),
        case2(
            "SEQUENCE_sum_5x5",
            "SUM(SEQUENCE(5,5,1,1))",
            "=SUM(SEQUENCE(5,5,1,1))",
            "number",
        ),
        case2(
            "SEQUENCE_start_10",
            "INDEX(SEQUENCE(5,1,10,10),3)",
            "=INDEX(SEQUENCE(5,1,10,10),3)",
            "number",
        ),
        case2(
            "TRANSPOSE_elem",
            "INDEX(TRANSPOSE(Data!A115:C117),1,3)",
            "=INDEX(TRANSPOSE(Data!A115:C117),1,3)",
            "number",
        ),
        case2(
            "FILTER_AND_SORT",
            "INDEX(SORT(FILTER(Data!B70:B89,Data!C70:C89=\"alpha\"),1,-1),1)",
            "=INDEX(SORT(FILTER(Data!B70:B89,Data!C70:C89=\"alpha\"),1,-1),1)",
            "number",
        ),
        case2(
            "RANDARRAY_type",
            "SUM(RANDARRAY(3,3))",
            "=SUM(RANDARRAY(3,3))",
            "number",
        ),
        case2(
            "MMULT_row1",
            "INDEX(MMULT(Data!A115:C117,TRANSPOSE(Data!A115:C115)),1)",
            "=INDEX(MMULT(Data!A115:C117,TRANSPOSE(Data!A115:C115)),1)",
            "number",
        ),
        case2(
            "MMULT_row3",
            "INDEX(MMULT(Data!A115:C117,TRANSPOSE(Data!A115:C115)),3)",
            "=INDEX(MMULT(Data!A115:C117,TRANSPOSE(Data!A115:C115)),3)",
            "number",
        ),
        case2(
            "CHOOSECOLS_elem",
            "INDEX(CHOOSECOLS(Data!A70:C89,2,3),1,1)",
            "=INDEX(CHOOSECOLS(Data!A70:C89,2,3),1,1)",
            "number",
        ),
        case(
            "XLOOKUP_wildcard",
            "XLOOKUP(\"ch*\",Data!A16:A25,Data!B16:B25,,2)",
            "=XLOOKUP(\"ch*\",Data!A16:A25,Data!B16:B25,,2)",
            "number",
        ),
        case(
            "XLOOKUP_not_found_def",
            "XLOOKUP(\"mango\",Data!A16:A25,Data!B16:B25,-1)",
            "=XLOOKUP(\"mango\",Data!A16:A25,Data!B16:B25,-1)",
            "number",
        ),
        case(
            "MATCH_wildcard",
            "MATCH(\"*berry\",Data!A16:A25,0)",
            "=MATCH(\"*berry\",Data!A16:A25,0)",
            "number",
        ),
        case(
            "INDEX_MATCH_price",
            "INDEX(Data!B16:B25,MATCH(\"grape\",Data!A16:A25,0))",
            "=INDEX(Data!B16:B25,MATCH(\"grape\",Data!A16:A25,0))",
            "number",
        ),
        case(
            "VLOOKUP_col3",
            "VLOOKUP(\"Q2\",Data!D17:G20,3,FALSE)",
            "=VLOOKUP(\"Q2\",Data!D17:G20,3,FALSE)",
            "number",
        ),
        case(
            "HLOOKUP_row2",
            "HLOOKUP(8,Data!A1:L2,2,FALSE)",
            "=HLOOKUP(8,Data!A1:L2,2,FALSE)",
            "number",
        ),
        case(
            "LOOKUP_vector",
            "LOOKUP(75,Data!A2:L2,Data!A1:L1)",
            "=LOOKUP(75,Data!A2:L2,Data!A1:L1)",
            "number",
        ),
        case(
            "XMATCH_exact",
            "XMATCH(\"fig\",Data!A16:A25)",
            "=XMATCH(\"fig\",Data!A16:A25)",
            "number",
        ),
        case(
            "XMATCH_wildcard",
            "XMATCH(\"*dew\",Data!A16:A25,2)",
            "=XMATCH(\"*dew\",Data!A16:A25,2)",
            "number",
        ),
        case(
            "DATEDIF_months",
            "DATEDIF(Data!A112,Data!A113,\"M\")",
            "=DATEDIF(Data!A112,Data!A113,\"M\")",
            "number",
        ),
        case(
            "DAYS_between",
            "DAYS(Data!A113,Data!A112)",
            "=DAYS(Data!A113,Data!A112)",
            "number",
        ),
        case(
            "WEEKDAY_sun",
            "WEEKDAY(Data!A112)",
            "=WEEKDAY(Data!A112)",
            "number",
        ),
        case(
            "WEEKNUM_basic",
            "WEEKNUM(Data!A112)",
            "=WEEKNUM(Data!A112)",
            "number",
        ),
        case(
            "ISOWEEKNUM_basic",
            "ISOWEEKNUM(Data!A112)",
            "=ISOWEEKNUM(Data!A112)",
            "number",
        ),
        case(
            "WORKDAY_plus10",
            "WORKDAY(Data!A112,10)",
            "=WORKDAY(Data!A112,10)",
            "number",
        ),
        case(
            "DATEVALUE_basic",
            "DATEVALUE(\"2024-06-15\")",
            "=DATEVALUE(\"2024-06-15\")",
            "number",
        ),
        case(
            "TIMEVALUE_basic",
            "TIMEVALUE(\"14:30:00\")",
            "=TIMEVALUE(\"14:30:00\")",
            "number",
        ),
        case(
            "HOUR_basic",
            "HOUR(TIMEVALUE(\"14:30:45\"))",
            "=HOUR(TIMEVALUE(\"14:30:45\"))",
            "number",
        ),
        case(
            "MINUTE_basic",
            "MINUTE(TIMEVALUE(\"14:30:45\"))",
            "=MINUTE(TIMEVALUE(\"14:30:45\"))",
            "number",
        ),
        case(
            "SECOND_basic",
            "SECOND(TIMEVALUE(\"14:30:45\"))",
            "=SECOND(TIMEVALUE(\"14:30:45\"))",
            "number",
        ),
        case("GCD_basic", "GCD(12,18,24)", "=GCD(12,18,24)", "number"),
        case("LCM_basic", "LCM(4,6,10)", "=LCM(4,6,10)", "number"),
        case("FACT_basic", "FACT(7)", "=FACT(7)", "number"),
        case("COMBIN_basic", "COMBIN(10,3)", "=COMBIN(10,3)", "number"),
        case("PERMUT_basic", "PERMUT(10,3)", "=PERMUT(10,3)", "number"),
        case(
            "MROUND_basic",
            "MROUND(7.3,0.5)",
            "=MROUND(7.3,0.5)",
            "number",
        ),
        case(
            "QUOTIENT_basic",
            "QUOTIENT(17,5)",
            "=QUOTIENT(17,5)",
            "number",
        ),
        case(
            "PRODUCT_range",
            "PRODUCT(Data!A100:A104)",
            "=PRODUCT(Data!A100:A104)",
            "number",
        ),
        case("SUMSQ_basic", "SUMSQ(3,4,5)", "=SUMSQ(3,4,5)", "number"),
        case(
            "SUMX2MY2_basic",
            "SUMX2MY2(Data!A100:A104,Data!A105:A109)",
            "=SUMX2MY2(Data!A100:A104,Data!A105:A109)",
            "number",
        ),
        case("SIN_zero", "SIN(Data!A120)", "=SIN(Data!A120)", "number"),
        case("SIN_pi_6", "SIN(Data!A121)", "=SIN(Data!A121)", "number"),
        case("COS_zero", "COS(Data!A120)", "=COS(Data!A120)", "number"),
        case("COS_pi_2", "COS(Data!A123)", "=COS(Data!A123)", "number"),
        case("TAN_pi_4", "TAN(Data!A122)", "=TAN(Data!A122)", "number"),
        case("ASIN_half", "ASIN(0.5)", "=ASIN(0.5)", "number"),
        case("ACOS_half", "ACOS(0.5)", "=ACOS(0.5)", "number"),
        case("ATAN_1", "ATAN(1)", "=ATAN(1)", "number"),
        case("ATAN2_basic", "ATAN2(1,1)", "=ATAN2(1,1)", "number"),
        case(
            "DEGREES_pi",
            "DEGREES(Data!A124)",
            "=DEGREES(Data!A124)",
            "number",
        ),
        case("RADIANS_180", "RADIANS(180)", "=RADIANS(180)", "number"),
        case("SINH_basic", "SINH(1)", "=SINH(1)", "number"),
        case("COSH_basic", "COSH(1)", "=COSH(1)", "number"),
        case("TANH_basic", "TANH(1)", "=TANH(1)", "number"),
        case("CODE_A", "CODE(Data!A128)", "=CODE(Data!A128)", "number"),
        case("CHAR_65", "CHAR(65)", "=CHAR(65)", "string"),
        case(
            "DOLLAR_basic",
            "DOLLAR(1234.567,2)",
            "=DOLLAR(1234.567,2)",
            "string",
        ),
        case(
            "FIXED_basic",
            "FIXED(1234.567,2,FALSE)",
            "=FIXED(1234.567,2,FALSE)",
            "string",
        ),
        case(
            "UNICODE_basic",
            "UNICODE(Data!A128)",
            "=UNICODE(Data!A128)",
            "number",
        ),
        case("UNICHAR_65", "UNICHAR(65)", "=UNICHAR(65)", "string"),
        case("T_number", "T(Data!A50)", "=T(Data!A50)", "string"),
        case("T_string", "T(Data!A126)", "=T(Data!A126)", "string"),
        case(
            "TEXTBEFORE_basic",
            "TEXTBEFORE(Data!A129,\"@\")",
            "=TEXTBEFORE(Data!A129,\"@\")",
            "string",
        ),
        case(
            "TEXTAFTER_basic",
            "TEXTAFTER(Data!A129,\"@\")",
            "=TEXTAFTER(Data!A129,\"@\")",
            "string",
        ),
        case(
            "LEFT_default",
            "LEFT(Data!A126)",
            "=LEFT(Data!A126)",
            "string",
        ),
        case("ROMAN_basic", "ROMAN(2024)", "=ROMAN(2024)", "string"),
        case("EXP_1", "EXP(1)", "=EXP(1)", "number"),
        case("EVEN_basic", "EVEN(3)", "=EVEN(3)", "number"),
        case("ODD_basic", "ODD(4)", "=ODD(4)", "number"),
        case("TRUNC_basic", "TRUNC(3.75,1)", "=TRUNC(3.75,1)", "number"),
        case(
            "DECIMAL_hex",
            "DECIMAL(\"FF\",16)",
            "=DECIMAL(\"FF\",16)",
            "number",
        ),
        case("BASE_dec", "BASE(255,16)", "=BASE(255,16)", "string"),
        case("SQRTPI_basic", "SQRTPI(2)", "=SQRTPI(2)", "number"),
        case(
            "MULTINOMIAL_b",
            "MULTINOMIAL(2,3,4)",
            "=MULTINOMIAL(2,3,4)",
            "number",
        ),
        case(
            "SERIESSUM_b",
            "SERIESSUM(2,0,1,{1,1,1})",
            "=SERIESSUM(2,0,1,{1,1,1})",
            "number",
        ),
        case("ROW_basic", "ROW(Data!A50)", "=ROW(Data!A50)", "number"),
        case(
            "COLUMN_basic",
            "COLUMN(Data!B50)",
            "=COLUMN(Data!B50)",
            "number",
        ),
        case(
            "ROWS_basic",
            "ROWS(Data!A1:A10)",
            "=ROWS(Data!A1:A10)",
            "number",
        ),
        case(
            "COLUMNS_basic",
            "COLUMNS(Data!A1:L1)",
            "=COLUMNS(Data!A1:L1)",
            "number",
        ),
        case("ADDRESS_basic", "ADDRESS(1,1)", "=ADDRESS(1,1)", "string"),
        case("ADDRESS_rel", "ADDRESS(1,1,4)", "=ADDRESS(1,1,4)", "string"),
        case(
            "INDIRECT_basic",
            "INDIRECT(\"Data!A50\")",
            "=INDIRECT(\"Data!A50\")",
            "number",
        ),
        case(
            "FORMULATEXT_b",
            "FORMULATEXT(Data!A121)",
            "=FORMULATEXT(Data!A121)",
            "string",
        ),
        case(
            "MAXIFS_basic",
            "MAXIFS(Data!B138:B142,Data!A138:A142,\"alice\")",
            "=MAXIFS(Data!B138:B142,Data!A138:A142,\"alice\")",
            "number",
        ),
        case(
            "MINIFS_basic",
            "MINIFS(Data!B138:B142,Data!A138:A142,\"alice\")",
            "=MINIFS(Data!B138:B142,Data!A138:A142,\"alice\")",
            "number",
        ),
        case(
            "AVERAGEIFS_basic",
            "AVERAGEIFS(Data!B138:B142,Data!A138:A142,\"bob\",Data!C138:C142,\"math\")",
            "=AVERAGEIFS(Data!B138:B142,Data!A138:A142,\"bob\",Data!C138:C142,\"math\")",
            "number",
        ),
        case(
            "AVERAGEIFS_multi",
            "AVERAGEIFS(Data!B138:B142,Data!A138:A142,\"alice\",Data!C138:C142,\"math\")",
            "=AVERAGEIFS(Data!B138:B142,Data!A138:A142,\"alice\",Data!C138:C142,\"math\")",
            "number",
        ),
        case(
            "SUMIFS_multi_cond",
            "SUMIFS(Data!B138:B142,Data!A138:A142,\"alice\",Data!C138:C142,\"math\")",
            "=SUMIFS(Data!B138:B142,Data!A138:A142,\"alice\",Data!C138:C142,\"math\")",
            "number",
        ),
        case(
            "COUNTIFS_name",
            "COUNTIFS(Data!A138:A142,\"alice\")",
            "=COUNTIFS(Data!A138:A142,\"alice\")",
            "number",
        ),
        case(
            "COUNTIFS_multi",
            "COUNTIFS(Data!A138:A142,\"alice\",Data!C138:C142,\"math\")",
            "=COUNTIFS(Data!A138:A142,\"alice\",Data!C138:C142,\"math\")",
            "number",
        ),
        case(
            "AGGREGATE_avg",
            "AGGREGATE(1,0,Data!A132:A136)",
            "=AGGREGATE(1,0,Data!A132:A136)",
            "number",
        ),
        case(
            "AGGREGATE_skip_err",
            "AGGREGATE(1,6,Data!A144:A148)",
            "=AGGREGATE(1,6,Data!A144:A148)",
            "number",
        ),
        case(
            "SUBTOTAL_sum",
            "SUBTOTAL(9,Data!A132:A136)",
            "=SUBTOTAL(9,Data!A132:A136)",
            "number",
        ),
        case("TRUE_val", "TRUE()", "=TRUE()", "boolean"),
        case("FALSE_val", "FALSE()", "=FALSE()", "boolean"),
        case("N_number", "N(42)", "=N(42)", "number"),
        case("N_bool", "N(TRUE)", "=N(TRUE)", "number"),
        case("N_text", "N(\"hello\")", "=N(\"hello\")", "number"),
        case("NA_error", "NA()", "=NA()", "error"),
        case("NOW_type", "NOW()", "=NOW()", "number"),
        case("TIME_basic", "TIME(14,30,0)", "=TIME(14,30,0)", "number"),
        case(
            "DATE_year_month",
            "YEAR(Data!A112)&\"-\"&MONTH(Data!A112)",
            "=YEAR(Data!A112)&\"-\"&MONTH(Data!A112)",
            "string",
        ),
        case(
            "YEARFRAC_basic",
            "YEARFRAC(Data!A112,Data!A113)",
            "=YEARFRAC(Data!A112,Data!A113)",
            "number",
        ),
        case(
            "EDATE_neg",
            "EDATE(Data!A113,-3)",
            "=EDATE(Data!A113,-3)",
            "number",
        ),
        case(
            "EOMONTH_neg",
            "EOMONTH(Data!A113,-1)",
            "=EOMONTH(Data!A113,-1)",
            "number",
        ),
        case(
            "WEEKDAY_mon",
            "WEEKDAY(Data!A112,2)",
            "=WEEKDAY(Data!A112,2)",
            "number",
        ),
        case(
            "NETWORKDAYS_INTL",
            "NETWORKDAYS.INTL(Data!A112,Data!A113,1)",
            "=NETWORKDAYS.INTL(Data!A112,Data!A113,1)",
            "number",
        ),
        case(
            "AVERAGEA_mixed",
            "AVERAGEA(Data!A50,Data!A52,Data!A51,Data!A56)",
            "=AVERAGEA(Data!A50,Data!A52,Data!A51,Data!A56)",
            "number",
        ),
        case(
            "MAXA_mixed",
            "MAXA(Data!A50,Data!A52,Data!A51,Data!A56)",
            "=MAXA(Data!A50,Data!A52,Data!A51,Data!A56)",
            "number",
        ),
        case(
            "MINA_mixed",
            "MINA(Data!A50,Data!A52,Data!A51,Data!A56)",
            "=MINA(Data!A50,Data!A52,Data!A51,Data!A56)",
            "number",
        ),
        case(
            "CORREL_basic",
            "CORREL(Data!A100:A104,Data!A105:A109)",
            "=CORREL(Data!A100:A104,Data!A105:A109)",
            "number",
        ),
        case(
            "INTERCEPT_basic",
            "INTERCEPT(Data!A105:A109,Data!A100:A104)",
            "=INTERCEPT(Data!A105:A109,Data!A100:A104)",
            "number",
        ),
        case(
            "SLOPE_basic",
            "SLOPE(Data!A105:A109,Data!A100:A104)",
            "=SLOPE(Data!A105:A109,Data!A100:A104)",
            "number",
        ),
        case(
            "FORECAST_basic",
            "FORECAST(30,Data!A105:A109,Data!A100:A104)",
            "=FORECAST(30,Data!A105:A109,Data!A100:A104)",
            "number",
        ),
        case2(
            "FREQUENCY_sum",
            "SUM(FREQUENCY(Data!A100:A109,{20,30,40}))",
            "=SUM(FREQUENCY(Data!A100:A109,{20,30,40}))",
            "number",
        ),
        case(
            "CONVERT_temp",
            "CONVERT(100,\"C\",\"F\")",
            "=CONVERT(100,\"C\",\"F\")",
            "number",
        ),
        case(
            "CONVERT_dist",
            "CONVERT(1,\"mi\",\"km\")",
            "=CONVERT(1,\"mi\",\"km\")",
            "number",
        ),
        case("DELTA_equal", "DELTA(5,5)", "=DELTA(5,5)", "number"),
        case("GESTEP_basic", "GESTEP(5,4)", "=GESTEP(5,4)", "number"),
        case(
            "VLOOKUP_wildcard_star",
            "VLOOKUP(\"*berry\",Data!A16:B25,2,FALSE)",
            "=VLOOKUP(\"*berry\",Data!A16:B25,2,FALSE)",
            "number",
        ),
        case(
            "VLOOKUP_wildcard_question",
            "VLOOKUP(\"d?te\",Data!A16:B25,2,FALSE)",
            "=VLOOKUP(\"d?te\",Data!A16:B25,2,FALSE)",
            "number",
        ),
        case(
            "VLOOKUP_wildcard_no_match",
            "VLOOKUP(\"z*\",Data!A16:B25,2,FALSE)",
            "=VLOOKUP(\"z*\",Data!A16:B25,2,FALSE)",
            "error",
        ),
        case(
            "HLOOKUP_wildcard_star",
            "HLOOKUP(\"*berry\",Data!A16:B25,2,FALSE)",
            "=HLOOKUP(\"*berry\",Data!A16:B25,2,FALSE)",
            "error",
        ),
        case(
            "MATCH_wildcard_question",
            "MATCH(\"ch???y\",Data!A16:A25,0)",
            "=MATCH(\"ch???y\",Data!A16:A25,0)",
            "number",
        ),
        case(
            "MATCH_wildcard_combined",
            "MATCH(\"g*p?\",Data!A16:A25,0)",
            "=MATCH(\"g*p?\",Data!A16:A25,0)",
            "number",
        ),
        case(
            "MATCH_wildcard_no_match",
            "MATCH(\"z*\",Data!A16:A25,0)",
            "=MATCH(\"z*\",Data!A16:A25,0)",
            "error",
        ),
        // Tier 1: broad function coverage
        case("ACOSH_2", "ACOSH(2)", "=ACOSH(2)", "number"),
        case("ACOT_1", "ACOT(1)", "=ACOT(1)", "number"),
        case("ACOTH_2", "ACOTH(2)", "=ACOTH(2)", "number"),
        case("ASINH_1", "ASINH(1)", "=ASINH(1)", "number"),
        case("ATANH_half", "ATANH(0.5)", "=ATANH(0.5)", "number"),
        case("COT_1", "COT(1)", "=COT(1)", "number"),
        case("COTH_2", "COTH(2)", "=COTH(2)", "number"),
        case("CSC_1", "CSC(1)", "=CSC(1)", "number"),
        case("CSCH_1", "CSCH(1)", "=CSCH(1)", "number"),
        case("SEC_0", "SEC(0)", "=SEC(0)", "number"),
        case("SECH_0", "SECH(0)", "=SECH(0)", "number"),
        case(
            "NORM_DIST_std",
            "NORM.DIST(0,0,1,TRUE)",
            "=NORM.DIST(0,0,1,TRUE)",
            "number",
        ),
        case(
            "NORM_DIST_pdf",
            "NORM.DIST(0,0,1,FALSE)",
            "=NORM.DIST(0,0,1,FALSE)",
            "number",
        ),
        case(
            "NORM_INV_50",
            "NORM.INV(0.5,0,1)",
            "=NORM.INV(0.5,0,1)",
            "number",
        ),
        case(
            "NORM_S_DIST_0",
            "NORM.S.DIST(0,TRUE)",
            "=NORM.S.DIST(0,TRUE)",
            "number",
        ),
        case(
            "NORM_S_INV_975",
            "NORM.S.INV(0.975)",
            "=NORM.S.INV(0.975)",
            "number",
        ),
        case(
            "BINOM_DIST_basic",
            "BINOM.DIST(3,10,0.5,FALSE)",
            "=BINOM.DIST(3,10,0.5,FALSE)",
            "number",
        ),
        case(
            "BINOM_DIST_cum",
            "BINOM.DIST(3,10,0.5,TRUE)",
            "=BINOM.DIST(3,10,0.5,TRUE)",
            "number",
        ),
        case(
            "POISSON_DIST_5",
            "POISSON.DIST(5,5,FALSE)",
            "=POISSON.DIST(5,5,FALSE)",
            "number",
        ),
        case(
            "POISSON_DIST_cum",
            "POISSON.DIST(5,5,TRUE)",
            "=POISSON.DIST(5,5,TRUE)",
            "number",
        ),
        case(
            "EXPON_DIST_basic",
            "EXPON.DIST(1,1,TRUE)",
            "=EXPON.DIST(1,1,TRUE)",
            "number",
        ),
        case("GAMMA_basic", "GAMMA(5)", "=GAMMA(5)", "number"),
        case(
            "GAMMA_DIST_basic",
            "GAMMA.DIST(2,2,1,TRUE)",
            "=GAMMA.DIST(2,2,1,TRUE)",
            "number",
        ),
        case("GAMMALN_basic", "GAMMALN(5)", "=GAMMALN(5)", "number"),
        case(
            "GAMMALN_PREC",
            "GAMMALN.PRECISE(5)",
            "=GAMMALN.PRECISE(5)",
            "number",
        ),
        case(
            "BETA_DIST_basic",
            "BETA.DIST(0.5,2,3,TRUE)",
            "=BETA.DIST(0.5,2,3,TRUE)",
            "number",
        ),
        case(
            "CHISQ_DIST_basic",
            "CHISQ.DIST(5,3,TRUE)",
            "=CHISQ.DIST(5,3,TRUE)",
            "number",
        ),
        case(
            "CHISQ_INV_basic",
            "CHISQ.INV(0.95,3)",
            "=CHISQ.INV(0.95,3)",
            "number",
        ),
        case(
            "F_DIST_basic",
            "F.DIST(2,5,10,TRUE)",
            "=F.DIST(2,5,10,TRUE)",
            "number",
        ),
        case(
            "F_INV_basic",
            "F.INV(0.95,5,10)",
            "=F.INV(0.95,5,10)",
            "number",
        ),
        case(
            "T_DIST_basic",
            "T.DIST(2,10,TRUE)",
            "=T.DIST(2,10,TRUE)",
            "number",
        ),
        case(
            "T_INV_basic",
            "T.INV(0.975,10)",
            "=T.INV(0.975,10)",
            "number",
        ),
        case(
            "WEIBULL_DIST_b",
            "WEIBULL.DIST(2,3,1,TRUE)",
            "=WEIBULL.DIST(2,3,1,TRUE)",
            "number",
        ),
        case(
            "LOGNORM_DIST_b",
            "LOGNORM.DIST(1,0,1,TRUE)",
            "=LOGNORM.DIST(1,0,1,TRUE)",
            "number",
        ),
        case(
            "HYPGEOM_DIST_b",
            "HYPGEOM.DIST(1,4,3,10,TRUE)",
            "=HYPGEOM.DIST(1,4,3,10,TRUE)",
            "number",
        ),
        case(
            "NEGBINOM_DIST_b",
            "NEGBINOM.DIST(3,5,0.5,TRUE)",
            "=NEGBINOM.DIST(3,5,0.5,TRUE)",
            "number",
        ),
        case(
            "CONFIDENCE_NORM",
            "CONFIDENCE.NORM(0.05,1,100)",
            "=CONFIDENCE.NORM(0.05,1,100)",
            "number",
        ),
        case("FISHER_basic", "FISHER(0.5)", "=FISHER(0.5)", "number"),
        case(
            "FISHERINV_basic",
            "FISHERINV(0.5)",
            "=FISHERINV(0.5)",
            "number",
        ),
        case(
            "STANDARDIZE_b",
            "STANDARDIZE(42,40,2)",
            "=STANDARDIZE(42,40,2)",
            "number",
        ),
        case("PHI_basic", "PHI(0)", "=PHI(0)", "number"),
        case("GAUSS_basic", "GAUSS(1)", "=GAUSS(1)", "number"),
        case(
            "PERCENTRANK_INC",
            "PERCENTRANK.INC(Data!A100:A109,25)",
            "=PERCENTRANK.INC(Data!A100:A109,25)",
            "number",
        ),
        case(
            "DDB_basic",
            "DDB(Data!A235,Data!B235,Data!C235,1)",
            "=DDB(Data!A235,Data!B235,Data!C235,1)",
            "number",
        ),
        case(
            "SYD_basic",
            "SYD(Data!A235,Data!B235,Data!C235,1)",
            "=SYD(Data!A235,Data!B235,Data!C235,1)",
            "number",
        ),
        case(
            "MIRR_basic",
            "MIRR(Data!A44:A48,0.1,0.12)",
            "=MIRR(Data!A44:A48,0.1,0.12)",
            "number",
        ),
        case(
            "XIRR_basic",
            "XIRR({-1000,300,420,680},{44927,45292,45658,46023})",
            "=XIRR({-1000,300,420,680},{44927,45292,45658,46023})",
            "number",
        ),
        case(
            "XNPV_basic",
            "XNPV(0.1,{-1000,300,420,680},{44927,45292,45658,46023})",
            "=XNPV(0.1,{-1000,300,420,680},{44927,45292,45658,46023})",
            "number",
        ),
        case(
            "DOLLARDE_basic",
            "DOLLARDE(1.02,16)",
            "=DOLLARDE(1.02,16)",
            "number",
        ),
        case(
            "DOLLARFR_basic",
            "DOLLARFR(1.125,16)",
            "=DOLLARFR(1.125,16)",
            "number",
        ),
        case(
            "DURATION_basic",
            "DURATION(Data!A234,Data!B234,Data!B233,Data!C233,2)",
            "=DURATION(Data!A234,Data!B234,Data!B233,Data!C233,2)",
            "number",
        ),
        case(
            "FVSCHEDULE_b",
            "FVSCHEDULE(10000,{0.05,0.06,0.07})",
            "=FVSCHEDULE(10000,{0.05,0.06,0.07})",
            "number",
        ),
        case(
            "ISPMT_basic",
            "ISPMT(0.06/12,1,360,200000)",
            "=ISPMT(0.06/12,1,360,200000)",
            "number",
        ),
        case(
            "PDURATION_basic",
            "PDURATION(0.05,1000,2000)",
            "=PDURATION(0.05,1000,2000)",
            "number",
        ),
        case(
            "RRI_basic",
            "RRI(10,1000,2000)",
            "=RRI(10,1000,2000)",
            "number",
        ),
        case2(
            "DROP_top2",
            "SUM(DROP(Data!A264:E264,0,2))",
            "=SUM(DROP(Data!A264:E264,0,2))",
            "number",
        ),
        case2(
            "TAKE_left3",
            "SUM(TAKE(Data!A264:E264,1,3))",
            "=SUM(TAKE(Data!A264:E264,1,3))",
            "number",
        ),
        case2(
            "HSTACK_basic",
            "INDEX(HSTACK(Data!A267:A268,Data!C267:C268),1,2)",
            "=INDEX(HSTACK(Data!A267:A268,Data!C267:C268),1,2)",
            "string",
        ),
        case2(
            "VSTACK_basic",
            "INDEX(VSTACK(Data!A267:C267,Data!A268:C268),2,1)",
            "=INDEX(VSTACK(Data!A267:C267,Data!A268:C268),2,1)",
            "string",
        ),
        case2(
            "TOCOL_basic",
            "INDEX(TOCOL(Data!A264:C264),2)",
            "=INDEX(TOCOL(Data!A264:C264),2)",
            "number",
        ),
        case2(
            "TOROW_basic",
            "INDEX(TOROW(Data!A264:A266),2)",
            "=INDEX(TOROW(Data!A264:A266),2)",
            "number",
        ),
        case2(
            "WRAPCOLS_basic",
            "INDEX(WRAPCOLS(SEQUENCE(6),3),1,2)",
            "=INDEX(WRAPCOLS(SEQUENCE(6),3),1,2)",
            "number",
        ),
        case2(
            "WRAPROWS_basic",
            "INDEX(WRAPROWS(SEQUENCE(6),3),2,1)",
            "=INDEX(WRAPROWS(SEQUENCE(6),3),2,1)",
            "number",
        ),
        case2(
            "CHOOSEROWS_b",
            "INDEX(CHOOSEROWS(Data!A264:E266,1,3),2,1)",
            "=INDEX(CHOOSEROWS(Data!A264:E266,1,3),2,1)",
            "number",
        ),
        case2(
            "EXPAND_basic",
            "COLUMNS(EXPAND(Data!A264:C264,1,5))",
            "=COLUMNS(EXPAND(Data!A264:C264,1,5))",
            "number",
        ),
        case(
            "DSUM_basic",
            "DSUM(Data!A150:D160,3,Data!A162:B163)",
            "=DSUM(Data!A150:D160,3,Data!A162:B163)",
            "number",
        ),
        case(
            "DCOUNT_basic",
            "DCOUNT(Data!A150:D160,3,Data!A162:B163)",
            "=DCOUNT(Data!A150:D160,3,Data!A162:B163)",
            "number",
        ),
        case(
            "DAVERAGE_basic",
            "DAVERAGE(Data!A150:D160,3,Data!A162:B163)",
            "=DAVERAGE(Data!A150:D160,3,Data!A162:B163)",
            "number",
        ),
        case(
            "DMAX_basic",
            "DMAX(Data!A150:D160,3,Data!A162:B163)",
            "=DMAX(Data!A150:D160,3,Data!A162:B163)",
            "number",
        ),
        case(
            "DMIN_basic",
            "DMIN(Data!A150:D160,3,Data!A162:B163)",
            "=DMIN(Data!A150:D160,3,Data!A162:B163)",
            "number",
        ),
        case(
            "DGET_basic",
            "DGET(Data!A150:D160,3,Data!F162:G163)",
            "=DGET(Data!A150:D160,3,Data!F162:G163)",
            "number",
        ),
        case(
            "DSTDEV_basic",
            "DSTDEV(Data!A150:D160,3,Data!A162:B163)",
            "=DSTDEV(Data!A150:D160,3,Data!A162:B163)",
            "number",
        ),
        case(
            "DVAR_basic",
            "DVAR(Data!A150:D160,3,Data!A162:B163)",
            "=DVAR(Data!A150:D160,3,Data!A162:B163)",
            "number",
        ),
        case2(
            "TEXTSPLIT_basic",
            "INDEX(TEXTSPLIT(\"hello-world-test\",\"-\"),1)",
            "=INDEX(TEXTSPLIT(\"hello-world-test\",\"-\"),1)",
            "string",
        ),
        case(
            "ERROR_TYPE_div0",
            "ERROR.TYPE(1/0)",
            "=ERROR.TYPE(1/0)",
            "number",
        ),
        case("ISERR_yes", "ISERR(1/0)", "=ISERR(1/0)", "boolean"),
        case(
            "ISNA_yes",
            "ISNA(MATCH(\"zzz\",Data!A16:A25,0))",
            "=ISNA(MATCH(\"zzz\",Data!A16:A25,0))",
            "boolean",
        ),
        case("ISEVEN_yes", "ISEVEN(4)", "=ISEVEN(4)", "boolean"),
        case("ISODD_yes", "ISODD(3)", "=ISODD(3)", "boolean"),
        case(
            "ISNONTEXT_num",
            "ISNONTEXT(42)",
            "=ISNONTEXT(42)",
            "boolean",
        ),
        case(
            "ISFORMULA_yes",
            "ISFORMULA(Data!A121)",
            "=ISFORMULA(Data!A121)",
            "boolean",
        ),
        case("BITAND_basic", "BITAND(12,10)", "=BITAND(12,10)", "number"),
        case("BITOR_basic", "BITOR(12,10)", "=BITOR(12,10)", "number"),
        case("BITXOR_basic", "BITXOR(12,10)", "=BITXOR(12,10)", "number"),
        case(
            "BITLSHIFT_basic",
            "BITLSHIFT(4,2)",
            "=BITLSHIFT(4,2)",
            "number",
        ),
        case(
            "BITRSHIFT_basic",
            "BITRSHIFT(16,2)",
            "=BITRSHIFT(16,2)",
            "number",
        ),
        case(
            "COMPLEX_basic",
            "IMABS(COMPLEX(3,4))",
            "=IMABS(COMPLEX(3,4))",
            "number",
        ),
        case("ERF_basic", "ERF(1)", "=ERF(1)", "number"),
        case("ERFC_basic", "ERFC(1)", "=ERFC(1)", "number"),
        case("BESSELI_basic", "BESSELI(1,0)", "=BESSELI(1,0)", "number"),
        case("BESSELJ_basic", "BESSELJ(1,0)", "=BESSELJ(1,0)", "number"),
        case(
            "IMABS_basic",
            "IMABS(\"3+4i\")",
            "=IMABS(\"3+4i\")",
            "number",
        ),
        case(
            "IMREAL_basic",
            "IMREAL(\"3+4i\")",
            "=IMREAL(\"3+4i\")",
            "number",
        ),
        case(
            "IMAGINARY_basic",
            "IMAGINARY(\"3+4i\")",
            "=IMAGINARY(\"3+4i\")",
            "number",
        ),
        case(
            "IMSUM_basic",
            "IMABS(IMSUM(\"3+4i\",\"1-2i\"))",
            "=IMABS(IMSUM(\"3+4i\",\"1-2i\"))",
            "number",
        ),
        case(
            "IMSUB_basic",
            "IMABS(IMSUB(\"3+4i\",\"1+2i\"))",
            "=IMABS(IMSUB(\"3+4i\",\"1+2i\"))",
            "number",
        ),
        // Tier 2: edge cases
        case(
            "VLOOKUP_empty_str",
            "VLOOKUP(\"\",Data!A16:B25,2,FALSE)",
            "=VLOOKUP(\"\",Data!A16:B25,2,FALSE)",
            "error",
        ),
        case(
            "VLOOKUP_num_str",
            "VLOOKUP(42,Data!A16:B25,2,FALSE)",
            "=VLOOKUP(42,Data!A16:B25,2,FALSE)",
            "error",
        ),
        case(
            "VLOOKUP_tilde_esc",
            "VLOOKUP(\"~*\",Data!A16:B25,2,FALSE)",
            "=VLOOKUP(\"~*\",Data!A16:B25,2,FALSE)",
            "error",
        ),
        case(
            "MATCH_last_dup",
            "MATCH(\"apple\",{\"apple\",\"banana\",\"apple\"},0)",
            "=MATCH(\"apple\",{\"apple\",\"banana\",\"apple\"},0)",
            "number",
        ),
        case(
            "MATCH_boolean",
            "MATCH(TRUE,{FALSE,TRUE,FALSE},0)",
            "=MATCH(TRUE,{FALSE,TRUE,FALSE},0)",
            "number",
        ),
        case(
            "INDEX_zero_both",
            "INDEX({1,2,3;4,5,6},0,0)",
            "=INDEX({1,2,3;4,5,6},0,0)",
            "number",
        ),
        case(
            "XLOOKUP_horiz",
            "XLOOKUP(5,Data!A1:L1,Data!A2:L2)",
            "=XLOOKUP(5,Data!A1:L1,Data!A2:L2)",
            "number",
        ),
        case(
            "XLOOKUP_binary_asc",
            "XLOOKUP(50,Data!A2:L2,Data!A1:L1,,0,2)",
            "=XLOOKUP(50,Data!A2:L2,Data!A1:L1,,0,2)",
            "number",
        ),
        case(
            "INDIRECT_dynamic",
            "INDIRECT(\"Data!A\"&50)",
            "=INDIRECT(\"Data!A\"&50)",
            "number",
        ),
        case(
            "INDIRECT_invalid",
            "INDIRECT(\"ZZZZ!A1\")",
            "=INDIRECT(\"ZZZZ!A1\")",
            "error",
        ),
        case("DATE_leap", "DATE(2024,2,29)", "=DATE(2024,2,29)", "number"),
        case(
            "DATE_noleap",
            "DATE(2023,2,29)",
            "=DATE(2023,2,29)",
            "number",
        ),
        case(
            "DATE_month_overflow",
            "DATE(2024,13,1)",
            "=DATE(2024,13,1)",
            "number",
        ),
        case(
            "DATE_neg_month",
            "DATE(2024,-1,1)",
            "=DATE(2024,-1,1)",
            "number",
        ),
        case(
            "DATEDIF_same",
            "DATEDIF(DATE(2024,1,1),DATE(2024,1,1),\"D\")",
            "=DATEDIF(DATE(2024,1,1),DATE(2024,1,1),\"D\")",
            "number",
        ),
        case(
            "DATEDIF_ym",
            "DATEDIF(DATE(2024,1,15),DATE(2024,7,20),\"YM\")",
            "=DATEDIF(DATE(2024,1,15),DATE(2024,7,20),\"YM\")",
            "number",
        ),
        case(
            "EDATE_year_wrap",
            "EDATE(DATE(2024,11,15),3)",
            "=EDATE(DATE(2024,11,15),3)",
            "number",
        ),
        case(
            "EOMONTH_leap",
            "EOMONTH(DATE(2024,1,15),1)",
            "=EOMONTH(DATE(2024,1,15),1)",
            "number",
        ),
        case("MOD_neg_divisor", "MOD(7,-3)", "=MOD(7,-3)", "number"),
        case("MOD_zero", "MOD(0,5)", "=MOD(0,5)", "number"),
        case("POWER_zero_zero", "POWER(0,0)", "=POWER(0,0)", "error"),
        case("LOG_zero", "LOG(0)", "=LOG(0)", "error"),
        case("SQRT_neg", "SQRT(-1)", "=SQRT(-1)", "error"),
        case("LN_zero", "LN(0)", "=LN(0)", "error"),
        case("DIV_zero_zero", "0/0", "=0/0", "error"),
        case(
            "LARGE_neg_k",
            "LARGE(Data!A2:L2,-1)",
            "=LARGE(Data!A2:L2,-1)",
            "error",
        ),
        case(
            "ROUND_neg_digits",
            "ROUND(1234.5,-2)",
            "=ROUND(1234.5,-2)",
            "number",
        ),
        case("INT_negative", "INT(-3.1)", "=INT(-3.1)", "number"),
        case(
            "SUMIFS_no_match",
            "SUMIFS(Data!C29:C37,Data!B29:B37,\"nonexist\")",
            "=SUMIFS(Data!C29:C37,Data!B29:B37,\"nonexist\")",
            "number",
        ),
        case(
            "AVERAGEIF_no_match",
            "AVERAGEIF(Data!B29:B37,\"nonexist\",Data!C29:C37)",
            "=AVERAGEIF(Data!B29:B37,\"nonexist\",Data!C29:C37)",
            "error",
        ),
        case(
            "COUNTIFS_gt_crit",
            "COUNTIFS(Data!C29:C37,\">15\")",
            "=COUNTIFS(Data!C29:C37,\">15\")",
            "number",
        ),
        case(
            "COUNTIFS_ne_crit",
            "COUNTIFS(Data!B29:B37,\"<>fruit\")",
            "=COUNTIFS(Data!B29:B37,\"<>fruit\")",
            "number",
        ),
        case(
            "SUMPRODUCT_bool",
            "SUMPRODUCT(--(Data!B29:B37=\"fruit\"))",
            "=SUMPRODUCT(--(Data!B29:B37=\"fruit\"))",
            "number",
        ),
        case(
            "SUM_all_empty",
            "SUM(Data!B40,Data!D40)",
            "=SUM(Data!B40,Data!D40)",
            "number",
        ),
        case(
            "AVERAGE_all_empty",
            "AVERAGE(Data!B40,Data!D40)",
            "=AVERAGE(Data!B40,Data!D40)",
            "error",
        ),
        case(
            "COUNT_all_empty",
            "COUNT(Data!B40,Data!D40)",
            "=COUNT(Data!B40,Data!D40)",
            "number",
        ),
        case(
            "IFERROR_chain",
            "IFERROR(IFERROR(1/0,SQRT(-1)),\"caught\")",
            "=IFERROR(IFERROR(1/0,SQRT(-1)),\"caught\")",
            "string",
        ),
        case(
            "IF_error_cond",
            "IF(ISERROR(1/0),\"err\",\"ok\")",
            "=IF(ISERROR(1/0),\"err\",\"ok\")",
            "string",
        ),
        case("AND_with_error", "AND(TRUE,1/0)", "=AND(TRUE,1/0)", "error"),
        case("OR_short_circuit", "OR(TRUE,1/0)", "=OR(TRUE,1/0)", "error"),
        case(
            "SWITCH_error_match",
            "SWITCH(1/0,1,\"one\",\"err\")",
            "=SWITCH(1/0,1,\"one\",\"err\")",
            "error",
        ),
    ]
}

fn append_tier3_cache_cases(cases: &mut Vec<FormulaCase>) {
    // Template cache: 50 formulas with identical structure, different absolute rows.
    for i in 0u32..50 {
        let row = 170 + i;
        cases.push(case(
            &format!("TCACHE_ref_{i}"),
            &format!("Data!B{row}"),
            &format!("=Data!B{row}"),
            "number",
        ));
    }

    // Binary-op cache: repeated shape across sparse rows.
    for i in 0u32..25 {
        let row = 170 + i * 2;
        cases.push(case(
            &format!("TCACHE_sum_{i}"),
            &format!("Data!A{row}+Data!B{row}"),
            &format!("=Data!A{row}+Data!B{row}"),
            "number",
        ));
    }

    // Lookup cache: the full fruit table through repeated exact VLOOKUPs.
    let fruits = [
        "apple",
        "banana",
        "cherry",
        "date",
        "elderberry",
        "fig",
        "grape",
        "honeydew",
        "kiwi",
        "lemon",
    ];
    for fruit in fruits {
        cases.push(case(
            &format!("LCACHE_vlookup_{fruit}"),
            &format!("VLOOKUP(\"{fruit}\",Data!A16:B25,2,FALSE)"),
            &format!("=VLOOKUP(\"{fruit}\",Data!A16:B25,2,FALSE)"),
            "number",
        ));
    }

    // MATCH cache on the same range and lookup shape.
    for fruit in fruits {
        cases.push(case(
            &format!("LCACHE_match_{fruit}"),
            &format!("MATCH(\"{fruit}\",Data!A16:A25,0)"),
            &format!("=MATCH(\"{fruit}\",Data!A16:A25,0)"),
            "number",
        ));
    }

    // Overlapping range cache: different slice lengths of the same stress column.
    for n in 2u32..=26 {
        let end = 170 + n;
        cases.push(case(
            &format!("RCACHE_sum_{n}"),
            &format!("SUM(Data!B170:B{end})"),
            &format!("=SUM(Data!B170:B{end})"),
            "number",
        ));
    }

    // Additional string reference cache over the 50-row stress block.
    for i in 0u32..50 {
        let row = 170 + i;
        cases.push(case(
            &format!("TCACHE_cat_{i}"),
            &format!("Data!C{row}"),
            &format!("=Data!C{row}"),
            "string",
        ));
    }

    // Additional conditional template cache with the same branch shape.
    for i in 0u32..50 {
        let row = 170 + i;
        cases.push(case(
            &format!("TCACHE_if_{i}"),
            &format!("IF(Data!C{row}=\"alpha\",Data!B{row},0)"),
            &format!("=IF(Data!C{row}=\"alpha\",Data!B{row},0)"),
            "number",
        ));
    }
}

fn append_tier4_validation_cases(cases: &mut Vec<FormulaCase>) {
    // Trig identity: SIN²(x) + COS²(x) = 1.
    for i in 0u32..10 {
        let row = 252 + i;
        cases.push(case(
            &format!("IDENTITY_sincos_{i}"),
            &format!("ROUND(SIN(Data!A{row})^2+COS(Data!A{row})^2,10)"),
            &format!("=ROUND(SIN(Data!A{row})^2+COS(Data!A{row})^2,10)"),
            "number",
        ));
    }

    // Cumulative SUM consistency across the validation range.
    for n in 0u32..10 {
        let end = 240 + n;
        cases.push(case(
            &format!("CUMSUM_{n}"),
            &format!("SUM(Data!A240:A{end})"),
            &format!("=SUM(Data!A240:A{end})"),
            "number",
        ));
    }

    // Cross-sheet addition matrix over the 50-row stress block.
    for i in 0u32..50 {
        let row = 170 + i;
        cases.push(case(
            &format!("XSHEET_add_{i}"),
            &format!("Data!A{row}+Data!B{row}"),
            &format!("=Data!A{row}+Data!B{row}"),
            "number",
        ));
    }

    // INDEX validation at different positions over the same range.
    for i in 1u32..=20 {
        cases.push(case(
            &format!("IDX_pos_{i}"),
            &format!("INDEX(Data!B170:B219,{i})"),
            &format!("=INDEX(Data!B170:B219,{i})"),
            "number",
        ));
    }

    // VLOOKUP validation with every key in the 50-row table.
    for i in 1u32..=50 {
        cases.push(case(
            &format!("VL50_{i}"),
            &format!("VLOOKUP({i},Data!A170:B219,2,FALSE)"),
            &format!("=VLOOKUP({i},Data!A170:B219,2,FALSE)"),
            "number",
        ));
    }

    // MATCH validation for every key in the 50-row table.
    for i in 1u32..=50 {
        cases.push(case(
            &format!("MATCH50_{i}"),
            &format!("MATCH({i},Data!A170:A219,0)"),
            &format!("=MATCH({i},Data!A170:A219,0)"),
            "number",
        ));
    }

    // Category position validation across the first half of the stress table.
    for i in 1u32..=25 {
        cases.push(case(
            &format!("CATIDX_{i}"),
            &format!("INDEX(Data!C170:C219,{i})"),
            &format!("=INDEX(Data!C170:C219,{i})"),
            "string",
        ));
    }
}

fn copy_fixture_into_repo(host_path: &Path) -> std::io::Result<()> {
    let repo_root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let output_path = repo_root.join(REPO_FIXTURE_RELATIVE_PATH);
    if let Some(parent) = output_path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    let _ = std::fs::copy(host_path, output_path)?;
    Ok(())
}

fn cell_addr(row: u32, col: u32) -> String {
    format!("{}{}", column_name(col), row)
}

fn column_name(mut col: u32) -> String {
    let mut chars = Vec::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        chars.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    chars.iter().rev().collect()
}
