use std::path::{Path, PathBuf};

use crate::{ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_excel_com::{BridgeError, ChainStep, Workbook};
use excel_com_protocol::{ResponseData, SheetRef};
use serde_json::json;

const VM_FIXTURE_PATH: &str = r"C:\temp\formula_parity.xlsx";
const HOST_FIXTURE_PATH: &str = "/tmp/duke-sheets-excel/formula_parity.xlsx";
const REPO_FIXTURE_RELATIVE_PATH: &str = "data/formula-parity.xlsx";

#[derive(Clone, Copy)]
enum FormulaKind {
    Formula,
    Formula2,
}

struct FormulaCase {
    id: &'static str,
    label: &'static str,
    formula: &'static str,
    expected_type: &'static str,
    kind: FormulaKind,
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
        populate_tests_sheet(&wb).expect("populate Tests sheet");

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

    Ok(())
}

fn populate_tests_sheet(wb: &Workbook<'_>) -> Result<(), BridgeError> {
    wb.set_cell_value("A1", "test_id")?;
    wb.set_cell_value("B1", "formula_label")?;
    wb.set_cell_value("C1", "formula")?;
    wb.set_cell_value("D1", "expected_type")?;

    let cases = [
        FormulaCase {
            id: "INDEX_2arg_single_row_3",
            label: "INDEX(Data!A2:L2,3)",
            formula: "=INDEX(Data!A2:L2,3)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_2arg_single_col_4",
            label: "INDEX(Data!A3:A14,4)",
            formula: "=INDEX(Data!A3:A14,4)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_3arg_matrix_q2_west",
            label: "INDEX(Data!E17:G20,2,3)",
            formula: "=INDEX(Data!E17:G20,2,3)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_3arg_row_vector_col_5",
            label: "INDEX(Data!A2:L2,1,5)",
            formula: "=INDEX(Data!A2:L2,1,5)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_MATCH_single_row_60",
            label: "INDEX(Data!A2:L2,MATCH(60,Data!A2:L2,0))",
            formula: "=INDEX(Data!A2:L2,MATCH(60,Data!A2:L2,0))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_exact_numeric_50",
            label: "MATCH(50,Data!A2:L2,0)",
            formula: "=MATCH(50,Data!A2:L2,0)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_approx_ascending_55",
            label: "MATCH(55,Data!A2:L2,1)",
            formula: "=MATCH(55,Data!A2:L2,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_approx_descending_55",
            label: "MATCH(55,Data!N2:N13,-1)",
            formula: "=MATCH(55,Data!N2:N13,-1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_exact_string_cherry",
            label: "MATCH(\"cherry\",Data!A16:A25,0)",
            formula: "=MATCH(\"cherry\",Data!A16:A25,0)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_not_found_orange",
            label: "MATCH(\"orange\",Data!A16:A25,0)",
            formula: "=MATCH(\"orange\",Data!A16:A25,0)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_exact_banana_price",
            label: "VLOOKUP(\"banana\",Data!A16:B25,2,FALSE)",
            formula: "=VLOOKUP(\"banana\",Data!A16:B25,2,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_approx_grade_88",
            label: "VLOOKUP(88,Data!P17:Q21,2,TRUE)",
            formula: "=VLOOKUP(88,Data!P17:Q21,2,TRUE)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_diff_col_q3_west",
            label: "VLOOKUP(\"Q3\",Data!D17:G20,4,FALSE)",
            formula: "=VLOOKUP(\"Q3\",Data!D17:G20,4,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_not_found_orange",
            label: "VLOOKUP(\"orange\",Data!A16:B25,2,FALSE)",
            formula: "=VLOOKUP(\"orange\",Data!A16:B25,2,FALSE)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "HLOOKUP_exact_month_4",
            label: "HLOOKUP(4,Data!A1:L2,2,FALSE)",
            formula: "=HLOOKUP(4,Data!A1:L2,2,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "HLOOKUP_approx_month_5_5",
            label: "HLOOKUP(5.5,Data!A1:L2,2,TRUE)",
            formula: "=HLOOKUP(5.5,Data!A1:L2,2,TRUE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "HLOOKUP_not_found_month_13",
            label: "HLOOKUP(13,Data!A1:L2,2,FALSE)",
            formula: "=HLOOKUP(13,Data!A1:L2,2,FALSE)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUMIFS_single_category_fruit",
            label: "SUMIFS(Data!C29:C37,Data!B29:B37,\"fruit\")",
            formula: "=SUMIFS(Data!C29:C37,Data!B29:B37,\"fruit\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUMIFS_multi_category_and_name",
            label: "SUMIFS(Data!C29:C37,Data!B29:B37,\"fruit\",Data!A29:A37,\"*a*\")",
            formula: "=SUMIFS(Data!C29:C37,Data!B29:B37,\"fruit\",Data!A29:A37,\"*a*\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "COUNTIFS_single_category_veg",
            label: "COUNTIFS(Data!B29:B37,\"veg\")",
            formula: "=COUNTIFS(Data!B29:B37,\"veg\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "COUNTIFS_multi_wildcard_i",
            label: "COUNTIFS(Data!B29:B37,\"fruit\",Data!A29:A37,\"*i*\")",
            formula: "=COUNTIFS(Data!B29:B37,\"fruit\",Data!A29:A37,\"*i*\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IF_basic_false",
            label: "IF(Data!A2>50,\"big\",\"small\")",
            formula: "=IF(Data!A2>50,\"big\",\"small\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IF_basic_true",
            label: "IF(Data!L2>50,\"big\",\"small\")",
            formula: "=IF(Data!L2>50,\"big\",\"small\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IF_nested_medium",
            label: "IF(Data!A2>50,\"big\",IF(Data!A2>5,\"medium\",\"small\"))",
            formula: "=IF(Data!A2>50,\"big\",IF(Data!A2>5,\"medium\",\"small\"))",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IFS_multi_condition_medium",
            label: "IFS(Data!A2>50,\"big\",Data!A2>5,\"medium\",TRUE,\"small\")",
            formula: "=IFS(Data!A2>50,\"big\",Data!A2>5,\"medium\",TRUE,\"small\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SWITCH_numeric_two",
            label: "SWITCH(2,1,\"one\",2,\"two\",\"other\")",
            formula: "=SWITCH(2,1,\"one\",2,\"two\",\"other\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SWITCH_default_other",
            label: "SWITCH(\"kiwi\",\"apple\",\"A\",\"banana\",\"B\",\"other\")",
            formula: "=SWITCH(\"kiwi\",\"apple\",\"A\",\"banana\",\"B\",\"other\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUM_row_values",
            label: "SUM(Data!A2:L2)",
            formula: "=SUM(Data!A2:L2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "AVERAGE_row_values",
            label: "AVERAGE(Data!A2:L2)",
            formula: "=AVERAGE(Data!A2:L2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MIN_col_values",
            label: "MIN(Data!A3:A14)",
            formula: "=MIN(Data!A3:A14)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MAX_col_values",
            label: "MAX(Data!A3:A14)",
            formula: "=MAX(Data!A3:A14)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "COUNT_col_values",
            label: "COUNT(Data!A3:A14)",
            formula: "=COUNT(Data!A3:A14)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "COUNTA_lookup_keys",
            label: "COUNTA(Data!A16:A25)",
            formula: "=COUNTA(Data!A16:A25)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUMPRODUCT_months_and_values",
            label: "SUMPRODUCT(Data!A1:L1,Data!A2:L2)",
            formula: "=SUMPRODUCT(Data!A1:L1,Data!A2:L2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ROUND_basic",
            label: "ROUND(1.2345,2)",
            formula: "=ROUND(1.2345,2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ROUNDUP_basic",
            label: "ROUNDUP(1.231,2)",
            formula: "=ROUNDUP(1.231,2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ROUNDDOWN_basic",
            label: "ROUNDDOWN(1.239,2)",
            formula: "=ROUNDDOWN(1.239,2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LEFT_banana_3",
            label: "LEFT(\"banana\",3)",
            formula: "=LEFT(\"banana\",3)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "RIGHT_banana_2",
            label: "RIGHT(\"banana\",2)",
            formula: "=RIGHT(\"banana\",2)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MID_elderberry_3_5",
            label: "MID(\"elderberry\",3,5)",
            formula: "=MID(\"elderberry\",3,5)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LEN_honeydew",
            label: "LEN(\"honeydew\")",
            formula: "=LEN(\"honeydew\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "FIND_berry_in_elderberry",
            label: "FIND(\"berry\",\"elderberry\")",
            formula: "=FIND(\"berry\",\"elderberry\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SEARCH_err_in_elderberry",
            label: "SEARCH(\"ERR\",\"elderberry\")",
            formula: "=SEARCH(\"ERR\",\"elderberry\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CONCATENATE_apple_pie",
            label: "CONCATENATE(\"apple\",\"-\",\"pie\")",
            formula: "=CONCATENATE(\"apple\",\"-\",\"pie\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CONCAT_duke_sheets",
            label: "CONCAT(\"duke\",\"-\",\"sheets\")",
            formula: "=CONCAT(\"duke\",\"-\",\"sheets\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TEXTJOIN_skip_empty",
            label: "TEXTJOIN(\",\",TRUE,\"apple\",\"\",\"banana\")",
            formula: "=TEXTJOIN(\",\",TRUE,\"apple\",\"\",\"banana\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "UPPER_kiwi",
            label: "UPPER(\"Kiwi\")",
            formula: "=UPPER(\"Kiwi\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LOWER_lemon",
            label: "LOWER(\"LEMON\")",
            formula: "=LOWER(\"LEMON\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PROPER_honeydew_melon",
            label: "PROPER(\"hONEYDEW MELON\")",
            formula: "=PROPER(\"hONEYDEW MELON\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TRIM_many_spaces",
            label: "TRIM(\"  too   many spaces  \")",
            formula: "=TRIM(\"  too   many spaces  \")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "DATE_2024_02_29",
            label: "DATE(2024,2,29)",
            formula: "=DATE(2024,2,29)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "YEAR_from_date",
            label: "YEAR(DATE(2024,2,29))",
            formula: "=YEAR(DATE(2024,2,29))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MONTH_from_date",
            label: "MONTH(DATE(2024,2,29))",
            formula: "=MONTH(DATE(2024,2,29))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "DAY_from_date",
            label: "DAY(DATE(2024,2,29))",
            formula: "=DAY(DATE(2024,2,29))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TODAY_type_only",
            label: "TODAY()",
            formula: "=TODAY()",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "EDATE_plus_one_month",
            label: "EDATE(DATE(2024,1,31),1)",
            formula: "=EDATE(DATE(2024,1,31),1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "EOMONTH_plus_one_month",
            label: "EOMONTH(DATE(2024,1,15),1)",
            formula: "=EOMONTH(DATE(2024,1,15),1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NETWORKDAYS_basic_range",
            label: "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10))",
            formula: "=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "AND_all_true",
            label: "AND(TRUE,1<2,2<3)",
            formula: "=AND(TRUE,1<2,2<3)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "OR_one_true",
            label: "OR(FALSE,2<1,3=3)",
            formula: "=OR(FALSE,2<1,3=3)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NOT_basic",
            label: "NOT(Data!A2>50)",
            formula: "=NOT(Data!A2>50)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "XOR_three_args",
            label: "XOR(TRUE,FALSE,TRUE)",
            formula: "=XOR(TRUE,FALSE,TRUE)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IFERROR_div_zero_fallback",
            label: "IFERROR(1/0,\"fallback\")",
            formula: "=IFERROR(1/0,\"fallback\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IFNA_match_missing",
            label: "IFNA(MATCH(\"orange\",Data!A16:A25,0),\"missing\")",
            formula: "=IFNA(MATCH(\"orange\",Data!A16:A25,0),\"missing\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "XLOOKUP_exact_kiwi",
            label: "XLOOKUP(\"kiwi\",Data!A16:A25,Data!B16:B25)",
            formula: "=XLOOKUP(\"kiwi\",Data!A16:A25,Data!B16:B25)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "XLOOKUP_if_not_found",
            label: "XLOOKUP(\"orange\",Data!A16:A25,Data!B16:B25,\"missing\")",
            formula: "=XLOOKUP(\"orange\",Data!A16:A25,Data!B16:B25,\"missing\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "XLOOKUP_reverse_last_apple",
            label: "XLOOKUP(\"apple\",Data!M17:M20,Data!N17:N20,\"missing\",0,-1)",
            formula: "=XLOOKUP(\"apple\",Data!M17:M20,Data!N17:N20,\"missing\",0,-1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "FILTER_second_fruit_name",
            label: "INDEX(FILTER(Data!A29:A37,Data!B29:B37=\"fruit\"),2)",
            formula: "=INDEX(FILTER(Data!A29:A37,Data!B29:B37=\"fruit\"),2)",
            expected_type: "string",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "SORT_first_name",
            label: "INDEX(SORT(Data!A29:A37),1)",
            formula: "=INDEX(SORT(Data!A29:A37),1)",
            expected_type: "string",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "UNIQUE_category_count",
            label: "COUNTA(UNIQUE(Data!B29:B37))",
            formula: "=COUNTA(UNIQUE(Data!B29:B37))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "SEQUENCE_sum_1_to_4",
            label: "SUM(SEQUENCE(4,1,1,1))",
            formula: "=SUM(SEQUENCE(4,1,1,1))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "VLOOKUP_approx_below_min",
            label: "VLOOKUP(-5,Data!P17:Q21,2,TRUE)",
            formula: "=VLOOKUP(-5,Data!P17:Q21,2,TRUE)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_approx_above_max",
            label: "VLOOKUP(99,Data!P17:Q21,2,TRUE)",
            formula: "=VLOOKUP(99,Data!P17:Q21,2,TRUE)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_approx_exact_bound",
            label: "VLOOKUP(90,Data!P17:Q21,2,TRUE)",
            formula: "=VLOOKUP(90,Data!P17:Q21,2,TRUE)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_approx_first_match",
            label: "VLOOKUP(0,Data!P17:Q21,2,TRUE)",
            formula: "=VLOOKUP(0,Data!P17:Q21,2,TRUE)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_2arg_row_out_of_bounds",
            label: "INDEX(Data!A2:L2,15)",
            formula: "=INDEX(Data!A2:L2,15)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_2arg_col_out_of_bounds",
            label: "INDEX(Data!A3:A14,15)",
            formula: "=INDEX(Data!A3:A14,15)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_2arg_row_position_1",
            label: "INDEX(Data!A2:L2,1)",
            formula: "=INDEX(Data!A2:L2,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_2arg_col_position_1",
            label: "INDEX(Data!A3:A14,1)",
            formula: "=INDEX(Data!A3:A14,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_2arg_row_last",
            label: "INDEX(Data!A2:L2,12)",
            formula: "=INDEX(Data!A2:L2,12)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MEDIAN_numbers",
            label: "MEDIAN(Data!A2:L2)",
            formula: "=MEDIAN(Data!A2:L2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "STDEV_numbers",
            label: "STDEV(Data!A2:L2)",
            formula: "=STDEV(Data!A2:L2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VAR_numbers",
            label: "VAR(Data!A2:L2)",
            formula: "=VAR(Data!A2:L2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LARGE_3rd",
            label: "LARGE(Data!A2:L2,3)",
            formula: "=LARGE(Data!A2:L2,3)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SMALL_2nd",
            label: "SMALL(Data!A2:L2,2)",
            formula: "=SMALL(Data!A2:L2,2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PERCENTILE_50",
            label: "PERCENTILE(Data!A2:L2,0.5)",
            formula: "=PERCENTILE(Data!A2:L2,0.5)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "QUARTILE_1",
            label: "QUARTILE(Data!A2:L2,1)",
            formula: "=QUARTILE(Data!A2:L2,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "RANK_30",
            label: "RANK(30,Data!A2:L2)",
            formula: "=RANK(30,Data!A2:L2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MODE_SNGL",
            label: "MODE.SNGL(Data!S1:S12)",
            formula: "=MODE.SNGL(Data!S1:S12)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "AVERAGEIF_fruit",
            label: "AVERAGEIF(Data!B29:B37,\"fruit\",Data!C29:C37)",
            formula: "=AVERAGEIF(Data!B29:B37,\"fruit\",Data!C29:C37)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ABS_negative",
            label: "ABS(-42.5)",
            formula: "=ABS(-42.5)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MOD_basic",
            label: "MOD(17,5)",
            formula: "=MOD(17,5)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "POWER_basic",
            label: "POWER(2,10)",
            formula: "=POWER(2,10)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SQRT_basic",
            label: "SQRT(144)",
            formula: "=SQRT(144)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LOG_base10",
            label: "LOG(1000,10)",
            formula: "=LOG(1000,10)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LOG10_basic",
            label: "LOG10(100)",
            formula: "=LOG10(100)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LN_basic",
            label: "LN(EXP(1))",
            formula: "=LN(EXP(1))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CEILING_basic",
            label: "CEILING(2.3,1)",
            formula: "=CEILING(2.3,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "FLOOR_MATH_basic",
            label: "FLOOR.MATH(2.7,1)",
            formula: "=FLOOR.MATH(2.7,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INT_basic",
            label: "INT(3.9)",
            formula: "=INT(3.9)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SIGN_negative",
            label: "SIGN(-42)",
            formula: "=SIGN(-42)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PI_value",
            label: "PI()",
            formula: "=PI()",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "RAND_type",
            label: "RAND()",
            formula: "=RAND()",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "RANDBETWEEN_type",
            label: "RANDBETWEEN(1,100)",
            formula: "=RANDBETWEEN(1,100)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISNUMBER_yes",
            label: "ISNUMBER(Data!A50)",
            formula: "=ISNUMBER(Data!A50)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISNUMBER_no",
            label: "ISNUMBER(Data!A51)",
            formula: "=ISNUMBER(Data!A51)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISTEXT_yes",
            label: "ISTEXT(Data!A51)",
            formula: "=ISTEXT(Data!A51)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISTEXT_no",
            label: "ISTEXT(Data!A50)",
            formula: "=ISTEXT(Data!A50)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISBLANK_yes",
            label: "ISBLANK(Data!A53)",
            formula: "=ISBLANK(Data!A53)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISBLANK_no",
            label: "ISBLANK(Data!A50)",
            formula: "=ISBLANK(Data!A50)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISERROR_yes",
            label: "ISERROR(1/0)",
            formula: "=ISERROR(1/0)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISERROR_no",
            label: "ISERROR(42)",
            formula: "=ISERROR(42)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISLOGICAL_yes",
            label: "ISLOGICAL(Data!A52)",
            formula: "=ISLOGICAL(Data!A52)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISLOGICAL_no",
            label: "ISLOGICAL(Data!A50)",
            formula: "=ISLOGICAL(Data!A50)",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TYPE_number",
            label: "TYPE(Data!A50)",
            formula: "=TYPE(Data!A50)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TYPE_text",
            label: "TYPE(Data!A51)",
            formula: "=TYPE(Data!A51)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TYPE_logical",
            label: "TYPE(Data!A52)",
            formula: "=TYPE(Data!A52)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUBSTITUTE_basic",
            label: "SUBSTITUTE(\"Hello World\",\"World\",\"Earth\")",
            formula: "=SUBSTITUTE(\"Hello World\",\"World\",\"Earth\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUBSTITUTE_nth",
            label: "SUBSTITUTE(\"mississippi\",\"s\",\"S\",2)",
            formula: "=SUBSTITUTE(\"mississippi\",\"s\",\"S\",2)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "REPLACE_basic",
            label: "REPLACE(\"Hello World\",7,5,\"Earth\")",
            formula: "=REPLACE(\"Hello World\",7,5,\"Earth\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "REPT_basic",
            label: "REPT(\"ab\",3)",
            formula: "=REPT(\"ab\",3)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "EXACT_match",
            label: "EXACT(\"hello\",\"hello\")",
            formula: "=EXACT(\"hello\",\"hello\")",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "EXACT_no_match",
            label: "EXACT(\"hello\",\"Hello\")",
            formula: "=EXACT(\"hello\",\"Hello\")",
            expected_type: "boolean",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VALUE_numeric",
            label: "VALUE(\"123.45\")",
            formula: "=VALUE(\"123.45\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TEXT_format",
            label: "TEXT(0.75,\"0.0%\")",
            formula: "=TEXT(0.75,\"0.0%\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CLEAN_basic",
            label: "CLEAN(CHAR(9)&\"hello\"&CHAR(10))",
            formula: "=CLEAN(CHAR(9)&\"hello\"&CHAR(10))",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NUMBERVALUE_basic",
            label: "NUMBERVALUE(\"1,234.56\",\".\",\",\")",
            formula: "=NUMBERVALUE(\"1,234.56\",\".\",\",\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PMT_basic",
            label: "PMT(Data!B43/Data!D43,Data!C43*Data!D43,-Data!A43)",
            formula: "=PMT(Data!B43/Data!D43,Data!C43*Data!D43,-Data!A43)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "FV_basic",
            label: "FV(0.05/12,10*12,-200)",
            formula: "=FV(0.05/12,10*12,-200)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PV_basic",
            label: "PV(0.08/12,20*12,-500)",
            formula: "=PV(0.08/12,20*12,-500)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NPER_basic",
            label: "NPER(0.06/12,-200,10000)",
            formula: "=NPER(0.06/12,-200,10000)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NPV_basic",
            label: "NPV(0.1,Data!A45:A48)+Data!A44",
            formula: "=NPV(0.1,Data!A45:A48)+Data!A44",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IRR_basic",
            label: "IRR(Data!A44:A48)",
            formula: "=IRR(Data!A44:A48)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ERROR_div_zero",
            label: "1/0",
            formula: "=1/0",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ERROR_ref",
            label: "INDEX(Data!A2:L2,99)",
            formula: "=INDEX(Data!A2:L2,99)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ERROR_value",
            label: "VALUE(\"not_a_number\")",
            formula: "=VALUE(\"not_a_number\")",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ERROR_nested_iferror",
            label: "IFERROR(IFERROR(1/0,SQRT(-1)),\"caught\")",
            formula: "=IFERROR(IFERROR(1/0,SQRT(-1)),\"caught\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUM_with_error",
            label: "SUM(1,2,1/0)",
            formula: "=SUM(1,2,1/0)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUM_with_blanks",
            label: "SUM(Data!A40:E40)",
            formula: "=SUM(Data!A40:E40)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "COUNTA_with_blanks",
            label: "COUNTA(Data!A40:E40)",
            formula: "=COUNTA(Data!A40:E40)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "COUNTBLANK_range",
            label: "COUNTBLANK(Data!A40:E40)",
            formula: "=COUNTBLANK(Data!A40:E40)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CONCATENATE_blank",
            label: "CONCATENATE(Data!A40,Data!B40,Data!C40)",
            formula: "=CONCATENATE(Data!A40,Data!B40,Data!C40)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IF_blank",
            label: "IF(Data!B40=\"\",\"empty\",\"full\")",
            formula: "=IF(Data!B40=\"\",\"empty\",\"full\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PRECISION_sum",
            label: "Data!A56+Data!B56",
            formula: "=Data!A56+Data!B56",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PRECISION_round",
            label: "ROUND(Data!A56+Data!B56,1)",
            formula: "=ROUND(Data!A56+Data!B56,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PRECISION_large",
            label: "Data!A57+Data!B57",
            formula: "=Data!A57+Data!B57",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_MATCH_2D",
            label:
                "INDEX(Data!E17:G20,MATCH(\"Q3\",Data!D17:D20,0),MATCH(\"South\",Data!E16:G16,0))",
            formula:
                "=INDEX(Data!E17:G20,MATCH(\"Q3\",Data!D17:D20,0),MATCH(\"South\",Data!E16:G16,0))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUMPRODUCT_IF",
            label: "SUMPRODUCT((Data!B29:B37=\"fruit\")*Data!C29:C37)",
            formula: "=SUMPRODUCT((Data!B29:B37=\"fruit\")*Data!C29:C37)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUMPRODUCT_MULTI",
            label: "SUMPRODUCT((Data!B29:B37=\"fruit\")*(Data!C29:C37>15))",
            formula: "=SUMPRODUCT((Data!B29:B37=\"fruit\")*(Data!C29:C37>15))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NESTED_XLOOKUP",
            label:
                "XLOOKUP(XLOOKUP(\"cherry\",Data!A16:A25,Data!B16:B25),Data!B16:B25,Data!A16:A25)",
            formula:
                "=XLOOKUP(XLOOKUP(\"cherry\",Data!A16:A25,Data!B16:B25),Data!B16:B25,Data!A16:A25)",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CHOOSE_basic",
            label: "CHOOSE(2,\"apple\",\"banana\",\"cherry\")",
            formula: "=CHOOSE(2,\"apple\",\"banana\",\"cherry\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CHOOSE_calc",
            label: "CHOOSE(MATCH(70,Data!P17:P21,1),\"F\",\"D\",\"C\",\"B\",\"A\")",
            formula: "=CHOOSE(MATCH(70,Data!P17:P21,1),\"F\",\"D\",\"C\",\"B\",\"A\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SORTBY_basic",
            label: "INDEX(SORTBY(Data!A29:A37,Data!C29:C37,-1),1)",
            formula: "=INDEX(SORTBY(Data!A29:A37,Data!C29:C37,-1),1)",
            expected_type: "string",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "FILTER_multi",
            label: "SUM(FILTER(Data!C29:C37,(Data!B29:B37=\"fruit\")*(Data!C29:C37>15)))",
            formula: "=SUM(FILTER(Data!C29:C37,(Data!B29:B37=\"fruit\")*(Data!C29:C37>15)))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "SEQUENCE_2d_sum",
            label: "SUM(SEQUENCE(3,3,1,1))",
            formula: "=SUM(SEQUENCE(3,3,1,1))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "TMPL_ref_row70",
            label: "Data!B70",
            formula: "=Data!B70",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_ref_row71",
            label: "Data!B71",
            formula: "=Data!B71",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_ref_row75",
            label: "Data!B75",
            formula: "=Data!B75",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_ref_row80",
            label: "Data!B80",
            formula: "=Data!B80",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_ref_row85",
            label: "Data!B85",
            formula: "=Data!B85",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_ref_row89",
            label: "Data!B89",
            formula: "=Data!B89",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_sum_row70",
            label: "Data!A70+Data!B70",
            formula: "=Data!A70+Data!B70",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_sum_row75",
            label: "Data!A75+Data!B75",
            formula: "=Data!A75+Data!B75",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_sum_row80",
            label: "Data!A80+Data!B80",
            formula: "=Data!A80+Data!B80",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_sum_row89",
            label: "Data!A89+Data!B89",
            formula: "=Data!A89+Data!B89",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_vlookup_70",
            label: "VLOOKUP(1,Data!A70:B89,2,FALSE)",
            formula: "=VLOOKUP(1,Data!A70:B89,2,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_vlookup_89",
            label: "VLOOKUP(20,Data!A70:B89,2,FALSE)",
            formula: "=VLOOKUP(20,Data!A70:B89,2,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_if_row70",
            label: "IF(Data!C70=\"alpha\",\"yes\",\"no\")",
            formula: "=IF(Data!C70=\"alpha\",\"yes\",\"no\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_if_row71",
            label: "IF(Data!C71=\"alpha\",\"yes\",\"no\")",
            formula: "=IF(Data!C71=\"alpha\",\"yes\",\"no\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_if_row72",
            label: "IF(Data!C72=\"alpha\",\"yes\",\"no\")",
            formula: "=IF(Data!C72=\"alpha\",\"yes\",\"no\")",
            expected_type: "string",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_index_r5",
            label: "INDEX(Data!B70:B89,5)",
            formula: "=INDEX(Data!B70:B89,5)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_index_r10",
            label: "INDEX(Data!B70:B89,10)",
            formula: "=INDEX(Data!B70:B89,10)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_index_r20",
            label: "INDEX(Data!B70:B89,20)",
            formula: "=INDEX(Data!B70:B89,20)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_countif_alpha",
            label: "COUNTIF(Data!C70:C89,\"alpha\")",
            formula: "=COUNTIF(Data!C70:C89,\"alpha\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TMPL_sumif_alpha",
            label: "SUMIF(Data!C70:C89,\"alpha\",Data!B70:B89)",
            formula: "=SUMIF(Data!C70:C89,\"alpha\",Data!B70:B89)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PMT_small_loan",
            label: "PMT(Data!B92/Data!D92,Data!C92*Data!D92,-Data!A92)",
            formula: "=PMT(Data!B92/Data!D92,Data!C92*Data!D92,-Data!A92)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "RATE_basic",
            label: "RATE(60,-200,10000)*12",
            formula: "=RATE(60,-200,10000)*12",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NPV_alt",
            label: "NPV(0.08,Data!A94:A97)+Data!A93",
            formula: "=NPV(0.08,Data!A94:A97)+Data!A93",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IRR_alt",
            label: "IRR(Data!A93:A97)",
            formula: "=IRR(Data!A93:A97)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "FV_monthly",
            label: "FV(0.04/12,5*12,-300)",
            formula: "=FV(0.04/12,5*12,-300)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PV_annuity",
            label: "PV(0.05,10,-1000)",
            formula: "=PV(0.05,10,-1000)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NPER_double",
            label: "NPER(0.07/12,-500,30000)",
            formula: "=NPER(0.07/12,-500,30000)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "IPMT_first",
            label: "IPMT(0.06/12,1,360,-200000)",
            formula: "=IPMT(0.06/12,1,360,-200000)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PPMT_first",
            label: "PPMT(0.06/12,1,360,-200000)",
            formula: "=PPMT(0.06/12,1,360,-200000)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CUMIPMT_year1",
            label: "CUMIPMT(0.06/12,360,200000,1,12,0)",
            formula: "=CUMIPMT(0.06/12,360,200000,1,12,0)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "CUMPRINC_year1",
            label: "CUMPRINC(0.06/12,360,200000,1,12,0)",
            formula: "=CUMPRINC(0.06/12,360,200000,1,12,0)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SLN_basic",
            label: "SLN(10000,1000,10)",
            formula: "=SLN(10000,1000,10)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "DB_year1",
            label: "DB(10000,1000,10,1)",
            formula: "=DB(10000,1000,10,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "EFFECT_basic",
            label: "EFFECT(0.06,12)",
            formula: "=EFFECT(0.06,12)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "NOMINAL_basic",
            label: "NOMINAL(0.0617,12)",
            formula: "=NOMINAL(0.0617,12)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "STDEV_S_set",
            label: "STDEV.S(Data!A100:A109)",
            formula: "=STDEV.S(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "STDEV_P_set",
            label: "STDEV.P(Data!A100:A109)",
            formula: "=STDEV.P(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VAR_S_set",
            label: "VAR.S(Data!A100:A109)",
            formula: "=VAR.S(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VAR_P_set",
            label: "VAR.P(Data!A100:A109)",
            formula: "=VAR.P(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PERCENTILE_INC_25",
            label: "PERCENTILE.INC(Data!A100:A109,0.25)",
            formula: "=PERCENTILE.INC(Data!A100:A109,0.25)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PERCENTILE_INC_75",
            label: "PERCENTILE.INC(Data!A100:A109,0.75)",
            formula: "=PERCENTILE.INC(Data!A100:A109,0.75)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PERCENTILE_EXC_25",
            label: "PERCENTILE.EXC(Data!A100:A109,0.25)",
            formula: "=PERCENTILE.EXC(Data!A100:A109,0.25)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "QUARTILE_INC_3",
            label: "QUARTILE.INC(Data!A100:A109,3)",
            formula: "=QUARTILE.INC(Data!A100:A109,3)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LARGE_1st",
            label: "LARGE(Data!A100:A109,1)",
            formula: "=LARGE(Data!A100:A109,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SMALL_1st",
            label: "SMALL(Data!A100:A109,1)",
            formula: "=SMALL(Data!A100:A109,1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "RANK_EQ_22",
            label: "RANK.EQ(22,Data!A100:A109)",
            formula: "=RANK.EQ(22,Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "RANK_AVG_22",
            label: "RANK.AVG(22,Data!A100:A109)",
            formula: "=RANK.AVG(22,Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "AVERAGE_set",
            label: "AVERAGE(Data!A100:A109)",
            formula: "=AVERAGE(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "GEOMEAN_set",
            label: "GEOMEAN(Data!A100:A109)",
            formula: "=GEOMEAN(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "HARMEAN_set",
            label: "HARMEAN(Data!A100:A109)",
            formula: "=HARMEAN(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TRIMMEAN_set",
            label: "TRIMMEAN(Data!A100:A109,0.2)",
            formula: "=TRIMMEAN(Data!A100:A109,0.2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "DEVSQ_set",
            label: "DEVSQ(Data!A100:A109)",
            formula: "=DEVSQ(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "AVEDEV_set",
            label: "AVEDEV(Data!A100:A109)",
            formula: "=AVEDEV(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "KURT_set",
            label: "KURT(Data!A100:A109)",
            formula: "=KURT(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SKEW_set",
            label: "SKEW(Data!A100:A109)",
            formula: "=SKEW(Data!A100:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "FILTER_gt25",
            label: "SUM(FILTER(Data!B70:B89,Data!B70:B89>500))",
            formula: "=SUM(FILTER(Data!B70:B89,Data!B70:B89>500))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "FILTER_alpha_sum",
            label: "SUM(FILTER(Data!B70:B89,Data!C70:C89=\"alpha\"))",
            formula: "=SUM(FILTER(Data!B70:B89,Data!C70:C89=\"alpha\"))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "FILTER_no_match",
            label: "FILTER(Data!B70:B89,Data!B70:B89>99999,\"none\")",
            formula: "=FILTER(Data!B70:B89,Data!B70:B89>99999,\"none\")",
            expected_type: "string",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "SORT_desc_first",
            label: "INDEX(SORT(Data!B70:B89,1,-1),1)",
            formula: "=INDEX(SORT(Data!B70:B89,1,-1),1)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "SORT_asc_first",
            label: "INDEX(SORT(Data!B70:B89),1)",
            formula: "=INDEX(SORT(Data!B70:B89),1)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "SORTBY_name_first",
            label: "INDEX(SORTBY(Data!A70:A89,Data!B70:B89,-1),1)",
            formula: "=INDEX(SORTBY(Data!A70:A89,Data!B70:B89,-1),1)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "UNIQUE_categories",
            label: "COUNTA(UNIQUE(Data!C70:C89))",
            formula: "=COUNTA(UNIQUE(Data!C70:C89))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "SEQUENCE_sum_5x5",
            label: "SUM(SEQUENCE(5,5,1,1))",
            formula: "=SUM(SEQUENCE(5,5,1,1))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "SEQUENCE_start_10",
            label: "INDEX(SEQUENCE(5,1,10,10),3)",
            formula: "=INDEX(SEQUENCE(5,1,10,10),3)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "TRANSPOSE_elem",
            label: "INDEX(TRANSPOSE(Data!A115:C117),1,3)",
            formula: "=INDEX(TRANSPOSE(Data!A115:C117),1,3)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "FILTER_AND_SORT",
            label: "INDEX(SORT(FILTER(Data!B70:B89,Data!C70:C89=\"alpha\"),1,-1),1)",
            formula: "=INDEX(SORT(FILTER(Data!B70:B89,Data!C70:C89=\"alpha\"),1,-1),1)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "RANDARRAY_type",
            label: "SUM(RANDARRAY(3,3))",
            formula: "=SUM(RANDARRAY(3,3))",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "MMULT_row1",
            label: "INDEX(MMULT(Data!A115:C117,TRANSPOSE(Data!A115:C115)),1)",
            formula: "=INDEX(MMULT(Data!A115:C117,TRANSPOSE(Data!A115:C115)),1)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "MMULT_row3",
            label: "INDEX(MMULT(Data!A115:C117,TRANSPOSE(Data!A115:C115)),3)",
            formula: "=INDEX(MMULT(Data!A115:C117,TRANSPOSE(Data!A115:C115)),3)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "CHOOSECOLS_elem",
            label: "INDEX(CHOOSECOLS(Data!A70:C89,2,3),1,1)",
            formula: "=INDEX(CHOOSECOLS(Data!A70:C89,2,3),1,1)",
            expected_type: "number",
            kind: FormulaKind::Formula2,
        },
        FormulaCase {
            id: "XLOOKUP_wildcard",
            label: "XLOOKUP(\"ch*\",Data!A16:A25,Data!B16:B25,,2)",
            formula: "=XLOOKUP(\"ch*\",Data!A16:A25,Data!B16:B25,,2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "XLOOKUP_not_found_def",
            label: "XLOOKUP(\"mango\",Data!A16:A25,Data!B16:B25,-1)",
            formula: "=XLOOKUP(\"mango\",Data!A16:A25,Data!B16:B25,-1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_wildcard",
            label: "MATCH(\"*berry\",Data!A16:A25,0)",
            formula: "=MATCH(\"*berry\",Data!A16:A25,0)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "INDEX_MATCH_price",
            label: "INDEX(Data!B16:B25,MATCH(\"grape\",Data!A16:A25,0))",
            formula: "=INDEX(Data!B16:B25,MATCH(\"grape\",Data!A16:A25,0))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_col3",
            label: "VLOOKUP(\"Q2\",Data!D17:G20,3,FALSE)",
            formula: "=VLOOKUP(\"Q2\",Data!D17:G20,3,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "HLOOKUP_row2",
            label: "HLOOKUP(8,Data!A1:L2,2,FALSE)",
            formula: "=HLOOKUP(8,Data!A1:L2,2,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LOOKUP_vector",
            label: "LOOKUP(75,Data!A2:L2,Data!A1:L1)",
            formula: "=LOOKUP(75,Data!A2:L2,Data!A1:L1)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "XMATCH_exact",
            label: "XMATCH(\"fig\",Data!A16:A25)",
            formula: "=XMATCH(\"fig\",Data!A16:A25)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "XMATCH_wildcard",
            label: "XMATCH(\"*dew\",Data!A16:A25,2)",
            formula: "=XMATCH(\"*dew\",Data!A16:A25,2)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "DATEDIF_months",
            label: "DATEDIF(Data!A112,Data!A113,\"M\")",
            formula: "=DATEDIF(Data!A112,Data!A113,\"M\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "DAYS_between",
            label: "DAYS(Data!A113,Data!A112)",
            formula: "=DAYS(Data!A113,Data!A112)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "WEEKDAY_sun",
            label: "WEEKDAY(Data!A112)",
            formula: "=WEEKDAY(Data!A112)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "WEEKNUM_basic",
            label: "WEEKNUM(Data!A112)",
            formula: "=WEEKNUM(Data!A112)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "ISOWEEKNUM_basic",
            label: "ISOWEEKNUM(Data!A112)",
            formula: "=ISOWEEKNUM(Data!A112)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "WORKDAY_plus10",
            label: "WORKDAY(Data!A112,10)",
            formula: "=WORKDAY(Data!A112,10)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "DATEVALUE_basic",
            label: "DATEVALUE(\"2024-06-15\")",
            formula: "=DATEVALUE(\"2024-06-15\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "TIMEVALUE_basic",
            label: "TIMEVALUE(\"14:30:00\")",
            formula: "=TIMEVALUE(\"14:30:00\")",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "HOUR_basic",
            label: "HOUR(TIMEVALUE(\"14:30:45\"))",
            formula: "=HOUR(TIMEVALUE(\"14:30:45\"))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MINUTE_basic",
            label: "MINUTE(TIMEVALUE(\"14:30:45\"))",
            formula: "=MINUTE(TIMEVALUE(\"14:30:45\"))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SECOND_basic",
            label: "SECOND(TIMEVALUE(\"14:30:45\"))",
            formula: "=SECOND(TIMEVALUE(\"14:30:45\"))",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "GCD_basic",
            label: "GCD(12,18,24)",
            formula: "=GCD(12,18,24)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "LCM_basic",
            label: "LCM(4,6,10)",
            formula: "=LCM(4,6,10)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "FACT_basic",
            label: "FACT(7)",
            formula: "=FACT(7)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "COMBIN_basic",
            label: "COMBIN(10,3)",
            formula: "=COMBIN(10,3)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PERMUT_basic",
            label: "PERMUT(10,3)",
            formula: "=PERMUT(10,3)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MROUND_basic",
            label: "MROUND(7.3,0.5)",
            formula: "=MROUND(7.3,0.5)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "QUOTIENT_basic",
            label: "QUOTIENT(17,5)",
            formula: "=QUOTIENT(17,5)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "PRODUCT_range",
            label: "PRODUCT(Data!A100:A104)",
            formula: "=PRODUCT(Data!A100:A104)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUMSQ_basic",
            label: "SUMSQ(3,4,5)",
            formula: "=SUMSQ(3,4,5)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "SUMX2MY2_basic",
            label: "SUMX2MY2(Data!A100:A104,Data!A105:A109)",
            formula: "=SUMX2MY2(Data!A100:A104,Data!A105:A109)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        // Wildcard support in VLOOKUP / HLOOKUP / MATCH
        FormulaCase {
            id: "VLOOKUP_wildcard_star",
            label: "VLOOKUP(\"*berry\",Data!A16:B25,2,FALSE)",
            formula: "=VLOOKUP(\"*berry\",Data!A16:B25,2,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_wildcard_question",
            label: "VLOOKUP(\"d?te\",Data!A16:B25,2,FALSE)",
            formula: "=VLOOKUP(\"d?te\",Data!A16:B25,2,FALSE)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "VLOOKUP_wildcard_no_match",
            label: "VLOOKUP(\"z*\",Data!A16:B25,2,FALSE)",
            formula: "=VLOOKUP(\"z*\",Data!A16:B25,2,FALSE)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "HLOOKUP_wildcard_star",
            label: "HLOOKUP(\"*berry\",Data!A16:B25,2,FALSE)",
            formula: "=HLOOKUP(\"*berry\",Data!A16:B25,2,FALSE)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_wildcard_question",
            label: "MATCH(\"ch???y\",Data!A16:A25,0)",
            formula: "=MATCH(\"ch???y\",Data!A16:A25,0)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_wildcard_combined",
            label: "MATCH(\"g*p?\",Data!A16:A25,0)",
            formula: "=MATCH(\"g*p?\",Data!A16:A25,0)",
            expected_type: "number",
            kind: FormulaKind::Formula,
        },
        FormulaCase {
            id: "MATCH_wildcard_no_match",
            label: "MATCH(\"z*\",Data!A16:A25,0)",
            formula: "=MATCH(\"z*\",Data!A16:A25,0)",
            expected_type: "error",
            kind: FormulaKind::Formula,
        },
    ];

    let mut row = 2u32;
    for case in cases {
        wb.set_cell_value(&cell_addr(row, 1), case.id)?;
        wb.set_cell_value(&cell_addr(row, 2), case.label)?;
        wb.set_cell_value(&cell_addr(row, 4), case.expected_type)?;
        match case.kind {
            FormulaKind::Formula => wb.set_cell_formula(&cell_addr(row, 3), case.formula)?,
            FormulaKind::Formula2 => wb.set_cell_formula2(&cell_addr(row, 3), case.formula)?,
        }
        row += 1;
    }

    Ok(())
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
