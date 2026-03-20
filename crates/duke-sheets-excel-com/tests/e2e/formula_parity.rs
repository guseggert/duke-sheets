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
