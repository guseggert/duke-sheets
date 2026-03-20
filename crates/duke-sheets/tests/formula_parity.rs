use std::path::PathBuf;

use duke_sheets::prelude::*;

#[derive(Debug, Clone)]
struct ParityCase {
    row: u32,
    id: String,
    label: String,
    expected_type: String,
    excel_value: CellValue,
}

#[test]
#[ignore = "requires data/formula-parity.xlsx — run `mise run test:excel` to generate"]
fn formula_parity_matches_excel_cached_values() {
    let fixture_path = formula_parity_path();
    if !fixture_path.exists() {
        println!(
            "skipping formula parity test: {} does not exist",
            fixture_path.display()
        );
        return;
    }

    let mut workbook = XlsxReader::read_file(&fixture_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture_path.display()));

    let cases = collect_parity_cases(&workbook);
    workbook.calculate().expect("calculate workbook");

    let tests_sheet = workbook
        .worksheet_by_name("Tests")
        .expect("Tests worksheet should exist");

    let mut failures: Vec<String> = Vec::new();
    for case in &cases {
        let actual_value = tests_sheet.get_value_at(case.row, 2);
        match compare_case(case, &actual_value) {
            Ok(()) => println!("PASS {} [{}]", case.id, case.label),
            Err(message) => {
                println!("FAIL {message}");
                failures.push(message);
            }
        }
    }

    assert!(
        failures.is_empty(),
        "{} formula parity case(s) failed:\n{}",
        failures.len(),
        failures.join("\n")
    );
}

fn formula_parity_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join("data/formula-parity.xlsx")
}

fn collect_parity_cases(workbook: &Workbook) -> Vec<ParityCase> {
    let tests_sheet = workbook
        .worksheet_by_name("Tests")
        .expect("Tests worksheet should exist");
    let used_range = tests_sheet
        .used_range()
        .expect("Tests sheet should not be empty");

    let mut cases = Vec::new();
    for row in 1..=used_range.end.row {
        let id = tests_sheet
            .get_value_at(row, 0)
            .as_string()
            .unwrap_or("")
            .trim()
            .to_string();
        if id.is_empty() {
            continue;
        }

        let label = tests_sheet
            .get_value_at(row, 1)
            .as_string()
            .unwrap_or("")
            .to_string();
        let expected_type = tests_sheet
            .get_value_at(row, 3)
            .as_string()
            .unwrap_or("")
            .trim()
            .to_string();

        cases.push(ParityCase {
            row,
            id,
            label,
            expected_type,
            excel_value: tests_sheet.get_value_at(row, 2),
        });
    }

    cases
}

fn compare_case(case: &ParityCase, actual_value: &CellValue) -> std::result::Result<(), String> {
    let context = format!("{} [{}] row {}", case.id, case.label, case.row + 1);

    match case.expected_type.as_str() {
        "number" => {
            let expected = strict_number(&case.excel_value).ok_or_else(|| {
                format!(
                    "{context}: expected cached Excel number, got {:?}",
                    case.excel_value
                )
            })?;
            let actual = strict_number(actual_value).ok_or_else(|| {
                format!("{context}: expected calculated number, got {actual_value:?}")
            })?;

            if is_type_only_case(case) {
                return Ok(());
            }

            if (expected - actual).abs() <= 1e-9 {
                Ok(())
            } else {
                Err(format!("{context}: expected {expected}, got {actual}"))
            }
        }
        "string" => {
            let expected = strict_string(&case.excel_value).ok_or_else(|| {
                format!(
                    "{context}: expected cached Excel string, got {:?}",
                    case.excel_value
                )
            })?;
            let actual = strict_string(actual_value).ok_or_else(|| {
                format!("{context}: expected calculated string, got {actual_value:?}")
            })?;
            if expected == actual {
                Ok(())
            } else {
                Err(format!("{context}: expected {expected:?}, got {actual:?}"))
            }
        }
        "boolean" => {
            let expected = strict_bool(&case.excel_value).ok_or_else(|| {
                format!(
                    "{context}: expected cached Excel boolean, got {:?}",
                    case.excel_value
                )
            })?;
            let actual = strict_bool(actual_value).ok_or_else(|| {
                format!("{context}: expected calculated boolean, got {actual_value:?}")
            })?;
            if expected == actual {
                Ok(())
            } else {
                Err(format!("{context}: expected {expected}, got {actual}"))
            }
        }
        "error" => match (&case.excel_value, actual_value) {
            (CellValue::Error(expected), CellValue::Error(actual)) if expected == actual => Ok(()),
            (CellValue::Error(expected), CellValue::Error(actual)) => Err(format!(
                "{context}: expected error {:?}, got {:?}",
                expected, actual
            )),
            (CellValue::Error(expected), other) => Err(format!(
                "{context}: expected error {:?}, got {other:?}",
                expected
            )),
            (other, _) => Err(format!(
                "{context}: expected cached Excel error, got {other:?}"
            )),
        },
        other => Err(format!("{context}: unsupported expected_type {other:?}")),
    }
}

fn is_type_only_case(case: &ParityCase) -> bool {
    case.id.contains("TODAY")
        || case.label.contains("TODAY()")
        || case.id.contains("RANDARRAY")
        || case.label.contains("RANDARRAY(")
        || case.id.contains("RAND")
}

fn strict_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(number) => Some(*number),
        _ => None,
    }
}

fn strict_string(value: &CellValue) -> Option<&str> {
    match value {
        CellValue::String(text) => Some(text.as_ref()),
        _ => None,
    }
}

fn strict_bool(value: &CellValue) -> Option<bool> {
    match value {
        CellValue::Boolean(boolean) => Some(*boolean),
        _ => None,
    }
}
