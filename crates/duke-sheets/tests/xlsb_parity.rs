use std::path::PathBuf;

use duke_sheets::prelude::*;
use duke_sheets_chart::ChartType;

#[derive(Debug, Clone)]
struct ParityCase {
    row: u32,
    id: String,
    label: String,
    expected_type: String,
    excel_value: CellValue,
}

#[test]
#[ignore = "requires data/formula-parity.xlsb — run `mise run generate:xlsb-parity`"]
fn xlsb_parity_matches_excel_cached_values() {
    let fixture_path = formula_parity_path();
    if !fixture_path.exists() {
        println!("skipping: {} does not exist", fixture_path.display());
        return;
    }

    let mut workbook = Workbook::open(&fixture_path)
        .unwrap_or_else(|e| panic!("open {}: {e}", fixture_path.display()));

    let cases = collect_parity_cases(&workbook);
    println!("XLSB parity: {} cases collected", cases.len());

    workbook.calculate().expect("calculate workbook");

    let tests_sheet = workbook
        .worksheet_by_name("Tests")
        .expect("Tests worksheet should exist");

    let mut failures: Vec<String> = Vec::new();
    for case in &cases {
        let calculated = tests_sheet.get_value_at(case.row, 2);
        match compare_case(case, &calculated) {
            Ok(()) => {}
            Err(message) => {
                println!("FAIL {message}");
                failures.push(message);
            }
        }
    }

    println!(
        "XLSB parity: {} cases checked, {} failures",
        cases.len(),
        failures.len()
    );

    assert!(
        failures.is_empty(),
        "{} XLSB formula parity case(s) failed:\n{}",
        failures.len(),
        failures.join("\n")
    );
}

#[test]
#[ignore = "requires data/formula-parity.xlsb — run `mise run generate:xlsb-parity`"]
fn xlsb_roundtrip_preserves_values() {
    let fixture_path = formula_parity_path();
    if !fixture_path.exists() {
        println!("skipping: {} does not exist", fixture_path.display());
        return;
    }

    let workbook = Workbook::open(&fixture_path).unwrap_or_else(|e| panic!("open: {e}"));

    let mut buf = Vec::new();
    duke_sheets::XlsbWriter::write(&workbook, std::io::Cursor::new(&mut buf))
        .expect("XlsbWriter::write");

    let wb2 = Workbook::from_bytes(&buf).expect("from_bytes roundtrip");

    assert_eq!(
        workbook.sheet_count(),
        wb2.sheet_count(),
        "sheet count mismatch"
    );

    let tests1 = workbook
        .worksheet_by_name("Tests")
        .expect("Tests sheet original");
    let tests2 = wb2
        .worksheet_by_name("Tests")
        .expect("Tests sheet roundtrip");

    let used = tests1.used_range().expect("used range");
    let mut mismatches = 0;
    for row in 0..=used.end.row {
        for col in 0..=used.end.col {
            let v1 = tests1.get_value_at(row, col);
            let v2 = tests2.get_value_at(row, col);
            if !values_equal(&v1, &v2) {
                if mismatches < 20 {
                    println!("MISMATCH ({row},{col}): {v1:?} vs {v2:?}");
                }
                mismatches += 1;
            }
        }
    }
    assert_eq!(
        mismatches, 0,
        "{mismatches} cell value mismatches in roundtrip"
    );
}

#[test]
#[ignore = "requires data/chart-parity.xlsb — run generate:xlsb-parity"]
fn xlsb_chart_parity_reads_successfully() {
    let path = chart_parity_path();
    if !path.exists() {
        println!("skipping: {} does not exist", path.display());
        return;
    }

    let wb = Workbook::open(&path).unwrap_or_else(|e| panic!("open chart-parity.xlsb: {e}"));

    println!("Chart XLSB parity: {} sheet(s)", wb.sheet_count());

    let ws = wb.worksheet(0).expect("first sheet");
    println!("  Sheet 0: \"{}\"", ws.name());

    let header_a = ws.get_value_at(0, 0);
    let header_b = ws.get_value_at(0, 1);
    let header_c = ws.get_value_at(0, 2);
    assert_eq!(header_a.as_string().unwrap(), "Month", "A1 header");
    assert_eq!(header_b.as_string().unwrap(), "Revenue", "B1 header");
    assert_eq!(header_c.as_string().unwrap(), "Profit", "C1 header");

    assert!(
        matches!(ws.get_value_at(1, 0), CellValue::Empty),
        "row 1 should be empty (spacing row in chart fixture)"
    );
    let val = ws.get_value_at(2, 1);
    assert!(
        matches!(val, CellValue::Number(n) if (n - 10.0).abs() < 0.001),
        "B3 should be 10.0, got {:?}",
        val
    );

    let charts = ws.charts();
    assert_eq!(charts.len(), 8, "expected 8 charts from XLSB fixture");

    let find_chart = |title: &str| -> &Chart {
        charts
            .iter()
            .find(|c| c.title.as_deref() == Some(title))
            .unwrap_or_else(|| panic!("chart titled '{title}' not found"))
    };

    let col_chart = find_chart("ColumnClustered");
    assert_eq!(col_chart.chart_type, ChartType::ColumnClustered);
    assert_eq!(col_chart.series.len(), 2, "ColumnClustered series count");

    let bar_chart = find_chart("BarClustered");
    assert_eq!(bar_chart.chart_type, ChartType::BarClustered);
    assert_eq!(bar_chart.series.len(), 2, "BarClustered series count");

    let line_chart = find_chart("Line");
    assert_eq!(line_chart.chart_type, ChartType::Line);
    assert_eq!(line_chart.series.len(), 2, "Line series count");

    let pie_chart = find_chart("Pie");
    assert_eq!(pie_chart.chart_type, ChartType::Pie);
    assert_eq!(pie_chart.series.len(), 1, "Pie series count");

    let doughnut = find_chart("Doughnut");
    assert_eq!(doughnut.chart_type, ChartType::Doughnut);

    let scatter = find_chart("XYScatter");
    assert_eq!(scatter.chart_type, ChartType::ScatterLines);

    let area = find_chart("Area");
    assert_eq!(area.chart_type, ChartType::Area);

    let radar = find_chart("Radar");
    assert_eq!(radar.chart_type, ChartType::Radar);

    println!("  All 8 charts verified: types, titles, series counts");
}

#[test]
#[ignore = "requires data/chart-parity.xlsb — run generate:xlsb-parity"]
fn xlsb_chart_roundtrip_preserves_data() {
    let path = chart_parity_path();
    if !path.exists() {
        println!("skipping: {} does not exist", path.display());
        return;
    }

    let wb = Workbook::open(&path).unwrap_or_else(|e| panic!("open: {e}"));

    let mut buf = Vec::new();
    duke_sheets::XlsbWriter::write(&wb, std::io::Cursor::new(&mut buf)).expect("XlsbWriter::write");

    let wb2 = Workbook::from_bytes(&buf).expect("from_bytes roundtrip");

    assert_eq!(wb.sheet_count(), wb2.sheet_count(), "sheet count mismatch");

    let ws1 = wb.worksheet(0).expect("original sheet");
    let ws2 = wb2.worksheet(0).expect("roundtrip sheet");

    let used = ws1.used_range().expect("used range");
    let mut mismatches = 0;
    for row in 0..=used.end.row {
        for col in 0..=used.end.col {
            let v1 = ws1.get_value_at(row, col);
            let v2 = ws2.get_value_at(row, col);
            if !values_equal(&v1, &v2) {
                if mismatches < 10 {
                    println!("MISMATCH ({row},{col}): {v1:?} vs {v2:?}");
                }
                mismatches += 1;
            }
        }
    }
    assert_eq!(
        mismatches, 0,
        "{mismatches} cell value mismatches in chart round-trip"
    );

    let d1 = ws1.raw_drawing_objects.len();
    let d2 = ws2.raw_drawing_objects.len();
    assert_eq!(d1, d2, "drawing object count changed in round-trip");
}

fn formula_parity_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join("data/formula-parity.xlsb")
}

fn chart_parity_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join("data/chart-parity.xlsb")
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
            if is_volatile_case(case) {
                return Ok(());
            }
            let eps = number_epsilon(case);
            if (expected - actual).abs() <= eps {
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
            if expected == actual || complex_strings_close(expected, actual) {
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

fn is_volatile_case(case: &ParityCase) -> bool {
    let id = case.id.to_ascii_uppercase();
    let label = case.label.to_ascii_uppercase();
    ["TODAY", "NOW", "RANDARRAY", "RANDBETWEEN", "RAND"]
        .into_iter()
        .any(|name| id.contains(name) || label.contains(&format!("{name}(")))
}

fn number_epsilon(case: &ParityCase) -> f64 {
    match case.id.as_str() {
        id if id.contains("XIRR") => 2e-9,
        id if id.contains("FORECAST_ETS") => 0.005,
        _ => 1e-9,
    }
}

fn strict_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(n) => Some(*n),
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
        CellValue::Boolean(b) => Some(*b),
        _ => None,
    }
}

fn complex_strings_close(a: &str, b: &str) -> bool {
    if !a.contains('i') && !b.contains('i') {
        return false;
    }
    fn parse_complex(s: &str) -> Option<(f64, f64)> {
        let s = s.trim();
        if s == "0" {
            return Some((0.0, 0.0));
        }
        if s.ends_with('i') && !s.contains('+') && !s[1..].contains('-') {
            let im = s[..s.len() - 1].parse::<f64>().ok()?;
            return Some((0.0, im));
        }
        let i_pos = s.find('i')?;
        let before_i = &s[..i_pos];
        let split = before_i.rfind(|c: char| (c == '+' || c == '-'))?;
        if split == 0 {
            let im = before_i.parse::<f64>().ok()?;
            return Some((0.0, im));
        }
        let real = s[..split].parse::<f64>().ok()?;
        let im = s[split..i_pos].parse::<f64>().ok()?;
        Some((real, im))
    }
    match (parse_complex(a), parse_complex(b)) {
        (Some((ar, ai)), Some((br, bi))) => (ar - br).abs() < 1e-6 && (ai - bi).abs() < 1e-6,
        _ => false,
    }
}

fn values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Empty, CellValue::Empty) => true,
        (CellValue::Number(a), CellValue::Number(b)) => (a - b).abs() < 1e-12,
        (CellValue::String(a), CellValue::String(b)) => a == b,
        (CellValue::Boolean(a), CellValue::Boolean(b)) => a == b,
        (CellValue::Error(a), CellValue::Error(b)) => a == b,
        _ => format!("{a:?}") == format!("{b:?}"),
    }
}
