//! Shared helpers for benchmark data generation.
#![allow(dead_code)]

use std::io::Cursor;
use std::time::Duration;

use criterion::Criterion;
use duke_sheets::{CellValue, Workbook, XlsxWriter};

/// Benchmark size presets: (label, rows, cols).
pub const SIZES: &[(&str, u32, u16)] = &[
    ("100_cells", 10, 10),
    ("1k_cells", 100, 10),
    ("10k_cells", 100, 100),
];

/// Faster criterion config: 1s warmup, 3s measurement, 50 samples.
pub fn fast_criterion() -> Criterion {
    Criterion::default()
        .warm_up_time(Duration::from_secs(1))
        .measurement_time(Duration::from_secs(3))
        .sample_size(50)
}

/// Generate a workbook with mixed cell data (strings, numbers, booleans).
pub fn generate_workbook(rows: u32, cols: u16) -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let _ = sheet.set_name("Data");

    for row in 0..rows {
        for col in 0..cols {
            let idx = row * cols as u32 + col as u32;
            match idx % 4 {
                0 => {
                    let _ = sheet.set_cell_value_at(row, col, format!("str_{}_{}", row, col));
                }
                1 => {
                    let _ = sheet.set_cell_value_at(row, col, idx as f64 * 1.5);
                }
                2 => {
                    let _ = sheet.set_cell_value_at(row, col, idx as f64);
                }
                _ => {
                    let _ = sheet.set_cell_value_at(row, col, CellValue::Boolean(row % 2 == 0));
                }
            }
        }
    }

    wb
}

/// Serialize a workbook to XLSX bytes.
pub fn workbook_to_xlsx_bytes(wb: &Workbook) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    XlsxWriter::write(wb, &mut buf).expect("write XLSX");
    buf.into_inner()
}

/// Generate CSV content as a string.
pub fn generate_csv_string(rows: u32, cols: u16) -> String {
    let mut csv = String::new();
    for col in 0..cols {
        if col > 0 {
            csv.push(',');
        }
        csv.push_str(&format!("col_{}", col));
    }
    csv.push('\n');

    for row in 0..rows {
        for col in 0..cols {
            if col > 0 {
                csv.push(',');
            }
            let idx = row * cols as u32 + col as u32;
            match idx % 3 {
                0 => csv.push_str(&format!("str_{}_{}", row, col)),
                1 => csv.push_str(&format!("{:.2}", idx as f64 * 1.5)),
                _ => csv.push_str(&format!("{}", idx)),
            }
        }
        csv.push('\n');
    }
    csv
}

// ── Formula benchmark data ──────────────────────────────────────────

pub fn simple_formulas() -> Vec<&'static str> {
    vec!["=1+2", "=A1", "=A1+B1", "=A1*2", "=-A1"]
}

pub fn medium_formulas() -> Vec<&'static str> {
    vec![
        "=SUM(A1:A100)",
        "=IF(A1>0,B1,C1)",
        "=VLOOKUP(A1,B1:D100,3,FALSE)",
        "=AVERAGE(A1:Z1)",
        "=INDEX(A1:Z100,MATCH(A1,A1:A100,0),1)",
    ]
}

pub fn complex_formulas() -> Vec<&'static str> {
    vec![
        "=SUMPRODUCT((A1:A100>0)*(B1:B100<100))",
        "=IF(AND(A1>0,B1<100),SUM(C1:C100),AVERAGE(D1:D100))",
        "=IFERROR(VLOOKUP(A1,Sheet2!A1:D1000,4,FALSE),\"N/A\")",
        "=SUMIFS(C1:C1000,A1:A1000,\">0\",B1:B1000,\"<100\")",
        "=INDEX(A1:Z100,MATCH(1,(A1:A100=E1)*(B1:B100=F1),0),3)",
    ]
}

// ── Calculation benchmark workbooks ─────────────────────────────────

/// Linear chain: A1=1, A2=A1+1, A3=A2+1, …
pub fn generate_linear_chain(depth: u32) -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let _ = sheet.set_cell_value_at(0, 0, 1.0);
    for row in 1..depth {
        let _ = sheet.set_cell_formula_at(row, 0, &format!("=A{}+1", row));
    }
    wb
}

/// Fan-out: row 1 has values, row 2 each cell sums A1 through its column.
pub fn generate_fanout(width: u16) -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    for col in 0..width {
        let _ = sheet.set_cell_value_at(0, col, col as f64 + 1.0);
        let end = col_to_letter(col);
        let _ = sheet.set_cell_formula_at(1, col, &format!("=SUM(A1:{}1)", end));
    }
    wb
}

/// Cross-sheet: Sheet1 has values, Sheet2 references Sheet1.
pub fn generate_cross_sheet(rows: u32) -> Workbook {
    let mut wb = Workbook::new();
    {
        let s1 = wb.worksheet_mut(0).unwrap();
        let _ = s1.set_name("Source");
        for row in 0..rows {
            let _ = s1.set_cell_value_at(row, 0, row as f64 + 1.0);
        }
    }
    let _ = wb.add_worksheet_with_name("Calc");
    {
        let s2 = wb.worksheet_mut(1).unwrap();
        for row in 0..rows {
            let _ = s2.set_cell_formula_at(row, 0, &format!("=Source!A{}*2", row + 1));
        }
        let _ = s2.set_cell_formula_at(rows, 0, &format!("=SUM(A1:A{})", rows));
    }
    wb
}

fn col_to_letter(col: u16) -> String {
    if col < 26 {
        String::from((b'A' + col as u8) as char)
    } else {
        format!(
            "{}{}",
            (b'A' + (col / 26 - 1) as u8) as char,
            (b'A' + (col % 26) as u8) as char
        )
    }
}
