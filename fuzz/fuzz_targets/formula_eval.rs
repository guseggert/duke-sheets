//! Fuzz target for the formula evaluation engine.
//!
//! Generates a workbook with seed data and random formulas, then runs
//! calculate().  Exercises: function dispatch, type coercion, array
//! expansion, cross-cell references, circular reference detection,
//! iterative convergence, and error propagation.

#![no_main]
use arbitrary::Unstructured;
use libfuzzer_sys::fuzz_target;

use duke_sheets::{CalculationOptions, WorkbookCalculationExt};
use duke_sheets_core::{CellValue, Workbook};

const MAX_DEPTH: usize = 4;

// Formula generation (reused structure from formula_parser target)

const FUNC_NAMES: &[&str] = &[
    "SUM",
    "AVERAGE",
    "COUNT",
    "IF",
    "VLOOKUP",
    "INDEX",
    "MATCH",
    "CONCATENATE",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "IFERROR",
    "SUMIF",
    "COUNTIF",
    "MIN",
    "MAX",
    "ROUND",
    "ABS",
    "NOW",
    "TODAY",
    "TEXT",
    "TRIM",
    "SUBSTITUTE",
    "AND",
    "OR",
    "NOT",
    "ISBLANK",
    "ISERROR",
    "MOD",
];

const BIN_OPS: &[&str] = &[
    "+", "-", "*", "/", "^", "=", "<>", "<", "<=", ">", ">=", "&",
];

fn col_to_letters(col: u16) -> String {
    if col < 26 {
        return String::from((b'A' + col as u8) as char);
    }
    let mut s = String::new();
    let mut c = col as u32;
    loop {
        s.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    s
}

/// Cell reference within the seed data grid (A1:E10)
fn arbitrary_cell_ref(u: &mut Unstructured) -> arbitrary::Result<String> {
    let col = u.int_in_range(0u16..=4)?;
    let row = u.int_in_range(1u32..=10)?;
    Ok(format!("{}{}", col_to_letters(col), row))
}

/// Range reference within the seed data grid
fn arbitrary_range_ref(u: &mut Unstructured) -> arbitrary::Result<String> {
    let col1 = u.int_in_range(0u16..=4)?;
    let row1 = u.int_in_range(1u32..=10)?;
    let col2 = u.int_in_range(col1..=4)?;
    let row2 = u.int_in_range(row1..=10)?;
    Ok(format!(
        "{}{}:{}{}",
        col_to_letters(col1),
        row1,
        col_to_letters(col2),
        row2,
    ))
}

enum Expr {
    Number(f64),
    Str(String),
    Bool(bool),
    CellRef(String),
    RangeRef(String),
    Function {
        name: &'static str,
        args: Vec<Expr>,
    },
    BinOp {
        op: &'static str,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    Negate(Box<Expr>),
    Parens(Box<Expr>),
}

fn arbitrary_expr(u: &mut Unstructured, depth: usize) -> arbitrary::Result<Expr> {
    let max_variant = if depth == 0 { 5 } else { 9 };
    let kind: u8 = u.int_in_range(0..=max_variant)?;

    match kind {
        0 => {
            let f: f64 = u.arbitrary()?;
            Ok(Expr::Number(if f.is_finite() { f } else { 0.0 }))
        }
        1 => {
            let len = u.int_in_range(0..=8)?;
            let mut s = String::with_capacity(len);
            for _ in 0..len {
                let c = u.int_in_range(b'a'..=b'z')? as char;
                s.push(c);
            }
            Ok(Expr::Str(s))
        }
        2 => Ok(Expr::Bool(u.arbitrary()?)),
        3 => Ok(Expr::CellRef(arbitrary_cell_ref(u)?)),
        4 | 5 => Ok(Expr::RangeRef(arbitrary_range_ref(u)?)),
        // Recursive
        6 => {
            let idx = u.int_in_range(0..=FUNC_NAMES.len() as u8 - 1)? as usize;
            let name = FUNC_NAMES[idx];
            let nargs = u.int_in_range(0..=3)?;
            let mut args = Vec::with_capacity(nargs);
            for _ in 0..nargs {
                args.push(arbitrary_expr(u, depth - 1)?);
            }
            Ok(Expr::Function { name, args })
        }
        7 => {
            let idx = u.int_in_range(0..=BIN_OPS.len() as u8 - 1)? as usize;
            Ok(Expr::BinOp {
                op: BIN_OPS[idx],
                left: Box::new(arbitrary_expr(u, depth - 1)?),
                right: Box::new(arbitrary_expr(u, depth - 1)?),
            })
        }
        8 => Ok(Expr::Negate(Box::new(arbitrary_expr(u, depth - 1)?))),
        _ => Ok(Expr::Parens(Box::new(arbitrary_expr(u, depth - 1)?))),
    }
}

impl Expr {
    fn render(&self) -> String {
        match self {
            Expr::Number(n) => format!("{n}"),
            Expr::Str(s) => format!("\"{}\"", s.replace('"', "\"\"")),
            Expr::Bool(b) => if *b { "TRUE" } else { "FALSE" }.into(),
            Expr::CellRef(r) | Expr::RangeRef(r) => r.clone(),
            Expr::Function { name, args } => {
                let a: Vec<String> = args.iter().map(|e| e.render()).collect();
                format!("{}({})", name, a.join(","))
            }
            Expr::BinOp { op, left, right } => {
                format!("{}{}{}", left.render(), op, right.render())
            }
            Expr::Negate(inner) => format!("-{}", inner.render()),
            Expr::Parens(inner) => format!("({})", inner.render()),
        }
    }
}

/// Populate a 5×10 grid with mixed seed data so formulas have something
/// to reference. The data pattern is deterministic - the fuzzer only
/// varies the formulas, not the seed values.
fn populate_seed_data(wb: &mut Workbook) {
    let sheet = wb.worksheet_mut(0).unwrap();
    for row in 0..10u32 {
        // Col A: integers 1..10
        let _ = sheet.set_cell_value_at(row, 0, CellValue::Number((row + 1) as f64));
        // Col B: negative floats
        let _ = sheet.set_cell_value_at(row, 1, CellValue::Number(-((row + 1) as f64) * 0.5));
        // Col C: strings
        let _ = sheet.set_cell_value_at(row, 2, CellValue::String(format!("txt{}", row).into()));
        // Col D: booleans (alternating)
        let _ = sheet.set_cell_value_at(row, 3, CellValue::Boolean(row % 2 == 0));
        // Col E: empty (left default) - tests ISBLANK, empty-cell coercion
    }
}

fuzz_target!(|data: &[u8]| {
    let mut u = Unstructured::new(data);

    let mut wb = Workbook::new();
    populate_seed_data(&mut wb);

    let sheet = wb.worksheet_mut(0).unwrap();

    // Generate 1-8 formulas placed in rows 11-18 (below seed data)
    let n_formulas = match u.int_in_range(1u8..=8) {
        Ok(n) => n,
        Err(_) => return,
    };

    for i in 0..n_formulas {
        let expr = match arbitrary_expr(&mut u, MAX_DEPTH) {
            Ok(e) => e,
            Err(_) => return,
        };
        let formula = format!("={}", expr.render());
        // Place in column A, rows 11..18
        let _ = sheet.set_cell_formula_at(10 + i as u32, 0, &formula);
    }

    // Also test inter-formula references: some formulas may reference
    // other generated formula cells (A11, A12, ...)
    if let Ok(true) = u.arbitrary::<bool>() {
        // Add a formula that SUMs the generated formula outputs
        let end_row = 10 + n_formulas as u32;
        let _ = sheet.set_cell_formula_at(end_row, 0, &format!("=SUM(A11:A{})", end_row));
    }

    // Run calculation - must not panic
    let opts = CalculationOptions {
        // Enable iterative calc so circular refs don't panic
        iterative: true,
        max_iterations: 10,
        max_change: 0.001,
        // Single thread for determinism
        max_threads: Some(1),
        ..Default::default()
    };
    let _ = wb.calculate_with_options(&opts);

    // Second calculation tests the cache path
    let _ = wb.calculate_with_options(&opts);
});
