//! Fuzz target for the formula parser.
//!
//! Generates structured formula strings via `Arbitrary` with depth
//! limiting to reach deep parser paths (nested expressions, cross-sheet
//! refs, structured refs, array literals) without blowing the stack.

#![no_main]
use arbitrary::Unstructured;
use libfuzzer_sys::fuzz_target;

const MAX_DEPTH: usize = 6;

/// A formula component that renders to a valid-ish formula fragment.
#[derive(Debug)]
enum Expr {
    Int(u16),
    Float(f64),
    Str(String),
    Bool(bool),
    Error(u8),
    CellRef {
        col: u16,
        row: u32,
        abs_col: bool,
        abs_row: bool,
    },
    Range {
        col1: u16,
        row1: u32,
        col2: u16,
        row2: u32,
    },
    SheetRef {
        sheet: String,
        col: u16,
        row: u32,
        quoted: bool,
    },
    Function {
        name_idx: u8,
        args: Vec<Expr>,
    },
    BinOp {
        op: u8,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    Negate(Box<Expr>),
    Percent(Box<Expr>),
    Parens(Box<Expr>),
    Array {
        rows: u8,
        cols: u8,
    },
    StructRef {
        table: String,
        column: String,
    },
    NameRef(String),
}

fn arbitrary_small_string(u: &mut Unstructured) -> arbitrary::Result<String> {
    let len = u.int_in_range(1..=12)?;
    let mut s = String::with_capacity(len);
    for _ in 0..len {
        let c = u.int_in_range(b'A'..=b'z')? as char;
        if c.is_alphanumeric() || c == '_' {
            s.push(c);
        } else {
            s.push('x');
        }
    }
    Ok(s)
}

fn arbitrary_col(u: &mut Unstructured) -> arbitrary::Result<u16> {
    u.int_in_range(0..=701)
}

fn arbitrary_row(u: &mut Unstructured) -> arbitrary::Result<u32> {
    u.int_in_range(1..=1_048_576)
}

/// Generate an Expr with bounded depth. Leaf nodes only when depth=0.
fn arbitrary_expr(u: &mut Unstructured, depth: usize) -> arbitrary::Result<Expr> {
    // At depth 0, only generate leaf nodes (no recursion)
    let max_variant = if depth == 0 { 10 } else { 16 };
    let kind: u8 = u.int_in_range(0..=max_variant)?;

    match kind {
        0 => Ok(Expr::Int(u.arbitrary()?)),
        1 => {
            let f: f64 = u.arbitrary()?;
            Ok(Expr::Float(if f.is_finite() { f } else { 0.0 }))
        }
        2 => Ok(Expr::Str(arbitrary_small_string(u)?)),
        3 => Ok(Expr::Bool(u.arbitrary()?)),
        4 => Ok(Expr::Error(u.int_in_range(0..=6)?)),
        5 => Ok(Expr::CellRef {
            col: arbitrary_col(u)?,
            row: arbitrary_row(u)?,
            abs_col: u.arbitrary()?,
            abs_row: u.arbitrary()?,
        }),
        6 => Ok(Expr::Range {
            col1: arbitrary_col(u)?,
            row1: arbitrary_row(u)?,
            col2: arbitrary_col(u)?,
            row2: arbitrary_row(u)?,
        }),
        7 => Ok(Expr::SheetRef {
            sheet: arbitrary_small_string(u)?,
            col: arbitrary_col(u)?,
            row: arbitrary_row(u)?,
            quoted: u.arbitrary()?,
        }),
        8 => Ok(Expr::Array {
            rows: u.int_in_range(1..=4)?,
            cols: u.int_in_range(1..=4)?,
        }),
        9 => Ok(Expr::StructRef {
            table: arbitrary_small_string(u)?,
            column: arbitrary_small_string(u)?,
        }),
        10 => Ok(Expr::NameRef(arbitrary_small_string(u)?)),
        // Recursive variants (only when depth > 0)
        11 => {
            let nargs = u.int_in_range(0..=4)?;
            let mut args = Vec::with_capacity(nargs);
            for _ in 0..nargs {
                args.push(arbitrary_expr(u, depth - 1)?);
            }
            Ok(Expr::Function {
                name_idx: u.int_in_range(0..=29)?,
                args,
            })
        }
        12 => Ok(Expr::BinOp {
            op: u.int_in_range(0..=11)?,
            left: Box::new(arbitrary_expr(u, depth - 1)?),
            right: Box::new(arbitrary_expr(u, depth - 1)?),
        }),
        13 => Ok(Expr::Negate(Box::new(arbitrary_expr(u, depth - 1)?))),
        14 => Ok(Expr::Percent(Box::new(arbitrary_expr(u, depth - 1)?))),
        15 | _ => Ok(Expr::Parens(Box::new(arbitrary_expr(u, depth - 1)?))),
    }
}

// --- Rendering ---

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
    "XLOOKUP",
    "FILTER",
    "SORT",
    "UNIQUE",
    "LET",
    "LAMBDA",
];

const BIN_OPS: &[&str] = &[
    "+", "-", "*", "/", "^", "=", "<>", "<", "<=", ">", ">=", "&",
];

const ERRORS: &[&str] = &[
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
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

impl Expr {
    fn render(&self) -> String {
        match self {
            Expr::Int(n) => n.to_string(),
            Expr::Float(f) => format!("{}", f),
            Expr::Str(s) => format!("\"{}\"", s.replace('"', "\"\"")),
            Expr::Bool(b) => if *b { "TRUE" } else { "FALSE" }.into(),
            Expr::Error(i) => ERRORS[*i as usize % ERRORS.len()].into(),
            Expr::CellRef {
                col,
                row,
                abs_col,
                abs_row,
            } => format!(
                "{}{}{}{}",
                if *abs_col { "$" } else { "" },
                col_to_letters(*col),
                if *abs_row { "$" } else { "" },
                row
            ),
            Expr::Range {
                col1,
                row1,
                col2,
                row2,
            } => format!(
                "{}{}:{}{}",
                col_to_letters(*col1),
                row1,
                col_to_letters(*col2),
                row2,
            ),
            Expr::SheetRef {
                sheet,
                col,
                row,
                quoted,
            } => {
                if *quoted {
                    format!("'{}'!{}{}", sheet, col_to_letters(*col), row)
                } else {
                    format!("{}!{}{}", sheet, col_to_letters(*col), row)
                }
            }
            Expr::Function { name_idx, args } => {
                let name = FUNC_NAMES[*name_idx as usize % FUNC_NAMES.len()];
                let arg_strs: Vec<String> = args.iter().map(|a| a.render()).collect();
                format!("{}({})", name, arg_strs.join(","))
            }
            Expr::BinOp { op, left, right } => {
                let op_str = BIN_OPS[*op as usize % BIN_OPS.len()];
                format!("{}{}{}", left.render(), op_str, right.render())
            }
            Expr::Negate(inner) => format!("-{}", inner.render()),
            Expr::Percent(inner) => format!("{}%", inner.render()),
            Expr::Parens(inner) => format!("({})", inner.render()),
            Expr::Array { rows, cols } => {
                let mut row_strs = Vec::new();
                for r in 0..*rows as usize {
                    let mut col_strs = Vec::new();
                    for c in 0..*cols as usize {
                        col_strs.push(format!("{}", r * *cols as usize + c + 1));
                    }
                    row_strs.push(col_strs.join(","));
                }
                format!("{{{}}}", row_strs.join(";"))
            }
            Expr::StructRef { table, column } => format!("{}[{}]", table, column),
            Expr::NameRef(name) => name.clone(),
        }
    }
}

fuzz_target!(|data: &[u8]| {
    let mut u = Unstructured::new(data);

    // Strategy 1: Generate structured formula from Arbitrary with depth limit
    if let Ok(expr) = arbitrary_expr(&mut u, MAX_DEPTH) {
        let formula = format!("={}", expr.render());
        let _ = duke_sheets_formula::parse_formula(&formula);
    }

    // Strategy 2: Raw bytes as a string (catches weird Unicode, etc.)
    if let Ok(s) = std::str::from_utf8(data) {
        let _ = duke_sheets_formula::parse_formula(s);
        if !s.starts_with('=') {
            let _ = duke_sheets_formula::parse_formula(&format!("={}", s));
        }
    }
});
