//! Decompile parsed BIFF8 formula tokens (RPN) into an infix formula string.
//!
//! Uses a stack machine: operands push strings onto the stack, operators pop
//! operands and push the formatted result. At the end, the single remaining
//! stack entry is the formula text.

use super::function_table;
use super::token_parser::ParsedToken;

/// Operator precedence levels (higher = binds tighter).
///
/// Used to determine when parentheses are needed around binary operator
/// operands.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Prec(u8);

const PREC_COMPARE: Prec = Prec(1); // < <= = >= > <>
const PREC_CONCAT: Prec = Prec(2); // &
const PREC_ADD: Prec = Prec(3); // + -
const PREC_MUL: Prec = Prec(4); // * /
const PREC_POWER: Prec = Prec(5); // ^
const PREC_UNARY: Prec = Prec(6); // unary +, -, %
const PREC_RANGE: Prec = Prec(7); // : (range), (space) (intersection), , (union)
const PREC_ATOM: Prec = Prec(8); // literals, refs, function calls

/// A value on the decompiler stack.
#[derive(Debug, Clone)]
struct StackEntry {
    /// The formatted text for this sub-expression.
    text: String,
    /// The precedence level of the outermost operator, used to decide
    /// whether wrapping in parentheses is needed.
    prec: Prec,
}

impl StackEntry {
    fn atom(text: String) -> Self {
        Self {
            text,
            prec: PREC_ATOM,
        }
    }

    fn with_prec(text: String, prec: Prec) -> Self {
        Self { text, prec }
    }

    /// Return the text, wrapping in parentheses if the entry's precedence
    /// is lower than `min_prec`.
    fn text_with_parens(&self, min_prec: Prec) -> String {
        if self.prec < min_prec {
            format!("({})", self.text)
        } else {
            self.text.clone()
        }
    }
}

/// Format a 0-based column index as an Excel column letter (A, B, ..., Z, AA, ...).
fn col_to_letter(mut col: u16) -> String {
    let mut result = String::new();
    loop {
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result
}

/// Format a cell reference (row, col) in A1 notation.
fn format_ref(row: u16, col: u16) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

/// Format an area reference.
fn format_area(first_row: u16, last_row: u16, first_col: u16, last_col: u16) -> String {
    if first_row == last_row && first_col == last_col {
        format_ref(first_row, first_col)
    } else {
        format!(
            "{}:{}",
            format_ref(first_row, first_col),
            format_ref(last_row, last_col)
        )
    }
}

/// Format a BIFF8 error code byte as an Excel error string.
fn format_error(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#UNKNOWN!",
    }
}

/// Format a floating-point number for display in a formula.
///
/// Strips trailing zeros and unnecessary decimal points, matching Excel's
/// compact representation.
fn format_number(val: f64) -> String {
    if val == val.trunc() && val.abs() < 1e15 {
        // Integer value — no decimal point needed
        format!("{}", val as i64)
    } else {
        // Use enough precision to round-trip, then strip trailing zeros
        let s = format!("{}", val);
        s
    }
}

/// Decompile a sequence of parsed tokens into an infix formula string.
///
/// The `sheet_names` slice is used for 3D references (Phase 2) — for now
/// it can be empty.
///
/// Returns the formula text (without leading `=`), or an empty string if
/// decompilation fails.
pub fn decompile(tokens: &[ParsedToken], sheet_names: &[String]) -> String {
    let mut stack: Vec<StackEntry> = Vec::new();

    for token in tokens {
        match token {
            // ---- Binary operators ----
            ParsedToken::Add => binary_op(&mut stack, "+", PREC_ADD),
            ParsedToken::Sub => binary_op(&mut stack, "-", PREC_ADD),
            ParsedToken::Mul => binary_op(&mut stack, "*", PREC_MUL),
            ParsedToken::Div => binary_op(&mut stack, "/", PREC_MUL),
            ParsedToken::Power => binary_op_right(&mut stack, "^", PREC_POWER),
            ParsedToken::Concat => binary_op(&mut stack, "&", PREC_CONCAT),
            ParsedToken::Lt => binary_op(&mut stack, "<", PREC_COMPARE),
            ParsedToken::Le => binary_op(&mut stack, "<=", PREC_COMPARE),
            ParsedToken::Eq => binary_op(&mut stack, "=", PREC_COMPARE),
            ParsedToken::Ge => binary_op(&mut stack, ">=", PREC_COMPARE),
            ParsedToken::Gt => binary_op(&mut stack, ">", PREC_COMPARE),
            ParsedToken::Ne => binary_op(&mut stack, "<>", PREC_COMPARE),
            ParsedToken::Isect => {
                // Intersection uses a space as operator
                let right = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                let left = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                stack.push(StackEntry::with_prec(
                    format!(
                        "{} {}",
                        left.text_with_parens(PREC_RANGE),
                        right.text_with_parens(PREC_RANGE)
                    ),
                    PREC_RANGE,
                ));
            }
            ParsedToken::List => {
                // Union operator (comma in reference context)
                let right = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                let left = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                stack.push(StackEntry::with_prec(
                    format!("{},{}", left.text, right.text),
                    PREC_RANGE,
                ));
            }
            ParsedToken::Range => {
                // Range operator (:)
                let right = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                let left = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                stack.push(StackEntry::with_prec(
                    format!("{}:{}", left.text, right.text),
                    PREC_RANGE,
                ));
            }

            // ---- Unary operators ----
            ParsedToken::Uplus => {
                let operand = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                stack.push(StackEntry::with_prec(
                    format!("+{}", operand.text_with_parens(PREC_UNARY)),
                    PREC_UNARY,
                ));
            }
            ParsedToken::Uminus => {
                let operand = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                stack.push(StackEntry::with_prec(
                    format!("-{}", operand.text_with_parens(PREC_UNARY)),
                    PREC_UNARY,
                ));
            }
            ParsedToken::Percent => {
                let operand = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                stack.push(StackEntry::with_prec(
                    format!("{}%", operand.text_with_parens(PREC_UNARY)),
                    PREC_UNARY,
                ));
            }
            ParsedToken::Paren => {
                let operand = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
                stack.push(StackEntry::atom(format!("({})", operand.text)));
            }

            // ---- Constants ----
            ParsedToken::MissArg => {
                stack.push(StackEntry::atom(String::new()));
            }
            ParsedToken::Str(s) => {
                // Escape double quotes inside the string
                let escaped = s.replace('"', "\"\"");
                stack.push(StackEntry::atom(format!("\"{}\"", escaped)));
            }
            ParsedToken::Err(code) => {
                stack.push(StackEntry::atom(format_error(*code).to_string()));
            }
            ParsedToken::Bool(val) => {
                stack.push(StackEntry::atom(
                    if *val { "TRUE" } else { "FALSE" }.to_string(),
                ));
            }
            ParsedToken::Int(val) => {
                stack.push(StackEntry::atom(val.to_string()));
            }
            ParsedToken::Num(val) => {
                stack.push(StackEntry::atom(format_number(*val)));
            }

            // ---- Cell references ----
            ParsedToken::Ref { row, col, .. } => {
                stack.push(StackEntry::atom(format_ref(*row, *col)));
            }
            ParsedToken::Area {
                first_row,
                last_row,
                first_col,
                last_col,
                ..
            } => {
                stack.push(StackEntry::atom(format_area(
                    *first_row, *last_row, *first_col, *last_col,
                )));
            }
            ParsedToken::RefErr | ParsedToken::AreaErr => {
                stack.push(StackEntry::atom("#REF!".to_string()));
            }

            // ---- Functions ----
            ParsedToken::Func { func_idx } => {
                let name = function_table::function_name(*func_idx);
                let argc = function_table::function_argc(*func_idx);
                let arg_count = if argc <= 253 { argc as usize } else { 0 };
                decompile_function(&mut stack, name, *func_idx, arg_count);
            }
            ParsedToken::FuncVar { argc, func_idx } => {
                let raw_idx = *func_idx & 0x7FFF; // strip CE flag
                let name = function_table::function_name(raw_idx);
                decompile_function(&mut stack, name, raw_idx, *argc as usize);
            }

            // ---- tAttr sub-types ----
            ParsedToken::AttrSum => {
                // Optimized SUM with single argument
                decompile_function(&mut stack, "SUM", 4, 1);
            }
            ParsedToken::AttrVolatile
            | ParsedToken::AttrIf { .. }
            | ParsedToken::AttrChoose { .. }
            | ParsedToken::AttrSkip { .. }
            | ParsedToken::AttrAssign => {
                // No-ops for decompilation — these are optimization hints
            }
            ParsedToken::AttrSpace { .. } => {
                // Whitespace preservation — ignore for now
            }

            // ---- Phase 2 stubs ----
            ParsedToken::Name { name_idx } => {
                // Without NAME records parsed, emit a placeholder
                stack.push(StackEntry::atom(format!("_name{}", name_idx)));
            }
            ParsedToken::NameX {
                extern_sheet_idx: _,
                name_idx,
            } => {
                stack.push(StackEntry::atom(format!("_namex{}", name_idx)));
            }
            ParsedToken::Ref3d {
                extern_sheet_idx,
                row,
                col,
                ..
            } => {
                let ref_str = format_ref(*row, *col);
                let sheet = resolve_sheet_name(sheet_names, *extern_sheet_idx);
                stack.push(StackEntry::atom(format!("{}!{}", sheet, ref_str)));
            }
            ParsedToken::Area3d {
                extern_sheet_idx,
                first_row,
                last_row,
                first_col,
                last_col,
                ..
            } => {
                let area_str = format_area(*first_row, *last_row, *first_col, *last_col);
                let sheet = resolve_sheet_name(sheet_names, *extern_sheet_idx);
                stack.push(StackEntry::atom(format!("{}!{}", sheet, area_str)));
            }
            ParsedToken::RefErr3d { .. } | ParsedToken::AreaErr3d { .. } => {
                stack.push(StackEntry::atom("#REF!".to_string()));
            }

            // tExp — array/shared formula indicator
            ParsedToken::Exp { .. } => {
                // This means the cell's formula is stored elsewhere
                // (ARRAY or SHAREDFMLA record). We can't decompile it from
                // the FORMULA record alone. Push empty so we don't break the stack.
            }

            // MemFunc — no-op for decompilation; sub-expression tokens handle it
            ParsedToken::MemFunc { .. } => {}

            ParsedToken::Unknown(_) => {
                stack.push(StackEntry::atom("<?>".to_string()));
            }
        }
    }

    // The final result is the single entry remaining on the stack
    match stack.len() {
        0 => String::new(),
        1 => stack.pop().unwrap().text,
        _ => {
            // Multiple entries — shouldn't happen for valid formulas.
            // Join them to avoid losing data.
            stack
                .into_iter()
                .map(|e| e.text)
                .collect::<Vec<_>>()
                .join("")
        }
    }
}

/// Helper: pop two operands, format as "left op right" with precedence-based parens.
fn binary_op(stack: &mut Vec<StackEntry>, op: &str, prec: Prec) {
    let right = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
    let left = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));

    // For left-associative operators: left needs parens if strictly lower prec,
    // right needs parens if lower-or-equal (to handle left-associativity).
    let left_text = left.text_with_parens(prec);
    let right_text = right.text_with_parens(Prec(prec.0 + 1));

    stack.push(StackEntry::with_prec(
        format!("{}{}{}", left_text, op, right_text),
        prec,
    ));
}

/// Helper: pop two operands for right-associative operator (^).
fn binary_op_right(stack: &mut Vec<StackEntry>, op: &str, prec: Prec) {
    let right = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
    let left = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));

    let left_text = left.text_with_parens(Prec(prec.0 + 1));
    let right_text = right.text_with_parens(prec);

    stack.push(StackEntry::with_prec(
        format!("{}{}{}", left_text, op, right_text),
        prec,
    ));
}

/// Helper: decompile a function call by popping `argc` arguments from the stack.
fn decompile_function(stack: &mut Vec<StackEntry>, name: &str, func_idx: u16, argc: usize) {
    // Collect args in reverse (they were pushed in order, so stack has last arg on top)
    let mut args: Vec<String> = Vec::with_capacity(argc);
    for _ in 0..argc {
        let entry = stack.pop().unwrap_or_else(|| StackEntry::atom("?".into()));
        args.push(entry.text);
    }
    args.reverse();

    let display_name = if name.is_empty() {
        format!("_xlfn.{}", func_idx)
    } else {
        name.to_string()
    };

    stack.push(StackEntry::atom(format!(
        "{}({})",
        display_name,
        args.join(",")
    )));
}

/// Resolve an EXTERNSHEET index to a sheet name string.
///
/// For Phase 1 without EXTERNSHEET records, we just use the sheet_names
/// array as a simple lookup: the EXTERNSHEET index often maps directly to
/// the sheet index for internal references.
///
/// Phase 2 will properly parse EXTERNSHEET records and do the full
/// (sup_book, first_sheet, last_sheet) resolution.
fn resolve_sheet_name(sheet_names: &[String], extern_sheet_idx: u16) -> String {
    let idx = extern_sheet_idx as usize;
    if idx < sheet_names.len() {
        let name = &sheet_names[idx];
        // Quote sheet name if it contains spaces or special characters
        if needs_quoting(name) {
            format!("'{}'", name.replace('\'', "''"))
        } else {
            name.clone()
        }
    } else {
        format!("_sheet{}", extern_sheet_idx)
    }
}

/// Check if a sheet name needs to be quoted in a formula.
fn needs_quoting(name: &str) -> bool {
    if name.is_empty() {
        return true;
    }
    // Must quote if contains spaces, special chars, or starts with a digit
    name.contains(' ')
        || name.contains('\'')
        || name.contains('!')
        || name.contains(':')
        || name.contains('[')
        || name.contains(']')
        || name.chars().next().map_or(false, |c| c.is_ascii_digit())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::biff::formula::token_parser::ParsedToken;

    #[test]
    fn test_simple_addition() {
        // =1+2
        let tokens = vec![ParsedToken::Int(1), ParsedToken::Int(2), ParsedToken::Add];
        assert_eq!(decompile(&tokens, &[]), "1+2");
    }

    #[test]
    fn test_subtraction() {
        let tokens = vec![ParsedToken::Int(5), ParsedToken::Int(3), ParsedToken::Sub];
        assert_eq!(decompile(&tokens, &[]), "5-3");
    }

    #[test]
    fn test_precedence_mul_add() {
        // =A1+B1*C1 → RPN: A1 B1 C1 * +
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Ref {
                row: 0,
                col: 1,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Ref {
                row: 0,
                col: 2,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Mul,
            ParsedToken::Add,
        ];
        assert_eq!(decompile(&tokens, &[]), "A1+B1*C1");
    }

    #[test]
    fn test_precedence_needs_parens() {
        // =(A1+B1)*C1 → RPN: A1 B1 + C1 *
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Ref {
                row: 0,
                col: 1,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Add,
            ParsedToken::Ref {
                row: 0,
                col: 2,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Mul,
        ];
        assert_eq!(decompile(&tokens, &[]), "(A1+B1)*C1");
    }

    #[test]
    fn test_unary_minus() {
        // =-A1
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Uminus,
        ];
        assert_eq!(decompile(&tokens, &[]), "-A1");
    }

    #[test]
    fn test_percent() {
        // =50%
        let tokens = vec![ParsedToken::Int(50), ParsedToken::Percent];
        assert_eq!(decompile(&tokens, &[]), "50%");
    }

    #[test]
    fn test_string_literal() {
        let tokens = vec![ParsedToken::Str("hello".to_string())];
        assert_eq!(decompile(&tokens, &[]), "\"hello\"");
    }

    #[test]
    fn test_string_with_quotes() {
        let tokens = vec![ParsedToken::Str("say \"hi\"".to_string())];
        assert_eq!(decompile(&tokens, &[]), "\"say \"\"hi\"\"\"");
    }

    #[test]
    fn test_bool_true() {
        let tokens = vec![ParsedToken::Bool(true)];
        assert_eq!(decompile(&tokens, &[]), "TRUE");
    }

    #[test]
    fn test_error_constant() {
        let tokens = vec![ParsedToken::Err(0x2A)];
        assert_eq!(decompile(&tokens, &[]), "#N/A");
    }

    #[test]
    fn test_float_number() {
        let tokens = vec![ParsedToken::Num(3.14)];
        assert_eq!(decompile(&tokens, &[]), "3.14");
    }

    #[test]
    fn test_integer_float() {
        let tokens = vec![ParsedToken::Num(100.0)];
        assert_eq!(decompile(&tokens, &[]), "100");
    }

    #[test]
    fn test_area_ref() {
        // =A1:C3
        let tokens = vec![ParsedToken::Area {
            first_row: 0,
            last_row: 2,
            first_col: 0,
            last_col: 2,
            first_row_rel: true,
            first_col_rel: true,
            last_row_rel: true,
            last_col_rel: true,
        }];
        assert_eq!(decompile(&tokens, &[]), "A1:C3");
    }

    #[test]
    fn test_func_sum() {
        // =SUM(A1:A10)
        let tokens = vec![
            ParsedToken::Area {
                first_row: 0,
                last_row: 9,
                first_col: 0,
                last_col: 0,
                first_row_rel: true,
                first_col_rel: true,
                last_row_rel: true,
                last_col_rel: true,
            },
            ParsedToken::AttrSum,
        ];
        assert_eq!(decompile(&tokens, &[]), "SUM(A1:A10)");
    }

    #[test]
    fn test_funcvar_if() {
        // =IF(A1>0,1,0)
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Int(0),
            ParsedToken::Gt,
            ParsedToken::AttrIf { offset: 0 },
            ParsedToken::Int(1),
            ParsedToken::AttrSkip { offset: 0 },
            ParsedToken::Int(0),
            ParsedToken::FuncVar {
                argc: 3,
                func_idx: 1,
            },
        ];
        assert_eq!(decompile(&tokens, &[]), "IF(A1>0,1,0)");
    }

    #[test]
    fn test_func_len() {
        // =LEN(A1) — fixed-arg function
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Func { func_idx: 32 },
        ];
        assert_eq!(decompile(&tokens, &[]), "LEN(A1)");
    }

    #[test]
    fn test_missing_arg() {
        // =MATCH(1,A1:A10,) — trailing missing arg
        let tokens = vec![
            ParsedToken::Int(1),
            ParsedToken::Area {
                first_row: 0,
                last_row: 9,
                first_col: 0,
                last_col: 0,
                first_row_rel: true,
                first_col_rel: true,
                last_row_rel: true,
                last_col_rel: true,
            },
            ParsedToken::MissArg,
            ParsedToken::FuncVar {
                argc: 3,
                func_idx: 64,
            },
        ];
        assert_eq!(decompile(&tokens, &[]), "MATCH(1,A1:A10,)");
    }

    #[test]
    fn test_concat_operator() {
        // ="hello"&A1
        let tokens = vec![
            ParsedToken::Str("hello".to_string()),
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Concat,
        ];
        assert_eq!(decompile(&tokens, &[]), "\"hello\"&A1");
    }

    #[test]
    fn test_comparison_operators() {
        // =A1>=B1
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Ref {
                row: 0,
                col: 1,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Ge,
        ];
        assert_eq!(decompile(&tokens, &[]), "A1>=B1");
    }

    #[test]
    fn test_nested_functions() {
        // =SUM(LEN(A1),LEN(B1)) — SUM with 2 args via FuncVar
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Func { func_idx: 32 }, // LEN
            ParsedToken::Ref {
                row: 0,
                col: 1,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Func { func_idx: 32 }, // LEN
            ParsedToken::FuncVar {
                argc: 2,
                func_idx: 4, // SUM
            },
        ];
        assert_eq!(decompile(&tokens, &[]), "SUM(LEN(A1),LEN(B1))");
    }

    #[test]
    fn test_power_right_associative() {
        // =2^3^4 → RPN: 2 3 4 ^ ^
        // Right-associative: should be 2^3^4 (not (2^3)^4)
        let tokens = vec![
            ParsedToken::Int(2),
            ParsedToken::Int(3),
            ParsedToken::Int(4),
            ParsedToken::Power,
            ParsedToken::Power,
        ];
        assert_eq!(decompile(&tokens, &[]), "2^3^4");
    }

    #[test]
    fn test_ref_err() {
        let tokens = vec![ParsedToken::RefErr];
        assert_eq!(decompile(&tokens, &[]), "#REF!");
    }

    #[test]
    fn test_paren_display() {
        // tParen wraps top of stack
        let tokens = vec![ParsedToken::Int(42), ParsedToken::Paren];
        assert_eq!(decompile(&tokens, &[]), "(42)");
    }

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(255), "IV"); // max BIFF8 column
    }

    #[test]
    fn test_empty_tokens() {
        assert_eq!(decompile(&[], &[]), "");
    }

    #[test]
    fn test_volatile_ignored() {
        // tAttrVolatile before NOW() — should not appear in output
        let tokens = vec![
            ParsedToken::AttrVolatile,
            ParsedToken::Func { func_idx: 74 }, // NOW
        ];
        assert_eq!(decompile(&tokens, &[]), "NOW()");
    }

    #[test]
    fn test_left_associativity_subtraction() {
        // =1-2-3 → RPN: 1 2 - 3 -
        // Should produce "1-2-3" not "1-(2-3)"
        let tokens = vec![
            ParsedToken::Int(1),
            ParsedToken::Int(2),
            ParsedToken::Sub,
            ParsedToken::Int(3),
            ParsedToken::Sub,
        ];
        assert_eq!(decompile(&tokens, &[]), "1-2-3");
    }

    #[test]
    fn test_right_operand_needs_parens() {
        // =1-(2+3) → RPN: 1 2 3 + -
        let tokens = vec![
            ParsedToken::Int(1),
            ParsedToken::Int(2),
            ParsedToken::Int(3),
            ParsedToken::Add,
            ParsedToken::Sub,
        ];
        assert_eq!(decompile(&tokens, &[]), "1-(2+3)");
    }

    #[test]
    fn test_3d_ref_with_sheet_names() {
        let sheet_names = vec!["Sheet1".to_string(), "Sheet2".to_string()];
        let tokens = vec![ParsedToken::Ref3d {
            extern_sheet_idx: 1,
            row: 0,
            col: 0,
            row_relative: false,
            col_relative: false,
        }];
        assert_eq!(decompile(&tokens, &sheet_names), "Sheet2!A1");
    }

    #[test]
    fn test_3d_ref_quoted_sheet() {
        let sheet_names = vec!["My Sheet".to_string()];
        let tokens = vec![ParsedToken::Ref3d {
            extern_sheet_idx: 0,
            row: 4,
            col: 1,
            row_relative: false,
            col_relative: false,
        }];
        assert_eq!(decompile(&tokens, &sheet_names), "'My Sheet'!B5");
    }
}
