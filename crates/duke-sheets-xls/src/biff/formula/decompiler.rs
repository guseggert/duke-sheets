//! Decompile parsed BIFF8 formula tokens (RPN) into an infix formula string.
//!
//! Uses a stack machine: operands push strings onto the stack, operators pop
//! operands and push the formatted result. At the end, the single remaining
//! stack entry is the formula text.

use super::function_table;
use super::token_parser::ParsedToken;
use super::{FormulaContext, SupBook};

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

/// Format a cell reference (row, col) in A1 notation with optional `$` for
/// absolute dimensions.  When `row_rel` is false the row is absolute (`$1`),
/// when `col_rel` is false the column is absolute (`$A`).
fn format_ref(row: u16, col: u16, row_rel: bool, col_rel: bool) -> String {
    let col_prefix = if col_rel { "" } else { "$" };
    let row_prefix = if row_rel { "" } else { "$" };
    format!(
        "{}{}{}{}",
        col_prefix,
        col_to_letter(col),
        row_prefix,
        row + 1
    )
}

/// Format an area reference with optional `$` for absolute dimensions.
fn format_area(
    first_row: u16,
    last_row: u16,
    first_col: u16,
    last_col: u16,
    first_row_rel: bool,
    first_col_rel: bool,
    last_row_rel: bool,
    last_col_rel: bool,
) -> String {
    if first_row == last_row && first_col == last_col {
        format_ref(first_row, first_col, first_row_rel, first_col_rel)
    } else {
        format!(
            "{}:{}",
            format_ref(first_row, first_col, first_row_rel, first_col_rel),
            format_ref(last_row, last_col, last_row_rel, last_col_rel)
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
/// The `ctx` provides EXTERNSHEET/SUPBOOK/NAME data for resolving 3D
/// references and defined names.
///
/// Returns the formula text (without leading `=`), or an empty string if
/// decompilation fails.
pub fn decompile(tokens: &[ParsedToken], ctx: &FormulaContext) -> String {
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
            ParsedToken::Ref {
                row,
                col,
                row_relative,
                col_relative,
            } => {
                stack.push(StackEntry::atom(format_ref(
                    *row,
                    *col,
                    *row_relative,
                    *col_relative,
                )));
            }
            ParsedToken::Area {
                first_row,
                last_row,
                first_col,
                last_col,
                first_row_rel,
                first_col_rel,
                last_row_rel,
                last_col_rel,
            } => {
                stack.push(StackEntry::atom(format_area(
                    *first_row,
                    *last_row,
                    *first_col,
                    *last_col,
                    *first_row_rel,
                    *first_col_rel,
                    *last_row_rel,
                    *last_col_rel,
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

            // ---- Defined names ----
            ParsedToken::Name { name_idx } => {
                // name_idx is 1-based
                let idx = (*name_idx as usize).wrapping_sub(1);
                let name = ctx
                    .names
                    .get(idx)
                    .map(|nr| nr.name.clone())
                    .unwrap_or_else(|| format!("_name{}", name_idx));
                stack.push(StackEntry::atom(name));
            }
            ParsedToken::NameX {
                extern_sheet_idx,
                name_idx,
            } => {
                // NameX: name_idx is 1-based into the external workbook's name table.
                // For self-ref SUPBOOKs this is the same as tName.
                let resolved = resolve_namex(ctx, *extern_sheet_idx, *name_idx);
                stack.push(StackEntry::atom(resolved));
            }

            // ---- 3D references ----
            ParsedToken::Ref3d {
                extern_sheet_idx,
                row,
                col,
                row_relative,
                col_relative,
            } => {
                let ref_str = format_ref(*row, *col, *row_relative, *col_relative);
                let sheet = resolve_sheet_prefix(ctx, *extern_sheet_idx);
                stack.push(StackEntry::atom(format!("{}!{}", sheet, ref_str)));
            }
            ParsedToken::Area3d {
                extern_sheet_idx,
                first_row,
                last_row,
                first_col,
                last_col,
                first_row_rel,
                first_col_rel,
                last_row_rel,
                last_col_rel,
            } => {
                let area_str = format_area(
                    *first_row,
                    *last_row,
                    *first_col,
                    *last_col,
                    *first_row_rel,
                    *first_col_rel,
                    *last_row_rel,
                    *last_col_rel,
                );
                let sheet = resolve_sheet_prefix(ctx, *extern_sheet_idx);
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

/// Resolve an EXTERNSHEET index to a sheet prefix string for 3D references.
///
/// Uses the EXTERNSHEET → SUPBOOK → sheet name resolution chain:
/// 1. Look up the EXTERNSHEET entry by `extern_sheet_idx`
/// 2. Get the SUPBOOK it points to
/// 3. For self-ref SUPBOOKs, resolve `itabFirst`/`itabLast` to sheet names
/// 4. For multi-sheet ranges (first != last), emit `"Sheet1:Sheet3"`
fn resolve_sheet_prefix(ctx: &FormulaContext, extern_sheet_idx: u16) -> String {
    let eidx = extern_sheet_idx as usize;

    // If no EXTERNSHEET data, fall back to direct index into sheet_names
    if ctx.extern_sheet.is_empty() {
        return direct_sheet_lookup(&ctx.sheet_names, extern_sheet_idx);
    }

    let entry = match ctx.extern_sheet.get(eidx) {
        Some(e) => e,
        None => return format!("_sheet{}", extern_sheet_idx),
    };

    let supbook = match ctx.supbooks.get(entry.sup_book_idx as usize) {
        Some(sb) => sb,
        None => return format!("_sheet{}", extern_sheet_idx),
    };

    match supbook {
        SupBook::SelfRef { .. } => {
            // 0xFFFE means workbook-level (e.g. workbook-scoped name) — no sheet prefix
            if entry.first_sheet == 0xFFFE {
                return String::new();
            }

            let first_name = direct_sheet_lookup(&ctx.sheet_names, entry.first_sheet);

            if entry.first_sheet == entry.last_sheet {
                // Single sheet
                first_name
            } else {
                // Multi-sheet range: Sheet1:Sheet3
                // If either name needs quoting, quote the whole range
                let first_raw = ctx
                    .sheet_names
                    .get(entry.first_sheet as usize)
                    .cloned()
                    .unwrap_or_default();
                let last_raw = ctx
                    .sheet_names
                    .get(entry.last_sheet as usize)
                    .cloned()
                    .unwrap_or_default();
                if needs_quoting(&first_raw) || needs_quoting(&last_raw) {
                    format!(
                        "'{}'",
                        format!(
                            "{}:{}",
                            first_raw.replace('\'', "''"),
                            last_raw.replace('\'', "''")
                        )
                    )
                } else {
                    format!("{}:{}", first_raw, last_raw)
                }
            }
        }
        SupBook::AddIn => {
            // Add-in function — no sheet prefix needed
            String::new()
        }
        SupBook::External { path, sheets } => {
            // External workbook reference: [path]SheetName
            let sheet_name = sheets
                .get(entry.first_sheet as usize)
                .cloned()
                .unwrap_or_else(|| format!("_extsheet{}", entry.first_sheet));
            format!("[{}]{}", path, sheet_name)
        }
    }
}

/// Direct lookup into sheet_names with quoting.
fn direct_sheet_lookup(sheet_names: &[String], idx: u16) -> String {
    let i = idx as usize;
    if i < sheet_names.len() {
        let name = &sheet_names[i];
        if needs_quoting(name) {
            format!("'{}'", name.replace('\'', "''"))
        } else {
            name.clone()
        }
    } else {
        format!("_sheet{}", idx)
    }
}

/// Resolve a NameX token (external name reference).
///
/// For self-ref SUPBOOKs, the name_idx is 1-based into the workbook's
/// NAME record array (same as tName). For external SUPBOOKs, we emit
/// a placeholder since we don't parse external name tables.
fn resolve_namex(ctx: &FormulaContext, extern_sheet_idx: u16, name_idx: u16) -> String {
    let eidx = extern_sheet_idx as usize;

    if let Some(entry) = ctx.extern_sheet.get(eidx) {
        if let Some(supbook) = ctx.supbooks.get(entry.sup_book_idx as usize) {
            match supbook {
                SupBook::SelfRef { .. } | SupBook::AddIn => {
                    // 1-based index into workbook's NAME records
                    let idx = (name_idx as usize).wrapping_sub(1);
                    return ctx
                        .names
                        .get(idx)
                        .map(|nr| nr.name.clone())
                        .unwrap_or_else(|| format!("_namex{}", name_idx));
                }
                SupBook::External { path, .. } => {
                    return format!("[{}]_name{}", path, name_idx);
                }
            }
        }
    }

    format!("_namex{}", name_idx)
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
    use crate::biff::formula::{ExternSheetEntry, NameRecord};

    /// Create an empty FormulaContext for tests that don't need sheet/name data.
    fn empty_ctx() -> FormulaContext {
        FormulaContext::new(vec![])
    }

    /// Create a FormulaContext with just sheet names (no EXTERNSHEET/SUPBOOK).
    /// This tests the fallback path (direct index into sheet_names).
    fn ctx_with_sheets(names: Vec<String>) -> FormulaContext {
        FormulaContext::new(names)
    }

    /// Create a FormulaContext with self-ref SUPBOOK and EXTERNSHEET entries.
    fn ctx_with_self_ref(sheet_names: Vec<String>) -> FormulaContext {
        let sheet_count = sheet_names.len() as u16;
        let supbooks = vec![SupBook::SelfRef { sheet_count }];
        // Create one EXTERNSHEET entry per sheet, all pointing to supbook 0
        let extern_sheet: Vec<ExternSheetEntry> = (0..sheet_count)
            .map(|i| ExternSheetEntry {
                sup_book_idx: 0,
                first_sheet: i,
                last_sheet: i,
            })
            .collect();
        FormulaContext {
            sheet_names,
            extern_sheet,
            supbooks,
            names: Vec::new(),
        }
    }

    #[test]
    fn test_simple_addition() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Int(1), ParsedToken::Int(2), ParsedToken::Add];
        assert_eq!(decompile(&tokens, &ctx), "1+2");
    }

    #[test]
    fn test_subtraction() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Int(5), ParsedToken::Int(3), ParsedToken::Sub];
        assert_eq!(decompile(&tokens, &ctx), "5-3");
    }

    #[test]
    fn test_precedence_mul_add() {
        let ctx = empty_ctx();
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
        assert_eq!(decompile(&tokens, &ctx), "A1+B1*C1");
    }

    #[test]
    fn test_precedence_needs_parens() {
        let ctx = empty_ctx();
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
        assert_eq!(decompile(&tokens, &ctx), "(A1+B1)*C1");
    }

    #[test]
    fn test_unary_minus() {
        let ctx = empty_ctx();
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Uminus,
        ];
        assert_eq!(decompile(&tokens, &ctx), "-A1");
    }

    #[test]
    fn test_percent() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Int(50), ParsedToken::Percent];
        assert_eq!(decompile(&tokens, &ctx), "50%");
    }

    #[test]
    fn test_string_literal() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Str("hello".to_string())];
        assert_eq!(decompile(&tokens, &ctx), "\"hello\"");
    }

    #[test]
    fn test_string_with_quotes() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Str("say \"hi\"".to_string())];
        assert_eq!(decompile(&tokens, &ctx), "\"say \"\"hi\"\"\"");
    }

    #[test]
    fn test_bool_true() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Bool(true)];
        assert_eq!(decompile(&tokens, &ctx), "TRUE");
    }

    #[test]
    fn test_error_constant() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Err(0x2A)];
        assert_eq!(decompile(&tokens, &ctx), "#N/A");
    }

    #[test]
    fn test_float_number() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Num(3.14)];
        assert_eq!(decompile(&tokens, &ctx), "3.14");
    }

    #[test]
    fn test_integer_float() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Num(100.0)];
        assert_eq!(decompile(&tokens, &ctx), "100");
    }

    #[test]
    fn test_area_ref() {
        let ctx = empty_ctx();
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
        assert_eq!(decompile(&tokens, &ctx), "A1:C3");
    }

    #[test]
    fn test_func_sum() {
        let ctx = empty_ctx();
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
        assert_eq!(decompile(&tokens, &ctx), "SUM(A1:A10)");
    }

    #[test]
    fn test_funcvar_if() {
        let ctx = empty_ctx();
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
        assert_eq!(decompile(&tokens, &ctx), "IF(A1>0,1,0)");
    }

    #[test]
    fn test_func_len() {
        let ctx = empty_ctx();
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Func { func_idx: 32 },
        ];
        assert_eq!(decompile(&tokens, &ctx), "LEN(A1)");
    }

    #[test]
    fn test_missing_arg() {
        let ctx = empty_ctx();
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
        assert_eq!(decompile(&tokens, &ctx), "MATCH(1,A1:A10,)");
    }

    #[test]
    fn test_concat_operator() {
        let ctx = empty_ctx();
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
        assert_eq!(decompile(&tokens, &ctx), "\"hello\"&A1");
    }

    #[test]
    fn test_comparison_operators() {
        let ctx = empty_ctx();
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
        assert_eq!(decompile(&tokens, &ctx), "A1>=B1");
    }

    #[test]
    fn test_nested_functions() {
        let ctx = empty_ctx();
        let tokens = vec![
            ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Func { func_idx: 32 },
            ParsedToken::Ref {
                row: 0,
                col: 1,
                row_relative: true,
                col_relative: true,
            },
            ParsedToken::Func { func_idx: 32 },
            ParsedToken::FuncVar {
                argc: 2,
                func_idx: 4,
            },
        ];
        assert_eq!(decompile(&tokens, &ctx), "SUM(LEN(A1),LEN(B1))");
    }

    #[test]
    fn test_power_right_associative() {
        let ctx = empty_ctx();
        let tokens = vec![
            ParsedToken::Int(2),
            ParsedToken::Int(3),
            ParsedToken::Int(4),
            ParsedToken::Power,
            ParsedToken::Power,
        ];
        assert_eq!(decompile(&tokens, &ctx), "2^3^4");
    }

    #[test]
    fn test_ref_err() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::RefErr];
        assert_eq!(decompile(&tokens, &ctx), "#REF!");
    }

    #[test]
    fn test_paren_display() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Int(42), ParsedToken::Paren];
        assert_eq!(decompile(&tokens, &ctx), "(42)");
    }

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(255), "IV");
    }

    #[test]
    fn test_empty_tokens() {
        let ctx = empty_ctx();
        assert_eq!(decompile(&[], &ctx), "");
    }

    #[test]
    fn test_volatile_ignored() {
        let ctx = empty_ctx();
        let tokens = vec![
            ParsedToken::AttrVolatile,
            ParsedToken::Func { func_idx: 74 },
        ];
        assert_eq!(decompile(&tokens, &ctx), "NOW()");
    }

    #[test]
    fn test_left_associativity_subtraction() {
        let ctx = empty_ctx();
        let tokens = vec![
            ParsedToken::Int(1),
            ParsedToken::Int(2),
            ParsedToken::Sub,
            ParsedToken::Int(3),
            ParsedToken::Sub,
        ];
        assert_eq!(decompile(&tokens, &ctx), "1-2-3");
    }

    #[test]
    fn test_right_operand_needs_parens() {
        let ctx = empty_ctx();
        let tokens = vec![
            ParsedToken::Int(1),
            ParsedToken::Int(2),
            ParsedToken::Int(3),
            ParsedToken::Add,
            ParsedToken::Sub,
        ];
        assert_eq!(decompile(&tokens, &ctx), "1-(2+3)");
    }

    #[test]
    fn test_3d_ref_with_self_ref_supbook() {
        let ctx = ctx_with_self_ref(vec!["Sheet1".to_string(), "Sheet2".to_string()]);
        let tokens = vec![ParsedToken::Ref3d {
            extern_sheet_idx: 1,
            row: 0,
            col: 0,
            row_relative: false,
            col_relative: false,
        }];
        assert_eq!(decompile(&tokens, &ctx), "Sheet2!$A$1");
    }

    #[test]
    fn test_3d_ref_fallback_no_externsheet() {
        // No EXTERNSHEET data — falls back to direct sheet_names lookup
        let ctx = ctx_with_sheets(vec!["Sheet1".to_string(), "Sheet2".to_string()]);
        let tokens = vec![ParsedToken::Ref3d {
            extern_sheet_idx: 1,
            row: 0,
            col: 0,
            row_relative: false,
            col_relative: false,
        }];
        assert_eq!(decompile(&tokens, &ctx), "Sheet2!$A$1");
    }

    #[test]
    fn test_3d_ref_quoted_sheet() {
        let ctx = ctx_with_self_ref(vec!["My Sheet".to_string()]);
        let tokens = vec![ParsedToken::Ref3d {
            extern_sheet_idx: 0,
            row: 4,
            col: 1,
            row_relative: false,
            col_relative: false,
        }];
        assert_eq!(decompile(&tokens, &ctx), "'My Sheet'!$B$5");
    }

    #[test]
    fn test_3d_area_multi_sheet_range() {
        // EXTERNSHEET entry with first_sheet=0, last_sheet=2 → Sheet1:Sheet3
        let ctx = FormulaContext {
            sheet_names: vec![
                "Sheet1".to_string(),
                "Sheet2".to_string(),
                "Sheet3".to_string(),
            ],
            supbooks: vec![SupBook::SelfRef { sheet_count: 3 }],
            extern_sheet: vec![ExternSheetEntry {
                sup_book_idx: 0,
                first_sheet: 0,
                last_sheet: 2,
            }],
            names: Vec::new(),
        };
        let tokens = vec![ParsedToken::Area3d {
            extern_sheet_idx: 0,
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 0,
            first_row_rel: false,
            first_col_rel: false,
            last_row_rel: false,
            last_col_rel: false,
        }];
        assert_eq!(decompile(&tokens, &ctx), "Sheet1:Sheet3!$A$1:$A$10");
    }

    #[test]
    fn test_name_lookup() {
        let ctx = FormulaContext {
            sheet_names: vec!["Sheet1".to_string()],
            supbooks: vec![],
            extern_sheet: vec![],
            names: vec![NameRecord {
                name: "MyRange".to_string(),
                sheet_idx: 0,
                is_builtin: false,
            }],
        };
        // tName with name_idx=1 (1-based) → "MyRange"
        let tokens = vec![ParsedToken::Name { name_idx: 1 }];
        assert_eq!(decompile(&tokens, &ctx), "MyRange");
    }

    #[test]
    fn test_name_lookup_unknown() {
        let ctx = empty_ctx();
        // No names defined — should fall back to placeholder
        let tokens = vec![ParsedToken::Name { name_idx: 5 }];
        assert_eq!(decompile(&tokens, &ctx), "_name5");
    }

    #[test]
    fn test_absolute_ref() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Ref {
            row: 0,
            col: 0,
            row_relative: false,
            col_relative: false,
        }];
        assert_eq!(decompile(&tokens, &ctx), "$A$1");
    }

    #[test]
    fn test_mixed_ref_absolute_col() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Ref {
            row: 4,
            col: 2,
            row_relative: true,
            col_relative: false,
        }];
        assert_eq!(decompile(&tokens, &ctx), "$C5");
    }

    #[test]
    fn test_mixed_ref_absolute_row() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Ref {
            row: 4,
            col: 2,
            row_relative: false,
            col_relative: true,
        }];
        assert_eq!(decompile(&tokens, &ctx), "C$5");
    }

    #[test]
    fn test_mixed_area_ref() {
        let ctx = empty_ctx();
        let tokens = vec![ParsedToken::Area {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 2,
            first_row_rel: false,
            first_col_rel: false,
            last_row_rel: true,
            last_col_rel: true,
        }];
        assert_eq!(decompile(&tokens, &ctx), "$A$1:C10");
    }

    #[test]
    fn test_namex_self_ref() {
        let ctx = FormulaContext {
            sheet_names: vec!["Sheet1".to_string()],
            supbooks: vec![SupBook::SelfRef { sheet_count: 1 }],
            extern_sheet: vec![ExternSheetEntry {
                sup_book_idx: 0,
                first_sheet: 0xFFFE,
                last_sheet: 0xFFFE,
            }],
            names: vec![NameRecord {
                name: "TaxRate".to_string(),
                sheet_idx: 0,
                is_builtin: false,
            }],
        };
        let tokens = vec![ParsedToken::NameX {
            extern_sheet_idx: 0,
            name_idx: 1,
        }];
        assert_eq!(decompile(&tokens, &ctx), "TaxRate");
    }
}
