use std::collections::HashMap;

use duke_sheets_core::{CellAddress, CellError};
use duke_sheets_formula::ast::{
    BinaryOperator, CellReference, FormulaExpr, RangeReference, UnaryOperator,
};
use duke_sheets_formula::decompile::function_table::{
    expr_calls_volatile_function, function_arg_class, function_index, function_is_biff8_addin,
    function_is_fixed_arity, function_returns_reference, name_body_operand_class, OperandClass,
};
use duke_sheets_formula::parse_formula;

use super::ptg;

pub(crate) struct CompileContext {
    pub sheet_names: Vec<String>,
    /// Maps uppercase function name (e.g. "IFS") to 1-based name index
    /// for _xlfn.* functions not in the standard FTAB.
    pub xlfn_names: HashMap<String, u32>,
    /// User-defined names in workbook emission order. NameRef lookups
    /// match case-insensitively and emit a 1-based index into this
    /// list as a PtgName payload.
    pub defined_names: Vec<String>,
}

pub(crate) struct CompiledFormula {
    pub rgce: Vec<u8>,
    pub rgcb: Vec<u8>,
}

/// Compile a cell formula. The top-level expression is value-class.
pub(crate) fn compile_formula(text: &str, ctx: &CompileContext) -> Result<CompiledFormula, String> {
    compile_with_top_class(text, ctx, false)
}

/// Compile a defined-name / built-in-name body (the `refers_to` formula).
///
/// Unlike a cell formula, a name body that is a bare reference or range must
/// be reference-class: a value-class range would make Excel apply implicit
/// intersection when the name is used (collapsing `Data!A1:A3` to a single
/// cell), so e.g. `SUM(Numbers)` would sum one cell instead of the range.
/// Mirrors the XLS writer's `name_body_operand_class`.
pub(crate) fn compile_name_body(
    text: &str,
    ctx: &CompileContext,
) -> Result<CompiledFormula, String> {
    compile_with_top_class(text, ctx, true)
}

fn compile_with_top_class(
    text: &str,
    ctx: &CompileContext,
    name_body: bool,
) -> Result<CompiledFormula, String> {
    let formula = if text.starts_with('=') {
        text
    } else {
        &format!("={text}")
    };
    let expr = parse_formula(formula).map_err(|e| format!("{e}"))?;
    let top_class = if name_body {
        name_body_operand_class(&expr)
    } else {
        OperandClass::V
    };
    let mut rgce = Vec::new();
    let mut rgcb = Vec::new();
    // A formula calling any volatile function (NOW, RAND, OFFSET, INDIRECT,
    // ...) is prefixed with PtgAttrVolatile so Excel recalculates it every
    // change. [MS-XLS] §2.5.198.42; BIFF12 uses the identical token shape.
    if expr_calls_volatile_function(&expr) {
        rgce.push(ptg::PTG_ATTR);
        rgce.push(ptg::ATTR_VOLATILE);
        rgce.extend_from_slice(&0u16.to_le_bytes());
    }
    emit_expr(&expr, ctx, &mut rgce, &mut rgcb, top_class)?;
    Ok(CompiledFormula { rgce, rgcb })
}

fn emit_expr(
    expr: &FormulaExpr,
    ctx: &CompileContext,
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
    class: OperandClass,
) -> Result<(), String> {
    match expr {
        FormulaExpr::Number(n) => emit_number(*n, out),
        FormulaExpr::String(s) => emit_string(s, out),
        FormulaExpr::Boolean(b) => emit_bool(*b, out),
        FormulaExpr::Error(e) => emit_error(e, out),
        FormulaExpr::Empty => emit_miss_arg(out),

        FormulaExpr::CellRef(cell_ref) => emit_cell_ref(cell_ref, ctx, out, class),
        FormulaExpr::RangeRef(range_ref) => emit_range_ref(range_ref, ctx, out, class),

        FormulaExpr::BinaryOp { op, left, right } => {
            // Range/union/intersect operate on references (R-class children);
            // arithmetic / comparison / concat always take values regardless
            // of context. Matches [MS-XLS] §2.5.198 value-vs-reference rules.
            let child_class = match op {
                BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect => {
                    OperandClass::R
                }
                _ => OperandClass::V,
            };
            emit_expr(left, ctx, out, extra, child_class)?;
            emit_expr(right, ctx, out, extra, child_class)?;
            emit_binary_op(*op, out)
        }
        FormulaExpr::UnaryOp { op, operand } => {
            emit_expr(operand, ctx, out, extra, class)?;
            emit_unary_op(*op, out)
        }

        FormulaExpr::Function { name, args } => emit_function(name, args, ctx, out, extra, class),

        FormulaExpr::Array(rows) => emit_array(rows, out, extra),

        FormulaExpr::NameRef(name) => {
            // Resolve to a 1-based index into the BrtName record stream.
            // _xlfn.* records are written before user names, so user
            // names start at ilbl xlfn_count + 1. Match
            // case-insensitively per Excel's convention.
            let upper = name.to_ascii_uppercase();
            let idx_0based = ctx
                .defined_names
                .iter()
                .position(|n| n.to_ascii_uppercase() == upper);
            match idx_0based {
                Some(i) => {
                    // PtgName takes the class of its position (R when used as
                    // a reference operand, V at value positions).
                    let ilbl = ctx.xlfn_names.len() as u32 + i as u32 + 1;
                    out.push(class_ptg(ptg::PTG_NAME, class));
                    out.extend_from_slice(&ilbl.to_le_bytes());
                    Ok(())
                }
                None => {
                    log::warn!("unknown named range '{}', emitting #NAME?", name);
                    emit_error(&CellError::Name, out)
                }
            }
        }
        FormulaExpr::StructuredRef(_) => {
            log::warn!("structured reference compilation not supported, emitting #REF!");
            emit_error(&CellError::Ref, out)
        }
        FormulaExpr::ExternalRef(_) => {
            log::warn!("external workbook reference compilation not supported, emitting #REF!");
            emit_error(&CellError::Ref, out)
        }
    }
}

/// Apply an [`OperandClass`] to an R-class base ptg byte. R keeps the stored
/// (R-class) constant; V converts via [`ptg::v_class`].
fn class_ptg(r_base: u8, class: OperandClass) -> u8 {
    match class {
        OperandClass::R => r_base,
        OperandClass::V => ptg::v_class(r_base),
    }
}

/// Effective class for a function's PtgFunc/PtgFuncVar token: reference-class
/// functions (IF, CHOOSE, OFFSET, INDIRECT, INDEX) take the surrounding
/// context class; pure value functions are always V. [MS-XLS] §2.5.198.103.
fn func_token_class(iftab: u16, context: OperandClass) -> OperandClass {
    if function_returns_reference(iftab) {
        context
    } else {
        OperandClass::V
    }
}

fn emit_number(n: f64, out: &mut Vec<u8>) -> Result<(), String> {
    if n >= 0.0 && n <= 65535.0 && n == n.floor() && !n.is_nan() {
        out.push(ptg::PTG_INT);
        out.extend_from_slice(&(n as u16).to_le_bytes());
    } else {
        out.push(ptg::PTG_NUM);
        out.extend_from_slice(&n.to_le_bytes());
    }
    Ok(())
}

fn emit_string(s: &str, out: &mut Vec<u8>) -> Result<(), String> {
    let utf16: Vec<u16> = s.encode_utf16().collect();
    out.push(ptg::PTG_STR);
    out.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
    for code_unit in &utf16 {
        out.extend_from_slice(&code_unit.to_le_bytes());
    }
    Ok(())
}

fn emit_bool(b: bool, out: &mut Vec<u8>) -> Result<(), String> {
    out.push(ptg::PTG_BOOL);
    out.push(if b { 0x01 } else { 0x00 });
    Ok(())
}

fn emit_error(e: &CellError, out: &mut Vec<u8>) -> Result<(), String> {
    out.push(ptg::PTG_ERR);
    out.push(error_byte(e));
    Ok(())
}

fn emit_miss_arg(out: &mut Vec<u8>) -> Result<(), String> {
    out.push(ptg::PTG_MISS_ARG);
    Ok(())
}

fn emit_array(
    rows: &[Vec<FormulaExpr>],
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
) -> Result<(), String> {
    // PtgArray is A-class (0x60): array constants always carry the array
    // data type. 1 ptg byte + 14 reserved bytes in the rgce. Verified
    // against native Excel XLSB authoring.
    out.push(ptg::a_class(ptg::PTG_ARRAY));
    out.extend_from_slice(&[0u8; 14]);

    let nr = rows.len();
    let nc = rows.first().map_or(0, |r| r.len());

    // BIFF12 rgcb: rows(u32) + cols(u32) (actual counts, rows first), then
    // row-major elements. Verified against native authoring: {1,2,3} (1x3)
    // emits 01 00 00 00 03 00 00 00.
    extra.extend_from_slice(&(nr as u32).to_le_bytes());
    extra.extend_from_slice(&(nc as u32).to_le_bytes());

    for row in rows {
        for expr in row {
            match expr {
                FormulaExpr::Number(n) => {
                    extra.push(0x00);
                    extra.extend_from_slice(&n.to_le_bytes());
                }
                // SerAr element bodies per [MS-XLSB] (cross-checked
                // against LO importArrayToken): string carries a
                // 16-bit cch, bool is a single byte with no padding,
                // error is a single byte plus 3 reserved bytes.
                FormulaExpr::String(s) => {
                    extra.push(0x01);
                    let utf16: Vec<u16> = s.encode_utf16().collect();
                    extra.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
                    for cu in &utf16 {
                        extra.extend_from_slice(&cu.to_le_bytes());
                    }
                }
                FormulaExpr::Boolean(b) => {
                    extra.push(0x02);
                    extra.push(if *b { 1 } else { 0 });
                }
                FormulaExpr::Error(e) => {
                    extra.push(0x04);
                    extra.push(error_byte(e));
                    extra.extend_from_slice(&[0u8; 3]);
                }
                _ => {
                    extra.push(0x10);
                    extra.extend_from_slice(&[0u8; 8]);
                }
            }
        }
    }
    Ok(())
}

fn error_byte(e: &CellError) -> u8 {
    match e {
        CellError::Null => 0x00,
        CellError::Div0 => 0x07,
        CellError::Value => 0x0F,
        CellError::Ref => 0x17,
        CellError::Name => 0x1D,
        CellError::Num => 0x24,
        CellError::Na => 0x2A,
        CellError::GettingData => 0x2B,
        _ => 0x0F,
    }
}

fn encode_col_word(addr: &CellAddress) -> u16 {
    let mut w = addr.col;
    if !addr.row_absolute {
        w |= 0x4000;
    }
    if !addr.col_absolute {
        w |= 0x8000;
    }
    w
}

fn emit_cell_ref(
    cell_ref: &CellReference,
    ctx: &CompileContext,
    out: &mut Vec<u8>,
    class: OperandClass,
) -> Result<(), String> {
    match &cell_ref.sheet {
        None => {
            // tRef — class of its position (V for value use, R inside an
            // aggregator like SUM(A1,A2) which iterates its operands).
            out.push(class_ptg(ptg::PTG_REF, class));
            out.extend_from_slice(&cell_ref.address.row.to_le_bytes());
            out.extend_from_slice(&encode_col_word(&cell_ref.address).to_le_bytes());
        }
        Some(sheet_name) => {
            let sheet_idx = resolve_sheet_index(sheet_name, ctx)?;
            // tRef3d
            out.push(class_ptg(ptg::PTG_REF_3D, class));
            out.extend_from_slice(&(sheet_idx as u16).to_le_bytes());
            out.extend_from_slice(&cell_ref.address.row.to_le_bytes());
            out.extend_from_slice(&encode_col_word(&cell_ref.address).to_le_bytes());
        }
    }
    Ok(())
}

fn emit_range_ref(
    range_ref: &RangeReference,
    ctx: &CompileContext,
    out: &mut Vec<u8>,
    class: OperandClass,
) -> Result<(), String> {
    let start = &range_ref.range.start;
    let end = &range_ref.range.end;

    match &range_ref.sheet {
        None => {
            // tArea — class of its position. R-class is the common case
            // (aggregators, intersection/union/range operands); V-class
            // causes Excel to collapse the range to its last cell during
            // SUM-style evaluation, so a value-context bare range is V.
            out.push(class_ptg(ptg::PTG_AREA, class));
            out.extend_from_slice(&start.row.to_le_bytes());
            out.extend_from_slice(&end.row.to_le_bytes());
            out.extend_from_slice(&encode_col_word(start).to_le_bytes());
            out.extend_from_slice(&encode_col_word(end).to_le_bytes());
        }
        Some(sheet_name) => {
            let sheet_idx = resolve_sheet_index(sheet_name, ctx)?;
            // tArea3d — cross-sheet ranges are R-class in the common case.
            out.push(class_ptg(ptg::PTG_AREA_3D, class));
            out.extend_from_slice(&(sheet_idx as u16).to_le_bytes());
            out.extend_from_slice(&start.row.to_le_bytes());
            out.extend_from_slice(&end.row.to_le_bytes());
            out.extend_from_slice(&encode_col_word(start).to_le_bytes());
            out.extend_from_slice(&encode_col_word(end).to_le_bytes());
        }
    }
    Ok(())
}

fn resolve_sheet_index(sheet_name: &str, ctx: &CompileContext) -> Result<usize, String> {
    ctx.sheet_names
        .iter()
        .position(|n| n.eq_ignore_ascii_case(sheet_name))
        .ok_or_else(|| format!("unknown sheet '{sheet_name}'"))
}

fn emit_binary_op(op: BinaryOperator, out: &mut Vec<u8>) -> Result<(), String> {
    let byte = match op {
        BinaryOperator::Add => ptg::PTG_ADD,
        BinaryOperator::Subtract => ptg::PTG_SUB,
        BinaryOperator::Multiply => ptg::PTG_MUL,
        BinaryOperator::Divide => ptg::PTG_DIV,
        BinaryOperator::Power => ptg::PTG_POWER,
        BinaryOperator::Concat => ptg::PTG_CONCAT,
        BinaryOperator::LessThan => ptg::PTG_LT,
        BinaryOperator::LessEqual => ptg::PTG_LE,
        BinaryOperator::Equal => ptg::PTG_EQ,
        BinaryOperator::GreaterEqual => ptg::PTG_GE,
        BinaryOperator::GreaterThan => ptg::PTG_GT,
        BinaryOperator::NotEqual => ptg::PTG_NE,
        BinaryOperator::Range => ptg::PTG_RANGE,
        BinaryOperator::Union => ptg::PTG_LIST,
        BinaryOperator::Intersect => ptg::PTG_ISECT,
    };
    out.push(byte);
    Ok(())
}

fn emit_unary_op(op: UnaryOperator, out: &mut Vec<u8>) -> Result<(), String> {
    let byte = match op {
        UnaryOperator::Plus => ptg::PTG_UPLUS,
        UnaryOperator::Negate => ptg::PTG_UMINUS,
        UnaryOperator::Percent => ptg::PTG_PERCENT,
        UnaryOperator::Paren => ptg::PTG_PAREN,
        UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange => {
            // These are dynamic array operators with no classical Ptg encoding.
            // Skip silently - the formula still evaluates via the cached value.
            return Ok(());
        }
    };
    out.push(byte);
    Ok(())
}

fn emit_function(
    name: &str,
    args: &[FormulaExpr],
    ctx: &CompileContext,
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
    class: OperandClass,
) -> Result<(), String> {
    let lookup_name = name
        .strip_prefix("_xlfn.")
        .or_else(|| name.strip_prefix("_XLFN."))
        .unwrap_or(name);

    if let Some(func_idx) = function_index(lookup_name) {
        // Single-arg SUM → optimized PtgAttrSum form.
        if func_idx == 4 && args.len() == 1 && !matches!(args[0], FormulaExpr::Empty) {
            let arg_class = function_arg_class(4, 0);
            emit_expr(&args[0], ctx, out, extra, arg_class)?;
            out.push(ptg::PTG_ATTR);
            out.push(ptg::ATTR_SUM);
            out.extend_from_slice(&0u16.to_le_bytes());
            return Ok(());
        }
        // IF short-circuit (PtgAttrIf / PtgAttrGoto), 2- and 3-arg forms.
        if func_idx == 1 && (args.len() == 2 || args.len() == 3) && emit_optimized_if(args, ctx, out, extra, class)? {
            return Ok(());
        }
        // CHOOSE jump table (PtgAttrChoose).
        if func_idx == 100 && args.len() >= 2 && emit_optimized_choose(args, ctx, out, extra, class)? {
            return Ok(());
        }

        if args.len() > u8::MAX as usize {
            return Err(format!("function '{name}' has too many arguments"));
        }
        // Analysis-ToolPak functions (Ftab 384..=476) take by-reference
        // (R-class) arguments in Excel's native XLSB emission — the same rule
        // as the BIFF8 add-in form. Other functions use their per-position
        // metadata class.
        let addin = function_is_biff8_addin(func_idx);
        for (i, arg) in args.iter().enumerate() {
            if matches!(arg, FormulaExpr::Empty) {
                emit_miss_arg(out)?;
            } else {
                let arg_class = if addin {
                    OperandClass::R
                } else {
                    function_arg_class(func_idx, i)
                };
                emit_expr(arg, ctx, out, extra, arg_class)?;
            }
        }

        let tok_class = func_token_class(func_idx, class);
        if function_is_fixed_arity(func_idx, args.len()) {
            out.push(class_ptg(ptg::PTG_FUNC, tok_class));
            out.extend_from_slice(&func_idx.to_le_bytes());
        } else {
            out.push(class_ptg(ptg::PTG_FUNC_VAR, tok_class));
            out.push(args.len() as u8);
            out.extend_from_slice(&func_idx.to_le_bytes());
        }
    } else if let Some(&name_idx) = ctx.xlfn_names.get(&lookup_name.to_ascii_uppercase()) {
        // tName ref to the _xlfn.* BrtName record is pushed BEFORE the
        // args: PtgFuncVar with iftab 0xFF treats the bottom-most of
        // its argc operands as the function name.
        out.push(ptg::v_class(ptg::PTG_NAME));
        out.extend_from_slice(&name_idx.to_le_bytes());
        for arg in args {
            emit_expr(arg, ctx, out, extra, OperandClass::V)?;
        }
        // tFuncVar (V-class) with func_idx=0xFF: argc includes the tName ref
        out.push(ptg::v_class(ptg::PTG_FUNC_VAR));
        out.push((args.len() + 1) as u8);
        out.extend_from_slice(&0x00FFu16.to_le_bytes());
    } else {
        return Err(format!("unknown function '{name}'"));
    }

    Ok(())
}

/// Emit IF with the MS-XLS PtgAttrIf / PtgAttrGoto short-circuit chain.
/// BIFF12 uses the identical token shape as BIFF8; offsets are computed from
/// the actual (wider) BIFF12 branch byte sizes. Returns `Ok(false)` on u16
/// offset overflow so the caller falls back to plain PtgFuncVar emission.
/// Mirrors the XLS writer's `emit_optimized_if`.
fn emit_optimized_if(
    args: &[FormulaExpr],
    ctx: &CompileContext,
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
    class: OperandClass,
) -> Result<bool, String> {
    if args.iter().any(|a| matches!(a, FormulaExpr::Empty)) {
        return Ok(false);
    }
    let cond = &args[0];
    let t_branch = &args[1];
    let f_branch = args.get(2);
    let argc = args.len() as u8;

    // Compile every part into scratch buffers first so an overflow bail
    // leaves `out` untouched for the caller's fallback.
    let mut cond_bytes = Vec::new();
    emit_expr(cond, ctx, &mut cond_bytes, extra, function_arg_class(1, 0))?;
    let mut t_bytes = Vec::new();
    emit_expr(t_branch, ctx, &mut t_bytes, extra, function_arg_class(1, 1))?;
    let mut f_bytes = Vec::new();
    if let Some(f) = f_branch {
        emit_expr(f, ctx, &mut f_bytes, extra, function_arg_class(1, 2))?;
    }

    let attr_if_offset = t_bytes.len() + 4;
    if attr_if_offset > u16::MAX as usize {
        return Ok(false);
    }
    let skip_after_t = if f_branch.is_some() {
        let s = f_bytes.len() + 7;
        if s > u16::MAX as usize {
            return Ok(false);
        }
        Some(s as u16)
    } else {
        None
    };

    out.extend_from_slice(&cond_bytes);
    out.push(ptg::PTG_ATTR);
    out.push(ptg::ATTR_IF);
    out.extend_from_slice(&(attr_if_offset as u16).to_le_bytes());
    out.extend_from_slice(&t_bytes);
    if let Some(skip) = skip_after_t {
        out.push(ptg::PTG_ATTR);
        out.push(ptg::ATTR_SKIP);
        out.extend_from_slice(&skip.to_le_bytes());
        out.extend_from_slice(&f_bytes);
    }
    out.push(ptg::PTG_ATTR);
    out.push(ptg::ATTR_SKIP);
    out.extend_from_slice(&3u16.to_le_bytes());
    // IF is reference-class; token takes the context class.
    out.push(class_ptg(ptg::PTG_FUNC_VAR, func_token_class(1, class)));
    out.push(argc);
    out.extend_from_slice(&1u16.to_le_bytes());
    Ok(true)
}

/// Emit CHOOSE with the MS-XLS PtgAttrChoose jump table. BIFF12 uses the
/// identical token shape as BIFF8. Returns `Ok(false)` on u16 offset overflow.
/// Mirrors the XLS writer's `emit_optimized_choose`.
fn emit_optimized_choose(
    args: &[FormulaExpr],
    ctx: &CompileContext,
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
    class: OperandClass,
) -> Result<bool, String> {
    if args.len() < 2 || args.iter().any(|a| matches!(a, FormulaExpr::Empty)) {
        return Ok(false);
    }
    let selector = &args[0];
    let choices = &args[1..];
    let nc = choices.len();
    if nc > u16::MAX as usize {
        return Ok(false);
    }
    let argc = args.len() as u8;

    let mut selector_bytes = Vec::new();
    emit_expr(selector, ctx, &mut selector_bytes, extra, function_arg_class(100, 0))?;

    let mut choice_bytes: Vec<Vec<u8>> = Vec::with_capacity(nc);
    for (i, c) in choices.iter().enumerate() {
        let mut buf = Vec::new();
        emit_expr(c, ctx, &mut buf, extra, function_arg_class(100, i + 1))?;
        choice_bytes.push(buf);
    }

    let table_size = (nc + 1) * 2;
    let mut offsets: Vec<u16> = Vec::with_capacity(nc + 1);
    let mut running = table_size;
    for choice in &choice_bytes {
        if running > u16::MAX as usize {
            return Ok(false);
        }
        offsets.push(running as u16);
        running += choice.len() + 4; // PtgAttrSkip after the choice
    }
    if running > u16::MAX as usize {
        return Ok(false);
    }
    offsets.push(running as u16);

    let mut skip_offsets: Vec<u16> = vec![0; nc];
    let mut remaining_after = 0usize;
    for k in (0..nc).rev() {
        if k + 1 == nc {
            skip_offsets[k] = 3;
        } else {
            let nxt = choice_bytes[k + 1].len() + 4 + remaining_after;
            if nxt + 3 > u16::MAX as usize {
                return Ok(false);
            }
            skip_offsets[k] = (nxt + 3) as u16;
            remaining_after = nxt;
        }
    }

    out.extend_from_slice(&selector_bytes);
    out.push(ptg::PTG_ATTR);
    out.push(ptg::ATTR_CHOOSE);
    out.extend_from_slice(&(nc as u16).to_le_bytes());
    for off in &offsets {
        out.extend_from_slice(&off.to_le_bytes());
    }
    for (k, choice) in choice_bytes.iter().enumerate() {
        out.extend_from_slice(choice);
        out.push(ptg::PTG_ATTR);
        out.push(ptg::ATTR_SKIP);
        out.extend_from_slice(&skip_offsets[k].to_le_bytes());
    }
    out.push(class_ptg(ptg::PTG_FUNC_VAR, func_token_class(100, class)));
    out.push(argc);
    out.extend_from_slice(&100u16.to_le_bytes());
    Ok(true)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::biff12::token_parser::parse_tokens_with_extra;
    use duke_sheets_formula::decompile::ParsedToken;

    fn ctx() -> CompileContext {
        CompileContext {
            sheet_names: vec!["Sheet1".to_string(), "Sheet2".to_string()],
            xlfn_names: HashMap::new(),
            defined_names: Vec::new(),
        }
    }

    fn compile(text: &str) -> CompiledFormula {
        compile_formula(text, &ctx()).unwrap()
    }

    fn compile_and_parse(text: &str) -> Vec<ParsedToken> {
        let compiled = compile(text);
        parse_tokens_with_extra(&compiled.rgce, &compiled.rgcb)
    }

    #[test]
    fn int_add() {
        let tokens = compile_and_parse("=1+2");
        assert_eq!(
            tokens,
            vec![ParsedToken::Int(1), ParsedToken::Int(2), ParsedToken::Add]
        );
    }

    #[test]
    fn float_number() {
        let tokens = compile_and_parse("=3.14");
        assert_eq!(tokens, vec![ParsedToken::Num(3.14)]);
    }

    #[test]
    fn negative_number() {
        let tokens = compile_and_parse("=-5");
        assert_eq!(tokens, vec![ParsedToken::Int(5), ParsedToken::Uminus]);
    }

    #[test]
    fn string_literal() {
        let tokens = compile_and_parse("=\"hello\"");
        assert_eq!(tokens, vec![ParsedToken::Str("hello".to_string())]);
    }

    #[test]
    fn boolean_values() {
        assert_eq!(compile_and_parse("=TRUE"), vec![ParsedToken::Bool(true)]);
        assert_eq!(compile_and_parse("=FALSE"), vec![ParsedToken::Bool(false)]);
    }

    #[test]
    fn error_values() {
        assert_eq!(compile_and_parse("=#N/A"), vec![ParsedToken::Err(0x2A)]);
        assert_eq!(compile_and_parse("=#VALUE!"), vec![ParsedToken::Err(0x0F)]);
    }

    #[test]
    fn cell_ref_absolute() {
        let tokens = compile_and_parse("=$A$1");
        assert_eq!(
            tokens,
            vec![ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: false,
                col_relative: false,
            }]
        );
    }

    #[test]
    fn cell_ref_relative() {
        let tokens = compile_and_parse("=A1");
        assert_eq!(
            tokens,
            vec![ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            }]
        );
    }

    #[test]
    fn range_ref() {
        let tokens = compile_and_parse("=A1:B10");
        assert_eq!(
            tokens,
            vec![ParsedToken::Area {
                first_row: 0,
                last_row: 9,
                first_col: 0,
                last_col: 1,
                first_row_rel: true,
                first_col_rel: true,
                last_row_rel: true,
                last_col_rel: true,
            }]
        );
    }

    #[test]
    fn sum_attr_optimization() {
        let tokens = compile_and_parse("=SUM(A1:A10)");
        assert_eq!(tokens.len(), 2);
        assert!(matches!(tokens[0], ParsedToken::Area { .. }));
        assert_eq!(tokens[1], ParsedToken::AttrSum);
    }

    #[test]
    fn if_function() {
        // IF now uses the MS-XLS PtgAttrIf / PtgAttrGoto short-circuit chain,
        // matching Excel's native XLSB emission. Branches are Bool literals
        // (2 bytes each): attr_if_offset = 2+4 = 6, skip_after_t = 2+7 = 9,
        // trailing skip = 3.
        let tokens = compile_and_parse("=IF(A1>0,TRUE,FALSE)");
        assert_eq!(tokens.len(), 9, "tokens={tokens:?}");
        assert!(matches!(tokens[0], ParsedToken::Ref { .. }));
        assert_eq!(tokens[1], ParsedToken::Int(0));
        assert_eq!(tokens[2], ParsedToken::Gt);
        assert_eq!(tokens[3], ParsedToken::AttrIf { offset: 6 });
        assert_eq!(tokens[4], ParsedToken::Bool(true));
        assert_eq!(tokens[5], ParsedToken::AttrSkip { offset: 9 });
        assert_eq!(tokens[6], ParsedToken::Bool(false));
        assert_eq!(tokens[7], ParsedToken::AttrSkip { offset: 3 });
        assert_eq!(tokens[8], ParsedToken::FuncVar { argc: 3, func_idx: 1 });
    }

    #[test]
    fn cross_sheet_ref() {
        let tokens = compile_and_parse("=Sheet2!A1");
        assert_eq!(
            tokens,
            vec![ParsedToken::Ref3d {
                extern_sheet_idx: 1,
                row: 0,
                col: 0,
                row_relative: true,
                col_relative: true,
            }]
        );
    }

    #[test]
    fn cross_sheet_range() {
        let tokens = compile_and_parse("=Sheet2!A1:B5");
        assert_eq!(
            tokens,
            vec![ParsedToken::Area3d {
                extern_sheet_idx: 1,
                first_row: 0,
                last_row: 4,
                first_col: 0,
                last_col: 1,
                first_row_rel: true,
                first_col_rel: true,
                last_row_rel: true,
                last_col_rel: true,
            }]
        );
    }

    #[test]
    fn comparison_operators() {
        let tokens = compile_and_parse("=1<2");
        assert_eq!(
            tokens,
            vec![ParsedToken::Int(1), ParsedToken::Int(2), ParsedToken::Lt]
        );
    }

    #[test]
    fn concat_operator() {
        let tokens = compile_and_parse("=\"a\"&\"b\"");
        assert_eq!(
            tokens,
            vec![
                ParsedToken::Str("a".to_string()),
                ParsedToken::Str("b".to_string()),
                ParsedToken::Concat,
            ]
        );
    }

    #[test]
    fn empty_arg() {
        let tokens = compile_and_parse("=IF(TRUE,,0)");
        assert!(tokens.contains(&ParsedToken::MissArg));
    }

    #[test]
    fn fixed_arg_func() {
        let tokens = compile_and_parse("=LEN(\"abc\")");
        let last = tokens.last().unwrap();
        assert_eq!(*last, ParsedToken::Func { func_idx: 32 });
    }

    #[test]
    fn unknown_sheet_errors() {
        let result = compile_formula("=NoSuchSheet!A1", &ctx());
        assert!(result.is_err());
    }
}
