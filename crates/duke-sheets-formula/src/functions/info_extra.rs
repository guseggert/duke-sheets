use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::{CellAddress, CellError};

pub fn fn_iserr(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array { .. }) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(matches!(
        v,
        FormulaValue::Error(e) if *e != CellError::Na
    )))
}

pub fn fn_iseven(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array { .. }) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let n = match v {
        FormulaValue::Number(n) => *n,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    Ok(FormulaValue::Boolean((n.trunc() as i64) % 2 == 0))
}

pub fn fn_isodd(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array { .. }) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let n = match v {
        FormulaValue::Number(n) => *n,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    Ok(FormulaValue::Boolean((n.trunc() as i64) % 2 != 0))
}

pub fn fn_islogical(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array { .. }) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::Boolean(_))))
}

pub fn fn_isnontext(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array { .. }) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(!matches!(v, FormulaValue::String(_))))
}

pub fn fn_isref(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array { .. }) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(false))
}

pub fn fn_error_type(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    let n = match v {
        FormulaValue::Error(CellError::Null) => Some(1.0),
        FormulaValue::Error(CellError::Div0) => Some(2.0),
        FormulaValue::Error(CellError::Value) => Some(3.0),
        FormulaValue::Error(CellError::Ref) => Some(4.0),
        FormulaValue::Error(CellError::Name) => Some(5.0),
        FormulaValue::Error(CellError::Num) => Some(6.0),
        FormulaValue::Error(CellError::Na) => Some(7.0),
        FormulaValue::Error(CellError::GettingData) => Some(8.0),
        _ => None,
    };

    match n {
        Some(n) => Ok(FormulaValue::Number(n)),
        None => Ok(FormulaValue::Error(CellError::Na)),
    }
}

pub fn fn_type(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    let n = match v {
        FormulaValue::Number(_) => 1.0,
        FormulaValue::String(_) => 2.0,
        FormulaValue::Boolean(_) => 4.0,
        FormulaValue::Error(_) => 16.0,
        FormulaValue::Array { .. } => 64.0,
        FormulaValue::Empty => 128.0,
    };

    Ok(FormulaValue::Number(n))
}

pub fn fn_cell(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Extract info_type from first argument
    let info_type = match args.first() {
        Some(FormulaValue::String(s)) => s.to_ascii_lowercase(),
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(_) => return Ok(FormulaValue::Error(CellError::Value)),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Get reference value (second arg) or Empty if omitted
    let ref_val = args.get(1).cloned().unwrap_or(FormulaValue::Empty);

    // Propagate errors from the reference argument
    if let FormulaValue::Error(e) = &ref_val {
        return Ok(FormulaValue::Error(*e));
    }

    match info_type.as_str() {
        "address" => {
            // Return absolute cell address as text, e.g. "$A$1"
            let col_letters = CellAddress::column_to_letters(ctx.current_col);
            let row_num = ctx.current_row + 1; // 1-based
            Ok(FormulaValue::String(format!(
                "${}${}",
                col_letters, row_num
            )))
        }
        "col" => {
            // Column number (1-based)
            Ok(FormulaValue::Number((ctx.current_col as f64) + 1.0))
        }
        "row" => {
            // Row number (1-based)
            Ok(FormulaValue::Number((ctx.current_row as f64) + 1.0))
        }
        "contents" => {
            // Return the cell value; Empty becomes 0 like Excel
            match ref_val {
                FormulaValue::Empty => Ok(FormulaValue::Number(0.0)),
                other => Ok(other),
            }
        }
        "type" => {
            // "b" for blank, "l" for label (text), "v" for value (anything else)
            let type_str = match &ref_val {
                FormulaValue::Empty => "b",
                FormulaValue::String(s) if s.is_empty() => "b",
                FormulaValue::String(_) => "l",
                _ => "v", // Number, Boolean, Error
            };
            Ok(FormulaValue::String(type_str.to_string()))
        }
        "filename" => {
            // Cannot determine file path in standalone engine
            Ok(FormulaValue::String(String::new()))
        }
        "format" => {
            // Default: General format
            Ok(FormulaValue::String("G".to_string()))
        }
        "width" => {
            // Default column width (8 characters)
            Ok(FormulaValue::Number(8.0))
        }
        "protect" => {
            // Default: cell is protected
            Ok(FormulaValue::Number(1.0))
        }
        "prefix" => {
            // No alignment prefix available
            Ok(FormulaValue::String(String::new()))
        }
        "parentheses" => {
            // Not formatted with parentheses
            Ok(FormulaValue::Number(0.0))
        }
        "color" => {
            // Not formatted in color for negatives
            Ok(FormulaValue::Number(0.0))
        }
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

pub fn fn_info(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let _ = args;
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_sheet(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let _ = args;
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_sheets(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let _ = args;
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_isformula(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array { .. }) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(false))
}

pub fn fn_isomitted(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array { .. }) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(false))
}

// ---------- Stubs for functions that require external services ----------
// These return #N/A because they cannot work as standalone computations.

/// Helper: stub that propagates errors, otherwise returns #N/A.
fn stub_na(args: &[FormulaValue]) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }
    Ok(FormulaValue::Error(CellError::Na))
}

/// STOCKHISTORY(...) — Stub: requires Microsoft's live stock data feed.
pub fn fn_stockhistory(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// CALL(...) — Stub: calls a Windows DLL procedure at runtime.
pub fn fn_call(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// REGISTER.ID(...) — Stub: returns the register ID of a loaded DLL.
pub fn fn_register_id(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// CUBEKPIMEMBER(...) — Stub: requires OLAP server connection.
pub fn fn_cubekpimember(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// CUBEMEMBER(...) — Stub: requires OLAP server connection.
pub fn fn_cubemember(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// CUBEMEMBERPROPERTY(...) — Stub: requires OLAP server connection.
pub fn fn_cubememberproperty(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// CUBERANKEDMEMBER(...) — Stub: requires OLAP server connection.
pub fn fn_cuberankedmember(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// CUBESET(...) — Stub: requires OLAP server connection.
pub fn fn_cubeset(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// CUBESETCOUNT(...) — Stub: requires OLAP server connection.
pub fn fn_cubesetcount(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    stub_na(args)
}

/// CUBEVALUE(...) — Stub: requires OLAP server connection.
pub fn fn_cubevalue(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    stub_na(args)
}
#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    #[test]
    fn info_extra_iserr() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_iserr(&[FormulaValue::Error(CellError::Div0)], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            fn_iserr(&[FormulaValue::Error(CellError::Na)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_iserr(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_iserr(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_iseven() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_iseven(&[FormulaValue::Number(2.9)], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            fn_iseven(&[FormulaValue::Number(-3.9)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_iseven(&[FormulaValue::Number(0.0)], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            fn_iseven(&[FormulaValue::String("2".to_string())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_isodd() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_isodd(&[FormulaValue::Number(3.9)], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            fn_isodd(&[FormulaValue::Number(-2.1)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isodd(&[FormulaValue::Number(0.0)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isodd(&[FormulaValue::Boolean(true)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_islogical() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_islogical(&[FormulaValue::Boolean(true)], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            fn_islogical(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_islogical(&[FormulaValue::String("TRUE".to_string())], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_islogical(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_isnontext() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_isnontext(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            fn_isnontext(&[FormulaValue::Empty], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            fn_isnontext(&[FormulaValue::String("x".to_string())], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isnontext(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_isref() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_isref(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isref(&[FormulaValue::String("A1".to_string())], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isref(&[FormulaValue::Error(CellError::Ref)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isref(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_error_type() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Null)], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Div0)], &ctx).unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Value)], &ctx).unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Ref)], &ctx).unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Name)], &ctx).unwrap(),
            FormulaValue::Number(5.0)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Num)], &ctx).unwrap(),
            FormulaValue::Number(6.0)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Na)], &ctx).unwrap(),
            FormulaValue::Number(7.0)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::GettingData)], &ctx).unwrap(),
            FormulaValue::Number(8.0)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Spill)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn info_extra_type() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_type(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            fn_type(&[FormulaValue::String("x".to_string())], &ctx).unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            fn_type(&[FormulaValue::Boolean(true)], &ctx).unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            fn_type(&[FormulaValue::Error(CellError::Value)], &ctx).unwrap(),
            FormulaValue::Number(16.0)
        );
        assert_eq!(
            fn_type(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Number(64.0)
        );
        assert_eq!(
            fn_type(&[FormulaValue::Empty], &ctx).unwrap(),
            FormulaValue::Number(128.0)
        );
    }

    #[test]
    fn info_extra_cell_stub() {
        let ctx = EvaluationContext::simple();
        // "filename" now returns empty string instead of #N/A
        assert_eq!(
            fn_cell(&[FormulaValue::String("filename".to_string())], &ctx).unwrap(),
            FormulaValue::String(String::new())
        );
        // No args → #VALUE! (info_type is required)
        assert_eq!(
            fn_cell(&[], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        // Array as info_type → #VALUE!
        assert_eq!(
            fn_cell(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_info_stub() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_info(&[FormulaValue::String("osversion".to_string())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_info(&[], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_info(&[FormulaValue::Boolean(true)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn info_extra_sheet_stub() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_sheet(&[], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_sheet(&[FormulaValue::String("Sheet1".to_string())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_sheet(&[FormulaValue::Error(CellError::Ref)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn info_extra_sheets_stub() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_sheets(&[], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_sheets(&[FormulaValue::String("Sheet1".to_string())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_sheets(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn info_extra_isformula() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_isformula(&[FormulaValue::String("A1".to_string())], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isformula(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isformula(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_isomitted() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_isomitted(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isomitted(&[FormulaValue::String("x".to_string())], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            fn_isomitted(&[FormulaValue::Array { data: vec![], source: None }], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_eval_helper_smoke() {
        assert_eq!(eval("=NA()").unwrap(), FormulaValue::Error(CellError::Na));
    }

    // ---------- Stub tests: explicitly skipped external-service functions ----------
    // These functions are registered so formulas parse, but return #N/A because
    // they require runtime resources unavailable in a standalone spreadsheet engine.

    #[test]
    fn stub_stockhistory_returns_na() {
        // Requires Microsoft's live stock data feed.
        assert_eq!(
            eval("=STOCKHISTORY(\"MSFT\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn stub_call_returns_na() {
        // Calls a Windows DLL procedure at runtime.
        assert_eq!(
            eval("=CALL(\"kernel32\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn stub_register_id_returns_na() {
        // Returns the register ID of a loaded DLL.
        assert_eq!(
            eval("=REGISTER.ID(\"lib\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn stub_cube_functions_return_na() {
        // All 7 CUBE functions require an OLAP server connection.
        assert_eq!(
            eval("=CUBEKPIMEMBER(\"conn\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            eval("=CUBEMEMBER(\"conn\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            eval("=CUBEMEMBERPROPERTY(\"conn\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            eval("=CUBERANKEDMEMBER(\"conn\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            eval("=CUBESET(\"conn\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            eval("=CUBESETCOUNT(\"conn\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            eval("=CUBEVALUE(\"conn\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_iserr_docs() {
        // MS docs: ISERR returns TRUE for any error EXCEPT #N/A.

        // #DIV/0! is caught by ISERR
        assert_eq!(eval("=ISERR(1/0)").unwrap(), FormulaValue::Boolean(true));

        // #N/A is NOT caught by ISERR (key distinction from ISERROR)
        assert_eq!(eval("=ISERR(NA())").unwrap(), FormulaValue::Boolean(false));

        // Non-error values return FALSE
        assert_eq!(eval("=ISERR(1)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISERR(0)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(
            eval("=ISERR(\"text\")").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(eval("=ISERR(TRUE)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISERR(FALSE)").unwrap(), FormulaValue::Boolean(false));

        // #VALUE! error is caught by ISERR — test via direct function call
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_iserr(&[FormulaValue::Error(CellError::Value)], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );
    }

    #[test]
    fn test_iseven_docs() {
        // MS docs: ISEVEN returns TRUE if number is even (truncates decimals).
        assert_eq!(eval("=ISEVEN(-1)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISEVEN(2.5)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISEVEN(5)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISEVEN(0)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISEVEN(4)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISEVEN(-2)").unwrap(), FormulaValue::Boolean(true));
    }

    #[test]
    fn test_isodd_docs() {
        // MS docs: ISODD returns TRUE if number is odd (truncates decimals).
        assert_eq!(eval("=ISODD(-1)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISODD(1)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISODD(2.5)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISODD(5)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISODD(0)").unwrap(), FormulaValue::Boolean(false));
    }

    #[test]
    fn test_islogical_docs() {
        // MS docs: =ISLOGICAL(TRUE) → TRUE, =ISLOGICAL("TRUE") → FALSE
        assert_eq!(
            eval("=ISLOGICAL(TRUE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=ISLOGICAL(FALSE)").unwrap(),
            FormulaValue::Boolean(true)
        );

        // String "TRUE" is text, not a logical value
        assert_eq!(
            eval("=ISLOGICAL(\"TRUE\")").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=ISLOGICAL(\"FALSE\")").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Numbers are not logical
        assert_eq!(eval("=ISLOGICAL(1)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISLOGICAL(0)").unwrap(), FormulaValue::Boolean(false));

        // Errors are not logical
        assert_eq!(
            eval("=ISLOGICAL(NA())").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Result of a comparison IS logical
        assert_eq!(
            eval("=ISLOGICAL(1>0)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=ISLOGICAL(1=1)").unwrap(),
            FormulaValue::Boolean(true)
        );
    }

    #[test]
    fn test_isnontext_docs() {
        // MS docs: ISNONTEXT returns TRUE if value is any item that is NOT text.

        // Numbers are not text
        assert_eq!(eval("=ISNONTEXT(1)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISNONTEXT(0)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(
            eval("=ISNONTEXT(3.14)").unwrap(),
            FormulaValue::Boolean(true)
        );

        // Booleans are not text
        assert_eq!(
            eval("=ISNONTEXT(TRUE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=ISNONTEXT(FALSE)").unwrap(),
            FormulaValue::Boolean(true)
        );

        // Errors are not text
        assert_eq!(
            eval("=ISNONTEXT(NA())").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=ISNONTEXT(1/0)").unwrap(),
            FormulaValue::Boolean(true)
        );

        // Text strings return FALSE
        assert_eq!(
            eval("=ISNONTEXT(\"hello\")").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=ISNONTEXT(\"Region1\")").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Empty string is still text → FALSE
        assert_eq!(
            eval("=ISNONTEXT(\"\")").unwrap(),
            FormulaValue::Boolean(false)
        );
    }

    #[test]
    fn test_error_type_docs() {
        let ctx = EvaluationContext::simple();

        // #NULL! → 1
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Null)], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );
        // #DIV/0! → 2
        assert_eq!(eval("=ERROR.TYPE(1/0)").unwrap(), FormulaValue::Number(2.0));
        // #VALUE! → 3
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Value)], &ctx).unwrap(),
            FormulaValue::Number(3.0)
        );
        // #REF! → 4
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Ref)], &ctx).unwrap(),
            FormulaValue::Number(4.0)
        );
        // #NAME? → 5
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Name)], &ctx).unwrap(),
            FormulaValue::Number(5.0)
        );
        // #NUM! → 6
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::Num)], &ctx).unwrap(),
            FormulaValue::Number(6.0)
        );
        // #N/A → 7
        assert_eq!(
            eval("=ERROR.TYPE(NA())").unwrap(),
            FormulaValue::Number(7.0)
        );
        // #GETTING_DATA → 8
        assert_eq!(
            fn_error_type(&[FormulaValue::Error(CellError::GettingData)], &ctx).unwrap(),
            FormulaValue::Number(8.0)
        );
        // Non-error → #N/A
        assert_eq!(
            fn_error_type(&[FormulaValue::Number(42.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::String("hello".into())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_error_type(&[FormulaValue::Boolean(true)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_type_docs() {
        // =TYPE("Smith") → 2 (Text)
        assert_eq!(eval("=TYPE(\"Smith\")").unwrap(), FormulaValue::Number(2.0));

        // =TYPE(error) → 16 (Error) — use direct call since 2+"Smith" propagates error
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_type(&[FormulaValue::Error(CellError::Value)], &ctx).unwrap(),
            FormulaValue::Number(16.0)
        );

        // =TYPE({1,2;3,4}) → 64 (Array)
        assert_eq!(
            eval("=TYPE({1,2;3,4})").unwrap(),
            FormulaValue::Number(64.0)
        );

        // Type table coverage from docs
        assert_eq!(eval("=TYPE(1)").unwrap(), FormulaValue::Number(1.0)); // Number → 1
        assert_eq!(eval("=TYPE(\"text\")").unwrap(), FormulaValue::Number(2.0)); // Text → 2
        assert_eq!(eval("=TYPE(TRUE)").unwrap(), FormulaValue::Number(4.0)); // Logical → 4
        assert_eq!(eval("=TYPE(1/0)").unwrap(), FormulaValue::Number(16.0)); // Error → 16
        assert_eq!(eval("=TYPE({1,2,3})").unwrap(), FormulaValue::Number(64.0));
        // Array → 64
    }

    #[test]
    fn test_cell_docs() {
        // === Tests using EvaluationContext::simple() (row=0, col=0, sheet=0) ===
        let ctx = EvaluationContext::simple();

        // "address" — returns "$A$1" for cell at (row=0, col=0)
        assert_eq!(
            fn_cell(&[FormulaValue::String("address".to_string())], &ctx).unwrap(),
            FormulaValue::String("$A$1".to_string())
        );

        // "col" — column number, 1-based (col=0 → 1)
        assert_eq!(
            fn_cell(&[FormulaValue::String("col".to_string())], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );

        // "row" — row number, 1-based (row=0 → 1)
        assert_eq!(
            fn_cell(&[FormulaValue::String("row".to_string())], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );

        // "contents" — with no reference, returns 0 (Empty → 0 like Excel)
        assert_eq!(
            fn_cell(&[FormulaValue::String("contents".to_string())], &ctx).unwrap(),
            FormulaValue::Number(0.0)
        );

        // "contents" — with a number reference
        assert_eq!(
            fn_cell(
                &[
                    FormulaValue::String("contents".to_string()),
                    FormulaValue::Number(42.0)
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Number(42.0)
        );

        // "contents" — with a string reference
        assert_eq!(
            fn_cell(
                &[
                    FormulaValue::String("contents".to_string()),
                    FormulaValue::String("hello".to_string()),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::String("hello".to_string())
        );

        // "type" — "b" for blank (empty/no reference)
        assert_eq!(
            fn_cell(&[FormulaValue::String("type".to_string())], &ctx).unwrap(),
            FormulaValue::String("b".to_string())
        );

        // "type" — "l" for label (text)
        assert_eq!(
            fn_cell(
                &[
                    FormulaValue::String("type".to_string()),
                    FormulaValue::String("Smith".to_string()),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::String("l".to_string())
        );

        // "type" — "v" for value (number)
        assert_eq!(
            fn_cell(
                &[
                    FormulaValue::String("type".to_string()),
                    FormulaValue::Number(3.14),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::String("v".to_string())
        );

        // "type" — "v" for boolean
        assert_eq!(
            fn_cell(
                &[
                    FormulaValue::String("type".to_string()),
                    FormulaValue::Boolean(true),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::String("v".to_string())
        );

        // "type" — "b" for empty string (treated as blank)
        assert_eq!(
            fn_cell(
                &[
                    FormulaValue::String("type".to_string()),
                    FormulaValue::String(String::new()),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::String("b".to_string())
        );

        // "filename" — returns empty string (no file context)
        assert_eq!(
            fn_cell(&[FormulaValue::String("filename".to_string())], &ctx).unwrap(),
            FormulaValue::String(String::new())
        );

        // "format" — returns "G" (General)
        assert_eq!(
            fn_cell(&[FormulaValue::String("format".to_string())], &ctx).unwrap(),
            FormulaValue::String("G".to_string())
        );

        // "width" — returns 8 (default column width)
        assert_eq!(
            fn_cell(&[FormulaValue::String("width".to_string())], &ctx).unwrap(),
            FormulaValue::Number(8.0)
        );

        // "protect" — returns 1 (default protected)
        assert_eq!(
            fn_cell(&[FormulaValue::String("protect".to_string())], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );

        // "prefix" — returns empty string
        assert_eq!(
            fn_cell(&[FormulaValue::String("prefix".to_string())], &ctx).unwrap(),
            FormulaValue::String(String::new())
        );

        // "parentheses" — returns 0
        assert_eq!(
            fn_cell(&[FormulaValue::String("parentheses".to_string())], &ctx,).unwrap(),
            FormulaValue::Number(0.0)
        );

        // "color" — returns 0
        assert_eq!(
            fn_cell(&[FormulaValue::String("color".to_string())], &ctx).unwrap(),
            FormulaValue::Number(0.0)
        );

        // === Case-insensitive info_type ===
        assert_eq!(
            fn_cell(&[FormulaValue::String("ROW".to_string())], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            fn_cell(&[FormulaValue::String("Address".to_string())], &ctx).unwrap(),
            FormulaValue::String("$A$1".to_string())
        );

        // === Unknown info_type → #VALUE! ===
        assert_eq!(
            fn_cell(&[FormulaValue::String("unknown".to_string())], &ctx,).unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // === Error propagation from reference arg ===
        assert_eq!(
            fn_cell(
                &[
                    FormulaValue::String("type".to_string()),
                    FormulaValue::Error(CellError::Div0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Error(CellError::Div0)
        );

        // === Non-string info_type → #VALUE! ===
        assert_eq!(
            fn_cell(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            fn_cell(&[FormulaValue::Boolean(true)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // === With a different cell position ===
        let ctx2 = EvaluationContext::new(None, 0, 4, 2); // row=4, col=2 → C5
        assert_eq!(
            fn_cell(&[FormulaValue::String("address".to_string())], &ctx2).unwrap(),
            FormulaValue::String("$C$5".to_string())
        );
        assert_eq!(
            fn_cell(&[FormulaValue::String("row".to_string())], &ctx2).unwrap(),
            FormulaValue::Number(5.0)
        );
        assert_eq!(
            fn_cell(&[FormulaValue::String("col".to_string())], &ctx2).unwrap(),
            FormulaValue::Number(3.0)
        );

        // === MS docs example: =IF(CELL("type",A1)="v",A1*2,0) ===
        // Simulate with value reference (number → "v")
        assert_eq!(
            fn_cell(
                &[
                    FormulaValue::String("type".to_string()),
                    FormulaValue::Number(10.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::String("v".to_string())
        );
    }
}
