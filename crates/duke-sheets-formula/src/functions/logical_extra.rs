use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

pub fn fn_let(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // LET(name1, value1, ..., nameN, valueN, calculation)
    // Args come in pre-evaluated. The last arg is the calculation result.
    // Must have odd number of args >= 3 (pairs + final calculation)
    if args.len() < 3 || args.len() % 2 == 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    for (i, arg) in args.iter().enumerate() {
        if i % 2 == 1 && i < args.len() - 1 {
            if let FormulaValue::Error(e) = arg {
                return Ok(FormulaValue::Error(*e));
            }
        }
    }

    Ok(args.last().unwrap().clone())
}

pub fn fn_lambda(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.is_empty() {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(args.last().unwrap().clone())
}

pub fn fn_map(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_reduce(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_scan(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_bycol(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_byrow(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_makearray(
    _args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    #[test]
    fn test_eval_helper_works() {
        let result = eval("=1+1").unwrap();
        assert_eq!(result, FormulaValue::Number(2.0));
    }

    #[test]
    fn test_let_passthrough() {
        let ctx = EvaluationContext::simple();
        let result = fn_let(
            &[
                FormulaValue::String("x".into()),
                FormulaValue::Number(5.0),
                FormulaValue::Number(5.0),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Number(5.0));
    }

    #[test]
    fn test_lambda_passthrough() {
        let ctx = EvaluationContext::simple();
        let result = fn_lambda(
            &[FormulaValue::String("x".into()), FormulaValue::Number(42.0)],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Number(42.0));
    }

    #[test]
    fn test_map_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_map(
            &[FormulaValue::Number(1.0), FormulaValue::Number(2.0)],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_reduce_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_reduce(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Array(vec![vec![FormulaValue::Number(1.0)]]),
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_scan_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_scan(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Array(vec![vec![FormulaValue::Number(1.0)]]),
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_bycol_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_bycol(
            &[
                FormulaValue::Array(vec![vec![FormulaValue::Number(1.0)]]),
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_byrow_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_byrow(
            &[
                FormulaValue::Array(vec![vec![FormulaValue::Number(1.0)]]),
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_makearray_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_makearray(
            &[
                FormulaValue::Number(2.0),
                FormulaValue::Number(3.0),
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }
}
