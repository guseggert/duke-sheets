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

    #[test]
    fn test_let_docs() {
        // https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
        // LET is a pass-through: evaluator pre-evaluates args, fn_let returns the last one.
        let ctx = EvaluationContext::simple();

        // Docs Example 1: =LET(x, 5, SUM(x, 1)) → 6
        let result = fn_let(
            &[
                FormulaValue::String("x".into()),
                FormulaValue::Number(5.0),
                FormulaValue::Number(6.0),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Number(6.0));

        // Via eval — string name, inline SUM(5,1)
        assert_eq!(
            eval("=LET(\"x\",5,SUM(5,1))").unwrap(),
            FormulaValue::Number(6.0)
        );

        // Inline arithmetic: =LET("x", 5, 5+1) → 6
        assert_eq!(
            eval("=LET(\"x\",5,5+1)").unwrap(),
            FormulaValue::Number(6.0)
        );

        // Two name/value pairs + calculation (5 args)
        let result = fn_let(
            &[
                FormulaValue::String("filterCriteria".into()),
                FormulaValue::String("Fred".into()),
                FormulaValue::String("filteredRange".into()),
                FormulaValue::String("filtered_data".into()),
                FormulaValue::String("Fred".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::String("Fred".into()));

        // Two pairs via eval: =LET("a", 10, "b", 20, 10+20) → 30
        assert_eq!(
            eval("=LET(\"a\",10,\"b\",20,10+20)").unwrap(),
            FormulaValue::Number(30.0)
        );

        // Minimum 3 args (one pair + calculation)
        let result = fn_let(
            &[
                FormulaValue::String("a".into()),
                FormulaValue::Number(10.0),
                FormulaValue::Number(10.0),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Number(10.0));

        // Boolean result
        let result = fn_let(
            &[
                FormulaValue::String("flag".into()),
                FormulaValue::Boolean(true),
                FormulaValue::Boolean(true),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Boolean(true));

        // String result via eval
        assert_eq!(
            eval("=LET(\"name\",\"hello\",\"hello\")").unwrap(),
            FormulaValue::String("hello".into())
        );

        // Error propagation: error in value position bubbles up
        let result = fn_let(
            &[
                FormulaValue::String("x".into()),
                FormulaValue::Error(CellError::Div0),
                FormulaValue::Number(42.0),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Div0));

        // Via eval: =LET("x", 1/0, 42) → #DIV/0!
        assert_eq!(
            eval("=LET(\"x\",1/0,42)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );

        // Error in name position is NOT checked — passes through to last arg
        let result = fn_let(
            &[
                FormulaValue::Error(CellError::Value),
                FormulaValue::Number(5.0),
                FormulaValue::Number(5.0),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Number(5.0));

        // Error in calculation position returned as-is
        let result = fn_let(
            &[
                FormulaValue::String("x".into()),
                FormulaValue::Number(5.0),
                FormulaValue::Error(CellError::Num),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Num));

        // Arg count: must be odd ≥ 3, else #VALUE!
        assert_eq!(
            fn_let(&[], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            fn_let(&[FormulaValue::Number(5.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            fn_let(
                &[FormulaValue::String("x".into()), FormulaValue::Number(5.0)],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            eval("=LET(\"a\",1,\"b\",2)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // 3 pairs (7 args)
        let result = fn_let(
            &[
                FormulaValue::String("x".into()),
                FormulaValue::Number(1.0),
                FormulaValue::String("y".into()),
                FormulaValue::Number(2.0),
                FormulaValue::String("z".into()),
                FormulaValue::Number(3.0),
                FormulaValue::Number(6.0),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Number(6.0));

        // 3 pairs via eval
        assert_eq!(
            eval("=LET(\"x\",1,\"y\",2,\"z\",3,1+2+3)").unwrap(),
            FormulaValue::Number(6.0)
        );

        // Nested function in calculation
        assert_eq!(
            eval("=LET(\"x\",10,IF(10>5,\"big\",\"small\"))").unwrap(),
            FormulaValue::String("big".into())
        );
        assert_eq!(
            eval("=LET(\"n\",100,SQRT(100))").unwrap(),
            FormulaValue::Number(10.0)
        );
    }

    #[test]
    fn test_lambda_docs() {
        // https://support.microsoft.com/en-us/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67
        // LAMBDA is a pass-through: returns the last argument.
        let ctx = EvaluationContext::simple();

        // Docs Step 2: =LAMBDA(number, number + 1)(1) => 2
        // Pass-through returns last arg
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("number".into()),
                    FormulaValue::Number(6.0)
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );

        // Example 1 (ToCelsius): two args => returns calculation (last arg)
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("temp".into()),
                    FormulaValue::Number(40.0)
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Number(40.0)
        );

        // Example 2 (Hypotenuse): three args => returns last
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("a".into()),
                    FormulaValue::String("b".into()),
                    FormulaValue::Number(5.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Number(5.0)
        );

        // Single arg = just the calculation, no parameters
        assert_eq!(
            fn_lambda(&[FormulaValue::Number(42.0)], &ctx).unwrap(),
            FormulaValue::Number(42.0)
        );

        // Boolean calculation result
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("x".into()),
                    FormulaValue::Boolean(true)
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Boolean(true)
        );

        // String calculation result
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("text".into()),
                    FormulaValue::String("hello".into()),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::String("hello".into())
        );

        // Error propagation: if calculation is an error, return it
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("x".into()),
                    FormulaValue::Error(CellError::Value),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // Empty args => #VALUE!
        assert_eq!(
            fn_lambda(&[], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }
}
