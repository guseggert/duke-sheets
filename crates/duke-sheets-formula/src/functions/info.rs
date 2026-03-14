//! Information functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

/// ISBLANK(value)
pub fn fn_isblank(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::Empty)))
}

/// ISNUMBER(value)
pub fn fn_isnumber(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::Number(_))))
}

/// ISTEXT(value)
pub fn fn_istext(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::String(_))))
}

/// ISERROR(value)
pub fn fn_iserror(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::Error(_))))
}

/// ISNA(value)
pub fn fn_isna(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(
        v,
        FormulaValue::Error(CellError::Na)
    )))
}

/// NA()
pub fn fn_na(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
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
    fn test_isblank_docs() {
        // MS docs: ISBLANK returns TRUE if value refers to an empty cell.
        // We can't reference empty cells via eval, so test with literals.

        // Number is not blank
        assert_eq!(eval("=ISBLANK(0)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISBLANK(1)").unwrap(), FormulaValue::Boolean(false));

        // Text is not blank (even empty string)
        assert_eq!(
            eval("=ISBLANK(\"\")").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=ISBLANK(\"hello\")").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Boolean is not blank
        assert_eq!(
            eval("=ISBLANK(TRUE)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=ISBLANK(FALSE)").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Error values are not blank
        assert_eq!(eval("=ISBLANK(1/0)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(
            eval("=ISBLANK(NA())").unwrap(),
            FormulaValue::Boolean(false)
        );
    }

    #[test]
    fn test_isnumber_docs() {
        // Docs Example 1: =ISNUMBER(4) → TRUE
        assert_eq!(eval("=ISNUMBER(4)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(
            eval("=ISNUMBER(330.92)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(eval("=ISNUMBER(0)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISNUMBER(-5)").unwrap(), FormulaValue::Boolean(true));

        // Docs remark: numeric text is NOT converted — remains text
        assert_eq!(
            eval("=ISNUMBER(\"19\")").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Text is not a number
        assert_eq!(
            eval("=ISNUMBER(\"hello\")").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Booleans are not numbers
        assert_eq!(
            eval("=ISNUMBER(TRUE)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=ISNUMBER(FALSE)").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Errors are not numbers
        assert_eq!(
            eval("=ISNUMBER(NA())").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Result of arithmetic IS a number
        assert_eq!(eval("=ISNUMBER(2+3)").unwrap(), FormulaValue::Boolean(true));
    }

    #[test]
    fn test_istext_docs() {
        // Docs Example 2: =ISTEXT("Region1") → TRUE
        assert_eq!(
            eval("=ISTEXT(\"hello\")").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=ISTEXT(\"Region1\")").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(eval("=ISTEXT(\"\")").unwrap(), FormulaValue::Boolean(true));

        // Numeric text "19" IS text
        assert_eq!(
            eval("=ISTEXT(\"19\")").unwrap(),
            FormulaValue::Boolean(true)
        );

        // Numbers are not text
        assert_eq!(eval("=ISTEXT(1)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISTEXT(0)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISTEXT(3.14)").unwrap(), FormulaValue::Boolean(false));

        // Booleans are not text
        assert_eq!(eval("=ISTEXT(TRUE)").unwrap(), FormulaValue::Boolean(false));

        // Errors are not text
        assert_eq!(eval("=ISTEXT(NA())").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISTEXT(1/0)").unwrap(), FormulaValue::Boolean(false));
    }

    #[test]
    fn test_iserror_docs() {
        // MS docs: ISERROR returns TRUE for ANY error value

        // #DIV/0!
        assert_eq!(eval("=ISERROR(1/0)").unwrap(), FormulaValue::Boolean(true));

        // #N/A — ISERROR catches it (unlike ISERR)
        assert_eq!(eval("=ISERROR(NA())").unwrap(), FormulaValue::Boolean(true));

        // #VALUE! — test via direct function call since arithmetic error propagates
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_iserror(&[FormulaValue::Error(CellError::Value)], &ctx).unwrap(),
            FormulaValue::Boolean(true)
        );

        // Non-error values return FALSE
        assert_eq!(eval("=ISERROR(1)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(
            eval("=ISERROR(\"text\")").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=ISERROR(TRUE)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(eval("=ISERROR(0)").unwrap(), FormulaValue::Boolean(false));
    }

    #[test]
    fn test_isna_docs() {
        // MS docs: ISNA returns TRUE only for the #N/A error value.

        // #N/A is the only error ISNA catches
        assert_eq!(eval("=ISNA(NA())").unwrap(), FormulaValue::Boolean(true));

        // #DIV/0! is NOT #N/A
        assert_eq!(eval("=ISNA(1/0)").unwrap(), FormulaValue::Boolean(false));

        // #VALUE! is NOT #N/A — test via direct function call
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_isna(&[FormulaValue::Error(CellError::Value)], &ctx).unwrap(),
            FormulaValue::Boolean(false)
        );

        // Non-error values return FALSE
        assert_eq!(eval("=ISNA(1)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(
            eval("=ISNA(\"text\")").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(eval("=ISNA(TRUE)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=ISNA(0)").unwrap(), FormulaValue::Boolean(false));
    }

    #[test]
    fn test_na_docs() {
        // NA() returns the #N/A error value
        assert_eq!(eval("=NA()").unwrap(), FormulaValue::Error(CellError::Na));

        // Docs remark: ISNA recognizes the NA() result
        assert_eq!(eval("=ISNA(NA())").unwrap(), FormulaValue::Boolean(true));
    }

    #[test]
    fn test_n_docs() {
        // N(7) - number stays as number
        assert_eq!(eval("=N(7)").unwrap(), FormulaValue::Number(7.0));

        // N("Even") - text returns 0
        assert_eq!(eval("=N(\"Even\")").unwrap(), FormulaValue::Number(0.0));

        // N(TRUE) - TRUE returns 1
        assert_eq!(eval("=N(TRUE)").unwrap(), FormulaValue::Number(1.0));

        // N(FALSE) - FALSE returns 0
        assert_eq!(eval("=N(FALSE)").unwrap(), FormulaValue::Number(0.0));

        // N("7") - text string "7" returns 0
        assert_eq!(eval("=N(\"7\")").unwrap(), FormulaValue::Number(0.0));

        // N(error) - errors pass through
        assert_eq!(
            eval("=N(1/0)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
    }
}
