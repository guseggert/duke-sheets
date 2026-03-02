use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

pub fn fn_iserr(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(matches!(
        v,
        FormulaValue::Error(e) if *e != CellError::Na
    )))
}

pub fn fn_iseven(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let n = match v {
        FormulaValue::Number(n) => *n,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    Ok(FormulaValue::Boolean((n.trunc() as i64) % 2 == 0))
}

pub fn fn_isodd(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
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
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::Boolean(_))))
}

pub fn fn_isnontext(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(!matches!(v, FormulaValue::String(_))))
}

pub fn fn_isref(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(false))
}

pub fn fn_error_type(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
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
    let v = args.get(0).unwrap();
    let n = match v {
        FormulaValue::Number(_) => 1.0,
        FormulaValue::String(_) => 2.0,
        FormulaValue::Boolean(_) => 4.0,
        FormulaValue::Error(_) => 16.0,
        FormulaValue::Array(_) => 64.0,
        FormulaValue::Empty => 128.0,
    };

    Ok(FormulaValue::Number(n))
}

pub fn fn_cell(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let _ = args;
    Ok(FormulaValue::Error(CellError::Na))
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
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(false))
}

pub fn fn_isomitted(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::Boolean(false))
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
            fn_iserr(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
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
            fn_islogical(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
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
            fn_isnontext(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
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
            fn_isref(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
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
            fn_type(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
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
        assert_eq!(
            fn_cell(&[FormulaValue::String("filename".to_string())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_cell(&[], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_cell(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
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
            fn_sheets(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
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
            fn_isformula(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
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
            fn_isomitted(&[FormulaValue::Array(vec![])], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn info_extra_eval_helper_smoke() {
        assert_eq!(eval("=NA()").unwrap(), FormulaValue::Error(CellError::Na));
    }
}
