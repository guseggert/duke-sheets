//! Information functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

/// ISBLANK(value)
pub fn fn_isblank(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::Empty)))
}

/// ISNUMBER(value)
pub fn fn_isnumber(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::Number(_))))
}

/// ISTEXT(value)
pub fn fn_istext(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::String(_))))
}

/// ISERROR(value)
pub fn fn_iserror(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::Boolean(matches!(v, FormulaValue::Error(_))))
}

/// ISNA(value)
pub fn fn_isna(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
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
