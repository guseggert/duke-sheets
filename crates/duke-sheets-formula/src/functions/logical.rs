//! Logical functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

/// IF function
pub fn fn_if(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let condition = args.get(0).ok_or_else(|| {
        crate::error::FormulaError::Argument("IF requires at least 2 arguments".into())
    })?;

    let if_true = args.get(1).ok_or_else(|| {
        crate::error::FormulaError::Argument("IF requires at least 2 arguments".into())
    })?;

    let if_false = args.get(2);

    // Evaluate condition
    let condition_bool = match condition {
        FormulaValue::Boolean(b) => *b,
        FormulaValue::Number(n) => *n != 0.0,
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if condition_bool {
        Ok(if_true.clone())
    } else {
        Ok(if_false.cloned().unwrap_or(FormulaValue::Boolean(false)))
    }
}

/// AND function
pub fn fn_and(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for arg in args {
        match arg {
            FormulaValue::Boolean(false) => return Ok(FormulaValue::Boolean(false)),
            FormulaValue::Number(n) if *n == 0.0 => return Ok(FormulaValue::Boolean(false)),
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array(arr) => {
                for row in arr {
                    for cell in row {
                        match cell {
                            FormulaValue::Boolean(false) => {
                                return Ok(FormulaValue::Boolean(false))
                            }
                            FormulaValue::Number(n) if *n == 0.0 => {
                                return Ok(FormulaValue::Boolean(false))
                            }
                            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                            _ => {}
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(FormulaValue::Boolean(true))
}

/// OR function
pub fn fn_or(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for arg in args {
        match arg {
            FormulaValue::Boolean(true) => return Ok(FormulaValue::Boolean(true)),
            FormulaValue::Number(n) if *n != 0.0 => return Ok(FormulaValue::Boolean(true)),
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array(arr) => {
                for row in arr {
                    for cell in row {
                        match cell {
                            FormulaValue::Boolean(true) => return Ok(FormulaValue::Boolean(true)),
                            FormulaValue::Number(n) if *n != 0.0 => {
                                return Ok(FormulaValue::Boolean(true))
                            }
                            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                            _ => {}
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(FormulaValue::Boolean(false))
}

/// NOT function
pub fn fn_not(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arg = args
        .get(0)
        .ok_or_else(|| crate::error::FormulaError::Argument("NOT requires 1 argument".into()))?;

    match arg {
        FormulaValue::Boolean(b) => Ok(FormulaValue::Boolean(!b)),
        FormulaValue::Number(n) => Ok(FormulaValue::Boolean(*n == 0.0)),
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// IFERROR(value, value_if_error) - Returns value_if_error if value is an error, otherwise returns value
/// Reference: LibreOffice ScInterpreter::ScIfError
pub fn fn_iferror(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let value = args.get(0).ok_or_else(|| {
        crate::error::FormulaError::Argument("IFERROR requires 2 arguments".into())
    })?;

    let value_if_error = args.get(1).ok_or_else(|| {
        crate::error::FormulaError::Argument("IFERROR requires 2 arguments".into())
    })?;

    // If the first argument is an error, return the second argument
    // Otherwise, return the first argument as-is
    match value {
        FormulaValue::Error(_) => Ok(value_if_error.clone()),
        _ => Ok(value.clone()),
    }
}

/// IFNA(value, value_if_na) - Returns value_if_na if value is #N/A error, otherwise returns value
/// Similar to IFERROR but only catches #N/A errors
pub fn fn_ifna(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let value = args
        .get(0)
        .ok_or_else(|| crate::error::FormulaError::Argument("IFNA requires 2 arguments".into()))?;

    let value_if_na = args
        .get(1)
        .ok_or_else(|| crate::error::FormulaError::Argument("IFNA requires 2 arguments".into()))?;

    // Only catch #N/A errors, propagate all other errors
    match value {
        FormulaValue::Error(CellError::Na) => Ok(value_if_na.clone()),
        _ => Ok(value.clone()),
    }
}

/// TRUE() - Returns the logical value TRUE
pub fn fn_true(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    Ok(FormulaValue::Boolean(true))
}

/// FALSE() - Returns the logical value FALSE
pub fn fn_false(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    Ok(FormulaValue::Boolean(false))
}

/// XOR(logical1, [logical2], ...) - Returns logical exclusive OR of all arguments
/// Returns TRUE if an odd number of arguments are TRUE
pub fn fn_xor(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut true_count = 0;

    for arg in args {
        match arg {
            FormulaValue::Boolean(true) => true_count += 1,
            FormulaValue::Number(n) if *n != 0.0 => true_count += 1,
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array(arr) => {
                for row in arr {
                    for cell in row {
                        match cell {
                            FormulaValue::Boolean(true) => true_count += 1,
                            FormulaValue::Number(n) if *n != 0.0 => true_count += 1,
                            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                            _ => {}
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // XOR is true if odd number of TRUE values
    Ok(FormulaValue::Boolean(true_count % 2 == 1))
}

/// IFS(condition1, value1, [condition2, value2], ...) - Checks conditions and returns corresponding value
/// Reference: LibreOffice ScInterpreter::ScIfs_MS, Microsoft IFS function
///
/// Evaluates conditions in order and returns the value for the first TRUE condition.
/// Returns #N/A if no condition is TRUE (no default/else clause).
pub fn fn_ifs(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Must have at least 2 arguments (one condition-value pair)
    if args.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Must have even number of arguments (condition-value pairs)
    if args.len() % 2 != 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Process condition-value pairs
    let mut i = 0;
    while i < args.len() {
        let condition = &args[i];
        let value = &args[i + 1];

        // Check for error in condition
        if let FormulaValue::Error(e) = condition {
            return Ok(FormulaValue::Error(*e));
        }

        // Evaluate condition as boolean
        let condition_bool = match condition {
            FormulaValue::Boolean(b) => *b,
            FormulaValue::Number(n) => *n != 0.0,
            FormulaValue::String(s) => {
                let upper = s.to_uppercase();
                if upper == "TRUE" {
                    true
                } else if upper == "FALSE" {
                    false
                } else {
                    return Ok(FormulaValue::Error(CellError::Value));
                }
            }
            FormulaValue::Empty => false,
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        };

        if condition_bool {
            // Found a TRUE condition, return its value
            return Ok(value.clone());
        }

        i += 2;
    }

    // No TRUE condition found, return #N/A
    Ok(FormulaValue::Error(CellError::Na))
}

/// SWITCH(expression, value1, result1, [value2, result2], ..., [default])
/// Reference: LibreOffice ScInterpreter::ScSwitch_MS, Microsoft SWITCH function
///
/// Evaluates expression against values and returns the result for the first match.
/// If odd number of args after expression, last arg is the default.
/// Returns #N/A if no match and no default.
pub fn fn_switch(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Must have at least 3 arguments (expression, value1, result1)
    if args.len() < 3 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let expression = &args[0];

    // Check for error in expression
    if let FormulaValue::Error(e) = expression {
        return Ok(FormulaValue::Error(*e));
    }

    // Determine if we have a default value
    // If (args.len() - 1) is odd, we have pairs + default
    // If (args.len() - 1) is even, we have only pairs
    let remaining = args.len() - 1; // args after expression
    let has_default = remaining % 2 == 1;
    let num_pairs = if has_default {
        (remaining - 1) / 2
    } else {
        remaining / 2
    };

    // Process value-result pairs
    for pair_idx in 0..num_pairs {
        let value_idx = 1 + pair_idx * 2;
        let result_idx = value_idx + 1;

        let value = &args[value_idx];
        let result = &args[result_idx];

        // Check for error in value
        if let FormulaValue::Error(e) = value {
            return Ok(FormulaValue::Error(*e));
        }

        // Compare expression with value
        if values_match(expression, value) {
            return Ok(result.clone());
        }
    }

    // No match found
    if has_default {
        // Return the default value (last argument)
        Ok(args.last().unwrap().clone())
    } else {
        // No default, return #N/A
        Ok(FormulaValue::Error(CellError::Na))
    }
}

/// Helper function to compare two values for SWITCH
fn values_match(a: &FormulaValue, b: &FormulaValue) -> bool {
    match (a, b) {
        (FormulaValue::Number(x), FormulaValue::Number(y)) => (x - y).abs() < 1e-10,
        (FormulaValue::Boolean(x), FormulaValue::Boolean(y)) => x == y,
        (FormulaValue::String(x), FormulaValue::String(y)) => x.eq_ignore_ascii_case(y),
        (FormulaValue::Empty, FormulaValue::Empty) => true,

        // Number/Boolean coercion
        (FormulaValue::Number(n), FormulaValue::Boolean(b))
        | (FormulaValue::Boolean(b), FormulaValue::Number(n)) => {
            let b_num = if *b { 1.0 } else { 0.0 };
            (n - b_num).abs() < 1e-10
        }

        // Empty coercions
        (FormulaValue::Empty, FormulaValue::Number(n))
        | (FormulaValue::Number(n), FormulaValue::Empty) => n.abs() < 1e-10,
        (FormulaValue::Empty, FormulaValue::String(s))
        | (FormulaValue::String(s), FormulaValue::Empty) => s.is_empty(),

        _ => false,
    }
}
