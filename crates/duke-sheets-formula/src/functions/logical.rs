//! Logical functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

/// IF function
pub fn fn_if(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let condition = args.first().ok_or_else(|| {
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
        .first()
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
    let value = args.first().ok_or_else(|| {
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
        .first()
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

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    fn number(v: FormulaValue) -> f64 {
        match v {
            FormulaValue::Number(n) => n,
            other => panic!("Expected number, got {:?}", other),
        }
    }

    fn assert_close(actual: f64, expected: f64) {
        assert!(
            (actual - expected).abs() < 1e-9,
            "actual={actual}, expected={expected}"
        );
    }

    #[test]
    fn test_if_docs() {
        // Docs example: =IF(C2="Yes",1,2)
        assert_eq!(
            eval(r#"=IF("Yes"="Yes",1,2)"#).unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval(r#"=IF("No"="Yes",1,2)"#).unwrap(),
            FormulaValue::Number(2.0)
        );

        // Docs example: =IF(C2=1,"Yes","No")
        assert_eq!(
            eval(r#"=IF(1=1,"Yes","No")"#).unwrap(),
            FormulaValue::String("Yes".into())
        );
        assert_eq!(
            eval(r#"=IF(0=1,"Yes","No")"#).unwrap(),
            FormulaValue::String("No".into())
        );

        // Docs example: =IF(C2>B2,"Over Budget","Within Budget")
        assert_eq!(
            eval(r#"=IF(2>1,"Over Budget","Within Budget")"#).unwrap(),
            FormulaValue::String("Over Budget".into())
        );
        assert_eq!(
            eval(r#"=IF(1>2,"Over Budget","Within Budget")"#).unwrap(),
            FormulaValue::String("Within Budget".into())
        );

        // Docs example: =IF(C2>B2,C2-B2,0)
        assert_eq!(eval("=IF(2>1,2-1,0)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(eval("=IF(1>2,1-2,0)").unwrap(), FormulaValue::Number(0.0));

        // Docs example: =IF(E7="Yes",F5*0.0825,0)
        assert_eq!(
            eval(r#"=IF("Yes"="Yes",100*0.0825,0)"#).unwrap(),
            FormulaValue::Number(8.25)
        );
        assert_eq!(
            eval(r#"=IF("No"="Yes",100*0.0825,0)"#).unwrap(),
            FormulaValue::Number(0.0)
        );

        // Docs syntax example: =IF(A2>B2,"Over Budget","OK")
        assert_eq!(
            eval(r#"=IF(3>2,"Over Budget","OK")"#).unwrap(),
            FormulaValue::String("Over Budget".into())
        );
        assert_eq!(
            eval(r#"=IF(2>3,"Over Budget","OK")"#).unwrap(),
            FormulaValue::String("OK".into())
        );

        // Docs syntax example: =IF(A2=B2,B4-A4,"")
        assert_eq!(
            eval(r#"=IF(5=5,10-3,"")"#).unwrap(),
            FormulaValue::Number(7.0)
        );
        assert_eq!(
            eval(r#"=IF(5=6,10-3,"")"#).unwrap(),
            FormulaValue::String("".into())
        );

        // Docs: TRUE or FALSE literals as conditions
        assert_eq!(eval("=IF(TRUE,1,2)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(eval("=IF(FALSE,1,2)").unwrap(), FormulaValue::Number(2.0));

        // Docs: omitted value_if_false returns FALSE
        assert_eq!(eval("=IF(TRUE,1)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(eval("=IF(FALSE,1)").unwrap(), FormulaValue::Boolean(false));

        // Docs: numeric condition (non-zero=TRUE, zero=FALSE)
        assert_eq!(
            eval(r#"=IF(1,"Yes","No")"#).unwrap(),
            FormulaValue::String("Yes".into())
        );
        assert_eq!(
            eval(r#"=IF(0,"Yes","No")"#).unwrap(),
            FormulaValue::String("No".into())
        );
    }

    #[test]
    fn test_and_docs() {
        // Technical Details: AND returns TRUE if all args TRUE
        assert_eq!(
            eval("=AND(TRUE,TRUE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=AND(TRUE,FALSE)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=AND(FALSE,TRUE)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=AND(FALSE,FALSE)").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Single argument
        assert_eq!(eval("=AND(TRUE)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=AND(FALSE)").unwrap(), FormulaValue::Boolean(false));

        // Docs Example 1: =AND(A2>1,A2<100) where A2=50 → TRUE
        assert_eq!(
            eval("=AND(50>1,50<100)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=AND(0>1,50<100)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=AND(50>1,200<100)").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Docs Example 2: =IF(AND(A2<A3,A2<100),A2,"The value is out of range")
        assert_eq!(
            eval("=IF(AND(50<100,50<100),50,\"The value is out of range\")").unwrap(),
            FormulaValue::Number(50.0)
        );

        // Docs Example 3: A3=100, 100<100 is FALSE
        assert_eq!(
            eval("=IF(AND(100>1,100<100),100,\"The value is out of range\")").unwrap(),
            FormulaValue::String("The value is out of range".into())
        );

        // Docs Bonus Calculation: =IF(AND(sales>=goal,accounts>=goal),sales*rate,0)
        assert_eq!(
            eval("=IF(AND(125000>=100000,55>=50),125000*0.12,0)").unwrap(),
            FormulaValue::Number(15000.0)
        );
        assert_eq!(
            eval("=IF(AND(95000>=100000,55>=50),95000*0.12,0)").unwrap(),
            FormulaValue::Number(0.0)
        );
        assert_eq!(
            eval("=IF(AND(125000>=100000,40>=50),125000*0.12,0)").unwrap(),
            FormulaValue::Number(0.0)
        );

        // Numbers coerce (non-zero=TRUE, zero=FALSE)
        assert_eq!(eval("=AND(1,1)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=AND(1,0)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=AND(0,0)").unwrap(), FormulaValue::Boolean(false));

        // Multiple conditions
        assert_eq!(
            eval("=AND(TRUE,TRUE,TRUE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=AND(TRUE,TRUE,TRUE,TRUE,TRUE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=AND(TRUE,TRUE,FALSE,TRUE,TRUE)").unwrap(),
            FormulaValue::Boolean(false)
        );

        // Error propagation
        assert_eq!(
            eval("=AND(TRUE,1/0)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
        assert_eq!(
            eval("=AND(1/0,TRUE)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
    }

    #[test]
    fn test_or_docs() {
        // Basic OR examples
        assert_eq!(eval("=OR(TRUE,TRUE)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(
            eval("=OR(TRUE,FALSE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=OR(1=1,2=2,3=3)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=OR(1=2,2=3,3=4)").unwrap(),
            FormulaValue::Boolean(false)
        );

        // OR with IF (A2=50, A3=100)
        assert_eq!(
            eval("=OR(50>1,50<100)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=IF(OR(50>1,50<100),100,\"The value is out of range\")").unwrap(),
            FormulaValue::Number(100.0)
        );
        assert_eq!(
            eval("=IF(OR(50<0,50>50),50,\"The value is out of range\")").unwrap(),
            FormulaValue::String("The value is out of range".into())
        );

        // Sales commission: IF(OR(sales>=goal, accounts>=goal), sales*rate, 0)
        assert_close(
            number(eval("=IF(OR(10260>=8500,9>=5),10260*0.02,0)").unwrap()),
            205.2,
        );
        assert_eq!(
            eval("=IF(OR(15700>=8500,7>=5),15700*0.02,0)").unwrap(),
            FormulaValue::Number(314.0)
        );
        assert_eq!(
            eval("=IF(OR(13275>=8500,5>=5),13275*0.02,0)").unwrap(),
            FormulaValue::Number(265.5)
        );
        assert_eq!(
            eval("=IF(OR(9100>=8500,3>=5),9100*0.02,0)").unwrap(),
            FormulaValue::Number(182.0)
        );
        // Neither goal met, commission=0
        assert_eq!(
            eval("=IF(OR(7480>=8500,4>=5),7480*0.02,0)").unwrap(),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_not_docs() {
        // NOT(TRUE) = FALSE, NOT(FALSE) = TRUE
        assert_eq!(eval("=NOT(TRUE)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=NOT(FALSE)").unwrap(), FormulaValue::Boolean(true));

        // =NOT(1+1=2) → FALSE
        assert_eq!(eval("=NOT(1+1=2)").unwrap(), FormulaValue::Boolean(false));

        // =NOT(50>100) → TRUE (50 is NOT greater than 100)
        assert_eq!(eval("=NOT(50>100)").unwrap(), FormulaValue::Boolean(true));

        // =IF(AND(NOT(50>1),NOT(50<100)),50,"The value is out of range")
        assert_eq!(
            eval(r#"=IF(AND(NOT(50>1),NOT(50<100)),50,"The value is out of range")"#).unwrap(),
            FormulaValue::String("The value is out of range".to_string())
        );

        // =IF(OR(NOT(100<0),NOT(100>50)),100,"The value is out of range")
        assert_eq!(
            eval(r#"=IF(OR(NOT(100<0),NOT(100>50)),100,"The value is out of range")"#).unwrap(),
            FormulaValue::Number(100.0)
        );
    }

    #[test]
    fn test_iferror_docs() {
        // =IFERROR(210/35, "Error in calculation") → 6.0
        assert_eq!(
            eval("=IFERROR(210/35, \"Error in calculation\")").unwrap(),
            FormulaValue::Number(6.0)
        );

        // =IFERROR(55/0, "Error in calculation") → "Error in calculation"
        assert_eq!(
            eval("=IFERROR(55/0, \"Error in calculation\")").unwrap(),
            FormulaValue::String("Error in calculation".into())
        );

        // =IFERROR(0/23, "Error in calculation") → 0.0
        assert_eq!(
            eval("=IFERROR(0/23, \"Error in calculation\")").unwrap(),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_ifna_docs() {
        // IFNA returns value_if_na when formula returns #N/A, otherwise returns formula result

        // #N/A → returns alternate value
        assert_eq!(
            eval("=IFNA(NA(), \"Not found\")").unwrap(),
            FormulaValue::String("Not found".into())
        );

        // Non-error value passes through
        assert_eq!(
            eval("=IFNA(5, \"Not found\")").unwrap(),
            FormulaValue::Number(5.0)
        );

        // String passes through
        assert_eq!(
            eval(r#"=IFNA("Seattle", "Not found")"#).unwrap(),
            FormulaValue::String("Seattle".into())
        );

        // Boolean passes through
        assert_eq!(
            eval("=IFNA(TRUE(), \"Not found\")").unwrap(),
            FormulaValue::Boolean(true)
        );

        // #N/A with numeric alternate value
        assert_eq!(eval("=IFNA(NA(), 0)").unwrap(), FormulaValue::Number(0.0));

        // Other errors NOT caught by IFNA — they propagate
        // #DIV/0!
        assert_eq!(
            eval("=IFNA(1/0, \"Not found\")").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );

        // value_if_na is itself NA()
        assert_eq!(
            eval("=IFNA(NA(), NA())").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_ifs_docs() {
        // Grade scoring: =IFS(score>89,"A",score>79,"B",...,TRUE,"F")
        assert_eq!(
            eval(r#"=IFS(45>89,"A",45>79,"B",45>69,"C",45>59,"D",TRUE,"F")"#).unwrap(),
            FormulaValue::String("F".into())
        );
        assert_eq!(
            eval(r#"=IFS(90>89,"A",90>79,"B",90>69,"C",90>59,"D",TRUE,"F")"#).unwrap(),
            FormulaValue::String("A".into())
        );
        assert_eq!(
            eval(r#"=IFS(78>89,"A",78>79,"B",78>69,"C",78>59,"D",TRUE,"F")"#).unwrap(),
            FormulaValue::String("C".into())
        );
        assert_eq!(
            eval(r#"=IFS(80>89,"A",80>79,"B",80>69,"C",80>59,"D",TRUE,"F")"#).unwrap(),
            FormulaValue::String("B".into())
        );
        assert_eq!(
            eval(r#"=IFS(60>89,"A",60>79,"B",60>69,"C",60>59,"D",TRUE,"F")"#).unwrap(),
            FormulaValue::String("D".into())
        );

        // Day of week mapping
        assert_eq!(
            eval(r#"=IFS(1=1,"Sunday",1=2,"Monday",1=3,"Tuesday",1=4,"Wednesday",1=5,"Thursday",1=6,"Friday",1=7,"Saturday")"#).unwrap(),
            FormulaValue::String("Sunday".into())
        );
        assert_eq!(
            eval(r#"=IFS(3=1,"Sunday",3=2,"Monday",3=3,"Tuesday",3=4,"Wednesday",3=5,"Thursday",3=6,"Friday",3=7,"Saturday")"#).unwrap(),
            FormulaValue::String("Tuesday".into())
        );
        assert_eq!(
            eval(r#"=IFS(7=1,"Sunday",7=2,"Monday",7=3,"Tuesday",7=4,"Wednesday",7=5,"Thursday",7=6,"Friday",7=7,"Saturday")"#).unwrap(),
            FormulaValue::String("Saturday".into())
        );

        // TRUE as default catch-all
        assert_eq!(
            eval(r#"=IFS(FALSE,"never",TRUE,"default")"#).unwrap(),
            FormulaValue::String("default".into())
        );

        // No TRUE conditions → #N/A
        assert_eq!(
            eval(r#"=IFS(1>2,"A",3>4,"B")"#).unwrap(),
            FormulaValue::Error(CellError::Na)
        );

        // First TRUE wins (order matters)
        assert_eq!(
            eval(r#"=IFS(1<2,"first",2<3,"second",TRUE,"third")"#).unwrap(),
            FormulaValue::String("first".into())
        );
    }

    #[test]
    fn test_switch_docs() {
        // SWITCH(2,...) matches value 2 => "Monday"
        assert_eq!(
            eval("=SWITCH(2,1,\"Sunday\",2,\"Monday\",3,\"Tuesday\",\"No match\")").unwrap(),
            FormulaValue::String("Monday".into())
        );

        // No match and no default => #N/A
        assert_eq!(
            eval("=SWITCH(99,1,\"Sunday\",2,\"Monday\",3,\"Tuesday\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );

        // No match, default is "No match"
        assert_eq!(
            eval("=SWITCH(99,1,\"Sunday\",2,\"Monday\",3,\"Tuesday\",\"No match\")").unwrap(),
            FormulaValue::String("No match".into())
        );

        // No match for 1 or 7, default is "weekday"
        assert_eq!(
            eval("=SWITCH(2,1,\"Sunday\",7,\"Saturday\",\"weekday\")").unwrap(),
            FormulaValue::String("weekday".into())
        );

        // Matches value 3 => "Tuesday"
        assert_eq!(
            eval("=SWITCH(3,1,\"Sunday\",2,\"Monday\",3,\"Tuesday\",\"No match\")").unwrap(),
            FormulaValue::String("Tuesday".into())
        );
    }

    #[test]
    fn test_xor_docs() {
        // =XOR(3>0,2<9) => FALSE (both TRUE, even count)
        assert_eq!(eval("=XOR(3>0,2<9)").unwrap(), FormulaValue::Boolean(false));

        // =XOR(3>12,4>6) => FALSE (both FALSE, zero TRUE)
        assert_eq!(
            eval("=XOR(3>12,4>6)").unwrap(),
            FormulaValue::Boolean(false)
        );
    }

    #[test]
    fn test_true_docs() {
        assert_eq!(eval("=TRUE()").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=TRUE").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(
            eval("=IF(TRUE(),\"yes\",\"no\")").unwrap(),
            FormulaValue::String("yes".into()),
        );
    }

    #[test]
    fn test_false_docs() {
        assert_eq!(eval("=FALSE()").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=FALSE").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(
            eval("=IF(FALSE(),\"yes\",\"no\")").unwrap(),
            FormulaValue::String("no".into()),
        );
    }
}
