//! Math functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

/// SUM function
pub fn fn_sum(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut sum = 0.0;

    for arg in args {
        match arg {
            FormulaValue::Number(n) => sum += n,
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array { data: arr, .. } => {
                for row in arr {
                    for cell in row {
                        if let FormulaValue::Number(n) = cell {
                            sum += n;
                        } else if let FormulaValue::Error(e) = cell {
                            return Ok(FormulaValue::Error(*e));
                        }
                    }
                }
            }
            _ => {} // Ignore non-numeric
        }
    }

    Ok(FormulaValue::Number(sum))
}

/// AVERAGE function
pub fn fn_average(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut sum = 0.0;
    let mut count = 0;

    for arg in args {
        match arg {
            FormulaValue::Number(n) => {
                sum += n;
                count += 1;
            }
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array { data: arr, .. } => {
                for row in arr {
                    for cell in row {
                        if let FormulaValue::Number(n) = cell {
                            sum += n;
                            count += 1;
                        } else if let FormulaValue::Error(e) = cell {
                            return Ok(FormulaValue::Error(*e));
                        }
                    }
                }
            }
            _ => {} // Ignore non-numeric
        }
    }

    if count == 0 {
        Ok(FormulaValue::Error(CellError::Div0))
    } else {
        Ok(FormulaValue::Number(sum / count as f64))
    }
}

/// MIN function
pub fn fn_min(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut min: Option<f64> = None;

    for arg in args {
        match arg {
            FormulaValue::Number(n) => {
                min = Some(min.map_or(*n, |m| m.min(*n)));
            }
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array { data: arr, .. } => {
                for row in arr {
                    for cell in row {
                        if let FormulaValue::Number(n) = cell {
                            min = Some(min.map_or(*n, |m| m.min(*n)));
                        } else if let FormulaValue::Error(e) = cell {
                            return Ok(FormulaValue::Error(*e));
                        }
                    }
                }
            }
            _ => {} // Ignore non-numeric
        }
    }

    Ok(min.map_or(FormulaValue::Number(0.0), FormulaValue::Number))
}

/// MAX function
pub fn fn_max(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut max: Option<f64> = None;

    for arg in args {
        match arg {
            FormulaValue::Number(n) => {
                max = Some(max.map_or(*n, |m| m.max(*n)));
            }
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array { data: arr, .. } => {
                for row in arr {
                    for cell in row {
                        if let FormulaValue::Number(n) = cell {
                            max = Some(max.map_or(*n, |m| m.max(*n)));
                        } else if let FormulaValue::Error(e) = cell {
                            return Ok(FormulaValue::Error(*e));
                        }
                    }
                }
            }
            _ => {} // Ignore non-numeric
        }
    }

    Ok(max.map_or(FormulaValue::Number(0.0), FormulaValue::Number))
}

/// COUNT function
pub fn fn_count(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut count = 0;

    for arg in args {
        match arg {
            FormulaValue::Number(_) => count += 1,
            FormulaValue::Array { data: arr, .. } => {
                for row in arr {
                    for cell in row {
                        if matches!(cell, FormulaValue::Number(_)) {
                            count += 1;
                        }
                    }
                }
            }
            _ => {} // Don't count non-numeric
        }
    }

    Ok(FormulaValue::Number(count as f64))
}

/// RAND() - Returns a random number between 0 and 1
/// This is a volatile function that returns a different value on each calculation.
pub fn fn_rand(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    use rand::Rng;
    let mut rng = rand::thread_rng();
    Ok(FormulaValue::Number(rng.gen::<f64>()))
}

/// RANDBETWEEN(bottom, top) - Returns a random integer between bottom and top (inclusive)
/// This is a volatile function that returns a different value on each calculation.
pub fn fn_randbetween(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    use rand::Rng;

    let bottom = match args.first() {
        Some(FormulaValue::Number(n)) => n.ceil() as i64,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let top = match args.get(1) {
        Some(FormulaValue::Number(n)) => n.floor() as i64,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if bottom > top {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut rng = rand::thread_rng();
    let result = rng.gen_range(bottom..=top);
    Ok(FormulaValue::Number(result as f64))
}

/// ABS(number) - Returns the absolute value of a number
/// Reference: LibreOffice ScInterpreter::ScAbs
pub fn fn_abs(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(n.abs())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)), // Empty treated as 0
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// ROUND(number, [num_digits]) - Rounds a number to a specified number of digits
/// Uses "round half away from zero" mode (standard Excel/LibreOffice rounding)
/// Reference: LibreOffice ScInterpreter::ScRound with rtl_math_RoundingMode_Corrected
pub fn fn_round(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let num_digits = match args.get(1) {
        Some(FormulaValue::Number(n)) => *n as i32,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Handle rounding with the multiplier approach
    // For negative digits, we round to the left of the decimal point
    let multiplier = 10_f64.powi(num_digits);

    // Use "round half away from zero" - the standard Excel behavior
    // For positive numbers: round(2.5) = 3, round(2.4) = 2
    // For negative numbers: round(-2.5) = -3, round(-2.4) = -2
    let result = if number >= 0.0 {
        (number * multiplier + 0.5).floor() / multiplier
    } else {
        (number * multiplier - 0.5).ceil() / multiplier
    };

    Ok(FormulaValue::Number(result))
}

/// MOD(number, divisor) - Returns the remainder after division
/// Uses Excel/LibreOffice formula: number - divisor * floor(number/divisor)
/// The result has the same sign as the divisor (unlike Rust's % operator)
/// Reference: LibreOffice ScInterpreter::ScMod
pub fn fn_mod(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let divisor = match args.get(1) {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Division by zero
    if divisor == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }

    // Excel MOD formula: number - divisor * floor(number/divisor)
    // This ensures the result has the same sign as the divisor
    let result = number - divisor * (number / divisor).floor();

    // Validate the result is in expected range (matching LibreOffice behavior)
    let valid = if divisor > 0.0 {
        result >= 0.0 && result < divisor
    } else {
        result <= 0.0 && result > divisor
    };

    if valid {
        Ok(FormulaValue::Number(result))
    } else {
        // Edge case: floating point precision issues - return error
        Ok(FormulaValue::Error(CellError::Value))
    }
}

/// INT(number) - Rounds a number down to the nearest integer
/// Always rounds toward negative infinity (floor)
/// Reference: LibreOffice ScInterpreter::ScInt uses approxFloor
pub fn fn_int(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(n.floor())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)),
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// TRUNC(number, [num_digits]) - Truncates a number to an integer by removing the fractional part
/// Unlike INT, TRUNC rounds toward zero
/// Reference: LibreOffice ScInterpreter::ScTrunc
pub fn fn_trunc(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let num_digits = match args.get(1) {
        Some(FormulaValue::Number(n)) => *n as i32,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let multiplier = 10_f64.powi(num_digits);
    let result = (number * multiplier).trunc() / multiplier;
    Ok(FormulaValue::Number(result))
}

/// SIGN(number) - Returns the sign of a number: 1 if positive, -1 if negative, 0 if zero
/// Reference: LibreOffice ScInterpreter::ScSign
pub fn fn_sign(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => {
            let sign = if *n > 0.0 {
                1.0
            } else if *n < 0.0 {
                -1.0
            } else {
                0.0
            };
            Ok(FormulaValue::Number(sign))
        }
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)),
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// SQRT(number) - Returns the positive square root of a number
/// Returns #NUM! error for negative numbers
/// Reference: LibreOffice ScInterpreter::ScSqrt
pub fn fn_sqrt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => {
            if *n >= 0.0 {
                Ok(FormulaValue::Number(n.sqrt()))
            } else {
                Ok(FormulaValue::Error(CellError::Num))
            }
        }
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)),
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// POWER(number, power) - Returns the result of a number raised to a power
/// Equivalent to number^power
/// Reference: LibreOffice ScInterpreter::ScPow
pub fn fn_power(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let power = match args.get(1) {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if number == 0.0 && power == 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let result = number.powf(power);

    // Check for invalid results
    if result.is_nan() || result.is_infinite() {
        // Cases like 0^(-1) or negative^(non-integer)
        Ok(FormulaValue::Error(CellError::Num))
    } else {
        Ok(FormulaValue::Number(result))
    }
}

/// LOG(number, [base]) - Returns the logarithm of a number to a specified base
/// Default base is 10
/// Reference: LibreOffice ScInterpreter::ScLog
pub fn fn_log(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => return Ok(FormulaValue::Error(CellError::Num)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let base = match args.get(1) {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 10.0, // Default base is 10
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Validate inputs
    if number <= 0.0 || base <= 0.0 || base == 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(number.ln() / base.ln()))
}

/// LOG10(number) - Returns the base-10 logarithm of a number
/// Reference: LibreOffice ScInterpreter::ScLog10
pub fn fn_log10(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => {
            if *n > 0.0 {
                Ok(FormulaValue::Number(n.log10()))
            } else {
                Ok(FormulaValue::Error(CellError::Num))
            }
        }
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Error(CellError::Num)),
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// LN(number) - Returns the natural logarithm of a number
/// Reference: LibreOffice ScInterpreter::ScLn
pub fn fn_ln(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => {
            if *n > 0.0 {
                Ok(FormulaValue::Number(n.ln()))
            } else {
                Ok(FormulaValue::Error(CellError::Num))
            }
        }
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Error(CellError::Num)),
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// EXP(number) - Returns e raised to the power of a given number
/// Reference: LibreOffice ScInterpreter::ScExp
pub fn fn_exp(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => {
            let result = n.exp();
            if result.is_infinite() {
                Ok(FormulaValue::Error(CellError::Num))
            } else {
                Ok(FormulaValue::Number(result))
            }
        }
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(1.0)), // e^0 = 1
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// PI() - Returns the value of pi (3.14159265358979...)
pub fn fn_pi(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    Ok(FormulaValue::Number(std::f64::consts::PI))
}

/// SIN(number) - Returns the sine of an angle (in radians)
/// Reference: LibreOffice ScInterpreter::ScSin
pub fn fn_sin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(n.sin())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)), // sin(0) = 0
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// COS(number) - Returns the cosine of an angle (in radians)
/// Reference: LibreOffice ScInterpreter::ScCos
pub fn fn_cos(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(n.cos())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(1.0)), // cos(0) = 1
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// TAN(number) - Returns the tangent of an angle (in radians)
/// Reference: LibreOffice ScInterpreter::ScTan
pub fn fn_tan(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => {
            let result = n.tan();
            if result.is_infinite() || result.is_nan() {
                Ok(FormulaValue::Error(CellError::Num))
            } else {
                Ok(FormulaValue::Number(result))
            }
        }
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)), // tan(0) = 0
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// ASIN(number) - Returns the arcsine (inverse sine) of a number, in radians
/// The result is between -PI/2 and PI/2
/// Reference: LibreOffice ScInterpreter::ScArcSin
pub fn fn_asin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => {
            if *n < -1.0 || *n > 1.0 {
                Ok(FormulaValue::Error(CellError::Num))
            } else {
                Ok(FormulaValue::Number(n.asin()))
            }
        }
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)), // asin(0) = 0
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// ACOS(number) - Returns the arccosine (inverse cosine) of a number, in radians
/// The result is between 0 and PI
/// Reference: LibreOffice ScInterpreter::ScArcCos
pub fn fn_acos(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => {
            if *n < -1.0 || *n > 1.0 {
                Ok(FormulaValue::Error(CellError::Num))
            } else {
                Ok(FormulaValue::Number(n.acos()))
            }
        }
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(std::f64::consts::FRAC_PI_2)), // acos(0) = PI/2
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// ATAN(number) - Returns the arctangent (inverse tangent) of a number, in radians
/// The result is between -PI/2 and PI/2
/// Reference: LibreOffice ScInterpreter::ScArcTan
pub fn fn_atan(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(n.atan())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)), // atan(0) = 0
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// ATAN2(x_num, y_num) - Returns the arctangent from x and y coordinates
/// Returns angle in radians between -PI and PI
/// Note: Excel's ATAN2(x,y) is equivalent to math atan2(y,x)
/// Reference: LibreOffice ScInterpreter::ScArcTan2
pub fn fn_atan2(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let y = match args.get(1) {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if x == 0.0 && y == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }

    // Excel ATAN2(x,y) = atan2(y,x) in standard math notation
    Ok(FormulaValue::Number(y.atan2(x)))
}

/// DEGREES(angle) - Converts radians to degrees
/// Reference: LibreOffice ScInterpreter::ScDeg
pub fn fn_degrees(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(n.to_degrees())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)),
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// RADIANS(angle) - Converts degrees to radians
/// Reference: LibreOffice ScInterpreter::ScRad
pub fn fn_radians(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(n.to_radians())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => Ok(FormulaValue::Number(0.0)),
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// ROUNDUP(number, num_digits) - Rounds a number up, away from zero
/// Reference: LibreOffice ScInterpreter::ScRoundUp
pub fn fn_roundup(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let num_digits = match args.get(1) {
        Some(FormulaValue::Number(n)) => *n as i32,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let multiplier = 10_f64.powi(num_digits);

    // Round away from zero
    let result = if number >= 0.0 {
        (number * multiplier).ceil() / multiplier
    } else {
        (number * multiplier).floor() / multiplier
    };

    Ok(FormulaValue::Number(result))
}

/// ROUNDDOWN(number, num_digits) - Rounds a number down, toward zero
/// Reference: LibreOffice ScInterpreter::ScRoundDown
pub fn fn_rounddown(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let num_digits = match args.get(1) {
        Some(FormulaValue::Number(n)) => *n as i32,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let multiplier = 10_f64.powi(num_digits);

    // Round toward zero (truncate)
    let result = (number * multiplier).trunc() / multiplier;

    Ok(FormulaValue::Number(result))
}

/// CEILING.MATH(number, [significance], [mode]) - Rounds a number up to the nearest integer or multiple of significance
/// Reference: LibreOffice ScInterpreter::ScCeil_MS / ScCeil_Precise
pub fn fn_ceiling_math(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let significance = match args.get(1) {
        Some(FormulaValue::Number(n)) => {
            if *n == 0.0 {
                return Ok(FormulaValue::Number(0.0));
            }
            n.abs()
        }
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 1.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let mode = match args.get(2) {
        Some(FormulaValue::Number(n)) => *n != 0.0,
        Some(FormulaValue::Boolean(b)) => *b,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        _ => false,
    };

    // For negative numbers with mode=true, round away from zero (more negative)
    let result = if number < 0.0 && mode {
        -((-number / significance).ceil() * significance)
    } else {
        (number / significance).ceil() * significance
    };

    Ok(FormulaValue::Number(result))
}

/// FLOOR.MATH(number, [significance], [mode]) - Rounds a number down to the nearest integer or multiple of significance
/// Reference: LibreOffice ScInterpreter::ScFloor_MS / ScFloor_Precise
pub fn fn_floor_math(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let significance = match args.get(1) {
        Some(FormulaValue::Number(n)) => {
            if *n == 0.0 {
                return Ok(FormulaValue::Number(0.0));
            }
            n.abs()
        }
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 1.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let mode = match args.get(2) {
        Some(FormulaValue::Number(n)) => *n != 0.0,
        Some(FormulaValue::Boolean(b)) => *b,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        _ => false,
    };

    // For negative numbers with mode=true, round toward zero (less negative)
    let result = if number < 0.0 && mode {
        -((-number / significance).floor() * significance)
    } else {
        (number / significance).floor() * significance
    };

    Ok(FormulaValue::Number(result))
}

/// ODD(number) - Rounds a number up to the nearest odd integer
pub fn fn_odd(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if number == 0.0 {
        return Ok(FormulaValue::Number(1.0));
    }

    let sign = if number >= 0.0 { 1.0 } else { -1.0 };
    let abs_num = number.abs();
    let ceiling = abs_num.ceil();

    let result = if ceiling as i64 % 2 == 0 {
        ceiling + 1.0
    } else {
        ceiling
    };

    Ok(FormulaValue::Number(result * sign))
}

/// EVEN(number) - Rounds a number up to the nearest even integer
pub fn fn_even(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0.0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if number == 0.0 {
        return Ok(FormulaValue::Number(0.0));
    }

    let sign = if number >= 0.0 { 1.0 } else { -1.0 };
    let abs_num = number.abs();
    let ceiling = abs_num.ceil();

    let result = if ceiling as i64 % 2 == 1 {
        ceiling + 1.0
    } else {
        ceiling
    };

    Ok(FormulaValue::Number(result * sign))
}

/// SUMIF(range, criteria, [sum_range]) - Adds the cells specified by a given criteria
/// Reference: LibreOffice ScInterpreter::ScSumIf / IterateParametersIf
///
/// Criteria can be:
/// - A number: exact match (e.g., 5)
/// - A text string: case-insensitive match (e.g., "apple")
/// - A comparison expression: ">5", ">=10", "<100", "<=50", "<>0", "=5"
/// - Wildcards: "*" matches any characters, "?" matches single character
pub fn fn_sumif(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    use super::criteria::CriteriaMatcher;

    // Get the range (first argument)
    let range = match args.first() {
        Some(FormulaValue::Array { data: arr, .. }) => arr,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => {
            // Single value treated as 1x1 array
            return fn_sumif_single(v, args);
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Get the criteria (second argument)
    let criteria = match args.get(1) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Parse the criteria into a matcher
    let matcher = CriteriaMatcher::new(criteria);

    // Get sum_range (third argument) or use range
    let sum_range = match args.get(2) {
        Some(FormulaValue::Array { data: arr, .. }) => arr,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(_) | None => range, // If no sum_range or single value, use range
    };

    // Sum values where criteria matches
    let mut sum = 0.0;

    for (row_idx, row) in range.iter().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            // Check if cell matches criteria
            if matcher.matches(cell) {
                // Get corresponding cell from sum_range
                // If sum_range is smaller, use the same dimensions as range starting from top-left
                if let Some(sum_row) = sum_range.get(row_idx) {
                    if let Some(sum_cell) = sum_row.get(col_idx) {
                        if let FormulaValue::Number(n) = sum_cell {
                            sum += n;
                        } else if let FormulaValue::Error(e) = sum_cell {
                            return Ok(FormulaValue::Error(*e));
                        }
                        // Non-numeric values are ignored
                    }
                }
            }
        }
    }

    Ok(FormulaValue::Number(sum))
}

/// Handle SUMIF with single-value range
fn fn_sumif_single(value: &FormulaValue, args: &[FormulaValue]) -> FormulaResult<FormulaValue> {
    use super::criteria::CriteriaMatcher;

    let criteria = match args.get(1) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let matcher = CriteriaMatcher::new(criteria);

    // Get sum value (third arg or use value)
    let sum_value = match args.get(2) {
        Some(v) => v,
        None => value,
    };

    if matcher.matches(value) {
        match sum_value {
            FormulaValue::Number(n) => Ok(FormulaValue::Number(*n)),
            FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
            _ => Ok(FormulaValue::Number(0.0)),
        }
    } else {
        Ok(FormulaValue::Number(0.0))
    }
}

/// SUMPRODUCT(array1, [array2], [array3], ...) - Multiplies corresponding elements and sums the products
/// Reference: LibreOffice ScInterpreter::ScSumProduct, Microsoft SUMPRODUCT function
///
/// All arrays must have the same dimensions. Non-numeric values are treated as 0.
/// Returns #VALUE! if array dimensions don't match.
pub fn fn_sumproduct(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    if args.is_empty() {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Get dimensions from first array
    let first = &args[0];
    let (base_rows, base_cols, base_values) = match first {
        FormulaValue::Array { data: arr, .. } => {
            let rows = arr.len();
            let cols = arr.first().map(|r| r.len()).unwrap_or(0);
            if rows == 0 || cols == 0 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            // Flatten to values
            let values: Vec<f64> = arr
                .iter()
                .flat_map(|row| {
                    row.iter().map(|v| match v {
                        FormulaValue::Number(n) => *n,
                        FormulaValue::Boolean(true) => 1.0,
                        FormulaValue::Boolean(false) => 0.0,
                        FormulaValue::Error(_) => f64::NAN,
                        _ => 0.0, // Non-numeric treated as 0
                    })
                })
                .collect();
            (rows, cols, values)
        }
        FormulaValue::Number(n) => (1, 1, vec![*n]),
        FormulaValue::Boolean(b) => (1, 1, vec![if *b { 1.0 } else { 0.0 }]),
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
        FormulaValue::Empty => (1, 1, vec![0.0]),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Start with base values as the product accumulator
    let mut products = base_values;

    // Multiply with each subsequent array
    for arg in args.iter().skip(1) {
        let (rows, cols, values) = match arg {
            FormulaValue::Array { data: arr, .. } => {
                let rows = arr.len();
                let cols = arr.first().map(|r| r.len()).unwrap_or(0);
                let values: Vec<f64> = arr
                    .iter()
                    .flat_map(|row| {
                        row.iter().map(|v| match v {
                            FormulaValue::Number(n) => *n,
                            FormulaValue::Boolean(true) => 1.0,
                            FormulaValue::Boolean(false) => 0.0,
                            FormulaValue::Error(_) => f64::NAN,
                            _ => 0.0,
                        })
                    })
                    .collect();
                (rows, cols, values)
            }
            FormulaValue::Number(n) => (1, 1, vec![*n]),
            FormulaValue::Boolean(b) => (1, 1, vec![if *b { 1.0 } else { 0.0 }]),
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Empty => (1, 1, vec![0.0]),
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        };

        // Check dimensions match
        if rows != base_rows || cols != base_cols {
            return Ok(FormulaValue::Error(CellError::Value));
        }

        // Multiply element-wise
        for (i, val) in values.iter().enumerate() {
            products[i] *= val;
        }
    }

    // Sum all products, propagating any NaN (errors)
    let mut sum = 0.0;
    for p in &products {
        if p.is_nan() {
            // Error in array element - for now treat as 0 like Excel does for text
            // (actual errors should have been caught earlier)
            continue;
        }
        sum += p;
    }

    Ok(FormulaValue::Number(sum))
}

/// SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
/// Reference: LibreOffice ScInterpreter::ScSumIfs, Microsoft SUMIFS function
///
/// Sums cells in sum_range where ALL criteria are met.
/// All ranges must have the same dimensions.
pub fn fn_sumifs(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    use super::criteria::CriteriaMatcher;

    // Must have at least 3 arguments: sum_range, criteria_range1, criteria1
    if args.len() < 3 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Must have odd number of arguments (sum_range + pairs)
    if args.len() % 2 != 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Get sum_range (first argument)
    let sum_range = match &args[0] {
        FormulaValue::Array { data: arr, .. } => arr,
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
        v => {
            // Single value - treat as 1x1 array
            return fn_sumifs_single(v, &args[1..]);
        }
    };

    let (rows, cols) = array_dims(sum_range);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Number(0.0));
    }

    // Build criteria matchers from pairs
    let num_pairs = (args.len() - 1) / 2;
    let mut criteria_ranges: Vec<&Vec<Vec<FormulaValue>>> = Vec::with_capacity(num_pairs);
    let mut matchers: Vec<CriteriaMatcher> = Vec::with_capacity(num_pairs);

    for i in 0..num_pairs {
        let range_idx = 1 + i * 2;
        let criteria_idx = range_idx + 1;

        // Get criteria range
        let range = match &args[range_idx] {
            FormulaValue::Array { data: arr, .. } => arr,
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        };

        // Validate dimensions match
        let (r, c) = array_dims(range);
        if r != rows || c != cols {
            return Ok(FormulaValue::Error(CellError::Value));
        }

        criteria_ranges.push(range);

        // Get criteria and create matcher
        let criteria = &args[criteria_idx];
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }
        matchers.push(CriteriaMatcher::new(criteria));
    }

    // Sum values where ALL criteria match
    let mut sum = 0.0;

    for row_idx in 0..rows {
        for col_idx in 0..cols {
            // Check if all criteria match for this cell
            let mut all_match = true;
            for (range, matcher) in criteria_ranges.iter().zip(matchers.iter()) {
                let cell = &range[row_idx][col_idx];
                if !matcher.matches(cell) {
                    all_match = false;
                    break;
                }
            }

            if all_match {
                // Add the corresponding sum_range value
                let sum_cell = &sum_range[row_idx][col_idx];
                match sum_cell {
                    FormulaValue::Number(n) => sum += n,
                    FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                    _ => {} // Non-numeric ignored
                }
            }
        }
    }

    Ok(FormulaValue::Number(sum))
}

/// Helper for SUMIFS with single-value ranges
fn fn_sumifs_single(
    sum_value: &FormulaValue,
    criteria_args: &[FormulaValue],
) -> FormulaResult<FormulaValue> {
    use super::criteria::CriteriaMatcher;

    // Each pair: criteria_range, criteria
    // For single values, criteria_range must also be single value
    let num_pairs = criteria_args.len() / 2;

    for i in 0..num_pairs {
        let range_idx = i * 2;
        let criteria_idx = range_idx + 1;

        let range_value = &criteria_args[range_idx];
        let criteria = &criteria_args[criteria_idx];

        if let FormulaValue::Error(e) = range_value {
            return Ok(FormulaValue::Error(*e));
        }
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }

        let matcher = CriteriaMatcher::new(criteria);
        if !matcher.matches(range_value) {
            return Ok(FormulaValue::Number(0.0));
        }
    }

    // All criteria matched
    match sum_value {
        FormulaValue::Number(n) => Ok(FormulaValue::Number(*n)),
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::Number(0.0)),
    }
}

/// Helper to get array dimensions
fn array_dims(arr: &[Vec<FormulaValue>]) -> (usize, usize) {
    let rows = arr.len();
    let cols = arr.first().map(|r| r.len()).unwrap_or(0);
    (rows, cols)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::functions::EvaluationContext;

    fn num(v: FormulaValue) -> f64 {
        match v {
            FormulaValue::Number(n) => n,
            other => panic!("expected number, got {:?}", other),
        }
    }

    // Batch 1 docs tests: ABS, ROUND, ROUNDUP, ROUNDDOWN, TRUNC, SIGN

    #[test]
    fn test_abs_docs() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_abs(&[FormulaValue::Number(2.0)], &ctx).unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            fn_abs(&[FormulaValue::Number(-2.0)], &ctx).unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            fn_abs(&[FormulaValue::Number(-4.0)], &ctx).unwrap(),
            FormulaValue::Number(4.0)
        );
    }

    #[test]
    fn test_round_docs() {
        let ctx = EvaluationContext::simple();
        assert!(
            (num(fn_round(
                &[FormulaValue::Number(2.15), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap())
                - 2.2)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_round(
                &[FormulaValue::Number(2.149), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap())
                - 2.1)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_round(
                &[FormulaValue::Number(-1.475), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap())
                - (-1.48))
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_round(
                &[FormulaValue::Number(21.5), FormulaValue::Number(-1.0)],
                &ctx
            )
            .unwrap())
                - 20.0)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_round(
                &[FormulaValue::Number(626.3), FormulaValue::Number(-3.0)],
                &ctx
            )
            .unwrap())
                - 1000.0)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_round(
                &[FormulaValue::Number(1.98), FormulaValue::Number(-1.0)],
                &ctx
            )
            .unwrap())
                - 0.0)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_round(
                &[FormulaValue::Number(-50.55), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap())
                - (-100.0))
                .abs()
                < 1e-10
        );
    }

    #[test]
    fn test_roundup_docs() {
        let ctx = EvaluationContext::simple();
        assert!(
            (num(fn_roundup(
                &[FormulaValue::Number(3.2), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap())
                - 4.0)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_roundup(
                &[FormulaValue::Number(76.9), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap())
                - 77.0)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_roundup(
                &[FormulaValue::Number(3.14159), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap())
                - 3.142)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_roundup(
                &[FormulaValue::Number(-3.14159), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap())
                - (-3.2))
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_roundup(
                &[
                    FormulaValue::Number(31415.92654),
                    FormulaValue::Number(-2.0)
                ],
                &ctx
            )
            .unwrap())
                - 31500.0)
                .abs()
                < 1e-10
        );
    }

    #[test]
    fn test_rounddown_docs() {
        let ctx = EvaluationContext::simple();
        assert!(
            (num(fn_rounddown(
                &[FormulaValue::Number(3.2), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap())
                - 3.0)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_rounddown(
                &[FormulaValue::Number(76.9), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap())
                - 76.0)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_rounddown(
                &[FormulaValue::Number(3.14159), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap())
                - 3.141)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_rounddown(
                &[FormulaValue::Number(-3.14159), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap())
                - (-3.1))
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_rounddown(
                &[
                    FormulaValue::Number(31415.92654),
                    FormulaValue::Number(-2.0)
                ],
                &ctx
            )
            .unwrap())
                - 31400.0)
                .abs()
                < 1e-10
        );
    }

    #[test]
    fn test_trunc_docs() {
        let ctx = EvaluationContext::simple();
        assert!(
            (num(fn_trunc(
                &[FormulaValue::Number(8.9), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap())
                - 8.0)
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_trunc(
                &[FormulaValue::Number(-8.9), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap())
                - (-8.0))
                .abs()
                < 1e-10
        );
        assert!(
            (num(fn_trunc(
                &[FormulaValue::Number(0.45), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap())
                - 0.0)
                .abs()
                < 1e-10
        );
    }

    #[test]
    fn test_sign_docs() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_sign(&[FormulaValue::Number(10.0)], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            fn_sign(&[FormulaValue::Number(0.0)], &ctx).unwrap(),
            FormulaValue::Number(0.0)
        );
        assert_eq!(
            fn_sign(&[FormulaValue::Number(-0.00001)], &ctx).unwrap(),
            FormulaValue::Number(-1.0)
        );
    }

    // Batch 2 docs tests: MOD, INT, EVEN, ODD, POWER, SQRT

    #[test]
    fn test_mod_docs() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_mod(
                &[FormulaValue::Number(3.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            fn_mod(
                &[FormulaValue::Number(-3.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            fn_mod(
                &[FormulaValue::Number(3.0), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-1.0)
        );
        assert_eq!(
            fn_mod(
                &[FormulaValue::Number(-3.0), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-1.0)
        );
    }

    #[test]
    fn test_int_docs() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_int(&[FormulaValue::Number(8.9)], &ctx).unwrap(),
            FormulaValue::Number(8.0)
        );
        assert_eq!(
            fn_int(&[FormulaValue::Number(-8.9)], &ctx).unwrap(),
            FormulaValue::Number(-9.0)
        );
        assert_eq!(
            fn_int(&[FormulaValue::Number(19.5)], &ctx).unwrap(),
            FormulaValue::Number(19.0)
        );
    }

    #[test]
    fn test_even_docs() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_even(&[FormulaValue::Number(1.5)], &ctx).unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            fn_even(&[FormulaValue::Number(3.0)], &ctx).unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            fn_even(&[FormulaValue::Number(2.0)], &ctx).unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            fn_even(&[FormulaValue::Number(-1.0)], &ctx).unwrap(),
            FormulaValue::Number(-2.0)
        );
    }

    #[test]
    fn test_odd_docs() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_odd(&[FormulaValue::Number(1.5)], &ctx).unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            fn_odd(&[FormulaValue::Number(3.0)], &ctx).unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            fn_odd(&[FormulaValue::Number(2.0)], &ctx).unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            fn_odd(&[FormulaValue::Number(-1.0)], &ctx).unwrap(),
            FormulaValue::Number(-1.0)
        );
        assert_eq!(
            fn_odd(&[FormulaValue::Number(-2.0)], &ctx).unwrap(),
            FormulaValue::Number(-3.0)
        );
    }

    #[test]
    fn test_power_docs() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_power(
                &[FormulaValue::Number(5.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(25.0)
        );
        assert!(
            (num(fn_power(
                &[FormulaValue::Number(98.6), FormulaValue::Number(3.2)],
                &ctx
            )
            .unwrap())
                - 2401077.222)
                .abs()
                < 0.001
        );
        assert!(
            (num(fn_power(
                &[FormulaValue::Number(4.0), FormulaValue::Number(1.25)],
                &ctx
            )
            .unwrap())
                - 5.656854249)
                .abs()
                < 1e-9
        );
        assert_eq!(
            fn_power(
                &[FormulaValue::Number(0.0), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_sqrt_docs() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_sqrt(&[FormulaValue::Number(16.0)], &ctx).unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            fn_sqrt(&[FormulaValue::Number(-16.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    // Batch 3 docs tests: SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2

    #[test]
    fn test_sin_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_sin(&[FormulaValue::Number(std::f64::consts::PI)], &ctx).unwrap());
        assert!((v - 0.0).abs() < 1e-10);
        let v = num(fn_sin(&[FormulaValue::Number(std::f64::consts::FRAC_PI_2)], &ctx).unwrap());
        assert!((v - 1.0).abs() < 1e-10);
        let v = num(fn_sin(
            &[FormulaValue::Number(30.0 * std::f64::consts::PI / 180.0)],
            &ctx,
        )
        .unwrap());
        assert!((v - 0.5).abs() < 1e-10);
    }

    #[test]
    fn test_cos_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_cos(&[FormulaValue::Number(1.047)], &ctx).unwrap());
        assert!((v - 0.5001711).abs() < 1e-6);
        let v = num(fn_cos(
            &[FormulaValue::Number(60.0 * std::f64::consts::PI / 180.0)],
            &ctx,
        )
        .unwrap());
        assert!((v - 0.5).abs() < 1e-10);
    }

    #[test]
    fn test_tan_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_tan(&[FormulaValue::Number(0.785)], &ctx).unwrap());
        assert!((v - 0.99920).abs() < 1e-4);
        let v = num(fn_tan(
            &[FormulaValue::Number(45.0 * std::f64::consts::PI / 180.0)],
            &ctx,
        )
        .unwrap());
        assert!((v - 1.0).abs() < 1e-10);
    }

    #[test]
    fn test_asin_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_asin(&[FormulaValue::Number(-0.5)], &ctx).unwrap());
        assert!((v - (-0.523598776)).abs() < 1e-8);
        assert!((v * 180.0 / std::f64::consts::PI - (-30.0)).abs() < 1e-10);
    }

    #[test]
    fn test_acos_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_acos(&[FormulaValue::Number(-0.5)], &ctx).unwrap());
        assert!((v - 2.094395102).abs() < 1e-8);
        assert!((v * 180.0 / std::f64::consts::PI - 120.0).abs() < 1e-10);
    }

    #[test]
    fn test_atan_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_atan(&[FormulaValue::Number(1.0)], &ctx).unwrap());
        assert!((v - 0.785398163).abs() < 1e-8);
        assert!((v * 180.0 / std::f64::consts::PI - 45.0).abs() < 1e-10);
    }

    #[test]
    fn test_atan2_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_atan2(
            &[FormulaValue::Number(1.0), FormulaValue::Number(1.0)],
            &ctx,
        )
        .unwrap());
        assert!((v - 0.785398163).abs() < 1e-8);
        let v = num(fn_atan2(
            &[FormulaValue::Number(-1.0), FormulaValue::Number(-1.0)],
            &ctx,
        )
        .unwrap());
        assert!((v - (-2.35619449)).abs() < 1e-7);
        assert!((v * 180.0 / std::f64::consts::PI - (-135.0)).abs() < 1e-10);
    }

    // Batch 4 docs tests: DEGREES, RADIANS, EXP, LN, LOG, LOG10, PI

    #[test]
    fn test_degrees_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_degrees(&[FormulaValue::Number(std::f64::consts::PI)], &ctx).unwrap());
        assert!((v - 180.0).abs() < 1e-10);
    }

    #[test]
    fn test_radians_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_radians(&[FormulaValue::Number(270.0)], &ctx).unwrap());
        assert!((v - std::f64::consts::PI * 1.5).abs() < 1e-10);
    }

    #[test]
    fn test_exp_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_exp(&[FormulaValue::Number(1.0)], &ctx).unwrap());
        assert!((v - std::f64::consts::E).abs() < 1e-10);
        let v = num(fn_exp(&[FormulaValue::Number(2.0)], &ctx).unwrap());
        assert!((v - 2.0_f64.exp()).abs() < 1e-10);
    }

    #[test]
    fn test_ln_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_ln(&[FormulaValue::Number(86.0)], &ctx).unwrap());
        assert!((v - 86.0_f64.ln()).abs() < 1e-10);
        let v = num(fn_ln(&[FormulaValue::Number(2.7182818)], &ctx).unwrap());
        assert!((v - 2.7182818_f64.ln()).abs() < 1e-10);
        let exp3 = fn_exp(&[FormulaValue::Number(3.0)], &ctx).unwrap();
        let v = num(fn_ln(&[exp3], &ctx).unwrap());
        assert!((v - 3.0).abs() < 1e-10);
    }

    #[test]
    fn test_log_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_log(&[FormulaValue::Number(10.0)], &ctx).unwrap());
        assert!((v - 1.0).abs() < 1e-10);
        let v = num(fn_log(
            &[FormulaValue::Number(8.0), FormulaValue::Number(2.0)],
            &ctx,
        )
        .unwrap());
        assert!((v - 3.0).abs() < 1e-10);
        let v = num(fn_log(
            &[FormulaValue::Number(86.0), FormulaValue::Number(2.7182818)],
            &ctx,
        )
        .unwrap());
        assert!((v - 86.0_f64.ln() / 2.7182818_f64.ln()).abs() < 1e-10);
    }

    #[test]
    fn test_log10_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_log10(&[FormulaValue::Number(86.0)], &ctx).unwrap());
        assert!((v - 86.0_f64.log10()).abs() < 1e-10);
        let v = num(fn_log10(&[FormulaValue::Number(10.0)], &ctx).unwrap());
        assert!((v - 1.0).abs() < 1e-10);
        let v = num(fn_log10(&[FormulaValue::Number(100000.0)], &ctx).unwrap());
        assert!((v - 5.0).abs() < 1e-10);
    }

    #[test]
    fn test_pi_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_pi(&[], &ctx).unwrap());
        assert!((v - std::f64::consts::PI).abs() < 1e-10);
    }

    // Batch 5 docs tests: CEILING.MATH, FLOOR.MATH, SUM, SUMIF

    #[test]
    fn test_ceiling_math_docs() {
        let ctx = EvaluationContext::simple();
        // =CEILING.MATH(24.3, 5) = 25
        assert_eq!(
            fn_ceiling_math(
                &[FormulaValue::Number(24.3), FormulaValue::Number(5.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(25.0)
        );
        // =CEILING.MATH(6.7) = 7
        assert_eq!(
            fn_ceiling_math(&[FormulaValue::Number(6.7)], &ctx).unwrap(),
            FormulaValue::Number(7.0)
        );
        // =CEILING.MATH(-8.1, 2) = -8
        assert_eq!(
            fn_ceiling_math(
                &[FormulaValue::Number(-8.1), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-8.0)
        );
        // =CEILING.MATH(-5.5, 2, -1) = -6
        assert_eq!(
            fn_ceiling_math(
                &[
                    FormulaValue::Number(-5.5),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(-1.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-6.0)
        );
    }

    #[test]
    fn test_floor_math_docs() {
        let ctx = EvaluationContext::simple();
        // =FLOOR.MATH(24.3, 5) = 20
        assert_eq!(
            fn_floor_math(
                &[FormulaValue::Number(24.3), FormulaValue::Number(5.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(20.0)
        );
        // =FLOOR.MATH(6.7) = 6
        assert_eq!(
            fn_floor_math(&[FormulaValue::Number(6.7)], &ctx).unwrap(),
            FormulaValue::Number(6.0)
        );
        // =FLOOR.MATH(-8.1, 2) = -10
        assert_eq!(
            fn_floor_math(
                &[FormulaValue::Number(-8.1), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-10.0)
        );
        // =FLOOR.MATH(-5.5, 2, -1) = -4
        assert_eq!(
            fn_floor_math(
                &[
                    FormulaValue::Number(-5.5),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(-1.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-4.0)
        );
    }

    #[test]
    fn test_sum_docs() {
        let ctx = EvaluationContext::simple();
        // =SUM(2, 3) = 5
        assert_eq!(
            fn_sum(
                &[FormulaValue::Number(2.0), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(5.0)
        );
        // =SUM(14598.93, 65437.90, 78496.23) = 158533.06
        assert!(
            (num(fn_sum(
                &[
                    FormulaValue::Number(14598.93),
                    FormulaValue::Number(65437.90),
                    FormulaValue::Number(78496.23)
                ],
                &ctx
            )
            .unwrap())
                - 158533.06)
                .abs()
                < 1e-10
        );
        // =SUM(array) sums all elements
        let arr = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::Number(1.0), FormulaValue::Number(2.0)],
                vec![FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
            ],
            source: None,
        };
        assert_eq!(fn_sum(&[arr], &ctx).unwrap(), FormulaValue::Number(10.0));
    }

    #[test]
    fn test_sumif_docs() {
        let ctx = EvaluationContext::simple();
        // Docs data: property values 100000, 200000, 300000, 400000
        // commissions: 7000, 14000, 21000, 28000
        let values = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::Number(100000.0)],
                vec![FormulaValue::Number(200000.0)],
                vec![FormulaValue::Number(300000.0)],
                vec![FormulaValue::Number(400000.0)],
            ],
            source: None,
        };
        let commissions = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::Number(7000.0)],
                vec![FormulaValue::Number(14000.0)],
                vec![FormulaValue::Number(21000.0)],
                vec![FormulaValue::Number(28000.0)],
            ],
            source: None,
        };
        // =SUMIF(values, ">160000", commissions) = 63000
        assert_eq!(
            num(fn_sumif(
                &[
                    values.clone(),
                    FormulaValue::String(">160000".to_string()),
                    commissions.clone()
                ],
                &ctx
            )
            .unwrap()),
            63000.0
        );
        // =SUMIF(values, ">160000") = 900000 (sum of values > 160000)
        assert_eq!(
            num(fn_sumif(
                &[values.clone(), FormulaValue::String(">160000".to_string())],
                &ctx
            )
            .unwrap()),
            900000.0
        );
        // =SUMIF(values, 300000, commissions) = 21000
        assert_eq!(
            num(fn_sumif(&[values, FormulaValue::Number(300000.0), commissions], &ctx).unwrap()),
            21000.0
        );
    }

    // Batch 6 docs tests: SUMIFS, SUMPRODUCT

    #[test]
    fn test_sumifs_docs() {
        let ctx = EvaluationContext::simple();
        // Docs data: 8 rows of quantity/product/salesperson
        let sum_range = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::Number(5.0)],
                vec![FormulaValue::Number(4.0)],
                vec![FormulaValue::Number(15.0)],
                vec![FormulaValue::Number(3.0)],
                vec![FormulaValue::Number(22.0)],
                vec![FormulaValue::Number(12.0)],
                vec![FormulaValue::Number(10.0)],
                vec![FormulaValue::Number(33.0)],
            ],
            source: None,
        };
        let products = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::String("Apples".to_string())],
                vec![FormulaValue::String("Apples".to_string())],
                vec![FormulaValue::String("Artichokes".to_string())],
                vec![FormulaValue::String("Artichokes".to_string())],
                vec![FormulaValue::String("Bananas".to_string())],
                vec![FormulaValue::String("Bananas".to_string())],
                vec![FormulaValue::String("Carrots".to_string())],
                vec![FormulaValue::String("Carrots".to_string())],
            ],
            source: None,
        };
        let salesperson = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::String("Tom".to_string())],
                vec![FormulaValue::String("Sarah".to_string())],
                vec![FormulaValue::String("Tom".to_string())],
                vec![FormulaValue::String("Sarah".to_string())],
                vec![FormulaValue::String("Tom".to_string())],
                vec![FormulaValue::String("Sarah".to_string())],
                vec![FormulaValue::String("Tom".to_string())],
                vec![FormulaValue::String("Sarah".to_string())],
            ],
            source: None,
        };
        // Docs Ex1: =SUMIFS(A2:A9, B2:B9, "=A*", C2:C9, "Tom") = 20
        // Products starting with A sold by Tom: Apples(5) + Artichokes(15)
        assert_eq!(
            num(fn_sumifs(
                &[
                    sum_range.clone(),
                    products.clone(),
                    FormulaValue::String("a*".to_string()),
                    salesperson.clone(),
                    FormulaValue::String("Tom".to_string())
                ],
                &ctx
            )
            .unwrap()),
            20.0
        );
        // All Tom's sales: 5 + 15 + 22 + 10 = 52
        assert_eq!(
            num(fn_sumifs(
                &[
                    sum_range,
                    salesperson,
                    FormulaValue::String("Tom".to_string())
                ],
                &ctx
            )
            .unwrap()),
            52.0
        );
    }

    #[test]
    fn test_sumproduct_docs() {
        let ctx = EvaluationContext::simple();
        // SUMPRODUCT({3,4;8,6;1,9}, {2,7;6,7;5,3}) = 156
        let a = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
                vec![FormulaValue::Number(8.0), FormulaValue::Number(6.0)],
                vec![FormulaValue::Number(1.0), FormulaValue::Number(9.0)],
            ],
            source: None,
        };
        let b = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::Number(2.0), FormulaValue::Number(7.0)],
                vec![FormulaValue::Number(6.0), FormulaValue::Number(7.0)],
                vec![FormulaValue::Number(5.0), FormulaValue::Number(3.0)],
            ],
            source: None,
        };
        assert_eq!(num(fn_sumproduct(&[a.clone(), b], &ctx).unwrap()), 156.0);
        // Single array: sums all elements = 3+4+8+6+1+9 = 31
        assert_eq!(num(fn_sumproduct(&[a], &ctx).unwrap()), 31.0);
    }

    #[test]
    fn test_sumproduct_docs_dimension_mismatch() {
        let ctx = EvaluationContext::simple();
        let a = FormulaValue::Array {
            data: vec![vec![FormulaValue::Number(1.0), FormulaValue::Number(2.0)]],
            source: None,
        };
        let b = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::Number(1.0)],
                vec![FormulaValue::Number(2.0)],
            ],
            source: None,
        };
        assert_eq!(
            fn_sumproduct(&[a, b], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    // Batch 14 docs tests: RAND, RANDBETWEEN

    #[test]
    fn test_rand_docs() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_rand(&[], &ctx).unwrap());
        assert!(v >= 0.0 && v < 1.0);
    }

    #[test]
    fn test_randbetween_docs() {
        let ctx = EvaluationContext::simple();
        // =RANDBETWEEN(1, 100) -> integer in [1, 100]
        let v = num(fn_randbetween(
            &[FormulaValue::Number(1.0), FormulaValue::Number(100.0)],
            &ctx,
        )
        .unwrap());
        assert!(v >= 1.0 && v <= 100.0);
        assert_eq!(v, v.floor()); // must be integer
                                  // =RANDBETWEEN(-1, 1) -> integer in [-1, 1]
        let v = num(fn_randbetween(
            &[FormulaValue::Number(-1.0), FormulaValue::Number(1.0)],
            &ctx,
        )
        .unwrap());
        assert!(v >= -1.0 && v <= 1.0);
        assert_eq!(v, v.floor());
        // bottom > top -> #NUM!
        assert_eq!(
            fn_randbetween(
                &[FormulaValue::Number(100.0), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }
}
