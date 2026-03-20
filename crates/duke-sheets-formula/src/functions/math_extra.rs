//! Additional math and trigonometry functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

fn get_number(arg: Option<&FormulaValue>) -> Result<f64, FormulaValue> {
    match arg {
        Some(FormulaValue::Number(n)) => Ok(*n),
        Some(FormulaValue::Empty) => Ok(0.0),
        Some(FormulaValue::Error(e)) => Err(FormulaValue::Error(*e)),
        _ => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn to_int_trunc(n: f64) -> Result<i64, FormulaValue> {
    if !n.is_finite() {
        return Err(FormulaValue::Error(CellError::Num));
    }
    let t = n.trunc();
    if t < i64::MIN as f64 || t > i64::MAX as f64 {
        return Err(FormulaValue::Error(CellError::Num));
    }
    Ok(t as i64)
}

fn collect_numbers(value: &FormulaValue, out: &mut Vec<f64>) -> Option<CellError> {
    match value {
        FormulaValue::Number(n) => out.push(*n),
        FormulaValue::Array { data: arr, .. } => {
            for row in arr {
                for cell in row {
                    match cell {
                        FormulaValue::Number(n) => out.push(*n),
                        FormulaValue::Error(e) => return Some(*e),
                        _ => {}
                    }
                }
            }
        }
        FormulaValue::Error(e) => return Some(*e),
        _ => {}
    }
    None
}

fn collect_numbers_skip_errors(value: &FormulaValue, out: &mut Vec<f64>) {
    match value {
        FormulaValue::Number(n) => out.push(*n),
        FormulaValue::Array { data: arr, .. } => {
            for row in arr {
                for cell in row {
                    if let FormulaValue::Number(n) = cell {
                        out.push(*n);
                    }
                }
            }
        }
        _ => {}
    }
}

fn collect_values(value: &FormulaValue, out: &mut Vec<FormulaValue>) -> Option<CellError> {
    match value {
        FormulaValue::Array { data: arr, .. } => {
            for row in arr {
                for cell in row {
                    if let FormulaValue::Error(e) = cell {
                        return Some(*e);
                    }
                    out.push(cell.clone());
                }
            }
        }
        FormulaValue::Error(e) => return Some(*e),
        _ => out.push(value.clone()),
    }
    None
}

fn gcd_i64(mut a: i64, mut b: i64) -> i64 {
    a = a.abs();
    b = b.abs();
    while b != 0 {
        let r = a % b;
        a = b;
        b = r;
    }
    a
}

fn lcm_i64(a: i64, b: i64) -> i64 {
    if a == 0 || b == 0 {
        0
    } else {
        (a / gcd_i64(a, b)) * b
    }
}

fn factorial(n: i64) -> Option<f64> {
    if n < 0 {
        return None;
    }
    let mut result = 1.0;
    for i in 1..=n {
        result *= i as f64;
        if !result.is_finite() {
            return None;
        }
    }
    Some(result)
}

fn parse_matrix(value: &FormulaValue) -> Result<Vec<Vec<f64>>, FormulaValue> {
    match value {
        FormulaValue::Array { data: rows, .. } => {
            if rows.is_empty() {
                return Err(FormulaValue::Error(CellError::Value));
            }
            let cols = rows[0].len();
            if cols == 0 {
                return Err(FormulaValue::Error(CellError::Value));
            }
            let mut out = Vec::with_capacity(rows.len());
            for row in rows {
                if row.len() != cols {
                    return Err(FormulaValue::Error(CellError::Value));
                }
                let mut out_row = Vec::with_capacity(cols);
                for cell in row {
                    match cell {
                        FormulaValue::Number(n) => out_row.push(*n),
                        FormulaValue::Error(e) => return Err(FormulaValue::Error(*e)),
                        _ => return Err(FormulaValue::Error(CellError::Value)),
                    }
                }
                out.push(out_row);
            }
            Ok(out)
        }
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        _ => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn percentile_inc(sorted: &[f64], k: f64) -> Option<f64> {
    if sorted.is_empty() || !(0.0..=1.0).contains(&k) {
        return None;
    }
    if sorted.len() == 1 {
        return Some(sorted[0]);
    }
    let n = sorted.len() as f64;
    let pos = (n - 1.0) * k;
    let lo = pos.floor() as usize;
    let hi = pos.ceil() as usize;
    if lo == hi {
        Some(sorted[lo])
    } else {
        let frac = pos - lo as f64;
        Some(sorted[lo] + frac * (sorted[hi] - sorted[lo]))
    }
}

fn percentile_exc(sorted: &[f64], k: f64) -> Option<f64> {
    if sorted.is_empty() {
        return None;
    }
    let n = sorted.len() as f64;
    let low = 1.0 / (n + 1.0);
    let high = n / (n + 1.0);
    if k <= low || k >= high {
        return None;
    }
    let pos = k * (n + 1.0);
    let lo = pos.floor() as usize;
    let hi = pos.ceil() as usize;
    if lo == hi {
        Some(sorted[lo - 1])
    } else {
        let frac = pos - lo as f64;
        let a = sorted[lo - 1];
        let b = sorted[hi - 1];
        Some(a + frac * (b - a))
    }
}

fn stdev_s(values: &[f64]) -> Option<f64> {
    if values.len() < 2 {
        return None;
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let var = values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (values.len() - 1) as f64;
    Some(var.sqrt())
}

fn stdev_p(values: &[f64]) -> Option<f64> {
    if values.is_empty() {
        return None;
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let var = values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / values.len() as f64;
    Some(var.sqrt())
}

fn var_s(values: &[f64]) -> Option<f64> {
    if values.len() < 2 {
        return None;
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    Some(values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (values.len() - 1) as f64)
}

fn var_p(values: &[f64]) -> Option<f64> {
    if values.is_empty() {
        return None;
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    Some(values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / values.len() as f64)
}

fn mode_sngl(values: &[f64]) -> Option<f64> {
    use std::collections::HashMap;
    if values.is_empty() {
        return None;
    }
    let mut counts: HashMap<i64, usize> = HashMap::new();
    for v in values {
        let key = (*v * 1e12).round() as i64;
        *counts.entry(key).or_insert(0) += 1;
    }
    let mut best_key = None;
    let mut best_count = 1usize;
    for (k, c) in counts {
        if c > best_count {
            best_count = c;
            best_key = Some(k);
        }
    }
    best_key.map(|k| k as f64 / 1e12)
}

fn flatten_array_like(value: &FormulaValue) -> Result<Vec<f64>, FormulaValue> {
    match value {
        FormulaValue::Number(n) => Ok(vec![*n]),
        FormulaValue::Array { data: rows, .. } => {
            let mut out = Vec::new();
            for row in rows {
                for cell in row {
                    match cell {
                        FormulaValue::Number(n) => out.push(*n),
                        FormulaValue::Error(e) => return Err(FormulaValue::Error(*e)),
                        _ => return Err(FormulaValue::Error(CellError::Value)),
                    }
                }
            }
            Ok(out)
        }
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        _ => Err(FormulaValue::Error(CellError::Value)),
    }
}

pub fn fn_acosh(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x < 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(x.acosh()))
}

pub fn fn_asinh(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(x.asinh()))
}

pub fn fn_atanh(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x <= -1.0 || x >= 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(x.atanh()))
}

/// ACOTH(number) — inverse hyperbolic cotangent
pub fn fn_acoth(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    // acoth(x) = atanh(1/x), domain: |x| > 1
    if x.abs() <= 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number((1.0 / x).atanh()))
}

/// ACOT(number) — inverse cotangent (arccotangent)
pub fn fn_acot(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    // acot(x) = atan(1/x), but acot(0) = PI/2
    let result = if x == 0.0 {
        std::f64::consts::FRAC_PI_2
    } else {
        let v = (1.0 / x).atan();
        // Ensure result is in range [0, pi]
        if v < 0.0 {
            v + std::f64::consts::PI
        } else {
            v
        }
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_cosh(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(x.cosh()))
}

pub fn fn_sinh(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(x.sinh()))
}

pub fn fn_tanh(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(x.tanh()))
}

pub fn fn_cot(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let t = x.tan();
    if t.abs() < 1e-15 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number(1.0 / t))
}

pub fn fn_coth(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let t = x.tanh();
    if t.abs() < 1e-15 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number(1.0 / t))
}

pub fn fn_csc(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let s = x.sin();
    if s.abs() < 1e-15 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number(1.0 / s))
}

pub fn fn_csch(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let s = x.sinh();
    if s.abs() < 1e-15 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number(1.0 / s))
}

pub fn fn_sec(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let c = x.cos();
    if c.abs() < 1e-15 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number(1.0 / c))
}

pub fn fn_sech(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(1.0 / x.cosh()))
}

pub fn fn_combin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let k = match get_number(args.get(1)).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if n < 0 || k < 0 || k > n {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let k = k.min(n - k);
    let mut result = 1.0;
    for i in 1..=k {
        result = result * (n - k + i) as f64 / i as f64;
        if !result.is_finite() {
            return Ok(FormulaValue::Error(CellError::Num));
        }
    }
    Ok(FormulaValue::Number(result))
}

pub fn fn_combina(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let k = match get_number(args.get(1)).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if n < 0 || k < 0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    if n == 0 {
        return Ok(FormulaValue::Number(if k == 0 { 1.0 } else { 0.0 }));
    }
    fn_combin(
        &[
            FormulaValue::Number((n + k - 1) as f64),
            FormulaValue::Number(k as f64),
        ],
        _ctx,
    )
}

pub fn fn_fact(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    match factorial(n) {
        Some(v) => Ok(FormulaValue::Number(v)),
        None => Ok(FormulaValue::Error(CellError::Num)),
    }
}

pub fn fn_factdouble(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let n = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if n < 0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let mut result = 1.0;
    let mut i = if n == 0 { 1 } else { n };
    while i > 1 {
        result *= i as f64;
        if !result.is_finite() {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        i -= 2;
    }
    Ok(FormulaValue::Number(result))
}

pub fn fn_multinomial(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    if args.is_empty() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let mut ints = Vec::new();
    let mut sum = 0i64;
    for arg in args {
        match arg {
            FormulaValue::Array { data: rows, .. } => {
                for row in rows {
                    for cell in row {
                        match cell {
                            FormulaValue::Number(n) => {
                                let i = match to_int_trunc(*n) {
                                    Ok(v) => v,
                                    Err(e) => return Ok(e),
                                };
                                if i < 0 {
                                    return Ok(FormulaValue::Error(CellError::Num));
                                }
                                ints.push(i);
                                sum += i;
                            }
                            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                            _ => return Ok(FormulaValue::Error(CellError::Value)),
                        }
                    }
                }
            }
            FormulaValue::Number(n) => {
                let i = match to_int_trunc(*n) {
                    Ok(v) => v,
                    Err(e) => return Ok(e),
                };
                if i < 0 {
                    return Ok(FormulaValue::Error(CellError::Num));
                }
                ints.push(i);
                sum += i;
            }
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        }
    }
    let mut result = match factorial(sum) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    for i in ints {
        let f = match factorial(i) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        result /= f;
    }
    Ok(FormulaValue::Number(result))
}

pub fn fn_gcd(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.is_empty() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let mut values = Vec::new();
    for arg in args {
        if let Some(e) = collect_numbers(arg, &mut values) {
            return Ok(FormulaValue::Error(e));
        }
    }
    if values.is_empty() {
        return Ok(FormulaValue::Number(0.0));
    }
    let mut g = 0i64;
    for v in values {
        let i = match to_int_trunc(v) {
            Ok(i) => i,
            Err(e) => return Ok(e),
        };
        g = gcd_i64(g, i);
    }
    Ok(FormulaValue::Number(g as f64))
}

pub fn fn_lcm(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.is_empty() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let mut values = Vec::new();
    for arg in args {
        if let Some(e) = collect_numbers(arg, &mut values) {
            return Ok(FormulaValue::Error(e));
        }
    }
    if values.is_empty() {
        return Ok(FormulaValue::Number(0.0));
    }
    let mut l = 1i64;
    for v in values {
        let i = match to_int_trunc(v) {
            Ok(i) => i.abs(),
            Err(e) => return Ok(e),
        };
        l = lcm_i64(l, i);
    }
    Ok(FormulaValue::Number(l as f64))
}

pub fn fn_product(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut values = Vec::new();
    for arg in args {
        if let Some(e) = collect_numbers(arg, &mut values) {
            return Ok(FormulaValue::Error(e));
        }
    }
    if values.is_empty() {
        return Ok(FormulaValue::Number(0.0));
    }
    let mut product = 1.0;
    for v in values {
        product *= v;
    }
    Ok(FormulaValue::Number(product))
}

pub fn fn_quotient(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let num = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let den = match get_number(args.get(1)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if den == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number((num / den).trunc()))
}

pub fn fn_mround(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let num = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let multiple = match get_number(args.get(1)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if multiple == 0.0 {
        return Ok(FormulaValue::Number(0.0));
    }
    if num * multiple < 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let q = num / multiple;
    let rounded = if q >= 0.0 {
        (q + 0.5).floor()
    } else {
        (q - 0.5).ceil()
    };
    Ok(FormulaValue::Number(rounded * multiple))
}

pub fn fn_sumsq(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut values = Vec::new();
    for arg in args {
        if let Some(e) = collect_numbers(arg, &mut values) {
            return Ok(FormulaValue::Error(e));
        }
    }
    Ok(FormulaValue::Number(
        values.into_iter().map(|v| v * v).sum(),
    ))
}

pub fn fn_sqrtpi(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if n < 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number((n * std::f64::consts::PI).sqrt()))
}

pub fn fn_base(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let radix = match get_number(args.get(1)).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if !(2..=36).contains(&radix) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    if number < 0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let min_len = if args.len() >= 3 {
        match get_number(args.get(2)).and_then(to_int_trunc) {
            Ok(v) if v >= 0 => v as usize,
            Ok(_) => return Ok(FormulaValue::Error(CellError::Num)),
            Err(e) => return Ok(e),
        }
    } else {
        0
    };
    let mut chars = Vec::new();
    let mut n = number;
    if n == 0 {
        chars.push('0');
    }
    while n > 0 {
        let d = (n % radix) as u8;
        let ch = if d < 10 {
            (b'0' + d) as char
        } else {
            (b'A' + (d - 10)) as char
        };
        chars.push(ch);
        n /= radix;
    }
    chars.reverse();
    let mut s: String = chars.into_iter().collect();
    while s.len() < min_len {
        s.insert(0, '0');
    }
    Ok(FormulaValue::String(s))
}

pub fn fn_decimal(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match args.first() {
        Some(FormulaValue::String(s)) => s.trim().to_uppercase(),
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let radix = match get_number(args.get(1)).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if !(2..=36).contains(&radix) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    if text.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let mut value = 0f64;
    for ch in text.chars() {
        let d = match ch {
            '0'..='9' => (ch as u8 - b'0') as i64,
            'A'..='Z' => (ch as u8 - b'A' + 10) as i64,
            _ => return Ok(FormulaValue::Error(CellError::Num)),
        };
        if d >= radix {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        value = value * radix as f64 + d as f64;
    }
    Ok(FormulaValue::Number(value))
}

fn roman_table(form: i64) -> &'static [(i64, &'static str)] {
    // Each form level allows progressively more aggressive subtractive pairs.
    // Form 0: classic (IV, IX, XL, XC, CD, CM)
    // Form 1+: add cross-boundary pairs VL(45), LD(450), LM(950)
    // Form 2+: add XD(490), XM(990), IL(49)
    // Form 3+: add VD(495), VM(995), IC(99)
    // Form 4+: add ID(499), IM(999)
    match form {
        0 => &[
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ],
        1 => &[
            (1000, "M"),
            (950, "LM"),
            (900, "CM"),
            (500, "D"),
            (450, "LD"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (45, "VL"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ],
        2 => &[
            (1000, "M"),
            (990, "XM"),
            (950, "LM"),
            (900, "CM"),
            (500, "D"),
            (490, "XD"),
            (450, "LD"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (49, "IL"),
            (45, "VL"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ],
        3 => &[
            (1000, "M"),
            (995, "VM"),
            (990, "XM"),
            (950, "LM"),
            (900, "CM"),
            (500, "D"),
            (495, "VD"),
            (490, "XD"),
            (450, "LD"),
            (400, "CD"),
            (100, "C"),
            (99, "IC"),
            (90, "XC"),
            (50, "L"),
            (49, "IL"),
            (45, "VL"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ],
        _ => &[
            (1000, "M"),
            (999, "IM"),
            (995, "VM"),
            (990, "XM"),
            (950, "LM"),
            (900, "CM"),
            (500, "D"),
            (499, "ID"),
            (495, "VD"),
            (490, "XD"),
            (450, "LD"),
            (400, "CD"),
            (100, "C"),
            (99, "IC"),
            (90, "XC"),
            (50, "L"),
            (49, "IL"),
            (45, "VL"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ],
    }
}

pub fn fn_roman(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let form = if args.len() >= 2 {
        match get_number(args.get(1)).and_then(to_int_trunc) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        }
    } else {
        0
    };
    if !(0..=3999).contains(&n) || !(0..=4).contains(&form) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    if n == 0 {
        return Ok(FormulaValue::String(String::new()));
    }
    let table = roman_table(form);
    let mut result = String::new();
    let mut remaining = n;
    for &(value, numeral) in table {
        while remaining >= value {
            result.push_str(numeral);
            remaining -= value;
        }
    }
    Ok(FormulaValue::String(result))
}

pub fn fn_arabic(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let s = match args.first() {
        Some(FormulaValue::String(s)) => s.trim().to_uppercase(),
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    if s.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    fn val(c: char) -> Option<i64> {
        match c {
            'I' => Some(1),
            'V' => Some(5),
            'X' => Some(10),
            'L' => Some(50),
            'C' => Some(100),
            'D' => Some(500),
            'M' => Some(1000),
            _ => None,
        }
    }
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0usize;
    let mut total = 0i64;
    while i < chars.len() {
        let curr = match val(chars[i]) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        if i + 1 < chars.len() {
            let next = match val(chars[i + 1]) {
                Some(v) => v,
                None => return Ok(FormulaValue::Error(CellError::Num)),
            };
            if curr < next {
                let valid = matches!(
                    (curr, next),
                    (1, 5) | (1, 10) | (10, 50) | (10, 100) | (100, 500) | (100, 1000)
                );
                if !valid {
                    return Ok(FormulaValue::Error(CellError::Num));
                }
                total += next - curr;
                i += 2;
                continue;
            }
        }
        total += curr;
        i += 1;
    }
    if !(1..=3999).contains(&total) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(total as f64))
}

pub fn fn_ceiling_precise(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let significance = if args.len() >= 2 {
        match get_number(args.get(1)) {
            Ok(v) if v != 0.0 => v.abs(),
            Ok(_) => return Ok(FormulaValue::Number(0.0)),
            Err(e) => return Ok(e),
        }
    } else {
        1.0
    };
    Ok(FormulaValue::Number(
        (number / significance).ceil() * significance,
    ))
}

pub fn fn_floor_precise(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let significance = if args.len() >= 2 {
        match get_number(args.get(1)) {
            Ok(v) if v != 0.0 => v.abs(),
            Ok(_) => return Ok(FormulaValue::Number(0.0)),
            Err(e) => return Ok(e),
        }
    } else {
        1.0
    };
    Ok(FormulaValue::Number(
        (number / significance).floor() * significance,
    ))
}

pub fn fn_iso_ceiling(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    fn_ceiling_precise(args, _ctx)
}

#[allow(clippy::needless_range_loop)]
pub fn fn_mdeterm(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut a = match parse_matrix(args.first().unwrap_or(&FormulaValue::Empty)) {
        Ok(m) => m,
        Err(e) => return Ok(e),
    };
    let n = a.len();
    if a.iter().any(|r| r.len() != n) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let mut det = 1.0;
    for i in 0..n {
        let mut pivot = i;
        for r in (i + 1)..n {
            if a[r][i].abs() > a[pivot][i].abs() {
                pivot = r;
            }
        }
        if a[pivot][i].abs() < 1e-15 {
            return Ok(FormulaValue::Number(0.0));
        }
        if pivot != i {
            a.swap(i, pivot);
            det = -det;
        }
        det *= a[i][i];
        for r in (i + 1)..n {
            let factor = a[r][i] / a[i][i];
            for c in i..n {
                a[r][c] -= factor * a[i][c];
            }
        }
    }
    Ok(FormulaValue::Number(det))
}

#[allow(clippy::needless_range_loop)]
pub fn fn_minverse(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let a = match parse_matrix(args.first().unwrap_or(&FormulaValue::Empty)) {
        Ok(m) => m,
        Err(e) => return Ok(e),
    };
    let n = a.len();
    if a.iter().any(|r| r.len() != n) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let mut aug = vec![vec![0.0; 2 * n]; n];
    for r in 0..n {
        for c in 0..n {
            aug[r][c] = a[r][c];
        }
        aug[r][n + r] = 1.0;
    }

    for i in 0..n {
        let mut pivot = i;
        for r in (i + 1)..n {
            if aug[r][i].abs() > aug[pivot][i].abs() {
                pivot = r;
            }
        }
        if aug[pivot][i].abs() < 1e-15 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        if pivot != i {
            aug.swap(i, pivot);
        }
        let p = aug[i][i];
        for c in 0..(2 * n) {
            aug[i][c] /= p;
        }
        for r in 0..n {
            if r == i {
                continue;
            }
            let f = aug[r][i];
            for c in 0..(2 * n) {
                aug[r][c] -= f * aug[i][c];
            }
        }
    }

    let mut inv = vec![vec![FormulaValue::Number(0.0); n]; n];
    for r in 0..n {
        for c in 0..n {
            inv[r][c] = FormulaValue::Number(aug[r][n + c]);
        }
    }
    Ok(FormulaValue::Array { data: inv, source: None })
}

#[allow(clippy::needless_range_loop)]
pub fn fn_mmult(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let a = match parse_matrix(args.first().unwrap_or(&FormulaValue::Empty)) {
        Ok(m) => m,
        Err(e) => return Ok(e),
    };
    let b = match parse_matrix(args.get(1).unwrap_or(&FormulaValue::Empty)) {
        Ok(m) => m,
        Err(e) => return Ok(e),
    };
    let a_rows = a.len();
    let a_cols = a[0].len();
    let b_rows = b.len();
    let b_cols = b[0].len();
    if a_cols != b_rows {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let mut out = vec![vec![FormulaValue::Number(0.0); b_cols]; a_rows];
    for r in 0..a_rows {
        for c in 0..b_cols {
            let mut sum = 0.0;
            for k in 0..a_cols {
                sum += a[r][k] * b[k][c];
            }
            out[r][c] = FormulaValue::Number(sum);
        }
    }
    Ok(FormulaValue::Array { data: out, source: None })
}

pub fn fn_munit(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if n <= 0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let n = n as usize;
    let mut out = vec![vec![FormulaValue::Number(0.0); n]; n];
    for (i, row) in out.iter_mut().enumerate() {
        row[i] = FormulaValue::Number(1.0);
    }
    Ok(FormulaValue::Array { data: out, source: None })
}

pub fn fn_randarray(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    use rand::Rng;
    let mut rng = rand::thread_rng();

    if args.is_empty() {
        return Ok(FormulaValue::Number(rng.r#gen::<f64>()));
    }

    let rows = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) if v > 0 => v as usize,
        Ok(_) => return Ok(FormulaValue::Error(CellError::Num)),
        Err(e) => return Ok(e),
    };
    let cols = if args.len() >= 2 {
        match get_number(args.get(1)).and_then(to_int_trunc) {
            Ok(v) if v > 0 => v as usize,
            Ok(_) => return Ok(FormulaValue::Error(CellError::Num)),
            Err(e) => return Ok(e),
        }
    } else {
        1
    };
    let min = if args.len() >= 3 {
        match get_number(args.get(2)) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        }
    } else {
        0.0
    };
    let max = if args.len() >= 4 {
        match get_number(args.get(3)) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        }
    } else {
        1.0
    };
    if max < min {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let whole = if args.len() >= 5 {
        match args.get(4) {
            Some(FormulaValue::Boolean(b)) => *b,
            Some(FormulaValue::Number(n)) => *n != 0.0,
            Some(FormulaValue::Empty) => false,
            Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        }
    } else {
        false
    };

    let mut out = vec![vec![FormulaValue::Number(0.0); cols]; rows];
    for row in out.iter_mut().take(rows) {
        for cell in row.iter_mut().take(cols) {
            let n = if whole {
                let lo = min.ceil() as i64;
                let hi = max.floor() as i64;
                if lo > hi {
                    return Ok(FormulaValue::Error(CellError::Num));
                }
                rng.gen_range(lo..=hi) as f64
            } else {
                rng.gen_range(min..=max)
            };
            *cell = FormulaValue::Number(n);
        }
    }
    Ok(FormulaValue::Array { data: out, source: None })
}

pub fn fn_seriessum(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match get_number(args.first()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let n = match get_number(args.get(1)).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let m = match get_number(args.get(2)).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let coeffs = match args.get(3) {
        Some(v) => match flatten_array_like(v) {
            Ok(values) => values,
            Err(e) => return Ok(e),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let mut sum = 0.0;
    for (i, c) in coeffs.iter().enumerate() {
        let exp = n + (i as i64) * m;
        sum += c * x.powf(exp as f64);
    }
    Ok(FormulaValue::Number(sum))
}

pub fn fn_sumx2my2(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match flatten_array_like(args.first().unwrap_or(&FormulaValue::Empty)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let y = match flatten_array_like(args.get(1).unwrap_or(&FormulaValue::Empty)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x.len() != y.len() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let sum = x
        .iter()
        .zip(y.iter())
        .map(|(a, b)| a * a - b * b)
        .sum::<f64>();
    Ok(FormulaValue::Number(sum))
}

pub fn fn_sumx2py2(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match flatten_array_like(args.first().unwrap_or(&FormulaValue::Empty)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let y = match flatten_array_like(args.get(1).unwrap_or(&FormulaValue::Empty)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x.len() != y.len() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let sum = x
        .iter()
        .zip(y.iter())
        .map(|(a, b)| a * a + b * b)
        .sum::<f64>();
    Ok(FormulaValue::Number(sum))
}

pub fn fn_sumxmy2(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match flatten_array_like(args.first().unwrap_or(&FormulaValue::Empty)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let y = match flatten_array_like(args.get(1).unwrap_or(&FormulaValue::Empty)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x.len() != y.len() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let sum = x
        .iter()
        .zip(y.iter())
        .map(|(a, b)| (a - b) * (a - b))
        .sum::<f64>();
    Ok(FormulaValue::Number(sum))
}

pub fn fn_aggregate(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    if args.len() < 3 {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let fn_num = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let option = match get_number(args.get(1)).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    // Options that skip errors: 2,3,6,7 (and combinations)
    let skip_errors = matches!(option, 2 | 3 | 6 | 7);

    let mut numbers = Vec::new();
    if skip_errors {
        if (14..=19).contains(&fn_num) {
            collect_numbers_skip_errors(&args[2], &mut numbers);
        } else {
            for arg in args.iter().skip(2) {
                collect_numbers_skip_errors(arg, &mut numbers);
            }
        }
    } else {
        if (14..=19).contains(&fn_num) {
            if let Some(e) = collect_numbers(&args[2], &mut numbers) {
                return Ok(FormulaValue::Error(e));
            }
        } else {
            for arg in args.iter().skip(2) {
                if let Some(e) = collect_numbers(arg, &mut numbers) {
                    return Ok(FormulaValue::Error(e));
                }
            }
        }
    }

    match fn_num {
        1 => {
            if numbers.is_empty() {
                Ok(FormulaValue::Error(CellError::Div0))
            } else {
                Ok(FormulaValue::Number(
                    numbers.iter().sum::<f64>() / numbers.len() as f64,
                ))
            }
        }
        2 => Ok(FormulaValue::Number(numbers.len() as f64)),
        3 => {
            let mut vals = Vec::new();
            for arg in args.iter().skip(2) {
                if let Some(e) = collect_values(arg, &mut vals) {
                    return Ok(FormulaValue::Error(e));
                }
            }
            let count = vals
                .iter()
                .filter(|v| match v {
                    FormulaValue::Empty => false,
                    FormulaValue::String(s) if s.is_empty() => false,
                    _ => true,
                })
                .count();
            Ok(FormulaValue::Number(count as f64))
        }
        4 => {
            if numbers.is_empty() {
                Ok(FormulaValue::Number(0.0))
            } else {
                Ok(FormulaValue::Number(
                    numbers.iter().copied().fold(f64::NEG_INFINITY, f64::max),
                ))
            }
        }
        5 => {
            if numbers.is_empty() {
                Ok(FormulaValue::Number(0.0))
            } else {
                Ok(FormulaValue::Number(
                    numbers.iter().copied().fold(f64::INFINITY, f64::min),
                ))
            }
        }
        6 => {
            if numbers.is_empty() {
                Ok(FormulaValue::Number(0.0))
            } else {
                Ok(FormulaValue::Number(numbers.iter().product()))
            }
        }
        7 => match stdev_s(&numbers) {
            Some(v) => Ok(FormulaValue::Number(v)),
            None => Ok(FormulaValue::Error(CellError::Div0)),
        },
        8 => match stdev_p(&numbers) {
            Some(v) => Ok(FormulaValue::Number(v)),
            None => Ok(FormulaValue::Error(CellError::Div0)),
        },
        9 => Ok(FormulaValue::Number(numbers.iter().sum())),
        10 => match var_s(&numbers) {
            Some(v) => Ok(FormulaValue::Number(v)),
            None => Ok(FormulaValue::Error(CellError::Div0)),
        },
        11 => match var_p(&numbers) {
            Some(v) => Ok(FormulaValue::Number(v)),
            None => Ok(FormulaValue::Error(CellError::Div0)),
        },
        12 => {
            if numbers.is_empty() {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            let n = numbers.len();
            let median = if n % 2 == 1 {
                numbers[n / 2]
            } else {
                (numbers[n / 2 - 1] + numbers[n / 2]) / 2.0
            };
            Ok(FormulaValue::Number(median))
        }
        13 => match mode_sngl(&numbers) {
            Some(v) => Ok(FormulaValue::Number(v)),
            None => Ok(FormulaValue::Error(CellError::Num)),
        },
        14 => {
            if args.len() < 4 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            let k = match get_number(args.get(3)).and_then(to_int_trunc) {
                Ok(v) if v >= 1 => v as usize,
                Ok(_) => return Ok(FormulaValue::Error(CellError::Num)),
                Err(e) => return Ok(e),
            };
            numbers.sort_by(|a, b| b.partial_cmp(a).unwrap_or(std::cmp::Ordering::Equal));
            if k > numbers.len() {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            Ok(FormulaValue::Number(numbers[k - 1]))
        }
        15 => {
            if args.len() < 4 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            let k = match get_number(args.get(3)).and_then(to_int_trunc) {
                Ok(v) if v >= 1 => v as usize,
                Ok(_) => return Ok(FormulaValue::Error(CellError::Num)),
                Err(e) => return Ok(e),
            };
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            if k > numbers.len() {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            Ok(FormulaValue::Number(numbers[k - 1]))
        }
        16 => {
            if args.len() < 4 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            let k = match get_number(args.get(3)) {
                Ok(v) => v,
                Err(e) => return Ok(e),
            };
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            match percentile_inc(&numbers, k) {
                Some(v) => Ok(FormulaValue::Number(v)),
                None => Ok(FormulaValue::Error(CellError::Num)),
            }
        }
        17 => {
            if args.len() < 4 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            let q = match get_number(args.get(3)).and_then(to_int_trunc) {
                Ok(v) if (0..=4).contains(&v) => v as f64 / 4.0,
                Ok(_) => return Ok(FormulaValue::Error(CellError::Num)),
                Err(e) => return Ok(e),
            };
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            match percentile_inc(&numbers, q) {
                Some(v) => Ok(FormulaValue::Number(v)),
                None => Ok(FormulaValue::Error(CellError::Num)),
            }
        }
        18 => {
            if args.len() < 4 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            let k = match get_number(args.get(3)) {
                Ok(v) => v,
                Err(e) => return Ok(e),
            };
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            match percentile_exc(&numbers, k) {
                Some(v) => Ok(FormulaValue::Number(v)),
                None => Ok(FormulaValue::Error(CellError::Num)),
            }
        }
        19 => {
            if args.len() < 4 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            let q = match get_number(args.get(3)).and_then(to_int_trunc) {
                Ok(v) if (1..=3).contains(&v) => v as f64 / 4.0,
                Ok(_) => return Ok(FormulaValue::Error(CellError::Num)),
                Err(e) => return Ok(e),
            };
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            match percentile_exc(&numbers, q) {
                Some(v) => Ok(FormulaValue::Number(v)),
                None => Ok(FormulaValue::Error(CellError::Num)),
            }
        }
        _ => Ok(FormulaValue::Error(CellError::Num)),
    }
}

pub fn fn_subtotal(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let fn_num = match get_number(args.first()).and_then(to_int_trunc) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let normalized = match fn_num {
        1..=11 => fn_num,
        101..=111 => fn_num - 100,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let mut agg_args = vec![
        FormulaValue::Number(normalized as f64),
        FormulaValue::Number(0.0),
    ];
    agg_args.extend_from_slice(&args[1..]);
    fn_aggregate(&agg_args, _ctx)
}

/// PERCENTOF(data_subset, data_all) — percentage that a subset makes up of a given data set
pub fn fn_percentof(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let mut subset_nums = Vec::new();
    if let Some(e) = collect_numbers(&args[0], &mut subset_nums) {
        return Ok(FormulaValue::Error(e));
    }

    let mut all_nums = Vec::new();
    if let Some(e) = collect_numbers(&args[1], &mut all_nums) {
        return Ok(FormulaValue::Error(e));
    }

    let subset_sum: f64 = subset_nums.iter().sum();
    let all_sum: f64 = all_nums.iter().sum();

    if all_sum == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }

    Ok(FormulaValue::Number(subset_sum / all_sum))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    fn num(v: FormulaValue) -> f64 {
        match v {
            FormulaValue::Number(n) => n,
            other => panic!("expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_eval_helper_compiles() {
        let out = eval("=1+2").unwrap();
        assert_eq!(out, FormulaValue::Number(3.0));
    }

    #[test]
    fn test_acosh() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_acosh(&[FormulaValue::Number(1.0)], &ctx).unwrap());
        assert!((v - 0.0).abs() < 1e-10);
        let v = num(fn_acosh(&[FormulaValue::Number(2.0)], &ctx).unwrap());
        assert!((v - 1.3169578969248166).abs() < 1e-10);
        assert_eq!(
            fn_acosh(&[FormulaValue::Number(0.5)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_asinh() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_asinh(&[FormulaValue::Number(0.0)], &ctx).unwrap());
        assert!((v - 0.0).abs() < 1e-10);
        let v = num(fn_asinh(&[FormulaValue::Number(1.0)], &ctx).unwrap());
        assert!((v - 0.881373587019543).abs() < 1e-10);
        let v = num(fn_asinh(&[FormulaValue::Number(-2.0)], &ctx).unwrap());
        assert!((v + 1.4436354751788103).abs() < 1e-10);
    }

    #[test]
    fn test_atanh() {
        let ctx = EvaluationContext::simple();
        let v = num(fn_atanh(&[FormulaValue::Number(0.5)], &ctx).unwrap());
        assert!((v - 0.5493061443340548).abs() < 1e-10);
        assert_eq!(
            fn_atanh(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            fn_atanh(&[FormulaValue::Number(-1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_hyperbolic_and_reciprocals() {
        let ctx = EvaluationContext::simple();
        assert!((num(fn_cosh(&[FormulaValue::Number(0.0)], &ctx).unwrap()) - 1.0).abs() < 1e-10);
        assert!((num(fn_sinh(&[FormulaValue::Number(0.0)], &ctx).unwrap()) - 0.0).abs() < 1e-10);
        assert!((num(fn_tanh(&[FormulaValue::Number(0.0)], &ctx).unwrap()) - 0.0).abs() < 1e-10);
        assert_eq!(
            fn_cot(&[FormulaValue::Number(0.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
        assert_eq!(
            fn_coth(&[FormulaValue::Number(0.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
        assert_eq!(
            fn_csc(&[FormulaValue::Number(0.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
        assert_eq!(
            fn_csch(&[FormulaValue::Number(0.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
        assert!((num(fn_sec(&[FormulaValue::Number(0.0)], &ctx).unwrap()) - 1.0).abs() < 1e-10);
        assert!((num(fn_sech(&[FormulaValue::Number(0.0)], &ctx).unwrap()) - 1.0).abs() < 1e-10);
    }

    #[test]
    fn test_combinatorics() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_combin(
                &[FormulaValue::Number(5.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(10.0)
        );
        assert_eq!(
            fn_combin(
                &[FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            fn_combina(
                &[FormulaValue::Number(3.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );
        assert_eq!(
            fn_fact(&[FormulaValue::Number(5.0)], &ctx).unwrap(),
            FormulaValue::Number(120.0)
        );
        assert_eq!(
            fn_fact(&[FormulaValue::Number(-1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            fn_factdouble(&[FormulaValue::Number(7.0)], &ctx).unwrap(),
            FormulaValue::Number(105.0)
        );
    }

    #[test]
    fn test_multinomial_gcd_lcm() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_multinomial(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(3.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(60.0)
        );
        assert_eq!(
            fn_gcd(
                &[FormulaValue::Number(24.0), FormulaValue::Number(18.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );
        assert_eq!(
            fn_lcm(
                &[FormulaValue::Number(6.0), FormulaValue::Number(8.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(24.0)
        );
    }

    #[test]
    fn test_product_quotient_mround_sumsq_sqrtpi() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_product(
                &[
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(4.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(24.0)
        );
        assert_eq!(
            fn_quotient(
                &[FormulaValue::Number(7.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            fn_quotient(
                &[FormulaValue::Number(7.0), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
        assert_eq!(
            fn_mround(
                &[FormulaValue::Number(10.0), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(9.0)
        );
        assert_eq!(
            fn_sumsq(
                &[FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(25.0)
        );
        let v = num(fn_sqrtpi(&[FormulaValue::Number(2.0)], &ctx).unwrap());
        assert!((v - (2.0 * std::f64::consts::PI).sqrt()).abs() < 1e-10);
    }

    #[test]
    fn test_base_decimal_roman_arabic() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_base(
                &[FormulaValue::Number(31.0), FormulaValue::Number(16.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("1F".to_string())
        );
        assert_eq!(
            fn_decimal(
                &[
                    FormulaValue::String("1F".to_string()),
                    FormulaValue::Number(16.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(31.0)
        );
        assert_eq!(
            fn_roman(
                &[FormulaValue::Number(1999.0), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("MCMXCIX".to_string())
        );
        assert_eq!(
            fn_roman(
                &[FormulaValue::Number(499.0), FormulaValue::Number(4.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("ID".to_string())
        );
        assert_eq!(
            fn_arabic(&[FormulaValue::String("MCMXCIX".to_string())], &ctx).unwrap(),
            FormulaValue::Number(1999.0)
        );
        assert_eq!(
            fn_arabic(&[FormulaValue::String("IC".to_string())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_precise_rounding() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_ceiling_precise(&[FormulaValue::Number(4.3)], &ctx).unwrap(),
            FormulaValue::Number(5.0)
        );
        assert_eq!(
            fn_ceiling_precise(&[FormulaValue::Number(-4.3)], &ctx).unwrap(),
            FormulaValue::Number(-4.0)
        );
        assert_eq!(
            fn_floor_precise(&[FormulaValue::Number(4.3)], &ctx).unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            fn_floor_precise(&[FormulaValue::Number(-4.3)], &ctx).unwrap(),
            FormulaValue::Number(-5.0)
        );
        assert_eq!(
            fn_iso_ceiling(&[FormulaValue::Number(-4.3)], &ctx).unwrap(),
            FormulaValue::Number(-4.0)
        );
    }

    #[test]
    fn test_matrix_functions() {
        let ctx = EvaluationContext::simple();
        let matrix = FormulaValue::Array { data: vec![
            vec![FormulaValue::Number(1.0), FormulaValue::Number(2.0)],
            vec![FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
        ], source: None };
        let det = num(fn_mdeterm(std::slice::from_ref(&matrix), &ctx).unwrap());
        assert!((det + 2.0).abs() < 1e-10);
        let inv = fn_minverse(std::slice::from_ref(&matrix), &ctx).unwrap();
        if let FormulaValue::Array { data: rows, .. } = inv {
            if let FormulaValue::Number(v) = rows[0][0] {
                assert!((v + 2.0).abs() < 1e-10);
            } else {
                panic!("expected number");
            }
        } else {
            panic!("expected array");
        }
        let mm = fn_mmult(&[matrix.clone(), matrix], &ctx).unwrap();
        if let FormulaValue::Array { data: rows, .. } = mm {
            assert_eq!(rows[0][0], FormulaValue::Number(7.0));
            assert_eq!(rows[1][1], FormulaValue::Number(22.0));
        } else {
            panic!("expected array");
        }
        let id = fn_munit(&[FormulaValue::Number(3.0)], &ctx).unwrap();
        if let FormulaValue::Array { data: rows, .. } = id {
            assert_eq!(rows[0][0], FormulaValue::Number(1.0));
            assert_eq!(rows[1][1], FormulaValue::Number(1.0));
            assert_eq!(rows[2][2], FormulaValue::Number(1.0));
        } else {
            panic!("expected array");
        }
    }

    #[test]
    fn test_randarray_and_seriessum() {
        let ctx = EvaluationContext::simple();
        let one = fn_randarray(&[], &ctx).unwrap();
        if let FormulaValue::Number(n) = one {
            assert!((0.0..=1.0).contains(&n));
        } else {
            panic!("expected number");
        }
        let arr = fn_randarray(
            &[
                FormulaValue::Number(2.0),
                FormulaValue::Number(3.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(5.0),
                FormulaValue::Boolean(true),
            ],
            &ctx,
        )
        .unwrap();
        if let FormulaValue::Array { data: rows, .. } = arr {
            assert_eq!(rows.len(), 2);
            assert_eq!(rows[0].len(), 3);
        } else {
            panic!("expected array");
        }
        let coeffs = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(1.0),
            FormulaValue::Number(2.0),
            FormulaValue::Number(3.0),
        ]], source: None };
        let s = num(fn_seriessum(
            &[
                FormulaValue::Number(2.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(2.0),
                coeffs,
            ],
            &ctx,
        )
        .unwrap());
        assert!((s - 114.0).abs() < 1e-10);
    }

    #[test]
    fn test_cross_array_functions() {
        let ctx = EvaluationContext::simple();
        let x = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(1.0),
            FormulaValue::Number(2.0),
        ]], source: None };
        let y = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(3.0),
            FormulaValue::Number(4.0),
        ]], source: None };
        assert_eq!(
            fn_sumx2my2(&[x.clone(), y.clone()], &ctx).unwrap(),
            FormulaValue::Number(-20.0)
        );
        assert_eq!(
            fn_sumx2py2(&[x.clone(), y.clone()], &ctx).unwrap(),
            FormulaValue::Number(30.0)
        );
        assert_eq!(
            fn_sumxmy2(&[x, y], &ctx).unwrap(),
            FormulaValue::Number(8.0)
        );
    }

    #[test]
    fn test_aggregate_and_subtotal() {
        let ctx = EvaluationContext::simple();
        let data = FormulaValue::Array { data: vec![
            vec![FormulaValue::Number(1.0), FormulaValue::Number(2.0)],
            vec![FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
        ], source: None };
        assert_eq!(
            fn_aggregate(
                &[
                    FormulaValue::Number(9.0),
                    FormulaValue::Number(0.0),
                    data.clone()
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(10.0)
        );
        assert_eq!(
            fn_aggregate(
                &[
                    FormulaValue::Number(4.0),
                    FormulaValue::Number(0.0),
                    data.clone()
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            fn_aggregate(
                &[
                    FormulaValue::Number(14.0),
                    FormulaValue::Number(0.0),
                    data.clone(),
                    FormulaValue::Number(2.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            fn_subtotal(
                &[
                    FormulaValue::Number(9.0),
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(3.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );
        assert_eq!(
            fn_subtotal(
                &[
                    FormulaValue::Number(109.0),
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            fn_subtotal(
                &[FormulaValue::Number(99.0), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    // ================================================================
    // Docs-based tests (from Microsoft support pages)
    // ================================================================

    // ── MROUND docs ──
    // https://support.microsoft.com/en-us/office/mround-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427

    #[test]
    fn test_mround_docs() {
        let ctx = EvaluationContext::simple();
        // =MROUND(10, 3) = 9
        assert_eq!(
            fn_mround(
                &[FormulaValue::Number(10.0), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(9.0)
        );
        // =MROUND(-10, -3) = -9
        assert_eq!(
            fn_mround(
                &[FormulaValue::Number(-10.0), FormulaValue::Number(-3.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-9.0)
        );
        // =MROUND(1.3, 0.2) = 1.4
        assert!(
            (num(fn_mround(
                &[FormulaValue::Number(1.3), FormulaValue::Number(0.2)],
                &ctx
            )
            .unwrap())
                - 1.4)
                .abs()
                < 1e-10
        );
        // =MROUND(5, -2) = #NUM! (different signs)
        assert_eq!(
            fn_mround(
                &[FormulaValue::Number(5.0), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    // ── PRODUCT docs ──
    // https://support.microsoft.com/en-us/office/product-function-8e6b5b24-90ee-4650-aeec-80982a0512ce

    #[test]
    fn test_product_docs() {
        let ctx = EvaluationContext::simple();
        // =PRODUCT(5, 15, 30) = 2250
        assert_eq!(
            num(fn_product(
                &[
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(15.0),
                    FormulaValue::Number(30.0)
                ],
                &ctx
            )
            .unwrap()),
            2250.0
        );
        // =PRODUCT(5, 15, 30, 2) = 4500
        assert_eq!(
            num(fn_product(
                &[
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(15.0),
                    FormulaValue::Number(30.0),
                    FormulaValue::Number(2.0)
                ],
                &ctx
            )
            .unwrap()),
            4500.0
        );
    }

    // ── SUMSQ docs ──
    // https://support.microsoft.com/en-us/office/sumsq-function-e3313c02-51cc-4963-aae6-31442d9ec307

    #[test]
    fn test_sumsq_docs() {
        let ctx = EvaluationContext::simple();
        // =SUMSQ(3, 4) = 25
        assert_eq!(
            num(fn_sumsq(
                &[FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
                &ctx
            )
            .unwrap()),
            25.0
        );
    }

    // ── ACOSH docs ──
    // https://support.microsoft.com/en-us/office/acosh-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe

    #[test]
    fn test_acosh_docs() {
        let ctx = EvaluationContext::simple();
        // =ACOSH(1) = 0
        assert_eq!(
            num(fn_acosh(&[FormulaValue::Number(1.0)], &ctx).unwrap()),
            0.0
        );
        // =ACOSH(10) = 2.993222846...
        assert!(
            (num(fn_acosh(&[FormulaValue::Number(10.0)], &ctx).unwrap()) - 2.993222846).abs()
                < 1e-6
        );
    }

    // ── ASINH docs ──
    // https://support.microsoft.com/en-us/office/asinh-function-4e00475a-067a-43cf-926a-765b0249717c

    #[test]
    fn test_asinh_docs() {
        let ctx = EvaluationContext::simple();
        // =ASINH(-2.5) = -1.647231146...
        assert!(
            (num(fn_asinh(&[FormulaValue::Number(-2.5)], &ctx).unwrap()) - (-1.647231146)).abs()
                < 1e-6
        );
        // =ASINH(10) = 2.998222950...
        assert!(
            (num(fn_asinh(&[FormulaValue::Number(10.0)], &ctx).unwrap()) - 2.998222950).abs()
                < 1e-6
        );
    }

    // ── ATANH docs ──
    // https://support.microsoft.com/en-us/office/atanh-function-3cd65768-0de7-4f1d-b312-d01c8c930d90

    #[test]
    fn test_atanh_docs() {
        let ctx = EvaluationContext::simple();
        // =ATANH(0.76159416) = 1.00000001 (approx)
        assert!(
            (num(fn_atanh(&[FormulaValue::Number(0.76159416)], &ctx).unwrap()) - 1.00000001).abs()
                < 1e-6
        );
        // =ATANH(-0.1) = -0.10033535...
        assert!(
            (num(fn_atanh(&[FormulaValue::Number(-0.1)], &ctx).unwrap()) - (-0.10033535)).abs()
                < 1e-6
        );
    }

    // ── COSH docs ──
    // https://support.microsoft.com/en-us/office/cosh-function-e460d426-c471-43e8-889a-2ab3a29c45c6

    #[test]
    fn test_cosh_docs() {
        let ctx = EvaluationContext::simple();
        // =COSH(4) = 27.30823284...
        assert!(
            (num(fn_cosh(&[FormulaValue::Number(4.0)], &ctx).unwrap()) - 27.30823284).abs() < 1e-5
        );
        // =COSH(EXP(1)) = 7.61012514...
        assert!(
            (num(fn_cosh(&[FormulaValue::Number(std::f64::consts::E)], &ctx).unwrap())
                - 7.61012514)
                .abs()
                < 1e-5
        );
    }

    // ── SINH docs ──
    // https://support.microsoft.com/en-us/office/sinh-function-c0571628-3f37-4bbc-bcc0-69db67c9bcb0

    #[test]
    fn test_sinh_docs() {
        let ctx = EvaluationContext::simple();
        // =SINH(1) = 1.175201194...
        assert!(
            (num(fn_sinh(&[FormulaValue::Number(1.0)], &ctx).unwrap()) - 1.175201194).abs() < 1e-6
        );
    }

    // ── TANH docs ──
    // https://support.microsoft.com/en-us/office/tanh-function-017222f0-a0c3-4f69-9787-b3202295dc6c

    #[test]
    fn test_tanh_docs() {
        let ctx = EvaluationContext::simple();
        // =TANH(-2) = -0.96402758...
        assert!(
            (num(fn_tanh(&[FormulaValue::Number(-2.0)], &ctx).unwrap()) - (-0.96402758)).abs()
                < 1e-6
        );
        // =TANH(0) = 0
        assert_eq!(
            num(fn_tanh(&[FormulaValue::Number(0.0)], &ctx).unwrap()),
            0.0
        );
        // =TANH(0.5) = 0.46211716...
        assert!(
            (num(fn_tanh(&[FormulaValue::Number(0.5)], &ctx).unwrap()) - 0.46211716).abs() < 1e-6
        );
    }

    // ── COT docs ──
    // https://support.microsoft.com/en-us/office/cot-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a

    #[test]
    fn test_cot_docs() {
        let ctx = EvaluationContext::simple();
        // =COT(2) = -0.45765755...
        assert!(
            (num(fn_cot(&[FormulaValue::Number(2.0)], &ctx).unwrap()) - (-0.45765755)).abs() < 1e-6
        );
        // =COT(30*PI()/180) = 1.73205080... (cot of 30 degrees)
        let rad30 = 30.0 * std::f64::consts::PI / 180.0;
        assert!(
            (num(fn_cot(&[FormulaValue::Number(rad30)], &ctx).unwrap()) - 1.73205080).abs() < 1e-5
        );
    }

    // ── COTH docs ──
    // https://support.microsoft.com/en-us/office/coth-function-2d78f1fd-7a21-47c1-9c76-5b1a44f7a1ee

    #[test]
    fn test_coth_docs() {
        let ctx = EvaluationContext::simple();
        // =COTH(2) = 1.03731472...
        assert!(
            (num(fn_coth(&[FormulaValue::Number(2.0)], &ctx).unwrap()) - 1.03731472).abs() < 1e-6
        );
    }

    // ── ACOT docs ──
    // https://support.microsoft.com/en-us/office/acot-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905

    #[test]
    fn test_acot_docs() {
        let ctx = EvaluationContext::simple();
        // =ACOT(2) = 0.4636 (docs example)
        assert!((num(fn_acot(&[FormulaValue::Number(2.0)], &ctx).unwrap()) - 0.4636).abs() < 1e-4);
        // acot(0) = PI/2
        assert!(
            (num(fn_acot(&[FormulaValue::Number(0.0)], &ctx).unwrap())
                - std::f64::consts::FRAC_PI_2)
                .abs()
                < 1e-10
        );
        // acot(1) = PI/4
        assert!(
            (num(fn_acot(&[FormulaValue::Number(1.0)], &ctx).unwrap())
                - std::f64::consts::FRAC_PI_4)
                .abs()
                < 1e-10
        );
        // acot(-1) = 3*PI/4 (result in range 0..pi)
        assert!(
            (num(fn_acot(&[FormulaValue::Number(-1.0)], &ctx).unwrap())
                - 3.0 * std::f64::consts::FRAC_PI_4)
                .abs()
                < 1e-10
        );
        // Non-numeric input → #VALUE!
        assert_eq!(
            fn_acot(&[FormulaValue::String("text".into())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    // ── ACOTH docs ──
    // https://support.microsoft.com/en-us/office/acoth-function-cc49480f-f684-4171-9fc5-73e4e852300f

    #[test]
    fn test_acoth_docs() {
        let ctx = EvaluationContext::simple();
        // =ACOTH(6) = 0.168 (docs example)
        assert!((num(fn_acoth(&[FormulaValue::Number(6.0)], &ctx).unwrap()) - 0.168).abs() < 1e-3);
        // |x| <= 1 → #NUM!
        assert_eq!(
            fn_acoth(&[FormulaValue::Number(0.5)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            fn_acoth(&[FormulaValue::Number(-0.5)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            fn_acoth(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            fn_acoth(&[FormulaValue::Number(-1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        // Non-numeric input → #VALUE!
        assert_eq!(
            fn_acoth(&[FormulaValue::String("text".into())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    // ── CSC docs ──
    // https://support.microsoft.com/en-us/office/csc-function-07379cb5-f7a8-44b7-ade8-da5803d22a3a

    #[test]
    fn test_csc_docs() {
        let ctx = EvaluationContext::simple();
        // =CSC(PI()/4) = 1.41421356... (sqrt(2))
        let v = num(fn_csc(&[FormulaValue::Number(std::f64::consts::FRAC_PI_4)], &ctx).unwrap());
        assert!((v - std::f64::consts::SQRT_2).abs() < 1e-6);
    }

    // ── CSCH docs ──
    // https://support.microsoft.com/en-us/office/csch-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb

    #[test]
    fn test_csch_docs() {
        let ctx = EvaluationContext::simple();
        // =CSCH(1.5) = 0.46964244...
        assert!(
            (num(fn_csch(&[FormulaValue::Number(1.5)], &ctx).unwrap()) - 0.46964244).abs() < 1e-6
        );
    }

    // ── SEC docs ──
    // https://support.microsoft.com/en-us/office/sec-function-ff224717-9c87-4170-9b58-d069ced6d5f7

    #[test]
    fn test_sec_docs() {
        let ctx = EvaluationContext::simple();
        // =SEC(45*PI()/180) = 1.41421356... (sqrt(2))
        let rad45 = 45.0 * std::f64::consts::PI / 180.0;
        assert!(
            (num(fn_sec(&[FormulaValue::Number(rad45)], &ctx).unwrap()) - std::f64::consts::SQRT_2)
                .abs()
                < 1e-6
        );
        // =SEC(1) = 1.85081571...
        assert!(
            (num(fn_sec(&[FormulaValue::Number(1.0)], &ctx).unwrap()) - 1.85081571).abs() < 1e-6
        );
    }

    // ── SECH docs ──
    // https://support.microsoft.com/en-us/office/sech-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f

    #[test]
    fn test_sech_docs() {
        let ctx = EvaluationContext::simple();
        // =SECH(45) = 5.73E-20 (docs example)
        assert!(
            (num(fn_sech(&[FormulaValue::Number(45.0)], &ctx).unwrap()) - 5.73e-20).abs() < 1e-21
        );
        // =SECH(30) = 1.87E-13 (docs example)
        assert!(
            (num(fn_sech(&[FormulaValue::Number(30.0)], &ctx).unwrap()) - 1.87e-13).abs() < 1e-14
        );
        // =SECH(0) = 1
        assert_eq!(
            num(fn_sech(&[FormulaValue::Number(0.0)], &ctx).unwrap()),
            1.0
        );
    }

    // ── COMBIN docs ──
    // https://support.microsoft.com/en-us/office/combin-function-12a3f276-0a21-423a-8de6-06990aaf638a

    #[test]
    fn test_combin_docs() {
        let ctx = EvaluationContext::simple();
        // =COMBIN(8, 2) = 28
        assert_eq!(
            fn_combin(
                &[FormulaValue::Number(8.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(28.0)
        );
    }

    // ── COMBINA docs ──
    // https://support.microsoft.com/en-us/office/combina-function-efb49efc-4a5b-4f00-9785-28a4ee5c8610

    #[test]
    fn test_combina_docs() {
        let ctx = EvaluationContext::simple();
        // =COMBINA(4, 3) = 20
        assert_eq!(
            fn_combina(
                &[FormulaValue::Number(4.0), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(20.0)
        );
        // =COMBINA(10, 3) = 220
        assert_eq!(
            fn_combina(
                &[FormulaValue::Number(10.0), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(220.0)
        );
    }

    // ── FACT docs ──
    // https://support.microsoft.com/en-us/office/fact-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3

    #[test]
    fn test_fact_docs() {
        let ctx = EvaluationContext::simple();
        // =FACT(5) = 120
        assert_eq!(
            fn_fact(&[FormulaValue::Number(5.0)], &ctx).unwrap(),
            FormulaValue::Number(120.0)
        );
        // =FACT(1.9) = 1 (truncated to 1)
        assert_eq!(
            fn_fact(&[FormulaValue::Number(1.9)], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );
        // =FACT(0) = 1
        assert_eq!(
            fn_fact(&[FormulaValue::Number(0.0)], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );
        // =FACT(-1) = #NUM!
        assert_eq!(
            fn_fact(&[FormulaValue::Number(-1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        // =FACT(1) = 1
        assert_eq!(
            fn_fact(&[FormulaValue::Number(1.0)], &ctx).unwrap(),
            FormulaValue::Number(1.0)
        );
    }

    // ── FACTDOUBLE docs ──
    // https://support.microsoft.com/en-us/office/factdouble-function-e67697ac-d214-48eb-b7b7-cce2589ecac8

    #[test]
    fn test_factdouble_docs() {
        let ctx = EvaluationContext::simple();
        // =FACTDOUBLE(6) = 48 (6*4*2)
        assert_eq!(
            fn_factdouble(&[FormulaValue::Number(6.0)], &ctx).unwrap(),
            FormulaValue::Number(48.0)
        );
        // =FACTDOUBLE(7) = 105 (7*5*3*1)
        assert_eq!(
            fn_factdouble(&[FormulaValue::Number(7.0)], &ctx).unwrap(),
            FormulaValue::Number(105.0)
        );
    }

    // ── MULTINOMIAL docs ──
    // https://support.microsoft.com/en-us/office/multinomial-function-6fa6ef1b-7628-4e08-9c81-7e1645f24bf0

    #[test]
    fn test_multinomial_docs() {
        let ctx = EvaluationContext::simple();
        // =MULTINOMIAL(2, 3, 4) = 1260
        assert_eq!(
            fn_multinomial(
                &[
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(4.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(1260.0)
        );
    }

    // ── GCD docs ──
    // https://support.microsoft.com/en-us/office/gcd-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a

    #[test]
    fn test_gcd_docs() {
        let ctx = EvaluationContext::simple();
        // =GCD(5, 2) = 1
        assert_eq!(
            fn_gcd(
                &[FormulaValue::Number(5.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(1.0)
        );
        // =GCD(24, 36) = 12
        assert_eq!(
            fn_gcd(
                &[FormulaValue::Number(24.0), FormulaValue::Number(36.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(12.0)
        );
        // =GCD(7, 1) = 1
        assert_eq!(
            fn_gcd(
                &[FormulaValue::Number(7.0), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(1.0)
        );
        // =GCD(128, 80) = 16
        assert_eq!(
            fn_gcd(
                &[FormulaValue::Number(128.0), FormulaValue::Number(80.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(16.0)
        );
    }

    // ── LCM docs ──
    // https://support.microsoft.com/en-us/office/lcm-function-7152b67a-8bb5-4075-ae5c-06ede5563c94

    #[test]
    fn test_lcm_docs() {
        let ctx = EvaluationContext::simple();
        // =LCM(5, 2) = 10
        assert_eq!(
            fn_lcm(
                &[FormulaValue::Number(5.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(10.0)
        );
        // =LCM(24, 36) = 72
        assert_eq!(
            fn_lcm(
                &[FormulaValue::Number(24.0), FormulaValue::Number(36.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(72.0)
        );
    }

    // ── QUOTIENT docs ──
    // https://support.microsoft.com/en-us/office/quotient-function-9f7bf099-2a18-4282-8fa4-65290cc99dee

    #[test]
    fn test_quotient_docs() {
        let ctx = EvaluationContext::simple();
        // =QUOTIENT(5, 2) = 2
        assert_eq!(
            fn_quotient(
                &[FormulaValue::Number(5.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(2.0)
        );
        // =QUOTIENT(4.5, 3.1) = 1
        assert_eq!(
            fn_quotient(
                &[FormulaValue::Number(4.5), FormulaValue::Number(3.1)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(1.0)
        );
        // =QUOTIENT(-10, 3) = -3
        assert_eq!(
            fn_quotient(
                &[FormulaValue::Number(-10.0), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-3.0)
        );
    }

    // ── SQRTPI docs ──
    // https://support.microsoft.com/en-us/office/sqrtpi-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4

    #[test]
    fn test_sqrtpi_docs() {
        let ctx = EvaluationContext::simple();
        // =SQRTPI(1) = 1.77245385...
        assert!(
            (num(fn_sqrtpi(&[FormulaValue::Number(1.0)], &ctx).unwrap()) - 1.77245385).abs() < 1e-6
        );
        // =SQRTPI(2) = 2.50662827...
        assert!(
            (num(fn_sqrtpi(&[FormulaValue::Number(2.0)], &ctx).unwrap()) - 2.50662827).abs() < 1e-6
        );
    }

    // ── BASE docs ──
    // https://support.microsoft.com/en-us/office/base-function-2ef61411-aef9-4b19-9e1b-4ed8ec21363e

    #[test]
    fn test_base_docs() {
        let ctx = EvaluationContext::simple();
        // =BASE(7, 2) = "111"
        assert_eq!(
            fn_base(
                &[FormulaValue::Number(7.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("111".to_string())
        );
        // =BASE(100, 16) = "64"
        assert_eq!(
            fn_base(
                &[FormulaValue::Number(100.0), FormulaValue::Number(16.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("64".to_string())
        );
        // =BASE(15, 2, 10) = "0000001111" (padded to 10 chars)
        assert_eq!(
            fn_base(
                &[
                    FormulaValue::Number(15.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(10.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("0000001111".to_string())
        );
    }

    // ── DECIMAL docs ──
    // https://support.microsoft.com/en-us/office/decimal-function-ee554665-6176-46ef-82de-0a283658da2e

    #[test]
    fn test_decimal_docs() {
        let ctx = EvaluationContext::simple();
        // =DECIMAL("FF", 16) = 255
        assert_eq!(
            fn_decimal(
                &[
                    FormulaValue::String("FF".to_string()),
                    FormulaValue::Number(16.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(255.0)
        );
        // =DECIMAL("111", 2) = 7
        assert_eq!(
            fn_decimal(
                &[
                    FormulaValue::String("111".to_string()),
                    FormulaValue::Number(2.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(7.0)
        );
        // =DECIMAL("zap", 36) = 45745 (docs: 35*36^2 + 10*36 + 25)
        assert_eq!(
            fn_decimal(
                &[
                    FormulaValue::String("zap".to_string()),
                    FormulaValue::Number(36.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(45745.0)
        );
    }

    // ── ROMAN docs ──
    // https://support.microsoft.com/en-us/office/roman-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5

    #[test]
    fn test_roman_docs() {
        let ctx = EvaluationContext::simple();
        // =ROMAN(499, 0) = "CDXCIX" (classic)
        assert_eq!(
            fn_roman(
                &[FormulaValue::Number(499.0), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("CDXCIX".to_string())
        );
        // =ROMAN(499, 1) = "LDVLIV"
        assert_eq!(
            fn_roman(
                &[FormulaValue::Number(499.0), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("LDVLIV".to_string())
        );
        // =ROMAN(499, 2) = "XDIX"
        assert_eq!(
            fn_roman(
                &[FormulaValue::Number(499.0), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("XDIX".to_string())
        );
        // =ROMAN(499, 3) = "VDIV"
        assert_eq!(
            fn_roman(
                &[FormulaValue::Number(499.0), FormulaValue::Number(3.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("VDIV".to_string())
        );
        // =ROMAN(499, 4) = "ID"
        assert_eq!(
            fn_roman(
                &[FormulaValue::Number(499.0), FormulaValue::Number(4.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::String("ID".to_string())
        );
    }

    // ── ARABIC docs ──
    // https://support.microsoft.com/en-us/office/arabic-function-9a8da418-c17b-4ef9-a657-9370a30a674f

    #[test]
    fn test_arabic_docs() {
        let ctx = EvaluationContext::simple();
        // =ARABIC("LVII") = 57
        assert_eq!(
            fn_arabic(&[FormulaValue::String("LVII".to_string())], &ctx).unwrap(),
            FormulaValue::Number(57.0)
        );
        // =ARABIC("MCMXII") = 1912
        assert_eq!(
            fn_arabic(&[FormulaValue::String("MCMXII".to_string())], &ctx).unwrap(),
            FormulaValue::Number(1912.0)
        );
    }

    // ── CEILING.PRECISE docs ──
    // https://support.microsoft.com/en-us/office/ceiling-precise-function-f366a774-527a-4c92-ba49-af0a196e66cb

    #[test]
    fn test_ceiling_precise_docs() {
        let ctx = EvaluationContext::simple();
        // =CEILING.PRECISE(4.3) = 5
        assert_eq!(
            fn_ceiling_precise(&[FormulaValue::Number(4.3)], &ctx).unwrap(),
            FormulaValue::Number(5.0)
        );
        // =CEILING.PRECISE(-4.3) = -4
        assert_eq!(
            fn_ceiling_precise(&[FormulaValue::Number(-4.3)], &ctx).unwrap(),
            FormulaValue::Number(-4.0)
        );
        // =CEILING.PRECISE(4.3, 2) = 6
        assert_eq!(
            fn_ceiling_precise(
                &[FormulaValue::Number(4.3), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );
        // =CEILING.PRECISE(4.3, -2) = 6 (sign of significance is ignored)
        assert_eq!(
            fn_ceiling_precise(
                &[FormulaValue::Number(4.3), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );
        // =CEILING.PRECISE(-4.3, 2) = -4
        assert_eq!(
            fn_ceiling_precise(
                &[FormulaValue::Number(-4.3), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-4.0)
        );
        // =CEILING.PRECISE(-4.3, -2) = -4
        assert_eq!(
            fn_ceiling_precise(
                &[FormulaValue::Number(-4.3), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-4.0)
        );
    }

    // ── FLOOR.PRECISE docs ──
    // https://support.microsoft.com/en-us/office/floor-precise-function-f769b468-1452-4617-8dc3-02f842a0702e

    #[test]
    fn test_floor_precise_docs() {
        let ctx = EvaluationContext::simple();
        // =FLOOR.PRECISE(-3.2, -1) = -4
        assert_eq!(
            fn_floor_precise(
                &[FormulaValue::Number(-3.2), FormulaValue::Number(-1.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-4.0)
        );
        // =FLOOR.PRECISE(3.2, 1) = 3
        assert_eq!(
            fn_floor_precise(
                &[FormulaValue::Number(3.2), FormulaValue::Number(1.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(3.0)
        );
        // =FLOOR.PRECISE(-3.2, 2) = -4
        assert_eq!(
            fn_floor_precise(
                &[FormulaValue::Number(-3.2), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-4.0)
        );
        // =FLOOR.PRECISE(3.2, -2) = 2 (sign of significance is ignored)
        assert_eq!(
            fn_floor_precise(
                &[FormulaValue::Number(3.2), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(2.0)
        );
        // =FLOOR.PRECISE(3.2) = 3
        assert_eq!(
            fn_floor_precise(&[FormulaValue::Number(3.2)], &ctx).unwrap(),
            FormulaValue::Number(3.0)
        );
    }

    // ── ISO.CEILING docs ──
    // https://support.microsoft.com/en-us/office/iso-ceiling-function-e587bb73-6cc2-4113-b664-ff5b09859a83

    #[test]
    fn test_iso_ceiling_docs() {
        let ctx = EvaluationContext::simple();
        // =ISO.CEILING(4.3) = 5
        assert_eq!(
            fn_iso_ceiling(&[FormulaValue::Number(4.3)], &ctx).unwrap(),
            FormulaValue::Number(5.0)
        );
        // =ISO.CEILING(-4.3) = -4
        assert_eq!(
            fn_iso_ceiling(&[FormulaValue::Number(-4.3)], &ctx).unwrap(),
            FormulaValue::Number(-4.0)
        );
        // =ISO.CEILING(4.3, 2) = 6
        assert_eq!(
            fn_iso_ceiling(
                &[FormulaValue::Number(4.3), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );
        // =ISO.CEILING(-4.3, 2) = -4
        assert_eq!(
            fn_iso_ceiling(
                &[FormulaValue::Number(-4.3), FormulaValue::Number(2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-4.0)
        );
        // =ISO.CEILING(-4.3, -2) = -4 (sign of significance ignored)
        assert_eq!(
            fn_iso_ceiling(
                &[FormulaValue::Number(-4.3), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(-4.0)
        );
        // =ISO.CEILING(4.3, -2) = 6 (sign of significance ignored)
        assert_eq!(
            fn_iso_ceiling(
                &[FormulaValue::Number(4.3), FormulaValue::Number(-2.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );
    }

    // ── CEILING / FLOOR compatibility via eval ──
    // These are registered as separate functions from CEILING.MATH/FLOOR.MATH

    #[test]
    fn test_ceiling_compat_docs() {
        // =CEILING(2.5, 1) = 3
        assert_eq!(num(eval("=CEILING(2.5, 1)").unwrap()), 3.0);
        // =CEILING(-2.5, -2) = -4
        assert_eq!(num(eval("=CEILING(-2.5, -2)").unwrap()), -4.0);
        // =CEILING(-2.5, 2) = #NUM! (compat: different signs)
        assert_eq!(
            eval("=CEILING(-2.5, 2)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        // =CEILING(1.5, 0.1) = 1.5
        assert!((num(eval("=CEILING(1.5, 0.1)").unwrap()) - 1.5).abs() < 1e-10);
        // =CEILING(0.234, 0.01) = 0.24
        assert!((num(eval("=CEILING(0.234, 0.01)").unwrap()) - 0.24).abs() < 1e-10);
    }

    #[test]
    fn test_floor_compat_docs() {
        // =FLOOR(3.7, 2) = 2
        assert_eq!(num(eval("=FLOOR(3.7, 2)").unwrap()), 2.0);
        // =FLOOR(-2.5, -2) = -2
        assert_eq!(num(eval("=FLOOR(-2.5, -2)").unwrap()), -2.0);
        // =FLOOR(0.234, 0.01) = 0.23
        assert!((num(eval("=FLOOR(0.234, 0.01)").unwrap()) - 0.23).abs() < 1e-10);
        // =FLOOR(-2.5, 2) = #NUM! (different signs in Excel)
        assert_eq!(
            eval("=FLOOR(-2.5, 2)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    // ── SERIESSUM docs ──
    // https://support.microsoft.com/en-us/office/seriessum-function-a3ab25b5-1093-4f5b-b084-96c49087f637

    #[test]
    fn test_seriessum_docs() {
        let ctx = EvaluationContext::simple();
        // Approximate COS(PI/4) using first 4 terms of Taylor series:
        // SERIESSUM(PI()/4, 0, 2, {1, -1/FACT(2), 1/FACT(4), -1/FACT(6)})
        let x = std::f64::consts::FRAC_PI_4;
        let coeffs = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(1.0),
            FormulaValue::Number(-1.0 / 2.0),   // -1/2!
            FormulaValue::Number(1.0 / 24.0),   // 1/4!
            FormulaValue::Number(-1.0 / 720.0), // -1/6!
        ]], source: None };
        let s = num(fn_seriessum(
            &[
                FormulaValue::Number(x),
                FormulaValue::Number(0.0),
                FormulaValue::Number(2.0),
                coeffs,
            ],
            &ctx,
        )
        .unwrap());
        // Should approximate cos(pi/4) ≈ 0.70710678
        assert!((s - x.cos()).abs() < 0.001);
    }

    // ── SUMX2MY2 docs ──
    // https://support.microsoft.com/en-us/office/sumx2my2-function-9e599cc5-5399-48e9-a5e0-e87571bf4c05

    #[test]
    fn test_sumx2my2_docs() {
        let ctx = EvaluationContext::simple();
        // Docs: array_x = {2,3,9,1,8,7,5}, array_y = {6,5,11,7,5,4,4}
        // SUM(x²-y²) = (4-36)+(9-25)+(81-121)+(1-49)+(64-25)+(49-16)+(25-16) = -55
        let x = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(2.0),
            FormulaValue::Number(3.0),
            FormulaValue::Number(9.0),
            FormulaValue::Number(1.0),
            FormulaValue::Number(8.0),
            FormulaValue::Number(7.0),
            FormulaValue::Number(5.0),
        ]], source: None };
        let y = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(6.0),
            FormulaValue::Number(5.0),
            FormulaValue::Number(11.0),
            FormulaValue::Number(7.0),
            FormulaValue::Number(5.0),
            FormulaValue::Number(4.0),
            FormulaValue::Number(4.0),
        ]], source: None };
        assert_eq!(
            fn_sumx2my2(&[x, y], &ctx).unwrap(),
            FormulaValue::Number(-55.0)
        );
    }

    // ── SUMX2PY2 docs ──
    // https://support.microsoft.com/en-us/office/sumx2py2-function-826b60b4-0aa2-4571-8d33-3b8b97f73584

    #[test]
    fn test_sumx2py2_docs() {
        let ctx = EvaluationContext::simple();
        // Docs: same data as SUMX2MY2
        // SUM(x²+y²) = (4+36)+(9+25)+(81+121)+(1+49)+(64+25)+(49+16)+(25+16) = 521
        let x = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(2.0),
            FormulaValue::Number(3.0),
            FormulaValue::Number(9.0),
            FormulaValue::Number(1.0),
            FormulaValue::Number(8.0),
            FormulaValue::Number(7.0),
            FormulaValue::Number(5.0),
        ]], source: None };
        let y = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(6.0),
            FormulaValue::Number(5.0),
            FormulaValue::Number(11.0),
            FormulaValue::Number(7.0),
            FormulaValue::Number(5.0),
            FormulaValue::Number(4.0),
            FormulaValue::Number(4.0),
        ]], source: None };
        assert_eq!(
            fn_sumx2py2(&[x, y], &ctx).unwrap(),
            FormulaValue::Number(521.0)
        );
    }

    // ── SUMXMY2 docs ──
    // https://support.microsoft.com/en-us/office/sumxmy2-function-9d144ac1-4d79-43de-b524-e2ecee23b299

    #[test]
    fn test_sumxmy2_docs() {
        let ctx = EvaluationContext::simple();
        // Docs: same data
        // SUM((x-y)²) = (2-6)²+(3-5)²+...+(5-4)² = 16+4+4+36+9+9+1 = 79
        let x = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(2.0),
            FormulaValue::Number(3.0),
            FormulaValue::Number(9.0),
            FormulaValue::Number(1.0),
            FormulaValue::Number(8.0),
            FormulaValue::Number(7.0),
            FormulaValue::Number(5.0),
        ]], source: None };
        let y = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(6.0),
            FormulaValue::Number(5.0),
            FormulaValue::Number(11.0),
            FormulaValue::Number(7.0),
            FormulaValue::Number(5.0),
            FormulaValue::Number(4.0),
            FormulaValue::Number(4.0),
        ]], source: None };
        assert_eq!(
            fn_sumxmy2(&[x, y], &ctx).unwrap(),
            FormulaValue::Number(79.0)
        );
    }

    // ── MDETERM docs ──
    // https://support.microsoft.com/en-us/office/mdeterm-function-e7bfa857-3532-4ae0-8a87-02aaef2c8b17

    #[test]
    fn test_mdeterm_docs() {
        let ctx = EvaluationContext::simple();
        // Docs: {1,3;7,2} -> det = 1*2 - 3*7 = -19
        let m = FormulaValue::Array { data: vec![
            vec![FormulaValue::Number(1.0), FormulaValue::Number(3.0)],
            vec![FormulaValue::Number(7.0), FormulaValue::Number(2.0)],
        ], source: None };
        assert!((num(fn_mdeterm(std::slice::from_ref(&m), &ctx).unwrap()) - (-19.0)).abs() < 1e-10);
        // 3x3: {3,6,1;1,1,0;-1,2,3} -> det = 3*(3-0) - 6*(3-0) + 1*(2-(-1)) = 9-18+3 = -6
        let m3 = FormulaValue::Array { data: vec![
            vec![
                FormulaValue::Number(3.0),
                FormulaValue::Number(6.0),
                FormulaValue::Number(1.0),
            ],
            vec![
                FormulaValue::Number(1.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(0.0),
            ],
            vec![
                FormulaValue::Number(-1.0),
                FormulaValue::Number(2.0),
                FormulaValue::Number(3.0),
            ],
        ], source: None };
        assert!((num(fn_mdeterm(std::slice::from_ref(&m3), &ctx).unwrap()) - (-6.0)).abs() < 1e-10);
        // Non-square matrix -> #VALUE!
        let nonsq = FormulaValue::Array { data: vec![
            vec![
                FormulaValue::Number(1.0),
                FormulaValue::Number(2.0),
                FormulaValue::Number(3.0),
            ],
            vec![
                FormulaValue::Number(4.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(6.0),
            ],
        ], source: None };
        assert_eq!(
            fn_mdeterm(std::slice::from_ref(&nonsq), &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    // ── MINVERSE docs ──
    // https://support.microsoft.com/en-us/office/minverse-function-11f55086-adde-4c9f-8eb9-59da2d72efc6

    #[test]
    fn test_minverse_docs() {
        let ctx = EvaluationContext::simple();
        // {4, -1; 2, 0} -> inverse is {0, 0.5; -1, 2} (det=2)
        let m = FormulaValue::Array { data: vec![
            vec![FormulaValue::Number(4.0), FormulaValue::Number(-1.0)],
            vec![FormulaValue::Number(2.0), FormulaValue::Number(0.0)],
        ], source: None };
        let inv = fn_minverse(std::slice::from_ref(&m), &ctx).unwrap();
        if let FormulaValue::Array { data: rows, .. } = inv {
            assert!((num(rows[0][0].clone()) - 0.0).abs() < 1e-10);
            assert!((num(rows[0][1].clone()) - 0.5).abs() < 1e-10);
            assert!((num(rows[1][0].clone()) - (-1.0)).abs() < 1e-10);
            assert!((num(rows[1][1].clone()) - 2.0).abs() < 1e-10);
        } else {
            panic!("expected array");
        }
        // {1, 2; 3, 4} -> inverse is {-2, 1; 1.5, -0.5} (det=-2)
        let m2 = FormulaValue::Array { data: vec![
            vec![FormulaValue::Number(1.0), FormulaValue::Number(2.0)],
            vec![FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
        ], source: None };
        let inv2 = fn_minverse(std::slice::from_ref(&m2), &ctx).unwrap();
        if let FormulaValue::Array { data: rows, .. } = inv2 {
            assert!((num(rows[0][0].clone()) - (-2.0)).abs() < 1e-10);
            assert!((num(rows[0][1].clone()) - 1.0).abs() < 1e-10);
            assert!((num(rows[1][0].clone()) - 1.5).abs() < 1e-10);
            assert!((num(rows[1][1].clone()) - (-0.5)).abs() < 1e-10);
        } else {
            panic!("expected array");
        }
    }

    // ── MMULT docs ──
    // https://support.microsoft.com/en-us/office/mmult-function-40593ed7-a3cd-4b6b-b9a3-e4ad3c7245eb

    #[test]
    fn test_mmult_docs() {
        let ctx = EvaluationContext::simple();
        // {1, 3; 7, 2} * {2, 0; 0, 2} = {2, 6; 14, 4}
        let a = FormulaValue::Array { data: vec![
            vec![FormulaValue::Number(1.0), FormulaValue::Number(3.0)],
            vec![FormulaValue::Number(7.0), FormulaValue::Number(2.0)],
        ], source: None };
        let b = FormulaValue::Array { data: vec![
            vec![FormulaValue::Number(2.0), FormulaValue::Number(0.0)],
            vec![FormulaValue::Number(0.0), FormulaValue::Number(2.0)],
        ], source: None };
        let result = fn_mmult(&[a, b], &ctx).unwrap();
        if let FormulaValue::Array { data: rows, .. } = result {
            assert_eq!(rows[0][0], FormulaValue::Number(2.0));
            assert_eq!(rows[0][1], FormulaValue::Number(6.0));
            assert_eq!(rows[1][0], FormulaValue::Number(14.0));
            assert_eq!(rows[1][1], FormulaValue::Number(4.0));
        } else {
            panic!("expected array");
        }
    }

    // ── MUNIT docs ──
    // https://support.microsoft.com/en-us/office/munit-function-c9fe916a-dc26-4105-997d-ba22799853a3

    #[test]
    fn test_munit_docs() {
        let ctx = EvaluationContext::simple();
        // MUNIT(3) = {{1,0,0};{0,1,0};{0,0,1}}
        let id = fn_munit(&[FormulaValue::Number(3.0)], &ctx).unwrap();
        if let FormulaValue::Array { data: rows, .. } = id {
            assert_eq!(rows.len(), 3);
            for i in 0..3 {
                for j in 0..3 {
                    let expected = if i == j { 1.0 } else { 0.0 };
                    assert_eq!(rows[i][j], FormulaValue::Number(expected));
                }
            }
        } else {
            panic!("expected array");
        }
    }

    // ── AGGREGATE docs ──
    // https://support.microsoft.com/en-us/office/aggregate-function-43b9278e-6aa7-4f17-92b6-e19993fa26df

    #[test]
    fn test_aggregate_docs() {
        let ctx = EvaluationContext::simple();
        // Data with errors: 72, 50, 96, 57, 83
        let data = FormulaValue::Array { data: vec![
            vec![FormulaValue::Number(72.0)],
            vec![FormulaValue::Number(50.0)],
            vec![FormulaValue::Number(96.0)],
            vec![FormulaValue::Number(57.0)],
            vec![FormulaValue::Number(83.0)],
        ], source: None };
        // AGGREGATE(4, 6, data) = MAX ignoring errors = 96
        assert_eq!(
            num(fn_aggregate(
                &[
                    FormulaValue::Number(4.0),
                    FormulaValue::Number(6.0),
                    data.clone()
                ],
                &ctx
            )
            .unwrap()),
            96.0
        );
        // AGGREGATE(5, 6, data) = MIN = 50
        assert_eq!(
            num(fn_aggregate(
                &[
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(6.0),
                    data.clone()
                ],
                &ctx
            )
            .unwrap()),
            50.0
        );
        // AGGREGATE(9, 6, data) = SUM = 358
        assert_eq!(
            num(fn_aggregate(
                &[
                    FormulaValue::Number(9.0),
                    FormulaValue::Number(6.0),
                    data.clone()
                ],
                &ctx
            )
            .unwrap()),
            358.0
        );
        // AGGREGATE(1, 6, data) = AVERAGE = 71.6
        assert!(
            (num(fn_aggregate(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(6.0),
                    data.clone()
                ],
                &ctx
            )
            .unwrap())
                - 71.6)
                .abs()
                < 1e-10
        );
        // AGGREGATE(14, 6, data, 2) = LARGE k=2 = 83
        assert_eq!(
            num(fn_aggregate(
                &[
                    FormulaValue::Number(14.0),
                    FormulaValue::Number(6.0),
                    data.clone(),
                    FormulaValue::Number(2.0)
                ],
                &ctx
            )
            .unwrap()),
            83.0
        );
        // AGGREGATE(15, 6, data, 3) = SMALL k=3 = 72
        assert_eq!(
            num(fn_aggregate(
                &[
                    FormulaValue::Number(15.0),
                    FormulaValue::Number(6.0),
                    data,
                    FormulaValue::Number(3.0)
                ],
                &ctx
            )
            .unwrap()),
            72.0
        );
    }

    // ── SUBTOTAL docs ──
    // https://support.microsoft.com/en-us/office/subtotal-function-7b027003-f060-4ade-9040-e478765b9939

    #[test]
    fn test_subtotal_docs() {
        let ctx = EvaluationContext::simple();
        // Data: 120, 10, 150, 23
        // SUBTOTAL(9, ...) = SUM = 303
        assert_eq!(
            num(fn_subtotal(
                &[
                    FormulaValue::Number(9.0),
                    FormulaValue::Number(120.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(150.0),
                    FormulaValue::Number(23.0),
                ],
                &ctx
            )
            .unwrap()),
            303.0
        );
        // SUBTOTAL(1, ...) = AVERAGE = 75.75
        assert!(
            (num(fn_subtotal(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(120.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(150.0),
                    FormulaValue::Number(23.0),
                ],
                &ctx
            )
            .unwrap())
                - 75.75)
                .abs()
                < 1e-10
        );
        // SUBTOTAL(2, ...) = COUNT = 4
        assert_eq!(
            num(fn_subtotal(
                &[
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(120.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(150.0),
                    FormulaValue::Number(23.0),
                ],
                &ctx
            )
            .unwrap()),
            4.0
        );
        // SUBTOTAL(4, ...) = MAX = 150
        assert_eq!(
            num(fn_subtotal(
                &[
                    FormulaValue::Number(4.0),
                    FormulaValue::Number(120.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(150.0),
                    FormulaValue::Number(23.0),
                ],
                &ctx
            )
            .unwrap()),
            150.0
        );
        // SUBTOTAL(5, ...) = MIN = 10
        assert_eq!(
            num(fn_subtotal(
                &[
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(120.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(150.0),
                    FormulaValue::Number(23.0),
                ],
                &ctx
            )
            .unwrap()),
            10.0
        );
    }

    // ── RANDARRAY docs ──
    // https://support.microsoft.com/en-us/office/randarray-function-21261e55-3bec-4885-86a6-8b0a47fd4d33

    #[test]
    fn test_randarray_docs() {
        let ctx = EvaluationContext::simple();
        // Default: single random number in [0, 1)
        let one = fn_randarray(&[], &ctx).unwrap();
        if let FormulaValue::Number(n) = one {
            assert!(n >= 0.0 && n < 1.0);
        } else {
            panic!("expected number for default RANDARRAY");
        }
        // RANDARRAY(3, 4) -> 3 rows x 4 cols of floats in [0, 1)
        let arr = fn_randarray(
            &[FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
            &ctx,
        )
        .unwrap();
        if let FormulaValue::Array { data: rows, .. } = arr {
            assert_eq!(rows.len(), 3);
            assert_eq!(rows[0].len(), 4);
        } else {
            panic!("expected array");
        }
        // RANDARRAY(5, 1, 1, 100, TRUE) -> 5 integers in [1, 100]
        let arr = fn_randarray(
            &[
                FormulaValue::Number(5.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(100.0),
                FormulaValue::Boolean(true),
            ],
            &ctx,
        )
        .unwrap();
        if let FormulaValue::Array { data: rows, .. } = arr {
            assert_eq!(rows.len(), 5);
            for row in &rows {
                if let FormulaValue::Number(n) = row[0] {
                    assert!(n >= 1.0 && n <= 100.0);
                    assert_eq!(n, n.floor()); // must be integer
                }
            }
        } else {
            panic!("expected array");
        }
        // RANDARRAY(1, 1, 100, 1) -> min > max = #NUM!
        let err = fn_randarray(
            &[
                FormulaValue::Number(1.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(100.0),
                FormulaValue::Number(1.0),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(err, FormulaValue::Error(CellError::Num));
    }

    // ── PERCENTOF docs ──
    // https://support.microsoft.com/en-us/office/percentof-function-7c66da0a-ac30-45d0-bfc7-834a8bd7c962

    #[test]
    fn test_percentof_docs() {
        let ctx = EvaluationContext::simple();
        // Docs Example 1: =PERCENTOF(C3:C4, C3:C14)
        // Subset: Bib-Shorts(7600) + Bike Racks(56100) = 63700
        // All: 7600+56100+2100+11100+55900+59600+74800+48300+9700+61400+7000+48500 = 442100
        // Result: 63700/442100 ≈ 0.14407 (displayed as 14.41%)
        let subset = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(7600.0),
            FormulaValue::Number(56100.0),
        ]], source: None };
        let all = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(7600.0),
            FormulaValue::Number(56100.0),
            FormulaValue::Number(2100.0),
            FormulaValue::Number(11100.0),
            FormulaValue::Number(55900.0),
            FormulaValue::Number(59600.0),
            FormulaValue::Number(74800.0),
            FormulaValue::Number(48300.0),
            FormulaValue::Number(9700.0),
            FormulaValue::Number(61400.0),
            FormulaValue::Number(7000.0),
            FormulaValue::Number(48500.0),
        ]], source: None };
        let result = num(fn_percentof(&[subset, all], &ctx).unwrap());
        // 63700 / 442100 = 0.144073717...
        assert!((result - 63700.0 / 442100.0).abs() < 1e-10);

        // Verify it matches SUM(subset)/SUM(all) semantics
        // Simple case: single values
        assert_eq!(
            num(fn_percentof(
                &[FormulaValue::Number(25.0), FormulaValue::Number(100.0)],
                &ctx
            )
            .unwrap()),
            0.25
        );

        // Half of total
        let half_sub = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(10.0),
            FormulaValue::Number(10.0),
        ]], source: None };
        let half_all = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(10.0),
            FormulaValue::Number(10.0),
            FormulaValue::Number(10.0),
            FormulaValue::Number(10.0),
        ]], source: None };
        assert!((num(fn_percentof(&[half_sub, half_all], &ctx).unwrap()) - 0.5).abs() < 1e-10);

        // Division by zero: all values sum to 0 -> #DIV/0!
        assert_eq!(
            fn_percentof(
                &[FormulaValue::Number(0.0), FormulaValue::Number(0.0)],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Div0)
        );

        // Subset is zero, all is non-zero -> 0.0
        assert_eq!(
            num(fn_percentof(
                &[FormulaValue::Number(0.0), FormulaValue::Number(100.0)],
                &ctx
            )
            .unwrap()),
            0.0
        );

        // Error propagation from subset
        assert_eq!(
            fn_percentof(
                &[
                    FormulaValue::Error(CellError::Value),
                    FormulaValue::Number(100.0)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // Error propagation from all
        assert_eq!(
            fn_percentof(
                &[
                    FormulaValue::Number(25.0),
                    FormulaValue::Error(CellError::Na)
                ],
                &ctx
            )
            .unwrap(),
            FormulaValue::Error(CellError::Na)
        );

        // Subset equals all -> 100%
        let same = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(10.0),
            FormulaValue::Number(20.0),
            FormulaValue::Number(30.0),
        ]], source: None };
        assert!((num(fn_percentof(&[same.clone(), same], &ctx).unwrap()) - 1.0).abs() < 1e-10);

        // Non-numeric values in arrays are ignored (text, booleans, empty)
        let subset_mixed = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(10.0),
            FormulaValue::String("text".to_string()),
            FormulaValue::Boolean(true),
            FormulaValue::Empty,
        ]], source: None };
        let all_mixed = FormulaValue::Array { data: vec![vec![
            FormulaValue::Number(10.0),
            FormulaValue::Number(40.0),
            FormulaValue::String("ignored".to_string()),
        ]], source: None };
        assert!((num(fn_percentof(&[subset_mixed, all_mixed], &ctx).unwrap()) - 0.2).abs() < 1e-10);
    }
}
