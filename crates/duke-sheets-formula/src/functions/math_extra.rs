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
        FormulaValue::Array(arr) => {
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

fn collect_values(value: &FormulaValue, out: &mut Vec<FormulaValue>) -> Option<CellError> {
    match value {
        FormulaValue::Array(arr) => {
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
        FormulaValue::Array(rows) => {
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
        FormulaValue::Array(rows) => {
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
            FormulaValue::Array(rows) => {
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

fn roman_digit(value: i64, one: &str, five: &str, ten: &str, form: i64) -> String {
    let mut out = String::new();
    if form == 0 {
        match value {
            0 => {}
            1..=3 => out.push_str(&one.repeat(value as usize)),
            4 => out.push_str(&(String::from(one) + five)),
            5 => out.push_str(five),
            6..=8 => out.push_str(&(String::from(five) + &one.repeat((value - 5) as usize))),
            9 => out.push_str(&(String::from(one) + ten)),
            _ => {}
        }
        return out;
    }

    match value {
        0 => {}
        1..=3 => out.push_str(&one.repeat(value as usize)),
        4 => {
            if form >= 4 {
                out.push_str(&one.repeat(4));
            } else {
                out.push_str(&(String::from(one) + five));
            }
        }
        5 => out.push_str(five),
        6..=8 => out.push_str(&(String::from(five) + &one.repeat((value - 5) as usize))),
        9 => {
            if form >= 3 {
                out.push_str(&(String::from(five) + &one.repeat(4)));
            } else {
                out.push_str(&(String::from(one) + ten));
            }
        }
        _ => {}
    }
    out
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
    let thousands = n / 1000;
    let hundreds = (n / 100) % 10;
    let tens = (n / 10) % 10;
    let ones = n % 10;

    let mut s = String::new();
    s.push_str(&"M".repeat(thousands as usize));
    s.push_str(&roman_digit(hundreds, "C", "D", "M", form));
    s.push_str(&roman_digit(tens, "X", "L", "C", form));
    s.push_str(&roman_digit(ones, "I", "V", "X", form));
    Ok(FormulaValue::String(s))
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
    Ok(FormulaValue::Array(inv))
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
    Ok(FormulaValue::Array(out))
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
    Ok(FormulaValue::Array(out))
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
    Ok(FormulaValue::Array(out))
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
    match get_number(args.get(1)) {
        Ok(_) => {}
        Err(e) => return Ok(e),
    }

    let mut numbers = Vec::new();
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
        4 => Ok(FormulaValue::Number(
            numbers
                .iter()
                .copied()
                .fold(f64::NEG_INFINITY, f64::max)
                .max(0.0),
        )),
        5 => Ok(FormulaValue::Number(
            numbers
                .iter()
                .copied()
                .fold(f64::INFINITY, f64::min)
                .min(0.0),
        )),
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
    fn_aggregate(
        &[
            FormulaValue::Number(normalized as f64),
            FormulaValue::Number(0.0),
            FormulaValue::Array(args[1..].iter().cloned().map(|v| vec![v]).collect()),
        ],
        _ctx,
    )
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
            FormulaValue::String("CCCCLXXXXVIIII".to_string())
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
        let matrix = FormulaValue::Array(vec![
            vec![FormulaValue::Number(1.0), FormulaValue::Number(2.0)],
            vec![FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
        ]);
        let det = num(fn_mdeterm(std::slice::from_ref(&matrix), &ctx).unwrap());
        assert!((det + 2.0).abs() < 1e-10);
        let inv = fn_minverse(std::slice::from_ref(&matrix), &ctx).unwrap();
        if let FormulaValue::Array(rows) = inv {
            if let FormulaValue::Number(v) = rows[0][0] {
                assert!((v + 2.0).abs() < 1e-10);
            } else {
                panic!("expected number");
            }
        } else {
            panic!("expected array");
        }
        let mm = fn_mmult(&[matrix.clone(), matrix], &ctx).unwrap();
        if let FormulaValue::Array(rows) = mm {
            assert_eq!(rows[0][0], FormulaValue::Number(7.0));
            assert_eq!(rows[1][1], FormulaValue::Number(22.0));
        } else {
            panic!("expected array");
        }
        let id = fn_munit(&[FormulaValue::Number(3.0)], &ctx).unwrap();
        if let FormulaValue::Array(rows) = id {
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
        if let FormulaValue::Array(rows) = arr {
            assert_eq!(rows.len(), 2);
            assert_eq!(rows[0].len(), 3);
        } else {
            panic!("expected array");
        }
        let coeffs = FormulaValue::Array(vec![vec![
            FormulaValue::Number(1.0),
            FormulaValue::Number(2.0),
            FormulaValue::Number(3.0),
        ]]);
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
        let x = FormulaValue::Array(vec![vec![
            FormulaValue::Number(1.0),
            FormulaValue::Number(2.0),
        ]]);
        let y = FormulaValue::Array(vec![vec![
            FormulaValue::Number(3.0),
            FormulaValue::Number(4.0),
        ]]);
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
        let data = FormulaValue::Array(vec![
            vec![FormulaValue::Number(1.0), FormulaValue::Number(2.0)],
            vec![FormulaValue::Number(3.0), FormulaValue::Number(4.0)],
        ]);
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
}
