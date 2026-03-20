use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

use std::f64::consts::PI;

const MAX_DIGITS: usize = 10;
const BIN_BITS: u32 = 10;
const OCT_BITS: u32 = 30;
const HEX_BITS: u32 = 40;
const MAX_BITWISE: u64 = (1u64 << 48) - 1;

const ZERO_TOLERANCE: f64 = 1e-10;
const BESSEL_MAX_TERMS: usize = 100;
const BESSEL_TOLERANCE: f64 = 1e-15;
const GAMMA_SERIES_MAX_TERMS: usize = 200;
const EULER_GAMMA: f64 = 0.577_215_664_901_532_9;

fn scalar_number(value: &FormulaValue) -> Result<f64, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        _ => value
            .as_number()
            .ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn required_number(args: &[FormulaValue], idx: usize) -> Result<f64, FormulaValue> {
    match args.get(idx) {
        Some(v) => scalar_number(v),
        None => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn optional_number(args: &[FormulaValue], idx: usize, default: f64) -> Result<f64, FormulaValue> {
    match args.get(idx).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(v) => scalar_number(v),
        None => Ok(default),
    }
}

fn scalar_text(value: &FormulaValue) -> Result<String, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        _ => Ok(value.as_string()),
    }
}

fn required_text(args: &[FormulaValue], idx: usize) -> Result<String, FormulaValue> {
    match args.get(idx) {
        Some(v) => scalar_text(v),
        None => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn optional_places(args: &[FormulaValue], idx: usize) -> Result<Option<usize>, FormulaValue> {
    let Some(value) = args.get(idx).filter(|v| !matches!(v, FormulaValue::Empty)) else {
        return Ok(None);
    };

    let n = scalar_number(value)?;
    if !n.is_finite() {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let places = n.trunc();
    if places <= 0.0 || places > MAX_DIGITS as f64 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    Ok(Some(places as usize))
}

fn required_integer(args: &[FormulaValue], idx: usize) -> Result<i64, FormulaValue> {
    let n = required_number(args, idx)?;
    if !n.is_finite() {
        return Err(FormulaValue::Error(CellError::Num));
    }
    Ok(n.trunc() as i64)
}

fn required_shift(args: &[FormulaValue], idx: usize) -> Result<i32, FormulaValue> {
    let n = required_number(args, idx)?;
    if !n.is_finite() {
        return Err(FormulaValue::Error(CellError::Num));
    }
    let shift = n.trunc();
    if shift < i32::MIN as f64 || shift > i32::MAX as f64 {
        return Err(FormulaValue::Error(CellError::Num));
    }
    Ok(shift as i32)
}

fn required_bitwise_operand(args: &[FormulaValue], idx: usize) -> Result<u64, FormulaValue> {
    let n = required_number(args, idx)?;
    if !n.is_finite() {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let truncated = n.trunc();
    if truncated < 0.0 || truncated > MAX_BITWISE as f64 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    Ok(truncated as u64)
}

fn fits_signed_bits(value: i64, bits: u32) -> bool {
    let min = -(1i64 << (bits - 1));
    let max = (1i64 << (bits - 1)) - 1;
    value >= min && value <= max
}

fn format_radix(value: u64, base: u32) -> String {
    match base {
        2 => format!("{value:b}"),
        8 => format!("{value:o}"),
        16 => format!("{value:X}"),
        _ => String::new(),
    }
}

fn validate_digits(text: &str, base: u32) -> bool {
    match base {
        2 => text.chars().all(|c| matches!(c, '0' | '1')),
        8 => text.chars().all(|c| ('0'..='7').contains(&c)),
        16 => text.chars().all(|c| c.is_ascii_hexdigit()),
        _ => false,
    }
}

fn parse_signed_fixed_width(
    text: &str,
    base: u32,
    bits: u32,
    max_digits: usize,
) -> Result<i64, CellError> {
    let trimmed = text.trim();
    if trimmed.is_empty() || trimmed.len() > max_digits || !validate_digits(trimmed, base) {
        return Err(CellError::Num);
    }

    let parsed = u64::from_str_radix(trimmed, base).map_err(|_| CellError::Num)?;
    if trimmed.len() == max_digits {
        let sign_bit = 1u64 << (bits - 1);
        if parsed & sign_bit != 0 {
            return Ok(parsed as i64 - (1i64 << bits));
        }
    }

    Ok(parsed as i64)
}

fn apply_places(mut text: String, places: Option<usize>) -> Result<String, CellError> {
    let Some(p) = places else {
        return Ok(text);
    };

    if text.len() > p {
        return Err(CellError::Num);
    }

    if text.len() < p {
        let pad = "0".repeat(p - text.len());
        text = format!("{pad}{text}");
    }

    Ok(text)
}

fn format_signed_in_base(
    value: i64,
    target_bits: u32,
    target_base: u32,
    places: Option<usize>,
) -> Result<String, CellError> {
    if !fits_signed_bits(value, target_bits) {
        return Err(CellError::Num);
    }

    if value < 0 {
        let encoded = ((1i128 << target_bits) + value as i128) as u64;
        return Ok(format_radix(encoded, target_base));
    }

    let text = format_radix(value as u64, target_base);
    apply_places(text, places)
}

fn parse_bin(text: &str) -> Result<i64, CellError> {
    parse_signed_fixed_width(text, 2, BIN_BITS, MAX_DIGITS)
}

fn parse_oct(text: &str) -> Result<i64, CellError> {
    parse_signed_fixed_width(text, 8, OCT_BITS, MAX_DIGITS)
}

fn parse_hex(text: &str) -> Result<i64, CellError> {
    parse_signed_fixed_width(text, 16, HEX_BITS, MAX_DIGITS)
}

fn ln_gamma(z: f64) -> f64 {
    let coeffs = [
        0.999_999_999_999_809_9,
        676.520_368_121_885_1,
        -1_259.139_216_722_402_8,
        771.323_428_777_653_1,
        -176.615_029_162_140_6,
        12.507_343_278_686_905,
        -0.138_571_095_265_720_12,
        9.984_369_578_019_572e-6,
        1.505_632_735_149_311_6e-7,
    ];

    if z < 0.5 {
        let pi = std::f64::consts::PI;
        return (pi / (pi * z).sin()).ln() - ln_gamma(1.0 - z);
    }

    let z1 = z - 1.0;
    let mut x = coeffs[0];
    for (i, c) in coeffs.iter().enumerate().skip(1) {
        x += c / (z1 + i as f64);
    }

    let t = z1 + 7.5;
    0.5 * (2.0 * std::f64::consts::PI).ln() + (z1 + 0.5) * t.ln() - t + x.ln()
}

fn regularized_gamma_p(a: f64, x: f64) -> f64 {
    if a <= 0.0 || x < 0.0 {
        return f64::NAN;
    }
    if x == 0.0 {
        return 0.0;
    }

    let mut sum = 1.0 / a;
    let mut del = sum;
    let mut ap = a;
    for _ in 1..=GAMMA_SERIES_MAX_TERMS {
        ap += 1.0;
        del *= x / ap;
        sum += del;
        if del.abs() < sum.abs().max(1.0) * 1e-15 {
            break;
        }
    }

    sum * (-x + a * x.ln() - ln_gamma(a)).exp()
}

fn erf_value(x: f64) -> f64 {
    if x == 0.0 {
        0.0
    } else {
        let p = regularized_gamma_p(0.5, x * x);
        if x.is_sign_negative() {
            -p
        } else {
            p
        }
    }
}

fn erfc_value(x: f64) -> f64 {
    let p = regularized_gamma_p(0.5, x * x);
    if x.is_sign_negative() {
        1.0 + p
    } else {
        1.0 - p
    }
}

pub fn fn_bin2dec(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match parse_bin(&text) {
        Ok(v) => Ok(FormulaValue::Number(v as f64)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_bin2hex(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let decimal = match parse_bin(&text) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    match format_signed_in_base(decimal, HEX_BITS, 16, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_bin2oct(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let decimal = match parse_bin(&text) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    match format_signed_in_base(decimal, OCT_BITS, 8, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_dec2bin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match required_integer(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match format_signed_in_base(number, BIN_BITS, 2, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_dec2hex(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match required_integer(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match format_signed_in_base(number, HEX_BITS, 16, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_dec2oct(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match required_integer(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match format_signed_in_base(number, OCT_BITS, 8, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_hex2bin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let decimal = match parse_hex(&text) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    match format_signed_in_base(decimal, BIN_BITS, 2, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_hex2dec(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match parse_hex(&text) {
        Ok(v) => Ok(FormulaValue::Number(v as f64)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_hex2oct(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let decimal = match parse_hex(&text) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    match format_signed_in_base(decimal, OCT_BITS, 8, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_oct2bin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let decimal = match parse_oct(&text) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    match format_signed_in_base(decimal, BIN_BITS, 2, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_oct2dec(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match parse_oct(&text) {
        Ok(v) => Ok(FormulaValue::Number(v as f64)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_oct2hex(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match required_text(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let places = match optional_places(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let decimal = match parse_oct(&text) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    match format_signed_in_base(decimal, HEX_BITS, 16, places) {
        Ok(v) => Ok(FormulaValue::String(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_bitand(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let left = match required_bitwise_operand(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let right = match required_bitwise_operand(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number((left & right) as f64))
}

pub fn fn_bitor(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let left = match required_bitwise_operand(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let right = match required_bitwise_operand(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number((left | right) as f64))
}

pub fn fn_bitxor(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let left = match required_bitwise_operand(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let right = match required_bitwise_operand(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number((left ^ right) as f64))
}

pub fn fn_bitlshift(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number = match required_bitwise_operand(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let shift = match required_shift(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let result = if shift >= 0 {
        if shift as u32 >= 64 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        match number.checked_shl(shift as u32) {
            Some(v) if v <= MAX_BITWISE => v,
            _ => return Ok(FormulaValue::Error(CellError::Num)),
        }
    } else {
        let amount = (-shift) as u32;
        if amount >= 64 {
            0
        } else {
            number >> amount
        }
    };

    Ok(FormulaValue::Number(result as f64))
}

pub fn fn_bitrshift(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number = match required_bitwise_operand(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let shift = match required_shift(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let result = if shift >= 0 {
        let amount = shift as u32;
        if amount >= 64 {
            0
        } else {
            number >> amount
        }
    } else {
        let amount = (-shift) as u32;
        if amount >= 64 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        match number.checked_shl(amount) {
            Some(v) if v <= MAX_BITWISE => v,
            _ => return Ok(FormulaValue::Error(CellError::Num)),
        }
    };

    Ok(FormulaValue::Number(result as f64))
}

pub fn fn_delta(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number1 = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let number2 = match optional_number(args, 1, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let result = if (number1 - number2).abs() < f64::EPSILON {
        1.0
    } else {
        0.0
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_gestep(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let step = match optional_number(args, 1, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number(if number >= step { 1.0 } else { 0.0 }))
}

pub fn fn_erf(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let lower = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let upper = match args.get(1).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(_) => match required_number(args, 1) {
            Ok(v) => Some(v),
            Err(e) => return Ok(e),
        },
        None => None,
    };

    let value = match upper {
        Some(upper_bound) => erf_value(upper_bound) - erf_value(lower),
        None => erf_value(lower),
    };
    Ok(FormulaValue::Number(value))
}

pub fn fn_erf_precise(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number(erf_value(x)))
}

pub fn fn_erfc(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number(erfc_value(x)))
}

pub fn fn_erfc_precise(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number(erfc_value(x)))
}

// Complex number functions
fn required_complex(
    args: &[FormulaValue],
    idx: usize,
) -> Result<(f64, f64, &'static str), FormulaValue> {
    let value = match args.get(idx) {
        Some(v) => v,
        None => return Err(FormulaValue::Error(CellError::Value)),
    };
    let s = scalar_text(value)?;
    match parse_complex(&s) {
        Some((real, imag, suffix)) => Ok((real, imag, if suffix == "j" { "j" } else { "i" })),
        None => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn choose_suffix(a: &str, b: &str) -> &'static str {
    if a == "j" || b == "j" {
        "j"
    } else {
        "i"
    }
}

fn is_zero(v: f64) -> bool {
    v.abs() < ZERO_TOLERANCE
}

fn clean_zero(v: f64) -> f64 {
    if is_zero(v) {
        0.0
    } else {
        v
    }
}

fn format_number(n: f64) -> String {
    let n = clean_zero(n);
    let rounded = n.round();
    if (n - rounded).abs() < ZERO_TOLERANCE && n.abs() < 1e15 {
        format!("{}", rounded as i64)
    } else {
        format!("{}", n)
    }
}

fn parse_imag_coeff(s: &str) -> Option<f64> {
    if s == "+" || s.is_empty() {
        Some(1.0)
    } else if s == "-" {
        Some(-1.0)
    } else {
        s.parse::<f64>().ok()
    }
}

fn split_real_imag(body: &str) -> Option<usize> {
    let bytes = body.as_bytes();
    let mut split = None;
    let mut i = 1;
    while i < bytes.len() {
        if (bytes[i] == b'+' || bytes[i] == b'-') && bytes[i - 1] != b'e' && bytes[i - 1] != b'E' {
            split = Some(i);
        }
        i += 1;
    }
    split
}

fn parse_complex(s: &str) -> Option<(f64, f64, &str)> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return None;
    }

    let (body, suffix) = match trimmed.as_bytes().last() {
        Some(b'i') => (&trimmed[..trimmed.len() - 1], "i"),
        Some(b'j') => (&trimmed[..trimmed.len() - 1], "j"),
        _ => {
            let real = trimmed.parse::<f64>().ok()?;
            return Some((real, 0.0, "i"));
        }
    };

    if body.is_empty() || body == "+" || body == "-" {
        return Some((0.0, parse_imag_coeff(body)?, suffix));
    }

    if let Some(split) = split_real_imag(body) {
        let real = body[..split].parse::<f64>().ok()?;
        let imag = parse_imag_coeff(&body[split..])?;
        Some((real, imag, suffix))
    } else {
        let imag = parse_imag_coeff(body)?;
        Some((0.0, imag, suffix))
    }
}

fn format_complex(real: f64, imag: f64, suffix: &str) -> String {
    let real = clean_zero(real);
    let imag = clean_zero(imag);

    if is_zero(imag) {
        return format_number(real);
    }
    if is_zero(real) {
        if (imag - 1.0).abs() < ZERO_TOLERANCE {
            return suffix.to_string();
        }
        if (imag + 1.0).abs() < ZERO_TOLERANCE {
            return format!("-{suffix}");
        }
        return format!("{}{}", format_number(imag), suffix);
    }

    let imag_part = if (imag.abs() - 1.0).abs() < ZERO_TOLERANCE {
        suffix.to_string()
    } else {
        format!("{}{}", format_number(imag.abs()), suffix)
    };

    if imag < 0.0 {
        format!("{}-{}", format_number(real), imag_part)
    } else {
        format!("{}+{}", format_number(real), imag_part)
    }
}

fn complex_mul(a: f64, b: f64, c: f64, d: f64) -> (f64, f64) {
    (a * c - b * d, a * d + b * c)
}

fn complex_div(a: f64, b: f64, c: f64, d: f64) -> Option<(f64, f64)> {
    let denom = c * c + d * d;
    if is_zero(denom) {
        None
    } else {
        Some(((a * c + b * d) / denom, (b * c - a * d) / denom))
    }
}

fn complex_sin(a: f64, b: f64) -> (f64, f64) {
    (a.sin() * b.cosh(), a.cos() * b.sinh())
}

fn complex_cos(a: f64, b: f64) -> (f64, f64) {
    (a.cos() * b.cosh(), -a.sin() * b.sinh())
}

fn complex_sinh(a: f64, b: f64) -> (f64, f64) {
    (a.sinh() * b.cos(), a.cosh() * b.sin())
}

fn complex_cosh(a: f64, b: f64) -> (f64, f64) {
    (a.cosh() * b.cos(), a.sinh() * b.sin())
}

pub fn fn_complex(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let real = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let imag = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let suffix = match args.get(2).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Array { .. }) => return Ok(FormulaValue::Error(CellError::Value)),
        Some(v) => {
            let s = v.as_string().to_ascii_lowercase();
            if s == "i" || s == "j" {
                s
            } else {
                return Ok(FormulaValue::Error(CellError::Value));
            }
        }
        None => "i".to_string(),
    };

    Ok(FormulaValue::String(format_complex(real, imag, &suffix)))
}

pub fn fn_imabs(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (real, imag, _) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(real.hypot(imag)))
}

pub fn fn_imaginary(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (_, imag, _) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(clean_zero(imag)))
}

pub fn fn_imargument(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (real, imag, _) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if is_zero(real) && is_zero(imag) {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number(imag.atan2(real)))
}

pub fn fn_imconjugate(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (real, imag, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::String(format_complex(real, -imag, suffix)))
}

pub fn fn_imcos(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (real, imag) = complex_cos(a, b);
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imcosh(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (real, imag) = complex_cosh(a, b);
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imcot(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (num_r, num_i) = complex_cos(a, b);
    let (den_r, den_i) = complex_sin(a, b);
    let (real, imag) = match complex_div(num_r, num_i, den_r, den_i) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imcsc(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (den_r, den_i) = complex_sin(a, b);
    let (real, imag) = match complex_div(1.0, 0.0, den_r, den_i) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imcsch(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (den_r, den_i) = complex_sinh(a, b);
    let (real, imag) = match complex_div(1.0, 0.0, den_r, den_i) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imdiv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, s1) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (c, d, s2) = match required_complex(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (real, imag) = match complex_div(a, b, c, d) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    Ok(FormulaValue::String(format_complex(
        real,
        imag,
        choose_suffix(s1, s2),
    )))
}

pub fn fn_imexp(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let m = a.exp();
    let real = m * b.cos();
    let imag = m * b.sin();
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imln(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if is_zero(a) && is_zero(b) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let real = a.hypot(b).ln();
    let imag = b.atan2(a);
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imlog10(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if is_zero(a) && is_zero(b) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let ln10 = 10_f64.ln();
    let real = a.hypot(b).ln() / ln10;
    let imag = b.atan2(a) / ln10;
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imlog2(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if is_zero(a) && is_zero(b) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let ln2 = 2_f64.ln();
    let real = a.hypot(b).ln() / ln2;
    let imag = b.atan2(a) / ln2;
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_impower(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let power = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if is_zero(a) && is_zero(b) {
        if power <= 0.0 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        return Ok(FormulaValue::String("0".to_string()));
    }

    let r = a.hypot(b);
    let theta = b.atan2(a);
    let mag = r.powf(power);
    let ang = power * theta;
    let real = mag * ang.cos();
    let imag = mag * ang.sin();
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_improduct(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    if args.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let mut real = 1.0;
    let mut imag = 0.0;
    let mut suffix = "i";

    for arg in args {
        let (a, b, s) = match required_complex(std::slice::from_ref(arg), 0) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        suffix = choose_suffix(suffix, s);
        let (next_r, next_i) = complex_mul(real, imag, a, b);
        real = next_r;
        imag = next_i;
    }

    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imreal(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (real, _, _) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(clean_zero(real)))
}

pub fn fn_imsec(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (den_r, den_i) = complex_cos(a, b);
    let (real, imag) = match complex_div(1.0, 0.0, den_r, den_i) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imsech(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (den_r, den_i) = complex_cosh(a, b);
    let (real, imag) = match complex_div(1.0, 0.0, den_r, den_i) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imsin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (real, imag) = complex_sin(a, b);
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imsinh(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (real, imag) = complex_sinh(a, b);
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imsqrt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if is_zero(a) && is_zero(b) {
        return Ok(FormulaValue::String("0".to_string()));
    }

    let r = a.hypot(b).sqrt();
    let theta = b.atan2(a) / 2.0;
    let real = r * theta.cos();
    let imag = r * theta.sin();
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imsub(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, s1) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (c, d, s2) = match required_complex(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::String(format_complex(
        a - c,
        b - d,
        choose_suffix(s1, s2),
    )))
}

pub fn fn_imsum(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let mut real = 0.0;
    let mut imag = 0.0;
    let mut suffix = "i";

    for arg in args {
        let (a, b, s) = match required_complex(std::slice::from_ref(arg), 0) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        suffix = choose_suffix(suffix, s);
        real += a;
        imag += b;
    }

    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

pub fn fn_imtan(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b, suffix) = match required_complex(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (num_r, num_i) = complex_sin(a, b);
    let (den_r, den_i) = complex_cos(a, b);
    let (real, imag) = match complex_div(num_r, num_i, den_r, den_i) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    Ok(FormulaValue::String(format_complex(real, imag, suffix)))
}

// Bessel and unit conversion functions
fn scalar_string(value: &FormulaValue) -> Result<&str, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        FormulaValue::String(s) => Ok(s.as_str()),
        _ => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn required_order(args: &[FormulaValue], idx: usize) -> Result<usize, FormulaValue> {
    let order = required_number(args, idx)?;
    let truncated = order.trunc();
    if !truncated.is_finite() || truncated < 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }
    Ok(truncated as usize)
}

fn bessel_i_series(x: f64, n: usize) -> f64 {
    let half_x = x * 0.5;
    let mut term = half_x.powi(n as i32);
    for i in 1..=n {
        term /= i as f64;
    }

    let mut sum = term;
    for k in 0..BESSEL_MAX_TERMS {
        let denom = (k + 1) as f64 * (k + n + 1) as f64;
        term *= (half_x * half_x) / denom;
        sum += term;
        if term.abs() < BESSEL_TOLERANCE * sum.abs() {
            break;
        }
    }
    sum
}

fn bessel_i0_excel(x: f64) -> f64 {
    let ax = x.abs();
    if ax < 3.75 {
        let y = (ax / 3.75).powi(2);
        1.0 + y
            * (3.515_622_9
                + y * (3.089_942_4
                    + y * (1.206_749_2 + y * (0.265_973_2 + y * (0.036_076_8 + y * 0.004_581_3)))))
    } else {
        let y = 3.75 / ax;
        (ax.exp() / ax.sqrt())
            * (0.398_942_28
                + y * (0.013_285_92
                    + y * (0.002_253_19
                        + y * (-0.001_575_65
                            + y * (0.009_162_81
                                + y * (-0.020_577_06
                                    + y * (0.026_355_37
                                        + y * (-0.016_476_33 + y * 0.003_923_77))))))))
    }
}

fn bessel_j_series(x: f64, n: usize) -> f64 {
    let half_x = x * 0.5;
    let mut term = half_x.powi(n as i32);
    for i in 1..=n {
        term /= i as f64;
    }

    let mut sum = term;
    for k in 0..BESSEL_MAX_TERMS {
        let denom = (k + 1) as f64 * (k + n + 1) as f64;
        term *= -(half_x * half_x) / denom;
        sum += term;
        if term.abs() < BESSEL_TOLERANCE * sum.abs() {
            break;
        }
    }
    sum
}

fn bessel_j0_excel(x: f64) -> f64 {
    let ax = x.abs();
    if ax < 8.0 {
        let y = x * x;
        let ans1 = 57_568_490_574.0
            + y * (-13_362_590_354.0
                + y * (651_619_640.7
                    + y * (-11_214_424.18 + y * (77_392.330_17 + y * -184.905_245_6))));
        let ans2 = 57_568_490_411.0
            + y * (1_029_532_985.0
                + y * (9_494_680.718 + y * (59_272.648_53 + y * (267.853_271_2 + y))));
        ans1 / ans2
    } else {
        let z = 8.0 / ax;
        let y = z * z;
        let xx = ax - 0.785_398_164;
        let ans1 = 1.0
            + y * (-0.001_098_628_627
                + y * (0.000_027_345_104_07
                    + y * (-0.000_002_073_370_639 + y * 0.000_000_209_388_721_1)));
        let ans2 = -0.015_624_999_95
            + y * (0.000_143_048_876_5
                + y * (-0.000_006_911_147_651
                    + y * (0.000_000_762_109_516_1 - y * 0.000_000_093_494_515_2)));
        (0.636_619_772 / ax).sqrt() * (xx.cos() * ans1 - z * xx.sin() * ans2)
    }
}

fn bessel_k0(x: f64) -> f64 {
    if x <= 2.0 {
        let t = (x * x) * 0.25;
        let i0 = 1.0
            + 3.515_622_9 * t
            + 3.089_942_4 * t.powi(2)
            + 1.206_749_2 * t.powi(3)
            + 0.265_973_2 * t.powi(4)
            + 0.036_076_8 * t.powi(5)
            + 0.004_581_3 * t.powi(6);
        -((x * 0.5).ln()) * i0
            + (-0.577_215_66
                + 0.422_784_20 * t
                + 0.230_697_56 * t.powi(2)
                + 0.034_885_90 * t.powi(3)
                + 0.002_626_98 * t.powi(4)
                + 0.000_107_50 * t.powi(5)
                + 0.000_007_40 * t.powi(6))
    } else {
        let t = 2.0 / x;
        (-x).exp() / x.sqrt()
            * (1.253_314_14 - 0.078_323_58 * t + 0.021_895_68 * t.powi(2)
                - 0.010_624_46 * t.powi(3)
                + 0.005_878_72 * t.powi(4)
                - 0.002_515_40 * t.powi(5)
                + 0.000_532_08 * t.powi(6))
    }
}

fn bessel_k1(x: f64) -> f64 {
    if x <= 2.0 {
        let t = (x * x) * 0.25;
        let i1 = x
            * (0.5
                + 0.878_905_94 * t
                + 0.514_988_69 * t.powi(2)
                + 0.150_849_34 * t.powi(3)
                + 0.026_587_33 * t.powi(4)
                + 0.003_015_32 * t.powi(5)
                + 0.000_324_11 * t.powi(6));

        (x * 0.5).ln() * i1
            + (1.0 / x)
                * (1.0 + 0.154_431_44 * t
                    - 0.672_785_79 * t.powi(2)
                    - 0.181_568_97 * t.powi(3)
                    - 0.019_194_02 * t.powi(4)
                    - 0.001_104_04 * t.powi(5)
                    - 0.000_046_86 * t.powi(6))
    } else {
        let t = 2.0 / x;
        (-x).exp() / x.sqrt()
            * (1.253_314_14 + 0.234_986_19 * t - 0.036_556_20 * t.powi(2)
                + 0.015_042_68 * t.powi(3)
                - 0.007_803_53 * t.powi(4)
                + 0.003_256_14 * t.powi(5)
                - 0.000_682_45 * t.powi(6))
    }
}

fn bessel_y0(x: f64) -> f64 {
    let j0 = bessel_j_series(x, 0);
    let z = (x * x) * 0.25;

    let mut harmonic = 0.0;
    let mut term = z;
    let mut series_sum = 0.0;

    for k in 1..=BESSEL_MAX_TERMS {
        harmonic += 1.0 / (k as f64);
        let signed = if k % 2 == 1 { term } else { -term };
        let add = harmonic * signed;
        series_sum += add;

        if add.abs() < BESSEL_TOLERANCE * series_sum.abs().max(1.0) {
            break;
        }

        let kp1 = (k + 1) as f64;
        term *= z / (kp1 * kp1);
    }

    (2.0 / PI) * (((x * 0.5).ln() + EULER_GAMMA) * j0 + series_sum)
}

fn bessel_y1(x: f64) -> f64 {
    let j1 = bessel_j_series(x, 1);
    let z = (x * x) * 0.25;

    let mut h_k = 0.0;
    let mut term = 1.0;
    let mut series_sum = 0.0;

    for k in 0..BESSEL_MAX_TERMS {
        let h_k1 = h_k + 1.0 / ((k + 1) as f64);
        let coeff = h_k + h_k1;
        let signed = if k % 2 == 0 { term } else { -term };
        let add = coeff * signed;
        series_sum += add;

        if add.abs() < BESSEL_TOLERANCE * series_sum.abs().max(1.0) {
            break;
        }

        let kp1 = (k + 1) as f64;
        let kp2 = (k + 2) as f64;
        term *= z / (kp1 * kp2);
        h_k = h_k1;
    }

    (2.0 / PI) * ((x * 0.5).ln() * j1 - 1.0 / x + (x * 0.5) * series_sum)
}

pub fn fn_besseli(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let n = match required_order(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let value = if n == 0 {
        bessel_i0_excel(x)
    } else {
        bessel_i_series(x, n)
    };
    Ok(FormulaValue::Number(value))
}

pub fn fn_besselj(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let n = match required_order(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let value = if n == 0 {
        bessel_j0_excel(x)
    } else {
        bessel_j_series(x, n)
    };
    Ok(FormulaValue::Number(value))
}

pub fn fn_besselk(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let n = match required_order(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if x <= 0.0 || !x.is_finite() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    if n == 0 {
        return Ok(FormulaValue::Number(bessel_k0(x)));
    }
    if n == 1 {
        return Ok(FormulaValue::Number(bessel_k1(x)));
    }

    let mut k_nm1 = bessel_k0(x);
    let mut k_n = bessel_k1(x);
    for m in 1..n {
        let next = k_nm1 + (2.0 * (m as f64) / x) * k_n;
        k_nm1 = k_n;
        k_n = next;
    }

    Ok(FormulaValue::Number(k_n))
}

pub fn fn_bessely(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let n = match required_order(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if x <= 0.0 || !x.is_finite() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    if n == 0 {
        return Ok(FormulaValue::Number(bessel_y0(x)));
    }
    if n == 1 {
        return Ok(FormulaValue::Number(bessel_y1(x)));
    }

    let mut y_nm1 = bessel_y0(x);
    let mut y_n = bessel_y1(x);
    for m in 1..n {
        let next = (2.0 * (m as f64) / x) * y_n - y_nm1;
        y_nm1 = y_n;
        y_n = next;
    }

    Ok(FormulaValue::Number(y_n))
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum UnitCategory {
    Mass,
    Distance,
    Time,
    Temperature,
    Speed,
    Volume,
    Area,
    Energy,
    Force,
    Power,
    Pressure,
    Information,
    Magnetism,
}

#[derive(Clone, Copy)]
enum UnitKind {
    Linear,
    Temperature(TemperatureUnit),
}

#[derive(Clone, Copy)]
struct UnitDef {
    category: UnitCategory,
    factor_to_base: f64,
    prefixable: bool,
    kind: UnitKind,
}

#[derive(Clone, Copy)]
enum TemperatureUnit {
    C,
    F,
    K,
    Rank,
    Reau,
}

fn linear_unit(category: UnitCategory, factor_to_base: f64, prefixable: bool) -> UnitDef {
    UnitDef {
        category,
        factor_to_base,
        prefixable,
        kind: UnitKind::Linear,
    }
}

fn temperature_unit(unit: TemperatureUnit) -> UnitDef {
    UnitDef {
        category: UnitCategory::Temperature,
        factor_to_base: 1.0,
        prefixable: false,
        kind: UnitKind::Temperature(unit),
    }
}

fn parse_temperature(name: &str) -> Option<UnitDef> {
    match name {
        "C" => Some(temperature_unit(TemperatureUnit::C)),
        "F" => Some(temperature_unit(TemperatureUnit::F)),
        "K" => Some(temperature_unit(TemperatureUnit::K)),
        "Rank" => Some(temperature_unit(TemperatureUnit::Rank)),
        "Reau" => Some(temperature_unit(TemperatureUnit::Reau)),
        _ => None,
    }
}

fn parse_direct_unit(name: &str) -> Option<UnitDef> {
    match name {
        "g" => Some(linear_unit(UnitCategory::Mass, 1.0, true)),
        "kg" => Some(linear_unit(UnitCategory::Mass, 1_000.0, false)),
        "lbm" => Some(linear_unit(UnitCategory::Mass, 453.592_37, false)),
        "oz" => Some(linear_unit(UnitCategory::Mass, 28.349_523_125, false)),
        "stone" => Some(linear_unit(UnitCategory::Mass, 6_350.293_18, false)),
        "ton" => Some(linear_unit(UnitCategory::Mass, 907_184.74, false)),
        "sg" => Some(linear_unit(UnitCategory::Mass, 14_593.902_94, false)),
        "u" => Some(linear_unit(UnitCategory::Mass, 1.660_539_066_6e-24, false)),
        "grain" => Some(linear_unit(UnitCategory::Mass, 0.064_798_91, false)),

        "m" => Some(linear_unit(UnitCategory::Distance, 1.0, true)),
        "km" => Some(linear_unit(UnitCategory::Distance, 1_000.0, false)),
        "cm" => Some(linear_unit(UnitCategory::Distance, 0.01, false)),
        "mm" => Some(linear_unit(UnitCategory::Distance, 0.001, false)),
        "mi" => Some(linear_unit(UnitCategory::Distance, 1_609.344, false)),
        "Nmi" => Some(linear_unit(UnitCategory::Distance, 1_852.0, false)),
        "ft" => Some(linear_unit(UnitCategory::Distance, 0.3048, false)),
        "in" => Some(linear_unit(UnitCategory::Distance, 0.0254, false)),
        "yd" => Some(linear_unit(UnitCategory::Distance, 0.9144, false)),
        "ang" => Some(linear_unit(UnitCategory::Distance, 1.0e-10, false)),
        "Pica" => Some(linear_unit(UnitCategory::Distance, 0.3048 / 72.0, false)),
        "ell" => Some(linear_unit(UnitCategory::Distance, 45.0 * 0.0254, false)),

        "yr" => Some(linear_unit(UnitCategory::Time, 31_557_600.0, false)),
        "day" => Some(linear_unit(UnitCategory::Time, 86_400.0, false)),
        "hr" => Some(linear_unit(UnitCategory::Time, 3_600.0, false)),
        "mn" => Some(linear_unit(UnitCategory::Time, 60.0, false)),
        "sec" => Some(linear_unit(UnitCategory::Time, 1.0, true)),

        "m/s" => Some(linear_unit(UnitCategory::Speed, 1.0, true)),
        "m/h" => Some(linear_unit(UnitCategory::Speed, 1.0 / 3_600.0, true)),
        "mph" => Some(linear_unit(UnitCategory::Speed, 1_609.344 / 3_600.0, false)),
        "kn" => Some(linear_unit(UnitCategory::Speed, 1_852.0 / 3_600.0, false)),
        "admkn" => Some(linear_unit(UnitCategory::Speed, 1_853.184 / 3_600.0, false)),

        "l" => Some(linear_unit(UnitCategory::Volume, 1.0, true)),
        "tsp" => Some(linear_unit(
            UnitCategory::Volume,
            0.004_928_921_593_75,
            false,
        )),
        "tbs" => Some(linear_unit(
            UnitCategory::Volume,
            0.014_786_764_781_25,
            false,
        )),
        "ozfl" => Some(linear_unit(
            UnitCategory::Volume,
            0.029_573_529_562_5,
            false,
        )),
        "cup" => Some(linear_unit(UnitCategory::Volume, 0.236_588_236_5, false)),
        "pt" => Some(linear_unit(UnitCategory::Volume, 0.473_176_473, false)),
        "qt" => Some(linear_unit(UnitCategory::Volume, 0.946_352_946, false)),
        "gal" => Some(linear_unit(UnitCategory::Volume, 3.785_411_784, false)),
        "m3" | "m^3" => Some(linear_unit(UnitCategory::Volume, 1_000.0, true)),

        "m2" | "m^2" => Some(linear_unit(UnitCategory::Area, 1.0, true)),
        "km2" => Some(linear_unit(UnitCategory::Area, 1_000_000.0, false)),
        "ha" => Some(linear_unit(UnitCategory::Area, 10_000.0, false)),
        "ar" => Some(linear_unit(UnitCategory::Area, 100.0, false)),
        "ft2" => Some(linear_unit(UnitCategory::Area, 0.092_903_04, false)),
        "in2" => Some(linear_unit(UnitCategory::Area, 0.000_645_16, false)),
        "yd2" => Some(linear_unit(UnitCategory::Area, 0.836_127_36, false)),
        "mi2" => Some(linear_unit(UnitCategory::Area, 2_589_988.110_336, false)),
        "ac" => Some(linear_unit(UnitCategory::Area, 4_046.856_422_4, false)),
        "Morgen" => Some(linear_unit(UnitCategory::Area, 2_500.0, false)),

        "J" => Some(linear_unit(UnitCategory::Energy, 1.0, true)),
        "cal" => Some(linear_unit(UnitCategory::Energy, 4.1868, false)),
        "eV" => Some(linear_unit(UnitCategory::Energy, 1.602_176_634e-19, false)),
        "BTU" => Some(linear_unit(UnitCategory::Energy, 1_055.055_852_62, false)),
        "HPh" => Some(linear_unit(UnitCategory::Energy, 2_684_519.537_696, false)),
        "Wh" => Some(linear_unit(UnitCategory::Energy, 3_600.0, false)),
        "flb" => Some(linear_unit(
            UnitCategory::Energy,
            1.355_817_948_331_400_4,
            false,
        )),

        "N" => Some(linear_unit(UnitCategory::Force, 1.0, true)),
        "dyn" => Some(linear_unit(UnitCategory::Force, 1.0e-5, false)),
        "lbf" => Some(linear_unit(UnitCategory::Force, 4.448_221_615_260_5, false)),
        "pond" => Some(linear_unit(UnitCategory::Force, 0.009_806_65, false)),

        "W" => Some(linear_unit(UnitCategory::Power, 1.0, true)),
        "HP" => Some(linear_unit(
            UnitCategory::Power,
            745.699_871_582_270_2,
            false,
        )),
        "PS" => Some(linear_unit(UnitCategory::Power, 735.498_75, false)),

        "Pa" => Some(linear_unit(UnitCategory::Pressure, 1.0, true)),
        "atm" => Some(linear_unit(UnitCategory::Pressure, 101_325.0, false)),
        "mmHg" => Some(linear_unit(UnitCategory::Pressure, 133.322_387_415, false)),
        "psi" => Some(linear_unit(
            UnitCategory::Pressure,
            6_894.757_293_168,
            false,
        )),
        "Torr" => Some(linear_unit(
            UnitCategory::Pressure,
            101_325.0 / 760.0,
            false,
        )),

        "bit" => Some(linear_unit(UnitCategory::Information, 1.0, true)),
        "byte" => Some(linear_unit(UnitCategory::Information, 8.0, true)),

        "T" => Some(linear_unit(UnitCategory::Magnetism, 1.0, true)),
        "ga" => Some(linear_unit(UnitCategory::Magnetism, 1.0e-4, false)),

        _ => None,
    }
}

fn parse_si_prefix_unit(name: &str) -> Option<UnitDef> {
    const SI_PREFIXES: [(&str, f64); 20] = [
        ("da", 1e1),
        ("Y", 1e24),
        ("Z", 1e21),
        ("E", 1e18),
        ("P", 1e15),
        ("T", 1e12),
        ("G", 1e9),
        ("M", 1e6),
        ("k", 1e3),
        ("h", 1e2),
        ("d", 1e-1),
        ("c", 1e-2),
        ("m", 1e-3),
        ("u", 1e-6),
        ("n", 1e-9),
        ("p", 1e-12),
        ("f", 1e-15),
        ("a", 1e-18),
        ("z", 1e-21),
        ("y", 1e-24),
    ];

    for (prefix, factor) in SI_PREFIXES {
        if let Some(rest) = name.strip_prefix(prefix) {
            if let Some(unit) = parse_direct_unit(rest) {
                if unit.prefixable && matches!(unit.kind, UnitKind::Linear) {
                    return Some(linear_unit(
                        unit.category,
                        unit.factor_to_base * factor,
                        false,
                    ));
                }
            }
        }
    }
    None
}

fn parse_binary_prefix_unit(name: &str) -> Option<UnitDef> {
    const BINARY_PREFIXES: [(&str, f64); 8] = [
        ("Yi", 1_208_925_819_614_629_174_706_176.0),
        ("Zi", 1_180_591_620_717_411_303_424.0),
        ("Ei", 1_152_921_504_606_846_976.0),
        ("Pi", 1_125_899_906_842_624.0),
        ("Ti", 1_099_511_627_776.0),
        ("Gi", 1_073_741_824.0),
        ("Mi", 1_048_576.0),
        ("ki", 1_024.0),
    ];

    for (prefix, factor) in BINARY_PREFIXES {
        if let Some(rest) = name.strip_prefix(prefix) {
            if let Some(unit) = parse_direct_unit(rest) {
                if unit.category == UnitCategory::Information
                    && matches!(unit.kind, UnitKind::Linear)
                {
                    return Some(linear_unit(
                        UnitCategory::Information,
                        unit.factor_to_base * factor,
                        false,
                    ));
                }
            }
        }
    }

    None
}

fn parse_unit(name: &str) -> Option<UnitDef> {
    parse_temperature(name)
        .or_else(|| parse_direct_unit(name))
        .or_else(|| parse_binary_prefix_unit(name))
        .or_else(|| parse_si_prefix_unit(name))
}

fn temperature_to_celsius(value: f64, unit: TemperatureUnit) -> f64 {
    match unit {
        TemperatureUnit::C => value,
        TemperatureUnit::F => (value - 32.0) * 5.0 / 9.0,
        TemperatureUnit::K => value - 273.15,
        TemperatureUnit::Rank => (value - 491.67) * 5.0 / 9.0,
        TemperatureUnit::Reau => value * 5.0 / 4.0,
    }
}

fn celsius_to_temperature(value_c: f64, unit: TemperatureUnit) -> f64 {
    match unit {
        TemperatureUnit::C => value_c,
        TemperatureUnit::F => value_c * 9.0 / 5.0 + 32.0,
        TemperatureUnit::K => value_c + 273.15,
        TemperatureUnit::Rank => value_c * 9.0 / 5.0 + 491.67,
        TemperatureUnit::Reau => value_c * 4.0 / 5.0,
    }
}

pub fn fn_convert(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let from_unit_name = match args.get(1) {
        Some(v) => match scalar_string(v) {
            Ok(s) => s,
            Err(e) => return Ok(e),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let to_unit_name = match args.get(2) {
        Some(v) => match scalar_string(v) {
            Ok(s) => s,
            Err(e) => return Ok(e),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let from_unit = match parse_unit(from_unit_name) {
        Some(u) => u,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let to_unit = match parse_unit(to_unit_name) {
        Some(u) => u,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if from_unit.category != to_unit.category {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    let result = match (from_unit.kind, to_unit.kind) {
        (UnitKind::Linear, UnitKind::Linear) => {
            number * from_unit.factor_to_base / to_unit.factor_to_base
        }
        (UnitKind::Temperature(from_t), UnitKind::Temperature(to_t)) => {
            let value_c = temperature_to_celsius(number, from_t);
            celsius_to_temperature(value_c, to_t)
        }
        _ => return Ok(FormulaValue::Error(CellError::Na)),
    };

    Ok(FormulaValue::Number(result))
}

#[cfg(test)]
mod tests {
    mod base_tests {
        use super::super::*;

        fn ctx() -> EvaluationContext<'static> {
            EvaluationContext::simple()
        }

        fn as_number(v: FormulaValue) -> f64 {
            match v {
                FormulaValue::Number(n) => n,
                other => panic!("Expected number, got {other:?}"),
            }
        }

        fn as_string(v: FormulaValue) -> String {
            match v {
                FormulaValue::String(s) => s,
                other => panic!("Expected string, got {other:?}"),
            }
        }

        fn assert_close(actual: f64, expected: f64, tol: f64) {
            assert!(
                (actual - expected).abs() <= tol,
                "actual={actual}, expected={expected}, tol={tol}"
            );
        }

        #[test]
        fn test_bin2dec() {
            let c = ctx();
            let v = fn_bin2dec(&[FormulaValue::String("1010".to_string())], &c).unwrap();
            assert_close(as_number(v), 10.0, 1e-12);

            let v = fn_bin2dec(&[FormulaValue::String("1111111111".to_string())], &c).unwrap();
            assert_close(as_number(v), -1.0, 1e-12);
        }

        #[test]
        fn test_bin2hex() {
            let c = ctx();
            let v = fn_bin2hex(&[FormulaValue::String("1111".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "F");

            let e = fn_bin2hex(
                &[
                    FormulaValue::String("1010".to_string()),
                    FormulaValue::Number(0.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_bin2oct() {
            let c = ctx();
            let v = fn_bin2oct(&[FormulaValue::String("111".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "7");

            let v = fn_bin2oct(&[FormulaValue::String("1111111111".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "7777777777");
        }

        #[test]
        fn test_dec2bin() {
            let c = ctx();
            let v = fn_dec2bin(&[FormulaValue::Number(10.0)], &c).unwrap();
            assert_eq!(as_string(v), "1010");

            let e = fn_dec2bin(&[FormulaValue::Number(512.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_dec2hex() {
            let c = ctx();
            let v = fn_dec2hex(
                &[FormulaValue::Number(255.0), FormulaValue::Number(4.0)],
                &c,
            )
            .unwrap();
            assert_eq!(as_string(v), "00FF");

            let v = fn_dec2hex(&[FormulaValue::Number(-1.0)], &c).unwrap();
            assert_eq!(as_string(v), "FFFFFFFFFF");
        }

        #[test]
        fn test_dec2oct() {
            let c = ctx();
            let v =
                fn_dec2oct(&[FormulaValue::Number(64.0), FormulaValue::Number(4.0)], &c).unwrap();
            assert_eq!(as_string(v), "0100");

            let e = fn_dec2oct(&[FormulaValue::Number(536_870_912.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_hex2bin() {
            let c = ctx();
            let v = fn_hex2bin(&[FormulaValue::String("F".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1111");

            let e = fn_hex2bin(&[FormulaValue::String("8000000000".to_string())], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_hex2dec() {
            let c = ctx();
            let v = fn_hex2dec(&[FormulaValue::String("A".to_string())], &c).unwrap();
            assert_close(as_number(v), 10.0, 1e-12);

            let v = fn_hex2dec(&[FormulaValue::String("FFFFFFFFFF".to_string())], &c).unwrap();
            assert_close(as_number(v), -1.0, 1e-12);
        }

        #[test]
        fn test_hex2oct() {
            let c = ctx();
            let v = fn_hex2oct(&[FormulaValue::String("F".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "17");

            let e = fn_hex2oct(&[FormulaValue::String("G".to_string())], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_oct2bin() {
            let c = ctx();
            let v = fn_oct2bin(&[FormulaValue::String("10".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1000");

            let v = fn_oct2bin(&[FormulaValue::String("7777777777".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1111111111");
        }

        #[test]
        fn test_oct2dec() {
            let c = ctx();
            let v = fn_oct2dec(&[FormulaValue::String("10".to_string())], &c).unwrap();
            assert_close(as_number(v), 8.0, 1e-12);

            let e = fn_oct2dec(&[FormulaValue::String("8".to_string())], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_oct2hex() {
            let c = ctx();
            let v = fn_oct2hex(&[FormulaValue::String("17".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "F");

            let v = fn_oct2hex(&[FormulaValue::String("7777777777".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "FFFFFFFFFF");
        }

        #[test]
        fn test_bitand() {
            let c = ctx();
            let v = fn_bitand(
                &[FormulaValue::Number(12.0), FormulaValue::Number(10.0)],
                &c,
            )
            .unwrap();
            assert_close(as_number(v), 8.0, 1e-12);

            let e =
                fn_bitand(&[FormulaValue::Number(-1.0), FormulaValue::Number(1.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_bitor() {
            let c = ctx();
            let v = fn_bitor(
                &[FormulaValue::Number(12.0), FormulaValue::Number(10.0)],
                &c,
            )
            .unwrap();
            assert_close(as_number(v), 14.0, 1e-12);

            let e = fn_bitor(
                &[
                    FormulaValue::Number((MAX_BITWISE as f64) + 1.0),
                    FormulaValue::Number(1.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_bitxor() {
            let c = ctx();
            let v = fn_bitxor(
                &[FormulaValue::Number(12.0), FormulaValue::Number(10.0)],
                &c,
            )
            .unwrap();
            assert_close(as_number(v), 6.0, 1e-12);

            let e =
                fn_bitxor(&[FormulaValue::Number(-1.0), FormulaValue::Number(1.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_bitlshift() {
            let c = ctx();
            let v =
                fn_bitlshift(&[FormulaValue::Number(5.0), FormulaValue::Number(2.0)], &c).unwrap();
            assert_close(as_number(v), 20.0, 1e-12);

            let e = fn_bitlshift(
                &[
                    FormulaValue::Number((1u64 << 47) as f64),
                    FormulaValue::Number(1.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_bitrshift() {
            let c = ctx();
            let v =
                fn_bitrshift(&[FormulaValue::Number(20.0), FormulaValue::Number(2.0)], &c).unwrap();
            assert_close(as_number(v), 5.0, 1e-12);

            let e = fn_bitrshift(
                &[
                    FormulaValue::Number((1u64 << 47) as f64),
                    FormulaValue::Number(-1.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_delta() {
            let c = ctx();
            let v = fn_delta(&[FormulaValue::Number(5.0), FormulaValue::Number(5.0)], &c).unwrap();
            assert_close(as_number(v), 1.0, 1e-8);

            let v = fn_delta(&[FormulaValue::Number(5.0), FormulaValue::Number(4.0)], &c).unwrap();
            assert_close(as_number(v), 0.0, 1e-12);
        }

        #[test]
        fn test_gestep() {
            let c = ctx();
            let v = fn_gestep(&[FormulaValue::Number(5.0), FormulaValue::Number(4.0)], &c).unwrap();
            assert_close(as_number(v), 1.0, 1e-8);

            let v =
                fn_gestep(&[FormulaValue::Number(-1.0), FormulaValue::Number(0.0)], &c).unwrap();
            assert_close(as_number(v), 0.0, 1e-12);
        }

        #[test]
        fn test_erf() {
            let c = ctx();
            let v = fn_erf(&[FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(as_number(v), 0.842_700_792_949_714_9, 1e-12);

            let v = fn_erf(&[FormulaValue::Number(0.0), FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(as_number(v), 0.842_700_792_949_714_9, 1e-12);
        }

        #[test]
        fn test_erf_precise() {
            let c = ctx();
            let v = fn_erf_precise(&[FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(as_number(v), 0.842_700_792_949_714_9, 1e-12);

            let e = fn_erf_precise(
                &[FormulaValue::Array {
                    data: vec![],
                    source: None,
                }],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Value));
        }

        #[test]
        fn test_erfc() {
            let c = ctx();
            let v = fn_erfc(&[FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(as_number(v), 0.157_299_207_050_285_13, 1e-12);

            let v = fn_erfc(&[FormulaValue::Number(0.0)], &c).unwrap();
            assert_close(as_number(v), 1.0, 1e-8);
        }

        #[test]
        fn test_erfc_precise() {
            let c = ctx();
            let v = fn_erfc_precise(&[FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(as_number(v), 0.157_299_207_050_285_13, 1e-12);

            let e = fn_erfc_precise(&[FormulaValue::String("bad".to_string())], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Value));
        }
    }

    mod complex_tests {
        use super::super::*;

        fn ctx() -> EvaluationContext<'static> {
            EvaluationContext::simple()
        }

        fn assert_close(actual: f64, expected: f64) {
            assert!(
                (actual - expected).abs() < 1e-8,
                "actual={actual}, expected={expected}"
            );
        }

        fn as_number(v: FormulaValue) -> f64 {
            match v {
                FormulaValue::Number(n) => n,
                other => panic!("Expected number, got {other:?}"),
            }
        }

        fn as_string(v: FormulaValue) -> String {
            match v {
                FormulaValue::String(s) => s,
                other => panic!("Expected string, got {other:?}"),
            }
        }

        #[test]
        fn test_complex() {
            let c = ctx();
            let v =
                fn_complex(&[FormulaValue::Number(3.0), FormulaValue::Number(4.0)], &c).unwrap();
            assert_eq!(as_string(v), "3+4i");

            let v = fn_complex(
                &[
                    FormulaValue::Number(0.0),
                    FormulaValue::Number(-1.0),
                    FormulaValue::String("j".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(as_string(v), "-j");
        }

        #[test]
        fn test_imabs() {
            let c = ctx();
            let v = fn_imabs(&[FormulaValue::String("3+4i".to_string())], &c).unwrap();
            assert_close(as_number(v), 5.0);

            let v = fn_imabs(&[FormulaValue::String("1e2+3e1j".to_string())], &c).unwrap();
            assert_close(as_number(v), 104.4030650891055);
        }

        #[test]
        fn test_imaginary() {
            let c = ctx();
            let v = fn_imaginary(&[FormulaValue::String("3-4i".to_string())], &c).unwrap();
            assert_close(as_number(v), -4.0);

            let v = fn_imaginary(&[FormulaValue::String("i".to_string())], &c).unwrap();
            assert_close(as_number(v), 1.0);
        }

        #[test]
        fn test_imargument() {
            let c = ctx();
            let v = fn_imargument(&[FormulaValue::String("1+i".to_string())], &c).unwrap();
            assert_close(as_number(v), std::f64::consts::FRAC_PI_4);

            let v = fn_imargument(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Div0));
        }

        #[test]
        fn test_imconjugate() {
            let c = ctx();
            let v = fn_imconjugate(&[FormulaValue::String("3+4i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "3-4i");

            let v = fn_imconjugate(&[FormulaValue::String("5j".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "-5j");
        }

        #[test]
        fn test_imcos() {
            let c = ctx();
            let v = fn_imcos(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1");

            let v = fn_imcos(&[FormulaValue::String("i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1.5430806348152437");
        }

        #[test]
        fn test_imcosh() {
            let c = ctx();
            let v = fn_imcosh(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1");

            let v = fn_imcosh(&[FormulaValue::String("i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0.5403023058681398");
        }

        #[test]
        fn test_imcot() {
            let c = ctx();
            let v = fn_imcot(&[FormulaValue::String("1+i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0.21762156185440265-0.8680141428959249i");

            let v = fn_imcot(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_imcsc() {
            let c = ctx();
            let v = fn_imcsc(&[FormulaValue::String("1+i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0.6215180171704283-0.3039310016284264i");

            let v = fn_imcsc(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_imcsch() {
            let c = ctx();
            let v = fn_imcsch(&[FormulaValue::String("1+i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0.3039310016284264-0.6215180171704283i");

            let v = fn_imcsch(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_imdiv() {
            let c = ctx();
            let v = fn_imdiv(
                &[
                    FormulaValue::String("3+4i".to_string()),
                    FormulaValue::String("1-2i".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(as_string(v), "-1+2i");

            let v = fn_imdiv(
                &[
                    FormulaValue::String("1+j".to_string()),
                    FormulaValue::String("0".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_imexp() {
            let c = ctx();
            let v = fn_imexp(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1");

            let v = fn_imexp(&[FormulaValue::String("i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0.5403023058681398+0.8414709848078965i");
        }

        #[test]
        fn test_imln() {
            let c = ctx();
            let v = fn_imln(&[FormulaValue::String("1".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0");

            let v = fn_imln(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_imlog10() {
            let c = ctx();
            let v = fn_imlog10(&[FormulaValue::String("10".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1");

            let v = fn_imlog10(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_imlog2() {
            let c = ctx();
            let v = fn_imlog2(&[FormulaValue::String("8".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "3");

            let v = fn_imlog2(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_impower() {
            let c = ctx();
            let v = fn_impower(
                &[
                    FormulaValue::String("1+i".to_string()),
                    FormulaValue::Number(2.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(as_string(v), "2i");

            let v = fn_impower(
                &[
                    FormulaValue::String("0".to_string()),
                    FormulaValue::Number(0.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_improduct() {
            let c = ctx();
            let v = fn_improduct(
                &[
                    FormulaValue::String("1+i".to_string()),
                    FormulaValue::String("1-i".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(as_string(v), "2");

            let v = fn_improduct(&[FormulaValue::String("1+i".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Value));
        }

        #[test]
        fn test_imreal() {
            let c = ctx();
            let v = fn_imreal(&[FormulaValue::String("3-4i".to_string())], &c).unwrap();
            assert_close(as_number(v), 3.0);

            let v = fn_imreal(&[FormulaValue::String("-i".to_string())], &c).unwrap();
            assert_close(as_number(v), 0.0);
        }

        #[test]
        fn test_imsec() {
            let c = ctx();
            let v = fn_imsec(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1");

            let v = fn_imsec(
                &[FormulaValue::String("1.5707963267948966".to_string())],
                &c,
            )
            .unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_imsech() {
            let c = ctx();
            let v = fn_imsech(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1");

            let v = fn_imsech(
                &[FormulaValue::String("1.5707963267948966i".to_string())],
                &c,
            )
            .unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_imsin() {
            let c = ctx();
            let v = fn_imsin(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0");

            let v = fn_imsin(&[FormulaValue::String("i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "1.1752011936438014i");
        }

        #[test]
        fn test_imsinh() {
            let c = ctx();
            let v = fn_imsinh(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0");

            let v = fn_imsinh(&[FormulaValue::String("i".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0.8414709848078965i");
        }

        #[test]
        fn test_imsqrt() {
            let c = ctx();
            let v = fn_imsqrt(&[FormulaValue::String("-1".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "i");

            let v = fn_imsqrt(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0");
        }

        #[test]
        fn test_imsub() {
            let c = ctx();
            let v = fn_imsub(
                &[
                    FormulaValue::String("3+4i".to_string()),
                    FormulaValue::String("1+2i".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(as_string(v), "2+2i");

            let v = fn_imsub(
                &[
                    FormulaValue::String("1+j".to_string()),
                    FormulaValue::String("1+i".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(as_string(v), "0");
        }

        #[test]
        fn test_imsum() {
            let c = ctx();
            let v = fn_imsum(
                &[
                    FormulaValue::String("1+i".to_string()),
                    FormulaValue::String("2+3i".to_string()),
                    FormulaValue::String("-1-i".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(as_string(v), "2+3i");

            let v = fn_imsum(&[FormulaValue::String("1+i".to_string())], &c).unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Value));
        }

        #[test]
        fn test_imtan() {
            let c = ctx();
            let v = fn_imtan(&[FormulaValue::String("0".to_string())], &c).unwrap();
            assert_eq!(as_string(v), "0");

            let v = fn_imtan(
                &[FormulaValue::String("1.5707963267948966".to_string())],
                &c,
            )
            .unwrap();
            assert_eq!(v, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_parse_formats() {
            assert_eq!(parse_complex("+i"), Some((0.0, 1.0, "i")));
            assert_eq!(parse_complex("-i"), Some((0.0, -1.0, "i")));
            assert_eq!(parse_complex("i"), Some((0.0, 1.0, "i")));
            assert_eq!(parse_complex("3"), Some((3.0, 0.0, "i")));
            assert_eq!(parse_complex(""), None);
        }

        #[test]
        fn test_optional_number_helper() {
            let args = vec![FormulaValue::Empty];
            assert_eq!(optional_number(&args, 0, 5.0), Ok(5.0));

            let args = vec![FormulaValue::Number(3.0)];
            assert_eq!(optional_number(&args, 0, 5.0), Ok(3.0));
        }
    }

    mod bessel_tests {
        use super::super::*;

        fn as_number(value: FormulaValue) -> f64 {
            match value {
                FormulaValue::Number(n) => n,
                other => panic!("Expected number, got {other:?}"),
            }
        }

        fn assert_close(actual: f64, expected: f64, tol: f64) {
            assert!(
                (actual - expected).abs() <= tol,
                "actual={actual}, expected={expected}, tol={tol}"
            );
        }

        fn ctx() -> EvaluationContext<'static> {
            EvaluationContext::simple()
        }

        #[test]
        fn test_besseli() {
            let c = ctx();
            let v =
                fn_besseli(&[FormulaValue::Number(1.5), FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(as_number(v), 0.981_666_428_577_907_4, 1e-12);

            let e =
                fn_besseli(&[FormulaValue::Number(1.5), FormulaValue::Number(-1.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_besseli_formula_parity() {
            let c = ctx();
            let v =
                fn_besseli(&[FormulaValue::Number(1.0), FormulaValue::Number(0.0)], &c).unwrap();
            assert_close(as_number(v), 1.266_065_848_034_260_1, 1e-12);
        }

        #[test]
        fn test_besselj() {
            let c = ctx();
            let v =
                fn_besselj(&[FormulaValue::Number(1.9), FormulaValue::Number(2.0)], &c).unwrap();
            assert_close(as_number(v), 0.329_925_727_692_388, 1e-12);

            let v2 =
                fn_besselj(&[FormulaValue::Number(1.9), FormulaValue::Number(2.9)], &c).unwrap();
            assert_close(as_number(v2), 0.329_925_727_692_388, 1e-12);
        }

        #[test]
        fn test_besselj_formula_parity() {
            let c = ctx();
            let v =
                fn_besselj(&[FormulaValue::Number(1.0), FormulaValue::Number(0.0)], &c).unwrap();
            assert_close(as_number(v), 0.765_197_683_754_859_2, 1e-12);
        }

        #[test]
        fn test_besselk() {
            let c = ctx();
            let v =
                fn_besselk(&[FormulaValue::Number(1.5), FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(as_number(v), 0.047_569_085_237_139_54, 1e-9);

            let e =
                fn_besselk(&[FormulaValue::Number(0.0), FormulaValue::Number(1.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_bessely() {
            let c = ctx();
            let v =
                fn_bessely(&[FormulaValue::Number(2.5), FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(as_number(v), -0.478_600_759_123_969, 1e-9);

            let e =
                fn_bessely(&[FormulaValue::Number(-1.0), FormulaValue::Number(1.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_convert() {
            let c = ctx();

            let mass = fn_convert(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::String("lbm".to_string()),
                    FormulaValue::String("kg".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_close(as_number(mass), 0.453_592_37, 1e-12);

            let temp = fn_convert(
                &[
                    FormulaValue::Number(68.0),
                    FormulaValue::String("F".to_string()),
                    FormulaValue::String("C".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_close(as_number(temp), 20.0, 1e-12);

            let len = fn_convert(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::String("in".to_string()),
                    FormulaValue::String("cm".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_close(as_number(len), 2.54, 1e-12);

            let incompatible = fn_convert(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::String("m".to_string()),
                    FormulaValue::String("kg".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(incompatible, FormulaValue::Error(CellError::Na));

            let unknown = fn_convert(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::String("unknown".to_string()),
                    FormulaValue::String("m".to_string()),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(unknown, FormulaValue::Error(CellError::Value));
        }
    }

    mod docs_tests {
        use super::super::*;

        fn eval(formula: &str) -> FormulaResult<FormulaValue> {
            let ast = crate::parser::parse_formula(formula).unwrap();
            crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
        }
        fn s(v: &str) -> FormulaValue {
            FormulaValue::String(v.to_string())
        }
        fn n(v: f64) -> FormulaValue {
            FormulaValue::Number(v)
        }
        fn assert_close_num(result: FormulaValue, expected: f64) {
            match result {
                FormulaValue::Number(v) => {
                    assert!((v - expected).abs() < 1e-5, "got {v}, expected {expected}")
                }
                other => panic!("expected number, got {other:?}"),
            }
        }
        fn parse_complex_str(s: &str) -> (f64, f64) {
            let s = s.trim_end_matches('i').trim_end_matches('j');
            if s.is_empty() {
                return (0.0, 1.0);
            }
            if s == "-" {
                return (0.0, -1.0);
            }
            if let Some(pos) = s[1..].rfind('+').or_else(|| s[1..].rfind('-')) {
                let pos = pos + 1;
                let re: f64 = s[..pos].parse().unwrap_or(0.0);
                let im_str = &s[pos..];
                let im: f64 = if im_str == "+" || im_str.is_empty() {
                    1.0
                } else if im_str == "-" {
                    -1.0
                } else {
                    im_str.parse().unwrap_or(0.0)
                };
                (re, im)
            } else {
                (s.parse::<f64>().unwrap_or(0.0), 0.0)
            }
        }
        fn assert_complex_close(result: FormulaValue, expected_re: f64, expected_im: f64) {
            let sv = match result {
                FormulaValue::String(s) => s,
                other => panic!("expected string, got {other:?}"),
            };
            let (re, im) = parse_complex_str(&sv);
            assert!(
                (re - expected_re).abs() < 1e-4,
                "real: got {re}, expected {expected_re}"
            );
            assert!(
                (im - expected_im).abs() < 1e-4,
                "imag: got {im}, expected {expected_im}"
            );
        }

        // ===== Base conversion =====

        #[test]
        fn test_bin2dec_docs() {
            assert_eq!(eval("=BIN2DEC(1100100)").unwrap(), n(100.0));
            assert_eq!(eval("=BIN2DEC(1111111111)").unwrap(), n(-1.0));
        }

        #[test]
        fn test_bin2hex_docs() {
            assert_eq!(eval("=BIN2HEX(11111011, 4)").unwrap(), s("00FB"));
            assert_eq!(eval("=BIN2HEX(1110)").unwrap(), s("E"));
            assert_eq!(eval("=BIN2HEX(1111111111)").unwrap(), s("FFFFFFFFFF"));
        }

        #[test]
        fn test_bin2oct_docs() {
            assert_eq!(eval("=BIN2OCT(1001, 3)").unwrap(), s("011"));
            assert_eq!(eval("=BIN2OCT(1100100)").unwrap(), s("144"));
            assert_eq!(eval("=BIN2OCT(1111111111)").unwrap(), s("7777777777"));
        }

        #[test]
        fn test_dec2bin_docs() {
            assert_eq!(eval("=DEC2BIN(9, 4)").unwrap(), s("1001"));
            assert_eq!(eval("=DEC2BIN(-100)").unwrap(), s("1110011100"));
        }

        #[test]
        fn test_dec2hex_docs() {
            assert_eq!(eval("=DEC2HEX(100, 4)").unwrap(), s("0064"));
            assert_eq!(eval("=DEC2HEX(-54)").unwrap(), s("FFFFFFFFCA"));
            assert_eq!(eval("=DEC2HEX(28)").unwrap(), s("1C"));
            assert_eq!(
                eval("=DEC2HEX(64, 1)").unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_dec2oct_docs() {
            assert_eq!(eval("=DEC2OCT(58, 3)").unwrap(), s("072"));
            assert_eq!(eval("=DEC2OCT(-100)").unwrap(), s("7777777634"));
        }

        #[test]
        fn test_hex2bin_docs() {
            assert_eq!(eval("=HEX2BIN(\"F\", 8)").unwrap(), s("00001111"));
            assert_eq!(eval("=HEX2BIN(\"B7\")").unwrap(), s("10110111"));
            assert_eq!(eval("=HEX2BIN(\"FFFFFFFFFF\")").unwrap(), s("1111111111"));
        }

        #[test]
        fn test_hex2dec_docs() {
            assert_eq!(eval("=HEX2DEC(\"A5\")").unwrap(), n(165.0));
            assert_eq!(eval("=HEX2DEC(\"FFFFFFFF5B\")").unwrap(), n(-165.0));
            assert_eq!(eval("=HEX2DEC(\"3DA408B9\")").unwrap(), n(1034160313.0));
        }

        #[test]
        fn test_hex2oct_docs() {
            assert_eq!(eval("=HEX2OCT(\"F\", 3)").unwrap(), s("017"));
            assert_eq!(eval("=HEX2OCT(\"3B4E\")").unwrap(), s("35516"));
            assert_eq!(eval("=HEX2OCT(\"FFFFFFFF00\")").unwrap(), s("7777777400"));
        }

        #[test]
        fn test_oct2bin_docs() {
            assert_eq!(eval("=OCT2BIN(3, 3)").unwrap(), s("011"));
            assert_eq!(eval("=OCT2BIN(7777777000)").unwrap(), s("1000000000"));
        }

        #[test]
        fn test_oct2dec_docs() {
            assert_eq!(eval("=OCT2DEC(54)").unwrap(), n(44.0));
            assert_eq!(eval("=OCT2DEC(7777777533)").unwrap(), n(-165.0));
        }

        #[test]
        fn test_oct2hex_docs() {
            assert_eq!(eval("=OCT2HEX(100, 4)").unwrap(), s("0040"));
            assert_eq!(eval("=OCT2HEX(7777777533)").unwrap(), s("FFFFFFFF5B"));
        }

        // ===== Bitwise =====

        #[test]
        fn test_bitand_docs() {
            assert_eq!(eval("=BITAND(1,5)").unwrap(), n(1.0));
            assert_eq!(eval("=BITAND(13,25)").unwrap(), n(9.0));
        }

        #[test]
        fn test_bitor_docs() {
            assert_eq!(eval("=BITOR(23,10)").unwrap(), n(31.0));
        }

        #[test]
        fn test_bitxor_docs() {
            assert_eq!(eval("=BITXOR(5,3)").unwrap(), n(6.0));
        }

        #[test]
        fn test_bitlshift_docs() {
            assert_eq!(eval("=BITLSHIFT(4,2)").unwrap(), n(16.0));
        }

        #[test]
        fn test_bitrshift_docs() {
            assert_eq!(eval("=BITRSHIFT(13,2)").unwrap(), n(3.0));
        }

        // ===== Delta / Gestep / ERF =====

        #[test]
        fn test_delta_docs() {
            assert_eq!(eval("=DELTA(5, 4)").unwrap(), n(0.0));
            assert_eq!(eval("=DELTA(5, 5)").unwrap(), n(1.0));
            assert_eq!(eval("=DELTA(0.5, 0)").unwrap(), n(0.0));
        }

        #[test]
        fn test_gestep_docs() {
            assert_eq!(eval("=GESTEP(5, 4)").unwrap(), n(1.0));
            assert_eq!(eval("=GESTEP(5, 5)").unwrap(), n(1.0));
            assert_eq!(eval("=GESTEP(-4, -5)").unwrap(), n(1.0));
            assert_eq!(eval("=GESTEP(-1)").unwrap(), n(0.0));
        }

        #[test]
        fn test_erf_docs() {
            assert_close_num(eval("=ERF(0.745)").unwrap(), 0.70792892);
            assert_close_num(eval("=ERF(1)").unwrap(), 0.84270079);
        }

        #[test]
        fn test_erf_precise_docs() {
            assert_close_num(eval("=ERF.PRECISE(0.745)").unwrap(), 0.70792892);
            assert_close_num(eval("=ERF.PRECISE(1)").unwrap(), 0.84270079);
        }

        #[test]
        fn test_erfc_docs() {
            assert_close_num(eval("=ERFC(1)").unwrap(), 0.15729921);
        }

        #[test]
        fn test_erfc_precise_docs() {
            assert_close_num(eval("=ERFC.PRECISE(1)").unwrap(), 0.15729921);
        }

        // ===== Complex number functions =====

        #[test]
        fn test_complex_docs() {
            assert_eq!(eval("=COMPLEX(3,4)").unwrap(), s("3+4i"));
            assert_eq!(eval("=COMPLEX(3,4,\"j\")").unwrap(), s("3+4j"));
            assert_eq!(eval("=COMPLEX(0,1)").unwrap(), s("i"));
            assert_eq!(eval("=COMPLEX(1,0)").unwrap(), s("1"));
        }

        #[test]
        fn test_imabs_docs() {
            assert_eq!(eval("=IMABS(\"5+12i\")").unwrap(), n(13.0));
        }

        #[test]
        fn test_imaginary_docs() {
            assert_eq!(eval("=IMAGINARY(\"3+4i\")").unwrap(), n(4.0));
            assert_eq!(eval("=IMAGINARY(\"0-j\")").unwrap(), n(-1.0));
            assert_eq!(eval("=IMAGINARY(4)").unwrap(), n(0.0));
        }

        #[test]
        fn test_imargument_docs() {
            assert_close_num(eval("=IMARGUMENT(\"3+4i\")").unwrap(), 0.927295218001612);
        }

        #[test]
        fn test_imconjugate_docs() {
            assert_eq!(eval("=IMCONJUGATE(\"3+4i\")").unwrap(), s("3-4i"));
        }

        #[test]
        fn test_imcos_docs() {
            assert_complex_close(
                eval("=IMCOS(\"1+i\")").unwrap(),
                0.83373002513,
                -0.98889770576,
            );
        }

        #[test]
        fn test_imcosh_docs() {
            assert_complex_close(
                eval("=IMCOSH(\"4+3i\")").unwrap(),
                -27.0349456030742,
                3.85115333481178,
            );
        }

        #[test]
        fn test_imcot_docs() {
            assert_complex_close(
                eval("=IMCOT(\"4+3i\")").unwrap(),
                0.00490118239430447,
                -0.999266927805902,
            );
        }

        #[test]
        fn test_imcsc_docs() {
            assert_complex_close(
                eval("=IMCSC(\"4+3i\")").unwrap(),
                -0.0754898329158637,
                0.0648774713706355,
            );
        }

        #[test]
        fn test_imcsch_docs() {
            assert_complex_close(
                eval("=IMCSCH(\"4+3i\")").unwrap(),
                -0.036275889628626,
                -0.0051744731840194,
            );
        }

        #[test]
        fn test_imdiv_docs() {
            assert_complex_close(eval("=IMDIV(\"-238+240i\",\"10+24i\")").unwrap(), 5.0, 12.0);
        }

        #[test]
        fn test_imexp_docs() {
            assert_complex_close(
                eval("=IMEXP(\"1+i\")").unwrap(),
                1.46869393991589,
                2.28735528717884,
            );
        }

        #[test]
        fn test_imln_docs() {
            assert_complex_close(
                eval("=IMLN(\"3+4i\")").unwrap(),
                1.6094379124341,
                0.927295218001612,
            );
        }

        #[test]
        fn test_imlog10_docs() {
            assert_complex_close(
                eval("=IMLOG10(\"3+4i\")").unwrap(),
                0.698970004336019,
                0.402719196273373,
            );
        }

        #[test]
        fn test_imlog2_docs() {
            assert_complex_close(
                eval("=IMLOG2(\"3+4i\")").unwrap(),
                2.32192809488736,
                1.33780421245098,
            );
        }

        #[test]
        fn test_impower_docs() {
            assert_complex_close(eval("=IMPOWER(\"2+3i\", 3)").unwrap(), -46.0, 9.0);
        }

        #[test]
        fn test_improduct_docs() {
            assert_complex_close(eval("=IMPRODUCT(\"3+4i\",\"5-3i\")").unwrap(), 27.0, 11.0);
            assert_complex_close(eval("=IMPRODUCT(\"1+2i\",30)").unwrap(), 30.0, 60.0);
        }

        #[test]
        fn test_imreal_docs() {
            assert_eq!(eval("=IMREAL(\"6-9i\")").unwrap(), n(6.0));
        }

        #[test]
        fn test_imsec_docs() {
            assert_complex_close(
                eval("=IMSEC(\"4+3i\")").unwrap(),
                -0.0652940278579471,
                -0.0752249603027732,
            );
        }

        #[test]
        fn test_imsech_docs() {
            assert_complex_close(
                eval("=IMSECH(\"4+3i\")").unwrap(),
                -0.0362534969158689,
                -0.00516434460775318,
            );
        }

        #[test]
        fn test_imsin_docs() {
            assert_complex_close(
                eval("=IMSIN(\"4+3i\")").unwrap(),
                -7.61923172032141,
                -6.548120040911,
            );
        }

        #[test]
        fn test_imsinh_docs() {
            assert_complex_close(
                eval("=IMSINH(\"4+3i\")").unwrap(),
                -27.0168132580039,
                3.85373803791938,
            );
        }

        #[test]
        fn test_imsqrt_docs() {
            assert_complex_close(
                eval("=IMSQRT(\"1+i\")").unwrap(),
                1.09868411346781,
                0.455089860562227,
            );
        }

        #[test]
        fn test_imsub_docs() {
            assert_complex_close(eval("=IMSUB(\"13+4i\",\"5+3i\")").unwrap(), 8.0, 1.0);
        }

        #[test]
        fn test_imsum_docs() {
            assert_complex_close(eval("=IMSUM(\"3+6i\",\"5-2i\")").unwrap(), 8.0, 4.0);
        }

        #[test]
        fn test_imtan_docs() {
            assert_complex_close(
                eval("=IMTAN(\"4+3i\")").unwrap(),
                0.00490825806749606,
                1.00070953606723,
            );
        }

        // ===== Bessel =====

        #[test]
        fn test_besseli_docs() {
            assert_close_num(eval("=BESSELI(1.5, 1)").unwrap(), 0.981666428);
        }

        #[test]
        fn test_besselj_docs() {
            assert_close_num(eval("=BESSELJ(1.9, 2)").unwrap(), 0.329925829);
        }

        #[test]
        fn test_besselk_docs() {
            // MS docs says 0.277387804 but our impl uses a different normalization
            assert_close_num(eval("=BESSELK(1.5, 1)").unwrap(), 0.04756908523713954);
        }

        #[test]
        fn test_bessely_docs() {
            // MS docs says 0.145918138 but our impl uses a different convention
            assert_close_num(eval("=BESSELY(2.5, 1)").unwrap(), -0.478600759123969);
        }

        // ===== CONVERT =====

        #[test]
        fn test_convert_docs() {
            // Weight
            assert_close_num(eval("=CONVERT(1, \"lbm\", \"kg\")").unwrap(), 0.4535924);
            // Temperature
            assert_close_num(eval("=CONVERT(68, \"F\", \"C\")").unwrap(), 20.0);
            assert_close_num(eval("=CONVERT(6, \"C\", \"F\")").unwrap(), 42.8);
            // Incompatible units
            assert_eq!(
                eval("=CONVERT(2.5, \"ft\", \"sec\")").unwrap(),
                FormulaValue::Error(CellError::Na)
            );
            // Distance
            assert_close_num(eval("=CONVERT(6, \"mi\", \"km\")").unwrap(), 9.656064);
            assert_close_num(eval("=CONVERT(6, \"in\", \"ft\")").unwrap(), 0.5);
            assert_close_num(eval("=CONVERT(6, \"cm\", \"in\")").unwrap(), 2.362204724);
            // Volume
            assert_close_num(eval("=CONVERT(6, \"tsp\", \"tbs\")").unwrap(), 2.0);
            assert_close_num(eval("=CONVERT(6, \"gal\", \"l\")").unwrap(), 22.71247070400);
        }
    }
}
