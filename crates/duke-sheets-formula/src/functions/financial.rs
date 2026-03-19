use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

const MAX_ITERATIONS: usize = 100;
const TOLERANCE: f64 = 1e-10;

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

fn payment_type(v: f64) -> Option<f64> {
    let t = v.trunc();
    if (t - 0.0).abs() < f64::EPSILON {
        Some(0.0)
    } else if (t - 1.0).abs() < f64::EPSILON {
        Some(1.0)
    } else {
        None
    }
}

fn calc_pmt(rate: f64, nper: f64, pv: f64, fv: f64, when: f64) -> Result<f64, CellError> {
    if nper <= 0.0 {
        return Err(CellError::Num);
    }

    if rate.abs() < f64::EPSILON {
        return Ok(-(pv + fv) / nper);
    }

    let r1 = (1.0 + rate).powf(nper);
    let denom = (1.0 + rate * when) * (r1 - 1.0);
    if denom.abs() < f64::EPSILON {
        return Err(CellError::Num);
    }

    Ok(-rate * (pv * r1 + fv) / denom)
}

fn calc_fv(rate: f64, nper: f64, pmt: f64, pv: f64, when: f64) -> Result<f64, CellError> {
    if rate.abs() < f64::EPSILON {
        return Ok(-(pv + pmt * nper));
    }

    let r1 = (1.0 + rate).powf(nper);
    Ok(-pv * r1 - pmt * (r1 - 1.0) * (1.0 + rate * when) / rate)
}

fn calc_pv(rate: f64, nper: f64, pmt: f64, fv: f64, when: f64) -> Result<f64, CellError> {
    if rate.abs() < f64::EPSILON {
        return Ok(-(fv + pmt * nper));
    }

    let r1 = (1.0 + rate).powf(nper);
    if r1.abs() < f64::EPSILON {
        return Err(CellError::Num);
    }

    Ok((-fv - pmt * (r1 - 1.0) * (1.0 + rate * when) / rate) / r1)
}

fn calc_ipmt(
    rate: f64,
    per: f64,
    nper: f64,
    pv: f64,
    fv: f64,
    when: f64,
) -> Result<f64, CellError> {
    if nper <= 0.0 || per < 1.0 || per > nper {
        return Err(CellError::Num);
    }
    if rate.abs() < f64::EPSILON {
        return Ok(0.0);
    }

    let pmt = calc_pmt(rate, nper, pv, fv, when)?;

    if when == 1.0 && per <= 1.0 {
        return Ok(0.0);
    }

    let mut ipmt = calc_fv(rate, per - 1.0, pmt, pv, when)? * rate;
    if when == 1.0 {
        ipmt /= 1.0 + rate;
    }
    Ok(ipmt)
}

fn collect_numbers(value: &FormulaValue, out: &mut Vec<f64>) -> Result<(), FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { data: rows, .. } => {
            for row in rows {
                for cell in row {
                    collect_numbers(cell, out)?;
                }
            }
            Ok(())
        }
        _ => {
            let n = value
                .as_number()
                .ok_or(FormulaValue::Error(CellError::Value))?;
            out.push(n);
            Ok(())
        }
    }
}

fn collect_dates(value: &FormulaValue, out: &mut Vec<f64>) -> Result<(), FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { data: rows, .. } => {
            for row in rows {
                for cell in row {
                    collect_dates(cell, out)?;
                }
            }
            Ok(())
        }
        _ => {
            let n = value
                .as_number()
                .ok_or(FormulaValue::Error(CellError::Value))?;
            out.push(n.floor());
            Ok(())
        }
    }
}

fn rate_function(rate: f64, nper: f64, pmt: f64, pv: f64, fv: f64, when: f64) -> f64 {
    if rate.abs() < 1e-12 {
        pv + pmt * nper + fv
    } else {
        let r1 = (1.0 + rate).powf(nper);
        pv * r1 + pmt * (1.0 + rate * when) * (r1 - 1.0) / rate + fv
    }
}

pub fn fn_pmt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nper = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fv = match optional_number(args, 3, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when_raw = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    match calc_pmt(rate, nper, pv, fv, when) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_fv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nper = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pmt = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match optional_number(args, 3, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when_raw = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    match calc_fv(rate, nper, pmt, pv, when) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_pv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nper = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pmt = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fv = match optional_number(args, 3, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when_raw = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    match calc_pv(rate, nper, pmt, fv, when) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_nper(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pmt = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fv = match optional_number(args, 3, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when_raw = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    if rate.abs() < f64::EPSILON {
        if pmt.abs() < f64::EPSILON {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        return Ok(FormulaValue::Number(-(pv + fv) / pmt));
    }

    let a = pmt * (1.0 + rate * when) / rate;
    let num = a - fv;
    let den = a + pv;
    let base = 1.0 + rate;

    if den.abs() < f64::EPSILON || base <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let ratio = num / den;
    if ratio <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let nper = ratio.ln() / base.ln();
    if !nper.is_finite() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(nper))
}

pub fn fn_rate(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let nper = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pmt = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fv = match optional_number(args, 3, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when_raw = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    let guess = match optional_number(args, 5, 0.1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if nper <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut rate = if guess <= -0.999_999_999 { 0.1 } else { guess };

    for _ in 0..MAX_ITERATIONS {
        let f = rate_function(rate, nper, pmt, pv, fv, when);
        if f.abs() < TOLERANCE {
            return Ok(FormulaValue::Number(rate));
        }

        let h = 1e-7;
        let left_r = (rate - h).max(-0.999_999_999);
        let right_r = rate + h;
        let left = rate_function(left_r, nper, pmt, pv, fv, when);
        let right = rate_function(right_r, nper, pmt, pv, fv, when);
        let df = (right - left) / (right_r - left_r);

        if !df.is_finite() || df.abs() < 1e-14 {
            return Ok(FormulaValue::Error(CellError::Num));
        }

        let next = rate - f / df;
        if !next.is_finite() || next <= -1.0 {
            return Ok(FormulaValue::Error(CellError::Num));
        }

        if (next - rate).abs() < TOLERANCE {
            return Ok(FormulaValue::Number(next));
        }
        rate = next;
    }

    Ok(FormulaValue::Error(CellError::Num))
}

pub fn fn_ipmt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let per = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nper = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fv = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when_raw = match optional_number(args, 5, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    match calc_ipmt(rate, per.trunc(), nper.trunc(), pv, fv, when) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}

pub fn fn_ppmt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let per = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nper = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fv = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when_raw = match optional_number(args, 5, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    let nper_i = nper.trunc();
    let per_i = per.trunc();
    let pmt = match calc_pmt(rate, nper_i, pv, fv, when) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let ipmt = match calc_ipmt(rate, per_i, nper_i, pv, fv, when) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    Ok(FormulaValue::Number(pmt - ipmt))
}

pub fn fn_cumipmt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nper = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let start_period = match required_number(args, 3) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };
    let end_period = match required_number(args, 4) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };
    let when_raw = match required_number(args, 5) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    let nper_i = nper.trunc();
    if rate <= 0.0
        || nper_i <= 0.0
        || pv <= 0.0
        || start_period < 1.0
        || end_period < start_period
        || end_period > nper_i
    {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut total = 0.0;
    let mut p = start_period;
    while p <= end_period {
        let ip = match calc_ipmt(rate, p, nper_i, pv, 0.0, when) {
            Ok(v) => v,
            Err(e) => return Ok(FormulaValue::Error(e)),
        };
        total += ip;
        p += 1.0;
    }

    Ok(FormulaValue::Number(total))
}

pub fn fn_cumprinc(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nper = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let start_period = match required_number(args, 3) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };
    let end_period = match required_number(args, 4) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };
    let when_raw = match required_number(args, 5) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let when = match payment_type(when_raw) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    let nper_i = nper.trunc();
    if rate <= 0.0
        || nper_i <= 0.0
        || pv <= 0.0
        || start_period < 1.0
        || end_period < start_period
        || end_period > nper_i
    {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let pmt = match calc_pmt(rate, nper_i, pv, 0.0, when) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    let mut total = 0.0;
    let mut p = start_period;
    while p <= end_period {
        let ip = match calc_ipmt(rate, p, nper_i, pv, 0.0, when) {
            Ok(v) => v,
            Err(e) => return Ok(FormulaValue::Error(e)),
        };
        total += pmt - ip;
        p += 1.0;
    }

    Ok(FormulaValue::Number(total))
}

pub fn fn_npv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if rate <= -1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    if args.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let mut values = Vec::new();
    for arg in &args[1..] {
        if let Err(e) = collect_numbers(arg, &mut values) {
            return Ok(e);
        }
    }

    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut npv = 0.0;
    for (i, v) in values.iter().enumerate() {
        npv += v / (1.0 + rate).powf((i + 1) as f64);
    }

    Ok(FormulaValue::Number(npv))
}

pub fn fn_irr(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut values = Vec::new();
    if let Some(first) = args.first() {
        if let Err(e) = collect_numbers(first, &mut values) {
            return Ok(e);
        }
    } else {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let guess = match optional_number(args, 1, 0.1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if values.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let has_pos = values.iter().any(|v| *v > 0.0);
    let has_neg = values.iter().any(|v| *v < 0.0);
    if !has_pos || !has_neg {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut rate = if guess <= -0.999_999_999 { 0.1 } else { guess };

    for _ in 0..MAX_ITERATIONS {
        let mut f = 0.0;
        let mut df = 0.0;
        let base = 1.0 + rate;
        if base <= 0.0 {
            return Ok(FormulaValue::Error(CellError::Num));
        }

        for (i, v) in values.iter().enumerate() {
            let i_f = i as f64;
            f += v / base.powf(i_f);
            if i > 0 {
                df -= i_f * v / base.powf(i_f + 1.0);
            }
        }

        if f.abs() < TOLERANCE {
            return Ok(FormulaValue::Number(rate));
        }
        if df.abs() < 1e-14 || !df.is_finite() {
            return Ok(FormulaValue::Error(CellError::Num));
        }

        let next = rate - f / df;
        if !next.is_finite() || next <= -1.0 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        if (next - rate).abs() < TOLERANCE {
            return Ok(FormulaValue::Number(next));
        }
        rate = next;
    }

    Ok(FormulaValue::Error(CellError::Num))
}

pub fn fn_mirr(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut values = Vec::new();
    if let Some(first) = args.first() {
        if let Err(e) = collect_numbers(first, &mut values) {
            return Ok(e);
        }
    } else {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let finance_rate = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let reinvest_rate = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if values.len() < 2 || finance_rate <= -1.0 || reinvest_rate <= -1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let has_pos = values.iter().any(|v| *v > 0.0);
    let has_neg = values.iter().any(|v| *v < 0.0);
    if !has_pos || !has_neg {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let n = values.len() as f64;
    let mut pv_neg = 0.0;
    let mut fv_pos = 0.0;
    for (i, v) in values.iter().enumerate() {
        let i_f = i as f64;
        if *v < 0.0 {
            pv_neg += v / (1.0 + finance_rate).powf(i_f);
        } else if *v > 0.0 {
            fv_pos += v * (1.0 + reinvest_rate).powf(n - i_f - 1.0);
        }
    }

    if pv_neg.abs() < f64::EPSILON {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mirr = (-fv_pos / pv_neg).powf(1.0 / (n - 1.0)) - 1.0;
    if !mirr.is_finite() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(mirr))
}

pub fn fn_xnpv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if rate <= -1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut values = Vec::new();
    let mut dates = Vec::new();
    let Some(v) = args.get(1) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let Some(d) = args.get(2) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };

    if let Err(e) = collect_numbers(v, &mut values) {
        return Ok(e);
    }
    if let Err(e) = collect_dates(d, &mut dates) {
        return Ok(e);
    }

    if values.is_empty() || values.len() != dates.len() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let first_date = dates[0];
    let mut xnpv = 0.0;
    for (v, d) in values.iter().zip(dates.iter()) {
        xnpv += v / (1.0 + rate).powf((d - first_date) / 365.0);
    }

    Ok(FormulaValue::Number(xnpv))
}

pub fn fn_sln(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let cost = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let salvage = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let life = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if life <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number((cost - salvage) / life))
}

pub fn fn_syd(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let cost = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let salvage = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let life = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let per = match required_number(args, 3) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };

    if life <= 0.0 || per < 1.0 || per > life {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let dep = (cost - salvage) * (life - per + 1.0) * 2.0 / (life * (life + 1.0));
    Ok(FormulaValue::Number(dep))
}

pub fn fn_db(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let cost = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let salvage = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let life = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let period = match required_number(args, 3) {
        Ok(v) => v.trunc() as i32,
        Err(e) => return Ok(e),
    };
    let month = match optional_number(args, 4, 12.0) {
        Ok(v) => v.trunc() as i32,
        Err(e) => return Ok(e),
    };

    if cost <= 0.0
        || salvage < 0.0
        || life <= 0.0
        || period <= 0
        || !(1..=12).contains(&month)
        || salvage > cost
    {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut rate = 1.0 - (salvage / cost).powf(1.0 / life);
    rate = (rate * 1000.0).round() / 1000.0;

    let life_i = life.trunc() as i32;
    let mut accumulated = 0.0;
    let mut dep = cost * rate * (month as f64) / 12.0;
    if dep > cost - salvage {
        dep = cost - salvage;
    }

    if period == 1 {
        return Ok(FormulaValue::Number(dep.max(0.0)));
    }

    accumulated += dep;
    for p in 2..=period {
        let book = (cost - accumulated).max(salvage);
        dep = if p == life_i + 1 {
            book * rate * ((12 - month) as f64) / 12.0
        } else {
            book * rate
        };
        if dep > book - salvage {
            dep = (book - salvage).max(0.0);
        }
        accumulated += dep;
    }

    Ok(FormulaValue::Number(dep.max(0.0)))
}

pub fn fn_ddb(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let cost = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let salvage = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let life = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let period = match required_number(args, 3) {
        Ok(v) => v.trunc() as i32,
        Err(e) => return Ok(e),
    };
    let factor = match optional_number(args, 4, 2.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if cost <= 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0 || factor <= 0.0 || salvage > cost
    {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut book = cost;
    let mut dep = 0.0;

    for _ in 1..=period {
        dep = book * factor / life;
        if book - dep < salvage {
            dep = book - salvage;
        }
        if dep < 0.0 {
            dep = 0.0;
        }
        book -= dep;
        if book <= salvage {
            break;
        }
    }

    Ok(FormulaValue::Number(dep))
}

pub fn fn_effect(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let nominal_rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let npery = match required_number(args, 1) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };

    if nominal_rate <= 0.0 || npery < 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(
        (1.0 + nominal_rate / npery).powf(npery) - 1.0,
    ))
}

pub fn fn_nominal(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let effect_rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let npery = match required_number(args, 1) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };

    if effect_rate <= 0.0 || npery < 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(
        npery * ((1.0 + effect_rate).powf(1.0 / npery) - 1.0),
    ))
}

pub fn fn_pduration(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let rate = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pv = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fv = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if rate <= 0.0 || pv <= 0.0 || fv <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let periods = (fv.ln() - pv.ln()) / (1.0 + rate).ln();
    if !periods.is_finite() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(periods))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_close(actual: f64, expected: f64, tol: f64) {
        assert!(
            (actual - expected).abs() <= tol,
            "actual={actual}, expected={expected}, tol={tol}"
        );
    }

    fn as_number(v: FormulaValue) -> f64 {
        match v {
            FormulaValue::Number(n) => n,
            other => panic!("Expected number, got {other:?}"),
        }
    }

    fn ctx() -> EvaluationContext<'static> {
        EvaluationContext::simple()
    }

    #[test]
    fn test_pmt() {
        let c = ctx();
        let v = fn_pmt(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), -1073.6432460242795, 1e-6);

        let v = fn_pmt(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Number(10.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Empty,
                FormulaValue::Empty,
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), -100.0, 1e-9);

        let e = fn_pmt(
            &[
                FormulaValue::Number(0.01),
                FormulaValue::Number(0.0),
                FormulaValue::Number(1000.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_fv() {
        let c = ctx();
        let v = fn_fv(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(-1000.0),
                FormulaValue::Number(0.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 832258.6351716131, 1.0);

        let v = fn_fv(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Number(12.0),
                FormulaValue::Number(-100.0),
                FormulaValue::Number(-1000.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 2200.0, 1e-9);

        let e = fn_fv(
            &[
                FormulaValue::String("bad".to_string()),
                FormulaValue::Number(1.0),
                FormulaValue::Number(1.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Value));
    }

    #[test]
    fn test_pv() {
        let c = ctx();
        let v = fn_pv(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(-1073.6432460242795),
                FormulaValue::Number(0.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 200000.0, 1e-3);

        let v = fn_pv(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Number(10.0),
                FormulaValue::Number(-100.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 1000.0, 1e-9);

        let e = fn_pv(
            &[
                FormulaValue::Array { data: vec![], source: None },
                FormulaValue::Number(1.0),
                FormulaValue::Number(1.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Value));
    }

    #[test]
    fn test_nper() {
        let c = ctx();
        let v = fn_nper(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(-1073.6432460242795),
                FormulaValue::Number(200000.0),
                FormulaValue::Number(0.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 360.0, 1e-5);

        let v = fn_nper(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Number(-100.0),
                FormulaValue::Number(1000.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 10.0, 1e-9);

        let e = fn_nper(
            &[
                FormulaValue::Number(0.01),
                FormulaValue::Number(0.0),
                FormulaValue::Number(1000.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_rate() {
        let c = ctx();
        let v = fn_rate(
            &[
                FormulaValue::Number(360.0),
                FormulaValue::Number(-1073.6432460242795),
                FormulaValue::Number(200000.0),
                FormulaValue::Number(0.0),
                FormulaValue::Number(0.0),
                FormulaValue::Number(0.01),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.05 / 12.0, 1e-8);

        let v = fn_rate(
            &[
                FormulaValue::Number(10.0),
                FormulaValue::Number(-100.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(0.0),
                FormulaValue::Number(0.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.0, 1e-8);

        let e = fn_rate(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Number(-100.0),
                FormulaValue::Number(1000.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_ipmt() {
        let c = ctx();
        let v = fn_ipmt(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), -833.3333333333334, 1e-6);

        let v = fn_ipmt(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(10.0),
                FormulaValue::Number(1000.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.0, 1e-9);

        let e = fn_ipmt(
            &[
                FormulaValue::Number(0.01),
                FormulaValue::Number(11.0),
                FormulaValue::Number(10.0),
                FormulaValue::Number(1000.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_ppmt() {
        let c = ctx();
        let v = fn_ppmt(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v.clone()), -240.30991269094602, 1e-6);

        let pmt = fn_pmt(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
            ],
            &c,
        )
        .unwrap();
        let ipmt = fn_ipmt(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), as_number(pmt) - as_number(ipmt), 1e-9);

        let e = fn_ppmt(
            &[
                FormulaValue::Number(0.01),
                FormulaValue::Number(0.0),
                FormulaValue::Number(10.0),
                FormulaValue::Number(1000.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_cumipmt() {
        let c = ctx();
        let v = fn_cumipmt(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(12.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert!(as_number(v) < 0.0);

        let one = fn_cumipmt(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(one), -833.3333333333334, 1e-6);

        let e = fn_cumipmt(
            &[
                FormulaValue::Number(-0.01),
                FormulaValue::Number(10.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(2.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_cumprinc() {
        let c = ctx();
        let v = fn_cumprinc(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(12.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert!(as_number(v) < 0.0);

        let one = fn_cumprinc(
            &[
                FormulaValue::Number(0.05 / 12.0),
                FormulaValue::Number(360.0),
                FormulaValue::Number(200000.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(one), -240.30991269094602, 1e-6);

        let e = fn_cumprinc(
            &[
                FormulaValue::Number(0.01),
                FormulaValue::Number(10.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(3.0),
                FormulaValue::Number(2.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_npv() {
        let c = ctx();
        let v = fn_npv(
            &[
                FormulaValue::Number(0.1),
                FormulaValue::Number(-10000.0),
                FormulaValue::Number(3000.0),
                FormulaValue::Number(4200.0),
                FormulaValue::Number(6800.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 1188.4434123352205, 1e-6);

        let v = fn_npv(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Array { data: vec![vec![
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                ]], source: None },
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 3.0, 1e-9);

        let e = fn_npv(&[FormulaValue::Number(-1.0), FormulaValue::Number(1.0)], &c).unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_irr() {
        let c = ctx();
        let v = fn_irr(
            &[
                FormulaValue::Array { data: vec![vec![
                    FormulaValue::Number(-70000.0),
                    FormulaValue::Number(12000.0),
                    FormulaValue::Number(15000.0),
                    FormulaValue::Number(18000.0),
                    FormulaValue::Number(21000.0),
                    FormulaValue::Number(26000.0),
                ]], source: None },
                FormulaValue::Number(0.1),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.08663094803653158, 1e-8);

        let e = fn_irr(
            &[FormulaValue::Array { data: vec![vec![
                FormulaValue::Number(100.0),
                FormulaValue::Number(200.0),
            ]], source: None }],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));

        let e = fn_irr(
            &[
                FormulaValue::Array { data: vec![vec![FormulaValue::Number(-100.0)]], source: None },
                FormulaValue::Number(0.1),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_mirr() {
        let c = ctx();
        let v = fn_mirr(
            &[
                FormulaValue::Array { data: vec![vec![
                    FormulaValue::Number(-120000.0),
                    FormulaValue::Number(39000.0),
                    FormulaValue::Number(30000.0),
                    FormulaValue::Number(21000.0),
                    FormulaValue::Number(37000.0),
                    FormulaValue::Number(46000.0),
                ]], source: None },
                FormulaValue::Number(0.1),
                FormulaValue::Number(0.12),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.1260941303665752, 1e-8);

        let e = fn_mirr(
            &[
                FormulaValue::Array { data: vec![vec![
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                ]], source: None },
                FormulaValue::Number(0.1),
                FormulaValue::Number(0.1),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));

        let e = fn_mirr(
            &[
                FormulaValue::Array { data: vec![vec![
                    FormulaValue::Number(-1.0),
                    FormulaValue::Number(2.0),
                ]], source: None },
                FormulaValue::Number(-1.0),
                FormulaValue::Number(0.1),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_xnpv() {
        let c = ctx();
        let v = fn_xnpv(
            &[
                FormulaValue::Number(0.09),
                FormulaValue::Array { data: vec![vec![
                    FormulaValue::Number(-10000.0),
                    FormulaValue::Number(2750.0),
                    FormulaValue::Number(4250.0),
                    FormulaValue::Number(3250.0),
                    FormulaValue::Number(2750.0),
                ]], source: None },
                FormulaValue::Array { data: vec![vec![
                    FormulaValue::Number(39448.0),
                    FormulaValue::Number(39508.0),
                    FormulaValue::Number(39751.0),
                    FormulaValue::Number(39859.0),
                    FormulaValue::Number(39904.0),
                ]], source: None },
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 2086.6476020315354, 1e-6);

        let e = fn_xnpv(
            &[
                FormulaValue::Number(0.1),
                FormulaValue::Array { data: vec![vec![
                    FormulaValue::Number(-100.0),
                    FormulaValue::Number(50.0),
                ]], source: None },
                FormulaValue::Array { data: vec![vec![FormulaValue::Number(0.0)]], source: None },
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));

        let e = fn_xnpv(
            &[
                FormulaValue::Number(-1.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(1.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_sln() {
        let c = ctx();
        let v = fn_sln(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(5.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 1800.0, 1e-9);

        let v = fn_sln(
            &[
                FormulaValue::Number(1000.0),
                FormulaValue::Number(0.0),
                FormulaValue::Number(1.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 1000.0, 1e-9);

        let e = fn_sln(
            &[
                FormulaValue::Number(1000.0),
                FormulaValue::Number(0.0),
                FormulaValue::Number(0.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_syd() {
        let c = ctx();
        let v = fn_syd(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(1.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 3000.0, 1e-9);

        let v = fn_syd(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(5.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 600.0, 1e-9);

        let e = fn_syd(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(6.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_db() {
        let c = ctx();
        let v = fn_db(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(12.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 3690.0, 1e-9);

        let v = fn_db(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(2.0),
            ],
            &c,
        )
        .unwrap();
        assert!(as_number(v) > 0.0);

        let e = fn_db(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(-5.0),
                FormulaValue::Number(1.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_ddb() {
        let c = ctx();
        let v = fn_ddb(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(1.0),
                FormulaValue::Number(2.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 4000.0, 1e-9);

        let v = fn_ddb(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(5.0),
            ],
            &c,
        )
        .unwrap();
        assert!(as_number(v) >= 0.0);

        let e = fn_ddb(
            &[
                FormulaValue::Number(10000.0),
                FormulaValue::Number(11000.0),
                FormulaValue::Number(5.0),
                FormulaValue::Number(1.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_effect() {
        let c = ctx();
        let v = fn_effect(
            &[FormulaValue::Number(0.12), FormulaValue::Number(12.0)],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.12682503013196977, 1e-12);

        let v = fn_effect(&[FormulaValue::Number(0.1), FormulaValue::Number(1.0)], &c).unwrap();
        assert_close(as_number(v), 0.1, 1e-12);

        let e = fn_effect(&[FormulaValue::Number(0.0), FormulaValue::Number(12.0)], &c).unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_nominal() {
        let c = ctx();
        let v = fn_nominal(
            &[
                FormulaValue::Number(0.12682503013196977),
                FormulaValue::Number(12.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.12, 1e-10);

        let v = fn_nominal(&[FormulaValue::Number(0.1), FormulaValue::Number(1.0)], &c).unwrap();
        assert_close(as_number(v), 0.1, 1e-12);

        let e = fn_nominal(
            &[FormulaValue::Number(-0.1), FormulaValue::Number(12.0)],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    #[test]
    fn test_pduration() {
        let c = ctx();
        let v = fn_pduration(
            &[
                FormulaValue::Number(0.08),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(2000.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 9.006468342000597, 1e-10);

        let v = fn_pduration(
            &[
                FormulaValue::Number(0.1),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(1100.0),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 1.0, 1e-12);

        let e = fn_pduration(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Number(1000.0),
                FormulaValue::Number(2000.0),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));
    }

    fn n(x: f64) -> FormulaValue {
        FormulaValue::Number(x)
    }

    fn arr(values: &[f64]) -> FormulaValue {
        FormulaValue::Array { data: vec![values
            .iter()
            .map(|v| FormulaValue::Number(*v))
            .collect()], source: None }
    }

    // PMT docs: =PMT(0.08/12, 10, 10000) = -1037.03
    #[test]
    fn test_pmt_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_pmt(&[n(0.08 / 12.0), n(10.0), n(10000.0)], &c).unwrap()),
            -1037.03,
            1e-2,
        );
    }

    // PMT docs: type=1 (beginning of period)
    #[test]
    fn test_pmt_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_pmt(&[n(0.08 / 12.0), n(10.0), n(10000.0), n(0.0), n(1.0)], &c).unwrap()),
            -1030.16,
            1e-2,
        );
    }

    // PMT docs: saving for future value
    #[test]
    fn test_pmt_docs_3() {
        let c = ctx();
        assert_close(
            as_number(fn_pmt(&[n(0.06 / 12.0), n(216.0), n(0.0), n(50000.0)], &c).unwrap()),
            -129.08,
            1e-2,
        );
    }

    // FV docs: =FV(0.06/12, 10, -200, -500, 1) = 2581.40
    #[test]
    fn test_fv_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_fv(&[n(0.06 / 12.0), n(10.0), n(-200.0), n(-500.0), n(1.0)], &c).unwrap()),
            2581.40,
            1e-2,
        );
    }

    // FV docs: =FV(0.12/12, 12, -1000) = 12682.50
    #[test]
    fn test_fv_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_fv(&[n(0.12 / 12.0), n(12.0), n(-1000.0)], &c).unwrap()),
            12682.50,
            1e-2,
        );
    }

    // FV docs: type=1
    #[test]
    fn test_fv_docs_3() {
        let c = ctx();
        assert_close(
            as_number(fn_fv(&[n(0.11 / 12.0), n(35.0), n(-2000.0), n(0.0), n(1.0)], &c).unwrap()),
            82846.25,
            1e-2,
        );
    }

    // FV docs: with pv and type
    #[test]
    fn test_fv_docs_4() {
        let c = ctx();
        assert_close(
            as_number(
                fn_fv(
                    &[n(0.06 / 12.0), n(12.0), n(-100.0), n(-1000.0), n(1.0)],
                    &c,
                )
                .unwrap(),
            ),
            2301.40,
            1e-2,
        );
    }

    // PV docs: =PV(0.08/12, 240, 500, 0, 0) = -59777.15
    #[test]
    fn test_pv_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_pv(&[n(0.08 / 12.0), n(240.0), n(500.0), n(0.0), n(0.0)], &c).unwrap()),
            -59777.15,
            1e-2,
        );
    }

    // PV docs: without optional args
    #[test]
    fn test_pv_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_pv(&[n(0.08 / 12.0), n(240.0), n(500.0)], &c).unwrap()),
            -59777.15,
            1e-2,
        );
    }

    // NPER docs: =NPER(0.12/12, -100, -1000, 10000, 1) = 59.6738657
    #[test]
    fn test_nper_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_nper(
                    &[n(0.12 / 12.0), n(-100.0), n(-1000.0), n(10000.0), n(1.0)],
                    &c,
                )
                .unwrap(),
            ),
            59.6738657,
            1e-4,
        );
    }

    // NPER docs: type omitted
    #[test]
    fn test_nper_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_nper(&[n(0.12 / 12.0), n(-100.0), n(-1000.0), n(10000.0)], &c).unwrap()),
            60.0821229,
            1e-4,
        );
    }

    // NPER docs: fv omitted
    #[test]
    fn test_nper_docs_3() {
        let c = ctx();
        assert_close(
            as_number(fn_nper(&[n(0.12 / 12.0), n(-100.0), n(-1000.0)], &c).unwrap()),
            -9.57859404,
            1e-4,
        );
    }

    // RATE docs: =RATE(48, -200, 8000) ≈ 0.0077 per month
    #[test]
    fn test_rate_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_rate(&[n(48.0), n(-200.0), n(8000.0)], &c).unwrap()),
            0.0077,
            1e-4,
        );
    }

    // IPMT docs: =IPMT(0.10/12, 1, 36, 8000) = -66.67
    #[test]
    fn test_ipmt_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_ipmt(&[n(0.10 / 12.0), n(1.0), n(36.0), n(8000.0)], &c).unwrap()),
            -66.67,
            1e-2,
        );
    }

    // IPMT docs: annual rate, last year
    #[test]
    fn test_ipmt_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_ipmt(&[n(0.10), n(3.0), n(3.0), n(8000.0)], &c).unwrap()),
            -292.45,
            1e-2,
        );
    }

    // PPMT docs: =PPMT(0.10/12, 1, 24, 2000) = -75.62
    #[test]
    fn test_ppmt_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_ppmt(&[n(0.10 / 12.0), n(1.0), n(24.0), n(2000.0)], &c).unwrap()),
            -75.62,
            1e-2,
        );
    }

    // PPMT docs: annual, period 10
    #[test]
    fn test_ppmt_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_ppmt(&[n(0.08), n(10.0), n(10.0), n(200000.0)], &c).unwrap()),
            -27598.05,
            1e-2,
        );
    }

    // CUMIPMT docs: periods 13-24
    #[test]
    fn test_cumipmt_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_cumipmt(
                    &[
                        n(0.09 / 12.0),
                        n(360.0),
                        n(125000.0),
                        n(13.0),
                        n(24.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            -11135.23,
            1e-2,
        );
    }

    // CUMIPMT docs: first month only
    #[test]
    fn test_cumipmt_docs_2() {
        let c = ctx();
        assert_close(
            as_number(
                fn_cumipmt(
                    &[
                        n(0.09 / 12.0),
                        n(360.0),
                        n(125000.0),
                        n(1.0),
                        n(1.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            -937.5,
            1e-2,
        );
    }

    // CUMPRINC docs: periods 13-24
    #[test]
    fn test_cumprinc_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_cumprinc(
                    &[
                        n(0.09 / 12.0),
                        n(360.0),
                        n(125000.0),
                        n(13.0),
                        n(24.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            -934.11,
            1e-2,
        );
    }

    // CUMPRINC docs: first month only
    #[test]
    fn test_cumprinc_docs_2() {
        let c = ctx();
        assert_close(
            as_number(
                fn_cumprinc(
                    &[
                        n(0.09 / 12.0),
                        n(360.0),
                        n(125000.0),
                        n(1.0),
                        n(1.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            -68.28,
            1e-2,
        );
    }

    // NPV docs: =NPV(0.1, -10000, 3000, 4200, 6800) = 1188.44
    #[test]
    fn test_npv_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_npv(&[n(0.1), n(-10000.0), n(3000.0), n(4200.0), n(6800.0)], &c).unwrap()),
            1188.44,
            1e-2,
        );
    }

    // IRR docs: =IRR({-70000,12000,15000,18000,21000}) = -2.12%
    #[test]
    fn test_irr_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_irr(&[arr(&[-70000.0, 12000.0, 15000.0, 18000.0, 21000.0])], &c).unwrap()),
            -0.0212,
            1e-3,
        );
    }

    // IRR docs: =IRR({-70000,12000,15000,18000,21000,26000}) = 8.66%
    #[test]
    fn test_irr_docs_2() {
        let c = ctx();
        assert_close(
            as_number(
                fn_irr(
                    &[arr(&[
                        -70000.0, 12000.0, 15000.0, 18000.0, 21000.0, 26000.0,
                    ])],
                    &c,
                )
                .unwrap(),
            ),
            0.0866,
            1e-3,
        );
    }

    // MIRR docs: =MIRR({-120000,39000,30000,21000,37000,46000},0.10,0.12) = 12.61%
    #[test]
    fn test_mirr_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_mirr(
                    &[
                        arr(&[-120000.0, 39000.0, 30000.0, 21000.0, 37000.0, 46000.0]),
                        n(0.10),
                        n(0.12),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            0.1261,
            1e-3,
        );
    }

    // XNPV docs: rate=0.09, values and dates
    #[test]
    fn test_xnpv_docs() {
        let c = ctx();
        // dates: 2008-01-01=39448, 2008-03-01=39508, 2008-10-30=39751, 2009-02-15=39859, 2009-04-01=39904
        let values = arr(&[-10000.0, 2750.0, 4250.0, 3250.0, 2750.0]);
        let dates = arr(&[39448.0, 39508.0, 39751.0, 39859.0, 39904.0]);
        assert_close(
            as_number(fn_xnpv(&[n(0.09), values, dates], &c).unwrap()),
            2086.65,
            1e-2,
        );
    }

    // SLN docs: =SLN(30000, 7500, 10) = 2250
    #[test]
    fn test_sln_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_sln(&[n(30000.0), n(7500.0), n(10.0)], &c).unwrap()),
            2250.0,
            1e-2,
        );
    }

    // SYD docs: period 1
    #[test]
    fn test_syd_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_syd(&[n(30000.0), n(7500.0), n(10.0), n(1.0)], &c).unwrap()),
            4090.91,
            1e-2,
        );
    }

    // SYD docs: period 10
    #[test]
    fn test_syd_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_syd(&[n(30000.0), n(7500.0), n(10.0), n(10.0)], &c).unwrap()),
            409.09,
            1e-2,
        );
    }

    // DB docs: period 1 (month=7)
    #[test]
    fn test_db_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_db(&[n(1000000.0), n(100000.0), n(6.0), n(1.0), n(7.0)], &c).unwrap()),
            186083.33,
            1e-2,
        );
    }

    // DB docs: period 2
    #[test]
    fn test_db_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_db(&[n(1000000.0), n(100000.0), n(6.0), n(2.0), n(7.0)], &c).unwrap()),
            259639.42,
            1e-2,
        );
    }

    // DB docs: period 3
    #[test]
    fn test_db_docs_3() {
        let c = ctx();
        assert_close(
            as_number(fn_db(&[n(1000000.0), n(100000.0), n(6.0), n(3.0), n(7.0)], &c).unwrap()),
            176814.44,
            1e-2,
        );
    }

    // DB docs: period 4
    #[test]
    fn test_db_docs_4() {
        let c = ctx();
        assert_close(
            as_number(fn_db(&[n(1000000.0), n(100000.0), n(6.0), n(4.0), n(7.0)], &c).unwrap()),
            120410.64,
            1e-2,
        );
    }

    // DDB docs: daily life, period 1
    #[test]
    fn test_ddb_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_ddb(&[n(2400.0), n(300.0), n(3650.0), n(1.0)], &c).unwrap()),
            1.32,
            1e-2,
        );
    }

    // DDB docs: monthly life, period 1
    #[test]
    fn test_ddb_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_ddb(&[n(2400.0), n(300.0), n(120.0), n(1.0), n(2.0)], &c).unwrap()),
            40.00,
            1e-2,
        );
    }

    // DDB docs: yearly life, period 1
    #[test]
    fn test_ddb_docs_3() {
        let c = ctx();
        assert_close(
            as_number(fn_ddb(&[n(2400.0), n(300.0), n(10.0), n(1.0), n(2.0)], &c).unwrap()),
            480.00,
            1e-2,
        );
    }

    // DDB docs: factor 1.5, period 2
    #[test]
    fn test_ddb_docs_4() {
        let c = ctx();
        assert_close(
            as_number(fn_ddb(&[n(2400.0), n(300.0), n(10.0), n(2.0), n(1.5)], &c).unwrap()),
            306.00,
            1e-2,
        );
    }

    // DDB docs: period 10
    #[test]
    fn test_ddb_docs_5() {
        let c = ctx();
        assert_close(
            as_number(fn_ddb(&[n(2400.0), n(300.0), n(10.0), n(10.0)], &c).unwrap()),
            22.12,
            1e-2,
        );
    }

    // EFFECT docs: =EFFECT(0.0525, 4) = 0.0535427
    #[test]
    fn test_effect_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_effect(&[n(0.0525), n(4.0)], &c).unwrap()),
            0.0535427,
            1e-4,
        );
    }

    // NOMINAL docs: =NOMINAL(0.053543, 4) = 0.0525003
    #[test]
    fn test_nominal_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_nominal(&[n(0.053543), n(4.0)], &c).unwrap()),
            0.0525003,
            1e-4,
        );
    }

    // PDURATION docs: =PDURATION(0.025, 2000, 2200) = 3.86
    #[test]
    fn test_pduration_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_pduration(&[n(0.025), n(2000.0), n(2200.0)], &c).unwrap()),
            3.86,
            1e-2,
        );
    }
}
