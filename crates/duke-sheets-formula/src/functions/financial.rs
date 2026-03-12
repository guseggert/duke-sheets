use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

const MAX_ITERATIONS: usize = 100;
const TOLERANCE: f64 = 1e-10;

fn scalar_number(value: &FormulaValue) -> Result<f64, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
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
        FormulaValue::Array(rows) => {
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
        FormulaValue::Array(rows) => {
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
                FormulaValue::Array(vec![]),
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
                FormulaValue::Array(vec![vec![
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                ]]),
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
                FormulaValue::Array(vec![vec![
                    FormulaValue::Number(-70000.0),
                    FormulaValue::Number(12000.0),
                    FormulaValue::Number(15000.0),
                    FormulaValue::Number(18000.0),
                    FormulaValue::Number(21000.0),
                    FormulaValue::Number(26000.0),
                ]]),
                FormulaValue::Number(0.1),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.08663094803653158, 1e-8);

        let e = fn_irr(
            &[FormulaValue::Array(vec![vec![
                FormulaValue::Number(100.0),
                FormulaValue::Number(200.0),
            ]])],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));

        let e = fn_irr(
            &[
                FormulaValue::Array(vec![vec![FormulaValue::Number(-100.0)]]),
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
                FormulaValue::Array(vec![vec![
                    FormulaValue::Number(-120000.0),
                    FormulaValue::Number(39000.0),
                    FormulaValue::Number(30000.0),
                    FormulaValue::Number(21000.0),
                    FormulaValue::Number(37000.0),
                    FormulaValue::Number(46000.0),
                ]]),
                FormulaValue::Number(0.1),
                FormulaValue::Number(0.12),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 0.1260941303665752, 1e-8);

        let e = fn_mirr(
            &[
                FormulaValue::Array(vec![vec![
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                ]]),
                FormulaValue::Number(0.1),
                FormulaValue::Number(0.1),
            ],
            &c,
        )
        .unwrap();
        assert_eq!(e, FormulaValue::Error(CellError::Num));

        let e = fn_mirr(
            &[
                FormulaValue::Array(vec![vec![
                    FormulaValue::Number(-1.0),
                    FormulaValue::Number(2.0),
                ]]),
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
                FormulaValue::Array(vec![vec![
                    FormulaValue::Number(-10000.0),
                    FormulaValue::Number(2750.0),
                    FormulaValue::Number(4250.0),
                    FormulaValue::Number(3250.0),
                    FormulaValue::Number(2750.0),
                ]]),
                FormulaValue::Array(vec![vec![
                    FormulaValue::Number(39448.0),
                    FormulaValue::Number(39508.0),
                    FormulaValue::Number(39751.0),
                    FormulaValue::Number(39859.0),
                    FormulaValue::Number(39904.0),
                ]]),
            ],
            &c,
        )
        .unwrap();
        assert_close(as_number(v), 2086.6476020315354, 1e-6);

        let e = fn_xnpv(
            &[
                FormulaValue::Number(0.1),
                FormulaValue::Array(vec![vec![
                    FormulaValue::Number(-100.0),
                    FormulaValue::Number(50.0),
                ]]),
                FormulaValue::Array(vec![vec![FormulaValue::Number(0.0)]]),
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
}
