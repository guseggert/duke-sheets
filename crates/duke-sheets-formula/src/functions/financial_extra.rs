use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use chrono::{Datelike, NaiveDate};
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

fn required_serial(args: &[FormulaValue], idx: usize) -> Result<i64, FormulaValue> {
    Ok(required_number(args, idx)?.floor() as i64)
}

fn optional_basis(args: &[FormulaValue], idx: usize, default: i64) -> Result<i64, FormulaValue> {
    Ok(optional_number(args, idx, default as f64)?.trunc() as i64)
}

fn parse_frequency(v: f64) -> Option<i64> {
    let f = v.trunc() as i64;
    if matches!(f, 1 | 2 | 4) {
        Some(f)
    } else {
        None
    }
}

fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0) && ((year % 100 != 0) || (year % 400 == 0))
}

fn year_days(year: i32) -> f64 {
    if is_leap_year(year) {
        366.0
    } else {
        365.0
    }
}

fn serial_to_date(serial: i64) -> Option<NaiveDate> {
    if serial == 60 {
        return None;
    }
    let base = NaiveDate::from_ymd_opt(1899, 12, 31)?;
    let adjusted = if serial > 60 { serial - 1 } else { serial };
    base.checked_add_signed(chrono::Duration::days(adjusted))
}

fn date_to_serial(date: NaiveDate) -> i64 {
    let base = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
    let mut serial = (date - base).num_days();
    if date >= NaiveDate::from_ymd_opt(1900, 3, 1).unwrap() {
        serial += 1;
    }
    serial
}

fn last_day_of_month(year: i32, month: u32) -> Option<u32> {
    let first = NaiveDate::from_ymd_opt(year, month, 1)?;
    let (ny, nm) = if month == 12 {
        (year + 1, 1)
    } else {
        (year, month + 1)
    };
    let next = NaiveDate::from_ymd_opt(ny, nm, 1)?;
    Some((next - first).num_days() as u32)
}

fn add_months(date: NaiveDate, months: i32) -> Option<NaiveDate> {
    let total = date.year() * 12 + date.month0() as i32 + months;
    let year = total.div_euclid(12);
    let month0 = total.rem_euclid(12) as u32;
    let month = month0 + 1;
    let last = last_day_of_month(year, month)?;
    let day = date.day().min(last);
    NaiveDate::from_ymd_opt(year, month, day)
}

fn days360_us(mut sy: i32, mut sm: u32, mut sd: u32, mut ey: i32, mut em: u32, mut ed: u32) -> i32 {
    let mut sign = 1;
    if (ey, em, ed) < (sy, sm, sd) {
        sign = -1;
        (sy, sm, sd, ey, em, ed) = (ey, em, ed, sy, sm, sd);
    }

    if sd == 31 {
        sd = 30;
    }
    if ed == 31 && sd >= 30 {
        ed = 30;
    }

    sign * (360 * (ey - sy) + 30 * (em as i32 - sm as i32) + (ed as i32 - sd as i32))
}

fn days360_eu(mut sy: i32, mut sm: u32, mut sd: u32, mut ey: i32, mut em: u32, mut ed: u32) -> i32 {
    let mut sign = 1;
    if (ey, em, ed) < (sy, sm, sd) {
        sign = -1;
        (sy, sm, sd, ey, em, ed) = (ey, em, ed, sy, sm, sd);
    }

    if sd == 31 {
        sd = 30;
    }
    if ed == 31 {
        ed = 30;
    }

    sign * (360 * (ey - sy) + 30 * (em as i32 - sm as i32) + (ed as i32 - sd as i32))
}

fn day_count_days(start: i64, end: i64, basis: i64) -> Result<f64, FormulaValue> {
    let (sign, s, e) = if end >= start {
        (1.0, start, end)
    } else {
        (-1.0, end, start)
    };
    let sd = serial_to_date(s).ok_or(FormulaValue::Error(CellError::Num))?;
    let ed = serial_to_date(e).ok_or(FormulaValue::Error(CellError::Num))?;

    let days = match basis {
        0 => days360_us(
            sd.year(),
            sd.month(),
            sd.day(),
            ed.year(),
            ed.month(),
            ed.day(),
        ) as f64,
        4 => days360_eu(
            sd.year(),
            sd.month(),
            sd.day(),
            ed.year(),
            ed.month(),
            ed.day(),
        ) as f64,
        1..=3 => (e - s) as f64,
        _ => return Err(FormulaValue::Error(CellError::Num)),
    };

    Ok(sign * days)
}

fn yearfrac_basis(start: i64, end: i64, basis: i64) -> Result<f64, FormulaValue> {
    if !(0..=4).contains(&basis) {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let (sign, s, e) = if end >= start {
        (1.0, start, end)
    } else {
        (-1.0, end, start)
    };

    let sd = serial_to_date(s).ok_or(FormulaValue::Error(CellError::Num))?;
    let ed = serial_to_date(e).ok_or(FormulaValue::Error(CellError::Num))?;

    let frac = match basis {
        0 => day_count_days(s, e, 0)? / 360.0,
        4 => day_count_days(s, e, 4)? / 360.0,
        2 => (e - s) as f64 / 360.0,
        3 => (e - s) as f64 / 365.0,
        1 => {
            if sd.year() == ed.year() {
                (e - s) as f64 / year_days(sd.year())
            } else {
                let next_jan1 = NaiveDate::from_ymd_opt(sd.year() + 1, 1, 1)
                    .ok_or(FormulaValue::Error(CellError::Num))?;
                let this_jan1 = NaiveDate::from_ymd_opt(ed.year(), 1, 1)
                    .ok_or(FormulaValue::Error(CellError::Num))?;
                let mut total = (next_jan1 - sd).num_days() as f64 / year_days(sd.year());
                for y in (sd.year() + 1)..ed.year() {
                    total += 1.0;
                    if is_leap_year(y) {
                        total += 0.0;
                    }
                }
                total + (ed - this_jan1).num_days() as f64 / year_days(ed.year())
            }
        }
        _ => 0.0,
    };

    Ok(sign * frac)
}

#[derive(Clone)]
struct CouponInfo {
    pcd: NaiveDate,
    ncd: NaiveDate,
    num: i64,
    e: f64,
    a: f64,
    dsc: f64,
}

fn coupon_info(
    settlement: i64,
    maturity: i64,
    frequency: i64,
    basis: i64,
) -> Result<CouponInfo, FormulaValue> {
    if !matches!(frequency, 1 | 2 | 4) || !(0..=4).contains(&basis) {
        return Err(FormulaValue::Error(CellError::Num));
    }
    if settlement >= maturity {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let s = serial_to_date(settlement).ok_or(FormulaValue::Error(CellError::Num))?;
    let m = serial_to_date(maturity).ok_or(FormulaValue::Error(CellError::Num))?;
    let months = (12 / frequency) as i32;

    let mut ncd = m;
    let pcd = loop {
        let prev = add_months(ncd, -months).ok_or(FormulaValue::Error(CellError::Num))?;
        if s >= prev {
            break prev;
        }
        ncd = prev;
    };

    let mut num = 1i64;
    let mut d = ncd;
    while d < m {
        d = add_months(d, months).ok_or(FormulaValue::Error(CellError::Num))?;
        num += 1;
    }

    let pcd_serial = date_to_serial(pcd);
    let ncd_serial = date_to_serial(ncd);

    let e = match basis {
        0 | 2 | 4 => 360.0 / frequency as f64,
        3 => 365.0 / frequency as f64,
        1 => day_count_days(pcd_serial, ncd_serial, 1)?,
        _ => return Err(FormulaValue::Error(CellError::Num)),
    };
    let a = match basis {
        0 => day_count_days(pcd_serial, settlement, 0)?,
        4 => day_count_days(pcd_serial, settlement, 4)?,
        _ => day_count_days(pcd_serial, settlement, 1)?,
    };
    let dsc = match basis {
        0 => day_count_days(settlement, ncd_serial, 0)?,
        4 => day_count_days(settlement, ncd_serial, 4)?,
        _ => day_count_days(settlement, ncd_serial, 1)?,
    };

    Ok(CouponInfo {
        pcd,
        ncd,
        num,
        e,
        a,
        dsc,
    })
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

fn price_from_yield_core(
    settlement: i64,
    maturity: i64,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i64,
    basis: i64,
) -> Result<f64, FormulaValue> {
    if rate < 0.0 || redemption <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }
    if yld <= -(frequency as f64) {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let coup = coupon_info(settlement, maturity, frequency, basis)?;
    let per_y = yld / frequency as f64;
    let c = redemption * rate / frequency as f64;

    let mut pv = 0.0;
    for k in 1..=coup.num {
        let t = (k as f64 - 1.0) + coup.dsc / coup.e;
        let df = (1.0 + per_y).powf(t);
        if !df.is_finite() || df <= 0.0 {
            return Err(FormulaValue::Error(CellError::Num));
        }
        pv += c / df;
        if k == coup.num {
            pv += redemption / df;
        }
    }

    Ok(pv - c * coup.a / coup.e)
}

fn duration_core(
    settlement: i64,
    maturity: i64,
    coupon: f64,
    yld: f64,
    frequency: i64,
    basis: i64,
) -> Result<f64, FormulaValue> {
    if coupon < 0.0 || yld < 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }
    let coup = coupon_info(settlement, maturity, frequency, basis)?;
    let per_y = yld / frequency as f64;

    let mut pv_total = 0.0;
    let mut weighted = 0.0;
    let c = 100.0 * coupon / frequency as f64;

    for k in 1..=coup.num {
        let t_period = (k as f64 - 1.0) + coup.dsc / coup.e;
        let t_year = t_period / frequency as f64;
        let df = (1.0 + per_y).powf(t_period);
        if !df.is_finite() || df <= 0.0 {
            return Err(FormulaValue::Error(CellError::Num));
        }
        let mut cf = c;
        if k == coup.num {
            cf += 100.0;
        }
        let pv = cf / df;
        pv_total += pv;
        weighted += t_year * pv;
    }

    pv_total -= c * coup.a / coup.e;
    if pv_total <= 0.0 || !pv_total.is_finite() {
        return Err(FormulaValue::Error(CellError::Num));
    }

    Ok(weighted / pv_total)
}

pub fn fn_accrint(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let issue = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let first_interest = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let settlement = match required_serial(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let par = match optional_number(args, 4, 1000.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 5)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 6, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let calc_method = match optional_number(args, 7, 1.0) {
        Ok(v) => v != 0.0,
        Err(e) => return Ok(e),
    };

    if settlement <= issue || first_interest <= issue || rate < 0.0 || par <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let start = if calc_method || settlement <= first_interest {
        issue
    } else {
        let mut d = match serial_to_date(first_interest) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        let settlement_d = match serial_to_date(settlement) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        let months = (12 / frequency) as i32;
        while d <= settlement_d {
            d = match add_months(d, months) {
                Some(v) => v,
                None => return Ok(FormulaValue::Error(CellError::Num)),
            };
        }
        let prev = match add_months(d, -months) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        date_to_serial(prev).max(issue)
    };

    let yf = match yearfrac_basis(start, settlement, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(par * rate * yf))
}

pub fn fn_accrintm(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let issue = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let settlement = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let par = match optional_number(args, 3, 1000.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 4, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if settlement <= issue || rate < 0.0 || par <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let yf = match yearfrac_basis(issue, settlement, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(par * rate * yf))
}

fn amordegrc_coeff(rate: f64) -> f64 {
    if rate <= 0.0 {
        return 1.0;
    }
    let life = 1.0 / rate;
    if life < 3.0 {
        1.0
    } else if life <= 4.0 {
        1.5
    } else if life <= 6.0 {
        2.0
    } else {
        2.5
    }
}

fn amorlinc_like(
    cost: f64,
    date_purchased: i64,
    first_period: i64,
    salvage: f64,
    period: i64,
    rate: f64,
    basis: i64,
    degrc: bool,
) -> Result<f64, FormulaValue> {
    if cost <= 0.0
        || salvage < 0.0
        || rate <= 0.0
        || period < 0
        || first_period <= date_purchased
        || cost <= salvage
    {
        return Err(FormulaValue::Error(CellError::Num));
    }
    if !(0..=4).contains(&basis) {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let year_base = if basis == 3 { 365.0 } else { 360.0 };
    let first_days = day_count_days(date_purchased, first_period, basis)?.abs();
    let mut remaining = cost;
    let mut dep = cost * rate;
    if degrc {
        dep *= amordegrc_coeff(rate);
    }
    let mut first_dep = dep * (first_days / year_base);
    first_dep = first_dep.min(remaining - salvage).max(0.0);

    if period == 0 {
        return Ok(first_dep);
    }

    remaining -= first_dep;
    let mut p = 1i64;
    while p <= period {
        let mut this_dep = if degrc {
            remaining * rate * amordegrc_coeff(rate)
        } else {
            dep
        };
        if !degrc && this_dep > remaining - salvage {
            this_dep = remaining - salvage;
        }
        if degrc {
            let remaining_life = ((remaining - salvage) / (cost * rate).max(1e-12)).max(1.0);
            let sl = (remaining - salvage) / remaining_life;
            if sl > this_dep {
                this_dep = sl;
            }
            this_dep = this_dep.min(remaining - salvage).max(0.0);
        }
        if p == period {
            return Ok(this_dep.max(0.0));
        }
        remaining -= this_dep;
        if remaining <= salvage {
            return Ok(0.0);
        }
        p += 1;
    }

    Ok(0.0)
}

pub fn fn_amordegrc(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let cost = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let date_purchased = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let first_period = match required_serial(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let salvage = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let period = match required_number(args, 4) {
        Ok(v) => v.trunc() as i64,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 5) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 6, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match amorlinc_like(
        cost,
        date_purchased,
        first_period,
        salvage,
        period,
        rate,
        basis,
        true,
    ) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(e),
    }
}

pub fn fn_amorlinc(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let cost = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let date_purchased = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let first_period = match required_serial(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let salvage = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let period = match required_number(args, 4) {
        Ok(v) => v.trunc() as i64,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 5) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 6, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match amorlinc_like(
        cost,
        date_purchased,
        first_period,
        salvage,
        period,
        rate,
        basis,
        false,
    ) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(e),
    }
}

pub fn fn_coupdaybs(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 2)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 3, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match coupon_info(settlement, maturity, frequency, basis) {
        Ok(v) => Ok(FormulaValue::Number(v.a)),
        Err(e) => Ok(e),
    }
}

pub fn fn_coupdays(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 2)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 3, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match coupon_info(settlement, maturity, frequency, basis) {
        Ok(v) => Ok(FormulaValue::Number(v.e)),
        Err(e) => Ok(e),
    }
}

pub fn fn_coupdaysnc(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 2)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 3, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match coupon_info(settlement, maturity, frequency, basis) {
        Ok(v) => Ok(FormulaValue::Number(v.dsc)),
        Err(e) => Ok(e),
    }
}

pub fn fn_coupncd(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 2)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 3, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match coupon_info(settlement, maturity, frequency, basis) {
        Ok(v) => Ok(FormulaValue::Number(date_to_serial(v.ncd) as f64)),
        Err(e) => Ok(e),
    }
}

pub fn fn_coupnum(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 2)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 3, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match coupon_info(settlement, maturity, frequency, basis) {
        Ok(v) => Ok(FormulaValue::Number(v.num as f64)),
        Err(e) => Ok(e),
    }
}

pub fn fn_couppcd(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 2)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 3, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match coupon_info(settlement, maturity, frequency, basis) {
        Ok(v) => Ok(FormulaValue::Number(date_to_serial(v.pcd) as f64)),
        Err(e) => Ok(e),
    }
}

pub fn fn_disc(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pr = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let redemption = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 4, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if settlement >= maturity || pr <= 0.0 || redemption <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let yf = match yearfrac_basis(settlement, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if yf <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number((redemption - pr) / redemption / yf))
}

pub fn fn_dollarde(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let fractional_dollar = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fraction = match required_number(args, 1) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };

    if fraction < 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let sign = if fractional_dollar < 0.0 { -1.0 } else { 1.0 };
    let abs = fractional_dollar.abs();
    let int_part = abs.trunc();
    let frac_part = abs - int_part;
    let numerator = (frac_part * 100.0).round();
    Ok(FormulaValue::Number(
        sign * (int_part + numerator / fraction),
    ))
}

pub fn fn_dollarfr(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let decimal_dollar = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let fraction = match required_number(args, 1) {
        Ok(v) => v.trunc(),
        Err(e) => return Ok(e),
    };

    if fraction < 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let sign = if decimal_dollar < 0.0 { -1.0 } else { 1.0 };
    let abs = decimal_dollar.abs();
    let int_part = abs.trunc();
    let frac = abs - int_part;
    Ok(FormulaValue::Number(
        sign * (int_part + (frac * fraction) / 100.0),
    ))
}

pub fn fn_duration(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let coupon = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let yld = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 4)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 5, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match duration_core(settlement, maturity, coupon, yld, frequency, basis) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(e),
    }
}

pub fn fn_fvschedule(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let principal = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let mut sched = Vec::new();
    let Some(value) = args.get(1) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    if let Err(e) = collect_numbers(value, &mut sched) {
        return Ok(e);
    }

    let mut fv = principal;
    for r in sched {
        fv *= 1.0 + r;
    }
    Ok(FormulaValue::Number(fv))
}

pub fn fn_intrate(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let investment = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let redemption = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 4, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if settlement >= maturity || investment <= 0.0 || redemption <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let yf = match yearfrac_basis(settlement, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if yf <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(
        (redemption - investment) / investment / yf,
    ))
}

pub fn fn_ispmt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
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

    if nper <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(pv * rate * (per / nper - 1.0)))
}

pub fn fn_mduration(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let coupon = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let yld = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 4)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 5, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match duration_core(settlement, maturity, coupon, yld, frequency, basis) {
        Ok(v) => Ok(FormulaValue::Number(v / (1.0 + yld / frequency as f64))),
        Err(e) => Ok(e),
    }
}

pub fn fn_oddfprice(
    _args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_oddfyield(
    _args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_oddlprice(
    _args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_oddlyield(
    _args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_price(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let yld = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let redemption = match required_number(args, 4) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 5)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 6, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match price_from_yield_core(
        settlement, maturity, rate, yld, redemption, frequency, basis,
    ) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(e),
    }
}

pub fn fn_pricedisc(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let discount = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let redemption = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 4, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if settlement >= maturity || discount < 0.0 || redemption <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let yf = match yearfrac_basis(settlement, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(redemption * (1.0 - discount * yf)))
}

pub fn fn_pricemat(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let issue = match required_serial(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let yld = match required_number(args, 4) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 5, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if settlement >= maturity || issue > settlement || rate < 0.0 || yld < 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let yf_issue = match yearfrac_basis(issue, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let yf_settle = match yearfrac_basis(settlement, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(
        (100.0 + 100.0 * rate * yf_issue) / (1.0 + yld * yf_settle),
    ))
}

pub fn fn_received(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let investment = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let discount = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 4, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if settlement >= maturity || investment <= 0.0 || discount <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let yf = match yearfrac_basis(settlement, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let denom = 1.0 - discount * yf;
    if denom <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(investment / denom))
}

pub fn fn_rri(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let nper = match required_number(args, 0) {
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

    if nper <= 0.0 || pv <= 0.0 || fv <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number((fv / pv).powf(1.0 / nper) - 1.0))
}

pub fn fn_tbilleq(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let discount = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let dsm = (maturity - settlement) as f64;
    if settlement >= maturity || dsm > 365.0 || discount <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let eq = if dsm <= 182.0 {
        (365.0 * discount) / (360.0 - discount * dsm)
    } else {
        let x = discount * dsm / 360.0;
        (2.0 * ((x * x - 2.0 * x + 1.0).sqrt() - 1.0)) / (dsm / 365.0)
    };
    Ok(FormulaValue::Number(eq))
}

pub fn fn_tbillprice(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let discount = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let dsm = (maturity - settlement) as f64;
    if settlement >= maturity || dsm > 365.0 || discount <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(100.0 * (1.0 - discount * dsm / 360.0)))
}

pub fn fn_tbillyield(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pr = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let dsm = (maturity - settlement) as f64;
    if settlement >= maturity || dsm > 365.0 || pr <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number((100.0 - pr) * 360.0 / (pr * dsm)))
}

pub fn fn_vdb(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
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
    let start_period = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let end_period = match required_number(args, 4) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let factor = match optional_number(args, 5, 2.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let no_switch = match optional_number(args, 6, 0.0) {
        Ok(v) => v != 0.0,
        Err(e) => return Ok(e),
    };

    if cost < 0.0
        || salvage < 0.0
        || life <= 0.0
        || start_period < 0.0
        || end_period < start_period
        || factor <= 0.0
    {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut book = cost;
    let mut total = 0.0;
    let end_whole = end_period.ceil() as i64;

    for i in 0..end_whole {
        if book <= salvage {
            break;
        }
        let period_start = i as f64;
        let period_end = i as f64 + 1.0;
        let overlap = (end_period.min(period_end) - start_period.max(period_start)).max(0.0);
        if overlap <= 0.0 && period_start >= end_period {
            break;
        }

        let mut dep = book * factor / life;
        if !no_switch {
            let remaining_life = (life - i as f64).max(1.0);
            let sl = (book - salvage).max(0.0) / remaining_life;
            if sl > dep {
                dep = sl;
            }
        }
        dep = dep.min((book - salvage).max(0.0));

        if overlap > 0.0 {
            total += dep * overlap;
        }
        book -= dep;
    }

    Ok(FormulaValue::Number(total))
}

pub fn fn_xirr(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut values = Vec::new();
    let mut dates = Vec::new();

    let Some(v) = args.get(0) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let Some(d) = args.get(1) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    if let Err(e) = collect_numbers(v, &mut values) {
        return Ok(e);
    }
    if let Err(e) = collect_dates(d, &mut dates) {
        return Ok(e);
    }
    let guess = match optional_number(args, 2, 0.1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if values.is_empty() || values.len() != dates.len() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let has_pos = values.iter().any(|v| *v > 0.0);
    let has_neg = values.iter().any(|v| *v < 0.0);
    if !has_pos || !has_neg {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let d0 = dates[0];
    let mut rate = if guess <= -0.999_999_999 { 0.1 } else { guess };

    for _ in 0..MAX_ITERATIONS {
        if rate <= -1.0 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        let base = 1.0 + rate;
        let mut f = 0.0;
        let mut df = 0.0;
        for (v, d) in values.iter().zip(dates.iter()) {
            let t = (d - d0) / 365.0;
            let p = base.powf(t);
            f += v / p;
            df -= t * v / base.powf(t + 1.0);
        }

        if f.abs() < TOLERANCE {
            return Ok(FormulaValue::Number(rate));
        }
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

pub fn fn_yield(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pr = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let redemption = match required_number(args, 4) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 5)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 6, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if pr <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut y = rate.max(0.01);
    for _ in 0..MAX_ITERATIONS {
        let p = match price_from_yield_core(
            settlement, maturity, rate, y, redemption, frequency, basis,
        ) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        let f = p - pr;
        if f.abs() < TOLERANCE {
            return Ok(FormulaValue::Number(y));
        }

        let h = 1e-7;
        let yl = (y - h).max(-(frequency as f64) + 1e-9);
        let yr = y + h;
        let pl = match price_from_yield_core(
            settlement, maturity, rate, yl, redemption, frequency, basis,
        ) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        let prr = match price_from_yield_core(
            settlement, maturity, rate, yr, redemption, frequency, basis,
        ) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        let df = (prr - pl) / (yr - yl);
        if !df.is_finite() || df.abs() < 1e-14 {
            return Ok(FormulaValue::Error(CellError::Num));
        }

        let next = y - f / df;
        if !next.is_finite() || next <= -(frequency as f64) {
            return Ok(FormulaValue::Error(CellError::Num));
        }
        if (next - y).abs() < TOLERANCE {
            return Ok(FormulaValue::Number(next));
        }
        y = next;
    }

    Ok(FormulaValue::Error(CellError::Num))
}

pub fn fn_yielddisc(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pr = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let redemption = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 4, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if settlement >= maturity || pr <= 0.0 || redemption <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let yf = match yearfrac_basis(settlement, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if yf <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number((redemption - pr) / pr / yf))
}

pub fn fn_yieldmat(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let settlement = match required_serial(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let maturity = match required_serial(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let issue = match required_serial(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pr = match required_number(args, 4) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 5, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if settlement >= maturity || issue > settlement || pr <= 0.0 || rate < 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let yf_issue = match yearfrac_basis(issue, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let yf_settle = match yearfrac_basis(settlement, maturity, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if yf_settle <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(
        ((100.0 + 100.0 * rate * yf_issue) / pr - 1.0) / yf_settle,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    fn as_number(v: FormulaValue) -> f64 {
        match v {
            FormulaValue::Number(n) => n,
            other => panic!("expected number, got {other:?}"),
        }
    }

    fn assert_finite(formula: &str) {
        let out = eval(formula).unwrap();
        let n = as_number(out);
        assert!(n.is_finite());
    }

    #[test]
    fn test_accrint() {
        assert_finite("=ACCRINT(45292,45474,45658,0.08,1000,2,0,1)");
    }

    #[test]
    fn test_accrintm() {
        assert_finite("=ACCRINTM(45292,45658,0.08,1000,0)");
    }

    #[test]
    fn test_amordegrc() {
        assert_finite("=AMORDEGRC(10000,45292,45474,1000,1,0.15,0)");
    }

    #[test]
    fn test_amorlinc() {
        assert_finite("=AMORLINC(10000,45292,45474,1000,1,0.15,0)");
    }

    #[test]
    fn test_coupdaybs() {
        assert_finite("=COUPDAYBS(45474,47484,2,0)");
    }

    #[test]
    fn test_coupdays() {
        assert_finite("=COUPDAYS(45474,47484,2,0)");
    }

    #[test]
    fn test_coupdaysnc() {
        assert_finite("=COUPDAYSNC(45474,47484,2,0)");
    }

    #[test]
    fn test_coupncd() {
        assert_finite("=COUPNCD(45474,47484,2,0)");
    }

    #[test]
    fn test_coupnum() {
        assert_finite("=COUPNUM(45474,47484,2,0)");
    }

    #[test]
    fn test_couppcd() {
        assert_finite("=COUPPCD(45474,47484,2,0)");
    }

    #[test]
    fn test_disc() {
        assert_finite("=DISC(45474,45658,97.5,100,0)");
    }

    #[test]
    fn test_dollarde() {
        let v = eval("=DOLLARDE(1.02,16)").unwrap();
        assert!((as_number(v) - 1.125).abs() < 1e-9);
    }

    #[test]
    fn test_dollarfr() {
        let v = eval("=DOLLARFR(1.125,16)").unwrap();
        assert!((as_number(v) - 1.02).abs() < 1e-9);
    }

    #[test]
    fn test_duration() {
        assert_finite("=DURATION(45474,47484,0.05,0.06,2,0)");
    }

    #[test]
    fn test_fvschedule() {
        let v = eval("=FVSCHEDULE(100,{0.1,0.2})").unwrap();
        assert!((as_number(v) - 132.0).abs() < 1e-9);
    }

    #[test]
    fn test_intrate() {
        assert_finite("=INTRATE(45474,45658,950,1000,0)");
    }

    #[test]
    fn test_ispmt() {
        assert_finite("=ISPMT(0.1/12,1,36,8000)");
    }

    #[test]
    fn test_mduration() {
        assert_finite("=MDURATION(45474,47484,0.05,0.06,2,0)");
    }

    #[test]
    fn test_oddfprice() {
        assert_eq!(
            eval("=ODDFPRICE(1,2,3,4,5,6,7,8,9)").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_oddfyield() {
        assert_eq!(
            eval("=ODDFYIELD(1,2,3,4,5,6,7,8,9)").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_oddlprice() {
        assert_eq!(
            eval("=ODDLPRICE(1,2,3,4,5,6,7,8)").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_oddlyield() {
        assert_eq!(
            eval("=ODDLYIELD(1,2,3,4,5,6,7,8)").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_price() {
        assert_finite("=PRICE(45474,47484,0.05,0.06,100,2,0)");
    }

    #[test]
    fn test_pricedisc() {
        assert_finite("=PRICEDISC(45474,45658,0.05,100,0)");
    }

    #[test]
    fn test_pricemat() {
        assert_finite("=PRICEMAT(45474,45658,45292,0.08,0.07,0)");
    }

    #[test]
    fn test_received() {
        assert_finite("=RECEIVED(45474,45658,950,0.05,0)");
    }

    #[test]
    fn test_rri() {
        let v = eval("=RRI(10,100,200)").unwrap();
        assert!((as_number(v) - 0.0717734625).abs() < 1e-8);
    }

    #[test]
    fn test_tbilleq() {
        assert_finite("=TBILLEQ(45474,45658,0.05)");
    }

    #[test]
    fn test_tbillprice() {
        assert_finite("=TBILLPRICE(45474,45658,0.05)");
    }

    #[test]
    fn test_tbillyield() {
        assert_finite("=TBILLYIELD(45474,45658,97.5)");
    }

    #[test]
    fn test_vdb() {
        assert_finite("=VDB(10000,1000,5,1,2,2,0)");
    }

    #[test]
    fn test_xirr() {
        assert_finite("=XIRR({-10000,2750,4250,3250,2750},{45292,45474,45658,45840,46023},0.1)");
    }

    #[test]
    fn test_yield() {
        assert_finite("=YIELD(45474,47484,0.05,95,100,2,0)");
    }

    #[test]
    fn test_yielddisc() {
        assert_finite("=YIELDDISC(45474,45658,97.5,100,0)");
    }

    #[test]
    fn test_yieldmat() {
        assert_finite("=YIELDMAT(45474,45658,45292,0.08,98,0)");
    }
}
