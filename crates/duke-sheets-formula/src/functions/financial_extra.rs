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

fn odd_last_coupon_fractions(
    settlement: i64,
    maturity: i64,
    last_interest: i64,
    frequency: i64,
    basis: i64,
) -> Result<(f64, f64, f64), FormulaValue> {
    if !matches!(frequency, 1 | 2 | 4) || !(0..=4).contains(&basis) {
        return Err(FormulaValue::Error(CellError::Num));
    }
    if !(last_interest < settlement && settlement < maturity) {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let accrued = yearfrac_basis(last_interest, settlement, basis)?;
    let odd_period = yearfrac_basis(last_interest, maturity, basis)?;
    let discount_period = yearfrac_basis(settlement, maturity, basis)?;

    if accrued < 0.0 || odd_period <= 0.0 || discount_period <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    Ok((accrued, odd_period, discount_period))
}

fn odd_last_price_core(
    settlement: i64,
    maturity: i64,
    last_interest: i64,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i64,
    basis: i64,
) -> Result<f64, FormulaValue> {
    if rate < 0.0 || yld < 0.0 || redemption <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let (accrued, odd_period, discount_period) =
        odd_last_coupon_fractions(settlement, maturity, last_interest, frequency, basis)?;
    let denom = 1.0 + yld * discount_period;
    if !denom.is_finite() || denom <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    Ok((redemption + 100.0 * rate * odd_period) / denom - 100.0 * rate * accrued)
}

fn odd_last_yield_core(
    settlement: i64,
    maturity: i64,
    last_interest: i64,
    rate: f64,
    pr: f64,
    redemption: f64,
    frequency: i64,
    basis: i64,
) -> Result<f64, FormulaValue> {
    if rate < 0.0 || pr <= 0.0 || redemption <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let (accrued, odd_period, discount_period) =
        odd_last_coupon_fractions(settlement, maturity, last_interest, frequency, basis)?;
    let denom = pr + 100.0 * rate * accrued;
    if !denom.is_finite() || denom <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let yld = ((redemption + 100.0 * rate * odd_period) / denom - 1.0) / discount_period;
    if !yld.is_finite() {
        return Err(FormulaValue::Error(CellError::Num));
    }

    Ok(yld)
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

fn odd_coupon_day_count(start: i64, end: i64, basis: i64) -> Result<f64, FormulaValue> {
    match basis {
        0 => day_count_days(start, end, 0),
        4 => day_count_days(start, end, 4),
        1..=3 => day_count_days(start, end, 1),
        _ => Err(FormulaValue::Error(CellError::Num)),
    }
}

fn coupon_add_months(date: NaiveDate, months: i32) -> Option<NaiveDate> {
    let total = date.year() * 12 + date.month0() as i32 + months;
    let year = total.div_euclid(12);
    let month0 = total.rem_euclid(12) as u32;
    let month = month0 + 1;
    let last = last_day_of_month(year, month)?;
    let source_last = date.day() == last_day_of_month(date.year(), date.month())?;
    let day = if source_last {
        last
    } else {
        date.day().min(last)
    };
    NaiveDate::from_ymd_opt(year, month, day)
}

fn odd_coupon_period_length(
    start: NaiveDate,
    end: NaiveDate,
    frequency: i64,
    basis: i64,
) -> Result<f64, FormulaValue> {
    match basis {
        0 | 2 | 4 => Ok(360.0 / frequency as f64),
        3 => Ok(365.0 / frequency as f64),
        1 => odd_coupon_day_count(date_to_serial(start), date_to_serial(end), 1),
        _ => Err(FormulaValue::Error(CellError::Num)),
    }
}

fn regular_coupon_count(
    first_coupon: NaiveDate,
    maturity: NaiveDate,
    frequency: i64,
) -> Result<usize, FormulaValue> {
    let months = (12 / frequency) as i32;
    let mut date = first_coupon;
    let mut count = 0usize;

    while date < maturity {
        date = coupon_add_months(date, months).ok_or(FormulaValue::Error(CellError::Num))?;
        count += 1;
        if date > maturity {
            return Err(FormulaValue::Error(CellError::Num));
        }
    }

    if date != maturity {
        return Err(FormulaValue::Error(CellError::Num));
    }

    Ok(count)
}

fn odd_first_price_core(
    settlement: i64,
    maturity: i64,
    issue: i64,
    first_coupon: i64,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i64,
    basis: i64,
) -> Result<f64, FormulaValue> {
    if !matches!(frequency, 1 | 2 | 4) || !(0..=4).contains(&basis) {
        return Err(FormulaValue::Error(CellError::Num));
    }
    if rate < 0.0 || redemption <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }
    if yld <= -(frequency as f64) {
        return Err(FormulaValue::Error(CellError::Num));
    }
    if settlement >= maturity
        || issue >= settlement
        || first_coupon >= maturity
        || first_coupon <= settlement
    {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let settlement_date = serial_to_date(settlement).ok_or(FormulaValue::Error(CellError::Num))?;
    let maturity_date = serial_to_date(maturity).ok_or(FormulaValue::Error(CellError::Num))?;
    let issue_date = serial_to_date(issue).ok_or(FormulaValue::Error(CellError::Num))?;
    let first_coupon_date =
        serial_to_date(first_coupon).ok_or(FormulaValue::Error(CellError::Num))?;

    let months = (12 / frequency) as i32;
    let mut quasi_dates = vec![first_coupon_date];
    let mut date = first_coupon_date;
    loop {
        let prev = coupon_add_months(date, -months).ok_or(FormulaValue::Error(CellError::Num))?;
        quasi_dates.push(prev);
        if prev <= issue_date {
            break;
        }
        date = prev;
    }
    quasi_dates.reverse();

    let remaining_coupons = regular_coupon_count(first_coupon_date, maturity_date, frequency)?;
    let coupon_payment = redemption * rate / frequency as f64;
    let per_yield = yld / frequency as f64;
    let base = 1.0 + per_yield;
    if !base.is_finite() || base <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let mut first_coupon_amount = 0.0;
    let mut accrued_interest = 0.0;
    let mut time_to_first_coupon = None;

    for (idx, window) in quasi_dates.windows(2).enumerate() {
        let period_start = window[0];
        let period_end = window[1];
        let period_length = odd_coupon_period_length(period_start, period_end, frequency, basis)?;
        let effective_start = issue_date.max(period_start);
        let period_end_serial = date_to_serial(period_end);

        let first_coupon_days =
            odd_coupon_day_count(date_to_serial(effective_start), period_end_serial, basis)?;
        first_coupon_amount += coupon_payment * first_coupon_days / period_length;

        let accrual_end = settlement_date.min(period_end);
        if accrual_end > effective_start {
            let accrued_days = odd_coupon_day_count(
                date_to_serial(effective_start),
                date_to_serial(accrual_end),
                basis,
            )?;
            accrued_interest += coupon_payment * accrued_days / period_length;
        }

        if time_to_first_coupon.is_none() && settlement_date < period_end {
            let dsc = odd_coupon_day_count(settlement, period_end_serial, basis)?;
            let whole_quasi_periods = (quasi_dates.len() - idx - 2) as f64;
            time_to_first_coupon = Some(dsc / period_length + whole_quasi_periods);
        }
    }

    let time_to_first_coupon = time_to_first_coupon.ok_or(FormulaValue::Error(CellError::Num))?;
    let first_discount = base.powf(time_to_first_coupon);
    if !first_discount.is_finite() || first_discount <= 0.0 {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let mut price = first_coupon_amount / first_discount;
    for k in 1..=remaining_coupons {
        let t = time_to_first_coupon + k as f64;
        let discount = base.powf(t);
        if !discount.is_finite() || discount <= 0.0 {
            return Err(FormulaValue::Error(CellError::Num));
        }
        price += coupon_payment / discount;
        if k == remaining_coupons {
            price += redemption / discount;
        }
    }

    Ok(price - accrued_interest)
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

#[allow(clippy::too_many_arguments)]
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

    let year_base = match basis {
        1 => {
            // Actual/actual: use actual days in the year of the purchase date
            if let Some(d) = serial_to_date(date_purchased) {
                if is_leap_year(d.year()) {
                    366.0
                } else {
                    365.0
                }
            } else {
                365.0
            }
        }
        3 => 365.0,
        _ => 360.0,
    };
    let first_days = day_count_days(date_purchased, first_period, basis)?.abs();
    let mut remaining = cost;
    let mut dep = cost * rate;
    if degrc {
        dep *= amordegrc_coeff(rate);
    }
    let mut first_dep = dep * (first_days / year_base);
    if degrc {
        first_dep = first_dep.round();
    }
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
            this_dep = this_dep.round();
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
    let issue = match required_serial(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let first_coupon = match required_serial(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 4) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let yld = match required_number(args, 5) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let redemption = match required_number(args, 6) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 7)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 8, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if rate < 0.0 || yld < 0.0 || redemption <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    match odd_first_price_core(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
    ) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(e),
    }
}

pub fn fn_oddfyield(
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
    let issue = match required_serial(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let first_coupon = match required_serial(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let rate = match required_number(args, 4) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let pr = match required_number(args, 5) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let redemption = match required_number(args, 6) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 7)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 8, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if rate < 0.0 || pr <= 0.0 || redemption <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut y = rate.max(0.01);
    for _ in 0..MAX_ITERATIONS {
        let price = match odd_first_price_core(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            y,
            redemption,
            frequency,
            basis,
        ) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        let f = price - pr;
        if f.abs() < TOLERANCE {
            return Ok(FormulaValue::Number(y));
        }

        let h = 1e-7;
        let yl = (y - h).max(-(frequency as f64) + 1e-9);
        let yr = y + h;
        let pl = match odd_first_price_core(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yl,
            redemption,
            frequency,
            basis,
        ) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        let prr = match odd_first_price_core(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yr,
            redemption,
            frequency,
            basis,
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

pub fn fn_oddlprice(
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
    let last_interest = match required_serial(args, 2) {
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
    let redemption = match required_number(args, 5) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 6)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 7, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match odd_last_price_core(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
    ) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(e),
    }
}

pub fn fn_oddlyield(
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
    let last_interest = match required_serial(args, 2) {
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
    let redemption = match required_number(args, 5) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let frequency = match required_number(args, 6)
        .and_then(|v| parse_frequency(v).ok_or(FormulaValue::Error(CellError::Num)))
    {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match optional_basis(args, 7, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match odd_last_yield_core(
        settlement,
        maturity,
        last_interest,
        rate,
        pr,
        redemption,
        frequency,
        basis,
    ) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(e),
    }
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
    let yf_a = match yearfrac_basis(issue, settlement, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(
        (100.0 + 100.0 * rate * yf_issue) / (1.0 + yld * yf_settle) - 100.0 * rate * yf_a,
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

    let Some(v) = args.first() else {
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
    let yf_a = match yearfrac_basis(issue, settlement, basis) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if yf_settle <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let term = pr / 100.0 + rate * yf_a;
    Ok(FormulaValue::Number(
        ((1.0 + rate * yf_issue) / term - 1.0) / yf_settle,
    ))
}

/// Helper: look up euro conversion rate and calc precision for a currency code.
/// Returns (rate_per_eur, calc_precision_decimal_places).
fn euro_currency(code: &str) -> Option<(f64, i32)> {
    match code {
        "EUR" => Some((1.0, 2)),
        "BEF" => Some((40.3399, 0)),
        "LUF" => Some((40.3399, 0)),
        "DEM" => Some((1.95583, 2)),
        "ESP" => Some((166.386, 0)),
        "FRF" => Some((6.55957, 2)),
        "IEP" => Some((0.787564, 2)),
        "ITL" => Some((1936.27, 0)),
        "NLG" => Some((2.20371, 2)),
        "ATS" => Some((13.7603, 2)),
        "PTE" => Some((200.482, 0)),
        "FIM" => Some((5.94573, 2)),
        "GRD" => Some((340.750, 0)),
        "SIT" => Some((239.640, 2)),
        "CYP" => Some((0.585274, 2)),
        "MTL" => Some((0.429300, 2)),
        "SKK" => Some((30.1260, 2)),
        "EEK" => Some((15.6466, 2)),
        "LVL" => Some((0.702804, 2)),
        "LTL" => Some((3.45280, 2)),
        "HRK" => Some((7.53450, 2)),
        _ => None,
    }
}

/// Round a value to n significant digits.
fn round_significant(x: f64, n: u32) -> f64 {
    if x == 0.0 {
        return 0.0;
    }
    let magnitude = x.abs().log10().floor() as i32;
    let factor = 10f64.powi(n as i32 - 1 - magnitude);
    (x * factor).round() / factor
}

/// Round a value to n decimal places (half-up).
fn round_decimal(x: f64, dp: i32) -> f64 {
    let factor = 10f64.powi(dp);
    (x * factor + 0.5f64.copysign(x)).floor().copysign(x) / factor
}

/// EUROCONVERT(number, source, target, [full_precision], [triangulation_precision])
pub fn fn_euroconvert(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let source = match args.get(1) {
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let target = match args.get(2) {
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let full_precision = match args.get(3).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(v) => v.as_bool().unwrap_or(false),
        None => false,
    };
    let tri_precision = match args.get(4).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(v) => {
            let n = match scalar_number(v) {
                Ok(n) => n.trunc() as i32,
                Err(e) => return Ok(e),
            };
            if n < 3 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            Some(n as u32)
        }
        None => None,
    };

    let src = source.trim().to_ascii_uppercase();
    let tgt = target.trim().to_ascii_uppercase();

    // Same currency: return unchanged
    if src == tgt {
        return Ok(FormulaValue::Number(number));
    }

    let (src_rate, _) = match euro_currency(&src) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let (tgt_rate, tgt_prec) = match euro_currency(&tgt) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Step 1: convert to EUR
    let eur_value = if src == "EUR" {
        number
    } else {
        number / src_rate
    };

    // Step 2: optionally round intermediate EUR to significant digits
    let eur_rounded = match tri_precision {
        Some(prec) => round_significant(eur_value, prec),
        None => eur_value,
    };

    // Step 3: convert from EUR to target
    let result = if tgt == "EUR" {
        eur_rounded
    } else {
        eur_rounded * tgt_rate
    };

    // Step 4: apply rounding if not full precision
    if full_precision {
        Ok(FormulaValue::Number(result))
    } else {
        Ok(FormulaValue::Number(round_decimal(result, tgt_prec)))
    }
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
    fn test_oddfprice_docs() {
        let v = eval("=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)").unwrap();
        assert!((as_number(v) - 113.59771747407883).abs() < 1e-9);
    }

    #[test]
    fn test_oddfyield_docs() {
        let v = eval("=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)").unwrap();
        assert!((as_number(v) - 0.07724554159781691).abs() < 1e-9);
    }

    #[test]
    fn test_oddlprice() {
        assert_eq!(
            eval("=ODDLPRICE(1,2,3,4,5,6,7,8)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_oddlyield() {
        assert_eq!(
            eval("=ODDLYIELD(1,2,3,4,5,6,7,8)").unwrap(),
            FormulaValue::Error(CellError::Num)
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

    // ---------- EUROCONVERT tests ----------

    #[test]
    fn test_euroconvert_same_currency() {
        // Same currency returns unchanged
        assert_eq!(
            eval("=EUROCONVERT(100,\"DEM\",\"DEM\")").unwrap(),
            FormulaValue::Number(100.0)
        );
    }

    #[test]
    fn test_euroconvert_dem_to_eur() {
        // 100 DEM -> EUR: 100 / 1.95583 = 51.129... -> rounded to 2dp = 51.13
        let result = eval("=EUROCONVERT(100,\"DEM\",\"EUR\")").unwrap();
        match result {
            FormulaValue::Number(n) => assert!((n - 51.13).abs() < 0.005, "got {}", n),
            other => panic!("expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_euroconvert_eur_to_dem() {
        // 100 EUR -> DEM: 100 * 1.95583 = 195.583 -> rounded to 2dp = 195.58
        let result = eval("=EUROCONVERT(100,\"EUR\",\"DEM\")").unwrap();
        match result {
            FormulaValue::Number(n) => assert!((n - 195.58).abs() < 0.005, "got {}", n),
            other => panic!("expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_euroconvert_cross_currency() {
        // 1 FRF -> DEM: 1 / 6.55957 * 1.95583 = 0.2981 -> rounded to 2dp = 0.30
        let result = eval("=EUROCONVERT(1,\"FRF\",\"DEM\")").unwrap();
        match result {
            FormulaValue::Number(n) => assert!((n - 0.30).abs() < 0.005, "got {}", n),
            other => panic!("expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_euroconvert_full_precision() {
        // 1 FRF -> DEM with full precision
        let result = eval("=EUROCONVERT(1,\"FRF\",\"DEM\",TRUE)").unwrap();
        match result {
            FormulaValue::Number(n) => {
                // Should be ~0.29808...
                assert!((n - 0.298164).abs() < 0.001, "got {}", n);
            }
            other => panic!("expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_euroconvert_triangulation_precision() {
        // 1 FRF -> DEM with triangulation precision 3
        // 1 / 6.55957 = 0.15244... -> rounded to 3 sig figs = 0.152
        // 0.152 * 1.95583 = 0.29728...
        let result = eval("=EUROCONVERT(1,\"FRF\",\"DEM\",TRUE,3)").unwrap();
        match result {
            FormulaValue::Number(n) => assert!((n - 0.29729).abs() < 0.001, "got {}", n),
            other => panic!("expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_euroconvert_integer_currencies() {
        // ITL (lira) has calc_precision 0 -> rounds to integer
        // 100 EUR -> ITL: 100 * 1936.27 = 193627.0
        let result = eval("=EUROCONVERT(100,\"EUR\",\"ITL\")").unwrap();
        match result {
            FormulaValue::Number(n) => assert!((n - 193627.0).abs() < 0.5, "got {}", n),
            other => panic!("expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_euroconvert_zero() {
        assert_eq!(
            eval("=EUROCONVERT(0,\"DEM\",\"EUR\")").unwrap(),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_euroconvert_negative() {
        // Negative values convert normally
        let result = eval("=EUROCONVERT(-100,\"DEM\",\"EUR\")").unwrap();
        match result {
            FormulaValue::Number(n) => assert!(n < 0.0 && (n + 51.13).abs() < 0.01, "got {}", n),
            other => panic!("expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_euroconvert_invalid_currency() {
        assert_eq!(
            eval("=EUROCONVERT(100,\"USD\",\"EUR\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            eval("=EUROCONVERT(100,\"EUR\",\"GBP\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_euroconvert_triangulation_too_small() {
        // Triangulation precision < 3 -> #VALUE!
        assert_eq!(
            eval("=EUROCONVERT(1,\"FRF\",\"DEM\",FALSE,2)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_euroconvert_all_currencies() {
        // Verify all 21 currencies are recognized
        let codes = [
            "EUR", "BEF", "LUF", "DEM", "ESP", "FRF", "IEP", "ITL", "NLG", "ATS", "PTE", "FIM",
            "GRD", "SIT", "CYP", "MTL", "SKK", "EEK", "LVL", "LTL", "HRK",
        ];
        for code in &codes {
            let formula = format!("=EUROCONVERT(100,\"EUR\",\"{}\")", code);
            let result = eval(&formula).unwrap();
            match result {
                FormulaValue::Number(_) => {} // OK
                other => panic!("currency {} failed: {:?}", code, other),
            }
        }
    }

    // ===== Docs-based tests =====

    fn n(x: f64) -> FormulaValue {
        FormulaValue::Number(x)
    }

    fn arr(values: &[f64]) -> FormulaValue {
        FormulaValue::Array(vec![values
            .iter()
            .map(|v| FormulaValue::Number(*v))
            .collect()])
    }

    fn ctx() -> EvaluationContext<'static> {
        EvaluationContext::simple()
    }

    fn assert_close(actual: f64, expected: f64, tol: f64) {
        assert!(
            (actual - expected).abs() <= tol,
            "actual={actual}, expected={expected}, tol={tol}"
        );
    }

    // XIRR docs: =XIRR(values, dates, 0.1) = 0.373362535
    #[test]
    fn test_xirr_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_xirr(
                    &[
                        arr(&[-10000.0, 2750.0, 4250.0, 3250.0, 2750.0]),
                        arr(&[39448.0, 39508.0, 39751.0, 39859.0, 39904.0]),
                        n(0.1),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            0.373363,
            1e-3,
        );
    }

    // ISPMT docs: =ISPMT(0.10, 0, 4, 4000) = -400
    #[test]
    fn test_ispmt_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_ispmt(&[n(0.10), n(0.0), n(4.0), n(4000.0)], &c).unwrap()),
            -400.0,
            1e-2,
        );
    }

    // ISPMT docs: period 1
    #[test]
    fn test_ispmt_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_ispmt(&[n(0.10), n(1.0), n(4.0), n(4000.0)], &c).unwrap()),
            -300.0,
            1e-2,
        );
    }

    // ISPMT docs: period 2
    #[test]
    fn test_ispmt_docs_3() {
        let c = ctx();
        assert_close(
            as_number(fn_ispmt(&[n(0.10), n(2.0), n(4.0), n(4000.0)], &c).unwrap()),
            -200.0,
            1e-2,
        );
    }

    // ISPMT docs: period 3
    #[test]
    fn test_ispmt_docs_4() {
        let c = ctx();
        assert_close(
            as_number(fn_ispmt(&[n(0.10), n(3.0), n(4.0), n(4000.0)], &c).unwrap()),
            -100.0,
            1e-2,
        );
    }

    // VDB docs: daily, first day
    #[test]
    fn test_vdb_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_vdb(&[n(2400.0), n(300.0), n(3650.0), n(0.0), n(1.0)], &c).unwrap()),
            1.32,
            1e-2,
        );
    }

    // VDB docs: monthly, first month
    #[test]
    fn test_vdb_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_vdb(&[n(2400.0), n(300.0), n(120.0), n(0.0), n(1.0)], &c).unwrap()),
            40.00,
            1e-2,
        );
    }

    // VDB docs: yearly, first year
    #[test]
    fn test_vdb_docs_3() {
        let c = ctx();
        assert_close(
            as_number(fn_vdb(&[n(2400.0), n(300.0), n(10.0), n(0.0), n(1.0)], &c).unwrap()),
            480.00,
            1e-2,
        );
    }

    // VDB docs: months 6-18
    #[test]
    fn test_vdb_docs_4() {
        let c = ctx();
        assert_close(
            as_number(fn_vdb(&[n(2400.0), n(300.0), n(120.0), n(6.0), n(18.0)], &c).unwrap()),
            396.31,
            1e-2,
        );
    }

    // VDB docs: months 6-18, factor=1.5
    #[test]
    fn test_vdb_docs_5() {
        let c = ctx();
        assert_close(
            as_number(
                fn_vdb(
                    &[n(2400.0), n(300.0), n(120.0), n(6.0), n(18.0), n(1.5)],
                    &c,
                )
                .unwrap(),
            ),
            311.81,
            1e-2,
        );
    }

    // RRI docs: =RRI(96, 10000, 11000) = 0.0009933
    #[test]
    fn test_rri_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_rri(&[n(96.0), n(10000.0), n(11000.0)], &c).unwrap()),
            0.0009933,
            1e-5,
        );
    }

    // FVSCHEDULE docs: =FVSCHEDULE(1, {0.09, 0.11, 0.1}) = 1.33089
    #[test]
    fn test_fvschedule_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_fvschedule(&[n(1.0), arr(&[0.09, 0.11, 0.1])], &c).unwrap()),
            1.33089,
            1e-4,
        );
    }

    // DOLLARDE docs: =DOLLARDE(1.02, 16) = 1.125
    #[test]
    fn test_dollarde_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_dollarde(&[n(1.02), n(16.0)], &c).unwrap()),
            1.125,
            1e-4,
        );
    }

    // DOLLARDE docs: =DOLLARDE(1.1, 32) = 1.3125
    #[test]
    fn test_dollarde_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_dollarde(&[n(1.1), n(32.0)], &c).unwrap()),
            1.3125,
            1e-4,
        );
    }

    // DOLLARFR docs: =DOLLARFR(1.125, 16) = 1.02
    #[test]
    fn test_dollarfr_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_dollarfr(&[n(1.125), n(16.0)], &c).unwrap()),
            1.02,
            1e-4,
        );
    }

    // DOLLARFR docs: =DOLLARFR(1.125, 32) = 1.04
    #[test]
    fn test_dollarfr_docs_2() {
        let c = ctx();
        assert_close(
            as_number(fn_dollarfr(&[n(1.125), n(32.0)], &c).unwrap()),
            1.04,
            1e-4,
        );
    }

    // ACCRINT docs: basis=0
    #[test]
    fn test_accrint_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_accrint(
                    &[
                        n(39508.0),
                        n(39691.0),
                        n(39569.0),
                        n(0.1),
                        n(1000.0),
                        n(2.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            16.6667,
            1e-2,
        );
    }

    // ACCRINTM docs: basis=3 (Actual/365)
    #[test]
    fn test_accrintm_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_accrintm(&[n(39539.0), n(39614.0), n(0.1), n(1000.0), n(3.0)], &c).unwrap(),
            ),
            20.5479,
            1e-2,
        );
    }

    // INTRATE docs: basis=2 (Actual/360)
    #[test]
    fn test_intrate_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_intrate(
                    &[n(39493.0), n(39583.0), n(1000000.0), n(1014420.0), n(2.0)],
                    &c,
                )
                .unwrap(),
            ),
            0.05768,
            1e-4,
        );
    }

    // COUPDAYBS docs: DATE(2011,1,25)=40568, DATE(2011,11,15)=40862, freq=2, basis=1
    #[test]
    fn test_coupdaybs_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_coupdaybs(&[n(40568.0), n(40862.0), n(2.0), n(1.0)], &c).unwrap()),
            71.0,
            1e-2,
        );
    }

    // COUPDAYS docs
    #[test]
    fn test_coupdays_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_coupdays(&[n(40568.0), n(40862.0), n(2.0), n(1.0)], &c).unwrap()),
            181.0,
            1e-2,
        );
    }

    // COUPDAYSNC docs
    #[test]
    fn test_coupdaysnc_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_coupdaysnc(&[n(40568.0), n(40862.0), n(2.0), n(1.0)], &c).unwrap()),
            110.0,
            1e-2,
        );
    }

    // COUPNCD docs: returns DATE(2011,5,15) = 40678
    #[test]
    fn test_coupncd_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_coupncd(&[n(40568.0), n(40862.0), n(2.0), n(1.0)], &c).unwrap()),
            40678.0,
            1e-2,
        );
    }

    // COUPNUM docs: DATE(2007,1,25)=39107, DATE(2008,11,15)=39767
    #[test]
    fn test_coupnum_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_coupnum(&[n(39107.0), n(39767.0), n(2.0), n(1.0)], &c).unwrap()),
            4.0,
            1e-2,
        );
    }

    // COUPPCD docs: returns DATE(2010,11,15) = 40497
    #[test]
    fn test_couppcd_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_couppcd(&[n(40568.0), n(40862.0), n(2.0), n(1.0)], &c).unwrap()),
            40497.0,
            1e-2,
        );
    }

    // ===== Batch 8 docs tests: DISC, DURATION, MDURATION, PRICE, RECEIVED =====

    // DISC docs: DATE(2018,7,1)=43282, DATE(2048,1,1)=54058
    #[test]
    fn test_disc_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_disc(&[n(43282.0), n(54058.0), n(97.975), n(100.0), n(1.0)], &c).unwrap()),
            0.000688,
            1e-4,
        );
    }

    // DURATION docs: DATE(2018,7,1)=43282, DATE(2048,1,1)=54058
    #[test]
    fn test_duration_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_duration(
                    &[n(43282.0), n(54058.0), n(0.08), n(0.09), n(2.0), n(1.0)],
                    &c,
                )
                .unwrap(),
            ),
            10.9191,
            1e-2,
        );
    }

    // MDURATION docs: DATE(2008,1,1)=39448, DATE(2016,1,1)=42370
    #[test]
    fn test_mduration_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_mduration(
                    &[n(39448.0), n(42370.0), n(0.08), n(0.09), n(2.0), n(1.0)],
                    &c,
                )
                .unwrap(),
            ),
            5.736,
            1e-2,
        );
    }

    #[test]
    fn test_oddlprice_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_oddlprice(
                    &[
                        n(39485.0),
                        n(39614.0),
                        n(39370.0),
                        n(0.0375),
                        n(0.0405),
                        n(100.0),
                        n(2.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            99.87828601472134,
            1e-10,
        );
    }

    #[test]
    fn test_oddlyield_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_oddlyield(
                    &[
                        n(39558.0),
                        n(39614.0),
                        n(39440.0),
                        n(0.0375),
                        n(99.875),
                        n(100.0),
                        n(2.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            0.04519223562916898,
            1e-10,
        );
    }

    // PRICE docs: DATE(2008,2,15)=39493, DATE(2017,11,15)=43054
    #[test]
    fn test_price_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_price(
                    &[
                        n(39493.0),
                        n(43054.0),
                        n(0.0575),
                        n(0.065),
                        n(100.0),
                        n(2.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            94.63,
            1e-1,
        );
    }

    // RECEIVED docs: DATE(2008,2,15)=39493, DATE(2008,5,15)=39583
    #[test]
    fn test_received_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_received(
                    &[n(39493.0), n(39583.0), n(1000000.0), n(0.0575), n(2.0)],
                    &c,
                )
                .unwrap(),
            ),
            1014584.65,
            1e-1,
        );
    }

    // ===== Batch 9 docs tests: PRICEDISC, PRICEMAT, TBILLEQ, TBILLPRICE, TBILLYIELD =====

    // PRICEDISC docs: DATE(2008,2,16)=39494, DATE(2008,3,1)=39508
    #[test]
    fn test_pricedisc_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_pricedisc(&[n(39494.0), n(39508.0), n(0.0525), n(100.0), n(2.0)], &c).unwrap(),
            ),
            99.79583,
            1e-3,
        );
    }

    // PRICEMAT docs: DATE(2008,2,15)=39493, DATE(2008,4,13)=39551, DATE(2007,11,11)=39397
    #[test]
    fn test_pricemat_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_pricemat(
                    &[
                        n(39493.0),
                        n(39551.0),
                        n(39397.0),
                        n(0.061),
                        n(0.061),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            99.98449,
            1e-2,
        );
    }

    // TBILLEQ docs: DATE(2008,3,31)=39538, DATE(2008,6,1)=39600
    #[test]
    fn test_tbilleq_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_tbilleq(&[n(39538.0), n(39600.0), n(0.0914)], &c).unwrap()),
            0.09415,
            1e-4,
        );
    }

    // TBILLPRICE docs: DATE(2008,3,31)=39538, DATE(2008,6,1)=39600
    #[test]
    fn test_tbillprice_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_tbillprice(&[n(39538.0), n(39600.0), n(0.09)], &c).unwrap()),
            98.45,
            1e-3,
        );
    }

    // TBILLYIELD docs: DATE(2008,3,31)=39538, DATE(2008,6,1)=39600
    #[test]
    fn test_tbillyield_docs() {
        let c = ctx();
        assert_close(
            as_number(fn_tbillyield(&[n(39538.0), n(39600.0), n(98.45)], &c).unwrap()),
            0.09142,
            1e-4,
        );
    }

    // ===== Batch 10 docs tests: YIELD, YIELDDISC, YIELDMAT, AMORDEGRC, AMORLINC =====

    // YIELD docs: DATE(2008,2,15)=39493, DATE(2016,11,15)=42689
    #[test]
    fn test_yield_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_yield(
                    &[
                        n(39493.0),
                        n(42689.0),
                        n(0.0575),
                        n(95.04287),
                        n(100.0),
                        n(2.0),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            0.065,
            1e-3,
        );
    }

    // YIELDDISC docs: DATE(2008,2,16)=39494, DATE(2008,3,1)=39508
    #[test]
    fn test_yielddisc_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_yielddisc(&[n(39494.0), n(39508.0), n(99.795), n(100.0), n(2.0)], &c).unwrap(),
            ),
            0.052823,
            1e-3,
        );
    }

    // YIELDMAT docs: DATE(2008,3,15)=39522, DATE(2008,11,3)=39755, DATE(2007,11,8)=39394
    #[test]
    fn test_yieldmat_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_yieldmat(
                    &[
                        n(39522.0),
                        n(39755.0),
                        n(39394.0),
                        n(0.0625),
                        n(100.0123),
                        n(0.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            0.060954,
            1e-3,
        );
    }

    // AMORDEGRC docs: cost=2400, date_purchased=DATE(2008,8,19)=39679, first_period=DATE(2008,12,31)=39813
    #[test]
    fn test_amordegrc_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_amordegrc(
                    &[
                        n(2400.0),
                        n(39679.0),
                        n(39813.0),
                        n(300.0),
                        n(1.0),
                        n(0.15),
                        n(1.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            776.0,
            1e-1,
        );
    }

    // AMORLINC docs: cost=2400, date_purchased=DATE(2008,8,19)=39679, first_period=DATE(2008,12,31)=39813
    #[test]
    fn test_amorlinc_docs() {
        let c = ctx();
        assert_close(
            as_number(
                fn_amorlinc(
                    &[
                        n(2400.0),
                        n(39679.0),
                        n(39813.0),
                        n(300.0),
                        n(1.0),
                        n(0.15),
                        n(1.0),
                    ],
                    &c,
                )
                .unwrap(),
            ),
            360.0,
            1e-1,
        );
    }
}
