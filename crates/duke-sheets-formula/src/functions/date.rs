//! Date/time functions
//!
//! This module implements a small subset of Excel date functions using Excel serial date numbers.
//!
//! Notes:
//! - Excel stores dates as serial numbers (days since a base date).
//! - In the 1900 date system, Excel includes the historical "1900 leap year" bug, inserting
//!   a non-existent day 1900-02-29 as serial 60.

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use chrono::{Datelike, Duration, NaiveDate, Timelike};
use duke_sheets_core::CellError;
use std::collections::HashSet;

fn ctx_date_1904(ctx: &EvaluationContext) -> bool {
    ctx.workbook
        .map(|wb| wb.settings().date_1904)
        .unwrap_or(false)
}

fn to_i32_trunc(v: &FormulaValue) -> Option<i32> {
    v.as_number().map(|n| n.trunc() as i32)
}

fn is_leap_gregorian(year: i32) -> bool {
    (year % 4 == 0) && ((year % 100 != 0) || (year % 400 == 0))
}

fn days_in_year_excel1900(year: i32) -> i64 {
    if year == 1900 || is_leap_gregorian(year) {
        366
    } else {
        365
    }
}

fn days_in_month_excel1900(year: i32, month: u32) -> i64 {
    match month {
        1 => 31,
        2 => {
            if year == 1900 || is_leap_gregorian(year) {
                29
            } else {
                28
            }
        }
        3 => 31,
        4 => 30,
        5 => 31,
        6 => 30,
        7 => 31,
        8 => 31,
        9 => 30,
        10 => 31,
        11 => 30,
        12 => 31,
        _ => 30,
    }
}

/// Compute the Excel serial (1900 system, with the 1900 leap-year bug) for the first day of
/// the given month/year.
///
/// Returns serial where 1900-01-01 == 1.
fn excel1900_serial_month_start(year: i32, month: u32) -> i64 {
    // We assume year is in a reasonable range (Excel supports up to 9999).
    let mut days: i64 = 0;
    if year >= 1900 {
        for y in 1900..year {
            days += days_in_year_excel1900(y);
        }
        for m in 1..month {
            days += days_in_month_excel1900(year, m);
        }
        1 + days
    } else {
        // For years before 1900, fall back to Gregorian chrono-based serial.
        // (The Excel bug does not apply before 1900.)
        let base = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
        let d = match NaiveDate::from_ymd_opt(year, month, 1) {
            Some(d) => d,
            None => return 0,
        };
        (d - base).num_days()
    }
}

fn excel1900_serial_from_ymd(year: i32, month: u32, day: i32) -> i64 {
    excel1900_serial_month_start(year, month) + (day as i64) - 1
}

fn excel1904_serial_from_date(date: NaiveDate) -> i64 {
    let base = NaiveDate::from_ymd_opt(1904, 1, 1).unwrap();
    (date - base).num_days()
}

fn excel1904_date_from_serial(serial: i64) -> Option<NaiveDate> {
    let base = NaiveDate::from_ymd_opt(1904, 1, 1)?;
    base.checked_add_signed(Duration::days(serial))
}

fn excel1900_date_from_serial(serial: i64) -> Option<(i32, u32, u32)> {
    // Serial 60 is the fictional 1900-02-29.
    if serial == 60 {
        return Some((1900, 2, 29));
    }
    let base = NaiveDate::from_ymd_opt(1899, 12, 31)?;
    let adjusted = if serial > 60 { serial - 1 } else { serial };
    let date = base.checked_add_signed(Duration::days(adjusted))?;
    Some((date.year(), date.month(), date.day()))
}

fn numeric_scalar(value: &FormulaValue) -> Result<f64, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => value
            .as_number()
            .ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn serial_scalar(value: &FormulaValue) -> Result<i64, FormulaValue> {
    Ok(numeric_scalar(value)?.floor() as i64)
}

fn string_scalar(value: &FormulaValue) -> Result<String, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        FormulaValue::String(s) => Ok(s.clone()),
        _ => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn ymd_from_serial_ctx(
    serial: i64,
    ctx: &EvaluationContext,
) -> Result<(i32, u32, u32), FormulaValue> {
    if ctx_date_1904(ctx) {
        let d = excel1904_date_from_serial(serial).ok_or(FormulaValue::Error(CellError::Num))?;
        Ok((d.year(), d.month(), d.day()))
    } else {
        excel1900_date_from_serial(serial).ok_or(FormulaValue::Error(CellError::Num))
    }
}

fn serial_from_ymd_ctx(
    year: i32,
    month: u32,
    day: u32,
    ctx: &EvaluationContext,
) -> Result<i64, FormulaValue> {
    if ctx_date_1904(ctx) {
        let d =
            NaiveDate::from_ymd_opt(year, month, day).ok_or(FormulaValue::Error(CellError::Num))?;
        Ok(excel1904_serial_from_date(d))
    } else {
        Ok(excel1900_serial_from_ymd(year, month, day as i32))
    }
}

fn days_in_month_gregorian(year: i32, month: u32) -> Option<u32> {
    let first = NaiveDate::from_ymd_opt(year, month, 1)?;
    let (ny, nm) = if month == 12 {
        (year + 1, 1)
    } else {
        (year, month + 1)
    };
    let next = NaiveDate::from_ymd_opt(ny, nm, 1)?;
    Some((next - first).num_days() as u32)
}

fn days_in_month_ctx(year: i32, month: u32, ctx: &EvaluationContext) -> Option<u32> {
    if ctx_date_1904(ctx) {
        days_in_month_gregorian(year, month)
    } else {
        Some(days_in_month_excel1900(year, month) as u32)
    }
}

fn ordinal_day_excel1900(year: i32, month: u32, day: u32) -> u32 {
    let mut total = 0u32;
    for m in 1..month {
        total += days_in_month_excel1900(year, m) as u32;
    }
    total + day
}

fn weekday_sun0_from_serial(serial: i64, ctx: &EvaluationContext) -> Result<u32, FormulaValue> {
    if ctx_date_1904(ctx) {
        let date = excel1904_date_from_serial(serial).ok_or(FormulaValue::Error(CellError::Num))?;
        Ok(date.weekday().num_days_from_sunday())
    } else {
        if serial == 60 {
            return Ok(4);
        }
        let base = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
        let adjusted = if serial > 60 { serial - 1 } else { serial };
        let date = base
            .checked_add_signed(Duration::days(adjusted))
            .ok_or(FormulaValue::Error(CellError::Num))?;
        Ok(date.weekday().num_days_from_sunday())
    }
}

fn parse_holidays(value: Option<&FormulaValue>) -> Result<HashSet<i64>, FormulaValue> {
    let mut out = HashSet::new();
    let Some(v) = value else {
        return Ok(out);
    };

    match v {
        FormulaValue::Error(e) => return Err(FormulaValue::Error(*e)),
        FormulaValue::Array(rows) => {
            for row in rows {
                for cell in row {
                    match cell {
                        FormulaValue::Error(e) => return Err(FormulaValue::Error(*e)),
                        FormulaValue::Empty => {}
                        _ => {
                            let n = numeric_scalar(cell)?;
                            out.insert(n.floor() as i64);
                        }
                    }
                }
            }
        }
        _ => {
            out.insert(numeric_scalar(v)?.floor() as i64);
        }
    }

    Ok(out)
}

fn is_weekend(serial: i64, ctx: &EvaluationContext) -> Result<bool, FormulaValue> {
    let dow = weekday_sun0_from_serial(serial, ctx)?;
    Ok(dow == 0 || dow == 6)
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

/// DATE(year, month, day)
pub fn fn_date(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Validate argument shapes
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
        if matches!(v, FormulaValue::Array(_)) {
            return Ok(FormulaValue::Error(CellError::Value));
        }
    }

    let mut year = to_i32_trunc(args.first().unwrap()).unwrap_or(0);
    let month = to_i32_trunc(args.get(1).unwrap()).unwrap_or(0);
    let day = to_i32_trunc(args.get(2).unwrap()).unwrap_or(0);

    // Excel: years 0..1899 are treated as 1900..3799
    if (0..1900).contains(&year) {
        year += 1900;
    }

    // Basic bounds (Excel supports 0..9999 in DATE)
    if !(0..=9999).contains(&year) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    // Normalize month overflow/underflow.
    // Use 0-based month index to handle negatives correctly.
    let total_months = (year as i64) * 12 + (month as i64 - 1);
    let norm_year = total_months.div_euclid(12) as i32;
    let norm_month0 = total_months.rem_euclid(12) as u32; // 0..11
    let norm_month = norm_month0 + 1;

    if ctx_date_1904(ctx) {
        // 1904 system: use chrono (no leap-year bug)
        let first = match NaiveDate::from_ymd_opt(norm_year, norm_month, 1) {
            Some(d) => d,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        let date = match first.checked_add_signed(Duration::days((day as i64) - 1)) {
            Some(d) => d,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        Ok(FormulaValue::Number(excel1904_serial_from_date(date) as f64))
    } else {
        // 1900 system with bug calendar: compute directly in the buggy serial system.
        Ok(FormulaValue::Number(
            excel1900_serial_from_ymd(norm_year, norm_month, day) as f64,
        ))
    }
}

/// YEAR(serial)
pub fn fn_year(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if let FormulaValue::Error(e) = v {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let n = match v.as_number() {
        Some(n) => n,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let serial = n.floor() as i64;

    if ctx_date_1904(ctx) {
        let date = match excel1904_date_from_serial(serial) {
            Some(d) => d,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        Ok(FormulaValue::Number(date.year() as f64))
    } else {
        let (y, _m, _d) = match excel1900_date_from_serial(serial) {
            Some(parts) => parts,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        Ok(FormulaValue::Number(y as f64))
    }
}

/// MONTH(serial)
pub fn fn_month(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if let FormulaValue::Error(e) = v {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let n = match v.as_number() {
        Some(n) => n,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let serial = n.floor() as i64;

    if ctx_date_1904(ctx) {
        let date = match excel1904_date_from_serial(serial) {
            Some(d) => d,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        Ok(FormulaValue::Number(date.month() as f64))
    } else {
        let (_y, m, _d) = match excel1900_date_from_serial(serial) {
            Some(parts) => parts,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        Ok(FormulaValue::Number(m as f64))
    }
}

/// DAY(serial)
pub fn fn_day(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if let FormulaValue::Error(e) = v {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let n = match v.as_number() {
        Some(n) => n,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let serial = n.floor() as i64;

    if ctx_date_1904(ctx) {
        let date = match excel1904_date_from_serial(serial) {
            Some(d) => d,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        Ok(FormulaValue::Number(date.day() as f64))
    } else {
        let (_y, _m, d) = match excel1900_date_from_serial(serial) {
            Some(parts) => parts,
            None => return Ok(FormulaValue::Error(CellError::Num)),
        };
        Ok(FormulaValue::Number(d as f64))
    }
}

pub fn fn_time(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let hour = numeric_scalar(args.first().unwrap());
    let minute = numeric_scalar(args.get(1).unwrap());
    let second = numeric_scalar(args.get(2).unwrap());

    let (hour, minute, second) = match (hour, minute, second) {
        (Ok(h), Ok(m), Ok(s)) => (h.trunc() as i64, m.trunc() as i64, s.trunc() as i64),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return Ok(e),
    };

    if hour < 0 || minute < 0 || second < 0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let total_seconds = hour * 3600 + minute * 60 + second;
    let seconds_of_day = total_seconds.rem_euclid(86400);
    Ok(FormulaValue::Number(seconds_of_day as f64 / 86400.0))
}

pub fn fn_hour(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match numeric_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if n < 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let frac = n.fract();
    let secs = ((frac * 86400.0).round() as i64).rem_euclid(86400);
    Ok(FormulaValue::Number((secs / 3600) as f64))
}

pub fn fn_minute(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match numeric_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if n < 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let frac = n.fract();
    let secs = ((frac * 86400.0).round() as i64).rem_euclid(86400);
    Ok(FormulaValue::Number(((secs % 3600) / 60) as f64))
}

pub fn fn_second(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match numeric_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if n < 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let frac = n.fract();
    let secs = ((frac * 86400.0).round() as i64).rem_euclid(86400);
    Ok(FormulaValue::Number((secs % 60) as f64))
}

pub fn fn_weekday(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let serial = match serial_scalar(args.first().unwrap()) {
        Ok(s) => s,
        Err(e) => return Ok(e),
    };

    let return_type = match args.get(1) {
        Some(v) => match numeric_scalar(v) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return Ok(e),
        },
        None => 1,
    };

    let dow_sun1 = match weekday_sun0_from_serial(serial, ctx) {
        Ok(d) => d + 1,
        Err(e) => return Ok(e),
    };

    let out = match return_type {
        1 => dow_sun1 as f64,
        2 => (((dow_sun1 + 5) % 7) + 1) as f64,
        3 => ((dow_sun1 + 5) % 7) as f64,
        11..=17 => {
            let start = ((return_type - 10) % 7) as u32;
            let dow_sun0 = dow_sun1 - 1;
            (((dow_sun0 + 7 - start) % 7) + 1) as f64
        }
        _ => return Ok(FormulaValue::Error(CellError::Num)),
    };

    Ok(FormulaValue::Number(out))
}

pub fn fn_weeknum(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let serial = match serial_scalar(args.first().unwrap()) {
        Ok(s) => s,
        Err(e) => return Ok(e),
    };

    let return_type = match args.get(1) {
        Some(v) => match numeric_scalar(v) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return Ok(e),
        },
        None => 1,
    };

    let (year, month, day) = match ymd_from_serial_ctx(serial, ctx) {
        Ok(parts) => parts,
        Err(e) => return Ok(e),
    };

    let week_start_sun0 = match return_type {
        1 => 0,
        2 => 1,
        _ => return Ok(FormulaValue::Error(CellError::Num)),
    };

    let ordinal = if ctx_date_1904(ctx) {
        let d = excel1904_date_from_serial(serial).ok_or(FormulaValue::Error(CellError::Num));
        match d {
            Ok(d) => d.ordinal(),
            Err(e) => return Ok(e),
        }
    } else {
        ordinal_day_excel1900(year, month, day)
    } as i32;

    let jan1_serial = if ctx_date_1904(ctx) {
        match serial_from_ymd_ctx(year, 1, 1, ctx) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        }
    } else {
        excel1900_serial_from_ymd(year, 1, 1)
    };
    let jan1_dow_sun0 = match weekday_sun0_from_serial(jan1_serial, ctx) {
        Ok(v) => v as i32,
        Err(e) => return Ok(e),
    };

    let offset = (7 + jan1_dow_sun0 - week_start_sun0) % 7;
    let week = ((ordinal + offset - 1) / 7) + 1;
    Ok(FormulaValue::Number(week as f64))
}

pub fn fn_isoweeknum(
    args: &[FormulaValue],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let serial = match serial_scalar(args.first().unwrap()) {
        Ok(s) => s,
        Err(e) => return Ok(e),
    };

    let date = if ctx_date_1904(ctx) {
        excel1904_date_from_serial(serial)
    } else {
        let base = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
        let adjusted = if serial > 60 { serial - 1 } else { serial };
        base.checked_add_signed(Duration::days(adjusted))
    };

    match date {
        Some(d) => Ok(FormulaValue::Number(d.iso_week().week() as f64)),
        None => Ok(FormulaValue::Error(CellError::Num)),
    }
}

pub fn fn_edate(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let start_serial = match serial_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let months = match numeric_scalar(args.get(1).unwrap()) {
        Ok(v) => v.trunc() as i64,
        Err(e) => return Ok(e),
    };

    let (year, month, day) = match ymd_from_serial_ctx(start_serial, ctx) {
        Ok(parts) => parts,
        Err(e) => return Ok(e),
    };

    let total_months = (year as i64) * 12 + (month as i64 - 1) + months;
    let new_year = total_months.div_euclid(12) as i32;
    let new_month = (total_months.rem_euclid(12) + 1) as u32;
    let dim = match days_in_month_ctx(new_year, new_month, ctx) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };
    let new_day = day.min(dim);

    match serial_from_ymd_ctx(new_year, new_month, new_day, ctx) {
        Ok(s) => Ok(FormulaValue::Number(s as f64)),
        Err(e) => Ok(e),
    }
}

pub fn fn_eomonth(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let start_serial = match serial_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let months = match numeric_scalar(args.get(1).unwrap()) {
        Ok(v) => v.trunc() as i64,
        Err(e) => return Ok(e),
    };

    let (year, month, _day) = match ymd_from_serial_ctx(start_serial, ctx) {
        Ok(parts) => parts,
        Err(e) => return Ok(e),
    };

    let total_months = (year as i64) * 12 + (month as i64 - 1) + months;
    let new_year = total_months.div_euclid(12) as i32;
    let new_month = (total_months.rem_euclid(12) + 1) as u32;
    let dim = match days_in_month_ctx(new_year, new_month, ctx) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    match serial_from_ymd_ctx(new_year, new_month, dim, ctx) {
        Ok(s) => Ok(FormulaValue::Number(s as f64)),
        Err(e) => Ok(e),
    }
}

pub fn fn_days(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let end_date = match serial_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let start_date = match serial_scalar(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number((end_date - start_date) as f64))
}

pub fn fn_days360(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let start_serial = match serial_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let end_serial = match serial_scalar(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let european = match args.get(2) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Array(_)) => return Ok(FormulaValue::Error(CellError::Value)),
        Some(v) => v.as_bool().ok_or(FormulaValue::Error(CellError::Value)),
        None => Ok(false),
    };
    let european = match european {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let (sy, sm, sd) = match ymd_from_serial_ctx(start_serial, ctx) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (ey, em, ed) = match ymd_from_serial_ctx(end_serial, ctx) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let days = if european {
        days360_eu(sy, sm, sd, ey, em, ed)
    } else {
        days360_us(sy, sm, sd, ey, em, ed)
    };

    Ok(FormulaValue::Number(days as f64))
}

pub fn fn_datedif(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let start_serial = match serial_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let end_serial = match serial_scalar(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if end_serial < start_serial {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let unit = match string_scalar(args.get(2).unwrap()) {
        Ok(s) => s.to_uppercase(),
        Err(e) => return Ok(e),
    };

    let (sy, sm, sd) = match ymd_from_serial_ctx(start_serial, ctx) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (ey, em, ed) = match ymd_from_serial_ctx(end_serial, ctx) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let out = match unit.as_str() {
        "D" => (end_serial - start_serial) as f64,
        "Y" => {
            let mut years = ey - sy;
            if (em, ed) < (sm, sd) {
                years -= 1;
            }
            years as f64
        }
        "M" => {
            let mut months = (ey - sy) * 12 + em as i32 - sm as i32;
            if ed < sd {
                months -= 1;
            }
            months as f64
        }
        "MD" => {
            if ed >= sd {
                (ed - sd) as f64
            } else {
                let dim = match days_in_month_ctx(sy, sm, ctx) {
                    Some(v) => v,
                    None => return Ok(FormulaValue::Error(CellError::Num)),
                };
                (ed + dim - sd) as f64
            }
        }
        "YM" => {
            let mut months = em as i32 - sm as i32;
            if months < 0 {
                months += 12;
            }
            if ed < sd {
                months = (months + 11) % 12;
            }
            months as f64
        }
        "YD" => {
            let dim_this_year = match days_in_month_ctx(ey, sm, ctx) {
                Some(v) => v,
                None => return Ok(FormulaValue::Error(CellError::Num)),
            };
            let day_this_year = sd.min(dim_this_year);
            let candidate_this_year = match serial_from_ymd_ctx(ey, sm, day_this_year, ctx) {
                Ok(v) => v,
                Err(e) => return Ok(e),
            };

            let anchor = if candidate_this_year <= end_serial {
                candidate_this_year
            } else {
                let dim_prev_year = match days_in_month_ctx(ey - 1, sm, ctx) {
                    Some(v) => v,
                    None => return Ok(FormulaValue::Error(CellError::Num)),
                };
                let day_prev_year = sd.min(dim_prev_year);
                match serial_from_ymd_ctx(ey - 1, sm, day_prev_year, ctx) {
                    Ok(v) => v,
                    Err(e) => return Ok(e),
                }
            };

            (end_serial - anchor) as f64
        }
        _ => return Ok(FormulaValue::Error(CellError::Num)),
    };

    Ok(FormulaValue::Number(out))
}

pub fn fn_yearfrac(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let start = match serial_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let end = match serial_scalar(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let basis = match args.get(2) {
        Some(v) => match numeric_scalar(v) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return Ok(e),
        },
        None => 0,
    };

    if !(0..=4).contains(&basis) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let (sign, s, e) = if end >= start {
        (1.0, start, end)
    } else {
        (-1.0, end, start)
    };

    let (sy, sm, sd) = match ymd_from_serial_ctx(s, ctx) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (ey, em, ed) = match ymd_from_serial_ctx(e, ctx) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let frac = match basis {
        0 => days360_us(sy, sm, sd, ey, em, ed) as f64 / 360.0,
        4 => days360_eu(sy, sm, sd, ey, em, ed) as f64 / 360.0,
        2 => (e - s) as f64 / 360.0,
        3 => (e - s) as f64 / 365.0,
        1 => {
            if sy == ey {
                let denom = if ctx_date_1904(ctx) {
                    if is_leap_gregorian(sy) {
                        366.0
                    } else {
                        365.0
                    }
                } else {
                    days_in_year_excel1900(sy) as f64
                };
                (e - s) as f64 / denom
            } else {
                let next_jan1 = match serial_from_ymd_ctx(sy + 1, 1, 1, ctx) {
                    Ok(v) => v,
                    Err(e) => return Ok(e),
                };
                let first_year_days = if ctx_date_1904(ctx) {
                    if is_leap_gregorian(sy) {
                        366.0
                    } else {
                        365.0
                    }
                } else {
                    days_in_year_excel1900(sy) as f64
                };
                let mut total = (next_jan1 - s) as f64 / first_year_days;

                for y in (sy + 1)..ey {
                    total += 1.0;
                    if !ctx_date_1904(ctx) && y == 1900 {
                        total += 1.0 / days_in_year_excel1900(1900) as f64;
                    }
                }

                let this_jan1 = match serial_from_ymd_ctx(ey, 1, 1, ctx) {
                    Ok(v) => v,
                    Err(e) => return Ok(e),
                };
                let last_year_days = if ctx_date_1904(ctx) {
                    if is_leap_gregorian(ey) {
                        366.0
                    } else {
                        365.0
                    }
                } else {
                    days_in_year_excel1900(ey) as f64
                };
                total + (e - this_jan1) as f64 / last_year_days
            }
        }
        _ => 0.0,
    };

    Ok(FormulaValue::Number(sign * frac))
}

pub fn fn_datevalue(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match string_scalar(args.first().unwrap()) {
        Ok(s) => s,
        Err(e) => return Ok(e),
    };
    let text = text.trim();

    let parsed = NaiveDate::parse_from_str(text, "%m/%d/%Y")
        .or_else(|_| NaiveDate::parse_from_str(text, "%Y-%m-%d"))
        .or_else(|_| NaiveDate::parse_from_str(text, "%Y/%m/%d"))
        .or_else(|_| NaiveDate::parse_from_str(text, "%d-%b-%y"))
        .or_else(|_| NaiveDate::parse_from_str(text, "%d-%b-%Y"));

    let date = match parsed {
        Ok(d) => d,
        Err(_) => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let serial = if ctx_date_1904(ctx) {
        excel1904_serial_from_date(date)
    } else {
        excel1900_serial_from_ymd(date.year(), date.month(), date.day() as i32)
    };
    Ok(FormulaValue::Number(serial as f64))
}

pub fn fn_timevalue(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let text = match string_scalar(args.first().unwrap()) {
        Ok(s) => s,
        Err(e) => return Ok(e),
    };
    let t = text.trim();
    let upper = t.to_uppercase();

    let parsed = chrono::NaiveTime::parse_from_str(t, "%H:%M:%S")
        .or_else(|_| chrono::NaiveTime::parse_from_str(t, "%H:%M"))
        .or_else(|_| chrono::NaiveTime::parse_from_str(&upper, "%I:%M:%S %p"))
        .or_else(|_| chrono::NaiveTime::parse_from_str(&upper, "%I:%M %p"))
        .or_else(|_| {
            // "Date information in time_text is ignored" — try datetime formats
            // and extract just the time portion.
            chrono::NaiveDateTime::parse_from_str(&upper, "%d-%b-%Y %I:%M %p")
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(&upper, "%d-%b-%Y %I:%M:%S %p"))
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(&upper, "%d-%b-%Y %H:%M:%S"))
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(&upper, "%d-%b-%Y %H:%M"))
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(t, "%m/%d/%Y %H:%M:%S"))
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(t, "%m/%d/%Y %H:%M"))
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(t, "%Y/%m/%d %H:%M:%S"))
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(t, "%Y/%m/%d %H:%M"))
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(t, "%Y-%m-%d %H:%M:%S"))
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(t, "%Y-%m-%d %H:%M"))
                .map(|dt| dt.time())
        });

    let tm = match parsed {
        Ok(v) => v,
        Err(_) => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let total_seconds = tm.num_seconds_from_midnight() as f64;
    Ok(FormulaValue::Number(total_seconds / 86400.0))
}

pub fn fn_networkdays(
    args: &[FormulaValue],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let start = match serial_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let end = match serial_scalar(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let holidays = match parse_holidays(args.get(2)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let (sign, s, e) = if end >= start {
        (1.0, start, end)
    } else {
        (-1.0, end, start)
    };

    let mut count = 0i64;
    for serial in s..=e {
        let weekend = match is_weekend(serial, ctx) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        if !weekend && !holidays.contains(&serial) {
            count += 1;
        }
    }

    Ok(FormulaValue::Number(sign * count as f64))
}

pub fn fn_workday(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let start = match serial_scalar(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let days = match numeric_scalar(args.get(1).unwrap()) {
        Ok(v) => v.trunc() as i64,
        Err(e) => return Ok(e),
    };
    let holidays = match parse_holidays(args.get(2)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if days == 0 {
        return Ok(FormulaValue::Number(start as f64));
    }

    let step = if days > 0 { 1 } else { -1 };
    let mut remaining = days.abs();
    let mut current = start;

    while remaining > 0 {
        current += step;
        let weekend = match is_weekend(current, ctx) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        if !weekend && !holidays.contains(&current) {
            remaining -= 1;
        }
    }

    Ok(FormulaValue::Number(current as f64))
}

/// NOW() - Returns current date and time as Excel serial number
/// This is a volatile function that recalculates on every calculation.
pub fn fn_now(_args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    use chrono::{Datelike, Local, Timelike};

    let now = Local::now();
    let year = now.year();
    let month = now.month();
    let day = now.day();

    // Get date serial
    let date_serial = if ctx_date_1904(ctx) {
        let date = NaiveDate::from_ymd_opt(year, month, day).unwrap();
        excel1904_serial_from_date(date)
    } else {
        excel1900_serial_from_ymd(year, month, day as i32)
    };

    // Add time as fraction of day
    let time_fraction =
        (now.hour() as f64 * 3600.0 + now.minute() as f64 * 60.0 + now.second() as f64) / 86400.0;

    Ok(FormulaValue::Number(date_serial as f64 + time_fraction))
}

/// TODAY() - Returns current date as Excel serial number
/// This is a volatile function that recalculates on every calculation.
pub fn fn_today(_args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    use chrono::{Datelike, Local};

    let today = Local::now();
    let year = today.year();
    let month = today.month();
    let day = today.day();

    let serial = if ctx_date_1904(ctx) {
        let date = NaiveDate::from_ymd_opt(year, month, day).unwrap();
        excel1904_serial_from_date(date)
    } else {
        excel1900_serial_from_ymd(year, month, day as i32)
    };

    Ok(FormulaValue::Number(serial as f64))
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
    fn test_time() {
        assert_close(number(eval("=TIME(1,30,0)").unwrap()), 0.0625);
        assert_close(number(eval("=TIME(25,0,0)").unwrap()), 1.0 / 24.0);
        assert_eq!(
            eval("=TIME(-1,0,0)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_hour() {
        assert_eq!(
            eval("=HOUR(TIME(14,30,15))").unwrap(),
            FormulaValue::Number(14.0)
        );
        assert_eq!(eval("=HOUR(1.5)").unwrap(), FormulaValue::Number(12.0));
        assert_eq!(
            eval("=HOUR(-1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_hour_docs() {
        // https://support.microsoft.com/en-us/office/hour-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7

        // =HOUR(A2) where A2=0.75 => 18 (75% of 24 hours)
        assert_eq!(eval("=HOUR(0.75)").unwrap(), FormulaValue::Number(18.0));

        // =HOUR(A3) where A3="7/18/2011 7:45" => 7 (hour portion of date/time)
        assert_eq!(
            eval("=HOUR(DATE(2011,7,18)+TIME(7,45,0))").unwrap(),
            FormulaValue::Number(7.0)
        );

        // =HOUR(A4) where A4="4/21/2012" => 0 (date with no time is 12:00 AM)
        assert_eq!(
            eval("=HOUR(DATE(2012,4,21))").unwrap(),
            FormulaValue::Number(0.0)
        );

        // Syntax: text string via TIMEVALUE("6:45 PM") => hour 18
        assert_eq!(
            eval("=HOUR(TIMEVALUE(\"6:45 PM\"))").unwrap(),
            FormulaValue::Number(18.0)
        );

        // Syntax: 0.78125 represents 6:45 PM => hour 18
        assert_eq!(eval("=HOUR(0.78125)").unwrap(), FormulaValue::Number(18.0));
    }

    #[test]
    fn test_minute() {
        assert_eq!(
            eval("=MINUTE(TIME(14,30,15))").unwrap(),
            FormulaValue::Number(30.0)
        );
        assert_eq!(
            eval("=MINUTE(TIME(0,59,59))").unwrap(),
            FormulaValue::Number(59.0)
        );
        assert_eq!(
            eval("=MINUTE(-0.1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_minute_docs() {
        // https://support.microsoft.com/en-us/office/minute-function-af728df0-05c4-4b07-9eed-a84801a60589

        // =MINUTE(A2) where A2="12:45:00 PM" => 45
        assert_eq!(
            eval("=MINUTE(TIME(12,45,0))").unwrap(),
            FormulaValue::Number(45.0)
        );

        // Syntax: text string via TIMEVALUE("6:45 PM") => minute 45
        assert_eq!(
            eval("=MINUTE(TIMEVALUE(\"6:45 PM\"))").unwrap(),
            FormulaValue::Number(45.0)
        );

        // Syntax: 0.78125 represents 6:45 PM => minute 45
        assert_eq!(
            eval("=MINUTE(0.78125)").unwrap(),
            FormulaValue::Number(45.0)
        );
    }

    #[test]
    fn test_second() {
        assert_eq!(
            eval("=SECOND(TIME(14,30,15))").unwrap(),
            FormulaValue::Number(15.0)
        );
        assert_eq!(
            eval("=SECOND(TIME(0,0,59))").unwrap(),
            FormulaValue::Number(59.0)
        );
        assert_eq!(
            eval("=SECOND(-0.5)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_weekday() {
        assert_eq!(
            eval("=WEEKDAY(DATE(2024,1,7))").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=WEEKDAY(DATE(2024,1,7),2)").unwrap(),
            FormulaValue::Number(7.0)
        );
        assert_eq!(
            eval("=WEEKDAY(DATE(2024,1,7),3)").unwrap(),
            FormulaValue::Number(6.0)
        );
    }

    #[test]
    fn test_weekday_docs() {
        // Microsoft docs: WEEKDAY function
        // https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a
        // Date used in examples: 2/14/2008 (a Thursday)

        // =WEEKDAY(A2) where A2=DATE(2008,2,14) => 5
        // return_type 1 (default): 1=Sunday..7=Saturday, Thursday=5
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14))").unwrap(),
            FormulaValue::Number(5.0)
        );

        // =WEEKDAY(A2, 2) => 4
        // return_type 2: 1=Monday..7=Sunday, Thursday=4
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),2)").unwrap(),
            FormulaValue::Number(4.0)
        );

        // =WEEKDAY(A2, 3) => 3
        // return_type 3: 0=Monday..6=Sunday, Thursday=3
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),3)").unwrap(),
            FormulaValue::Number(3.0)
        );

        // return_type 11: 1=Monday..7=Sunday, Thursday=4
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),11)").unwrap(),
            FormulaValue::Number(4.0)
        );

        // return_type 12: 1=Tuesday..7=Monday, Thursday=3
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),12)").unwrap(),
            FormulaValue::Number(3.0)
        );

        // return_type 13: 1=Wednesday..7=Tuesday, Thursday=2
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),13)").unwrap(),
            FormulaValue::Number(2.0)
        );

        // return_type 14: 1=Thursday..7=Wednesday, Thursday=1
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),14)").unwrap(),
            FormulaValue::Number(1.0)
        );

        // return_type 15: 1=Friday..7=Thursday, Thursday=7
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),15)").unwrap(),
            FormulaValue::Number(7.0)
        );

        // return_type 16: 1=Saturday..7=Friday, Thursday=6
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),16)").unwrap(),
            FormulaValue::Number(6.0)
        );

        // return_type 17: 1=Sunday..7=Saturday, Thursday=5
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),17)").unwrap(),
            FormulaValue::Number(5.0)
        );

        // Error cases from docs:
        // "If return_type is out of the range specified in the table above, a #NUM! error is returned."
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),0)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            eval("=WEEKDAY(DATE(2008,2,14),4)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_weeknum() {
        assert_eq!(
            eval("=WEEKNUM(DATE(2024,1,1),1)").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=WEEKNUM(DATE(2024,1,1),2)").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=WEEKNUM(DATE(2024,12,31),1)").unwrap(),
            FormulaValue::Number(53.0)
        );
    }

    #[test]
    fn test_weeknum_docs() {
        // Microsoft docs: WEEKNUM function
        // https://support.microsoft.com/en-us/office/weeknum-function-e5c43a03-b4ab-426c-b411-b18c13c75340
        // Date used in examples: 3/9/2012 (March 9, 2012, a Friday)

        // =WEEKNUM(A2) where A2=DATE(2012,3,9) => 10
        // return_type 1 (default): week begins on Sunday
        assert_eq!(
            eval("=WEEKNUM(DATE(2012,3,9))").unwrap(),
            FormulaValue::Number(10.0)
        );

        // =WEEKNUM(A2,1) => 10 (explicit return_type 1, same as default)
        assert_eq!(
            eval("=WEEKNUM(DATE(2012,3,9),1)").unwrap(),
            FormulaValue::Number(10.0)
        );

        // =WEEKNUM(A2,2) => 11
        // return_type 2: week begins on Monday
        assert_eq!(
            eval("=WEEKNUM(DATE(2012,3,9),2)").unwrap(),
            FormulaValue::Number(11.0)
        );

        // Error cases from docs:
        // "If Return_type is out of the range specified in the table above, a #NUM! error is returned."
        assert_eq!(
            eval("=WEEKNUM(DATE(2012,3,9),0)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            eval("=WEEKNUM(DATE(2012,3,9),3)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_isoweeknum() {
        assert_eq!(
            eval("=ISOWEEKNUM(DATE(2024,1,1))").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=ISOWEEKNUM(DATE(2024,12,31))").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=ISOWEEKNUM(DATE(2023,1,1))").unwrap(),
            FormulaValue::Number(52.0)
        );
    }

    #[test]
    fn test_edate() {
        assert_eq!(
            eval("=EDATE(DATE(2024,1,31),1)").unwrap(),
            FormulaValue::Number(45351.0)
        );
        assert_eq!(
            eval("=EDATE(DATE(2024,3,31),-1)").unwrap(),
            FormulaValue::Number(45351.0)
        );
        assert_eq!(
            eval("=EDATE(DATE(2024,1,15),12)").unwrap(),
            FormulaValue::Number(45672.0)
        );
    }

    #[test]
    fn test_edate_docs() {
        // https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5
        // Start date: 15-Jan-11 (January 15, 2011)

        // =EDATE(A2,1) => 15-Feb-11 (one month after)
        assert_eq!(
            eval("=EDATE(DATE(2011,1,15),1)").unwrap(),
            FormulaValue::Number(40589.0)
        );

        // =EDATE(A2,-1) => 15-Dec-10 (one month before)
        assert_eq!(
            eval("=EDATE(DATE(2011,1,15),-1)").unwrap(),
            FormulaValue::Number(40527.0)
        );

        // =EDATE(A2,2) => 15-Mar-11 (two months after)
        assert_eq!(
            eval("=EDATE(DATE(2011,1,15),2)").unwrap(),
            FormulaValue::Number(40617.0)
        );
    }

    #[test]
    fn test_eomonth() {
        assert_eq!(
            eval("=EOMONTH(DATE(2024,1,15),0)").unwrap(),
            FormulaValue::Number(45322.0)
        );
        assert_eq!(
            eval("=EOMONTH(DATE(2024,1,15),1)").unwrap(),
            FormulaValue::Number(45351.0)
        );
        assert_eq!(
            eval("=EOMONTH(DATE(2024,1,15),-1)").unwrap(),
            FormulaValue::Number(45291.0)
        );
    }

    #[test]
    fn test_eomonth_docs() {
        // https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628
        // Start date: 1-Jan-11

        // 0 months: end of same month => 1/31/2011
        assert_eq!(
            eval("=EOMONTH(DATE(2011,1,1),0)").unwrap(),
            FormulaValue::Number(40574.0)
        );

        // =EOMONTH(A2,1) => 2/28/2011
        // "Date of the last day of the month, one month after the date in A2."
        assert_eq!(
            eval("=EOMONTH(DATE(2011,1,1),1)").unwrap(),
            FormulaValue::Number(40602.0)
        );

        // =EOMONTH(A2,-3) => 10/31/2010
        // "Date of the last day of the month, three months before the date in A2."
        assert_eq!(
            eval("=EOMONTH(DATE(2011,1,1),-3)").unwrap(),
            FormulaValue::Number(40482.0)
        );
    }

    #[test]
    fn test_days() {
        assert_eq!(
            eval("=DAYS(DATE(2024,1,10),DATE(2024,1,1))").unwrap(),
            FormulaValue::Number(9.0)
        );
        assert_eq!(
            eval("=DAYS(DATE(2024,1,1),DATE(2024,1,10))").unwrap(),
            FormulaValue::Number(-9.0)
        );
        assert_eq!(eval("=DAYS(10.9,1.1)").unwrap(), FormulaValue::Number(9.0));
    }

    #[test]
    fn test_days_docs() {
        // https://support.microsoft.com/en-us/office/days-function-57740535-d549-4395-8728-0f07bff0b9df

        // =DAYS("15-MAR-2021","1-FEB-2021") => 42
        assert_eq!(
            eval("=DAYS(DATE(2021,3,15),DATE(2021,2,1))").unwrap(),
            FormulaValue::Number(42.0)
        );

        // =DAYS(A2,A3) where A2=31-DEC-2021, A3=1-JAN-2021 => 364
        assert_eq!(
            eval("=DAYS(DATE(2021,12,31),DATE(2021,1,1))").unwrap(),
            FormulaValue::Number(364.0)
        );
    }

    #[test]
    fn test_days360() {
        assert_eq!(
            eval("=DAYS360(DATE(2024,1,31),DATE(2024,2,29),FALSE)").unwrap(),
            FormulaValue::Number(29.0)
        );
        assert_eq!(
            eval("=DAYS360(DATE(2024,1,31),DATE(2024,2,29),TRUE)").unwrap(),
            FormulaValue::Number(29.0)
        );
        assert_eq!(
            eval("=DAYS360(DATE(2024,1,31),DATE(2024,3,31),FALSE)").unwrap(),
            FormulaValue::Number(60.0)
        );
    }

    #[test]
    fn test_datedif() {
        assert_eq!(
            eval("=DATEDIF(DATE(2020,1,15),DATE(2024,3,10),\"Y\")").unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            eval("=DATEDIF(DATE(2020,1,15),DATE(2024,3,10),\"M\")").unwrap(),
            FormulaValue::Number(49.0)
        );
        assert_eq!(
            eval("=DATEDIF(DATE(2024,1,31),DATE(2024,3,2),\"MD\")").unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            eval("=DATEDIF(DATE(2024,1,31),DATE(2024,3,2),\"YM\")").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=DATEDIF(DATE(2024,1,31),DATE(2024,3,2),\"YD\")").unwrap(),
            FormulaValue::Number(31.0)
        );
        assert_eq!(
            eval("=DATEDIF(DATE(2024,3,2),DATE(2024,1,31),\"D\")").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    /// Microsoft docs examples for DATEDIF:
    /// https://support.microsoft.com/en-us/office/datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c
    #[test]
    fn test_datedif_docs() {
        // Example 1: "Y" – two complete years in the period (2)
        assert_eq!(
            eval("=DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"Y\")").unwrap(),
            FormulaValue::Number(2.0)
        );

        // Example 2: "D" – 440 days between June 1, 2001 and August 15, 2002 (440)
        assert_eq!(
            eval("=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"D\")").unwrap(),
            FormulaValue::Number(440.0)
        );

        // Example 3: "YD" – 75 days between June 1 and August 15, ignoring years (75)
        assert_eq!(
            eval("=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YD\")").unwrap(),
            FormulaValue::Number(75.0)
        );

        // "M" – complete months for example-1 dates: 1/1/2001 to 1/1/2003 = 24
        assert_eq!(
            eval("=DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"M\")").unwrap(),
            FormulaValue::Number(24.0)
        );

        // "YM" – month diff ignoring days/years for example-2 dates: Jun to Aug = 2
        assert_eq!(
            eval("=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YM\")").unwrap(),
            FormulaValue::Number(2.0)
        );

        // "MD" – day diff ignoring months/years for example-2 dates: 15 - 1 = 14
        assert_eq!(
            eval("=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"MD\")").unwrap(),
            FormulaValue::Number(14.0)
        );

        // Remarks: start_date > end_date → #NUM!
        assert_eq!(
            eval("=DATEDIF(DATE(2003,1,1),DATE(2001,1,1),\"Y\")").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_yearfrac() {
        assert_close(
            number(eval("=YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),1)").unwrap()),
            1.0,
        );
        assert_close(
            number(eval("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),2)").unwrap()),
            182.0 / 360.0,
        );
        assert_close(
            number(eval("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)").unwrap()),
            182.0 / 365.0,
        );
    }

    #[test]
    fn test_yearfrac_docs() {
        // Microsoft docs: YEARFRAC function
        // https://support.microsoft.com/en-us/office/yearfrac-function-3844141e-c76d-4143-82b6-208454ddc6a8

        // --- Example table (start_date=1/1/2012, end_date=7/30/2012) ---

        // =YEARFRAC(A2,A3) → 0.58055556 (basis omitted, defaults to 0: US 30/360)
        // 30/360 days: (7-1)*30 + (30-1) = 209, 209/360 = 0.58055556
        assert_close(
            number(eval("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))").unwrap()),
            209.0 / 360.0,
        );

        // =YEARFRAC(A2,A3,1) → 0.57650273 (basis 1: Actual/actual, 366 day leap year)
        // Actual days: 211, year days: 366 (2012 is leap), 211/366 = 0.57650273
        assert_close(
            number(eval("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)").unwrap()),
            211.0 / 366.0,
        );

        // =YEARFRAC(A2,A3,3) → 0.57808219 (basis 3: Actual/365)
        // Actual days: 211, 211/365 = 0.57808219
        assert_close(
            number(eval("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)").unwrap()),
            211.0 / 365.0,
        );
    }

    #[test]
    fn test_datevalue() {
        assert_eq!(
            eval("=DATEVALUE(\"1/15/2024\")").unwrap(),
            FormulaValue::Number(45306.0)
        );
        assert_eq!(
            eval("=DATEVALUE(\"2024-01-15\")").unwrap(),
            FormulaValue::Number(45306.0)
        );
        assert_eq!(
            eval("=DATEVALUE(\"15-Jan-24\")").unwrap(),
            FormulaValue::Number(45306.0)
        );
    }

    #[test]
    fn test_datevalue_docs() {
        // Microsoft docs: DATEVALUE function
        // https://support.microsoft.com/en-us/office/datevalue-function-df8b07d4-7761-4a93-bc33-b7471bbff252

        // --- Example table ---

        // =DATEVALUE("8/22/2011") → 40777
        assert_eq!(
            eval("=DATEVALUE(\"8/22/2011\")").unwrap(),
            FormulaValue::Number(40777.0)
        );

        // =DATEVALUE("22-MAY-2011") → 40685
        assert_eq!(
            eval("=DATEVALUE(\"22-MAY-2011\")").unwrap(),
            FormulaValue::Number(40685.0)
        );

        // =DATEVALUE("2011/02/23") → 40597
        assert_eq!(
            eval("=DATEVALUE(\"2011/02/23\")").unwrap(),
            FormulaValue::Number(40597.0)
        );

        // =DATEVALUE(A2 & "/" & A3 & "/" & A4) where A2=11, A3=3, A4=2011
        // Equivalent to =DATEVALUE("11/3/2011") → 40850
        assert_eq!(
            eval("=DATEVALUE(\"11/3/2011\")").unwrap(),
            FormulaValue::Number(40850.0)
        );

        // --- Description examples ---

        // =DATEVALUE("1/1/2008") returns 39448
        assert_eq!(
            eval("=DATEVALUE(\"1/1/2008\")").unwrap(),
            FormulaValue::Number(39448.0)
        );

        // --- Syntax section: mentioned valid formats ---

        // "1/30/2008" – m/d/Y format
        assert_eq!(
            eval("=DATEVALUE(\"1/30/2008\")").unwrap(),
            FormulaValue::Number(39477.0)
        );

        // "30-Jan-2008" – d-MMM-YYYY format (same date as above)
        assert_eq!(
            eval("=DATEVALUE(\"30-Jan-2008\")").unwrap(),
            FormulaValue::Number(39477.0)
        );

        // --- Remarks section ---

        // "January 1, 1900 is serial number 1"
        assert_eq!(
            eval("=DATEVALUE(\"1/1/1900\")").unwrap(),
            FormulaValue::Number(1.0)
        );

        // "January 1, 2008 is serial number 39448"
        assert_eq!(
            eval("=DATEVALUE(\"1/1/2008\")").unwrap(),
            FormulaValue::Number(39448.0)
        );

        // --- Error cases (docs: returns #VALUE! for invalid date_text) ---

        assert_eq!(
            eval("=DATEVALUE(\"not a date\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        assert_eq!(
            eval("=DATEVALUE(\"\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // Numeric argument → #VALUE! (DATEVALUE expects text)
        assert_eq!(
            eval("=DATEVALUE(12345)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // --- Year-dependent example (skipped: "5-JUL" depends on current year,
        //     docs serial 39634 is inconsistent with stated year 2011) ---
    }

    #[test]
    fn test_timevalue() {
        assert_close(
            number(eval("=TIMEVALUE(\"14:30:00\")").unwrap()),
            14.5 / 24.0,
        );
        assert_close(number(eval("=TIMEVALUE(\"14:30\")").unwrap()), 14.5 / 24.0);
        assert_close(
            number(eval("=TIMEVALUE(\"2:30 PM\")").unwrap()),
            14.5 / 24.0,
        );
    }

    #[test]
    fn test_networkdays() {
        assert_eq!(
            eval("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10))").unwrap(),
            FormulaValue::Number(8.0)
        );
        assert_eq!(
            eval("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10),{DATE(2024,1,1),DATE(2024,1,8)})")
                .unwrap(),
            FormulaValue::Number(6.0)
        );
        assert_eq!(
            eval("=NETWORKDAYS(DATE(2024,1,10),DATE(2024,1,1))").unwrap(),
            FormulaValue::Number(-8.0)
        );
    }

    #[test]
    fn test_networkdays_docs() {
        // https://support.microsoft.com/en-us/office/networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7
        // Data: start=10/1/2012, end=3/1/2013
        // Holidays: 11/22/2012, 12/4/2012, 1/21/2013

        // =NETWORKDAYS(A2,A3) => 110
        // Number of workdays between 10/1/2012 and 3/1/2013
        assert_eq!(
            eval("=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))").unwrap(),
            FormulaValue::Number(110.0)
        );

        // =NETWORKDAYS(A2,A3,A4) => 109
        // With 11/22/2012 holiday as a non-working day
        assert_eq!(
            eval("=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),DATE(2012,11,22))").unwrap(),
            FormulaValue::Number(109.0)
        );

        // =NETWORKDAYS(A2,A3,A4:A6) => 107
        // With three holidays as non-working days
        assert_eq!(
            eval("=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),{DATE(2012,11,22),DATE(2012,12,4),DATE(2013,1,21)})")
                .unwrap(),
            FormulaValue::Number(107.0)
        );

        // Remarks: invalid date returns #VALUE!
        assert_eq!(
            eval("=NETWORKDAYS(\"not a date\",DATE(2013,3,1))").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_workday() {
        assert_eq!(
            eval("=WORKDAY(DATE(2024,1,1),5)").unwrap(),
            FormulaValue::Number(45299.0)
        );
        assert_eq!(
            eval("=WORKDAY(DATE(2024,1,1),5,{DATE(2024,1,3)})").unwrap(),
            FormulaValue::Number(45300.0)
        );
        assert_eq!(
            eval("=WORKDAY(DATE(2024,1,8),-5)").unwrap(),
            FormulaValue::Number(45292.0)
        );
    }

    #[test]
    fn test_workday_docs() {
        // https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33
        // Data: start=10/1/2008, days=151
        // Holidays: 11/26/2008, 12/4/2008, 1/21/2009

        // =WORKDAY(A2,A3) => 4/30/2009 (serial 39933)
        // Date 151 workdays from the start date
        assert_eq!(
            eval("=WORKDAY(DATE(2008,10,1),151)").unwrap(),
            FormulaValue::Number(39933.0)
        );

        // =WORKDAY(A2,A3,A4:A6) => 5/5/2009 (serial 39938)
        // Date 151 workdays from the start date, excluding holidays
        assert_eq!(
            eval(
                "=WORKDAY(DATE(2008,10,1),151,{DATE(2008,11,26),DATE(2008,12,4),DATE(2009,1,21)})"
            )
            .unwrap(),
            FormulaValue::Number(39938.0)
        );

        // Negative days: go backwards 151 workdays from result to get start date
        // 4/30/2009 minus 151 workdays => 10/1/2008 (serial 39722)
        assert_eq!(
            eval("=WORKDAY(DATE(2009,4,30),-151)").unwrap(),
            FormulaValue::Number(39722.0)
        );

        // Remarks: "If days is not an integer, it is truncated"
        assert_eq!(
            eval("=WORKDAY(DATE(2008,10,1),151.9)").unwrap(),
            FormulaValue::Number(39933.0)
        );

        // Remarks: "If any argument is not a valid date, WORKDAY returns the #VALUE! error value"
        assert_eq!(
            eval("=WORKDAY(\"not a date\",151)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_workday_intl_docs() {
        // https://support.microsoft.com/en-us/office/workday-intl-function-a378391c-9ba7-4678-8a39-39611a9bf81d

        // Doc example 1: Using a 0 for the Weekend argument results in a #NUM! error.
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),30,0)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );

        // Doc example 2: 90 workdays from 1/1/2012, counting only Sundays as
        // a weekend day (weekend argument is 11). Result: 41013.
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),90,11)").unwrap(),
            FormulaValue::Number(41013.0)
        );

        // Doc example 3: 30 workdays from 1/1/2012, counting only Saturdays as
        // a weekend day (weekend argument is 17). Serial number 40944 = 2/05/2012.
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),30,17)").unwrap(),
            FormulaValue::Number(40944.0)
        );

        // Weekend string "0000001" = Sunday only (position 7), equivalent to weekend=11.
        // Docs: "Each character represents a day starting with Monday. 1=non-workday, 0=workday."
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),90,\"0000001\")").unwrap(),
            FormulaValue::Number(41013.0)
        );

        // Weekend string "0000010" = Saturday only (position 6), equivalent to weekend=17.
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),30,\"0000010\")").unwrap(),
            FormulaValue::Number(40944.0)
        );

        // Docs: "0000011 would result in a weekend that is Saturday and Sunday"
        // This is equivalent to weekend=1 (or omitted).
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),30,\"0000011\")").unwrap(),
            eval("=WORKDAY.INTL(DATE(2012,1,1),30,1)").unwrap()
        );

        // Docs remark: "Also, 1111111 is an invalid string."
        // All days are weekends so no workday is reachable => #NUM!
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),30,\"1111111\")").unwrap(),
            FormulaValue::Error(CellError::Num)
        );

        // Docs remark: "If a weekend string is of invalid length ... returns the #VALUE! error value."
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),30,\"00011\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // Docs remark: "...or contains invalid characters, WORKDAY.INTL returns the #VALUE! error value."
        assert_eq!(
            eval("=WORKDAY.INTL(DATE(2012,1,1),30,\"000002X\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_day_docs() {
        // Docs example: =DAY(A2) where A2 is 15-Apr-11, result 15
        assert_eq!(
            eval("=DAY(DATE(2011,4,15))").unwrap(),
            FormulaValue::Number(15.0)
        );

        // Docs remark: January 1, 1900 is serial number 1
        assert_eq!(eval("=DAY(1)").unwrap(), FormulaValue::Number(1.0));

        // Docs remark: January 1, 2008 is serial number 39448
        assert_eq!(eval("=DAY(39448)").unwrap(), FormulaValue::Number(1.0));

        // Docs syntax: use DATE(2008,5,23) for the 23rd day of May, 2008
        assert_eq!(
            eval("=DAY(DATE(2008,5,23))").unwrap(),
            FormulaValue::Number(23.0)
        );
    }

    #[test]
    fn test_days360_docs() {
        // https://support.microsoft.com/en-us/office/days360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a

        // =DAYS360(A3,A4) where A3=30-Jan-11, A4=1-Feb-11 => 1
        // "Number of days between 1/30/2011 and 2/1/2011, based on a 360-day year."
        assert_eq!(
            eval("=DAYS360(DATE(2011,1,30),DATE(2011,2,1))").unwrap(),
            FormulaValue::Number(1.0)
        );

        // =DAYS360(A2,A5) where A2=1-Jan-11, A5=31-Dec-11 => 360
        // "Number of days between 1/1/2011 and 12/31/2011, based on a 360-day year."
        assert_eq!(
            eval("=DAYS360(DATE(2011,1,1),DATE(2011,12,31))").unwrap(),
            FormulaValue::Number(360.0)
        );

        // =DAYS360(A2,A4) where A2=1-Jan-11, A4=1-Feb-11 => 30
        // "Number of days between 1/1/2011 and 2/1/2011, based on a 360-day year."
        assert_eq!(
            eval("=DAYS360(DATE(2011,1,1),DATE(2011,2,1))").unwrap(),
            FormulaValue::Number(30.0)
        );
    }

    #[test]
    fn test_isoweeknum_docs() {
        // https://support.microsoft.com/en-us/office/isoweeknum-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e

        // Docs example: =ISOWEEKNUM(A2) where A2 is 3/9/2012, result 10
        // "Number of the week in the year that 3/9/2012 occurs, based on weeks beginning on the default, Monday"
        assert_eq!(
            eval("=ISOWEEKNUM(DATE(2012,3,9))").unwrap(),
            FormulaValue::Number(10.0)
        );

        // Docs remark: January 1, 1900 is serial number 1
        assert_eq!(eval("=ISOWEEKNUM(1)").unwrap(), FormulaValue::Number(1.0));

        // Docs remark: January 1, 2008 is serial number 39448
        assert_eq!(
            eval("=ISOWEEKNUM(39448)").unwrap(),
            FormulaValue::Number(1.0)
        );

        // Year boundary: Dec 31 can belong to ISO week 1 of the next year
        // 2012-12-31 is a Monday => ISO week 1 of 2013
        assert_eq!(
            eval("=ISOWEEKNUM(DATE(2012,12,31))").unwrap(),
            FormulaValue::Number(1.0)
        );

        // Year boundary: Jan 1 can belong to a high ISO week of the previous year
        // 2010-01-01 is a Friday => ISO week 53 of 2009
        assert_eq!(
            eval("=ISOWEEKNUM(DATE(2010,1,1))").unwrap(),
            FormulaValue::Number(53.0)
        );

        // Docs remark: non-valid date type returns #VALUE!
        assert_eq!(
            eval("=ISOWEEKNUM(\"not a date\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_date_docs() {
        // https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349

        // === Reference serial numbers stated in docs ===
        // "January 1, 1900 is serial number 1"
        assert_eq!(eval("=DATE(1900,1,1)").unwrap(), FormulaValue::Number(1.0));

        // "January 1, 2008 is serial number 39448 because it is 39,447 days after January 1, 1900"
        assert_eq!(
            eval("=DATE(2008,1,1)").unwrap(),
            FormulaValue::Number(39448.0)
        );

        // === Year argument: values 0-1899 add 1900 ===
        // "DATE(108,1,2) returns January 2, 2008 (1900+108)"
        // Jan 2, 2008 = 39448 + 1 = 39449
        assert_eq!(
            eval("=DATE(108,1,2)").unwrap(),
            FormulaValue::Number(39449.0)
        );
        assert_eq!(
            eval("=YEAR(DATE(108,1,2))").unwrap(),
            FormulaValue::Number(2008.0)
        );
        assert_eq!(
            eval("=MONTH(DATE(108,1,2))").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=DAY(DATE(108,1,2))").unwrap(),
            FormulaValue::Number(2.0)
        );

        // === Year argument: values 1900-9999 used directly ===
        // "DATE(2008,1,2) returns January 2, 2008"
        assert_eq!(
            eval("=DATE(2008,1,2)").unwrap(),
            FormulaValue::Number(39449.0)
        );

        // DATE(108,1,2) must equal DATE(2008,1,2) since 108+1900=2008
        assert_eq!(
            eval("=DATE(108,1,2)").unwrap(),
            eval("=DATE(2008,1,2)").unwrap()
        );

        // Year = 0 maps to 1900: DATE(0,1,1) = DATE(1900,1,1) = serial 1
        assert_eq!(eval("=DATE(0,1,1)").unwrap(), FormulaValue::Number(1.0));

        // Year = 1899 maps to 3799: DATE(1899,1,1) = DATE(3799,1,1)
        assert_eq!(
            eval("=DATE(1899,1,1)").unwrap(),
            eval("=DATE(3799,1,1)").unwrap()
        );
        assert_eq!(
            eval("=YEAR(DATE(1899,1,1))").unwrap(),
            FormulaValue::Number(3799.0)
        );

        // === Year argument: error conditions ===
        // "If year is less than 0 ... Excel returns the #NUM! error value"
        assert_eq!(
            eval("=DATE(-1,1,1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );

        // "if year ... is 10000 or greater, Excel returns the #NUM! error value"
        assert_eq!(
            eval("=DATE(10000,1,1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );

        // Boundary: year = 9999 is the maximum valid year
        assert_eq!(
            eval("=YEAR(DATE(9999,1,1))").unwrap(),
            FormulaValue::Number(9999.0)
        );

        // Large negative year also errors
        assert_eq!(
            eval("=DATE(-100,1,1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );

        // === Month overflow (month > 12) ===
        // "DATE(2008,14,2) returns the serial number representing February 2, 2009"
        // 2008 is a leap year (366 days). Jan 1, 2009 = 39448+366 = 39814.
        // Feb 2, 2009 = 39814 + 31 (Jan) + 1 = 39846
        assert_eq!(
            eval("=DATE(2008,14,2)").unwrap(),
            FormulaValue::Number(39846.0)
        );
        assert_eq!(
            eval("=YEAR(DATE(2008,14,2))").unwrap(),
            FormulaValue::Number(2009.0)
        );
        assert_eq!(
            eval("=MONTH(DATE(2008,14,2))").unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            eval("=DAY(DATE(2008,14,2))").unwrap(),
            FormulaValue::Number(2.0)
        );

        // === Month underflow (negative month) ===
        // "DATE(2008,-3,2) returns the serial number representing September 2, 2007"
        // Jan 1, 2007 = 39448-365 = 39083. Sep 2, 2007 = 39083 + 243 (Jan-Aug) + 1 = 39327
        assert_eq!(
            eval("=DATE(2008,-3,2)").unwrap(),
            FormulaValue::Number(39327.0)
        );
        assert_eq!(
            eval("=YEAR(DATE(2008,-3,2))").unwrap(),
            FormulaValue::Number(2007.0)
        );
        assert_eq!(
            eval("=MONTH(DATE(2008,-3,2))").unwrap(),
            FormulaValue::Number(9.0)
        );
        assert_eq!(
            eval("=DAY(DATE(2008,-3,2))").unwrap(),
            FormulaValue::Number(2.0)
        );

        // === Day overflow (day > days in month) ===
        // "DATE(2008,1,35) returns the serial number representing February 4, 2008"
        // Jan has 31 days; 35-31 = 4 days into Feb. Serial = 39448 + 34 = 39482
        assert_eq!(
            eval("=DATE(2008,1,35)").unwrap(),
            FormulaValue::Number(39482.0)
        );
        assert_eq!(
            eval("=YEAR(DATE(2008,1,35))").unwrap(),
            FormulaValue::Number(2008.0)
        );
        assert_eq!(
            eval("=MONTH(DATE(2008,1,35))").unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            eval("=DAY(DATE(2008,1,35))").unwrap(),
            FormulaValue::Number(4.0)
        );

        // === Day underflow (negative day) ===
        // "DATE(2008,1,-15) returns the serial number representing December 16, 2007"
        // Serial = 39448 + (-15) - 1 = 39432. Verify: Dec 31, 2007 = 39447; 39447-15 = 39432
        assert_eq!(
            eval("=DATE(2008,1,-15)").unwrap(),
            FormulaValue::Number(39432.0)
        );
        assert_eq!(
            eval("=YEAR(DATE(2008,1,-15))").unwrap(),
            FormulaValue::Number(2007.0)
        );
        assert_eq!(
            eval("=MONTH(DATE(2008,1,-15))").unwrap(),
            FormulaValue::Number(12.0)
        );
        assert_eq!(
            eval("=DAY(DATE(2008,1,-15))").unwrap(),
            FormulaValue::Number(16.0)
        );

        // === Docs composite example: anniversary calculation ===
        // "DATE(YEAR(C2)+5,MONTH(C2),DAY(C2))" with date 3/14/2012 gives 3/14/2017
        assert_eq!(
            eval("=DATE(YEAR(DATE(2012,3,14))+5,MONTH(DATE(2012,3,14)),DAY(DATE(2012,3,14)))")
                .unwrap(),
            eval("=DATE(2017,3,14)").unwrap()
        );
        assert_eq!(
            eval(
                "=YEAR(DATE(YEAR(DATE(2012,3,14))+5,MONTH(DATE(2012,3,14)),DAY(DATE(2012,3,14))))"
            )
            .unwrap(),
            FormulaValue::Number(2017.0)
        );
        assert_eq!(
            eval(
                "=MONTH(DATE(YEAR(DATE(2012,3,14))+5,MONTH(DATE(2012,3,14)),DAY(DATE(2012,3,14))))"
            )
            .unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            eval("=DAY(DATE(YEAR(DATE(2012,3,14))+5,MONTH(DATE(2012,3,14)),DAY(DATE(2012,3,14))))")
                .unwrap(),
            FormulaValue::Number(14.0)
        );

        // === Additional edge cases implied by docs behavioral notes ===

        // Month = 0: underflows to December of prior year
        // "month subtracts the magnitude of that number of months, plus 1, from the first month"
        assert_eq!(
            eval("=YEAR(DATE(2008,0,1))").unwrap(),
            FormulaValue::Number(2007.0)
        );
        assert_eq!(
            eval("=MONTH(DATE(2008,0,1))").unwrap(),
            FormulaValue::Number(12.0)
        );
        assert_eq!(
            eval("=DAY(DATE(2008,0,1))").unwrap(),
            FormulaValue::Number(1.0)
        );

        // Day = 0: underflows to last day of previous month
        // "day subtracts the magnitude of that number of days, plus one, from the first day"
        // DATE(2008,2,0) = Jan 31, 2008
        assert_eq!(
            eval("=YEAR(DATE(2008,2,0))").unwrap(),
            FormulaValue::Number(2008.0)
        );
        assert_eq!(
            eval("=MONTH(DATE(2008,2,0))").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=DAY(DATE(2008,2,0))").unwrap(),
            FormulaValue::Number(31.0)
        );

        // Combined month and day overflow: DATE(2008,13,32)
        // month 13 -> Jan 2009; day 32 -> Feb 1, 2009
        assert_eq!(
            eval("=YEAR(DATE(2008,13,32))").unwrap(),
            FormulaValue::Number(2009.0)
        );
        assert_eq!(
            eval("=MONTH(DATE(2008,13,32))").unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            eval("=DAY(DATE(2008,13,32))").unwrap(),
            FormulaValue::Number(1.0)
        );

        // Docs tip: "07" could mean "1907" or "2007" - two-digit year 7 maps to 1907
        assert_eq!(
            eval("=YEAR(DATE(7,1,1))").unwrap(),
            FormulaValue::Number(1907.0)
        );

        // Increase/decrease a date by days (docs section: add/subtract days)
        // DATE(2008,1,1)+7 = Jan 8, 2008 = 39448 + 7 = 39455
        assert_eq!(
            eval("=DATE(2008,1,1)+7").unwrap(),
            FormulaValue::Number(39455.0)
        );
        // DATE(2008,1,1)-7 = Dec 25, 2007 = 39448 - 7 = 39441
        assert_eq!(
            eval("=DATE(2008,1,1)-7").unwrap(),
            FormulaValue::Number(39441.0)
        );
    }

    #[test]
    fn test_month_docs() {
        // https://support.microsoft.com/en-us/office/month-function-579a2881-199b-48b2-ab90-ddba0eba86e8

        // Docs example: =MONTH(A2) where A2 is 15-Apr-11, result 4
        assert_eq!(
            eval("=MONTH(DATE(2011,4,15))").unwrap(),
            FormulaValue::Number(4.0)
        );

        // Docs remarks: January 1, 1900 is serial number 1
        assert_eq!(eval("=MONTH(1)").unwrap(), FormulaValue::Number(1.0));

        // Docs remarks: January 1, 2008 is serial number 39448
        assert_eq!(eval("=MONTH(39448)").unwrap(), FormulaValue::Number(1.0));
    }

    #[test]
    fn test_now_docs() {
        // https://support.microsoft.com/en-us/office/now-function-3337fd29-145a-4347-b2e6-20c904739c46

        // NOW() returns the serial number of the current date and time.
        // Syntax: NOW() - no arguments.
        let now_val = number(eval("=NOW()").unwrap());

        // Docs remarks: Jan 1, 1900 is serial number 1, Jan 1, 2025 is 45658.
        // Should be a recent date (after 2020 ≈ serial 44000).
        assert!(
            now_val > 44000.0,
            "NOW() should return a recent date serial, got {now_val}"
        );

        // Docs remarks: numbers to the right of the decimal point represent time;
        // the serial number 0.5 represents 12:00 noon.
        // The fractional part should be in [0.0, 1.0).
        let frac = now_val.fract();
        assert!(
            frac >= 0.0 && frac < 1.0,
            "Time fraction should be in [0,1), got {frac}"
        );

        // Integer part should match TODAY() (same calendar day).
        let today_val = number(eval("=TODAY()").unwrap());
        assert_eq!(
            now_val.floor(),
            today_val,
            "NOW() date part should equal TODAY()"
        );

        // Docs example: =NOW()-0.5 returns the date and time 12 hours ago.
        // Subtracting 0.5 from NOW() should still be a valid recent serial.
        let half_day_ago = number(eval("=NOW()-0.5").unwrap());
        assert_close(now_val - half_day_ago, 0.5);

        // Docs example: =NOW()+7 returns the date and time 7 days in the future.
        let week_ahead = number(eval("=NOW()+7").unwrap());
        assert_close(week_ahead - now_val, 7.0);

        // Docs example: =NOW()-2.25 returns 2 days and 6 hours ago.
        let two_and_quarter = number(eval("=NOW()-2.25").unwrap());
        assert_close(now_val - two_and_quarter, 2.25);

        // TODAY() should be a whole number (no time component).
        assert_eq!(
            today_val.fract(),
            0.0,
            "TODAY() should have no fractional part"
        );
    }

    #[test]
    fn test_second_docs() {
        // https://support.microsoft.com/en-us/office/second-function-740d1cfc-553c-4099-b668-80eaa24e8af1

        // Docs example: =SECOND(A3) where A3="4:48:18 PM" => 18
        assert_eq!(
            eval("=SECOND(TIME(16,48,18))").unwrap(),
            FormulaValue::Number(18.0)
        );

        // Docs example: =SECOND(A4) where A4="4:48 PM" => 0
        assert_eq!(
            eval("=SECOND(TIME(16,48,0))").unwrap(),
            FormulaValue::Number(0.0)
        );

        // Syntax: text string via TIMEVALUE("6:45 PM") => second 0
        assert_eq!(
            eval("=SECOND(TIMEVALUE(\"6:45 PM\"))").unwrap(),
            FormulaValue::Number(0.0)
        );

        // Syntax: 0.78125 represents 6:45 PM => second 0
        assert_eq!(eval("=SECOND(0.78125)").unwrap(), FormulaValue::Number(0.0));
    }

    #[test]
    fn test_today_docs() {
        // https://support.microsoft.com/en-us/office/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9

        // Docs: "The TODAY function returns the serial number of the current date."
        // Syntax: TODAY() — no arguments.
        let today = number(eval("=TODAY()").unwrap());

        // Docs: serial number is the date-time code used for date/time calculations.
        // TODAY() should be a whole number (no time component).
        assert_eq!(
            today,
            today.floor(),
            "TODAY() should be a whole number (no time component)"
        );

        // Docs: "January 1, 1900 is serial number 1, and January 1, 2008 is serial
        // number 39448 because it is 39,447 days after January 1, 1900."
        // Should be a recent date (after 2020 ≈ serial ~44000).
        assert!(
            today > 44000.0,
            "TODAY() should return a recent date serial, got {today}"
        );

        // Docs example: =TODAY()+5 returns the current date plus 5 days.
        let plus_five = number(eval("=TODAY()+5").unwrap());
        assert_close(plus_five - today, 5.0);

        // Docs example: =DATEVALUE("1/1/2030")-TODAY() returns days until 1/1/2030.
        // Result should be a number (positive or negative depending on current date).
        let days_diff = number(eval("=DATEVALUE(\"1/1/2030\")-TODAY()").unwrap());
        // Jan 1, 2030 serial = 47484. Verify the arithmetic is consistent.
        assert_close(days_diff, 47484.0 - today);

        // Docs example: =DAY(TODAY()) returns the current day of the month (1-31).
        let day = number(eval("=DAY(TODAY())").unwrap());
        assert!(
            day >= 1.0 && day <= 31.0,
            "DAY(TODAY()) should be 1..31, got {day}"
        );

        // Docs example: =MONTH(TODAY()) returns the current month of the year (1-12).
        let month = number(eval("=MONTH(TODAY())").unwrap());
        assert!(
            month >= 1.0 && month <= 12.0,
            "MONTH(TODAY()) should be 1..12, got {month}"
        );

        // Docs age-calculation example: =YEAR(TODAY())-1963
        // "returns the person's age as of this year's birthday"
        let year = number(eval("=YEAR(TODAY())").unwrap());
        assert!(
            year >= 2025.0 && year <= 9999.0,
            "YEAR(TODAY()) should be recent, got {year}"
        );
        let age = number(eval("=YEAR(TODAY())-1963").unwrap());
        assert_close(age, year - 1963.0);

        // TODAY() used as argument to YEAR (docs: "uses the TODAY function as an
        // argument for the YEAR function to obtain the current year").
        assert_eq!(eval("=YEAR(TODAY())").unwrap(), FormulaValue::Number(year));

        // Verify TODAY() reconstitutes via DATE(YEAR,MONTH,DAY).
        assert_eq!(
            number(eval("=DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY()))").unwrap()),
            today
        );
    }

    #[test]
    fn test_timevalue_docs() {
        // Microsoft docs: TIMEVALUE function
        // https://support.microsoft.com/en-us/office/timevalue-function-0b615c12-33d8-4431-bf3d-f3eb6d186645

        // =TIMEVALUE("2:24 AM") → 0.10
        // 2:24 AM = (2*60+24)/(24*60) = 144/1440 = 0.1
        assert_close(number(eval("=TIMEVALUE(\"2:24 AM\")").unwrap()), 0.1);

        // =TIMEVALUE("22-Aug-2011 6:35 AM") → 0.2743
        // 6:35 AM = (6*60+35)/(24*60) = 395/1440 = 0.274305..., date info ignored
        assert_close(
            number(eval("=TIMEVALUE(\"22-Aug-2011 6:35 AM\")").unwrap()),
            395.0 / 1440.0,
        );
    }

    #[test]
    fn test_time_docs() {
        // https://support.microsoft.com/en-us/office/time-function-9a5aff99-8f7d-4611-845e-747d0b8d5457

        // Example table row 1: =TIME(12,0,0) => 0.5
        assert_close(number(eval("=TIME(12,0,0)").unwrap()), 0.5);

        // Example table row 2: =TIME(16,48,10) => 0.7001157
        assert_close(number(eval("=TIME(16,48,10)").unwrap()), 60490.0 / 86400.0);

        // Hour overflow: TIME(27,0,0) = TIME(3,0,0) = .125
        assert_close(number(eval("=TIME(27,0,0)").unwrap()), 0.125);
        assert_close(number(eval("=TIME(3,0,0)").unwrap()), 0.125);

        // Minute overflow: TIME(0,750,0) = TIME(12,30,0) = .520833
        assert_close(number(eval("=TIME(0,750,0)").unwrap()), 45000.0 / 86400.0);
        assert_close(number(eval("=TIME(12,30,0)").unwrap()), 45000.0 / 86400.0);

        // Second overflow: TIME(0,0,2000) = TIME(0,33,20) = .023148
        assert_close(number(eval("=TIME(0,0,2000)").unwrap()), 2000.0 / 86400.0);
        assert_close(number(eval("=TIME(0,33,20)").unwrap()), 2000.0 / 86400.0);
    }

    #[test]
    fn test_year_docs() {
        // https://support.microsoft.com/en-us/office/year-function-c64f017a-1354-490d-981f-578e8ec8d3b9

        // Docs example: =YEAR(A3) where A3 is 7/5/2023, result 2023
        assert_eq!(
            eval("=YEAR(DATE(2023,7,5))").unwrap(),
            FormulaValue::Number(2023.0)
        );

        // Docs example: =YEAR(A4) where A4 is 7/5/2025, result 2025
        assert_eq!(
            eval("=YEAR(DATE(2025,7,5))").unwrap(),
            FormulaValue::Number(2025.0)
        );

        // Docs remarks: January 1, 1900 is serial number 1
        assert_eq!(eval("=YEAR(1)").unwrap(), FormulaValue::Number(1900.0));

        // Docs remarks: January 1, 2008 is serial number 39448
        assert_eq!(eval("=YEAR(39448)").unwrap(), FormulaValue::Number(2008.0));

        // Docs syntax: use DATE(2025,5,23) for the 23rd day of May, 2025
        assert_eq!(
            eval("=YEAR(DATE(2025,5,23))").unwrap(),
            FormulaValue::Number(2025.0)
        );
    }

    #[test]
    fn test_networkdays_intl_docs() {
        // https://support.microsoft.com/en-us/office/networkdays-intl-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28

        // Example 1: Default weekend (Saturday, Sunday)
        // "Results in 22 future workdays. Subtracts 9 nonworking weekend days
        // (5 Saturdays and 4 Sundays) from the 31 total days."
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))").unwrap(),
            FormulaValue::Number(22.0)
        );

        // Example 2: Negative result when start_date is later than end_date
        // "Results in -21, which is 21 workdays in the past."
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))").unwrap(),
            FormulaValue::Number(-21.0)
        );

        // Example 3: Weekend=7 (Friday, Saturday) with two holidays
        // "Results in 22 future workdays by subtracting 10 nonworking days
        // (4 Fridays, 4 Saturdays, 2 Holidays) from the 32 days."
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{DATE(2006,1,2),DATE(2006,1,16)})").unwrap(),
            FormulaValue::Number(22.0)
        );

        // Example 4: Weekend string "0010001" (Wednesday, Sunday) with holidays
        // "Results in 20 future workdays. Same time period as example directly
        // above, but with Sunday and Wednesday as weekend days."
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),\"0010001\",{DATE(2006,1,2),DATE(2006,1,16)})").unwrap(),
            FormulaValue::Number(20.0)
        );

        // --- Weekend number values (docs table) using Example 1 date range ---

        // Weekend 1 (Saturday, Sunday) - explicit, same as omitted
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),1)").unwrap(),
            FormulaValue::Number(22.0)
        );

        // Weekend 2 (Sunday, Monday)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),2)").unwrap(),
            FormulaValue::Number(21.0)
        );

        // Weekend 3 (Monday, Tuesday)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),3)").unwrap(),
            FormulaValue::Number(21.0)
        );

        // Weekend 4 (Tuesday, Wednesday)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),4)").unwrap(),
            FormulaValue::Number(22.0)
        );

        // Weekend 5 (Wednesday, Thursday)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),5)").unwrap(),
            FormulaValue::Number(23.0)
        );

        // Weekend 6 (Thursday, Friday)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),6)").unwrap(),
            FormulaValue::Number(23.0)
        );

        // Weekend 7 (Friday, Saturday)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),7)").unwrap(),
            FormulaValue::Number(23.0)
        );

        // --- Single-day weekend codes (docs table) ---

        // Weekend 11 (Sunday only)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),11)").unwrap(),
            FormulaValue::Number(26.0)
        );

        // Weekend 12 (Monday only)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),12)").unwrap(),
            FormulaValue::Number(26.0)
        );

        // Weekend 13 (Tuesday only)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),13)").unwrap(),
            FormulaValue::Number(26.0)
        );

        // Weekend 14 (Wednesday only)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),14)").unwrap(),
            FormulaValue::Number(27.0)
        );

        // Weekend 15 (Thursday only)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),15)").unwrap(),
            FormulaValue::Number(27.0)
        );

        // Weekend 16 (Friday only)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),16)").unwrap(),
            FormulaValue::Number(27.0)
        );

        // Weekend 17 (Saturday only)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),17)").unwrap(),
            FormulaValue::Number(27.0)
        );

        // --- Weekend string patterns (docs) ---

        // "0000011" = Saturday and Sunday (docs example, equivalent to code 1)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"0000011\")").unwrap(),
            FormulaValue::Number(22.0)
        );

        // "1111111" will always return 0 (docs: all days are weekends)
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"1111111\")").unwrap(),
            FormulaValue::Number(0.0)
        );

        // --- Error cases (docs Remarks section) ---

        // "If a weekend string is of invalid length ... returns the #VALUE! error"
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"001\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // "... or contains invalid characters, NETWORKDAYS.INTL returns the #VALUE! error"
        assert_eq!(
            eval("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"0000022\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }
}
