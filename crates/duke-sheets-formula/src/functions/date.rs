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
        .or_else(|_| NaiveDate::parse_from_str(text, "%d-%b-%y"));

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
        .or_else(|_| chrono::NaiveTime::parse_from_str(&upper, "%I:%M %p"));

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
}
