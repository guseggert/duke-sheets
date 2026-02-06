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
use chrono::{Datelike, Duration, NaiveDate};
use duke_sheets_core::CellError;

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
    if year == 1900 {
        366
    } else if is_leap_gregorian(year) {
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

    let mut year = to_i32_trunc(args.get(0).unwrap()).unwrap_or(0);
    let month = to_i32_trunc(args.get(1).unwrap()).unwrap_or(0);
    let day = to_i32_trunc(args.get(2).unwrap()).unwrap_or(0);

    // Excel: years 0..1899 are treated as 1900..3799
    if (0..1900).contains(&year) {
        year += 1900;
    }

    // Basic bounds (Excel supports 0..9999 in DATE)
    if year < 0 || year > 9999 {
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
    let v = args.get(0).unwrap();
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
    let v = args.get(0).unwrap();
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
    let v = args.get(0).unwrap();
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
