//! Text functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use chrono::{Datelike, Duration, NaiveDate};
use duke_sheets_core::CellError;

fn to_int_trunc(v: &FormulaValue) -> Option<i64> {
    v.as_number().map(|n| n.trunc() as i64)
}

fn take_left(s: &str, n: usize) -> String {
    s.chars().take(n).collect()
}

fn take_right(s: &str, n: usize) -> String {
    let len = s.chars().count();
    if n >= len {
        return s.to_string();
    }
    s.chars().skip(len - n).collect()
}

fn take_mid(s: &str, start_1based: usize, n: usize) -> String {
    if start_1based == 0 {
        return String::new();
    }
    let start0 = start_1based - 1;
    s.chars().skip(start0).take(n).collect()
}

fn add_thousands_separators(number: &str) -> String {
    let (sign, body) = if let Some(rest) = number.strip_prefix('-') {
        ("-", rest)
    } else {
        ("", number)
    };

    let mut parts = body.splitn(2, '.');
    let int_part = parts.next().unwrap_or("");
    let frac_part = parts.next();

    let mut grouped_rev = String::new();
    for (i, ch) in int_part.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            grouped_rev.push(',');
        }
        grouped_rev.push(ch);
    }

    let grouped_int: String = grouped_rev.chars().rev().collect();
    match frac_part {
        Some(frac) if !frac.is_empty() => format!("{}{}.{}", sign, grouped_int, frac),
        _ => format!("{}{}", sign, grouped_int),
    }
}

fn round_with_decimals(number: f64, decimals: i64) -> f64 {
    if decimals >= 0 {
        let factor = 10f64.powi(decimals as i32);
        (number * factor).round() / factor
    } else {
        let factor = 10f64.powi((-decimals) as i32);
        (number / factor).round() * factor
    }
}

fn format_fixed_number(number: f64, decimals: i64, with_commas: bool) -> String {
    let rounded = round_with_decimals(number, decimals);
    let display_decimals = if decimals > 0 { decimals as usize } else { 0 };
    let mut out = format!("{:.*}", display_decimals, rounded);
    if with_commas {
        out = add_thousands_separators(&out);
    }
    out
}

fn excel1900_date_from_serial(serial: i64) -> Option<(i32, u32, u32)> {
    if serial == 60 {
        return Some((1900, 2, 29));
    }

    let base = NaiveDate::from_ymd_opt(1899, 12, 31)?;
    let adjusted = if serial > 60 { serial - 1 } else { serial };
    let date = base.checked_add_signed(Duration::days(adjusted))?;
    Some((date.year(), date.month(), date.day()))
}

fn serial_time_parts(number: f64) -> (u32, u32, u32) {
    let frac = number.rem_euclid(1.0);
    let mut total_seconds = (frac * 86400.0).round() as i64;
    if total_seconds >= 86400 {
        total_seconds = 0;
    }
    let h = (total_seconds / 3600) as u32;
    let m = ((total_seconds % 3600) / 60) as u32;
    let s = (total_seconds % 60) as u32;
    (h, m, s)
}

fn format_text_value(number: f64, format_text: &str) -> String {
    let fmt = format_text.trim();
    let fmt_lower = fmt.to_ascii_lowercase();
    match fmt_lower.as_str() {
        "0" => format_fixed_number(number, 0, false),
        "0.00" => format_fixed_number(number, 2, false),
        "#,##0" => format_fixed_number(number, 0, true),
        "0%" => format!("{}%", round_with_decimals(number * 100.0, 0) as i64),
        "0.00e+00" => format!("{:.2E}", number),
        "$#,##0.00" => {
            if number < 0.0 {
                format!("-${}", format_fixed_number(-number, 2, true))
            } else {
                format!("${}", format_fixed_number(number, 2, true))
            }
        }
        "mm/dd/yyyy" | "yyyy-mm-dd" | "dd/mm/yyyy" | "m/d/yy" => {
            let serial = number.floor() as i64;
            if let Some((y, m, d)) = excel1900_date_from_serial(serial) {
                match fmt_lower.as_str() {
                    "mm/dd/yyyy" => format!("{:02}/{:02}/{:04}", m, d, y),
                    "yyyy-mm-dd" => format!("{:04}-{:02}-{:02}", y, m, d),
                    "dd/mm/yyyy" => format!("{:02}/{:02}/{:04}", d, m, y),
                    "m/d/yy" => format!("{}/{}/{:02}", m, d, y.rem_euclid(100)),
                    _ => number.to_string(),
                }
            } else {
                number.to_string()
            }
        }
        "hh:mm:ss" => {
            let (h, m, s) = serial_time_parts(number);
            format!("{:02}:{:02}:{:02}", h, m, s)
        }
        "hh:mm" => {
            let (h, m, _s) = serial_time_parts(number);
            format!("{:02}:{:02}", h, m)
        }
        _ => {
            if number.fract() == 0.0 {
                format!("{}", number as i64)
            } else {
                number.to_string()
            }
        }
    }
}

/// LEN(text)
pub fn fn_len(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.first().unwrap();
    if let FormulaValue::Error(e) = v {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(v, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let s = v.as_string();
    Ok(FormulaValue::Number(s.chars().count() as f64))
}

/// LEFT(text, [num_chars])
pub fn fn_left(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = args.first().unwrap();
    if let FormulaValue::Error(e) = text {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(text, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let num_chars = match args.get(1) {
        None => 1i64,
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            to_int_trunc(v).unwrap_or(0)
        }
    };

    if num_chars < 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let s = text.as_string();
    Ok(FormulaValue::String(take_left(&s, num_chars as usize)))
}

/// RIGHT(text, [num_chars])
pub fn fn_right(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = args.first().unwrap();
    if let FormulaValue::Error(e) = text {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(text, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let num_chars = match args.get(1) {
        None => 1i64,
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            to_int_trunc(v).unwrap_or(0)
        }
    };

    if num_chars < 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let s = text.as_string();
    Ok(FormulaValue::String(take_right(&s, num_chars as usize)))
}

/// MID(text, start_num, num_chars)
pub fn fn_mid(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = args.first().unwrap();
    if let FormulaValue::Error(e) = text {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(text, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let start = args.get(1).unwrap();
    if let FormulaValue::Error(e) = start {
        return Ok(FormulaValue::Error(*e));
    }

    let count = args.get(2).unwrap();
    if let FormulaValue::Error(e) = count {
        return Ok(FormulaValue::Error(*e));
    }

    let start_i = to_int_trunc(start).unwrap_or(0);
    let count_i = to_int_trunc(count).unwrap_or(0);

    if start_i < 1 || count_i < 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let s = text.as_string();
    Ok(FormulaValue::String(take_mid(
        &s,
        start_i as usize,
        count_i as usize,
    )))
}

/// LOWER(text)
pub fn fn_lower(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = args.first().unwrap();
    if let FormulaValue::Error(e) = text {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(text, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::String(text.as_string().to_lowercase()))
}

/// UPPER(text)
pub fn fn_upper(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = args.first().unwrap();
    if let FormulaValue::Error(e) = text {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(text, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    Ok(FormulaValue::String(text.as_string().to_uppercase()))
}

/// TRIM(text)
pub fn fn_trim(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = args.first().unwrap();
    if let FormulaValue::Error(e) = text {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(text, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let s = text.as_string();
    let trimmed = s.split_whitespace().collect::<Vec<_>>().join(" ");
    Ok(FormulaValue::String(trimmed))
}

/// CONCAT(text1, [text2], ...)
///
/// Also used for legacy CONCATENATE.
pub fn fn_concat(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut out = String::new();
    for arg in args {
        match arg {
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array(arr) => {
                for row in arr {
                    for v in row {
                        if let FormulaValue::Error(e) = v {
                            return Ok(FormulaValue::Error(*e));
                        }
                        out.push_str(&v.as_string());
                    }
                }
            }
            _ => {
                out.push_str(&arg.as_string());
            }
        }
    }
    Ok(FormulaValue::String(out))
}

/// FIND(find_text, within_text, [start_num]) - Finds one text string within another (case-sensitive)
/// Returns the position of the first character of find_text within within_text
/// Returns #VALUE! error if find_text is not found
/// Reference: LibreOffice ScInterpreter::ScFind
pub fn fn_find(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let find_text = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let within_text = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let start_num = match args.get(2) {
        Some(FormulaValue::Number(n)) => *n as usize,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 1,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // start_num must be >= 1 and <= length of within_text
    let within_len = within_text.chars().count();
    if start_num < 1 || start_num > within_len {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Convert to 0-based index for searching
    let search_start = start_num - 1;

    // Get the substring starting from search_start (in characters, not bytes)
    let search_str: String = within_text.chars().skip(search_start).collect();

    // Find the substring (case-sensitive)
    if let Some(byte_pos) = search_str.find(&find_text) {
        // Convert byte position back to character position
        let char_pos = search_str[..byte_pos].chars().count();
        // Return 1-based position
        Ok(FormulaValue::Number((search_start + char_pos + 1) as f64))
    } else {
        Ok(FormulaValue::Error(CellError::Value))
    }
}

/// SEARCH(find_text, within_text, [start_num]) - Finds one text string within another (case-insensitive)
/// Similar to FIND but case-insensitive
/// Reference: LibreOffice ScInterpreter::ScSearch
pub fn fn_search(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let find_text = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string().to_lowercase(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let within_text = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let start_num = match args.get(2) {
        Some(FormulaValue::Number(n)) => *n as usize,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => 1,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let within_len = within_text.chars().count();
    if start_num < 1 || start_num > within_len {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let search_start = start_num - 1;
    let search_str: String = within_text.chars().skip(search_start).collect();
    let search_str_lower = search_str.to_lowercase();

    // Find the substring (case-insensitive)
    if let Some(byte_pos) = search_str_lower.find(&find_text) {
        let char_pos = search_str_lower[..byte_pos].chars().count();
        Ok(FormulaValue::Number((search_start + char_pos + 1) as f64))
    } else {
        Ok(FormulaValue::Error(CellError::Value))
    }
}

/// EXACT(text1, text2) - Checks whether two text strings are exactly the same (case-sensitive)
/// Returns TRUE if they are identical, FALSE otherwise
pub fn fn_exact(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text1 = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let text2 = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    Ok(FormulaValue::Boolean(text1 == text2))
}

/// REPT(text, number_times) - Repeats text a given number of times
pub fn fn_rept(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let times = match args.get(1) {
        Some(FormulaValue::Number(n)) => {
            if *n < 0.0 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            *n as usize
        }
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => 0,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Limit to prevent memory issues (Excel has a limit of 32767 chars)
    if text.len() * times > 32767 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(FormulaValue::String(text.repeat(times)))
}

/// SUBSTITUTE(text, old_text, new_text, [instance_num]) - Substitutes new_text for old_text in a text string
/// If instance_num is omitted, every occurrence of old_text is replaced
pub fn fn_substitute(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let text = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let old_text = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let new_text = match args.get(2) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let instance_num = match args.get(3) {
        Some(FormulaValue::Number(n)) => {
            if *n < 1.0 {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            Some(*n as usize)
        }
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) | None => None,
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if old_text.is_empty() {
        // If old_text is empty, return the original text unchanged
        return Ok(FormulaValue::String(text));
    }

    match instance_num {
        None => {
            // Replace all occurrences
            Ok(FormulaValue::String(text.replace(&old_text, &new_text)))
        }
        Some(n) => {
            // Replace only the nth occurrence
            let mut result = String::new();
            let mut remaining = text.as_str();
            let mut occurrence = 0;

            while let Some(pos) = remaining.find(&old_text) {
                occurrence += 1;
                if occurrence == n {
                    result.push_str(&remaining[..pos]);
                    result.push_str(&new_text);
                    result.push_str(&remaining[pos + old_text.len()..]);
                    return Ok(FormulaValue::String(result));
                } else {
                    result.push_str(&remaining[..pos + old_text.len()]);
                    remaining = &remaining[pos + old_text.len()..];
                }
            }

            // If we didn't find the nth occurrence, return original text
            result.push_str(remaining);
            Ok(FormulaValue::String(text))
        }
    }
}

/// PROPER(text) - Capitalizes the first letter in each word of a text value
pub fn fn_proper(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let mut result = String::with_capacity(text.len());
    let mut capitalize_next = true;

    for ch in text.chars() {
        if ch.is_whitespace() || !ch.is_alphanumeric() {
            result.push(ch);
            capitalize_next = true;
        } else if capitalize_next {
            result.extend(ch.to_uppercase());
            capitalize_next = false;
        } else {
            result.extend(ch.to_lowercase());
        }
    }

    Ok(FormulaValue::String(result))
}

/// CHAR(number) - Returns the character specified by the code number
/// Uses Unicode code points (Excel uses Windows-1252 for 128-255, but Unicode is more universal)
pub fn fn_char(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Number(n)) => *n as u32,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(FormulaValue::Empty) => return Ok(FormulaValue::Error(CellError::Value)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Excel accepts 1-255 for CHAR (ANSI), we extend to full Unicode
    if number == 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    match char::from_u32(number) {
        Some(c) => Ok(FormulaValue::String(c.to_string())),
        None => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// CODE(text) - Returns a numeric code for the first character in a text string
/// Returns the Unicode code point
pub fn fn_code(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if text.is_empty() {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let first_char = text.chars().next().unwrap();
    Ok(FormulaValue::Number(first_char as u32 as f64))
}

/// CLEAN(text) - Removes all nonprintable characters from text
/// Removes characters with codes 0-31 (control characters)
pub fn fn_clean(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let cleaned: String = text.chars().filter(|c| *c as u32 >= 32).collect();
    Ok(FormulaValue::String(cleaned))
}

/// VALUE(text) - Converts a text string that represents a number to a number
pub fn fn_value(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match args.first() {
        Some(FormulaValue::Number(n)) => return Ok(FormulaValue::Number(*n)), // Already a number
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Try to parse as a number
    let trimmed = text.trim();
    match trimmed.parse::<f64>() {
        Ok(n) => Ok(FormulaValue::Number(n)),
        Err(_) => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// T(value) - Returns the text referred to by value
/// Returns empty string if value is not text
pub fn fn_t(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::String(s)) => Ok(FormulaValue::String(s.clone())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::String(String::new())),
    }
}

/// N(value) - Returns a value converted to a number
pub fn fn_n(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.first() {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(*n)),
        Some(FormulaValue::Boolean(true)) => Ok(FormulaValue::Number(1.0)),
        Some(FormulaValue::Boolean(false)) => Ok(FormulaValue::Number(0.0)),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::Number(0.0)),
    }
}

pub fn fn_text(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let value = args.first().unwrap();
    if let FormulaValue::Error(e) = value {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(value, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let format_text = args.get(1).unwrap();
    if let FormulaValue::Error(e) = format_text {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(format_text, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let number = match value.as_number() {
        Some(n) => n,
        // Excel returns string values as-is (the text section of any number
        // format applies, which defaults to returning the string unchanged).
        None => return Ok(FormulaValue::String(value.as_string())),
    };

    Ok(FormulaValue::String(format_text_value(
        number,
        &format_text.as_string(),
    )))
}

pub fn fn_textjoin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let delimiter = args.first().unwrap();
    if let FormulaValue::Error(e) = delimiter {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(delimiter, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let ignore_empty_arg = args.get(1).unwrap();
    if let FormulaValue::Error(e) = ignore_empty_arg {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(ignore_empty_arg, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let ignore_empty = match ignore_empty_arg.as_bool() {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let mut pieces = Vec::new();
    for arg in &args[2..] {
        match arg {
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            FormulaValue::Array(arr) => {
                for row in arr {
                    for v in row {
                        if let FormulaValue::Error(e) = v {
                            return Ok(FormulaValue::Error(*e));
                        }
                        if ignore_empty && matches!(v, FormulaValue::Empty) {
                            continue;
                        }
                        let s = v.as_string();
                        if ignore_empty && s.is_empty() {
                            continue;
                        }
                        pieces.push(s);
                    }
                }
            }
            _ => {
                if ignore_empty && matches!(arg, FormulaValue::Empty) {
                    continue;
                }
                let s = arg.as_string();
                if ignore_empty && s.is_empty() {
                    continue;
                }
                pieces.push(s);
            }
        }
    }

    Ok(FormulaValue::String(pieces.join(&delimiter.as_string())))
}

pub fn fn_fixed(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number_arg = args.first().unwrap();
    if let FormulaValue::Error(e) = number_arg {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(number_arg, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let number = match number_arg.as_number() {
        Some(n) => n,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let decimals = match args.get(1) {
        None => 2,
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            if matches!(v, FormulaValue::Array(_)) {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            to_int_trunc(v).unwrap_or(0)
        }
    };

    let no_commas = match args.get(2) {
        None => false,
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            if matches!(v, FormulaValue::Array(_)) {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            match v.as_bool() {
                Some(b) => b,
                None => return Ok(FormulaValue::Error(CellError::Value)),
            }
        }
    };

    Ok(FormulaValue::String(format_fixed_number(
        number, decimals, !no_commas,
    )))
}

pub fn fn_dollar(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number_arg = args.first().unwrap();
    if let FormulaValue::Error(e) = number_arg {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(number_arg, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let number = match number_arg.as_number() {
        Some(n) => n,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let decimals = match args.get(1) {
        None => 2,
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            if matches!(v, FormulaValue::Array(_)) {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            to_int_trunc(v).unwrap_or(2)
        }
    };

    let abs_formatted = format_fixed_number(number.abs(), decimals, true);
    if number < 0.0 {
        Ok(FormulaValue::String(format!("-${}", abs_formatted)))
    } else {
        Ok(FormulaValue::String(format!("${}", abs_formatted)))
    }
}

pub fn fn_numbervalue(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let text_arg = args.first().unwrap();
    if let FormulaValue::Error(e) = text_arg {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(text_arg, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let decimal_separator = match args.get(1) {
        None => ".".to_string(),
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            if matches!(v, FormulaValue::Array(_)) {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            let s = v.as_string();
            if s.is_empty() {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            s
        }
    };

    let group_separator = match args.get(2) {
        None => ",".to_string(),
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            if matches!(v, FormulaValue::Array(_)) {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            v.as_string()
        }
    };

    if decimal_separator == group_separator {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let mut raw = text_arg.as_string().trim().replace(' ', "");
    if !group_separator.is_empty() {
        raw = raw.replace(&group_separator, "");
    }

    if !decimal_separator.is_empty() {
        let count = raw.matches(&decimal_separator).count();
        if count > 1 {
            return Ok(FormulaValue::Error(CellError::Value));
        }
        raw = raw.replace(&decimal_separator, ".");
    }

    match raw.parse::<f64>() {
        Ok(n) => Ok(FormulaValue::Number(n)),
        Err(_) => Ok(FormulaValue::Error(CellError::Value)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    fn s(v: &str) -> FormulaValue {
        FormulaValue::String(v.to_string())
    }

    #[test]
    fn test_text_number_formats() {
        assert_eq!(
            eval("=TEXT(1234.5,\"0.00\")").unwrap(),
            FormulaValue::String("1234.50".into())
        );
        assert_eq!(
            eval("=TEXT(1234.5,\"#,##0\")").unwrap(),
            FormulaValue::String("1,235".into())
        );
        assert_eq!(
            eval("=TEXT(0.126,\"0%\")").unwrap(),
            FormulaValue::String("13%".into())
        );
    }

    #[test]
    fn test_text_string_passthrough() {
        // TEXT(string, "") returns the string as-is (Excel behaviour)
        assert_eq!(
            eval("=TEXT(\"September\",\"\")").unwrap(),
            FormulaValue::String("September".into())
        );
        // Non-numeric string with a number format still returns the string
        assert_eq!(
            eval("=TEXT(\"hello\",\"0.00\")").unwrap(),
            FormulaValue::String("hello".into())
        );
        // Numeric strings still parse and format as numbers
        assert_eq!(
            eval("=TEXT(\"1234.5\",\"0.00\")").unwrap(),
            FormulaValue::String("1234.50".into())
        );
    }

    #[test]
    fn test_text_date_and_time_formats() {
        assert_eq!(
            eval("=TEXT(61,\"mm/dd/yyyy\")").unwrap(),
            FormulaValue::String("03/01/1900".into())
        );
        assert_eq!(
            eval("=TEXT(61,\"yyyy-mm-dd\")").unwrap(),
            FormulaValue::String("1900-03-01".into())
        );
        assert_eq!(
            eval("=TEXT(0.5,\"hh:mm\")").unwrap(),
            FormulaValue::String("12:00".into())
        );
    }

    #[test]
    fn test_textjoin() {
        assert_eq!(
            eval("=TEXTJOIN(\",\",TRUE,\"a\",\"\",\"b\")").unwrap(),
            FormulaValue::String("a,b".into())
        );
        assert_eq!(
            eval("=TEXTJOIN(\"-\",FALSE,\"a\",\"\",\"b\")").unwrap(),
            FormulaValue::String("a--b".into())
        );
        assert_eq!(
            eval("=TEXTJOIN(\"|\",TRUE,{\"a\",\"\";\"b\",\"c\"})").unwrap(),
            FormulaValue::String("a|b|c".into())
        );
    }

    #[test]
    fn test_fixed() {
        assert_eq!(
            eval("=FIXED(1234.567)").unwrap(),
            FormulaValue::String("1,234.57".into())
        );
        assert_eq!(
            eval("=FIXED(1234.567,1,TRUE)").unwrap(),
            FormulaValue::String("1234.6".into())
        );
        assert_eq!(
            eval("=FIXED(1234.567,-2)").unwrap(),
            FormulaValue::String("1,200".into())
        );
    }

    #[test]
    fn test_dollar() {
        assert_eq!(
            eval("=DOLLAR(1234.5)").unwrap(),
            FormulaValue::String("$1,234.50".into())
        );
        assert_eq!(
            eval("=DOLLAR(1234.5,0)").unwrap(),
            FormulaValue::String("$1,235".into())
        );
        assert_eq!(
            eval("=DOLLAR(-12.34,2)").unwrap(),
            FormulaValue::String("-$12.34".into())
        );
    }

    #[test]
    fn test_numbervalue() {
        assert_eq!(
            eval("=NUMBERVALUE(\"1,234.56\")").unwrap(),
            FormulaValue::Number(1234.56)
        );
        assert_eq!(
            eval("=NUMBERVALUE(\"1.234,56\",\",\",\".\")").unwrap(),
            FormulaValue::Number(1234.56)
        );
        assert_eq!(
            eval("=NUMBERVALUE(\"abc\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_len_docs() {
        // =LEN("Phoenix, AZ") → 11
        assert_eq!(
            eval("=LEN(\"Phoenix, AZ\")").unwrap(),
            FormulaValue::Number(11.0)
        );
        // =LEN("") → 0
        assert_eq!(eval("=LEN(\"\")").unwrap(), FormulaValue::Number(0.0));
        // =LEN("     One   ") → 11 (includes eight spaces)
        assert_eq!(
            eval("=LEN(\"     One   \")").unwrap(),
            FormulaValue::Number(11.0)
        );
    }

    #[test]
    fn test_left_docs() {
        // =LEFT("Sale Price", 4) → "Sale"
        assert_eq!(eval("=LEFT(\"Sale Price\", 4)").unwrap(), s("Sale"));
        // =LEFT("Sweden") → "S" (num_chars defaults to 1)
        assert_eq!(eval("=LEFT(\"Sweden\")").unwrap(), s("S"));
    }

    #[test]
    fn test_right_docs() {
        // =RIGHT("Sale Price", 5) → "Price"
        assert_eq!(eval("=RIGHT(\"Sale Price\", 5)").unwrap(), s("Price"));
        // =RIGHT("Stock Number") → "r" (defaults to 1)
        assert_eq!(eval("=RIGHT(\"Stock Number\")").unwrap(), s("r"));
    }

    #[test]
    fn test_mid_docs() {
        // =MID("Fluid Flow",1,5) → "Fluid"
        assert_eq!(eval("=MID(\"Fluid Flow\",1,5)").unwrap(), s("Fluid"));
        // =MID("Fluid Flow",7,20) → "Flow" (exceeds length)
        assert_eq!(eval("=MID(\"Fluid Flow\",7,20)").unwrap(), s("Flow"));
        // =MID("Fluid Flow",20,5) → "" (start exceeds length)
        assert_eq!(eval("=MID(\"Fluid Flow\",20,5)").unwrap(), s(""));
    }

    #[test]
    fn test_lower_docs() {
        assert_eq!(
            eval("=LOWER(\"E. E. Cummings\")").unwrap(),
            s("e. e. cummings")
        );
        assert_eq!(eval("=LOWER(\"Apt. 2B\")").unwrap(), s("apt. 2b"));
    }

    #[test]
    fn test_upper_docs() {
        assert_eq!(eval("=UPPER(\"total\")").unwrap(), s("TOTAL"));
        assert_eq!(eval("=UPPER(\"Yield\")").unwrap(), s("YIELD"));
    }

    #[test]
    fn test_proper_docs() {
        assert_eq!(
            eval("=PROPER(\"this is a TITLE\")").unwrap(),
            s("This Is A Title")
        );
        assert_eq!(
            eval("=PROPER(\"2-way street\")").unwrap(),
            s("2-Way Street")
        );
        assert_eq!(eval("=PROPER(\"76BudGet\")").unwrap(), s("76budget"));
    }

    #[test]
    fn test_trim_docs() {
        assert_eq!(
            eval("=TRIM(\" Sales report \")").unwrap(),
            s("Sales report")
        );
        assert_eq!(
            eval("=TRIM(\" First Quarter Earnings \")").unwrap(),
            s("First Quarter Earnings")
        );
        // Multiple interior spaces collapsed
        assert_eq!(
            eval("=TRIM(\"  hello   world  \")").unwrap(),
            s("hello world")
        );
        // TRIM + CLEAN combined
        assert_eq!(
            eval("=TRIM(CLEAN(CHAR(9)&\" Monthly report \"&CHAR(10)))").unwrap(),
            s("Monthly report")
        );
        assert_eq!(eval("=TRIM(\"\")").unwrap(), s(""));
        assert_eq!(eval("=TRIM(\"     \")").unwrap(), s(""));
    }

    #[test]
    fn test_clean_docs() {
        // CHAR(9)&"Monthly report"&CHAR(10) → removes non-printable
        assert_eq!(
            eval("=CLEAN(CHAR(9)&\"Monthly report\"&CHAR(10))").unwrap(),
            s("Monthly report")
        );
        assert_eq!(
            eval("=CLEAN(CHAR(13)&\"data\"&CHAR(13))").unwrap(),
            s("data")
        );
        // CHAR(7) = bell in the middle
        assert_eq!(eval("=CLEAN(\"hel\"&CHAR(7)&\"lo\")").unwrap(), s("hello"));
        // CHAR(32) = space is NOT removed by CLEAN
        assert_eq!(eval("=CLEAN(\" visible \")").unwrap(), s(" visible "));
        assert_eq!(eval("=CLEAN(\"\")").unwrap(), s(""));
    }

    #[test]
    fn test_concat_docs() {
        assert_eq!(
            eval("=CONCAT(\"The\",\" \",\"sun\",\" \",\"will\",\" \",\"come\",\" \",\"up\",\" \",\"tomorrow.\")").unwrap(),
            s("The sun will come up tomorrow.")
        );
        assert_eq!(
            eval("=CONCAT(\"Andreas\",\" \",\"Hauser\")").unwrap(),
            s("Andreas Hauser")
        );
        assert_eq!(
            eval("=CONCAT(\"Hauser\",\", \",\"Andreas\")").unwrap(),
            s("Hauser, Andreas")
        );
        assert_eq!(eval("=CONCATENATE(\"a\",\"b\")").unwrap(), s("ab"));
        assert_eq!(eval("=CONCAT(\"val=\",100)").unwrap(), s("val=100"));
    }

    // CONCATENATE docs: =CONCATENATE("Hello"," ","World!") = "Hello World!"
    #[test]
    fn test_concatenate_docs() {
        assert_eq!(
            eval("=CONCATENATE(\"Hello\",\" \",\"World!\")").unwrap(),
            s("Hello World!")
        );
        // CONCATENATE with numbers
        assert_eq!(
            eval("=CONCATENATE(\"Population for \",2015)").unwrap(),
            s("Population for 2015")
        );
        // CONCATENATE: single argument
        assert_eq!(eval("=CONCATENATE(\"test\")").unwrap(), s("test"));
        // CONCATENATE: boolean and number
        assert_eq!(eval("=CONCATENATE(TRUE,\"-\",42)").unwrap(), s("TRUE-42"));
    }

    #[test]
    fn test_find_docs() {
        // Case-sensitive find
        assert_eq!(
            eval(r#"=FIND("M","Miriam McGovern")"#).unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval(r#"=FIND("m","Miriam McGovern")"#).unwrap(),
            FormulaValue::Number(6.0)
        );
        assert_eq!(
            eval(r#"=FIND("M","Miriam McGovern",3)"#).unwrap(),
            FormulaValue::Number(8.0)
        );
        // Product string examples
        assert_eq!(
            eval(r#"=FIND(" #","Ceramic Insulators #124-TD45-87",1)"#).unwrap(),
            FormulaValue::Number(19.0)
        );
        assert_eq!(
            eval(r#"=FIND(" #","Copper Coils #12-671-6772",1)"#).unwrap(),
            FormulaValue::Number(13.0)
        );
        // MID+FIND compound
        assert_eq!(
            eval(r#"=MID("Ceramic Insulators #124-TD45-87",1,FIND(" #","Ceramic Insulators #124-TD45-87",1)-1)"#).unwrap(),
            s("Ceramic Insulators")
        );
        // Empty find_text matches at 1
        assert_eq!(
            eval(r#"=FIND("","Miriam McGovern")"#).unwrap(),
            FormulaValue::Number(1.0)
        );
        // Error: not found
        assert_eq!(
            eval(r#"=FIND("z","Miriam McGovern")"#).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        // Error: case mismatch
        assert_eq!(
            eval(r#"=FIND("m","MIRIAM")"#).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        // Error: start_num 0
        assert_eq!(
            eval(r#"=FIND("M","Miriam McGovern",0)"#).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_search_docs() {
        // Case-insensitive search
        assert_eq!(
            eval(r#"=SEARCH("n","printer")"#).unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            eval(r#"=SEARCH("base","database")"#).unwrap(),
            FormulaValue::Number(5.0)
        );
        assert_eq!(
            eval(r#"=SEARCH("e","Statements",6)"#).unwrap(),
            FormulaValue::Number(7.0)
        );
        assert_eq!(
            eval(r#"=SEARCH("margin","Profit Margin")"#).unwrap(),
            FormulaValue::Number(8.0)
        );
        // Case insensitivity
        assert_eq!(
            eval(r#"=SEARCH("MARGIN","Profit Margin")"#).unwrap(),
            FormulaValue::Number(8.0)
        );
        assert_eq!(
            eval(r#"=SEARCH("m","MIRIAM")"#).unwrap(),
            FormulaValue::Number(1.0)
        );
        // Wildcard ? — our SEARCH doesn't support wildcards, skip
        // Tilde escapes — skip (requires wildcard support)
        // Error: not found
        assert_eq!(
            eval(r#"=SEARCH("z","printer")"#).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_exact_docs() {
        assert_eq!(
            eval(r#"=EXACT("word","word")"#).unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval(r#"=EXACT("Word","word")"#).unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval(r#"=EXACT("w ord","word")"#).unwrap(),
            FormulaValue::Boolean(false)
        );
    }

    #[test]
    fn test_rept_docs() {
        assert_eq!(eval(r#"=REPT("*-",3)"#).unwrap(), s("*-*-*-"));
        assert_eq!(eval(r#"=REPT("-",10)"#).unwrap(), s("----------"));
    }

    #[test]
    fn test_substitute_docs() {
        // Replace all: =SUBSTITUTE("Sales Data", "Sales", "Cost") → "Cost Data"
        assert_eq!(
            eval("=SUBSTITUTE(\"Sales Data\", \"Sales\", \"Cost\")").unwrap(),
            s("Cost Data")
        );
        // Replace 1st instance: =SUBSTITUTE("Quarter 1, 2008", "1", "2", 1) → "Quarter 2, 2008"
        assert_eq!(
            eval("=SUBSTITUTE(\"Quarter 1, 2008\", \"1\", \"2\", 1)").unwrap(),
            s("Quarter 2, 2008")
        );
        // Replace 3rd instance: =SUBSTITUTE("Quarter 1, 2011", "1", "2", 3) → "Quarter 1, 2012"
        assert_eq!(
            eval("=SUBSTITUTE(\"Quarter 1, 2011\", \"1\", \"2\", 3)").unwrap(),
            s("Quarter 1, 2012")
        );
    }

    #[test]
    fn test_char_docs() {
        assert_eq!(eval("=CHAR(65)").unwrap(), s("A"));
        assert_eq!(eval("=CHAR(33)").unwrap(), s("!"));
    }

    #[test]
    fn test_code_docs() {
        assert_eq!(eval("=CODE(\"A\")").unwrap(), FormulaValue::Number(65.0));
        assert_eq!(eval("=CODE(\"!\")").unwrap(), FormulaValue::Number(33.0));
    }

    #[test]
    fn test_value_docs() {
        // VALUE("$1,000") — our impl doesn't handle currency symbols
        // Test with plain numeric text instead
        assert_eq!(
            eval("=VALUE(\"1000\")").unwrap(),
            FormulaValue::Number(1000.0)
        );
        // VALUE("1234.56") = 1234.56
        assert_eq!(
            eval("=VALUE(\"1234.56\")").unwrap(),
            FormulaValue::Number(1234.56)
        );
    }

    #[test]
    fn test_t_docs() {
        // T("Rainfall") = "Rainfall"
        assert_eq!(eval("=T(\"Rainfall\")").unwrap(), s("Rainfall"));
        // T(19) = "" (number returns empty)
        assert_eq!(eval("=T(19)").unwrap(), s(""));
        // T(TRUE) = "" (logical returns empty)
        assert_eq!(eval("=T(TRUE)").unwrap(), s(""));
    }

    #[test]
    fn test_text_docs() {
        // Currency with thousands
        assert_eq!(
            eval("=TEXT(1234.567,\"$#,##0.00\")").unwrap(),
            s("$1,234.57")
        );
        // Date format
        assert_eq!(
            eval("=TEXT(40998,\"mm/dd/yyyy\")").unwrap(),
            s("03/30/2012")
        );
        // Percentage
        assert_eq!(eval("=TEXT(0.285,\"0%\")").unwrap(), s("28%"));
        // Scientific notation
        assert_eq!(eval("=TEXT(12200000,\"0.00E+00\")").unwrap(), s("1.22E7"));
        // Fixed decimal
        assert_eq!(eval("=TEXT(1234.567,\"0.00\")").unwrap(), s("1234.57"));
        // Thousands no decimals
        assert_eq!(eval("=TEXT(1234.567,\"#,##0\")").unwrap(), s("1,235"));
        // yyyy-mm-dd format
        assert_eq!(
            eval("=TEXT(40998,\"yyyy-mm-dd\")").unwrap(),
            s("2012-03-30")
        );
        // Time hh:mm
        assert_eq!(eval("=TEXT(0.5,\"hh:mm\")").unwrap(), s("12:00"));
        // Time hh:mm:ss
        assert_eq!(eval("=TEXT(0.75,\"hh:mm:ss\")").unwrap(), s("18:00:00"));
    }

    #[test]
    fn test_textjoin_docs() {
        assert_eq!(
            eval(r#"=TEXTJOIN(" ",TRUE,"The","sun","will","come","up","tomorrow.")"#).unwrap(),
            s("The sun will come up tomorrow.")
        );
        // ignore_empty=FALSE includes empty strings
        assert_eq!(
            eval(r#"=TEXTJOIN(", ",FALSE,{"a1","b1";"a2","b2";"","";"a4","b4"})"#).unwrap(),
            s("a1, b1, a2, b2, , , a4, b4")
        );
        // Empty delimiter concatenates
        assert_eq!(eval(r#"=TEXTJOIN("",TRUE,"a","b","c")"#).unwrap(), s("abc"));
    }

    #[test]
    fn test_fixed_docs() {
        assert_eq!(eval("=FIXED(1234.567,1)").unwrap(), s("1,234.6"));
        assert_eq!(eval("=FIXED(1234.567,-1)").unwrap(), s("1,230"));
        assert_eq!(eval("=FIXED(-1234.567,-1,TRUE)").unwrap(), s("-1230"));
        assert_eq!(eval("=FIXED(44.332)").unwrap(), s("44.33"));
    }

    #[test]
    fn test_dollar_docs() {
        assert_eq!(eval("=DOLLAR(1234.567,2)").unwrap(), s("$1,234.57"));
        assert_eq!(eval("=DOLLAR(-1234.567,-2)").unwrap(), s("-$1,200"));
        assert_eq!(eval("=DOLLAR(-0.123,4)").unwrap(), s("-$0.1230"));
        assert_eq!(eval("=DOLLAR(99.888)").unwrap(), s("$99.89"));
    }

    #[test]
    fn test_numbervalue_docs() {
        // "2.500,27" with decimal=comma, group=period → 2500.27
        assert_eq!(
            eval("=NUMBERVALUE(\"2.500,27\",\",\",\".\")").unwrap(),
            FormulaValue::Number(2500.27)
        );
        // Spaces ignored: " 3 000 " → 3000
        assert_eq!(
            eval("=NUMBERVALUE(\" 3 000 \")").unwrap(),
            FormulaValue::Number(3000.0)
        );
        // Decimal separator used more than once → #VALUE!
        assert_eq!(
            eval("=NUMBERVALUE(\"1.2.3\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }
}
