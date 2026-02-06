//! Text functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
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

/// LEN(text)
pub fn fn_len(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let v = args.get(0).unwrap();
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
    let text = args.get(0).unwrap();
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
    let text = args.get(0).unwrap();
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
    let text = args.get(0).unwrap();
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
    let text = args.get(0).unwrap();
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
    let text = args.get(0).unwrap();
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
    let text = args.get(0).unwrap();
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
    let find_text = match args.get(0) {
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
    let find_text = match args.get(0) {
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
    let text1 = match args.get(0) {
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
    let text = match args.get(0) {
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
    let text = match args.get(0) {
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
    let text = match args.get(0) {
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
    let number = match args.get(0) {
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
    let text = match args.get(0) {
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
    let text = match args.get(0) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => v.as_string(),
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let cleaned: String = text.chars().filter(|c| *c as u32 >= 32).collect();
    Ok(FormulaValue::String(cleaned))
}

/// VALUE(text) - Converts a text string that represents a number to a number
pub fn fn_value(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match args.get(0) {
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
    match args.get(0) {
        Some(FormulaValue::String(s)) => Ok(FormulaValue::String(s.clone())),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::String(String::new())),
    }
}

/// N(value) - Returns a value converted to a number
pub fn fn_n(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match args.get(0) {
        Some(FormulaValue::Number(n)) => Ok(FormulaValue::Number(*n)),
        Some(FormulaValue::Boolean(true)) => Ok(FormulaValue::Number(1.0)),
        Some(FormulaValue::Boolean(false)) => Ok(FormulaValue::Number(0.0)),
        Some(FormulaValue::Error(e)) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::Number(0.0)),
    }
}
