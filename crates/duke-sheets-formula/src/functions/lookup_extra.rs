//! Additional lookup, reference, web, and date-intl functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue, ImageInfo, ImageSizing};
use chrono::{Datelike, Duration, NaiveDate};
use duke_sheets_core::CellError;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

fn to_i64_trunc(v: &FormulaValue) -> Option<i64> {
    v.as_number().map(|n| n.trunc() as i64)
}

fn scalar_i64(v: &FormulaValue) -> Result<i64, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        _ => to_i64_trunc(v).ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn scalar_bool(v: &FormulaValue) -> Result<bool, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        FormulaValue::Empty => Ok(false),
        _ => v.as_bool().ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn scalar_string(v: &FormulaValue) -> Result<String, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        _ => Ok(v.as_string()),
    }
}

fn as_array(v: &FormulaValue) -> Vec<Vec<FormulaValue>> {
    match v {
        FormulaValue::Array { data: a, .. } => a.clone(),
        _ => vec![vec![v.clone()]],
    }
}

fn array_dims(a: &[Vec<FormulaValue>]) -> (usize, usize) {
    (a.len(), a.first().map(|r| r.len()).unwrap_or(0))
}

#[allow(clippy::needless_range_loop)]
fn flat_vector(a: &[Vec<FormulaValue>], by_col: bool) -> Vec<FormulaValue> {
    let (rows, cols) = array_dims(a);
    let mut out = Vec::new();
    if by_col {
        for c in 0..cols {
            for r in 0..rows {
                out.push(a[r].get(c).cloned().unwrap_or(FormulaValue::Empty));
            }
        }
    } else {
        for row in a {
            for v in row {
                out.push(v.clone());
            }
        }
    }
    out
}

fn compare_lookup_values(a: &FormulaValue, b: &FormulaValue) -> Option<Ordering> {
    match (a, b) {
        (FormulaValue::Number(x), FormulaValue::Number(y)) => x.partial_cmp(y),
        (FormulaValue::Boolean(x), FormulaValue::Boolean(y)) => Some(x.cmp(y)),
        (FormulaValue::String(x), FormulaValue::String(y)) => {
            Some(x.to_ascii_lowercase().cmp(&y.to_ascii_lowercase()))
        }
        (FormulaValue::Number(x), FormulaValue::String(s)) => {
            s.parse::<f64>().ok().and_then(|n| x.partial_cmp(&n))
        }
        (FormulaValue::String(s), FormulaValue::Number(x)) => {
            s.parse::<f64>().ok().and_then(|n| n.partial_cmp(x))
        }
        (FormulaValue::Empty, FormulaValue::Empty) => Some(Ordering::Equal),
        (FormulaValue::Empty, FormulaValue::Number(n)) => 0.0f64.partial_cmp(n),
        (FormulaValue::Number(n), FormulaValue::Empty) => n.partial_cmp(&0.0),
        (FormulaValue::Empty, FormulaValue::String(s)) => {
            Some("".cmp(s.to_ascii_lowercase().as_str()))
        }
        (FormulaValue::String(s), FormulaValue::Empty) => {
            Some(s.to_ascii_lowercase().as_str().cmp(""))
        }
        _ => None,
    }
}

fn to_col_letters(mut col_1based: i64) -> Option<String> {
    if col_1based < 1 {
        return None;
    }
    let mut out = String::new();
    while col_1based > 0 {
        let rem = ((col_1based - 1) % 26) as u8;
        out.insert(0, (b'A' + rem) as char);
        col_1based = (col_1based - 1) / 26;
    }
    Some(out)
}

fn ctx_date_1904(ctx: &EvaluationContext) -> bool {
    ctx.workbook
        .map(|wb| wb.settings().date_1904)
        .unwrap_or(false)
}

fn excel1904_date_from_serial(serial: i64) -> Option<NaiveDate> {
    let base = NaiveDate::from_ymd_opt(1904, 1, 1)?;
    base.checked_add_signed(Duration::days(serial))
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

fn parse_holidays(v: Option<&FormulaValue>) -> Result<HashSet<i64>, FormulaValue> {
    let mut out = HashSet::new();
    let Some(v) = v else {
        return Ok(out);
    };
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { data: rows, .. } => {
            for row in rows {
                for cell in row {
                    if let FormulaValue::Error(e) = cell {
                        return Err(FormulaValue::Error(*e));
                    }
                    if matches!(cell, FormulaValue::Empty) {
                        continue;
                    }
                    let n = cell
                        .as_number()
                        .ok_or(FormulaValue::Error(CellError::Value))?;
                    out.insert(n.floor() as i64);
                }
            }
            Ok(out)
        }
        _ => {
            let n = v.as_number().ok_or(FormulaValue::Error(CellError::Value))?;
            out.insert(n.floor() as i64);
            Ok(out)
        }
    }
}

fn parse_weekend_mask(v: Option<&FormulaValue>) -> Result<[bool; 7], FormulaValue> {
    // Indexes are Sunday=0..Saturday=6.
    fn mask_from_code(code: i64) -> Option<[bool; 7]> {
        let mut m = [false; 7];
        match code {
            1 => {
                m[0] = true;
                m[6] = true;
            }
            2 => {
                m[0] = true;
                m[1] = true;
            }
            3 => {
                m[1] = true;
                m[2] = true;
            }
            4 => {
                m[2] = true;
                m[3] = true;
            }
            5 => {
                m[3] = true;
                m[4] = true;
            }
            6 => {
                m[4] = true;
                m[5] = true;
            }
            7 => {
                m[5] = true;
                m[6] = true;
            }
            11 => m[0] = true,
            12 => m[1] = true,
            13 => m[2] = true,
            14 => m[3] = true,
            15 => m[4] = true,
            16 => m[5] = true,
            17 => m[6] = true,
            _ => return None,
        }
        Some(m)
    }

    let Some(v) = v else {
        return Ok(mask_from_code(1).unwrap());
    };

    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        FormulaValue::String(s) => {
            if s.len() != 7 || !s.chars().all(|c| c == '0' || c == '1') {
                return Err(FormulaValue::Error(CellError::Value));
            }
            let chars: Vec<char> = s.chars().collect();
            let mut m = [false; 7];
            // Weekend string is Monday..Sunday.
            m[1] = chars[0] == '1';
            m[2] = chars[1] == '1';
            m[3] = chars[2] == '1';
            m[4] = chars[3] == '1';
            m[5] = chars[4] == '1';
            m[6] = chars[5] == '1';
            m[0] = chars[6] == '1';
            Ok(m)
        }
        _ => {
            let code = to_i64_trunc(v).ok_or(FormulaValue::Error(CellError::Value))?;
            mask_from_code(code).ok_or(FormulaValue::Error(CellError::Num))
        }
    }
}

fn is_weekend_intl(
    serial: i64,
    weekend_mask: &[bool; 7],
    ctx: &EvaluationContext,
) -> Result<bool, FormulaValue> {
    let dow = weekday_sun0_from_serial(serial, ctx)? as usize;
    Ok(weekend_mask[dow])
}

/// ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])
pub fn fn_address(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let row = match scalar_i64(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let col = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if row < 1 || col < 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let abs_num = match args.get(2) {
        Some(v) => match scalar_i64(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => 1,
    };
    if !(1..=4).contains(&abs_num) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let a1 = match args.get(3) {
        Some(v) => match scalar_bool(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => true,
    };

    let sheet_text = match args.get(4) {
        Some(v) => match scalar_string(v) {
            Ok(s) => Some(s),
            Err(e) => return Ok(e),
        },
        None => None,
    };

    let addr = if a1 {
        let col_letters = match to_col_letters(col) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        };
        match abs_num {
            1 => format!("${}${}", col_letters, row),
            2 => format!("{}${}", col_letters, row),
            3 => format!("${}{}", col_letters, row),
            4 => format!("{}{}", col_letters, row),
            _ => unreachable!(),
        }
    } else {
        // R1C1 style: abs_num controls absolute vs relative for row/col
        // 1 = both absolute: R2C3
        // 2 = absolute row, relative col: R2C[3]
        // 3 = relative row, absolute col: R[2]C3
        // 4 = both relative: R[2]C[3]
        let row_part = match abs_num {
            1 | 2 => format!("R{}", row),
            3 | 4 => format!("R[{}]", row),
            _ => unreachable!(),
        };
        let col_part = match abs_num {
            1 | 3 => format!("C{}", col),
            2 | 4 => format!("C[{}]", col),
            _ => unreachable!(),
        };
        format!("{}{}", row_part, col_part)
    };

    let final_addr = match sheet_text {
        Some(s) if !s.is_empty() => {
            // Wrap sheet name in single quotes if it contains special chars,
            // brackets, spaces, or is an external reference.
            let needs_quotes = s.contains(' ')
                || s.contains('[')
                || s.contains(']')
                || s.contains('!')
                || s.contains('\'');
            if needs_quotes {
                format!("'{}'!{}", s, addr)
            } else {
                format!("{}!{}", s, addr)
            }
        }
        _ => addr,
    };

    Ok(FormulaValue::String(final_addr))
}

/// AREAS(reference)
///
/// Returns the number of areas in a reference.
/// The real implementation lives in `evaluate_areas` in the evaluator,
/// which intercepts the call before arguments are evaluated so it can
/// count union branches in the raw AST.  This fallback handles the
/// pre-evaluated path (direct calls, or if the special case is bypassed).
pub fn fn_areas(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if let Some(FormulaValue::Error(e)) = args.first() {
        return Ok(FormulaValue::Error(*e));
    }
    Ok(FormulaValue::Number(1.0))
}

/// CHOOSECOLS(array, col_num1, ...)
pub fn fn_choosecols(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let array = as_array(args.first().unwrap());
    let (rows, cols) = array_dims(&array);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let mut selected = Vec::new();
    for v in &args[1..] {
        let c = match scalar_i64(v) {
            Ok(n) => n,
            Err(e) => return Ok(e),
        };
        if c == 0 {
            return Ok(FormulaValue::Error(CellError::Value));
        }
        let idx = if c > 0 { c - 1 } else { cols as i64 + c };
        if idx < 0 || idx >= cols as i64 {
            return Ok(FormulaValue::Error(CellError::Value));
        }
        selected.push(idx as usize);
    }

    let mut out = Vec::with_capacity(rows);
    for row in &array {
        let mut new_row = Vec::with_capacity(selected.len());
        for &idx in &selected {
            new_row.push(row[idx].clone());
        }
        out.push(new_row);
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// CHOOSEROWS(array, row_num1, ...)
pub fn fn_chooserows(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let array = as_array(args.first().unwrap());
    let (rows, _cols) = array_dims(&array);
    if rows == 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let mut out = Vec::new();
    for v in &args[1..] {
        let r = match scalar_i64(v) {
            Ok(n) => n,
            Err(e) => return Ok(e),
        };
        if r == 0 {
            return Ok(FormulaValue::Error(CellError::Value));
        }
        let idx = if r > 0 { r - 1 } else { rows as i64 + r };
        if idx < 0 || idx >= rows as i64 {
            return Ok(FormulaValue::Error(CellError::Value));
        }
        out.push(array[idx as usize].clone());
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// DROP(array, rows, [columns])
pub fn fn_drop(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let array = as_array(args.first().unwrap());
    let (rows, cols) = array_dims(&array);
    let drop_rows = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let drop_cols = match args.get(2) {
        Some(v) => match scalar_i64(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => 0,
    };

    if drop_rows.unsigned_abs() as usize >= rows || drop_cols.unsigned_abs() as usize >= cols {
        return Ok(FormulaValue::Error(CellError::Calc));
    }

    let row_start = if drop_rows >= 0 {
        drop_rows as usize
    } else {
        0
    };
    let row_end = if drop_rows >= 0 {
        rows
    } else {
        rows - (-drop_rows as usize)
    };

    let col_start = if drop_cols >= 0 {
        drop_cols as usize
    } else {
        0
    };
    let col_end = if drop_cols >= 0 {
        cols
    } else {
        cols - (-drop_cols as usize)
    };

    let mut out = Vec::new();
    for row in &array[row_start..row_end] {
        out.push(row[col_start..col_end].to_vec());
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// EXPAND(array, rows, [columns], [pad_with])
pub fn fn_expand(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let array = as_array(args.first().unwrap());
    let (src_rows, src_cols) = array_dims(&array);

    let target_rows = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) if v > 0 => v as usize,
        Ok(_) => return Ok(FormulaValue::Error(CellError::Value)),
        Err(e) => return Ok(e),
    };
    let target_cols = match args.get(2) {
        Some(v) => match scalar_i64(v) {
            Ok(v) if v > 0 => v as usize,
            Ok(_) => return Ok(FormulaValue::Error(CellError::Value)),
            Err(e) => return Ok(e),
        },
        None => src_cols,
    };

    if target_rows < src_rows || target_cols < src_cols {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let pad = args
        .get(3)
        .cloned()
        .unwrap_or(FormulaValue::Error(CellError::Na));
    let mut out = vec![vec![pad.clone(); target_cols]; target_rows];
    for r in 0..src_rows {
        for c in 0..src_cols {
            out[r][c] = array[r][c].clone();
        }
    }
    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// FILTER(array, include, [if_empty])
pub fn fn_filter(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let array = as_array(args.first().unwrap());
    let include_arr = as_array(args.get(1).unwrap());
    let (rows, cols) = array_dims(&array);
    let (inc_rows, inc_cols) = array_dims(&include_arr);

    // Determine filter direction and extract include vector
    let is_row_filter;
    let include_vec;

    if inc_rows == rows && inc_cols == 1 {
        // Column vector matching rows → row filter
        is_row_filter = true;
        include_vec = include_arr
            .iter()
            .map(|r| r.first().cloned().unwrap_or(FormulaValue::Empty))
            .collect::<Vec<_>>();
    } else if inc_rows == 1 && inc_cols == rows {
        // Row vector matching rows → row filter
        is_row_filter = true;
        include_vec = include_arr[0].clone();
    } else if inc_rows == 1 && inc_cols == cols && cols != rows {
        // Row vector matching cols (unambiguous) → column filter
        is_row_filter = false;
        include_vec = include_arr[0].clone();
    } else if inc_rows == cols && inc_cols == 1 && cols != rows {
        // Column vector matching cols (unambiguous) → column filter
        is_row_filter = false;
        include_vec = include_arr
            .iter()
            .map(|r| r.first().cloned().unwrap_or(FormulaValue::Empty))
            .collect::<Vec<_>>();
    } else {
        return Ok(FormulaValue::Error(CellError::Value));
    };

    if is_row_filter {
        let mut out = Vec::new();
        for (i, row) in array.iter().enumerate() {
            if let FormulaValue::Error(e) = &include_vec[i] {
                return Ok(FormulaValue::Error(*e));
            }
            let keep = include_vec[i]
                .as_bool()
                .ok_or(FormulaValue::Error(CellError::Value));
            match keep {
                Ok(true) => out.push(row.clone()),
                Ok(false) => {}
                Err(e) => return Ok(e),
            }
        }
        if out.is_empty() {
            return Ok(args
                .get(2)
                .cloned()
                .unwrap_or(FormulaValue::Error(CellError::Calc)));
        }
        Ok(FormulaValue::Array {
            data: out,
            source: None,
        })
    } else {
        // Column filter: keep only columns where include is TRUE
        let mut col_indices = Vec::new();
        for (j, val) in include_vec.iter().enumerate() {
            if let FormulaValue::Error(e) = val {
                return Ok(FormulaValue::Error(*e));
            }
            let keep = val.as_bool().ok_or(FormulaValue::Error(CellError::Value));
            match keep {
                Ok(true) => col_indices.push(j),
                Ok(false) => {}
                Err(e) => return Ok(e),
            }
        }
        if col_indices.is_empty() {
            return Ok(args
                .get(2)
                .cloned()
                .unwrap_or(FormulaValue::Error(CellError::Calc)));
        }
        let mut out = Vec::new();
        for row in &array {
            let filtered_row: Vec<FormulaValue> = col_indices
                .iter()
                .map(|&j| row.get(j).cloned().unwrap_or(FormulaValue::Empty))
                .collect();
            out.push(filtered_row);
        }
        Ok(FormulaValue::Array {
            data: out,
            source: None,
        })
    }
}

/// FORMULATEXT(reference)
///
/// Returns the formula of a referenced cell as a string.
/// The real implementation lives in `evaluate_formulatext` in the evaluator,
/// which intercepts the call before arguments are evaluated so it can inspect
/// the raw cell reference.  This fallback handles the pre-evaluated path
/// (direct calls, or if the special case is ever bypassed).
pub fn fn_formulatext(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    // Propagate errors from evaluated argument
    if let Some(FormulaValue::Error(e)) = args.first() {
        return Ok(FormulaValue::Error(*e));
    }
    // With pre-evaluated args the original reference is lost — return #N/A.
    Ok(FormulaValue::Error(CellError::Na))
}

/// HSTACK(array1, array2, ...)
pub fn fn_hstack(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arrays: Vec<Vec<Vec<FormulaValue>>> = args.iter().map(as_array).collect();
    let max_rows = arrays.iter().map(|a| a.len()).max().unwrap_or(0);
    let mut out = Vec::with_capacity(max_rows);

    for r in 0..max_rows {
        let mut row = Vec::new();
        for arr in &arrays {
            let cols = arr.first().map(|x| x.len()).unwrap_or(0);
            if r < arr.len() {
                row.extend(arr[r].iter().cloned());
            } else {
                row.extend(std::iter::repeat_n(
                    FormulaValue::Error(CellError::Na),
                    cols,
                ));
            }
        }
        out.push(row);
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// LOOKUP(lookup_value, lookup_vector, [result_vector])
pub fn fn_lookup(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let lookup_value = args.first().unwrap().clone();
    let lookup_array = as_array(args.get(1).unwrap());
    let lookup_vec = flat_vector(&lookup_array, false);

    let result_vec = match args.get(2) {
        Some(v) => flat_vector(&as_array(v), false),
        None => lookup_vec.clone(),
    };

    if lookup_vec.len() != result_vec.len() || lookup_vec.is_empty() {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    let mut best: Option<usize> = None;
    for (i, item) in lookup_vec.iter().enumerate() {
        if let Some(ord) = compare_lookup_values(item, &lookup_value) {
            if ord == Ordering::Equal || ord == Ordering::Less {
                best = Some(i);
            }
        }
    }

    match best {
        Some(i) => Ok(result_vec[i].clone()),
        None => Ok(FormulaValue::Error(CellError::Na)),
    }
}

/// SORT(array, [sort_index], [sort_order], [by_col])
pub fn fn_sort(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut arr = as_array(args.first().unwrap());
    let sort_index = match args.get(1) {
        Some(v) => match scalar_i64(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => 1,
    };
    let sort_order = match args.get(2) {
        Some(v) => match scalar_i64(v) {
            Ok(v) if v == 1 || v == -1 => v,
            Ok(_) => return Ok(FormulaValue::Error(CellError::Value)),
            Err(e) => return Ok(e),
        },
        None => 1,
    };
    let by_col = match args.get(3) {
        Some(v) => match scalar_bool(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => false,
    };

    if by_col {
        let t = fn_transpose(
            &[FormulaValue::Array {
                data: arr,
                source: None,
            }],
            _ctx,
        )?;
        if let FormulaValue::Array { data: mut tr, .. } = t {
            let idx = (sort_index - 1) as usize;
            if tr.is_empty() || idx >= tr[0].len() {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            tr.sort_by(|a, b| compare_lookup_values(&a[idx], &b[idx]).unwrap_or(Ordering::Equal));
            if sort_order == -1 {
                tr.reverse();
            }
            return fn_transpose(
                &[FormulaValue::Array {
                    data: tr,
                    source: None,
                }],
                _ctx,
            );
        }
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let idx = (sort_index - 1) as usize;
    if arr.is_empty() || idx >= arr[0].len() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    arr.sort_by(|a, b| compare_lookup_values(&a[idx], &b[idx]).unwrap_or(Ordering::Equal));
    if sort_order == -1 {
        arr.reverse();
    }
    Ok(FormulaValue::Array {
        data: arr,
        source: None,
    })
}

/// SORTBY(array, by_array1, [sort_order1], ...)
pub fn fn_sortby(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.first().unwrap());
    if arr.is_empty() {
        return Ok(FormulaValue::Array {
            data: arr,
            source: None,
        });
    }

    let mut keys: Vec<(Vec<FormulaValue>, i64)> = Vec::new();
    let mut i = 1;
    while i < args.len() {
        let by = flat_vector(&as_array(args.get(i).unwrap()), false);
        let order = if i + 1 < args.len() {
            match scalar_i64(args.get(i + 1).unwrap()) {
                Ok(v) if v == 1 || v == -1 => v,
                Ok(_) => return Ok(FormulaValue::Error(CellError::Value)),
                Err(e) => return Ok(e),
            }
        } else {
            1
        };
        if by.len() != arr.len() {
            return Ok(FormulaValue::Error(CellError::Value));
        }
        keys.push((by, order));
        i += 2;
    }

    let mut idxs: Vec<usize> = (0..arr.len()).collect();
    idxs.sort_by(|&a, &b| {
        for (key, order) in &keys {
            let ord = compare_lookup_values(&key[a], &key[b]).unwrap_or(Ordering::Equal);
            if ord != Ordering::Equal {
                return if *order == -1 { ord.reverse() } else { ord };
            }
        }
        Ordering::Equal
    });

    let sorted = idxs.into_iter().map(|ix| arr[ix].clone()).collect();
    Ok(FormulaValue::Array {
        data: sorted,
        source: None,
    })
}

/// TAKE(array, rows, [columns])
pub fn fn_take(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.first().unwrap());
    let (rows, cols) = array_dims(&arr);
    let take_rows = match args.get(1) {
        Some(FormulaValue::Empty) | None => rows as i64,
        Some(v) => match scalar_i64(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
    };
    let take_cols = match args.get(2) {
        Some(FormulaValue::Empty) | None => cols as i64,
        Some(v) => match scalar_i64(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
    };

    if take_rows == 0 || take_cols == 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let rs = take_rows.unsigned_abs() as usize;
    let cs = take_cols.unsigned_abs() as usize;

    let row_slice = if take_rows > 0 {
        (0, rs.min(rows))
    } else {
        (rows.saturating_sub(rs.min(rows)), rows)
    };
    let col_slice = if take_cols > 0 {
        (0, cs.min(cols))
    } else {
        (cols.saturating_sub(cs.min(cols)), cols)
    };

    let mut out = Vec::new();
    for row in &arr[row_slice.0..row_slice.1] {
        out.push(row[col_slice.0..col_slice.1].to_vec());
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

fn flatten_with_ignore(arr: &[Vec<FormulaValue>], ignore: i64, by_col: bool) -> Vec<FormulaValue> {
    let values = flat_vector(arr, by_col);
    values
        .into_iter()
        .filter(|v| match ignore {
            1 => !matches!(v, FormulaValue::Empty),
            2 => !matches!(v, FormulaValue::Error(_)),
            3 => !matches!(v, FormulaValue::Empty | FormulaValue::Error(_)),
            _ => true,
        })
        .collect()
}

/// TOCOL(array, [ignore], [scan_by_column])
pub fn fn_tocol(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.first().unwrap());
    let ignore = match args.get(1) {
        Some(v) => match scalar_i64(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => 0,
    };
    let by_col = match args.get(2) {
        Some(v) => match scalar_bool(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => false,
    };

    let out = flatten_with_ignore(&arr, ignore, by_col)
        .into_iter()
        .map(|v| vec![v])
        .collect();
    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// TOROW(array, [ignore], [scan_by_column])
pub fn fn_torow(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.first().unwrap());
    let ignore = match args.get(1) {
        Some(v) => match scalar_i64(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => 0,
    };
    let by_col = match args.get(2) {
        Some(v) => match scalar_bool(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => false,
    };

    Ok(FormulaValue::Array {
        data: vec![flatten_with_ignore(&arr, ignore, by_col)],
        source: None,
    })
}

/// TRANSPOSE(array)
pub fn fn_transpose(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.first().unwrap());
    let (rows, cols) = array_dims(&arr);
    let mut out = vec![vec![FormulaValue::Empty; rows]; cols];
    for (r, row) in arr.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            out[c][r] = v.clone();
        }
    }
    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// UNIQUE(array, [by_col], [exactly_once])
pub fn fn_unique(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.first().unwrap());
    let by_col = match args.get(1) {
        Some(v) => match scalar_bool(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => false,
    };
    let exactly_once = match args.get(2) {
        Some(v) => match scalar_bool(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => false,
    };

    if by_col {
        let t = fn_transpose(
            &[FormulaValue::Array {
                data: arr,
                source: None,
            }],
            _ctx,
        )?;
        if let FormulaValue::Array {
            data: transposed, ..
        } = t
        {
            let unique_rows = fn_unique(
                &[
                    FormulaValue::Array {
                        data: transposed,
                        source: None,
                    },
                    FormulaValue::Boolean(false),
                    FormulaValue::Boolean(exactly_once),
                ],
                _ctx,
            )?;
            return fn_transpose(&[unique_rows], _ctx);
        }
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let mut counts: HashMap<String, usize> = HashMap::new();
    for row in &arr {
        let k = format!("{:?}", row);
        *counts.entry(k).or_insert(0) += 1;
    }

    let mut seen = HashSet::new();
    let mut out = Vec::new();
    for row in arr {
        let k = format!("{:?}", row);
        if exactly_once {
            if counts.get(&k).copied().unwrap_or(0) == 1 {
                out.push(row);
            }
        } else if seen.insert(k) {
            out.push(row);
        }
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// VSTACK(array1, array2, ...)
pub fn fn_vstack(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arrays: Vec<Vec<Vec<FormulaValue>>> = args.iter().map(as_array).collect();
    let max_cols = arrays
        .iter()
        .map(|a| a.first().map(|r| r.len()).unwrap_or(0))
        .max()
        .unwrap_or(0);

    let mut out = Vec::new();
    for arr in arrays {
        for row in arr {
            let mut r = row.clone();
            while r.len() < max_cols {
                r.push(FormulaValue::Error(CellError::Na));
            }
            out.push(r);
        }
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// WRAPCOLS(vector, wrap_count, [pad_with])
pub fn fn_wrapcols(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let vec_values = flat_vector(&as_array(args.first().unwrap()), false);
    let wrap_count = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) if v > 0 => v as usize,
        Ok(_) => return Ok(FormulaValue::Error(CellError::Value)),
        Err(e) => return Ok(e),
    };
    let pad = args
        .get(2)
        .cloned()
        .unwrap_or(FormulaValue::Error(CellError::Na));

    let cols = vec_values.len().div_ceil(wrap_count);
    let mut out = vec![vec![pad.clone(); cols]; wrap_count];

    for (i, v) in vec_values.into_iter().enumerate() {
        let col = i / wrap_count;
        let row = i % wrap_count;
        out[row][col] = v;
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// WRAPROWS(vector, wrap_count, [pad_with])
pub fn fn_wraprows(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let vec_values = flat_vector(&as_array(args.first().unwrap()), false);
    let wrap_count = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) if v > 0 => v as usize,
        Ok(_) => return Ok(FormulaValue::Error(CellError::Value)),
        Err(e) => return Ok(e),
    };
    let pad = args
        .get(2)
        .cloned()
        .unwrap_or(FormulaValue::Error(CellError::Na));

    let rows = vec_values.len().div_ceil(wrap_count);
    let mut out = vec![vec![pad.clone(); wrap_count]; rows];

    for (i, v) in vec_values.into_iter().enumerate() {
        let row = i / wrap_count;
        let col = i % wrap_count;
        out[row][col] = v;
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

/// HYPERLINK(link_location, [friendly_name])
pub fn fn_hyperlink(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let link = match scalar_string(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let display = match args.get(1) {
        Some(v) => match scalar_string(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => link,
    };

    Ok(FormulaValue::String(display))
}

/// NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])
pub fn fn_networkdays_intl(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let start = match scalar_i64(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let end = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let weekend_mask = match parse_weekend_mask(args.get(2)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let holidays = match parse_holidays(args.get(3)) {
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
        let weekend = match is_weekend_intl(serial, &weekend_mask, _ctx) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        if !weekend && !holidays.contains(&serial) {
            count += 1;
        }
    }

    Ok(FormulaValue::Number(sign * count as f64))
}

/// WORKDAY.INTL(start_date, days, [weekend], [holidays])
pub fn fn_workday_intl(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let start = match scalar_i64(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let days = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let weekend_mask = match parse_weekend_mask(args.get(2)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let holidays = match parse_holidays(args.get(3)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if days == 0 {
        return Ok(FormulaValue::Number(start as f64));
    }

    let mut current = start;
    let mut remaining = days.abs();
    let step = if days > 0 { 1 } else { -1 };

    while remaining > 0 {
        current += step;
        let weekend = match is_weekend_intl(current, &weekend_mask, _ctx) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        };
        if !weekend && !holidays.contains(&current) {
            remaining -= 1;
        }
    }

    Ok(FormulaValue::Number(current as f64))
}

/// ENCODEURL(text)
pub fn fn_encodeurl(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let text = match scalar_string(args.first().unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let mut out = String::new();
    for b in text.bytes() {
        if b.is_ascii_alphanumeric() || b"-_.~".contains(&b) {
            out.push(b as char);
        } else {
            out.push_str(&format!("%{:02X}", b));
        }
    }

    Ok(FormulaValue::String(out))
}

#[derive(Debug, Clone)]
enum XmlChild {
    Element(XmlNode),
    Text(String),
}

#[derive(Debug, Clone)]
struct XmlNode {
    name: String,
    attributes: HashMap<String, String>,
    children: Vec<XmlChild>,
}

impl XmlNode {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        let mut attributes = HashMap::new();
        for attr in e.attributes().flatten() {
            let value = attr.unescape_value().ok()?.into_owned();
            attributes.insert(
                String::from_utf8_lossy(attr.key.as_ref()).into_owned(),
                value,
            );
        }
        Some(Self {
            name: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
            attributes,
            children: Vec::new(),
        })
    }

    fn child_elements_named(&self, name: &str) -> Vec<&XmlNode> {
        self.children
            .iter()
            .filter_map(|child| match child {
                XmlChild::Element(node) if node.name == name => Some(node),
                _ => None,
            })
            .collect()
    }

    fn descendant_elements_named<'a>(&'a self, name: &str, out: &mut Vec<&'a XmlNode>) {
        if self.name == name {
            out.push(self);
        }
        for child in &self.children {
            if let XmlChild::Element(node) = child {
                node.descendant_elements_named(name, out);
            }
        }
    }

    fn text_content(&self) -> String {
        let mut out = String::new();
        for child in &self.children {
            match child {
                XmlChild::Element(node) => out.push_str(&node.text_content()),
                XmlChild::Text(text) => out.push_str(text),
            }
        }
        out
    }
}

#[derive(Debug, Clone)]
enum XPathStep {
    Root { name: String, index: Option<usize> },
    Child { name: String, index: Option<usize> },
    Descendant { name: String, index: Option<usize> },
    Attribute(String),
}

fn parse_xpath_segment(segment: &str) -> Option<(String, Option<usize>)> {
    let Some((name, index)) = segment.split_once('[') else {
        if segment.is_empty() {
            return None;
        }
        return Some((segment.to_string(), None));
    };
    let index = index.strip_suffix(']')?;
    let index = index.parse::<usize>().ok()?;
    if name.is_empty() || index == 0 {
        return None;
    }
    Some((name.to_string(), Some(index)))
}

fn parse_xpath(xpath: &str) -> Option<Vec<XPathStep>> {
    if xpath.is_empty() || !xpath.starts_with('/') {
        return None;
    }

    let bytes = xpath.as_bytes();
    let mut i = 0;
    let mut first = true;
    let mut steps = Vec::new();

    while i < bytes.len() {
        if bytes[i] != b'/' {
            return None;
        }

        let descendant = i + 1 < bytes.len() && bytes[i + 1] == b'/';
        i += if descendant { 2 } else { 1 };
        if i >= bytes.len() {
            return None;
        }

        let start = i;
        while i < bytes.len() && bytes[i] != b'/' {
            i += 1;
        }

        let segment = &xpath[start..i];
        if let Some(attribute) = segment.strip_prefix('@') {
            if attribute.is_empty() || i != bytes.len() {
                return None;
            }
            steps.push(XPathStep::Attribute(attribute.to_string()));
            return Some(steps);
        }

        let (name, index) = parse_xpath_segment(segment)?;
        steps.push(if descendant {
            XPathStep::Descendant { name, index }
        } else if first {
            XPathStep::Root { name, index }
        } else {
            XPathStep::Child { name, index }
        });
        first = false;
    }

    Some(steps)
}

fn append_xml_node(node: XmlNode, stack: &mut [XmlNode], root: &mut Option<XmlNode>) -> Option<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlChild::Element(node));
    } else if root.is_none() {
        *root = Some(node);
    } else {
        return None;
    }
    Some(())
}

fn parse_xml_document(xml: &str) -> Option<XmlNode> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut stack = Vec::new();
    let mut root = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => stack.push(XmlNode::from_start(&e)?),
            Ok(Event::Empty(e)) => {
                append_xml_node(XmlNode::from_start(&e)?, &mut stack, &mut root)?
            }
            Ok(Event::End(_)) => append_xml_node(stack.pop()?, &mut stack, &mut root)?,
            Ok(Event::Text(e)) => {
                let text = e.unescape().ok()?.into_owned();
                if let Some(node) = stack.last_mut() {
                    if !text.trim().is_empty() {
                        node.children.push(XmlChild::Text(text));
                    }
                } else if !text.trim().is_empty() {
                    return None;
                }
            }
            Ok(Event::CData(e)) => {
                let text = String::from_utf8_lossy(e.as_ref()).into_owned();
                if let Some(node) = stack.last_mut() {
                    if !text.trim().is_empty() {
                        node.children.push(XmlChild::Text(text));
                    }
                } else if !text.trim().is_empty() {
                    return None;
                }
            }
            Ok(Event::Comment(_))
            | Ok(Event::Decl(_))
            | Ok(Event::DocType(_))
            | Ok(Event::PI(_)) => {}
            Ok(Event::Eof) => break,
            Err(_) => return None,
        }
        buf.clear();
    }

    if !stack.is_empty() {
        return None;
    }

    root
}

fn pick_index(nodes: Vec<&XmlNode>, index: Option<usize>) -> Vec<&XmlNode> {
    match index {
        Some(index) => nodes
            .get(index.saturating_sub(1))
            .copied()
            .into_iter()
            .collect(),
        None => nodes,
    }
}

fn evaluate_xpath(root: &XmlNode, steps: &[XPathStep]) -> Option<String> {
    let mut current: Vec<&XmlNode> = Vec::new();

    for step in steps {
        match step {
            XPathStep::Root { name, index } => {
                if root.name != *name || index.unwrap_or(1) != 1 {
                    return None;
                }
                current.clear();
                current.push(root);
            }
            XPathStep::Child { name, index } => {
                let mut next = Vec::new();
                for node in &current {
                    next.extend(pick_index(node.child_elements_named(name), *index));
                }
                if next.is_empty() {
                    return None;
                }
                current = next;
            }
            XPathStep::Descendant { name, index } => {
                let mut next = Vec::new();
                let sources = if current.is_empty() {
                    vec![root]
                } else {
                    current.clone()
                };
                for node in sources {
                    let mut matches = Vec::new();
                    node.descendant_elements_named(name, &mut matches);
                    next.extend(pick_index(matches, *index));
                }
                if next.is_empty() {
                    return None;
                }
                current = next;
            }
            XPathStep::Attribute(name) => {
                for node in &current {
                    if let Some(value) = node.attributes.get(name) {
                        return Some(value.clone());
                    }
                }
                return None;
            }
        }
    }

    current.first().map(|node| node.text_content())
}

fn xml_result_to_value(text: String) -> FormulaValue {
    let trimmed = text.trim();
    if let Ok(number) = trimmed.parse::<f64>() {
        FormulaValue::Number(number)
    } else {
        FormulaValue::String(text)
    }
}

/// FILTERXML(xml, xpath)
pub fn fn_filterxml(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let Some(xml) = args.first() else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let xml = match scalar_string(xml) {
        Ok(value) => value,
        Err(err) => return Ok(err),
    };
    let xpath = match args.get(1) {
        Some(value) => match scalar_string(value) {
            Ok(value) => value,
            Err(err) => return Ok(err),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let Some(document) = parse_xml_document(&xml) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let Some(steps) = parse_xpath(&xpath) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let Some(result) = evaluate_xpath(&document, &steps) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };

    Ok(xml_result_to_value(result))
}

/// WEBSERVICE(url)
pub fn fn_webservice(
    args: &[FormulaValue],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let Some(url) = args.first() else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let url = match scalar_string(url) {
        Ok(value) => value,
        Err(err) => return Ok(err),
    };

    match ctx.web_service_fn {
        Some(handler) => match handler(&url) {
            Some(body) => Ok(FormulaValue::String(body)),
            None => Ok(FormulaValue::Error(CellError::Na)),
        },
        None => Ok(FormulaValue::Error(CellError::Na)),
    }
}

/// GETPIVOTDATA(...)
pub fn fn_getpivotdata(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }
    Ok(FormulaValue::Error(CellError::Na))
}

/// RTD(...)
pub fn fn_rtd(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let Some(prog_id) = args.first() else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let prog_id = match scalar_string(prog_id) {
        Ok(value) => value,
        Err(err) => return Ok(err),
    };
    let server = match args.get(1) {
        Some(value) => match scalar_string(value) {
            Ok(value) => value,
            Err(err) => return Ok(err),
        },
        None => String::new(),
    };

    let mut topics = Vec::with_capacity(args.len().saturating_sub(2));
    for value in args.iter().skip(2) {
        match scalar_string(value) {
            Ok(topic) => topics.push(topic),
            Err(err) => return Ok(err),
        }
    }

    match ctx.rtd_fn {
        Some(handler) => match handler(&prog_id, &server, &topics) {
            Some(value) => Ok(FormulaValue::String(value)),
            None => Ok(FormulaValue::Error(CellError::Na)),
        },
        None => Ok(FormulaValue::Error(CellError::Na)),
    }
}

/// IMAGE(...)
pub fn fn_image(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let Some(source) = args.first() else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let source = match scalar_string(source) {
        Ok(value) => value,
        Err(err) => return Ok(err),
    };
    let alt_text = match args.get(1) {
        Some(value) => match scalar_string(value) {
            Ok(value) => value,
            Err(err) => return Ok(err),
        },
        None => String::new(),
    };
    let sizing = match args.get(2) {
        Some(FormulaValue::Empty) | None => ImageSizing::FitCell,
        Some(value) => match value.as_number() {
            Some(number) if number.fract() == 0.0 => match number as i64 {
                0 => ImageSizing::FitCell,
                1 => ImageSizing::FillCell,
                2 => ImageSizing::OriginalSize,
                3 => ImageSizing::Custom,
                _ => return Ok(FormulaValue::Error(CellError::Value)),
            },
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        },
    };
    let height = match args.get(3) {
        Some(FormulaValue::Empty) | None => None,
        Some(value) => match value.as_number() {
            Some(number) => Some(number),
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
    };
    let width = match args.get(4) {
        Some(FormulaValue::Empty) | None => None,
        Some(value) => match value.as_number() {
            Some(number) => Some(number),
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
    };

    if let Some(image_sink) = ctx.image_sink {
        image_sink(
            ctx.current_sheet,
            ctx.current_row,
            ctx.current_col,
            ImageInfo {
                source: source.clone(),
                alt_text: alt_text.clone(),
                sizing,
                width,
                height,
            },
        );
    }

    let display = if alt_text.is_empty() {
        source
    } else {
        alt_text
    };
    Ok(FormulaValue::String(display))
}

/// Helper: check if a cell is considered "blank" for TRIMRANGE
fn is_blank_cell(v: &FormulaValue) -> bool {
    matches!(v, FormulaValue::Empty) || matches!(v, FormulaValue::String(s) if s.is_empty())
}

/// TRIMRANGE(range, [trim_rows], [trim_cols])
pub fn fn_trimrange(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let array = as_array(args.first().unwrap());
    let (rows, cols) = array_dims(&array);

    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Error(CellError::Calc));
    }

    let trim_rows = match args.get(1) {
        Some(v) => match scalar_i64(v) {
            Ok(v) if (0..=3).contains(&v) => v,
            Ok(_) => return Ok(FormulaValue::Error(CellError::Value)),
            Err(e) => return Ok(e),
        },
        None => 3,
    };
    let trim_cols = match args.get(2) {
        Some(v) => match scalar_i64(v) {
            Ok(v) if (0..=3).contains(&v) => v,
            Ok(_) => return Ok(FormulaValue::Error(CellError::Value)),
            Err(e) => return Ok(e),
        },
        None => 3,
    };

    // Determine row range to keep
    let mut row_start = 0;
    let mut row_end = rows;

    // Trim leading rows (trim_rows == 1 or 3)
    if trim_rows == 1 || trim_rows == 3 {
        while row_start < row_end && array[row_start].iter().all(is_blank_cell) {
            row_start += 1;
        }
    }

    // Trim trailing rows (trim_rows == 2 or 3)
    if trim_rows == 2 || trim_rows == 3 {
        while row_end > row_start && array[row_end - 1].iter().all(is_blank_cell) {
            row_end -= 1;
        }
    }

    // Determine column range to keep (check only remaining rows)
    let mut col_start = 0;
    let mut col_end = cols;

    // Trim leading cols (trim_cols == 1 or 3)
    if trim_cols == 1 || trim_cols == 3 {
        while col_start < col_end {
            let all_blank = (row_start..row_end)
                .all(|r| array[r].get(col_start).map(is_blank_cell).unwrap_or(true));
            if all_blank {
                col_start += 1;
            } else {
                break;
            }
        }
    }

    // Trim trailing cols (trim_cols == 2 or 3)
    if trim_cols == 2 || trim_cols == 3 {
        while col_end > col_start {
            let all_blank = (row_start..row_end)
                .all(|r| array[r].get(col_end - 1).map(is_blank_cell).unwrap_or(true));
            if all_blank {
                col_end -= 1;
            } else {
                break;
            }
        }
    }

    // If everything was trimmed, return #CALC!
    if row_start >= row_end || col_start >= col_end {
        return Ok(FormulaValue::Error(CellError::Calc));
    }

    // Build the result array
    let mut out = Vec::new();
    for row in &array[row_start..row_end] {
        out.push(row[col_start..col_end].to_vec());
    }

    Ok(FormulaValue::Array {
        data: out,
        source: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    fn eval_with_ctx(formula: &str, ctx: &EvaluationContext<'_>) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, ctx)
    }

    fn n(v: f64) -> FormulaValue {
        FormulaValue::Number(v)
    }

    fn s(v: &str) -> FormulaValue {
        FormulaValue::String(v.to_string())
    }

    #[test]
    fn test_address_areas_hyperlink() {
        assert_eq!(
            fn_address(&[n(2.0), n(3.0)], &EvaluationContext::simple()).unwrap(),
            s("$C$2")
        );
        assert_eq!(
            fn_address(
                &[
                    n(2.0),
                    n(3.0),
                    n(4.0),
                    FormulaValue::Boolean(true),
                    s("Sheet1")
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("Sheet1!C2")
        );
        assert_eq!(
            fn_areas(&[s("A1")], &EvaluationContext::simple()).unwrap(),
            n(1.0)
        );
        assert_eq!(
            fn_hyperlink(&[s("https://x"), s("click")], &EvaluationContext::simple()).unwrap(),
            s("click")
        );
    }

    #[test]
    fn test_choosecols_rows_take_drop_expand() {
        let arr = FormulaValue::Array {
            data: vec![vec![n(1.0), n(2.0)], vec![n(3.0), n(4.0)]],
            source: None,
        };
        assert_eq!(
            fn_choosecols(&[arr.clone(), n(2.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(2.0)], vec![n(4.0)]],
                source: None
            }
        );
        assert_eq!(
            fn_chooserows(&[arr.clone(), n(-1.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(3.0), n(4.0)]],
                source: None
            }
        );
        assert_eq!(
            fn_take(&[arr.clone(), n(1.0), n(1.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)]],
                source: None
            }
        );
        assert_eq!(
            fn_drop(&[arr.clone(), n(1.0), n(0.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(3.0), n(4.0)]],
                source: None
            }
        );
        assert_eq!(
            fn_expand(&[arr, n(3.0), n(3.0), s("x")], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0), s("x")],
                    vec![n(3.0), n(4.0), s("x")],
                    vec![s("x"), s("x"), s("x")]
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_filter_stack_and_transforms() {
        let arr = FormulaValue::Array {
            data: vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)]],
            source: None,
        };
        let include = FormulaValue::Array {
            data: vec![
                vec![FormulaValue::Boolean(false)],
                vec![FormulaValue::Boolean(true)],
                vec![FormulaValue::Boolean(true)],
            ],
            source: None,
        };
        assert_eq!(
            fn_filter(&[arr.clone(), include], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(2.0)], vec![n(3.0)]],
                source: None
            }
        );

        let h = fn_hstack(
            &[
                FormulaValue::Array {
                    data: vec![vec![n(1.0)], vec![n(2.0)]],
                    source: None,
                },
                FormulaValue::Array {
                    data: vec![vec![n(10.0)], vec![n(20.0)]],
                    source: None,
                },
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            h,
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(10.0)], vec![n(2.0), n(20.0)]],
                source: None
            }
        );

        let t = fn_transpose(&[h], &EvaluationContext::simple()).unwrap();
        assert_eq!(
            t,
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0)], vec![n(10.0), n(20.0)]],
                source: None
            }
        );
    }

    #[test]
    fn test_lookup_sort_unique_wrap() {
        let lookup = fn_lookup(
            &[
                n(2.5),
                FormulaValue::Array {
                    data: vec![vec![n(1.0), n(2.0), n(3.0)]],
                    source: None,
                },
                FormulaValue::Array {
                    data: vec![vec![s("a"), s("b"), s("c")]],
                    source: None,
                },
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(lookup, s("b"));

        let sorted = fn_sort(
            &[
                FormulaValue::Array {
                    data: vec![vec![n(2.0)], vec![n(1.0)]],
                    source: None,
                },
                n(1.0),
                n(1.0),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            sorted,
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)]],
                source: None
            }
        );

        let uniq = fn_unique(
            &[FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(1.0)], vec![n(2.0)]],
                source: None,
            }],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            uniq,
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)]],
                source: None
            }
        );

        let wrapped = fn_wraprows(
            &[
                FormulaValue::Array {
                    data: vec![vec![n(1.0), n(2.0), n(3.0)]],
                    source: None,
                },
                n(2.0),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            wrapped,
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0)],
                    vec![n(3.0), FormulaValue::Error(CellError::Na)]
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_tocol_torow_sortby_vstack() {
        let arr = FormulaValue::Array {
            data: vec![vec![n(1.0), n(2.0)], vec![n(3.0), FormulaValue::Empty]],
            source: None,
        };
        assert_eq!(
            fn_tocol(&[arr.clone(), n(1.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)]],
                source: None
            }
        );
        assert_eq!(
            fn_torow(&[arr, n(1.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0), n(3.0)]],
                source: None
            }
        );

        let sorted = fn_sortby(
            &[
                FormulaValue::Array {
                    data: vec![vec![s("b")], vec![s("a")]],
                    source: None,
                },
                FormulaValue::Array {
                    data: vec![vec![n(2.0)], vec![n(1.0)]],
                    source: None,
                },
                n(1.0),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            sorted,
            FormulaValue::Array {
                data: vec![vec![s("a")], vec![s("b")]],
                source: None
            }
        );

        let v = fn_vstack(
            &[
                FormulaValue::Array {
                    data: vec![vec![n(1.0)]],
                    source: None,
                },
                FormulaValue::Array {
                    data: vec![vec![n(2.0), n(3.0)]],
                    source: None,
                },
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            v,
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), FormulaValue::Error(CellError::Na)],
                    vec![n(2.0), n(3.0)]
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_date_intl_and_web_helpers() {
        // 2024-01-01 serial is 45292; Mon-Fri work week by default.
        assert_eq!(
            fn_networkdays_intl(&[n(45292.0), n(45301.0)], &EvaluationContext::simple()).unwrap(),
            n(8.0)
        );
        // Weekend only Sunday (11) makes Saturday a workday.
        assert_eq!(
            fn_workday_intl(&[n(45292.0), n(5.0), n(11.0)], &EvaluationContext::simple()).unwrap(),
            n(45297.0)
        );

        assert_eq!(
            fn_encodeurl(&[s("a b+c")], &EvaluationContext::simple()).unwrap(),
            s("a%20b%2Bc")
        );
        assert_eq!(
            fn_filterxml(&[s("<x/>"), s("//x")], &EvaluationContext::simple()).unwrap(),
            s("")
        );
        assert_eq!(
            fn_webservice(&[s("https://example.com")], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_getpivotdata(&[s("x")], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_rtd(&[s("x")], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        assert_eq!(
            fn_image(&[s("x")], &EvaluationContext::simple()).unwrap(),
            s("x")
        );

        let one = eval("=1").unwrap();
        assert_eq!(one, n(1.0));
    }

    #[test]
    fn test_filterxml_docs() {
        assert_eq!(
            eval(r#"=FILTERXML("<root><item>hello</item></root>","//item")"#).unwrap(),
            s("hello")
        );
        assert_eq!(
            eval(r#"=FILTERXML("<root><item code=""A1"">hello</item></root>","/root/item/@code")"#)
                .unwrap(),
            s("A1")
        );
        assert_eq!(
            eval(r#"=FILTERXML("<root><item>1</item><item>2</item></root>","/root/item[2]")"#)
                .unwrap(),
            n(2.0)
        );
        assert_eq!(
            eval(r#"=FILTERXML("<root><item><value>42.5</value></item></root>","//value")"#)
                .unwrap(),
            n(42.5)
        );
        assert_eq!(
            eval(r#"=FILTERXML("<root><item>hello</item></root>","//missing")"#).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            eval(r#"=FILTERXML("<root><item></root>","//item")"#).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_webservice_callback() {
        let handler = |url: &str| Some(format!("body:{}", url));
        let mut ctx = EvaluationContext::simple();
        ctx.web_service_fn = Some(&handler);

        assert_eq!(
            eval_with_ctx(r#"=WEBSERVICE("https://example.com")"#, &ctx).unwrap(),
            s("body:https://example.com")
        );
    }

    #[test]
    fn test_rtd_callback() {
        let handler = |prog_id: &str, server: &str, topics: &[String]| {
            Some(format!("{}|{}|{}", prog_id, server, topics.join("|")))
        };
        let mut ctx = EvaluationContext::simple();
        ctx.rtd_fn = Some(&handler);

        assert_eq!(
            eval_with_ctx(r#"=RTD("prog","srv","topic1","topic2")"#, &ctx).unwrap(),
            s("prog|srv|topic1|topic2")
        );
    }

    #[test]
    fn test_image_returns_display_string_and_emits_metadata() {
        let images = Mutex::new(Vec::new());
        let sink = |sheet: usize, row: u32, col: u16, info: ImageInfo| {
            images.lock().unwrap().push((sheet, row, col, info));
        };
        let mut ctx = EvaluationContext::simple();
        ctx.current_sheet = 2;
        ctx.current_row = 4;
        ctx.current_col = 7;
        ctx.image_sink = Some(&sink);

        assert_eq!(
            eval_with_ctx(
                r#"=IMAGE("https://example.com/img.png","Logo",3,48,96)"#,
                &ctx,
            )
            .unwrap(),
            s("Logo")
        );
        assert_eq!(
            eval(r#"=IMAGE("https://example.com/img.png")"#).unwrap(),
            s("https://example.com/img.png")
        );

        let images = images.lock().unwrap();
        assert_eq!(images.len(), 1);
        assert_eq!(
            images[0],
            (
                2,
                4,
                7,
                ImageInfo {
                    source: "https://example.com/img.png".to_string(),
                    alt_text: "Logo".to_string(),
                    sizing: ImageSizing::Custom,
                    width: Some(96.0),
                    height: Some(48.0),
                },
            )
        );
    }

    // ===== DOCS-BASED TESTS =====

    #[test]
    fn test_address_docs() {
        // Docs: =ADDRESS(2,3) -> $C$2 (absolute reference, default)
        assert_eq!(
            eval("=ADDRESS(2,3)").unwrap(),
            FormulaValue::String("$C$2".into())
        );
        // Docs: =ADDRESS(2,3,2) -> C$2 (absolute row; relative column)
        assert_eq!(
            eval("=ADDRESS(2,3,2)").unwrap(),
            FormulaValue::String("C$2".into())
        );
        // Docs: =ADDRESS(2,3,2,FALSE) -> R2C[3] (R1C1 style)
        assert_eq!(
            eval("=ADDRESS(2,3,2,FALSE)").unwrap(),
            FormulaValue::String("R2C[3]".into())
        );
        // Docs: =ADDRESS(2,3,1,FALSE,"[Book1]Sheet1") -> '[Book1]Sheet1'!R2C3
        assert_eq!(
            eval("=ADDRESS(2,3,1,FALSE,\"[Book1]Sheet1\")").unwrap(),
            FormulaValue::String("'[Book1]Sheet1'!R2C3".into())
        );
        // Docs: =ADDRESS(2,3,1,FALSE,"EXCEL SHEET") -> 'EXCEL SHEET'!R2C3
        assert_eq!(
            eval("=ADDRESS(2,3,1,FALSE,\"EXCEL SHEET\")").unwrap(),
            FormulaValue::String("'EXCEL SHEET'!R2C3".into())
        );
    }

    #[test]
    fn test_lookup_docs() {
        // Docs vector form: lookup_vector={4.14;4.19;5.17;5.77;6.39}
        //                   result_vector={"red";"orange";"yellow";"green";"blue"}

        // =LOOKUP(4.19,...) -> "orange" (exact match)
        assert_eq!(
            eval("=LOOKUP(4.19,{4.14;4.19;5.17;5.77;6.39},{\"red\";\"orange\";\"yellow\";\"green\";\"blue\"})").unwrap(),
            FormulaValue::String("orange".into())
        );
        // =LOOKUP(5.75,...) -> "yellow" (nearest smaller value is 5.17)
        assert_eq!(
            eval("=LOOKUP(5.75,{4.14;4.19;5.17;5.77;6.39},{\"red\";\"orange\";\"yellow\";\"green\";\"blue\"})").unwrap(),
            FormulaValue::String("yellow".into())
        );
        // =LOOKUP(7.66,...) -> "blue" (nearest smaller value is 6.39)
        assert_eq!(
            eval("=LOOKUP(7.66,{4.14;4.19;5.17;5.77;6.39},{\"red\";\"orange\";\"yellow\";\"green\";\"blue\"})").unwrap(),
            FormulaValue::String("blue".into())
        );
        // =LOOKUP(0,...) -> #N/A (0 < smallest value 4.14)
        assert_eq!(
            eval("=LOOKUP(0,{4.14;4.19;5.17;5.77;6.39},{\"red\";\"orange\";\"yellow\";\"green\";\"blue\"})").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_choose_docs() {
        // Docs: =CHOOSE(2,"1st","2nd","3rd","Finished") -> "2nd"
        assert_eq!(
            eval("=CHOOSE(2,\"1st\",\"2nd\",\"3rd\",\"Finished\")").unwrap(),
            s("2nd")
        );
        // Docs: =CHOOSE(4,"Nails","Screws","Nuts","Bolts") -> "Bolts"
        assert_eq!(
            eval("=CHOOSE(4,\"Nails\",\"Screws\",\"Nuts\",\"Bolts\")").unwrap(),
            s("Bolts")
        );
        // Docs: =CHOOSE(3,"Wide",115,"world",8) -> "world"
        assert_eq!(
            eval("=CHOOSE(3,\"Wide\",115,\"world\",8)").unwrap(),
            s("world")
        );
    }

    #[test]
    fn test_rows_docs() {
        // Docs: =ROWS(C1:E4) -> 4 (4x3 array)
        assert_eq!(eval("=ROWS({0,0,0;0,0,0;0,0,0;0,0,0})").unwrap(), n(4.0));
        // Docs: =ROWS({1,2,3;4,5,6}) -> 2
        assert_eq!(eval("=ROWS({1,2,3;4,5,6})").unwrap(), n(2.0));
    }

    #[test]
    fn test_columns_docs() {
        // Docs: =COLUMNS(C1:E4) -> 3 (4x3 array)
        assert_eq!(eval("=COLUMNS({0,0,0;0,0,0;0,0,0;0,0,0})").unwrap(), n(3.0));
        // Docs: =COLUMNS({1,2,3;4,5,6}) -> 3
        assert_eq!(eval("=COLUMNS({1,2,3;4,5,6})").unwrap(), n(3.0));
    }

    #[test]
    fn test_row_docs() {
        // Docs: ROW(C10)=10, ROW()=current — need cell context, skip those.
        // Test ROW with array arg: 3-row column -> {1;2;3}
        assert_eq!(
            eval("=ROW({0;0;0})").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)],],
                source: None
            }
        );
    }

    #[test]
    fn test_column_docs() {
        // Docs: COLUMN(D10)=4 — need cell context, skip.
        // Test COLUMN with array arg: 1x3 row -> {1,2,3}
        assert_eq!(
            eval("=COLUMN({0,0,0})").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0), n(3.0)]],
                source: None
            }
        );
    }

    #[test]
    fn test_choosecols_docs() {
        // Docs Example 1: 6x5 array, select columns 1, 3, 5, 1
        assert_eq!(
            eval("=CHOOSECOLS({1,2,3,4,5;6,7,8,9,10;11,12,13,14,15;16,17,18,19,20;21,22,23,24,25;26,27,28,29,30},1,3,5,1)").unwrap(),
            FormulaValue::Array { data: vec![
                vec![n(1.0), n(3.0), n(5.0), n(1.0)],
                vec![n(6.0), n(8.0), n(10.0), n(6.0)],
                vec![n(11.0), n(13.0), n(15.0), n(11.0)],
                vec![n(16.0), n(18.0), n(20.0), n(16.0)],
                vec![n(21.0), n(23.0), n(25.0), n(21.0)],
                vec![n(26.0), n(28.0), n(30.0), n(26.0)],
            ], source: None }
        );
        // Docs Example 3: select columns -1, -2 (last two reversed)
        assert_eq!(
            eval("=CHOOSECOLS({1,2,13,14;3,4,15,16;5,6,17,18;7,8,19,20;9,10,21,22;11,12,23,24},-1,-2)").unwrap(),
            FormulaValue::Array { data: vec![
                vec![n(14.0), n(13.0)],
                vec![n(16.0), n(15.0)],
                vec![n(18.0), n(17.0)],
                vec![n(20.0), n(19.0)],
                vec![n(22.0), n(21.0)],
                vec![n(24.0), n(23.0)],
            ], source: None }
        );
    }

    #[test]
    fn test_chooserows_docs() {
        // Docs Example 1: 6x2 array, select rows 1, 3, 5, 1
        assert_eq!(
            eval("=CHOOSEROWS({1,2;3,4;5,6;7,8;9,10;11,12},1,3,5,1)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0)],
                    vec![n(5.0), n(6.0)],
                    vec![n(9.0), n(10.0)],
                    vec![n(1.0), n(2.0)],
                ],
                source: None
            }
        );
        // Docs Example 3: select rows -1 (last), -2 (second-to-last)
        assert_eq!(
            eval("=CHOOSEROWS({1,2;3,4;5,6;7,8;9,10;11,12},-1,-2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(11.0), n(12.0)], vec![n(9.0), n(10.0)],],
                source: None
            }
        );
    }

    #[test]
    fn test_drop_docs() {
        // Docs: drop first 2 rows
        assert_eq!(
            eval("=DROP({1,2,3;4,5,6;7,8,9},2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(7.0), n(8.0), n(9.0)]],
                source: None
            }
        );
        // Docs: drop first 2 columns
        assert_eq!(
            eval("=DROP({1,2,3;4,5,6;7,8,9},,2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(3.0)], vec![n(6.0)], vec![n(9.0)]],
                source: None
            }
        );
        // Docs: drop last 2 rows
        assert_eq!(
            eval("=DROP({1,2,3;4,5,6;7,8,9},-2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0), n(3.0)]],
                source: None
            }
        );
        // Docs: drop first 2 rows and first 2 columns
        assert_eq!(
            eval("=DROP({1,2,3;4,5,6;7,8,9},2,2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(9.0)]],
                source: None
            }
        );
    }

    #[test]
    fn test_take_docs() {
        // Docs: take first 2 rows
        assert_eq!(
            eval("=TAKE({1,2,3;4,5,6;7,8,9},2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0), n(3.0)], vec![n(4.0), n(5.0), n(6.0)],],
                source: None
            }
        );
        // Docs: take first 2 columns
        assert_eq!(
            eval("=TAKE({1,2,3;4,5,6;7,8,9},,2)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0)],
                    vec![n(4.0), n(5.0)],
                    vec![n(7.0), n(8.0)],
                ],
                source: None
            }
        );
        // Docs: take last 2 rows
        assert_eq!(
            eval("=TAKE({1,2,3;4,5,6;7,8,9},-2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(4.0), n(5.0), n(6.0)], vec![n(7.0), n(8.0), n(9.0)],],
                source: None
            }
        );
        // Docs: take first 2 rows and first 2 columns
        assert_eq!(
            eval("=TAKE({1,2,3;4,5,6;7,8,9},2,2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0)], vec![n(4.0), n(5.0)],],
                source: None
            }
        );
    }

    #[test]
    fn test_expand_docs() {
        // Docs: expand 2x2 to 3x3, pad with #N/A
        assert_eq!(
            eval("=EXPAND({1,2;3,4},3,3)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0), FormulaValue::Error(CellError::Na)],
                    vec![n(3.0), n(4.0), FormulaValue::Error(CellError::Na)],
                    vec![
                        FormulaValue::Error(CellError::Na),
                        FormulaValue::Error(CellError::Na),
                        FormulaValue::Error(CellError::Na)
                    ],
                ],
                source: None
            }
        );
        // Docs: expand scalar 1 to 3x3, pad with "-"
        assert_eq!(
            eval("=EXPAND(1,3,3,\"-\")").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), s("-"), s("-")],
                    vec![s("-"), s("-"), s("-")],
                    vec![s("-"), s("-"), s("-")],
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_wrapcols_docs() {
        // Docs: wrap 7 elements into columns of 3, pad #N/A
        assert_eq!(
            eval("=WRAPCOLS({\"A\",\"B\",\"C\",\"D\",\"E\",\"F\",\"G\"},3)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![s("A"), s("D"), s("G")],
                    vec![s("B"), s("E"), FormulaValue::Error(CellError::Na)],
                    vec![s("C"), s("F"), FormulaValue::Error(CellError::Na)],
                ],
                source: None
            }
        );
        // Docs: wrap 7 elements into columns of 3, pad "x"
        assert_eq!(
            eval("=WRAPCOLS({\"A\",\"B\",\"C\",\"D\",\"E\",\"F\",\"G\"},3,\"x\")").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![s("A"), s("D"), s("G")],
                    vec![s("B"), s("E"), s("x")],
                    vec![s("C"), s("F"), s("x")],
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_wraprows_docs() {
        // Docs: wrap 7 elements into rows of 3, pad #N/A
        assert_eq!(
            eval("=WRAPROWS({\"A\",\"B\",\"C\",\"D\",\"E\",\"F\",\"G\"},3)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![s("A"), s("B"), s("C")],
                    vec![s("D"), s("E"), s("F")],
                    vec![
                        s("G"),
                        FormulaValue::Error(CellError::Na),
                        FormulaValue::Error(CellError::Na)
                    ],
                ],
                source: None
            }
        );
        // Docs: wrap 7 elements into rows of 3, pad "x"
        assert_eq!(
            eval("=WRAPROWS({\"A\",\"B\",\"C\",\"D\",\"E\",\"F\",\"G\"},3,\"x\")").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![s("A"), s("B"), s("C")],
                    vec![s("D"), s("E"), s("F")],
                    vec![s("G"), s("x"), s("x")],
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_transpose_docs() {
        // Docs: TRANSPOSE(A1:B4) — 4-row × 2-col transposed to 2-row × 4-col
        assert_eq!(
            eval("=TRANSPOSE({1,2;3,4;5,6;7,8})").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(3.0), n(5.0), n(7.0)],
                    vec![n(2.0), n(4.0), n(6.0), n(8.0)],
                ],
                source: None
            }
        );
        // 2×2 transpose: rows become columns
        assert_eq!(
            eval("=TRANSPOSE({1,2;3,4})").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(3.0)], vec![n(2.0), n(4.0)],],
                source: None
            }
        );
        // Single row → single column
        assert_eq!(
            eval("=TRANSPOSE({1,2,3,4})").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)], vec![n(4.0)],],
                source: None
            }
        );
    }

    #[test]
    fn test_tocol_docs() {
        // Docs Example 1: =TOCOL(A2:D4) — 3×4 text array, scan by row (default)
        assert_eq!(
            eval(r#"=TOCOL({"Ben","Peter","Mary","Sam";"John","Hillary","Jenny","James";"Agnes","Harry","Felicity","Joe"})"#).unwrap(),
            FormulaValue::Array { data: vec![
                vec![s("Ben")], vec![s("Peter")], vec![s("Mary")], vec![s("Sam")],
                vec![s("John")], vec![s("Hillary")], vec![s("Jenny")], vec![s("James")],
                vec![s("Agnes")], vec![s("Harry")], vec![s("Felicity")], vec![s("Joe")],
            ], source: None }
        );
        // Numeric 2×3, scan by row (default)
        assert_eq!(
            eval("=TOCOL({1,2,3;4,5,6})").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0)],
                    vec![n(2.0)],
                    vec![n(3.0)],
                    vec![n(4.0)],
                    vec![n(5.0)],
                    vec![n(6.0)],
                ],
                source: None
            }
        );
        // Docs Example 4 pattern: scan by column (scan_by_column=TRUE)
        assert_eq!(
            eval("=TOCOL({1,2,3;4,5,6},,TRUE)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0)],
                    vec![n(4.0)],
                    vec![n(2.0)],
                    vec![n(5.0)],
                    vec![n(3.0)],
                    vec![n(6.0)],
                ],
                source: None
            }
        );
        // Docs: ignore=0 keeps all values (explicit default)
        assert_eq!(
            eval("=TOCOL({1,2;3,4},0)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)], vec![n(4.0)],],
                source: None
            }
        );
        // Docs Example 4: scan by column on 3×4 text grid
        assert_eq!(
            eval(r#"=TOCOL({"Ben","Peter","Mary","Sam";"John","Hillary","Jenny","James";"Agnes","Harry","Felicity","Joe"},,TRUE)"#).unwrap(),
            FormulaValue::Array { data: vec![
                vec![s("Ben")], vec![s("John")], vec![s("Agnes")],
                vec![s("Peter")], vec![s("Hillary")], vec![s("Harry")],
                vec![s("Mary")], vec![s("Jenny")], vec![s("Felicity")],
                vec![s("Sam")], vec![s("James")], vec![s("Joe")],
            ], source: None }
        );
    }

    #[test]
    fn test_torow_docs() {
        // Docs Example 1: =TOROW(A2:D4) — 3×4 text array, scan by row (default)
        assert_eq!(
            eval(r#"=TOROW({"Ben","Peter","Mary","Sam";"John","Hillary","Jenny","James";"Agnes","Harry","Felicity","Joe"})"#).unwrap(),
            FormulaValue::Array { data: vec![vec![
                s("Ben"), s("Peter"), s("Mary"), s("Sam"),
                s("John"), s("Hillary"), s("Jenny"), s("James"),
                s("Agnes"), s("Harry"), s("Felicity"), s("Joe"),
            ]], source: None }
        );
        // Numeric 2×3, scan by row (default)
        assert_eq!(
            eval("=TOROW({1,2,3;4,5,6})").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0), n(3.0), n(4.0), n(5.0), n(6.0),]],
                source: None
            }
        );
        // Docs Example 4 pattern: scan by column
        assert_eq!(
            eval("=TOROW({1,2,3;4,5,6},,TRUE)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(4.0), n(2.0), n(5.0), n(3.0), n(6.0),]],
                source: None
            }
        );
        // Docs: ignore=0 keeps all values (explicit default)
        assert_eq!(
            eval("=TOROW({1,2;3,4},0)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0), n(3.0), n(4.0),]],
                source: None
            }
        );
        // Docs Example 4: scan by column on 3×4 text grid
        assert_eq!(
            eval(r#"=TOROW({"Ben","Peter","Mary","Sam";"John","Hillary","Jenny","James";"Agnes","Harry","Felicity","Joe"},,TRUE)"#).unwrap(),
            FormulaValue::Array { data: vec![vec![
                s("Ben"), s("John"), s("Agnes"),
                s("Peter"), s("Hillary"), s("Harry"),
                s("Mary"), s("Jenny"), s("Felicity"),
                s("Sam"), s("James"), s("Joe"),
            ]], source: None }
        );
    }

    #[test]
    fn test_hstack_docs() {
        // Docs Example 1: Two 2×3 text arrays horizontally appended
        assert_eq!(
            eval(r#"=HSTACK({"A","B","C";"D","E","F"},{"AA","BB","CC";"DD","EE","FF"})"#).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![s("A"), s("B"), s("C"), s("AA"), s("BB"), s("CC")],
                    vec![s("D"), s("E"), s("F"), s("DD"), s("EE"), s("FF")],
                ],
                source: None
            }
        );
        // Docs Example 2: Three arrays, different row counts (3,2,1), #N/A fill
        assert_eq!(
            eval(r#"=HSTACK({1,2;3,4;5,6},{"A","B";"C","D"},{"X","Y"})"#).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0), s("A"), s("B"), s("X"), s("Y")],
                    vec![
                        n(3.0),
                        n(4.0),
                        s("C"),
                        s("D"),
                        FormulaValue::Error(CellError::Na),
                        FormulaValue::Error(CellError::Na)
                    ],
                    vec![
                        n(5.0),
                        n(6.0),
                        FormulaValue::Error(CellError::Na),
                        FormulaValue::Error(CellError::Na),
                        FormulaValue::Error(CellError::Na),
                        FormulaValue::Error(CellError::Na)
                    ],
                ],
                source: None
            }
        );
        // Docs Example 3 variant: scalar appended horizontally
        assert_eq!(
            eval("=HSTACK({1;2;3},4)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(4.0)],
                    vec![n(2.0), FormulaValue::Error(CellError::Na)],
                    vec![n(3.0), FormulaValue::Error(CellError::Na)],
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_vstack_docs() {
        // Docs Example 1: Two 2×3 text arrays vertically appended
        assert_eq!(
            eval(r#"=VSTACK({"A","B","C";"D","E","F"},{"AA","BB","CC";"DD","EE","FF"})"#).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![s("A"), s("B"), s("C")],
                    vec![s("D"), s("E"), s("F")],
                    vec![s("AA"), s("BB"), s("CC")],
                    vec![s("DD"), s("EE"), s("FF")],
                ],
                source: None
            }
        );
        // Docs Example 2: Three arrays, all 2 cols (3+2+1=6 rows)
        assert_eq!(
            eval(r#"=VSTACK({1,2;3,4;5,6},{"A","B";"C","D"},{"X","Y"})"#).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0)],
                    vec![n(3.0), n(4.0)],
                    vec![n(5.0), n(6.0)],
                    vec![s("A"), s("B")],
                    vec![s("C"), s("D")],
                    vec![s("X"), s("Y")],
                ],
                source: None
            }
        );
        // Docs Example 3: Different column counts (2 vs 3), #N/A column padding
        assert_eq!(
            eval(r#"=VSTACK({1,2;3,4;5,6},{"A","B","C";"D","E","F"})"#).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0), FormulaValue::Error(CellError::Na)],
                    vec![n(3.0), n(4.0), FormulaValue::Error(CellError::Na)],
                    vec![n(5.0), n(6.0), FormulaValue::Error(CellError::Na)],
                    vec![s("A"), s("B"), s("C")],
                    vec![s("D"), s("E"), s("F")],
                ],
                source: None
            }
        );
        // Docs Example 3 variant: scalar appended vertically
        assert_eq!(
            eval("=VSTACK({1,2;3,4},5)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0)],
                    vec![n(3.0), n(4.0)],
                    vec![n(5.0), FormulaValue::Error(CellError::Na)],
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_sort_docs() {
        // Ascending (default sort_order=1)
        assert_eq!(
            eval("=SORT({622;961;691;445;378;483;650;783;142;404})").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(142.0)],
                    vec![n(378.0)],
                    vec![n(404.0)],
                    vec![n(445.0)],
                    vec![n(483.0)],
                    vec![n(622.0)],
                    vec![n(650.0)],
                    vec![n(691.0)],
                    vec![n(783.0)],
                    vec![n(961.0)],
                ],
                source: None
            }
        );
        // Descending (sort_order=-1) — exact docs example
        assert_eq!(
            eval("=SORT({622;961;691;445;378;483;650;783;142;404},1,-1)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(961.0)],
                    vec![n(783.0)],
                    vec![n(691.0)],
                    vec![n(650.0)],
                    vec![n(622.0)],
                    vec![n(483.0)],
                    vec![n(445.0)],
                    vec![n(404.0)],
                    vec![n(378.0)],
                    vec![n(142.0)],
                ],
                source: None
            }
        );
        // Multi-column sort by first column ascending (docs behavior pattern)
        assert_eq!(
            eval(r#"=SORT({3,"c";1,"a";2,"b"},1)"#).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), s("a")],
                    vec![n(2.0), s("b")],
                    vec![n(3.0), s("c")],
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_sortby_docs() {
        // Simple ascending sort by separate by_array
        assert_eq!(
            eval(r#"=SORTBY({"b";"a";"c"},{2;1;3})"#).unwrap(),
            FormulaValue::Array {
                data: vec![vec![s("a")], vec![s("b")], vec![s("c")]],
                source: None
            }
        );
        // Descending sort
        assert_eq!(
            eval(r#"=SORTBY({"b";"a";"c"},{2;1;3},-1)"#).unwrap(),
            FormulaValue::Array {
                data: vec![vec![s("c")], vec![s("b")], vec![s("a")]],
                source: None
            }
        );
        // Full docs example: 8 people sorted by age ascending
        assert_eq!(
            eval(concat!(
                r#"=SORTBY({"Tom",52;"Fred",65;"Amy",22;"Sal",73;"#,
                r#""Fritz",19;"Sravan",39;"Xi",19;"Hector",66},"#,
                "{52;65;22;73;19;39;19;66})"
            ))
            .unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![s("Fritz"), n(19.0)],
                    vec![s("Xi"), n(19.0)],
                    vec![s("Amy"), n(22.0)],
                    vec![s("Sravan"), n(39.0)],
                    vec![s("Tom"), n(52.0)],
                    vec![s("Fred"), n(65.0)],
                    vec![s("Hector"), n(66.0)],
                    vec![s("Sal"), n(73.0)],
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_unique_docs() {
        // Basic: extract distinct values preserving first-occurrence order
        assert_eq!(
            eval("=UNIQUE({1;2;1;3;2})").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)]],
                source: None
            }
        );
        // exactly_once=TRUE: only values appearing exactly once
        assert_eq!(
            eval("=UNIQUE({1;2;1;3;2;4},,TRUE)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(3.0)], vec![n(4.0)]],
                source: None
            }
        );
        // Unique strings (docs pattern: unique product names)
        assert_eq!(
            eval(r#"=UNIQUE({"Apple";"Grape";"Apple";"Banana";"Grape"})"#).unwrap(),
            FormulaValue::Array {
                data: vec![vec![s("Apple")], vec![s("Grape")], vec![s("Banana")],],
                source: None
            }
        );
    }

    #[test]
    fn test_filter_docs() {
        // Column filter with boolean include
        assert_eq!(
            eval("=FILTER({1;2;3},{TRUE;FALSE;TRUE})").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(3.0)]],
                source: None
            }
        );
        // Row filter
        assert_eq!(
            eval("=FILTER({1,2,3},{TRUE,FALSE,TRUE})").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(3.0)]],
                source: None
            }
        );
        // Multi-column filter (docs pattern: filter sales rows by product)
        assert_eq!(
            eval(r#"=FILTER({"Apple",6380;"Grape",5619;"Apple",4394;"Banana",5323},{TRUE;FALSE;TRUE;FALSE})"#).unwrap(),
            FormulaValue::Array { data: vec![
                vec![s("Apple"), n(6380.0)],
                vec![s("Apple"), n(4394.0)],
            ], source: None }
        );
        // if_empty argument when filter returns nothing
        assert_eq!(
            eval(r#"=FILTER({1;2;3},{FALSE;FALSE;FALSE},"No data")"#).unwrap(),
            s("No data")
        );
    }

    #[test]
    fn test_encodeurl_docs() {
        // Docs: =ENCODEURL("hello world") → "hello%20world"
        assert_eq!(
            eval(r#"=ENCODEURL("hello world")"#).unwrap(),
            s("hello%20world")
        );
        // Docs: reserved chars encoded
        assert_eq!(
            eval(r#"=ENCODEURL("http://example.com")"#).unwrap(),
            s("http%3A%2F%2Fexample.com")
        );
        // Alphanumeric and unreserved chars (-_.~) not encoded
        assert_eq!(
            eval(r#"=ENCODEURL("abc-123_test.txt~")"#).unwrap(),
            s("abc-123_test.txt~")
        );
    }

    #[test]
    fn test_trimrange_docs() {
        let e = FormulaValue::Empty;
        let ctx = EvaluationContext::simple();

        // --- Docs: TRIMRANGE trims blank rows/cols from edges ---

        // Basic: leading and trailing blank rows and columns trimmed (default trim_rows=3, trim_cols=3)
        // Array layout (3x3):
        //   "" "" ""
        //   "" 1  2
        //   "" 3  4
        assert_eq!(
            fn_trimrange(
                &[FormulaValue::Array {
                    data: vec![
                        vec![e.clone(), e.clone(), e.clone()],
                        vec![e.clone(), n(1.0), n(2.0)],
                        vec![e.clone(), n(3.0), n(4.0)],
                    ],
                    source: None
                }],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0)], vec![n(3.0), n(4.0)]],
                source: None
            }
        );

        // Already trimmed array — no change
        assert_eq!(
            fn_trimrange(
                &[FormulaValue::Array {
                    data: vec![vec![n(1.0), n(2.0)], vec![n(3.0), n(4.0)],],
                    source: None
                }],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0)], vec![n(3.0), n(4.0)]],
                source: None
            }
        );

        // Trailing blank rows and cols only
        // Array (3x3):
        //   1  2  ""
        //   3  4  ""
        //   "" "" ""
        assert_eq!(
            fn_trimrange(
                &[FormulaValue::Array {
                    data: vec![
                        vec![n(1.0), n(2.0), e.clone()],
                        vec![n(3.0), n(4.0), e.clone()],
                        vec![e.clone(), e.clone(), e.clone()],
                    ],
                    source: None
                }],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0)], vec![n(3.0), n(4.0)]],
                source: None
            }
        );

        // Single non-blank cell surrounded by blanks
        // Array (3x3):
        //   "" "" ""
        //   "" 42 ""
        //   "" "" ""
        assert_eq!(
            fn_trimrange(
                &[FormulaValue::Array {
                    data: vec![
                        vec![e.clone(), e.clone(), e.clone()],
                        vec![e.clone(), n(42.0), e.clone()],
                        vec![e.clone(), e.clone(), e.clone()],
                    ],
                    source: None
                }],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(42.0)]],
                source: None
            }
        );

        // All blank → #CALC!
        assert_eq!(
            fn_trimrange(
                &[FormulaValue::Array {
                    data: vec![vec![e.clone(), e.clone()], vec![e.clone(), e.clone()],],
                    source: None
                }],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Error(CellError::Calc)
        );

        // Single cell with value — unchanged
        assert_eq!(
            fn_trimrange(&[n(5.0)], &ctx).unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(5.0)]],
                source: None
            }
        );

        // --- Docs: trim_rows parameter ---

        // Docs table: TRIMRANGE(range,3,3) = Trim All (default)
        // Equivalent to A1.:.E10
        // Already tested above

        // trim_rows=0: None — don't trim any rows (cols still trimmed)
        // Array: ["",1,""; "",2,""; "",3,""] — col 0 and col 2 are all blank
        assert_eq!(
            fn_trimrange(
                &[
                    FormulaValue::Array {
                        data: vec![
                            vec![e.clone(), n(1.0), e.clone()],
                            vec![e.clone(), n(2.0), e.clone()],
                            vec![e.clone(), n(3.0), e.clone()],
                        ],
                        source: None
                    },
                    n(0.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)],],
                source: None
            }
        );

        // trim_rows=1: Trims leading blank rows only
        // Docs table: TRIMRANGE(range,1,1) = Trim Leading
        assert_eq!(
            fn_trimrange(
                &[
                    FormulaValue::Array {
                        data: vec![
                            vec![e.clone(), e.clone()],
                            vec![n(1.0), n(2.0)],
                            vec![e.clone(), e.clone()],
                        ],
                        source: None
                    },
                    n(1.0),
                    n(1.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0)], vec![e.clone(), e.clone()]],
                source: None
            }
        );

        // trim_rows=2: Trims trailing blank rows only
        // Docs table: TRIMRANGE(range,2,2) = Trim Trailing
        assert_eq!(
            fn_trimrange(
                &[
                    FormulaValue::Array {
                        data: vec![
                            vec![e.clone(), e.clone()],
                            vec![n(1.0), n(2.0)],
                            vec![e.clone(), e.clone()],
                        ],
                        source: None
                    },
                    n(2.0),
                    n(2.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![e.clone(), e.clone()], vec![n(1.0), n(2.0)]],
                source: None
            }
        );

        // --- Docs: trim_cols parameter ---

        // trim_cols=0: None — don't trim any columns
        assert_eq!(
            fn_trimrange(
                &[
                    FormulaValue::Array {
                        data: vec![
                            vec![e.clone(), n(1.0), e.clone()],
                            vec![e.clone(), n(2.0), e.clone()],
                        ],
                        source: None
                    },
                    n(3.0),
                    n(0.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![e.clone(), n(1.0), e.clone()],
                    vec![e.clone(), n(2.0), e.clone()],
                ],
                source: None
            }
        );

        // trim_cols=1: Trims leading blank columns only
        assert_eq!(
            fn_trimrange(
                &[
                    FormulaValue::Array {
                        data: vec![
                            vec![e.clone(), n(1.0), e.clone()],
                            vec![e.clone(), n(2.0), e.clone()],
                        ],
                        source: None
                    },
                    n(3.0),
                    n(1.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), e.clone()], vec![n(2.0), e.clone()]],
                source: None
            }
        );

        // trim_cols=2: Trims trailing blank columns only
        assert_eq!(
            fn_trimrange(
                &[
                    FormulaValue::Array {
                        data: vec![
                            vec![e.clone(), n(1.0), e.clone()],
                            vec![e.clone(), n(2.0), e.clone()],
                        ],
                        source: None
                    },
                    n(3.0),
                    n(2.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![e.clone(), n(1.0)], vec![e.clone(), n(2.0)]],
                source: None
            }
        );

        // --- Docs: Trim Refs equivalences ---

        // Trim All (.:.) — TRIMRANGE(range,3,3)
        // 4x3 with data in center, blanks on all edges
        let arr_with_edges = FormulaValue::Array {
            data: vec![
                vec![e.clone(), e.clone(), e.clone()],
                vec![e.clone(), s("X"), e.clone()],
                vec![e.clone(), s("Y"), e.clone()],
                vec![e.clone(), e.clone(), e.clone()],
            ],
            source: None,
        };
        assert_eq!(
            fn_trimrange(&[arr_with_edges.clone(), n(3.0), n(3.0)], &ctx).unwrap(),
            FormulaValue::Array {
                data: vec![vec![s("X")], vec![s("Y")]],
                source: None
            }
        );

        // Trim Trailing (:.) — TRIMRANGE(range,2,2)
        assert_eq!(
            fn_trimrange(&[arr_with_edges.clone(), n(2.0), n(2.0)], &ctx).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![e.clone(), e.clone()],
                    vec![e.clone(), s("X")],
                    vec![e.clone(), s("Y")],
                ],
                source: None
            }
        );

        // Trim Leading (.:) — TRIMRANGE(range,1,1)
        assert_eq!(
            fn_trimrange(&[arr_with_edges, n(1.0), n(1.0)], &ctx).unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![s("X"), e.clone()],
                    vec![s("Y"), e.clone()],
                    vec![e.clone(), e.clone()],
                ],
                source: None
            }
        );

        // Empty strings also count as blank
        assert_eq!(
            fn_trimrange(
                &[FormulaValue::Array {
                    data: vec![vec![s(""), s("")], vec![s(""), n(7.0)],],
                    source: None
                }],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(7.0)]],
                source: None
            }
        );

        // Interior blank rows/cols are preserved
        assert_eq!(
            fn_trimrange(
                &[FormulaValue::Array {
                    data: vec![
                        vec![e.clone(), e.clone(), e.clone()],
                        vec![n(1.0), e.clone(), n(2.0)],
                        vec![e.clone(), e.clone(), e.clone()],
                        vec![n(3.0), e.clone(), n(4.0)],
                        vec![e.clone(), e.clone(), e.clone()],
                    ],
                    source: None
                }],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), e.clone(), n(2.0)],
                    vec![e.clone(), e.clone(), e.clone()],
                    vec![n(3.0), e.clone(), n(4.0)],
                ],
                source: None
            }
        );
    }

    #[test]
    fn test_formulatext_docs() {
        use duke_sheets_core::Workbook;

        // Build a workbook: A2 has =TODAY(), B2 has plain value
        let mut wb = Workbook::new();
        {
            let ws = wb.worksheet_mut(0).unwrap();
            ws.set_cell_formula("A2", "=TODAY()").unwrap();
            ws.set_cell_value("B2", "hello").unwrap();
        }
        let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

        // Docs example: =FORMULATEXT(A2) returns "=TODAY()"
        assert_eq!(
            eval_with_ctx("=FORMULATEXT(A2)", &ctx).unwrap(),
            s("=TODAY()")
        );

        // Non-formula cell → #N/A
        assert_eq!(
            eval_with_ctx("=FORMULATEXT(B2)", &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );

        // Empty cell → #N/A
        assert_eq!(
            eval_with_ctx("=FORMULATEXT(C1)", &ctx).unwrap(),
            FormulaValue::Error(CellError::Na)
        );

        // Range reference → formula of upper-left cell
        assert_eq!(
            eval_with_ctx("=FORMULATEXT(A2:B2)", &ctx).unwrap(),
            s("=TODAY()")
        );

        // Error propagation
        assert_eq!(
            eval_with_ctx("=FORMULATEXT(#VALUE!)", &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // No workbook context → #N/A
        assert_eq!(
            eval("=FORMULATEXT(A1)").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_areas_docs() {
        use crate::ast::{BinaryOperator, CellReference, FormulaExpr, RangeReference};
        use duke_sheets_core::{CellAddress, CellRange};

        // Helper to build a Function AST node for AREAS with a single arg.
        let areas = |arg: FormulaExpr| -> FormulaExpr {
            FormulaExpr::Function {
                name: "AREAS".into(),
                args: vec![arg],
            }
        };
        let cell = |row: u32, col: u16| -> FormulaExpr {
            FormulaExpr::CellRef(CellReference {
                sheet: None,
                address: CellAddress::new(row, col),
            })
        };
        let range = |r0: u32, c0: u16, r1: u32, c1: u16| -> FormulaExpr {
            FormulaExpr::RangeRef(RangeReference {
                sheet: None,
                range: CellRange::new(CellAddress::new(r0, c0), CellAddress::new(r1, c1)),
            })
        };
        let union = |l: FormulaExpr, r: FormulaExpr| -> FormulaExpr {
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Union,
                left: Box::new(l),
                right: Box::new(r),
            }
        };
        let intersect = |l: FormulaExpr, r: FormulaExpr| -> FormulaExpr {
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Intersect,
                left: Box::new(l),
                right: Box::new(r),
            }
        };
        let eval_ast = |expr: FormulaExpr| -> FormulaValue {
            crate::evaluator::evaluate(&expr, &EvaluationContext::simple()).unwrap()
        };

        // Microsoft docs example 1: =AREAS(B2:D4) → 1
        assert_eq!(eval("=AREAS(B2:D4)").unwrap(), n(1.0));

        // Microsoft docs example 2: =AREAS((B2:D4,E5,F6:I9)) → 3
        // Parser doesn't support union syntax, so build AST manually.
        // B2:D4 = rows 1..3, cols 1..3; E5 = row 4, col 4; F6:I9 = rows 5..8, cols 5..8
        let arg2 = union(union(range(1, 1, 3, 3), cell(4, 4)), range(5, 5, 8, 8));
        assert_eq!(eval_ast(areas(arg2)), n(3.0));

        // Microsoft docs example 3: =AREAS(B2:D4 B2) → 1 (intersection)
        let arg3 = intersect(range(1, 1, 3, 3), cell(1, 1));
        assert_eq!(eval_ast(areas(arg3)), n(1.0));

        // Single cell reference
        assert_eq!(eval("=AREAS(A1)").unwrap(), n(1.0));

        // Two areas via union (manual AST)
        let arg_two = union(range(0, 0, 2, 2), range(3, 1, 4, 1));
        assert_eq!(eval_ast(areas(arg_two)), n(2.0));

        // Error propagation
        assert_eq!(
            eval("=AREAS(#REF!)").unwrap(),
            FormulaValue::Error(CellError::Ref)
        );
    }
}
