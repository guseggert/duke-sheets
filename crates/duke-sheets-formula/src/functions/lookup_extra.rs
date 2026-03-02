//! Additional lookup, reference, web, and date-intl functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use chrono::{Datelike, Duration, NaiveDate};
use duke_sheets_core::CellError;
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

fn to_i64_trunc(v: &FormulaValue) -> Option<i64> {
    v.as_number().map(|n| n.trunc() as i64)
}

fn scalar_i64(v: &FormulaValue) -> Result<i64, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => to_i64_trunc(v).ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn scalar_bool(v: &FormulaValue) -> Result<bool, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => v.as_bool().ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn scalar_string(v: &FormulaValue) -> Result<String, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => Ok(v.as_string()),
    }
}

fn as_array(v: &FormulaValue) -> Vec<Vec<FormulaValue>> {
    match v {
        FormulaValue::Array(a) => a.clone(),
        _ => vec![vec![v.clone()]],
    }
}

fn array_dims(a: &[Vec<FormulaValue>]) -> (usize, usize) {
    (a.len(), a.first().map(|r| r.len()).unwrap_or(0))
}

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
        FormulaValue::Array(rows) => {
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
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
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
    let row = match scalar_i64(args.get(0).unwrap()) {
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
        format!("R{}C{}", row, col)
    };

    let final_addr = match sheet_text {
        Some(s) if !s.is_empty() => format!("{}!{}", s, addr),
        _ => addr,
    };

    Ok(FormulaValue::String(final_addr))
}

/// AREAS(reference)
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
    let array = as_array(args.get(0).unwrap());
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

    Ok(FormulaValue::Array(out))
}

/// CHOOSEROWS(array, row_num1, ...)
pub fn fn_chooserows(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let array = as_array(args.get(0).unwrap());
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

    Ok(FormulaValue::Array(out))
}

/// DROP(array, rows, [columns])
pub fn fn_drop(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let array = as_array(args.get(0).unwrap());
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

    Ok(FormulaValue::Array(out))
}

/// EXPAND(array, rows, [columns], [pad_with])
pub fn fn_expand(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let array = as_array(args.get(0).unwrap());
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
    Ok(FormulaValue::Array(out))
}

/// FILTER(array, include, [if_empty])
pub fn fn_filter(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let array = as_array(args.get(0).unwrap());
    let include_arr = as_array(args.get(1).unwrap());
    let (rows, _cols) = array_dims(&array);

    let include_vec = if include_arr.len() == rows
        && include_arr.first().map(|r| r.len()).unwrap_or(0) == 1
    {
        include_arr
            .iter()
            .map(|r| r.first().cloned().unwrap_or(FormulaValue::Empty))
            .collect::<Vec<_>>()
    } else if include_arr.len() == 1 && include_arr.first().map(|r| r.len()).unwrap_or(0) == rows {
        include_arr[0].clone()
    } else {
        return Ok(FormulaValue::Error(CellError::Value));
    };

    let mut out = Vec::new();
    for (i, row) in array.iter().enumerate() {
        if let FormulaValue::Error(e) = include_vec[i] {
            return Ok(FormulaValue::Error(e));
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

    Ok(FormulaValue::Array(out))
}

/// FORMULATEXT(reference)
pub fn fn_formulatext(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    if let Some(FormulaValue::Error(e)) = args.first() {
        return Ok(FormulaValue::Error(*e));
    }
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

    Ok(FormulaValue::Array(out))
}

/// LOOKUP(lookup_value, lookup_vector, [result_vector])
pub fn fn_lookup(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let lookup_value = args.get(0).unwrap().clone();
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
    let mut arr = as_array(args.get(0).unwrap());
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
        let t = fn_transpose(&[FormulaValue::Array(arr)], _ctx)?;
        if let FormulaValue::Array(mut tr) = t {
            let idx = (sort_index - 1) as usize;
            if tr.is_empty() || idx >= tr[0].len() {
                return Ok(FormulaValue::Error(CellError::Value));
            }
            tr.sort_by(|a, b| compare_lookup_values(&a[idx], &b[idx]).unwrap_or(Ordering::Equal));
            if sort_order == -1 {
                tr.reverse();
            }
            return fn_transpose(&[FormulaValue::Array(tr)], _ctx);
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
    Ok(FormulaValue::Array(arr))
}

/// SORTBY(array, by_array1, [sort_order1], ...)
pub fn fn_sortby(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.get(0).unwrap());
    if arr.is_empty() {
        return Ok(FormulaValue::Array(arr));
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
    Ok(FormulaValue::Array(sorted))
}

/// TAKE(array, rows, [columns])
pub fn fn_take(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.get(0).unwrap());
    let (rows, cols) = array_dims(&arr);
    let take_rows = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let take_cols = match args.get(2) {
        Some(v) => match scalar_i64(v) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        None => cols as i64,
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

    Ok(FormulaValue::Array(out))
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
    let arr = as_array(args.get(0).unwrap());
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
    Ok(FormulaValue::Array(out))
}

/// TOROW(array, [ignore], [scan_by_column])
pub fn fn_torow(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.get(0).unwrap());
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

    Ok(FormulaValue::Array(vec![flatten_with_ignore(
        &arr, ignore, by_col,
    )]))
}

/// TRANSPOSE(array)
pub fn fn_transpose(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.get(0).unwrap());
    let (rows, cols) = array_dims(&arr);
    let mut out = vec![vec![FormulaValue::Empty; rows]; cols];
    for (r, row) in arr.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            out[c][r] = v.clone();
        }
    }
    Ok(FormulaValue::Array(out))
}

/// UNIQUE(array, [by_col], [exactly_once])
pub fn fn_unique(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arr = as_array(args.get(0).unwrap());
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
        let t = fn_transpose(&[FormulaValue::Array(arr)], _ctx)?;
        if let FormulaValue::Array(transposed) = t {
            let unique_rows = fn_unique(
                &[
                    FormulaValue::Array(transposed),
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

    Ok(FormulaValue::Array(out))
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

    Ok(FormulaValue::Array(out))
}

/// WRAPCOLS(vector, wrap_count, [pad_with])
pub fn fn_wrapcols(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let vec_values = flat_vector(&as_array(args.get(0).unwrap()), false);
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

    Ok(FormulaValue::Array(out))
}

/// WRAPROWS(vector, wrap_count, [pad_with])
pub fn fn_wraprows(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let vec_values = flat_vector(&as_array(args.get(0).unwrap()), false);
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

    Ok(FormulaValue::Array(out))
}

/// HYPERLINK(link_location, [friendly_name])
pub fn fn_hyperlink(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let link = match scalar_string(args.get(0).unwrap()) {
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
    let start = match scalar_i64(args.get(0).unwrap()) {
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
    let start = match scalar_i64(args.get(0).unwrap()) {
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
    let text = match scalar_string(args.get(0).unwrap()) {
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

/// FILTERXML(xml, xpath)
pub fn fn_filterxml(
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

/// WEBSERVICE(url)
pub fn fn_webservice(
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
pub fn fn_rtd(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }
    Ok(FormulaValue::Error(CellError::Na))
}

/// IMAGE(...)
pub fn fn_image(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }
    Ok(FormulaValue::Error(CellError::Na))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
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
        let arr = FormulaValue::Array(vec![vec![n(1.0), n(2.0)], vec![n(3.0), n(4.0)]]);
        assert_eq!(
            fn_choosecols(&[arr.clone(), n(2.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array(vec![vec![n(2.0)], vec![n(4.0)]])
        );
        assert_eq!(
            fn_chooserows(&[arr.clone(), n(-1.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array(vec![vec![n(3.0), n(4.0)]])
        );
        assert_eq!(
            fn_take(&[arr.clone(), n(1.0), n(1.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array(vec![vec![n(1.0)]])
        );
        assert_eq!(
            fn_drop(&[arr.clone(), n(1.0), n(0.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array(vec![vec![n(3.0), n(4.0)]])
        );
        assert_eq!(
            fn_expand(&[arr, n(3.0), n(3.0), s("x")], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array(vec![
                vec![n(1.0), n(2.0), s("x")],
                vec![n(3.0), n(4.0), s("x")],
                vec![s("x"), s("x"), s("x")]
            ])
        );
    }

    #[test]
    fn test_filter_stack_and_transforms() {
        let arr = FormulaValue::Array(vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)]]);
        let include = FormulaValue::Array(vec![
            vec![FormulaValue::Boolean(false)],
            vec![FormulaValue::Boolean(true)],
            vec![FormulaValue::Boolean(true)],
        ]);
        assert_eq!(
            fn_filter(&[arr.clone(), include], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array(vec![vec![n(2.0)], vec![n(3.0)]])
        );

        let h = fn_hstack(
            &[
                FormulaValue::Array(vec![vec![n(1.0)], vec![n(2.0)]]),
                FormulaValue::Array(vec![vec![n(10.0)], vec![n(20.0)]]),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            h,
            FormulaValue::Array(vec![vec![n(1.0), n(10.0)], vec![n(2.0), n(20.0)]])
        );

        let t = fn_transpose(&[h], &EvaluationContext::simple()).unwrap();
        assert_eq!(
            t,
            FormulaValue::Array(vec![vec![n(1.0), n(2.0)], vec![n(10.0), n(20.0)]])
        );
    }

    #[test]
    fn test_lookup_sort_unique_wrap() {
        let lookup = fn_lookup(
            &[
                n(2.5),
                FormulaValue::Array(vec![vec![n(1.0), n(2.0), n(3.0)]]),
                FormulaValue::Array(vec![vec![s("a"), s("b"), s("c")]]),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(lookup, s("b"));

        let sorted = fn_sort(
            &[
                FormulaValue::Array(vec![vec![n(2.0)], vec![n(1.0)]]),
                n(1.0),
                n(1.0),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            sorted,
            FormulaValue::Array(vec![vec![n(1.0)], vec![n(2.0)]])
        );

        let uniq = fn_unique(
            &[FormulaValue::Array(vec![
                vec![n(1.0)],
                vec![n(1.0)],
                vec![n(2.0)],
            ])],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(uniq, FormulaValue::Array(vec![vec![n(1.0)], vec![n(2.0)]]));

        let wrapped = fn_wraprows(
            &[
                FormulaValue::Array(vec![vec![n(1.0), n(2.0), n(3.0)]]),
                n(2.0),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            wrapped,
            FormulaValue::Array(vec![
                vec![n(1.0), n(2.0)],
                vec![n(3.0), FormulaValue::Error(CellError::Na)]
            ])
        );
    }

    #[test]
    fn test_tocol_torow_sortby_vstack() {
        let arr = FormulaValue::Array(vec![
            vec![n(1.0), n(2.0)],
            vec![n(3.0), FormulaValue::Empty],
        ]);
        assert_eq!(
            fn_tocol(&[arr.clone(), n(1.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array(vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)]])
        );
        assert_eq!(
            fn_torow(&[arr, n(1.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Array(vec![vec![n(1.0), n(2.0), n(3.0)]])
        );

        let sorted = fn_sortby(
            &[
                FormulaValue::Array(vec![vec![s("b")], vec![s("a")]]),
                FormulaValue::Array(vec![vec![n(2.0)], vec![n(1.0)]]),
                n(1.0),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            sorted,
            FormulaValue::Array(vec![vec![s("a")], vec![s("b")]])
        );

        let v = fn_vstack(
            &[
                FormulaValue::Array(vec![vec![n(1.0)]]),
                FormulaValue::Array(vec![vec![n(2.0), n(3.0)]]),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            v,
            FormulaValue::Array(vec![
                vec![n(1.0), FormulaValue::Error(CellError::Na)],
                vec![n(2.0), n(3.0)]
            ])
        );
    }

    #[test]
    fn test_date_intl_and_web_stubs() {
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
            FormulaValue::Error(CellError::Na)
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
            FormulaValue::Error(CellError::Na)
        );

        let one = eval("=1").unwrap();
        assert_eq!(one, n(1.0));
    }
}
