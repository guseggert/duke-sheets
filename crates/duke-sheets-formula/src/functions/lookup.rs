//! Lookup functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;
use std::cmp::Ordering;

fn to_i64_trunc(v: &FormulaValue) -> Option<i64> {
    v.as_number().map(|n| n.trunc() as i64)
}

fn values_equal(a: &FormulaValue, b: &FormulaValue) -> bool {
    match (a, b) {
        (FormulaValue::Number(x), FormulaValue::Number(y)) => x == y,
        (FormulaValue::Boolean(x), FormulaValue::Boolean(y)) => x == y,
        (FormulaValue::String(x), FormulaValue::String(y)) => x.eq_ignore_ascii_case(y),

        // Try numeric coercion between string/number
        (FormulaValue::Number(x), FormulaValue::String(s))
        | (FormulaValue::String(s), FormulaValue::Number(x)) => {
            s.parse::<f64>().ok().map(|n| n == *x).unwrap_or(false)
        }

        // Empty coercions
        (FormulaValue::Empty, FormulaValue::Empty) => true,
        (FormulaValue::Empty, FormulaValue::Number(n))
        | (FormulaValue::Number(n), FormulaValue::Empty) => *n == 0.0,
        (FormulaValue::Empty, FormulaValue::String(s))
        | (FormulaValue::String(s), FormulaValue::Empty) => s.is_empty(),

        _ => false,
    }
}

fn expect_array<'a>(v: &'a FormulaValue) -> Option<&'a Vec<Vec<FormulaValue>>> {
    match v {
        FormulaValue::Array(a) => Some(a),
        _ => None,
    }
}

fn array_dims(arr: &[Vec<FormulaValue>]) -> (usize, usize) {
    let rows = arr.len();
    let cols = arr.first().map(|r| r.len()).unwrap_or(0);
    (rows, cols)
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

fn wildcard_match(pattern: &str, text: &str) -> bool {
    let p: Vec<char> = pattern.chars().collect();
    let t: Vec<char> = text.chars().collect();
    let mut dp = vec![vec![false; t.len() + 1]; p.len() + 1];
    dp[0][0] = true;
    for i in 1..=p.len() {
        if p[i - 1] == '*' {
            dp[i][0] = dp[i - 1][0];
        }
    }
    for i in 1..=p.len() {
        for j in 1..=t.len() {
            if p[i - 1] == '*' {
                dp[i][j] = dp[i - 1][j] || dp[i][j - 1];
            } else if p[i - 1] == '?' || p[i - 1].eq_ignore_ascii_case(&t[j - 1]) {
                dp[i][j] = dp[i - 1][j - 1];
            }
        }
    }
    dp[p.len()][t.len()]
}

fn vector_from_array(arr: &[Vec<FormulaValue>]) -> Option<Vec<FormulaValue>> {
    let (rows, cols) = array_dims(arr);
    if rows == 0 || cols == 0 {
        return None;
    }
    if rows == 1 {
        return Some(arr[0].clone());
    }
    if cols == 1 {
        return Some(
            arr.iter()
                .map(|r| r.first().cloned().unwrap_or(FormulaValue::Empty))
                .collect(),
        );
    }
    None
}

fn parse_column_letters(s: &str) -> Option<u16> {
    let upper = s.to_ascii_uppercase();
    let mut col: u16 = 0;
    for c in upper.chars() {
        if !c.is_ascii_uppercase() {
            return None;
        }
        col = col
            .checked_mul(26)?
            .checked_add((c as u16) - ('A' as u16) + 1)?;
    }
    Some(col - 1)
}

fn parse_cell_address(addr: &str) -> Option<(u32, u16)> {
    let col_end = addr
        .find(|c: char| c.is_ascii_digit())
        .unwrap_or(addr.len());
    if col_end == 0 || col_end == addr.len() {
        return None;
    }
    let col = parse_column_letters(&addr[..col_end])?;
    let row: u32 = addr[col_end..].parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((row - 1, col))
}

/// INDEX(array, row_num, [column_num])
pub fn fn_index(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Propagate lookup errors in arguments
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    let arr = match expect_array(args.get(0).unwrap()) {
        Some(a) => a,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let (rows, cols) = array_dims(arr);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Error(CellError::Ref));
    }

    let row_num = to_i64_trunc(args.get(1).unwrap()).unwrap_or(0);
    if row_num < 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let col_num = match args.get(2) {
        Some(v) => to_i64_trunc(v).unwrap_or(0),
        None => 1,
    };
    if col_num < 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let r = (row_num - 1) as usize;
    let c = (col_num - 1) as usize;
    if r >= rows || c >= cols {
        return Ok(FormulaValue::Error(CellError::Ref));
    }
    Ok(arr[r][c].clone())
}

/// MATCH(lookup_value, lookup_array, [match_type])
///
/// Currently supports exact match only (match_type = 0). Other match types return #N/A.
pub fn fn_match(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let lookup_value = args.get(0).unwrap();
    if let FormulaValue::Error(e) = lookup_value {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(lookup_value, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let arr = match expect_array(args.get(1).unwrap()) {
        Some(a) => a,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let (rows, cols) = array_dims(arr);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    let match_type = match args.get(2) {
        None => 0,
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            to_i64_trunc(v).unwrap_or(0)
        }
    };

    if match_type != 0 {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    // MATCH expects a vector (single row or single column)
    if rows == 1 {
        for (i, v) in arr[0].iter().enumerate() {
            if values_equal(lookup_value, v) {
                return Ok(FormulaValue::Number((i + 1) as f64));
            }
        }
    } else if cols == 1 {
        for (i, row) in arr.iter().enumerate() {
            let v = row.get(0).unwrap_or(&FormulaValue::Empty);
            if values_equal(lookup_value, v) {
                return Ok(FormulaValue::Number((i + 1) as f64));
            }
        }
    } else {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    Ok(FormulaValue::Error(CellError::Na))
}

/// ROWS(array) - Returns the number of rows in a reference or array
/// Reference: LibreOffice ScInterpreter::ScRows, Microsoft ROWS function
pub fn fn_rows(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arg = args.get(0).unwrap_or(&FormulaValue::Empty);

    match arg {
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        FormulaValue::Array(arr) => {
            let rows = arr.len();
            Ok(FormulaValue::Number(rows as f64))
        }
        // Single value = 1 row
        _ => Ok(FormulaValue::Number(1.0)),
    }
}

/// COLUMNS(array) - Returns the number of columns in a reference or array
/// Reference: LibreOffice ScInterpreter::ScColumns, Microsoft COLUMNS function
pub fn fn_columns(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arg = args.get(0).unwrap_or(&FormulaValue::Empty);

    match arg {
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        FormulaValue::Array(arr) => {
            let cols = arr.first().map(|r| r.len()).unwrap_or(0);
            Ok(FormulaValue::Number(cols as f64))
        }
        // Single value = 1 column
        _ => Ok(FormulaValue::Number(1.0)),
    }
}

/// CHOOSE(index_num, value1, [value2], ...) - Returns a value from a list based on index
/// Reference: LibreOffice ScInterpreter::ScChooseJump, Microsoft CHOOSE function
///
/// index_num is 1-based and floored (2.9 -> 2)
/// Returns #VALUE! if index is out of range
pub fn fn_choose(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Get and validate index
    let index_arg = &args[0];
    if let FormulaValue::Error(e) = index_arg {
        return Ok(FormulaValue::Error(*e));
    }

    let index = match to_i64_trunc(index_arg) {
        Some(i) => i,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Index must be >= 1 and <= number of values
    let num_values = args.len() - 1; // exclude the index argument
    if index < 1 || index as usize > num_values {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Return the selected value (1-based index)
    Ok(args[index as usize].clone())
}

/// ROW([reference]) - Returns the row number of a reference
///
/// Reference: LibreOffice ScInterpreter::ScRow, Microsoft ROW function
///
/// - ROW() with no args returns the row of the current cell (1-indexed)
/// - ROW(reference) returns the row number of the reference
/// - ROW(range) returns an array of row numbers
///
/// Note: Since arguments are evaluated before reaching the function, we can't
/// distinguish between ROW(A1) and ROW(5) - both arrive as the value. For ranges,
/// we return an array of sequential row numbers based on array dimensions.
pub fn fn_row(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.is_empty() {
        // ROW() with no args - return current row (1-indexed like Excel)
        return Ok(FormulaValue::Number((ctx.current_row + 1) as f64));
    }

    let arg = &args[0];
    match arg {
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        FormulaValue::Array(arr) => {
            // For an array, return a column vector of row numbers
            // This matches Excel's behavior: ROW(A1:A5) returns {1;2;3;4;5}
            let rows = arr.len();
            if rows == 0 {
                return Ok(FormulaValue::Error(CellError::Value));
            }

            // Return column vector (rows x 1) with row indices
            // The actual row numbers should come from the reference, but since
            // we don't have that info, we return 1..n as relative row indices
            // This is a limitation - full support needs unevaluated reference passing
            let result: Vec<Vec<FormulaValue>> = (0..rows)
                .map(|i| {
                    vec![FormulaValue::Number(
                        (ctx.current_row + 1 + i as u32) as f64,
                    )]
                })
                .collect();
            Ok(FormulaValue::Array(result))
        }
        // Single value = assume it's from the current cell context
        _ => Ok(FormulaValue::Number((ctx.current_row + 1) as f64)),
    }
}

/// COLUMN([reference]) - Returns the column number of a reference
///
/// Reference: LibreOffice ScInterpreter::ScColumn, Microsoft COLUMN function
///
/// - COLUMN() with no args returns the column of the current cell (1-indexed)
/// - COLUMN(reference) returns the column number of the reference
/// - COLUMN(range) returns an array of column numbers
///
/// Note: Since arguments are evaluated before reaching the function, we can't
/// distinguish between COLUMN(A1) and COLUMN(1) - both arrive as the value. For ranges,
/// we return an array of sequential column numbers based on array dimensions.
pub fn fn_column(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.is_empty() {
        // COLUMN() with no args - return current column (1-indexed like Excel)
        return Ok(FormulaValue::Number((ctx.current_col + 1) as f64));
    }

    let arg = &args[0];
    match arg {
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        FormulaValue::Array(arr) => {
            // For an array, return a row vector of column numbers
            // This matches Excel's behavior: COLUMN(A1:E1) returns {1,2,3,4,5}
            let cols = arr.first().map(|r| r.len()).unwrap_or(0);
            if cols == 0 {
                return Ok(FormulaValue::Error(CellError::Value));
            }

            // Return row vector (1 x cols) with column indices
            // The actual column numbers should come from the reference, but since
            // we don't have that info, we return 1..n as relative column indices
            let result: Vec<FormulaValue> = (0..cols)
                .map(|i| FormulaValue::Number((ctx.current_col + 1 + i as u16) as f64))
                .collect();
            Ok(FormulaValue::Array(vec![result]))
        }
        // Single value = assume it's from the current cell context
        _ => Ok(FormulaValue::Number((ctx.current_col + 1) as f64)),
    }
}

/// VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
///
/// Currently implements exact match only. If range_lookup is TRUE or omitted, we still
/// perform exact match (no approximate matching).
pub fn fn_vlookup(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Propagate errors in arguments
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    let lookup_value = args.get(0).unwrap();
    if matches!(lookup_value, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let table = match expect_array(args.get(1).unwrap()) {
        Some(a) => a,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let (rows, cols) = array_dims(table);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    let col_index = to_i64_trunc(args.get(2).unwrap()).unwrap_or(0);
    if col_index < 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let col_index0 = (col_index - 1) as usize;
    if col_index0 >= cols {
        return Ok(FormulaValue::Error(CellError::Ref));
    }

    // range_lookup (ignored for now; exact match only)
    if let Some(v) = args.get(3) {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    for row in table {
        let key = row.get(0).unwrap_or(&FormulaValue::Empty);
        if values_equal(lookup_value, key) {
            return Ok(row.get(col_index0).cloned().unwrap_or(FormulaValue::Empty));
        }
    }

    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_hlookup(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    let lookup_value = args.get(0).unwrap();
    if matches!(lookup_value, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let table = match expect_array(args.get(1).unwrap()) {
        Some(a) => a,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let (rows, cols) = array_dims(table);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    let row_index = to_i64_trunc(args.get(2).unwrap()).unwrap_or(0);
    if row_index < 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let row_index0 = (row_index - 1) as usize;
    if row_index0 >= rows {
        return Ok(FormulaValue::Error(CellError::Ref));
    }

    let range_lookup = args
        .get(3)
        .map(|v| v.as_bool().unwrap_or(false))
        .unwrap_or(true);

    let first_row = &table[0];

    let col_idx = if range_lookup {
        let mut is_sorted = true;
        for i in 1..first_row.len() {
            let prev = &first_row[i - 1];
            let curr = &first_row[i];
            match compare_lookup_values(prev, curr) {
                Some(Ordering::Greater) | None => {
                    is_sorted = false;
                    break;
                }
                _ => {}
            }
        }
        if !is_sorted {
            return Ok(FormulaValue::Error(CellError::Na));
        }

        let mut lo = 0usize;
        let mut hi = first_row.len();
        let mut best: Option<usize> = None;
        while lo < hi {
            let mid = lo + (hi - lo) / 2;
            match compare_lookup_values(&first_row[mid], lookup_value) {
                Some(Ordering::Equal) => {
                    best = Some(mid);
                    break;
                }
                Some(Ordering::Less) => {
                    best = Some(mid);
                    lo = mid + 1;
                }
                Some(Ordering::Greater) => {
                    hi = mid;
                }
                None => return Ok(FormulaValue::Error(CellError::Na)),
            }
        }
        match best {
            Some(i) => i,
            None => return Ok(FormulaValue::Error(CellError::Na)),
        }
    } else {
        match first_row.iter().position(|k| values_equal(lookup_value, k)) {
            Some(i) => i,
            None => return Ok(FormulaValue::Error(CellError::Na)),
        }
    };

    Ok(table
        .get(row_index0)
        .and_then(|r| r.get(col_idx))
        .cloned()
        .unwrap_or(FormulaValue::Empty))
}

pub fn fn_xmatch(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    let lookup_value = args.get(0).unwrap();
    if matches!(lookup_value, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let lookup_arr = match expect_array(args.get(1).unwrap()).and_then(|a| vector_from_array(a)) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let match_mode = args.get(2).and_then(to_i64_trunc).unwrap_or(0);
    let search_mode = args.get(3).and_then(to_i64_trunc).unwrap_or(1);

    let indices: Vec<usize> = if search_mode == -1 {
        (0..lookup_arr.len()).rev().collect()
    } else {
        (0..lookup_arr.len()).collect()
    };

    let exact_pos = || {
        indices
            .iter()
            .copied()
            .find(|&i| values_equal(lookup_value, &lookup_arr[i]))
    };

    let result = match match_mode {
        0 => exact_pos(),
        2 => {
            let pattern = lookup_value.as_string();
            indices.iter().copied().find(|&i| {
                wildcard_match(
                    &pattern.to_ascii_lowercase(),
                    &lookup_arr[i].as_string().to_ascii_lowercase(),
                )
            })
        }
        -1 => {
            if let Some(i) = exact_pos() {
                Some(i)
            } else {
                let mut best: Option<usize> = None;
                for (i, v) in lookup_arr.iter().enumerate() {
                    if let Some(ord) = compare_lookup_values(v, lookup_value) {
                        if ord == Ordering::Less {
                            if let Some(prev) = best {
                                if compare_lookup_values(v, &lookup_arr[prev])
                                    == Some(Ordering::Greater)
                                {
                                    best = Some(i);
                                }
                            } else {
                                best = Some(i);
                            }
                        }
                    }
                }
                best
            }
        }
        1 => {
            if let Some(i) = exact_pos() {
                Some(i)
            } else {
                let mut best: Option<usize> = None;
                for (i, v) in lookup_arr.iter().enumerate() {
                    if let Some(ord) = compare_lookup_values(v, lookup_value) {
                        if ord == Ordering::Greater {
                            if let Some(prev) = best {
                                if compare_lookup_values(v, &lookup_arr[prev])
                                    == Some(Ordering::Less)
                                {
                                    best = Some(i);
                                }
                            } else {
                                best = Some(i);
                            }
                        }
                    }
                }
                best
            }
        }
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    match result {
        Some(i) => Ok(FormulaValue::Number((i + 1) as f64)),
        None => Ok(FormulaValue::Error(CellError::Na)),
    }
}

pub fn fn_xlookup(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    let lookup_value = args.get(0).unwrap().clone();
    let lookup_array = args.get(1).unwrap().clone();
    let return_arr = match expect_array(args.get(2).unwrap()).and_then(|a| vector_from_array(a)) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let lookup_vec = match expect_array(&lookup_array).and_then(|a| vector_from_array(a)) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if lookup_vec.len() != return_arr.len() {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let if_not_found = args
        .get(3)
        .cloned()
        .unwrap_or(FormulaValue::Error(CellError::Na));
    let match_mode = args.get(4).cloned().unwrap_or(FormulaValue::Number(0.0));
    let search_mode = args.get(5).cloned().unwrap_or(FormulaValue::Number(1.0));

    let lookup_array_value = FormulaValue::Array(vec![lookup_vec]);
    let xmatch_result = fn_xmatch(
        &[lookup_value, lookup_array_value, match_mode, search_mode],
        ctx,
    )?;

    match xmatch_result {
        FormulaValue::Number(pos) => {
            let idx = pos as usize;
            if idx == 0 || idx > return_arr.len() {
                Ok(FormulaValue::Error(CellError::Na))
            } else {
                Ok(return_arr[idx - 1].clone())
            }
        }
        FormulaValue::Error(CellError::Na) => Ok(if_not_found),
        FormulaValue::Error(e) => Ok(FormulaValue::Error(e)),
        _ => Ok(FormulaValue::Error(CellError::Na)),
    }
}

pub fn fn_indirect(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    if ctx.workbook.is_none() {
        return Ok(FormulaValue::Error(CellError::Ref));
    }

    let ref_text = args.get(0).unwrap();
    if matches!(ref_text, FormulaValue::Array(_)) {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let a1_style = args
        .get(1)
        .map(|v| v.as_bool().unwrap_or(false))
        .unwrap_or(true);
    if !a1_style {
        return Ok(FormulaValue::Error(CellError::Ref));
    }

    let text = ref_text.as_string();
    let (sheet_name, ref_part) = if let Some(pos) = text.rfind('!') {
        let raw_sheet = text[..pos].trim();
        let clean_sheet = raw_sheet.trim_matches('\'').to_string();
        (Some(clean_sheet), text[pos + 1..].trim().to_string())
    } else {
        (None, text.trim().to_string())
    };

    let ref_clean = ref_part.replace('$', "");
    if let Some(colon) = ref_clean.find(':') {
        let start = &ref_clean[..colon];
        let end = &ref_clean[colon + 1..];
        let (sr, sc) = match parse_cell_address(start) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Ref)),
        };
        let (er, ec) = match parse_cell_address(end) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Ref)),
        };
        Ok(ctx.get_range_values(sheet_name.as_deref(), sr, sc, er, ec))
    } else {
        let (row, col) = match parse_cell_address(&ref_clean) {
            Some(v) => v,
            None => return Ok(FormulaValue::Error(CellError::Ref)),
        };
        Ok(ctx.get_cell_value(sheet_name.as_deref(), row, col))
    }
}

pub fn fn_offset(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    if args.get(1).and_then(to_i64_trunc).is_none() || args.get(2).and_then(to_i64_trunc).is_none()
    {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    if let Some(h) = args.get(3).and_then(to_i64_trunc) {
        if h < 1 {
            return Ok(FormulaValue::Error(CellError::Value));
        }
    }
    if let Some(w) = args.get(4).and_then(to_i64_trunc) {
        if w < 1 {
            return Ok(FormulaValue::Error(CellError::Value));
        }
    }

    // TODO: full OFFSET behavior requires evaluator support for unevaluated references.
    Ok(FormulaValue::Error(CellError::Ref))
}

/// SEQUENCE(rows, [columns], [start], [step]) - Generates a sequence of numbers
///
/// Reference: Microsoft SEQUENCE function, LibreOffice SEQUENCE
///
/// - rows: Number of rows to return (required, must be >= 1)
/// - columns: Number of columns to return (default 1, must be >= 1)  
/// - start: Starting number (default 1)
/// - step: Increment between numbers (default 1)
///
/// Returns a 2D array filled with sequential numbers.
/// The array is filled row by row (left to right, top to bottom).
pub fn fn_sequence(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Check for errors in arguments
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    // rows (required)
    let rows = match args.get(0) {
        Some(v) => match to_i64_trunc(v) {
            Some(r) if r >= 1 => r as usize,
            Some(_) => return Ok(FormulaValue::Error(CellError::Value)),
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // columns (default 1)
    let cols = match args.get(1) {
        Some(v) if !matches!(v, FormulaValue::Empty) => match to_i64_trunc(v) {
            Some(c) if c >= 1 => c as usize,
            Some(_) => return Ok(FormulaValue::Error(CellError::Value)),
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        _ => 1,
    };

    // start (default 1)
    let start = match args.get(2) {
        Some(v) if !matches!(v, FormulaValue::Empty) => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        _ => 1.0,
    };

    // step (default 1)
    let step = match args.get(3) {
        Some(v) if !matches!(v, FormulaValue::Empty) => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        _ => 1.0,
    };

    // Limit array size to prevent memory issues (Excel limits to 1,048,576 rows × 16,384 columns)
    const MAX_CELLS: usize = 1_000_000;
    if rows * cols > MAX_CELLS {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Generate the sequence array
    let mut result: Vec<Vec<FormulaValue>> = Vec::with_capacity(rows);
    let mut current = start;

    for _ in 0..rows {
        let mut row: Vec<FormulaValue> = Vec::with_capacity(cols);
        for _ in 0..cols {
            row.push(FormulaValue::Number(current));
            current += step;
        }
        result.push(row);
    }

    Ok(FormulaValue::Array(result))
}

#[cfg(test)]
mod tests {
    use super::*;
    use duke_sheets_core::{CellValue, Workbook};

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    #[test]
    fn test_hlookup() {
        assert_eq!(
            eval("=HLOOKUP(2,{1,2,3;\"a\",\"b\",\"c\"},2,FALSE)").unwrap(),
            FormulaValue::String("b".into())
        );
        assert_eq!(
            eval("=HLOOKUP(2.5,{1,2,3;10,20,30},2,TRUE)").unwrap(),
            FormulaValue::Number(20.0)
        );
        assert_eq!(
            eval("=HLOOKUP(9,{1,2,3;10,20,30},2,FALSE)").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_xmatch() {
        assert_eq!(
            eval("=XMATCH(3,{1,2,3,4})").unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            eval("=XMATCH(\"ab*\",{\"aa\",\"abc\",\"def\"},2)").unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            eval("=XMATCH(2.5,{1,2,3},-1)").unwrap(),
            FormulaValue::Number(2.0)
        );
    }

    #[test]
    fn test_xlookup() {
        assert_eq!(
            eval("=XLOOKUP(2,{1,2,3},{\"a\",\"b\",\"c\"})").unwrap(),
            FormulaValue::String("b".into())
        );
        assert_eq!(
            eval("=XLOOKUP(4,{1,2,3},{\"a\",\"b\",\"c\"},\"NF\")").unwrap(),
            FormulaValue::String("NF".into())
        );
        // match_mode=1 (next larger): 2.5 not found, next larger is 3 → return 30
        // Note: ideally we'd test =XLOOKUP(2.5,...,,1) with an omitted 4th arg,
        // but the parser doesn't yet support empty arguments (,,). Using explicit
        // if_not_found here; the lookup succeeds so it's never reached.
        assert_eq!(
            eval("=XLOOKUP(2.5,{1,2,3},{10,20,30},\"NF\",1)").unwrap(),
            FormulaValue::Number(30.0)
        );
        // Default if_not_found is #N/A (tested via 4-arg form)
        assert_eq!(
            eval("=XLOOKUP(99,{1,2,3},{10,20,30})").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_indirect() {
        assert_eq!(
            eval("=INDIRECT(\"A1\")").unwrap(),
            FormulaValue::Error(CellError::Ref)
        );

        let mut workbook = Workbook::new();
        {
            let sheet = workbook.worksheet_mut(0).unwrap();
            sheet
                .set_cell_value_at(0, 0, CellValue::Number(42.0))
                .unwrap();
            sheet
                .set_cell_value_at(0, 1, CellValue::String("x".into()))
                .unwrap();
        }
        let ctx = EvaluationContext::new(Some(&workbook), 0, 0, 0);
        assert_eq!(
            fn_indirect(&[FormulaValue::String("A1".into())], &ctx).unwrap(),
            FormulaValue::Number(42.0)
        );
        assert_eq!(
            fn_indirect(&[FormulaValue::String("A1:B1".into())], &ctx).unwrap(),
            FormulaValue::Array(vec![vec![
                FormulaValue::Number(42.0),
                FormulaValue::String("x".into())
            ]])
        );
    }

    #[test]
    fn test_offset() {
        assert_eq!(
            eval("=OFFSET(1,1,1)").unwrap(),
            FormulaValue::Error(CellError::Ref)
        );
        assert_eq!(
            eval("=OFFSET(1,\"x\",1)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            eval("=OFFSET(1,1,1,0,1)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }
}
