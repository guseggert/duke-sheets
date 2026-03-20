//! Lookup functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;
use std::cmp::Ordering;

pub(crate) fn to_i64_trunc(v: &FormulaValue) -> Option<i64> {
    v.as_number().map(|n| n.trunc() as i64)
}

pub(crate) fn values_equal(a: &FormulaValue, b: &FormulaValue) -> bool {
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

fn expect_array(v: &FormulaValue) -> Option<&Vec<Vec<FormulaValue>>> {
    match v {
        FormulaValue::Array { data: a, .. } => Some(a),
        _ => None,
    }
}

fn array_dims(arr: &[Vec<FormulaValue>]) -> (usize, usize) {
    let rows = arr.len();
    let cols = arr.first().map(|r| r.len()).unwrap_or(0);
    (rows, cols)
}

/// Resolve the effective (row_num, col_num) for INDEX given the raw second
/// argument, an optional third argument, and the array dimensions.
///
/// In the two-argument form `INDEX(vector, position)` Excel treats the
/// position as an index into the vector's non-trivial dimension:
///   - single row  → position selects the column  (row fixed at 1)
///   - single col  → position selects the row     (col fixed at 1)
///   - position=0  → "return the whole vector"
///
/// In the three-argument form the values are used directly.
pub(crate) fn index_resolve_coords(
    raw_pos: i64,
    explicit_col: Option<i64>,
    rows: usize,
    cols: usize,
) -> (i64, i64) {
    match explicit_col {
        Some(c) => (raw_pos, c),
        None if rows == 1 => (1, raw_pos),  // pos=0 → (1,0) → return entire row
        None if cols == 1 => (raw_pos, 1),  // pos=0 → (0,1) → return entire column
        None => (raw_pos, 0),               // 2D, no col → return entire row
    }
}

pub(crate) fn compare_lookup_values(a: &FormulaValue, b: &FormulaValue) -> Option<Ordering> {
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

pub(crate) fn wildcard_match(pattern: &str, text: &str) -> bool {
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

    let arr = match expect_array(args.first().unwrap()) {
        Some(a) => a,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let (rows, cols) = array_dims(arr);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Error(CellError::Ref));
    }

    let raw_pos = to_i64_trunc(args.get(1).unwrap()).unwrap_or(0);
    if raw_pos < 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let explicit_col = args.get(2).map(|v| to_i64_trunc(v).unwrap_or(0));
    let (row_num, col_num) = index_resolve_coords(raw_pos, explicit_col, rows, cols);
    if col_num < 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // row_num=0: return entire column as a column vector
    if row_num == 0 {
        if col_num == 0 {
            return Ok(FormulaValue::Error(CellError::Value));
        }
        let c = (col_num - 1) as usize;
        if c >= cols {
            return Ok(FormulaValue::Error(CellError::Ref));
        }
        let column: Vec<Vec<FormulaValue>> = arr.iter().map(|row| vec![row[c].clone()]).collect();
        return Ok(FormulaValue::Array {
            data: column,
            source: None,
        });
    }

    // col_num=0: return entire row as a row vector
    if col_num == 0 {
        let r = (row_num - 1) as usize;
        if r >= rows {
            return Ok(FormulaValue::Error(CellError::Ref));
        }
        return Ok(FormulaValue::Array {
            data: vec![arr[r].clone()],
            source: None,
        });
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
/// match_type: 1 = largest value <= lookup_value (array must be ascending)
///             0 = exact match (default)
///            -1 = smallest value >= lookup_value (array must be descending)
pub fn fn_match(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let lookup_value = args.first().unwrap();
    if let FormulaValue::Error(e) = lookup_value {
        return Ok(FormulaValue::Error(*e));
    }
    if matches!(lookup_value, FormulaValue::Array { .. }) {
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
        None => 1,
        Some(v) => {
            if let FormulaValue::Error(e) = v {
                return Ok(FormulaValue::Error(*e));
            }
            let mt = to_i64_trunc(v).unwrap_or(1);
            if mt > 0 {
                1
            } else if mt < 0 {
                -1
            } else {
                0
            }
        }
    };

    // Extract the vector from the array
    let vec: Vec<&FormulaValue> = if rows == 1 {
        arr[0].iter().collect()
    } else if cols == 1 {
        arr.iter()
            .map(|row| row.first().unwrap_or(&FormulaValue::Empty))
            .collect()
    } else {
        return Ok(FormulaValue::Error(CellError::Na));
    };

    match match_type {
        0 => {
            // Exact match: try hash index (Tier 2) first
            if let Some(cache) = ctx.eval_cache {
                // Get the RangeSource from the array argument to build the cache key
                let array_arg = args.get(1).unwrap();
                if let FormulaValue::Array {
                    source: Some(src), ..
                } = array_arg
                {
                    let col_offset = if rows == 1 { 0u16 } else { 0u16 }; // single column/row
                    let range_key = (
                        src.sheet,
                        src.start_row,
                        src.start_col,
                        src.end_row,
                        src.end_col,
                    );
                    let index_key = (range_key, col_offset);

                    // Check for existing index
                    if let Some(existing) = cache.lookup_indexes.get(&index_key) {
                        return match existing.find(lookup_value) {
                            Some(pos) => Ok(FormulaValue::Number((pos + 1) as f64)),
                            None => Ok(FormulaValue::Error(CellError::Na)),
                        };
                    }

                    // Build index from the 1D vector
                    let owned_vec: Vec<FormulaValue> = vec.iter().map(|v| (*v).clone()).collect();
                    let index = crate::eval_cache::LookupIndex::build(&owned_vec);
                    let result = match index.find(lookup_value) {
                        Some(pos) => Ok(FormulaValue::Number((pos + 1) as f64)),
                        None => Ok(FormulaValue::Error(CellError::Na)),
                    };
                    cache
                        .lookup_indexes
                        .insert(index_key, std::sync::Arc::new(index));
                    return result;
                }
            }

            // Fallback: linear scan (no cache or no source metadata)
            for (i, v) in vec.iter().enumerate() {
                if values_equal(lookup_value, v) {
                    return Ok(FormulaValue::Number((i + 1) as f64));
                }
            }
            Ok(FormulaValue::Error(CellError::Na))
        }
        1 => {
            // Largest value <= lookup_value (array should be ascending)
            // Linear scan: find last value <= lookup_value
            let mut best: Option<usize> = None;
            for (i, v) in vec.iter().enumerate() {
                match compare_lookup_values(v, lookup_value) {
                    Some(Ordering::Less) | Some(Ordering::Equal) => {
                        best = Some(i);
                    }
                    Some(Ordering::Greater) => break,
                    None => {}
                }
            }
            match best {
                Some(i) => Ok(FormulaValue::Number((i + 1) as f64)),
                None => Ok(FormulaValue::Error(CellError::Na)),
            }
        }
        -1 => {
            // Smallest value >= lookup_value (array should be descending)
            // Linear scan: find last value >= lookup_value
            let mut best: Option<usize> = None;
            for (i, v) in vec.iter().enumerate() {
                match compare_lookup_values(v, lookup_value) {
                    Some(Ordering::Greater) | Some(Ordering::Equal) => {
                        best = Some(i);
                    }
                    Some(Ordering::Less) => break,
                    None => {}
                }
            }
            match best {
                Some(i) => Ok(FormulaValue::Number((i + 1) as f64)),
                None => Ok(FormulaValue::Error(CellError::Na)),
            }
        }
        _ => Ok(FormulaValue::Error(CellError::Na)),
    }
}

/// ROWS(array) - Returns the number of rows in a reference or array
/// Reference: LibreOffice ScInterpreter::ScRows, Microsoft ROWS function
pub fn fn_rows(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let arg = args.first().unwrap_or(&FormulaValue::Empty);

    match arg {
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        FormulaValue::Array { data: arr, .. } => {
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
    let arg = args.first().unwrap_or(&FormulaValue::Empty);

    match arg {
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        FormulaValue::Array { data: arr, .. } => {
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
        FormulaValue::Array { data: arr, .. } => {
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
            Ok(FormulaValue::Array {
                data: result,
                source: None,
            })
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
        FormulaValue::Array { data: arr, .. } => {
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
            Ok(FormulaValue::Array {
                data: vec![result],
                source: None,
            })
        }
        // Single value = assume it's from the current cell context
        _ => Ok(FormulaValue::Number((ctx.current_col + 1) as f64)),
    }
}

/// VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
///
/// Currently implements exact match only. If range_lookup is TRUE or omitted, we still
/// perform exact match (no approximate matching).
pub fn fn_vlookup(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Propagate errors in arguments
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    let lookup_value = args.first().unwrap();
    if matches!(lookup_value, FormulaValue::Array { .. }) {
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

    if let Some(FormulaValue::Error(e)) = args.get(3) {
        return Ok(FormulaValue::Error(*e));
    }
    let range_lookup = args
        .get(3)
        .map(|v| v.as_bool().unwrap_or(false))
        .unwrap_or(true);

    if range_lookup {
        // Approximate match: binary search for largest value <= lookup_value
        // First column must be sorted ascending.
        let first_col: Vec<&FormulaValue> = table
            .iter()
            .map(|row| row.first().unwrap_or(&FormulaValue::Empty))
            .collect();

        let mut is_sorted = true;
        for i in 1..first_col.len() {
            match compare_lookup_values(first_col[i - 1], first_col[i]) {
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
        let mut hi = first_col.len();
        let mut best: Option<usize> = None;
        while lo < hi {
            let mid = lo + (hi - lo) / 2;
            match compare_lookup_values(first_col[mid], lookup_value) {
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
            Some(r) => Ok(table[r]
                .get(col_index0)
                .cloned()
                .unwrap_or(FormulaValue::Empty)),
            None => Ok(FormulaValue::Error(CellError::Na)),
        }
    } else {
        // Exact match
        // Tier 2: Try hash index for first column
        if let Some(cache) = ctx.eval_cache {
            let array_arg = args.get(1).unwrap();
            if let FormulaValue::Array {
                source: Some(src), ..
            } = array_arg
            {
                let range_key = (
                    src.sheet,
                    src.start_row,
                    src.start_col,
                    src.end_row,
                    src.end_col,
                );
                let index_key = (range_key, 0u16); // first column

                if let Some(existing) = cache.lookup_indexes.get(&index_key) {
                    return match existing.find(lookup_value) {
                        Some(pos) => Ok(table[pos]
                            .get(col_index0)
                            .cloned()
                            .unwrap_or(FormulaValue::Empty)),
                        None => Ok(FormulaValue::Error(CellError::Na)),
                    };
                }

                // Build index from first column
                let first_col: Vec<FormulaValue> = table
                    .iter()
                    .map(|row| row.first().cloned().unwrap_or(FormulaValue::Empty))
                    .collect();
                let index = crate::eval_cache::LookupIndex::build(&first_col);
                let result = match index.find(lookup_value) {
                    Some(pos) => Ok(table[pos]
                        .get(col_index0)
                        .cloned()
                        .unwrap_or(FormulaValue::Empty)),
                    None => Ok(FormulaValue::Error(CellError::Na)),
                };
                cache
                    .lookup_indexes
                    .insert(index_key, std::sync::Arc::new(index));
                return result;
            }
        }

        // Fallback: linear scan (exact match)
        for row in table {
            let key = row.first().unwrap_or(&FormulaValue::Empty);
            if values_equal(lookup_value, key) {
                return Ok(row.get(col_index0).cloned().unwrap_or(FormulaValue::Empty));
            }
        }

        Ok(FormulaValue::Error(CellError::Na))
    }
}

pub fn fn_hlookup(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    let lookup_value = args.first().unwrap();
    if matches!(lookup_value, FormulaValue::Array { .. }) {
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

pub fn fn_xmatch(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    for v in args {
        if let FormulaValue::Error(e) = v {
            return Ok(FormulaValue::Error(*e));
        }
    }

    let lookup_value = args.first().unwrap();
    if matches!(lookup_value, FormulaValue::Array { .. }) {
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
        // Tier 2: Try hash index for exact match
        if let Some(cache) = ctx.eval_cache {
            let array_arg = args.get(1).unwrap();
            if let FormulaValue::Array {
                source: Some(src), ..
            } = array_arg
            {
                let range_key = (
                    src.sheet,
                    src.start_row,
                    src.start_col,
                    src.end_row,
                    src.end_col,
                );
                let index_key = (range_key, 0u16);

                if let Some(existing) = cache.lookup_indexes.get(&index_key) {
                    return existing.find(lookup_value);
                }

                let owned: Vec<FormulaValue> = lookup_arr.iter().cloned().collect();
                let index = crate::eval_cache::LookupIndex::build(&owned);
                let result = index.find(lookup_value);
                cache
                    .lookup_indexes
                    .insert(index_key, std::sync::Arc::new(index));
                return result;
            }
        }
        // Fallback: linear scan
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

    let lookup_value = args.first().unwrap().clone();
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

    let lookup_array_value = FormulaValue::Array {
        data: vec![lookup_vec],
        source: None,
    };
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

    let ref_text = args.first().unwrap();
    if matches!(ref_text, FormulaValue::Array { .. }) {
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

/// OFFSET's reference-aware behavior is intercepted in `evaluator.rs`; this
/// fallback preserves direct-call coercion and error propagation semantics.
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
    let rows = match args.first() {
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

    Ok(FormulaValue::Array {
        data: result,
        source: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use duke_sheets_core::{CellValue, Workbook};

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    fn eval_with_workbook(formula: &str, workbook: &Workbook) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        let ctx = EvaluationContext::new(Some(workbook), 0, 0, 0);
        crate::evaluator::evaluate(&ast, &ctx)
    }

    fn n(x: f64) -> FormulaValue {
        FormulaValue::Number(x)
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
        assert_eq!(
            eval("=XLOOKUP(2.5,{1,2,3},{10,20,30},,1)").unwrap(),
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
        let result = fn_indirect(&[FormulaValue::String("A1:B1".into())], &ctx).unwrap();
        if let FormulaValue::Array { data, .. } = &result {
            assert_eq!(
                *data,
                vec![vec![
                    FormulaValue::Number(42.0),
                    FormulaValue::String("x".into())
                ]]
            );
        } else {
            panic!("Expected Array, got {:?}", result);
        }
    }

    #[test]
    fn test_indirect_docs() {
        // Microsoft docs: https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261
        //
        // Data layout (matching the docs example table):
        //   A2="B2", A3="B3", A5=5
        //   B2=1.333, B3=45, B5=62
        // We call fn_indirect directly with the resolved string values.

        let mut workbook = Workbook::new();
        {
            let sheet = workbook.worksheet_mut(0).unwrap();
            // B2 (row=1, col=1) = 1.333
            sheet
                .set_cell_value_at(1, 1, CellValue::Number(1.333))
                .unwrap();
            // B3 (row=2, col=1) = 45
            sheet
                .set_cell_value_at(2, 1, CellValue::Number(45.0))
                .unwrap();
            // B5 (row=4, col=1) = 62
            sheet
                .set_cell_value_at(4, 1, CellValue::Number(62.0))
                .unwrap();
            // A1 (row=0, col=0) = 99 (for range and sheet-ref tests)
            sheet
                .set_cell_value_at(0, 0, CellValue::Number(99.0))
                .unwrap();
            // B1 (row=0, col=1) = "hello" (for range test)
            sheet
                .set_cell_value_at(0, 1, CellValue::String("hello".into()))
                .unwrap();
        }
        let ctx = EvaluationContext::new(Some(&workbook), 0, 0, 0);

        // Docs Example 1: =INDIRECT(A2) where A2="B2", B2=1.333 → 1.333
        assert_eq!(
            fn_indirect(&[FormulaValue::String("B2".into())], &ctx).unwrap(),
            FormulaValue::Number(1.333)
        );

        // Docs Example 2: =INDIRECT(A3) where A3="B3", B3=45 → 45
        assert_eq!(
            fn_indirect(&[FormulaValue::String("B3".into())], &ctx).unwrap(),
            FormulaValue::Number(45.0)
        );

        // Docs Example 4: =INDIRECT("B"&A5) where A5=5 → INDIRECT("B5"), B5=62
        // (concatenation already resolved before fn_indirect is called)
        assert_eq!(
            fn_indirect(&[FormulaValue::String("B5".into())], &ctx).unwrap(),
            FormulaValue::Number(62.0)
        );

        // INDIRECT with range: INDIRECT("A1:B1") → array of values
        let result = fn_indirect(&[FormulaValue::String("A1:B1".into())], &ctx).unwrap();
        if let FormulaValue::Array { data, .. } = &result {
            assert_eq!(
                *data,
                vec![vec![
                    FormulaValue::Number(99.0),
                    FormulaValue::String("hello".into()),
                ]]
            );
        } else {
            panic!("Expected Array, got {:?}", result);
        }

        // INDIRECT with invalid reference: INDIRECT("ZZZZZ999999") → #REF!
        assert_eq!(
            fn_indirect(&[FormulaValue::String("ZZZZZ999999".into())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Ref)
        );

        // INDIRECT with sheet reference: INDIRECT("Sheet1!A1") → 99
        assert_eq!(
            fn_indirect(&[FormulaValue::String("Sheet1!A1".into())], &ctx).unwrap(),
            FormulaValue::Number(99.0)
        );

        // INDIRECT with R1C1 style (a1=false): our impl returns #REF!
        assert_eq!(
            fn_indirect(
                &[
                    FormulaValue::String("R1C1".into()),
                    FormulaValue::Boolean(false)
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Error(CellError::Ref)
        );

        // Error propagation: INDIRECT(#VALUE!) → #VALUE!
        assert_eq!(
            fn_indirect(&[FormulaValue::Error(CellError::Value)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // INDIRECT with empty string: INDIRECT("") → #REF!
        assert_eq!(
            fn_indirect(&[FormulaValue::String("".into())], &ctx).unwrap(),
            FormulaValue::Error(CellError::Ref)
        );

        // INDIRECT with absolute refs (dollar signs): INDIRECT("$B$2") → 1.333
        assert_eq!(
            fn_indirect(&[FormulaValue::String("$B$2".into())], &ctx).unwrap(),
            FormulaValue::Number(1.333)
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

    #[test]
    fn test_offset_docs() {
        let mut workbook = Workbook::new();
        {
            let sheet = workbook.worksheet_mut(0).unwrap();
            sheet
                .set_cell_value_at(5, 1, CellValue::Number(4.0))
                .unwrap();
            sheet
                .set_cell_value_at(5, 2, CellValue::Number(10.0))
                .unwrap();
            sheet
                .set_cell_value_at(6, 1, CellValue::Number(8.0))
                .unwrap();
            sheet
                .set_cell_value_at(6, 2, CellValue::Number(3.0))
                .unwrap();
            sheet
                .set_cell_value_at(7, 1, CellValue::Number(3.0))
                .unwrap();
            sheet
                .set_cell_value_at(7, 2, CellValue::Number(6.0))
                .unwrap();
        }

        assert_eq!(
            eval_with_workbook("=OFFSET(D3,3,-2,1,1)", &workbook).unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            eval_with_workbook("=SUM(OFFSET(D3:F5,3,-2,3,3))", &workbook).unwrap(),
            FormulaValue::Number(34.0)
        );
        assert_eq!(
            eval_with_workbook("=OFFSET(D3,-3,-3)", &workbook).unwrap(),
            FormulaValue::Error(CellError::Ref)
        );
        assert_eq!(
            eval_with_workbook("=OFFSET(D3,\"x\",0)", &workbook).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            eval_with_workbook("=OFFSET(D3,0,0,0,1)", &workbook).unwrap(),
            FormulaValue::Error(CellError::Ref)
        );
        assert_eq!(
            eval_with_workbook("=OFFSET(1/0,0,0)", &workbook).unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
    }

    // ===== DOCS-BASED TESTS =====

    #[test]
    fn test_index_docs() {
        // Docs Example 1: A2:B3 = {Apples,Lemons;Bananas,Pears}
        // =INDEX(A2:B3,2,2) -> Pears
        assert_eq!(
            eval("=INDEX({\"Apples\",\"Lemons\";\"Bananas\",\"Pears\"},2,2)").unwrap(),
            FormulaValue::String("Pears".into())
        );
        // =INDEX(A2:B3,2,1) -> Bananas
        assert_eq!(
            eval("=INDEX({\"Apples\",\"Lemons\";\"Bananas\",\"Pears\"},2,1)").unwrap(),
            FormulaValue::String("Bananas".into())
        );
        // Docs Example 2: =INDEX({1,2;3,4},0,2) -> column {2;4}
        assert_eq!(
            eval("=INDEX({1,2;3,4},0,2)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![FormulaValue::Number(2.0)],
                    vec![FormulaValue::Number(4.0)],
                ],
                source: None
            }
        );
        // Docs Reference form: Fruits/Price/Count table
        // =INDEX(..., 2, 3) -> 38 (Bananas count)
        assert_eq!(
            eval("=INDEX({\"Apples\",0.69,40;\"Bananas\",0.34,38;\"Lemons\",0.55,15;\"Oranges\",0.25,25;\"Pears\",0.59,40},2,3)").unwrap(),
            FormulaValue::Number(38.0)
        );
    }

    #[test]
    fn test_index_two_arg_vector_lookup() {
        // Two-arg INDEX on a single-row vector: position selects column
        // =INDEX({10,20,30,40}, 3) -> 30
        assert_eq!(
            eval("=INDEX({10,20,30,40},3)").unwrap(),
            FormulaValue::Number(30.0)
        );

        // Two-arg INDEX on a single-column vector: position selects row
        // =INDEX({10;20;30;40}, 3) -> 30
        assert_eq!(
            eval("=INDEX({10;20;30;40},3)").unwrap(),
            FormulaValue::Number(30.0)
        );

        // Two-arg INDEX with cell range reference (single row)
        {
            let mut workbook = Workbook::new();
            {
                let sheet = workbook.worksheet_mut(0).unwrap();
                // A1:L1 = months 1..12
                for i in 0u16..12 {
                    sheet.set_cell_value_at(0, i, CellValue::Number((i + 1) as f64)).unwrap();
                }
                // A2:L2 = values 10,20,...120
                for i in 0u16..12 {
                    sheet.set_cell_value_at(1, i, CellValue::Number(((i + 1) * 10) as f64)).unwrap();
                }
            }
            // INDEX($A$2:$L$2, 3) should return 30 (3rd column)
            assert_eq!(
                eval_with_workbook("=INDEX($A$2:$L$2,3)", &workbook).unwrap(),
                FormulaValue::Number(30.0)
            );
            // INDEX($A$2:$L$2, MATCH(3,$A$1:$L$1,0)) should also return 30
            assert_eq!(
                eval_with_workbook("=INDEX($A$2:$L$2,MATCH(3,$A$1:$L$1,0))", &workbook).unwrap(),
                FormulaValue::Number(30.0)
            );
        }

        // Two-arg INDEX with cell range reference (single column)
        {
            let mut workbook = Workbook::new();
            {
                let sheet = workbook.worksheet_mut(0).unwrap();
                // A1:A4 = {10;20;30;40}
                for i in 0u32..4 {
                    sheet.set_cell_value_at(i, 0, CellValue::Number(((i + 1) * 10) as f64)).unwrap();
                }
            }
            assert_eq!(
                eval_with_workbook("=INDEX($A$1:$A$4,3)", &workbook).unwrap(),
                FormulaValue::Number(30.0)
            );
        }

        // position=0 on single row: return entire row vector
        assert_eq!(
            eval("=INDEX({10,20,30},0)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(10.0), n(20.0), n(30.0)]],
                source: None,
            }
        );

        // position=0 on single column: return entire column vector
        assert_eq!(
            eval("=INDEX({10;20;30},0)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(10.0)], vec![n(20.0)], vec![n(30.0)]],
                source: None,
            }
        );

        // Out of bounds
        assert_eq!(
            eval("=INDEX({10,20,30},5)").unwrap(),
            FormulaValue::Error(CellError::Ref)
        );
        assert_eq!(
            eval("=INDEX({10;20;30},5)").unwrap(),
            FormulaValue::Error(CellError::Ref)
        );

        // 2D array with only row_num: returns entire row (col_num defaults to 0)
        assert_eq!(
            eval("=INDEX({1,2,3;4,5,6},2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(4.0), n(5.0), n(6.0)]],
                source: None,
            }
        );

        // 1x1 array: rows==1 wins, so position=1 -> (1,1) -> scalar
        assert_eq!(
            eval("=INDEX({42},1)").unwrap(),
            FormulaValue::Number(42.0)
        );

        // 3-arg form on single-row vector still works
        assert_eq!(
            eval("=INDEX({10,20,30,40},1,3)").unwrap(),
            FormulaValue::Number(30.0)
        );
    }

    #[test]
    fn test_match_docs() {
        // Docs data: {25;38;40;41} (ascending)
        // =MATCH(39,{25;38;40;41},1) -> 2 (largest value <= 39 is 38, at position 2)
        assert_eq!(
            eval("=MATCH(39,{25;38;40;41},1)").unwrap(),
            FormulaValue::Number(2.0)
        );
        // =MATCH(41,{25;38;40;41},0) -> 4 (exact match at position 4)
        assert_eq!(
            eval("=MATCH(41,{25;38;40;41},0)").unwrap(),
            FormulaValue::Number(4.0)
        );
        // =MATCH(40,{25;38;40;41},-1) -> #N/A (values not in descending order)
        assert_eq!(
            eval("=MATCH(40,{25;38;40;41},-1)").unwrap(),
            FormulaValue::Error(CellError::Na)
        );
        // Docs Remarks: MATCH("b",{"a","b","c"},0) -> 2
        assert_eq!(
            eval("=MATCH(\"b\",{\"a\",\"b\",\"c\"},0)").unwrap(),
            FormulaValue::Number(2.0)
        );
    }

    #[test]
    fn test_vlookup_docs() {
        // Example 1: =VLOOKUP("Fontana",B2:E7,2,FALSE) → "Olivier"
        assert_eq!(
            eval(concat!(
                r#"=VLOOKUP("Fontana",{"Davis","Sara";"Fontana","Olivier";"#,
                r#""Leal","Ana";"Sousa","Pedro";"Burke","James";"Baran","Kim"},2,FALSE)"#
            ))
            .unwrap(),
            FormulaValue::String("Olivier".into())
        );
        // Example 2: =VLOOKUP(102,A2:C7,2,FALSE) → "Fontana"
        assert_eq!(
            eval(concat!(
                r#"=VLOOKUP(102,{101,"Davis";102,"Fontana";103,"Leal";"#,
                r#"104,"Sousa";105,"Burke";106,"Baran"},2,FALSE)"#
            ))
            .unwrap(),
            FormulaValue::String("Fontana".into())
        );
        // Example 3: IF(VLOOKUP(103,...,2)="Souse","Located","Not found") → "Not found"
        assert_eq!(
            eval(concat!(
                r#"=IF(VLOOKUP(103,{101,"Davis";102,"Fontana";103,"Leal";"#,
                r#"104,"Sousa";105,"Burke";106,"Baran"},2,FALSE)="Souse","Located","Not found")"#
            ))
            .unwrap(),
            FormulaValue::String("Not found".into())
        );
        // Example 5: IF(ISNA(VLOOKUP(105,...))=TRUE,"Employee not found",VLOOKUP(105,...)) → "Burke"
        assert_eq!(
            eval(concat!(
                r#"=IF(ISNA(VLOOKUP(105,{101,"Davis";102,"Fontana";103,"Leal";"#,
                r#"104,"Sousa";105,"Burke";106,"Baran"},2,FALSE))=TRUE,"#,
                r#""Employee not found","#,
                r#"VLOOKUP(105,{101,"Davis";102,"Fontana";103,"Leal";"#,
                r#"104,"Sousa";105,"Burke";106,"Baran"},2,FALSE))"#
            ))
            .unwrap(),
            FormulaValue::String("Burke".into())
        );
    }

    #[test]
    fn test_hlookup_docs() {
        // Example 1: =HLOOKUP("Axles",table,2,TRUE) → 4
        assert_eq!(
            eval(r#"=HLOOKUP("Axles",{"Axles","Bearings","Bolts";4,4,9;5,7,10;6,8,11},2,TRUE)"#)
                .unwrap(),
            FormulaValue::Number(4.0)
        );
        // Example 2: =HLOOKUP("Bearings",table,3,FALSE) → 7
        assert_eq!(
            eval(
                r#"=HLOOKUP("Bearings",{"Axles","Bearings","Bolts";4,4,9;5,7,10;6,8,11},3,FALSE)"#
            )
            .unwrap(),
            FormulaValue::Number(7.0)
        );
        // Example 3: =HLOOKUP("B",table,3,TRUE) → 5 (approx match, Axles < B)
        assert_eq!(
            eval(r#"=HLOOKUP("B",{"Axles","Bearings","Bolts";4,4,9;5,7,10;6,8,11},3,TRUE)"#)
                .unwrap(),
            FormulaValue::Number(5.0)
        );
        // Example 4: =HLOOKUP("Bolts",table,4) → 11 (range_lookup default TRUE)
        assert_eq!(
            eval(r#"=HLOOKUP("Bolts",{"Axles","Bearings","Bolts";4,4,9;5,7,10;6,8,11},4)"#)
                .unwrap(),
            FormulaValue::Number(11.0)
        );
        // Example 5: =HLOOKUP(3,{1,2,3;"a","b","c";"d","e","f"},2,TRUE) → "c"
        assert_eq!(
            eval(r#"=HLOOKUP(3,{1,2,3;"a","b","c";"d","e","f"},2,TRUE)"#).unwrap(),
            FormulaValue::String("c".into())
        );
    }

    #[test]
    fn test_xlookup_docs() {
        // Example 1: exact match country → phone code
        assert_eq!(
            eval(concat!(
                r#"=XLOOKUP("Brazil","#,
                r#"{"China","India","United States","Indonesia","Brazil"},"#,
                r#"{86,91,1,62,55})"#
            ))
            .unwrap(),
            FormulaValue::Number(55.0)
        );
        // Example 3: if_not_found
        assert_eq!(
            eval(concat!(
                r#"=XLOOKUP("NotFound","#,
                r#"{"China","India","United States","Indonesia","Brazil"},"#,
                r#"{86,91,1,62,55},"#,
                r#""Employee not found")"#
            ))
            .unwrap(),
            FormulaValue::String("Employee not found".into())
        );
        // Example 4: match_mode=1 (next larger): 25000 → 40125 threshold → 0.12 rate
        assert_eq!(
            eval(concat!(
                "=XLOOKUP(25000,",
                "{9875,40125,85525,163300,207350,518400},",
                "{0.10,0.12,0.22,0.24,0.32,0.35},",
                "0,1,1)"
            ))
            .unwrap(),
            FormulaValue::Number(0.12)
        );
        // Default: #N/A when no match
        assert_eq!(
            eval(r#"=XLOOKUP("xyz",{"a","b","c"},{1,2,3})"#).unwrap(),
            FormulaValue::Error(CellError::Na)
        );
    }

    #[test]
    fn test_xmatch_docs() {
        // Example 1: match_mode=1, "Gra" → next largest "Grape" at position 2
        assert_eq!(
            eval(r#"=XMATCH("Gra",{"Apple","Grape","Lemon","Orange","Peach"},1)"#).unwrap(),
            FormulaValue::Number(2.0)
        );
        // Example 2: match_mode=1, bonus threshold 35000 among ascending sales
        assert_eq!(
            eval("=XMATCH(35000,{10000,20000,30000,40000,50000,60000,70000},1)").unwrap(),
            FormulaValue::Number(4.0)
        );
        // Example 4: exact match, 4 in {5,4,3,2,1} → position 2
        assert_eq!(
            eval("=XMATCH(4,{5,4,3,2,1})").unwrap(),
            FormulaValue::Number(2.0)
        );
        // Example 4: match_mode=1 (next largest), 4.5 in {5,4,3,2,1} → 5 at position 1
        assert_eq!(
            eval("=XMATCH(4.5,{5,4,3,2,1},1)").unwrap(),
            FormulaValue::Number(1.0)
        );
    }

    #[test]
    fn test_sequence_docs() {
        // Docs: =SEQUENCE(4,5) → 4×5 array, 1..20
        assert_eq!(
            eval("=SEQUENCE(4,5)").unwrap(),
            FormulaValue::Array {
                data: vec![
                    vec![n(1.0), n(2.0), n(3.0), n(4.0), n(5.0)],
                    vec![n(6.0), n(7.0), n(8.0), n(9.0), n(10.0)],
                    vec![n(11.0), n(12.0), n(13.0), n(14.0), n(15.0)],
                    vec![n(16.0), n(17.0), n(18.0), n(19.0), n(20.0)],
                ],
                source: None
            }
        );
        // Docs: =SEQUENCE(4) → 4×1 column {1;2;3;4}
        assert_eq!(
            eval("=SEQUENCE(4)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0)], vec![n(2.0)], vec![n(3.0)], vec![n(4.0)],],
                source: None
            }
        );
        // Docs: =SEQUENCE(1,5) → 1×5 row {1,2,3,4,5}
        assert_eq!(
            eval("=SEQUENCE(1,5)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(1.0), n(2.0), n(3.0), n(4.0), n(5.0)],],
                source: None
            }
        );
        // Docs: start and step: =SEQUENCE(4,1,2,2) → {2;4;6;8}
        assert_eq!(
            eval("=SEQUENCE(4,1,2,2)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(2.0)], vec![n(4.0)], vec![n(6.0)], vec![n(8.0)],],
                source: None
            }
        );
        // Docs: descending: =SEQUENCE(10,1,100,-5)
        assert_eq!(
            eval("=SEQUENCE(3,1,100,-5)").unwrap(),
            FormulaValue::Array {
                data: vec![vec![n(100.0)], vec![n(95.0)], vec![n(90.0)],],
                source: None
            }
        );
    }
}
