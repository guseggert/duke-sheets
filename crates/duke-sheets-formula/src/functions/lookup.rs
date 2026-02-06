//! Lookup functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

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
