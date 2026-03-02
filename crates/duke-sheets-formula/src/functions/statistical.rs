//! Statistical functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

use super::criteria::CriteriaMatcher;

/// COUNTA(value1, [value2], ...) - Counts the number of non-empty cells
/// Unlike COUNT which only counts numbers, COUNTA counts any non-blank cell
/// including numbers, text, errors, and boolean values.
/// Reference: LibreOffice ScInterpreter::ScCount2 (ifCOUNT2)
pub fn fn_counta(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut count = 0;

    for arg in args {
        match arg {
            // Count numbers
            FormulaValue::Number(_) => count += 1,
            // Count text (non-empty strings)
            FormulaValue::String(s) => {
                if !s.is_empty() {
                    count += 1;
                }
            }
            // Count booleans
            FormulaValue::Boolean(_) => count += 1,
            // Count errors (COUNTA counts error cells as non-empty)
            FormulaValue::Error(_) => count += 1,
            // Empty cells are not counted
            FormulaValue::Empty => {}
            // Handle arrays - recursively count non-empty cells
            FormulaValue::Array(arr) => {
                for row in arr {
                    for cell in row {
                        match cell {
                            FormulaValue::Number(_) => count += 1,
                            FormulaValue::String(s) => {
                                if !s.is_empty() {
                                    count += 1;
                                }
                            }
                            FormulaValue::Boolean(_) => count += 1,
                            FormulaValue::Error(_) => count += 1,
                            FormulaValue::Empty => {}
                            FormulaValue::Array(_) => {
                                // Nested arrays are rare, but count as 1 if present
                                count += 1;
                            }
                        }
                    }
                }
            }
        }
    }

    Ok(FormulaValue::Number(count as f64))
}

/// COUNTBLANK(range) - Counts empty cells in a range
/// Reference: LibreOffice has similar functionality
pub fn fn_countblank(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let mut count = 0;

    for arg in args {
        match arg {
            FormulaValue::Empty => count += 1,
            FormulaValue::String(s) if s.is_empty() => count += 1,
            FormulaValue::Array(arr) => {
                for row in arr {
                    for cell in row {
                        match cell {
                            FormulaValue::Empty => count += 1,
                            FormulaValue::String(s) if s.is_empty() => count += 1,
                            _ => {}
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(FormulaValue::Number(count as f64))
}

/// COUNTIF(range, criteria) - Counts cells that meet a criteria
/// Reference: LibreOffice ScInterpreter::ScCountIf
///
/// Criteria can be:
/// - A number: exact match (e.g., 5)
/// - A text string: case-insensitive match (e.g., "apple")
/// - A comparison expression: ">5", ">=10", "<100", "<=50", "<>0", "=5"
/// - Wildcards: "*" matches any characters, "?" matches single character
pub fn fn_countif(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Get the range (first argument)
    let range = match args.get(0) {
        Some(FormulaValue::Array(arr)) => arr,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => {
            // Single value - check if it matches criteria
            let criteria = match args.get(1) {
                Some(c) => c,
                None => return Ok(FormulaValue::Error(CellError::Value)),
            };
            let matcher = CriteriaMatcher::new(criteria);
            let count = if matcher.matches(v) { 1.0 } else { 0.0 };
            return Ok(FormulaValue::Number(count));
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Get the criteria (second argument)
    let criteria = match args.get(1) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let matcher = CriteriaMatcher::new(criteria);

    // Count cells that match criteria
    let mut count = 0;

    for row in range {
        for cell in row {
            if matcher.matches(cell) {
                count += 1;
            }
        }
    }

    Ok(FormulaValue::Number(count as f64))
}

/// AVERAGEIF(range, criteria, [average_range]) - Returns the average of cells that meet a criteria
/// Reference: LibreOffice ScInterpreter::ScAverageIf / IterateParametersIf
///
/// Criteria can be:
/// - A number: exact match (e.g., 5)
/// - A text string: case-insensitive match (e.g., "apple")
/// - A comparison expression: ">5", ">=10", "<100", "<=50", "<>0", "=5"
/// - Wildcards: "*" matches any characters, "?" matches single character
pub fn fn_averageif(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    // Get the range (first argument)
    let range = match args.get(0) {
        Some(FormulaValue::Array(arr)) => arr,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => {
            // Single value treated as 1x1 array
            return fn_averageif_single(v, args);
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // Get the criteria (second argument)
    let criteria = match args.get(1) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let matcher = CriteriaMatcher::new(criteria);

    // Get average_range (third argument) or use range
    let avg_range = match args.get(2) {
        Some(FormulaValue::Array(arr)) => arr,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(_) | None => range,
    };

    // Sum and count values where criteria matches
    let mut sum = 0.0;
    let mut count = 0;

    for (row_idx, row) in range.iter().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            if matcher.matches(cell) {
                // Get corresponding cell from avg_range
                if let Some(avg_row) = avg_range.get(row_idx) {
                    if let Some(avg_cell) = avg_row.get(col_idx) {
                        if let FormulaValue::Number(n) = avg_cell {
                            sum += n;
                            count += 1;
                        } else if let FormulaValue::Error(e) = avg_cell {
                            return Ok(FormulaValue::Error(*e));
                        }
                        // Non-numeric values are ignored
                    }
                }
            }
        }
    }

    if count == 0 {
        Ok(FormulaValue::Error(CellError::Div0))
    } else {
        Ok(FormulaValue::Number(sum / count as f64))
    }
}

/// Handle AVERAGEIF with single-value range
fn fn_averageif_single(value: &FormulaValue, args: &[FormulaValue]) -> FormulaResult<FormulaValue> {
    let criteria = match args.get(1) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let matcher = CriteriaMatcher::new(criteria);

    // Get avg value (third arg or use value)
    let avg_value = match args.get(2) {
        Some(v) => v,
        None => value,
    };

    if matcher.matches(value) {
        match avg_value {
            FormulaValue::Number(n) => Ok(FormulaValue::Number(*n)),
            FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
            _ => Ok(FormulaValue::Error(CellError::Div0)), // Non-numeric can't average
        }
    } else {
        Ok(FormulaValue::Error(CellError::Div0)) // No matches
    }
}

/// MEDIAN(number1, [number2], ...) - Returns the median of the given numbers
/// The median is the middle value when numbers are sorted.
/// If there's an even count, returns the average of the two middle values.
/// Reference: LibreOffice ScInterpreter::ScMedian / GetMedian
pub fn fn_median(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();

    // Collect all numbers from arguments
    for arg in args {
        if let Some(err) = collect_numbers(arg, &mut numbers) {
            return Ok(FormulaValue::Error(err));
        }
    }

    if numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    // Sort the numbers (LibreOffice uses nth_element for efficiency, but sort is correct)
    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    let len = numbers.len();
    let median = if len % 2 == 1 {
        // Odd count: middle value (upper median)
        numbers[len / 2]
    } else {
        // Even count: average of two middle values
        (numbers[len / 2 - 1] + numbers[len / 2]) / 2.0
    };

    Ok(FormulaValue::Number(median))
}

/// LARGE(array, k) - Returns the k-th largest value in a data set
/// Reference: LibreOffice ScInterpreter::CalculateSmallLarge(false)
pub fn fn_large(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Get the array
    let mut numbers = Vec::new();
    if let Some(arg) = args.get(0) {
        if let Some(err) = collect_numbers(arg, &mut numbers) {
            return Ok(FormulaValue::Error(err));
        }
    } else {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Get k (LibreOffice uses approxCeil for LARGE)
    let k = match args.get(1) {
        Some(FormulaValue::Number(n)) => n.ceil() as usize,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // k must be >= 1 and <= array size
    if k == 0 || k > numbers.len() || numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    // Sort ascending, then get (nSize - k) element
    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    Ok(FormulaValue::Number(numbers[numbers.len() - k]))
}

/// SMALL(array, k) - Returns the k-th smallest value in a data set
/// Reference: LibreOffice ScInterpreter::CalculateSmallLarge(true)
pub fn fn_small(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Get the array
    let mut numbers = Vec::new();
    if let Some(arg) = args.get(0) {
        if let Some(err) = collect_numbers(arg, &mut numbers) {
            return Ok(FormulaValue::Error(err));
        }
    } else {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Get k (LibreOffice uses approxFloor for SMALL)
    let k = match args.get(1) {
        Some(FormulaValue::Number(n)) => n.floor() as usize,
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    // k must be >= 1 and <= array size
    if k == 0 || k > numbers.len() || numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    // Sort ascending, then get (k-1) element
    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    Ok(FormulaValue::Number(numbers[k - 1]))
}

pub fn fn_stdev_s(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    for arg in args {
        if let Some(err) = collect_numbers(arg, &mut numbers) {
            return Ok(FormulaValue::Error(err));
        }
    }

    if numbers.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let variance = calculate_variance(&numbers, true);
    Ok(FormulaValue::Number(variance.sqrt()))
}

pub fn fn_stdev_p(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    for arg in args {
        if let Some(err) = collect_numbers(arg, &mut numbers) {
            return Ok(FormulaValue::Error(err));
        }
    }

    if numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let variance = calculate_variance(&numbers, false);
    Ok(FormulaValue::Number(variance.sqrt()))
}

pub fn fn_var_s(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    for arg in args {
        if let Some(err) = collect_numbers(arg, &mut numbers) {
            return Ok(FormulaValue::Error(err));
        }
    }

    if numbers.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(calculate_variance(&numbers, true)))
}

pub fn fn_var_p(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    for arg in args {
        if let Some(err) = collect_numbers(arg, &mut numbers) {
            return Ok(FormulaValue::Error(err));
        }
    }

    if numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(calculate_variance(&numbers, false)))
}

pub fn fn_mode_sngl(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    for arg in args {
        if let Some(err) = collect_numbers(arg, &mut numbers) {
            return Ok(FormulaValue::Error(err));
        }
    }

    if numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut best_value = numbers[0];
    let mut best_count = 0usize;

    for (idx, value) in numbers.iter().enumerate() {
        let mut count = 0usize;
        for candidate in &numbers {
            if candidate == value {
                count += 1;
            }
        }

        if count > best_count {
            best_count = count;
            best_value = numbers[idx];
        }
    }

    Ok(FormulaValue::Number(best_value))
}

pub fn fn_maxifs(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.len() < 3 || args.len() % 2 != 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let max_range = match &args[0] {
        FormulaValue::Array(arr) => arr,
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
        v => return fn_maxifs_single(v, &args[1..]),
    };

    let (rows, cols) = array_dims(max_range);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Number(0.0));
    }

    let num_pairs = (args.len() - 1) / 2;
    let mut criteria_ranges: Vec<&Vec<Vec<FormulaValue>>> = Vec::with_capacity(num_pairs);
    let mut matchers: Vec<CriteriaMatcher> = Vec::with_capacity(num_pairs);

    for i in 0..num_pairs {
        let range_idx = 1 + i * 2;
        let criteria_idx = range_idx + 1;

        let range = match &args[range_idx] {
            FormulaValue::Array(arr) => arr,
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        };

        let (r, c) = array_dims(range);
        if r != rows || c != cols {
            return Ok(FormulaValue::Error(CellError::Value));
        }

        criteria_ranges.push(range);

        let criteria = &args[criteria_idx];
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }
        matchers.push(CriteriaMatcher::new(criteria));
    }

    let mut current_max: Option<f64> = None;

    for row_idx in 0..rows {
        for col_idx in 0..cols {
            let mut all_match = true;
            for (range, matcher) in criteria_ranges.iter().zip(matchers.iter()) {
                if !matcher.matches(&range[row_idx][col_idx]) {
                    all_match = false;
                    break;
                }
            }

            if all_match {
                match &max_range[row_idx][col_idx] {
                    FormulaValue::Number(n) => {
                        current_max = Some(match current_max {
                            Some(m) => m.max(*n),
                            None => *n,
                        });
                    }
                    FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                    _ => {}
                }
            }
        }
    }

    Ok(FormulaValue::Number(current_max.unwrap_or(0.0)))
}

pub fn fn_minifs(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.len() < 3 || args.len() % 2 != 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let min_range = match &args[0] {
        FormulaValue::Array(arr) => arr,
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
        v => return fn_minifs_single(v, &args[1..]),
    };

    let (rows, cols) = array_dims(min_range);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Number(0.0));
    }

    let num_pairs = (args.len() - 1) / 2;
    let mut criteria_ranges: Vec<&Vec<Vec<FormulaValue>>> = Vec::with_capacity(num_pairs);
    let mut matchers: Vec<CriteriaMatcher> = Vec::with_capacity(num_pairs);

    for i in 0..num_pairs {
        let range_idx = 1 + i * 2;
        let criteria_idx = range_idx + 1;

        let range = match &args[range_idx] {
            FormulaValue::Array(arr) => arr,
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        };

        let (r, c) = array_dims(range);
        if r != rows || c != cols {
            return Ok(FormulaValue::Error(CellError::Value));
        }

        criteria_ranges.push(range);

        let criteria = &args[criteria_idx];
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }
        matchers.push(CriteriaMatcher::new(criteria));
    }

    let mut current_min: Option<f64> = None;

    for row_idx in 0..rows {
        for col_idx in 0..cols {
            let mut all_match = true;
            for (range, matcher) in criteria_ranges.iter().zip(matchers.iter()) {
                if !matcher.matches(&range[row_idx][col_idx]) {
                    all_match = false;
                    break;
                }
            }

            if all_match {
                match &min_range[row_idx][col_idx] {
                    FormulaValue::Number(n) => {
                        current_min = Some(match current_min {
                            Some(m) => m.min(*n),
                            None => *n,
                        });
                    }
                    FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                    _ => {}
                }
            }
        }
    }

    Ok(FormulaValue::Number(current_min.unwrap_or(0.0)))
}

pub fn fn_rank_eq(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let mut numbers = Vec::new();
    match args.get(1) {
        Some(v) => {
            if let Some(err) = collect_numbers(v, &mut numbers) {
                return Ok(FormulaValue::Error(err));
            }
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    }

    if numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let ascending = match args.get(2) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n != 0.0,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => false,
    };

    let count = if ascending {
        numbers.iter().filter(|&&n| n < number).count()
    } else {
        numbers.iter().filter(|&&n| n > number).count()
    };

    Ok(FormulaValue::Number((count + 1) as f64))
}

pub fn fn_rank_avg(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match args.first() {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let mut numbers = Vec::new();
    match args.get(1) {
        Some(v) => {
            if let Some(err) = collect_numbers(v, &mut numbers) {
                return Ok(FormulaValue::Error(err));
            }
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    }

    if numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let ascending = match args.get(2) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n != 0.0,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => false,
    };

    let greater_or_less = if ascending {
        numbers.iter().filter(|&&n| n < number).count()
    } else {
        numbers.iter().filter(|&&n| n > number).count()
    };
    let equal = numbers.iter().filter(|&&n| n == number).count();

    if equal == 0 {
        return Ok(FormulaValue::Number((greater_or_less + 1) as f64));
    }

    let start = (greater_or_less + 1) as f64;
    let end = (greater_or_less + equal) as f64;
    Ok(FormulaValue::Number((start + end) / 2.0))
}

pub fn fn_percentile_inc(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    match args.first() {
        Some(v) => {
            if let Some(err) = collect_numbers(v, &mut numbers) {
                return Ok(FormulaValue::Error(err));
            }
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    }

    if numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    let k = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if !(0.0..=1.0).contains(&k) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(calculate_percentile_inc(&numbers, k)))
}

pub fn fn_percentile_exc(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    match args.first() {
        Some(v) => {
            if let Some(err) = collect_numbers(v, &mut numbers) {
                return Ok(FormulaValue::Error(err));
            }
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    }

    if numbers.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    let k = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let n = numbers.len() as f64;
    let lower_bound = 1.0 / (n + 1.0);
    let upper_bound = n / (n + 1.0);
    if k <= lower_bound || k >= upper_bound {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(calculate_percentile_exc(&numbers, k)))
}

pub fn fn_quartile_inc(
    args: &[FormulaValue],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let quart = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n.floor() as i32,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if !(0..=4).contains(&quart) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let k = quart as f64 / 4.0;
    fn_percentile_inc(&[args[0].clone(), FormulaValue::Number(k)], ctx)
}

pub fn fn_quartile_exc(
    args: &[FormulaValue],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let quart = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n.floor() as i32,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if !(1..=3).contains(&quart) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let k = quart as f64 / 4.0;
    fn_percentile_exc(&[args[0].clone(), FormulaValue::Number(k)], ctx)
}

pub fn fn_percentrank_inc(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    match args.first() {
        Some(v) => {
            if let Some(err) = collect_numbers(v, &mut numbers) {
                return Ok(FormulaValue::Error(err));
            }
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    }

    if numbers.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    let x = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let significance = match args.get(2) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) if n > 0.0 => n.floor() as i32,
            Some(_) => return Ok(FormulaValue::Error(CellError::Num)),
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => 3,
    };

    let rank = match calculate_percentrank(&numbers, x, false) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    Ok(FormulaValue::Number(truncate_significance(
        rank,
        significance,
    )))
}

pub fn fn_percentrank_exc(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let mut numbers = Vec::new();
    match args.first() {
        Some(v) => {
            if let Some(err) = collect_numbers(v, &mut numbers) {
                return Ok(FormulaValue::Error(err));
            }
        }
        None => return Ok(FormulaValue::Error(CellError::Value)),
    }

    if numbers.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    let x = match args.get(1) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    let significance = match args.get(2) {
        Some(FormulaValue::Error(e)) => return Ok(FormulaValue::Error(*e)),
        Some(v) => match v.as_number() {
            Some(n) if n > 0.0 => n.floor() as i32,
            Some(_) => return Ok(FormulaValue::Error(CellError::Num)),
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
        None => 3,
    };

    let rank = match calculate_percentrank(&numbers, x, true) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    Ok(FormulaValue::Number(truncate_significance(
        rank,
        significance,
    )))
}

pub fn fn_stdev(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_stdev_s(args, ctx)
}

pub fn fn_stdevp(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_stdev_p(args, ctx)
}

pub fn fn_var(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_var_s(args, ctx)
}

pub fn fn_varp(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_var_p(args, ctx)
}

pub fn fn_mode(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_mode_sngl(args, ctx)
}

pub fn fn_percentile(
    args: &[FormulaValue],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    fn_percentile_inc(args, ctx)
}

pub fn fn_quartile(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_quartile_inc(args, ctx)
}

pub fn fn_rank(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_rank_eq(args, ctx)
}

pub fn fn_percentrank(
    args: &[FormulaValue],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    fn_percentrank_inc(args, ctx)
}

fn fn_maxifs_single(
    max_value: &FormulaValue,
    criteria_args: &[FormulaValue],
) -> FormulaResult<FormulaValue> {
    let num_pairs = criteria_args.len() / 2;

    for i in 0..num_pairs {
        let range_idx = i * 2;
        let criteria_idx = range_idx + 1;

        let range_value = &criteria_args[range_idx];
        let criteria = &criteria_args[criteria_idx];

        if let FormulaValue::Error(e) = range_value {
            return Ok(FormulaValue::Error(*e));
        }
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }

        let matcher = CriteriaMatcher::new(criteria);
        if !matcher.matches(range_value) {
            return Ok(FormulaValue::Number(0.0));
        }
    }

    match max_value {
        FormulaValue::Number(n) => Ok(FormulaValue::Number(*n)),
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::Number(0.0)),
    }
}

fn fn_minifs_single(
    min_value: &FormulaValue,
    criteria_args: &[FormulaValue],
) -> FormulaResult<FormulaValue> {
    let num_pairs = criteria_args.len() / 2;

    for i in 0..num_pairs {
        let range_idx = i * 2;
        let criteria_idx = range_idx + 1;

        let range_value = &criteria_args[range_idx];
        let criteria = &criteria_args[criteria_idx];

        if let FormulaValue::Error(e) = range_value {
            return Ok(FormulaValue::Error(*e));
        }
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }

        let matcher = CriteriaMatcher::new(criteria);
        if !matcher.matches(range_value) {
            return Ok(FormulaValue::Number(0.0));
        }
    }

    match min_value {
        FormulaValue::Number(n) => Ok(FormulaValue::Number(*n)),
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::Number(0.0)),
    }
}

fn calculate_variance(numbers: &[f64], sample: bool) -> f64 {
    let n = numbers.len() as f64;
    let mean = numbers.iter().sum::<f64>() / n;
    let sum_sq = numbers.iter().map(|x| (x - mean) * (x - mean)).sum::<f64>();
    if sample {
        sum_sq / (n - 1.0)
    } else {
        sum_sq / n
    }
}

fn calculate_percentile_inc(sorted_numbers: &[f64], k: f64) -> f64 {
    let n = sorted_numbers.len();
    if n == 1 {
        return sorted_numbers[0];
    }

    let position = k * (n as f64 - 1.0);
    let lower_idx = position.floor() as usize;
    let upper_idx = position.ceil() as usize;

    if lower_idx == upper_idx {
        return sorted_numbers[lower_idx];
    }

    let fraction = position - lower_idx as f64;
    sorted_numbers[lower_idx] + (sorted_numbers[upper_idx] - sorted_numbers[lower_idx]) * fraction
}

fn calculate_percentile_exc(sorted_numbers: &[f64], k: f64) -> f64 {
    let position = k * (sorted_numbers.len() as f64 + 1.0) - 1.0;
    let lower_idx = position.floor() as usize;
    let upper_idx = position.ceil() as usize;

    if lower_idx == upper_idx {
        return sorted_numbers[lower_idx];
    }

    let fraction = position - lower_idx as f64;
    sorted_numbers[lower_idx] + (sorted_numbers[upper_idx] - sorted_numbers[lower_idx]) * fraction
}

fn calculate_percentrank(
    sorted_numbers: &[f64],
    x: f64,
    exclusive: bool,
) -> Result<f64, CellError> {
    let n = sorted_numbers.len();

    if x < sorted_numbers[0] {
        return if exclusive {
            Err(CellError::Num)
        } else {
            Ok(0.0)
        };
    }
    if x > sorted_numbers[n - 1] {
        return if exclusive {
            Err(CellError::Num)
        } else {
            Ok(1.0)
        };
    }

    for i in 0..n {
        if sorted_numbers[i] == x {
            if exclusive {
                return Ok((i as f64 + 1.0) / (n as f64 + 1.0));
            }
            return Ok(i as f64 / (n as f64 - 1.0));
        }

        if i + 1 < n && sorted_numbers[i] < x && x < sorted_numbers[i + 1] {
            let interval = sorted_numbers[i + 1] - sorted_numbers[i];
            let fraction = if interval == 0.0 {
                0.0
            } else {
                (x - sorted_numbers[i]) / interval
            };

            if exclusive {
                return Ok((i as f64 + 1.0 + fraction) / (n as f64 + 1.0));
            }
            return Ok((i as f64 + fraction) / (n as f64 - 1.0));
        }
    }

    Err(CellError::Num)
}

fn truncate_significance(value: f64, significance: i32) -> f64 {
    let factor = 10_f64.powi(significance);
    (value * factor).floor() / factor
}

/// Helper function to collect numbers from a FormulaValue into a vector
/// Returns Some(CellError) if an error is encountered, None otherwise
fn collect_numbers(value: &FormulaValue, numbers: &mut Vec<f64>) -> Option<CellError> {
    match value {
        FormulaValue::Number(n) => numbers.push(*n),
        FormulaValue::Error(e) => return Some(*e),
        FormulaValue::Array(arr) => {
            for row in arr {
                for cell in row {
                    match cell {
                        FormulaValue::Number(n) => numbers.push(*n),
                        FormulaValue::Error(e) => return Some(*e),
                        // Skip non-numeric values (text, booleans, empty)
                        _ => {}
                    }
                }
            }
        }
        // Skip non-numeric values
        _ => {}
    }
    None
}

/// COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
/// Reference: LibreOffice ScInterpreter::ScCountIfs, Microsoft COUNTIFS function
///
/// Counts cells where ALL criteria are met.
/// All criteria ranges must have the same dimensions.
pub fn fn_countifs(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // Must have at least 2 arguments: criteria_range1, criteria1
    if args.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Must have even number of arguments (pairs)
    if args.len() % 2 != 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Get first criteria range to establish dimensions
    let first_range = match &args[0] {
        FormulaValue::Array(arr) => arr,
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
        v => {
            // Single value
            return fn_countifs_single(v, &args[1..]);
        }
    };

    let (rows, cols) = array_dims(first_range);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Number(0.0));
    }

    // Build criteria matchers from pairs
    let num_pairs = args.len() / 2;
    let mut criteria_ranges: Vec<&Vec<Vec<FormulaValue>>> = Vec::with_capacity(num_pairs);
    let mut matchers: Vec<CriteriaMatcher> = Vec::with_capacity(num_pairs);

    for i in 0..num_pairs {
        let range_idx = i * 2;
        let criteria_idx = range_idx + 1;

        // Get criteria range
        let range = match &args[range_idx] {
            FormulaValue::Array(arr) => arr,
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        };

        // Validate dimensions match
        let (r, c) = array_dims(range);
        if r != rows || c != cols {
            return Ok(FormulaValue::Error(CellError::Value));
        }

        criteria_ranges.push(range);

        // Get criteria and create matcher
        let criteria = &args[criteria_idx];
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }
        matchers.push(CriteriaMatcher::new(criteria));
    }

    // Count cells where ALL criteria match
    let mut count = 0.0;

    for row_idx in 0..rows {
        for col_idx in 0..cols {
            // Check if all criteria match for this cell
            let mut all_match = true;
            for (range, matcher) in criteria_ranges.iter().zip(matchers.iter()) {
                let cell = &range[row_idx][col_idx];
                if !matcher.matches(cell) {
                    all_match = false;
                    break;
                }
            }

            if all_match {
                count += 1.0;
            }
        }
    }

    Ok(FormulaValue::Number(count))
}

/// Helper for COUNTIFS with single-value ranges
fn fn_countifs_single(
    first_value: &FormulaValue,
    remaining_args: &[FormulaValue],
) -> FormulaResult<FormulaValue> {
    // First pair: first_value is the range, remaining_args[0] is criteria
    let first_criteria = &remaining_args[0];
    if let FormulaValue::Error(e) = first_criteria {
        return Ok(FormulaValue::Error(*e));
    }

    let first_matcher = CriteriaMatcher::new(first_criteria);
    if !first_matcher.matches(first_value) {
        return Ok(FormulaValue::Number(0.0));
    }

    // Check remaining pairs
    let num_remaining_pairs = (remaining_args.len() - 1) / 2;
    for i in 0..num_remaining_pairs {
        let range_idx = 1 + i * 2;
        let criteria_idx = range_idx + 1;

        let range_value = &remaining_args[range_idx];
        let criteria = &remaining_args[criteria_idx];

        if let FormulaValue::Error(e) = range_value {
            return Ok(FormulaValue::Error(*e));
        }
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }

        let matcher = CriteriaMatcher::new(criteria);
        if !matcher.matches(range_value) {
            return Ok(FormulaValue::Number(0.0));
        }
    }

    // All criteria matched
    Ok(FormulaValue::Number(1.0))
}

/// AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
/// Reference: LibreOffice ScInterpreter::ScAverageIfs, Microsoft AVERAGEIFS function
///
/// Averages cells in average_range where ALL criteria are met.
/// All ranges must have the same dimensions.
/// Returns #DIV/0! if no cells meet all criteria.
pub fn fn_averageifs(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    // Must have at least 3 arguments: average_range, criteria_range1, criteria1
    if args.len() < 3 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Must have odd number of arguments (average_range + pairs)
    if args.len() % 2 != 1 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    // Get average_range (first argument)
    let avg_range = match &args[0] {
        FormulaValue::Array(arr) => arr,
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
        v => {
            // Single value
            return fn_averageifs_single(v, &args[1..]);
        }
    };

    let (rows, cols) = array_dims(avg_range);
    if rows == 0 || cols == 0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }

    // Build criteria matchers from pairs
    let num_pairs = (args.len() - 1) / 2;
    let mut criteria_ranges: Vec<&Vec<Vec<FormulaValue>>> = Vec::with_capacity(num_pairs);
    let mut matchers: Vec<CriteriaMatcher> = Vec::with_capacity(num_pairs);

    for i in 0..num_pairs {
        let range_idx = 1 + i * 2;
        let criteria_idx = range_idx + 1;

        // Get criteria range
        let range = match &args[range_idx] {
            FormulaValue::Array(arr) => arr,
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
            _ => return Ok(FormulaValue::Error(CellError::Value)),
        };

        // Validate dimensions match
        let (r, c) = array_dims(range);
        if r != rows || c != cols {
            return Ok(FormulaValue::Error(CellError::Value));
        }

        criteria_ranges.push(range);

        // Get criteria and create matcher
        let criteria = &args[criteria_idx];
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }
        matchers.push(CriteriaMatcher::new(criteria));
    }

    // Sum and count values where ALL criteria match
    let mut sum = 0.0;
    let mut count = 0;

    for row_idx in 0..rows {
        for col_idx in 0..cols {
            // Check if all criteria match for this cell
            let mut all_match = true;
            for (range, matcher) in criteria_ranges.iter().zip(matchers.iter()) {
                let cell = &range[row_idx][col_idx];
                if !matcher.matches(cell) {
                    all_match = false;
                    break;
                }
            }

            if all_match {
                // Add the corresponding avg_range value if numeric
                let avg_cell = &avg_range[row_idx][col_idx];
                match avg_cell {
                    FormulaValue::Number(n) => {
                        sum += n;
                        count += 1;
                    }
                    FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                    _ => {} // Non-numeric ignored
                }
            }
        }
    }

    if count == 0 {
        Ok(FormulaValue::Error(CellError::Div0))
    } else {
        Ok(FormulaValue::Number(sum / count as f64))
    }
}

/// Helper for AVERAGEIFS with single-value ranges
fn fn_averageifs_single(
    avg_value: &FormulaValue,
    criteria_args: &[FormulaValue],
) -> FormulaResult<FormulaValue> {
    // Each pair: criteria_range, criteria
    let num_pairs = criteria_args.len() / 2;

    for i in 0..num_pairs {
        let range_idx = i * 2;
        let criteria_idx = range_idx + 1;

        let range_value = &criteria_args[range_idx];
        let criteria = &criteria_args[criteria_idx];

        if let FormulaValue::Error(e) = range_value {
            return Ok(FormulaValue::Error(*e));
        }
        if let FormulaValue::Error(e) = criteria {
            return Ok(FormulaValue::Error(*e));
        }

        let matcher = CriteriaMatcher::new(criteria);
        if !matcher.matches(range_value) {
            return Ok(FormulaValue::Error(CellError::Div0)); // No match
        }
    }

    // All criteria matched
    match avg_value {
        FormulaValue::Number(n) => Ok(FormulaValue::Number(*n)),
        FormulaValue::Error(e) => Ok(FormulaValue::Error(*e)),
        _ => Ok(FormulaValue::Error(CellError::Div0)),
    }
}

/// Helper to get array dimensions
fn array_dims(arr: &[Vec<FormulaValue>]) -> (usize, usize) {
    let rows = arr.len();
    let cols = arr.first().map(|r| r.len()).unwrap_or(0);
    (rows, cols)
}

const MAX_NEWTON_ITERS: usize = 100;
const NEWTON_TOL: f64 = 1e-12;
const BETA_CF_MAX_ITERS: usize = 200;
const BETA_CF_TOL: f64 = 1e-15;
const GAMMA_SERIES_MAX_TERMS: usize = 200;
const SQRT_2PI: f64 = 2.506_628_274_631_000_2;

fn scalar_number(value: &FormulaValue) -> Result<f64, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => value
            .as_number()
            .ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn required_number(args: &[FormulaValue], idx: usize) -> Result<f64, FormulaValue> {
    match args.get(idx) {
        Some(v) => scalar_number(v),
        None => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn optional_number(args: &[FormulaValue], idx: usize, default: f64) -> Result<f64, FormulaValue> {
    match args.get(idx).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(v) => scalar_number(v),
        None => Ok(default),
    }
}

fn scalar_bool(value: &FormulaValue) -> Result<bool, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => value.as_bool().ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn required_bool(args: &[FormulaValue], idx: usize) -> Result<bool, FormulaValue> {
    match args.get(idx) {
        Some(v) => scalar_bool(v),
        None => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn is_integer(x: f64) -> bool {
    (x - x.round()).abs() <= 1e-10
}

fn to_nonneg_usize(x: f64) -> Option<usize> {
    if x.is_finite() && x >= 0.0 && is_integer(x) {
        Some(x.round() as usize)
    } else {
        None
    }
}

fn to_pos_usize(x: f64) -> Option<usize> {
    if x.is_finite() && x > 0.0 && is_integer(x) {
        Some(x.round() as usize)
    } else {
        None
    }
}

fn ln_gamma(z: f64) -> f64 {
    let coeffs = [
        0.999_999_999_999_809_9,
        676.520_368_121_885_1,
        -1_259.139_216_722_402_8,
        771.323_428_777_653_1,
        -176.615_029_162_140_6,
        12.507_343_278_686_905,
        -0.138_571_095_265_720_12,
        9.984_369_578_019_572e-6,
        1.505_632_735_149_311_6e-7,
    ];

    if z <= 0.0 && is_integer(z) {
        return f64::NAN;
    }
    if z < 0.5 {
        let pi = std::f64::consts::PI;
        return (pi / (pi * z).sin()).ln() - ln_gamma(1.0 - z);
    }

    let z1 = z - 1.0;
    let mut x = coeffs[0];
    for (i, c) in coeffs.iter().enumerate().skip(1) {
        x += c / (z1 + i as f64);
    }

    let t = z1 + 7.5;
    0.5 * (2.0 * std::f64::consts::PI).ln() + (z1 + 0.5) * t.ln() - t + x.ln()
}

fn gamma_fn(x: f64) -> f64 {
    ln_gamma(x).exp()
}

fn beta_fn(a: f64, b: f64) -> f64 {
    (ln_gamma(a) + ln_gamma(b) - ln_gamma(a + b)).exp()
}

fn beta_continued_fraction(x: f64, a: f64, b: f64) -> f64 {
    let tiny = 1e-300;
    let qab = a + b;
    let qap = a + 1.0;
    let qam = a - 1.0;
    let mut c = 1.0;
    let mut d = 1.0 - qab * x / qap;
    if d.abs() < tiny {
        d = tiny;
    }
    d = 1.0 / d;
    let mut h = d;

    for m in 1..=BETA_CF_MAX_ITERS {
        let m2 = 2.0 * m as f64;

        let aa = m as f64 * (b - m as f64) * x / ((qam + m2) * (a + m2));
        d = 1.0 + aa * d;
        if d.abs() < tiny {
            d = tiny;
        }
        c = 1.0 + aa / c;
        if c.abs() < tiny {
            c = tiny;
        }
        d = 1.0 / d;
        h *= d * c;

        let aa = -(a + m as f64) * (qab + m as f64) * x / ((a + m2) * (qap + m2));
        d = 1.0 + aa * d;
        if d.abs() < tiny {
            d = tiny;
        }
        c = 1.0 + aa / c;
        if c.abs() < tiny {
            c = tiny;
        }
        d = 1.0 / d;
        let delta = d * c;
        h *= delta;
        if (delta - 1.0).abs() < BETA_CF_TOL {
            break;
        }
    }

    h
}

fn regularized_beta(x: f64, a: f64, b: f64) -> f64 {
    if a <= 0.0 || b <= 0.0 {
        return f64::NAN;
    }
    if x <= 0.0 {
        return 0.0;
    }
    if x >= 1.0 {
        return 1.0;
    }

    let bt = (ln_gamma(a + b) - ln_gamma(a) - ln_gamma(b) + a * x.ln() + b * (1.0 - x).ln()).exp();
    let threshold = (a + 1.0) / (a + b + 2.0);
    if x < threshold {
        bt * beta_continued_fraction(x, a, b) / a
    } else {
        1.0 - bt * beta_continued_fraction(1.0 - x, b, a) / b
    }
}

fn regularized_gamma_p(a: f64, x: f64) -> f64 {
    if a <= 0.0 || x < 0.0 {
        return f64::NAN;
    }
    if x == 0.0 {
        return 0.0;
    }

    let mut sum = 1.0 / a;
    let mut del = sum;
    let mut ap = a;
    for _ in 1..=GAMMA_SERIES_MAX_TERMS {
        ap += 1.0;
        del *= x / ap;
        sum += del;
        if del.abs() < sum.abs() * 1e-15 {
            break;
        }
    }

    sum * (-x + a * x.ln() - ln_gamma(a)).exp()
}

fn regularized_gamma_q(a: f64, x: f64) -> f64 {
    1.0 - regularized_gamma_p(a, x)
}

fn normal_pdf(x: f64) -> f64 {
    (-0.5 * x * x).exp() / SQRT_2PI
}

fn normal_cdf(x: f64) -> f64 {
    if x == 0.0 {
        return 0.5;
    }
    let p = regularized_gamma_p(0.5, x * x / 2.0);
    if x > 0.0 {
        0.5 * (1.0 + p)
    } else {
        0.5 * (1.0 - p)
    }
}

fn inverse_normal(p: f64) -> f64 {
    let a = [
        -3.969_683_028_665_376e1,
        2.209_460_984_245_205e2,
        -2.759_285_104_469_687e2,
        1.383_577_518_672_69e2,
        -3.066_479_806_614_716e1,
        2.506_628_277_459_239,
    ];
    let b = [
        -5.447_609_879_822_406e1,
        1.615_858_368_580_409e2,
        -1.556_989_798_598_866e2,
        6.680_131_188_771_972e1,
        -1.328_068_155_288_572e1,
    ];
    let c = [
        -7.784_894_002_430_293e-3,
        -3.223_964_580_411_365e-1,
        -2.400_758_277_161_838,
        -2.549_732_539_343_734,
        4.374_664_141_464_968,
        2.938_163_982_698_783,
    ];
    let d = [
        7.784_695_709_041_462e-3,
        3.224_671_290_700_398e-1,
        2.445_134_137_142_996,
        3.754_408_661_907_416,
    ];

    let plow = 0.024_25;
    let phigh = 1.0 - plow;

    if p < plow {
        let q = (-2.0 * p.ln()).sqrt();
        return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5])
            / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
    }
    if p > phigh {
        let q = (-2.0 * (1.0 - p).ln()).sqrt();
        return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5])
            / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
    }

    let q = p - 0.5;
    let r = q * q;
    (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q
        / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1.0)
}

fn t_pdf(x: f64, df: f64) -> f64 {
    let num = gamma_fn((df + 1.0) / 2.0);
    let den = (df * std::f64::consts::PI).sqrt() * gamma_fn(df / 2.0);
    num / den * (1.0 + x * x / df).powf(-(df + 1.0) / 2.0)
}

fn t_cdf(x: f64, df: f64) -> f64 {
    let z = df / (df + x * x);
    let ib = regularized_beta(z, df / 2.0, 0.5);
    if x >= 0.0 {
        1.0 - 0.5 * ib
    } else {
        0.5 * ib
    }
}

fn chi2_pdf(x: f64, df: f64) -> f64 {
    if x <= 0.0 {
        if (df - 2.0).abs() < 1e-12 {
            return 0.5;
        }
        return 0.0;
    }
    let k = df / 2.0;
    x.powf(k - 1.0) * (-x / 2.0).exp() / (2.0_f64.powf(k) * gamma_fn(k))
}

fn chi2_cdf(x: f64, df: f64) -> f64 {
    regularized_gamma_p(df / 2.0, x / 2.0)
}

fn f_pdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    let a = d1 / 2.0;
    let b = d2 / 2.0;
    let num = (d1 / d2).powf(a) * x.powf(a - 1.0);
    let den = beta_fn(a, b) * (1.0 + d1 * x / d2).powf(a + b);
    num / den
}

fn f_cdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    let z = d1 * x / (d1 * x + d2);
    regularized_beta(z, d1 / 2.0, d2 / 2.0)
}

fn ln_binomial(n: usize, k: usize) -> f64 {
    ln_gamma((n + 1) as f64) - ln_gamma((k + 1) as f64) - ln_gamma((n - k + 1) as f64)
}

fn binom_pmf(k: usize, n: usize, p: f64) -> f64 {
    if p == 0.0 {
        return if k == 0 { 1.0 } else { 0.0 };
    }
    if p == 1.0 {
        return if k == n { 1.0 } else { 0.0 };
    }
    (ln_binomial(n, k) + (k as f64) * p.ln() + ((n - k) as f64) * (1.0 - p).ln()).exp()
}

fn hypergeom_pmf(
    sample_s: usize,
    number_sample: usize,
    population_s: usize,
    number_pop: usize,
) -> f64 {
    if sample_s > number_sample || sample_s > population_s {
        return 0.0;
    }
    if number_sample - sample_s > number_pop - population_s {
        return 0.0;
    }

    let ln_num = ln_binomial(population_s, sample_s)
        + ln_binomial(number_pop - population_s, number_sample - sample_s);
    let ln_den = ln_binomial(number_pop, number_sample);
    (ln_num - ln_den).exp()
}

fn poisson_pmf(x: usize, mean: f64) -> f64 {
    (-mean + (x as f64) * mean.ln() - ln_gamma((x + 1) as f64)).exp()
}

fn dist_collect_numbers(value: &FormulaValue, out: &mut Vec<f64>) -> Result<(), FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(rows) => {
            for row in rows {
                for cell in row {
                    dist_collect_numbers(cell, out)?;
                }
            }
            Ok(())
        }
        _ => {
            let n = value
                .as_number()
                .ok_or(FormulaValue::Error(CellError::Value))?;
            out.push(n);
            Ok(())
        }
    }
}

fn matrix_numbers(value: &FormulaValue) -> Result<Vec<Vec<f64>>, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(rows) => {
            if rows.is_empty() {
                return Err(FormulaValue::Error(CellError::Value));
            }
            let cols = rows[0].len();
            if cols == 0 {
                return Err(FormulaValue::Error(CellError::Value));
            }
            let mut out = Vec::with_capacity(rows.len());
            for row in rows {
                if row.len() != cols {
                    return Err(FormulaValue::Error(CellError::Value));
                }
                let mut out_row = Vec::with_capacity(cols);
                for cell in row {
                    out_row.push(
                        cell.as_number()
                            .ok_or(FormulaValue::Error(CellError::Value))?,
                    );
                }
                out.push(out_row);
            }
            Ok(out)
        }
        _ => Ok(vec![vec![scalar_number(value)?]]),
    }
}

fn inverse_newton<F, G>(target: f64, mut x: f64, cdf: F, pdf: G, lo: f64, hi: f64) -> Option<f64>
where
    F: Fn(f64) -> f64,
    G: Fn(f64) -> f64,
{
    let mut left = lo;
    let mut right = hi;
    for _ in 0..MAX_NEWTON_ITERS {
        let fx = cdf(x) - target;
        if fx.abs() < NEWTON_TOL {
            return Some(x.max(lo).min(hi));
        }

        if fx > 0.0 {
            right = right.min(x);
        } else {
            left = left.max(x);
        }

        let d = pdf(x);
        let mut next = if d.is_finite() && d.abs() > 1e-14 {
            x - fx / d
        } else {
            (left + right) / 2.0
        };

        if !next.is_finite() || next <= left || next >= right {
            next = (left + right) / 2.0;
        }

        if (next - x).abs() < NEWTON_TOL {
            return Some(next.max(lo).min(hi));
        }
        x = next;
    }
    None
}

fn mean(values: &[f64]) -> f64 {
    values.iter().sum::<f64>() / values.len() as f64
}

fn variance(values: &[f64], sample: bool) -> f64 {
    let m = mean(values);
    let denom = if sample {
        (values.len() - 1) as f64
    } else {
        values.len() as f64
    };
    values.iter().map(|v| (v - m) * (v - m)).sum::<f64>() / denom
}

pub fn fn_norm_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let mean_v = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let stdev = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if stdev <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let z = (x - mean_v) / stdev;
    let result = if cumulative {
        normal_cdf(z)
    } else {
        normal_pdf(z) / stdev
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_norm_s_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let z = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let result = if cumulative {
        normal_cdf(z)
    } else {
        normal_pdf(z)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_norm_inv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let mean_v = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let stdev = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if !(0.0..1.0).contains(&p) || stdev <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(mean_v + stdev * inverse_normal(p)))
}

pub fn fn_norm_s_inv(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if !(0.0..1.0).contains(&p) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(inverse_normal(p)))
}

pub fn fn_phi(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(normal_pdf(x)))
}

pub fn fn_binom_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number_s = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let trials = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let p = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let Some(k) = to_nonneg_usize(number_s) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    let Some(n) = to_nonneg_usize(trials) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    if p < 0.0 || p > 1.0 || k > n {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let result = if cumulative {
        (0..=k).map(|i| binom_pmf(i, n, p)).sum::<f64>()
    } else {
        binom_pmf(k, n, p)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_binom_dist_range(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let trials = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let p = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let number_s = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let number_s2 = match optional_number(args, 3, number_s) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let Some(n) = to_nonneg_usize(trials) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    let Some(k1) = to_nonneg_usize(number_s) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    let Some(k2) = to_nonneg_usize(number_s2) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    if p < 0.0 || p > 1.0 || k1 > k2 || k2 > n {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let prob = (k1..=k2).map(|k| binom_pmf(k, n, p)).sum::<f64>();
    Ok(FormulaValue::Number(prob))
}

pub fn fn_binom_inv(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let trials = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let p = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let alpha = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let Some(n) = to_nonneg_usize(trials) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    if p < 0.0 || p > 1.0 || !(0.0..=1.0).contains(&alpha) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut cdf = 0.0;
    for k in 0..=n {
        cdf += binom_pmf(k, n, p);
        if cdf >= alpha {
            return Ok(FormulaValue::Number(k as f64));
        }
    }
    Ok(FormulaValue::Number(n as f64))
}

pub fn fn_chisq_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if x < 0.0 || to_pos_usize(df).is_none() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let result = if cumulative {
        chi2_cdf(x, df)
    } else {
        chi2_pdf(x, df)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_chisq_dist_rt(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if x < 0.0 || to_pos_usize(df).is_none() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(regularized_gamma_q(df / 2.0, x / 2.0)))
}

pub fn fn_chisq_inv(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if !(0.0..1.0).contains(&p) || to_pos_usize(df).is_none() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut hi = df + 10.0 * (2.0 * df).sqrt() + 10.0;
    while chi2_cdf(hi, df) < p {
        hi *= 2.0;
        if hi > 1e10 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
    }

    let guess = (df
        * (1.0 - 2.0 / (9.0 * df) + inverse_normal(p) * (2.0 / (9.0 * df)).sqrt()).powi(3))
    .max(1e-9);
    let root = inverse_newton(p, guess, |x| chi2_cdf(x, df), |x| chi2_pdf(x, df), 0.0, hi);

    match root {
        Some(v) => Ok(FormulaValue::Number(v)),
        None => Ok(FormulaValue::Error(CellError::Num)),
    }
}

pub fn fn_chisq_inv_rt(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if !(0.0..1.0).contains(&p) || to_pos_usize(df).is_none() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    fn_chisq_inv(
        &[FormulaValue::Number(1.0 - p), FormulaValue::Number(df)],
        _ctx,
    )
}

pub fn fn_chisq_test(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let actual = match args.get(0) {
        Some(v) => match matrix_numbers(v) {
            Ok(m) => m,
            Err(e) => return Ok(e),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let expected = match args.get(1) {
        Some(v) => match matrix_numbers(v) {
            Ok(m) => m,
            Err(e) => return Ok(e),
        },
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if actual.len() != expected.len() || actual[0].len() != expected[0].len() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let rows = actual.len();
    let cols = actual[0].len();
    let mut stat = 0.0;
    for r in 0..rows {
        for c in 0..cols {
            let e = expected[r][c];
            if e <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let diff = actual[r][c] - e;
            stat += diff * diff / e;
        }
    }

    let df = if rows > 1 && cols > 1 {
        ((rows - 1) * (cols - 1)) as f64
    } else {
        (rows * cols - 1) as f64
    };
    if df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(regularized_gamma_q(
        df / 2.0,
        stat / 2.0,
    )))
}

pub fn fn_t_dist(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let result = if cumulative {
        t_cdf(x, df)
    } else {
        t_pdf(x, df)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_t_dist_2t(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x < 0.0 || df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let p = 2.0 * (1.0 - t_cdf(x, df));
    Ok(FormulaValue::Number(p.clamp(0.0, 1.0)))
}

pub fn fn_t_dist_rt(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number((1.0 - t_cdf(x, df)).clamp(0.0, 1.0)))
}

pub fn fn_t_inv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if !(0.0..1.0).contains(&p) || df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let guess = inverse_normal(p) * ((df - 2.0) / df).sqrt().max(0.5);
    let root = inverse_newton(p, guess, |x| t_cdf(x, df), |x| t_pdf(x, df), -1e6, 1e6);
    match root {
        Some(v) => Ok(FormulaValue::Number(v)),
        None => Ok(FormulaValue::Error(CellError::Num)),
    }
}

pub fn fn_t_inv_2t(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if p <= 0.0 || p > 1.0 || df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    if (p - 1.0).abs() < 1e-12 {
        return Ok(FormulaValue::Number(0.0));
    }
    fn_t_inv(
        &[
            FormulaValue::Number(1.0 - p / 2.0),
            FormulaValue::Number(df),
        ],
        _ctx,
    )
}

pub fn fn_t_test(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut a = Vec::new();
    let mut b = Vec::new();
    let tails = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let kind = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let Some(v1) = args.get(0) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let Some(v2) = args.get(1) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    if let Err(e) = dist_collect_numbers(v1, &mut a) {
        return Ok(e);
    }
    if let Err(e) = dist_collect_numbers(v2, &mut b) {
        return Ok(e);
    }
    if a.len() < 2 || b.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let tails_i = tails.round() as i32;
    let kind_i = kind.round() as i32;
    if !is_integer(tails) || !is_integer(kind) || !(tails_i == 1 || tails_i == 2) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let (t_stat, df) = match kind_i {
        1 => {
            if a.len() != b.len() || a.len() < 2 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let diffs: Vec<f64> = a.iter().zip(b.iter()).map(|(x, y)| x - y).collect();
            let m = mean(&diffs);
            let s2 = variance(&diffs, true);
            if s2 <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let t = m / (s2.sqrt() / (diffs.len() as f64).sqrt());
            (t.abs(), (diffs.len() - 1) as f64)
        }
        2 => {
            let n1 = a.len() as f64;
            let n2 = b.len() as f64;
            let m1 = mean(&a);
            let m2 = mean(&b);
            let v1 = variance(&a, true);
            let v2 = variance(&b, true);
            let pooled = ((n1 - 1.0) * v1 + (n2 - 1.0) * v2) / (n1 + n2 - 2.0);
            if pooled <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let t = (m1 - m2) / (pooled * (1.0 / n1 + 1.0 / n2)).sqrt();
            (t.abs(), n1 + n2 - 2.0)
        }
        3 => {
            let n1 = a.len() as f64;
            let n2 = b.len() as f64;
            let m1 = mean(&a);
            let m2 = mean(&b);
            let v1 = variance(&a, true);
            let v2 = variance(&b, true);
            let se2 = v1 / n1 + v2 / n2;
            if se2 <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let t = (m1 - m2) / se2.sqrt();
            let num = se2 * se2;
            let den = (v1 * v1) / (n1 * n1 * (n1 - 1.0)) + (v2 * v2) / (n2 * n2 * (n2 - 1.0));
            if den <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            (t.abs(), num / den)
        }
        _ => return Ok(FormulaValue::Error(CellError::Num)),
    };

    let one_tail = (1.0 - t_cdf(t_stat, df)).clamp(0.0, 1.0);
    let p = if tails_i == 2 {
        (2.0 * one_tail).clamp(0.0, 1.0)
    } else {
        one_tail
    };
    Ok(FormulaValue::Number(p))
}

pub fn fn_f_dist(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let d1 = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let d2 = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if x < 0.0 || d1 <= 0.0 || d2 <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let v = if cumulative {
        f_cdf(x, d1, d2)
    } else {
        f_pdf(x, d1, d2)
    };
    Ok(FormulaValue::Number(v))
}

pub fn fn_f_dist_rt(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let d1 = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let d2 = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x < 0.0 || d1 <= 0.0 || d2 <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(
        (1.0 - f_cdf(x, d1, d2)).clamp(0.0, 1.0),
    ))
}

pub fn fn_f_inv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let d1 = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let d2 = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if p <= 0.0 || p >= 1.0 || d1 <= 0.0 || d2 <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut hi = 1.0;
    while f_cdf(hi, d1, d2) < p {
        hi *= 2.0;
        if hi > 1e10 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
    }
    let root = inverse_newton(
        p,
        d2 / d1,
        |x| f_cdf(x, d1, d2),
        |x| f_pdf(x, d1, d2),
        0.0,
        hi,
    );
    match root {
        Some(v) => Ok(FormulaValue::Number(v)),
        None => Ok(FormulaValue::Error(CellError::Num)),
    }
}

pub fn fn_f_inv_rt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let d1 = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let d2 = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if p <= 0.0 || p >= 1.0 || d1 <= 0.0 || d2 <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    fn_f_inv(
        &[
            FormulaValue::Number(1.0 - p),
            FormulaValue::Number(d1),
            FormulaValue::Number(d2),
        ],
        _ctx,
    )
}

pub fn fn_f_test(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut a = Vec::new();
    let mut b = Vec::new();
    let Some(v1) = args.get(0) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let Some(v2) = args.get(1) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    if let Err(e) = dist_collect_numbers(v1, &mut a) {
        return Ok(e);
    }
    if let Err(e) = dist_collect_numbers(v2, &mut b) {
        return Ok(e);
    }
    if a.len() < 2 || b.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let va = variance(&a, true);
    let vb = variance(&b, true);
    if va <= 0.0 || vb <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let (f, d1, d2) = if va >= vb {
        (va / vb, (a.len() - 1) as f64, (b.len() - 1) as f64)
    } else {
        (vb / va, (b.len() - 1) as f64, (a.len() - 1) as f64)
    };
    let right = (1.0 - f_cdf(f, d1, d2)).clamp(0.0, 1.0);
    Ok(FormulaValue::Number((2.0 * right).clamp(0.0, 1.0)))
}

pub fn fn_gamma(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x <= 0.0 || (x <= 0.0 && is_integer(x)) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(gamma_fn(x)))
}

pub fn fn_gamma_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let result = if cumulative {
        regularized_gamma_p(alpha, x / beta)
    } else {
        x.powf(alpha - 1.0) * (-x / beta).exp() / (beta.powf(alpha) * gamma_fn(alpha))
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_gamma_inv(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if !(0.0..1.0).contains(&p) || alpha <= 0.0 || beta <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut hi = alpha * beta + 10.0 * (alpha * beta * beta).sqrt() + 10.0;
    while regularized_gamma_p(alpha, hi / beta) < p {
        hi *= 2.0;
        if hi > 1e10 {
            return Ok(FormulaValue::Error(CellError::Num));
        }
    }

    let pdf = |x: f64| {
        if x <= 0.0 {
            0.0
        } else {
            x.powf(alpha - 1.0) * (-x / beta).exp() / (beta.powf(alpha) * gamma_fn(alpha))
        }
    };
    let cdf = |x: f64| regularized_gamma_p(alpha, x / beta);
    let root = inverse_newton(p, alpha * beta, cdf, pdf, 0.0, hi);
    match root {
        Some(v) => Ok(FormulaValue::Number(v)),
        None => Ok(FormulaValue::Error(CellError::Num)),
    }
}

pub fn fn_gammaln(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if x <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(ln_gamma(x)))
}

pub fn fn_gammaln_precise(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    fn_gammaln(args, _ctx)
}

pub fn fn_beta_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let a = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let b = match optional_number(args, 5, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if alpha <= 0.0 || beta <= 0.0 || b <= a || x < a || x > b {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let xn = (x - a) / (b - a);
    let result = if cumulative {
        regularized_beta(xn, alpha, beta)
    } else {
        (x - a).powf(alpha - 1.0) * (b - x).powf(beta - 1.0)
            / ((b - a).powf(alpha + beta - 1.0) * beta_fn(alpha, beta))
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_expon_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let lambda = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if x < 0.0 || lambda <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let result = if cumulative {
        1.0 - (-lambda * x).exp()
    } else {
        lambda * (-lambda * x).exp()
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_hypgeom_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let sample_s = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let number_sample = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let population_s = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let number_pop = match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 4) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let Some(k) = to_nonneg_usize(sample_s) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    let Some(n) = to_nonneg_usize(number_sample) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    let Some(k_pop) = to_nonneg_usize(population_s) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    let Some(n_pop) = to_nonneg_usize(number_pop) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };

    if n > n_pop || k_pop > n_pop {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let min_k = n.saturating_sub(n_pop - k_pop);
    let max_k = n.min(k_pop);
    if k < min_k || k > max_k {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let result = if cumulative {
        (min_k..=k)
            .map(|i| hypergeom_pmf(i, n, k_pop, n_pop))
            .sum::<f64>()
    } else {
        hypergeom_pmf(k, n, k_pop, n_pop)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_negbinom_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let number_f = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let number_s = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let p = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let Some(k) = to_nonneg_usize(number_f) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    let Some(r) = to_pos_usize(number_s) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    if p < 0.0 || p > 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let pmf = |failures: usize| {
        if p == 0.0 {
            return 0.0;
        }
        if p == 1.0 {
            return if failures == 0 { 1.0 } else { 0.0 };
        }
        let ln_c = ln_binomial(failures + r - 1, failures);
        (ln_c + (r as f64) * p.ln() + (failures as f64) * (1.0 - p).ln()).exp()
    };

    let result = if cumulative {
        (0..=k).map(pmf).sum::<f64>()
    } else {
        pmf(k)
    };

    Ok(FormulaValue::Number(result))
}

pub fn fn_poisson_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let mean_v = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let Some(k) = to_nonneg_usize(x) else {
        return Ok(FormulaValue::Error(CellError::Num));
    };
    if mean_v <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let result = if cumulative {
        (0..=k).map(|i| poisson_pmf(i, mean_v)).sum::<f64>()
    } else {
        poisson_pmf(k, mean_v)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_weibull_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let xb = (x / beta).powf(alpha);
    let result = if cumulative {
        1.0 - (-xb).exp()
    } else {
        (alpha / beta) * (x / beta).powf(alpha - 1.0) * (-xb).exp()
    };

    Ok(FormulaValue::Number(result))
}

fn extra_scalar_number(value: &FormulaValue) -> Result<f64, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => value
            .as_number()
            .ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn extra_required_number(args: &[FormulaValue], idx: usize) -> Result<f64, FormulaValue> {
    match args.get(idx) {
        Some(v) => extra_scalar_number(v),
        None => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn extra_optional_number(
    args: &[FormulaValue],
    idx: usize,
    default: f64,
) -> Result<f64, FormulaValue> {
    match args.get(idx).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(v) => extra_scalar_number(v),
        None => Ok(default),
    }
}

fn extra_collect_numbers(value: &FormulaValue, out: &mut Vec<f64>) {
    match value {
        FormulaValue::Number(n) => out.push(*n),
        FormulaValue::Array(rows) => {
            for row in rows {
                for cell in row {
                    extra_collect_numbers(cell, out);
                }
            }
        }
        _ => {}
    }
}

fn extra_collect_values_a(value: &FormulaValue, out: &mut Vec<f64>) {
    match value {
        FormulaValue::Number(n) => out.push(*n),
        FormulaValue::Boolean(true) => out.push(1.0),
        FormulaValue::Boolean(false) => out.push(0.0),
        FormulaValue::String(_) => out.push(0.0),
        FormulaValue::Array(rows) => {
            for row in rows {
                for cell in row {
                    extra_collect_values_a(cell, out);
                }
            }
        }
        _ => {}
    }
}

fn extra_flatten_values(value: &FormulaValue, out: &mut Vec<FormulaValue>) {
    match value {
        FormulaValue::Array(rows) => {
            for row in rows {
                for cell in row {
                    extra_flatten_values(cell, out);
                }
            }
        }
        _ => out.push(value.clone()),
    }
}

fn extra_collect_numeric_args(args: &[FormulaValue]) -> Vec<f64> {
    let mut values = Vec::new();
    for arg in args {
        extra_collect_numbers(arg, &mut values);
    }
    values
}

fn extra_collect_numeric_args_a(args: &[FormulaValue]) -> Vec<f64> {
    let mut values = Vec::new();
    for arg in args {
        extra_collect_values_a(arg, &mut values);
    }
    values
}

fn extra_paired_numbers(
    a: &FormulaValue,
    b: &FormulaValue,
) -> Result<(Vec<f64>, Vec<f64>), FormulaValue> {
    let mut left_raw = Vec::new();
    let mut right_raw = Vec::new();
    extra_flatten_values(a, &mut left_raw);
    extra_flatten_values(b, &mut right_raw);

    if left_raw.len() != right_raw.len() {
        return Err(FormulaValue::Error(CellError::Na));
    }

    let mut left = Vec::new();
    let mut right = Vec::new();

    for (lv, rv) in left_raw.iter().zip(right_raw.iter()) {
        if let FormulaValue::Error(e) = lv {
            return Err(FormulaValue::Error(*e));
        }
        if let FormulaValue::Error(e) = rv {
            return Err(FormulaValue::Error(*e));
        }

        match (lv, rv) {
            (FormulaValue::Number(x), FormulaValue::Number(y)) => {
                left.push(*x);
                right.push(*y);
            }
            _ => {}
        }
    }

    Ok((left, right))
}

fn extra_mean(values: &[f64]) -> f64 {
    values.iter().sum::<f64>() / values.len() as f64
}

fn extra_variance(values: &[f64], sample: bool) -> f64 {
    let m = extra_mean(values);
    let sum_sq = values.iter().map(|v| (v - m) * (v - m)).sum::<f64>();
    if sample {
        sum_sq / (values.len() as f64 - 1.0)
    } else {
        sum_sq / values.len() as f64
    }
}

fn extra_correlation(x: &[f64], y: &[f64]) -> Result<f64, FormulaValue> {
    let mx = extra_mean(x);
    let my = extra_mean(y);

    let mut num = 0.0;
    let mut sxx = 0.0;
    let mut syy = 0.0;

    for (xi, yi) in x.iter().zip(y.iter()) {
        let dx = xi - mx;
        let dy = yi - my;
        num += dx * dy;
        sxx += dx * dx;
        syy += dy * dy;
    }

    let den = (sxx * syy).sqrt();
    if den == 0.0 {
        return Err(FormulaValue::Error(CellError::Div0));
    }

    Ok(num / den)
}

fn extra_regression_slope_intercept(x: &[f64], y: &[f64]) -> Result<(f64, f64), FormulaValue> {
    let mx = extra_mean(x);
    let my = extra_mean(y);
    let mut cov = 0.0;
    let mut varx = 0.0;

    for (xi, yi) in x.iter().zip(y.iter()) {
        let dx = xi - mx;
        cov += dx * (yi - my);
        varx += dx * dx;
    }

    if varx == 0.0 {
        return Err(FormulaValue::Error(CellError::Div0));
    }

    let slope = cov / varx;
    let intercept = my - slope * mx;
    Ok((slope, intercept))
}

fn extra_erf_approx(x: f64) -> f64 {
    let sign = if x < 0.0 { -1.0 } else { 1.0 };
    let ax = x.abs();
    let p = 0.3275911;
    let a1 = 0.254829592;
    let a2 = -0.284496736;
    let a3 = 1.421413741;
    let a4 = -1.453152027;
    let a5 = 1.061405429;
    let t = 1.0 / (1.0 + p * ax);
    let y = 1.0 - (((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t) * (-ax * ax).exp();
    sign * y
}

fn extra_normal_cdf(x: f64) -> f64 {
    if x == 0.0 {
        0.5
    } else {
        0.5 * (1.0 + extra_erf_approx(x / 2.0_f64.sqrt()))
    }
}

fn extra_inverse_standard_normal(p: f64) -> Option<f64> {
    if !(0.0 < p && p < 1.0) {
        return None;
    }

    let c0 = 2.515517;
    let c1 = 0.802853;
    let c2 = 0.010328;
    let d1 = 1.432788;
    let d2 = 0.189269;
    let d3 = 0.001308;

    let q = if p < 0.5 { p } else { 1.0 - p };
    let t = (-2.0 * q.ln()).sqrt();
    let x = t - (c0 + c1 * t + c2 * t * t) / (1.0 + d1 * t + d2 * t * t + d3 * t * t * t);

    if p < 0.5 {
        Some(-x)
    } else {
        Some(x)
    }
}

fn extra_ln_gamma(z: f64) -> f64 {
    let coeffs: [f64; 9] = [
        0.9999999999998099,
        676.5203681218851,
        -1259.1392167224028,
        771.3234287776531,
        -176.6150291621406,
        12.507343278686905,
        -0.13857109526572012,
        0.000009984369578019572,
        0.00000015056327351493116,
    ];

    if z < 0.5 {
        return std::f64::consts::PI.ln()
            - (std::f64::consts::PI * z).sin().ln()
            - extra_ln_gamma(1.0 - z);
    }

    let z1 = z - 1.0;
    let mut x = coeffs[0];
    for (i, c) in coeffs.iter().enumerate().skip(1) {
        x += c / (z1 + i as f64);
    }
    let t = z1 + 7.5;
    0.5 * (2.0 * std::f64::consts::PI).ln() + (z1 + 0.5) * t.ln() - t + x.ln()
}

fn extra_beta_continued_fraction(a: f64, b: f64, x: f64) -> f64 {
    let max_iter = 200;
    let eps = 3.0e-14;
    let fpmin = 1.0e-300;

    let mut c = 1.0;
    let mut d = 1.0 - (a + b) * x / (a + 1.0);
    if d.abs() < fpmin {
        d = fpmin;
    }
    d = 1.0 / d;
    let mut h = d;

    for m in 1..=max_iter {
        let m2 = 2 * m;
        let aa1 = m as f64 * (b - m as f64) * x / ((a + m2 as f64 - 1.0) * (a + m2 as f64));
        d = 1.0 + aa1 * d;
        if d.abs() < fpmin {
            d = fpmin;
        }
        c = 1.0 + aa1 / c;
        if c.abs() < fpmin {
            c = fpmin;
        }
        d = 1.0 / d;
        h *= d * c;

        let aa2 =
            -(a + m as f64) * (a + b + m as f64) * x / ((a + m2 as f64) * (a + m2 as f64 + 1.0));
        d = 1.0 + aa2 * d;
        if d.abs() < fpmin {
            d = fpmin;
        }
        c = 1.0 + aa2 / c;
        if c.abs() < fpmin {
            c = fpmin;
        }
        d = 1.0 / d;
        let del = d * c;
        h *= del;

        if (del - 1.0).abs() < eps {
            break;
        }
    }

    h
}

fn extra_regularized_beta(x: f64, a: f64, b: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    if x >= 1.0 {
        return 1.0;
    }

    let bt = (extra_ln_gamma(a + b) - extra_ln_gamma(a) - extra_ln_gamma(b)
        + a * x.ln()
        + b * (1.0 - x).ln())
    .exp();

    if x < (a + 1.0) / (a + b + 2.0) {
        bt * extra_beta_continued_fraction(a, b, x) / a
    } else {
        1.0 - bt * extra_beta_continued_fraction(b, a, 1.0 - x) / b
    }
}

fn extra_student_t_pdf(x: f64, df: f64) -> f64 {
    let log_num = extra_ln_gamma((df + 1.0) / 2.0);
    let log_den = 0.5 * (df * std::f64::consts::PI).ln() + extra_ln_gamma(df / 2.0);
    (log_num - log_den).exp() * (1.0 + x * x / df).powf(-(df + 1.0) / 2.0)
}

fn extra_student_t_cdf(x: f64, df: f64) -> f64 {
    if x == 0.0 {
        return 0.5;
    }

    let xb = df / (df + x * x);
    let ib = extra_regularized_beta(xb, df / 2.0, 0.5);
    if x > 0.0 {
        1.0 - 0.5 * ib
    } else {
        0.5 * ib
    }
}

fn extra_inverse_student_t(p: f64, df: f64) -> Option<f64> {
    if !(0.0 < p && p < 1.0) || df <= 0.0 {
        return None;
    }

    let mut x = extra_inverse_standard_normal(p)?;
    for _ in 0..30 {
        let cdf = extra_student_t_cdf(x, df);
        let pdf = extra_student_t_pdf(x, df);
        if pdf == 0.0 {
            break;
        }
        let delta = (cdf - p) / pdf;
        x -= delta;
        if delta.abs() < 1.0e-12 {
            break;
        }
    }
    Some(x)
}

fn extra_non_negative_integer(value: f64) -> Option<usize> {
    if !value.is_finite() || value < 0.0 {
        return None;
    }
    let rounded = value.round();
    if (value - rounded).abs() > 1.0e-9 {
        return None;
    }
    Some(rounded as usize)
}

fn extra_permutation(n: usize, k: usize) -> f64 {
    let mut acc = 1.0;
    for v in (n - k + 1)..=n {
        acc *= v as f64;
    }
    acc
}

pub fn fn_avedev(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args(args);
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let m = extra_mean(&values);
    let avg = values.iter().map(|v| (v - m).abs()).sum::<f64>() / values.len() as f64;
    Ok(FormulaValue::Number(avg))
}

pub fn fn_averagea(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args_a(args);
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number(
        values.iter().sum::<f64>() / values.len() as f64,
    ))
}

pub fn fn_devsq(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args(args);
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let m = extra_mean(&values);
    Ok(FormulaValue::Number(
        values.iter().map(|v| (v - m) * (v - m)).sum(),
    ))
}

pub fn fn_geomean(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args(args);
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    if values.iter().any(|v| *v <= 0.0) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let log_sum = values.iter().map(|v| v.ln()).sum::<f64>();
    Ok(FormulaValue::Number((log_sum / values.len() as f64).exp()))
}

pub fn fn_harmean(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args(args);
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    if values.iter().any(|v| *v <= 0.0) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let denom = values.iter().map(|v| 1.0 / v).sum::<f64>();
    if denom == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    Ok(FormulaValue::Number(values.len() as f64 / denom))
}

pub fn fn_kurt(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args(args);
    let n = values.len();
    if n < 4 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let m = extra_mean(&values);
    let s = extra_variance(&values, true).sqrt();
    if s == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }

    let sum4 = values.iter().map(|x| ((x - m) / s).powi(4)).sum::<f64>();
    let n_f = n as f64;
    let term1 = n_f * (n_f + 1.0) / ((n_f - 1.0) * (n_f - 2.0) * (n_f - 3.0)) * sum4;
    let term2 = 3.0 * (n_f - 1.0).powi(2) / ((n_f - 2.0) * (n_f - 3.0));
    Ok(FormulaValue::Number(term1 - term2))
}

pub fn fn_skew(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args(args);
    let n = values.len();
    if n < 3 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let m = extra_mean(&values);
    let s = extra_variance(&values, true).sqrt();
    if s == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    let n_f = n as f64;
    let sum3 = values.iter().map(|x| ((x - m) / s).powi(3)).sum::<f64>();
    Ok(FormulaValue::Number(
        n_f / ((n_f - 1.0) * (n_f - 2.0)) * sum3,
    ))
}

pub fn fn_skew_p(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args(args);
    let n = values.len();
    if n < 3 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let m = extra_mean(&values);
    let sigma = extra_variance(&values, false).sqrt();
    if sigma == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    let skew = values
        .iter()
        .map(|x| ((x - m) / sigma).powi(3))
        .sum::<f64>()
        / n as f64;
    Ok(FormulaValue::Number(skew))
}

pub fn fn_trimmean(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let percent = match extra_required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if !(0.0..1.0).contains(&percent) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut values = Vec::new();
    if let Some(v) = args.first() {
        extra_collect_numbers(v, &mut values);
    }
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    values.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    let total_trim = ((values.len() as f64 * percent).floor() as usize) & !1;
    let trim_each = total_trim / 2;
    if trim_each * 2 >= values.len() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let kept = &values[trim_each..(values.len() - trim_each)];
    Ok(FormulaValue::Number(extra_mean(kept)))
}

pub fn fn_standardize(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match extra_required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let m = match extra_required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let s = match extra_required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if s <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number((x - m) / s))
}

pub fn fn_correl(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (x, y) = match (args.first(), args.get(1)) {
        (Some(a), Some(b)) => match extra_paired_numbers(a, b) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    if x.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Na));
    }
    match extra_correlation(&x, &y) {
        Ok(v) => Ok(FormulaValue::Number(v)),
        Err(e) => Ok(e),
    }
}

pub fn fn_covariance_p(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (x, y) = match (args.first(), args.get(1)) {
        (Some(a), Some(b)) => match extra_paired_numbers(a, b) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    if x.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Na));
    }
    let mx = extra_mean(&x);
    let my = extra_mean(&y);
    let sum = x
        .iter()
        .zip(y.iter())
        .map(|(xi, yi)| (xi - mx) * (yi - my))
        .sum::<f64>();
    Ok(FormulaValue::Number(sum / x.len() as f64))
}

pub fn fn_covariance_s(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (x, y) = match (args.first(), args.get(1)) {
        (Some(a), Some(b)) => match extra_paired_numbers(a, b) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    if x.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    let mx = extra_mean(&x);
    let my = extra_mean(&y);
    let sum = x
        .iter()
        .zip(y.iter())
        .map(|(xi, yi)| (xi - mx) * (yi - my))
        .sum::<f64>();
    Ok(FormulaValue::Number(sum / (x.len() as f64 - 1.0)))
}

pub fn fn_pearson(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_correl(args, ctx)
}

pub fn fn_rsq(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let corr = match fn_correl(args, ctx)? {
        FormulaValue::Number(v) => v,
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    Ok(FormulaValue::Number(corr * corr))
}

pub fn fn_slope(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (y, x) = match (args.first(), args.get(1)) {
        (Some(a), Some(b)) => match extra_paired_numbers(a, b) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    if x.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Na));
    }
    match extra_regression_slope_intercept(&x, &y) {
        Ok((slope, _)) => Ok(FormulaValue::Number(slope)),
        Err(e) => Ok(e),
    }
}

pub fn fn_intercept(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (y, x) = match (args.first(), args.get(1)) {
        (Some(a), Some(b)) => match extra_paired_numbers(a, b) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    if x.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Na));
    }
    match extra_regression_slope_intercept(&x, &y) {
        Ok((_, intercept)) => Ok(FormulaValue::Number(intercept)),
        Err(e) => Ok(e),
    }
}

pub fn fn_fisher(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match extra_required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if !(-1.0 < x && x < 1.0) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(0.5 * ((1.0 + x) / (1.0 - x)).ln()))
}

pub fn fn_fisherinv(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let y = match extra_required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(y.tanh()))
}

pub fn fn_forecast_linear(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x_value = match extra_required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (y, x) = match (args.get(1), args.get(2)) {
        (Some(a), Some(b)) => match extra_paired_numbers(a, b) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    if x.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Na));
    }
    match extra_regression_slope_intercept(&x, &y) {
        Ok((slope, intercept)) => Ok(FormulaValue::Number(intercept + slope * x_value)),
        Err(e) => Ok(e),
    }
}

pub fn fn_forecast(args: &[FormulaValue], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_forecast_linear(args, ctx)
}

pub fn fn_frequency(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let mut data = Vec::new();
    let mut bins = Vec::new();
    if let Some(v) = args.first() {
        extra_collect_numbers(v, &mut data);
    }
    if let Some(v) = args.get(1) {
        extra_collect_numbers(v, &mut bins);
    }

    bins.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let mut counts = vec![0.0; bins.len() + 1];

    for value in data {
        let mut placed = false;
        for (i, bin) in bins.iter().enumerate() {
            if value <= *bin {
                counts[i] += 1.0;
                placed = true;
                break;
            }
        }
        if !placed {
            counts[bins.len()] += 1.0;
        }
    }

    let out = counts
        .into_iter()
        .map(|c| vec![FormulaValue::Number(c)])
        .collect::<Vec<_>>();
    Ok(FormulaValue::Array(out))
}

pub fn fn_maxa(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args_a(args);
    if values.is_empty() {
        return Ok(FormulaValue::Number(0.0));
    }
    Ok(FormulaValue::Number(
        values
            .iter()
            .copied()
            .fold(f64::NEG_INFINITY, |a, b| a.max(b)),
    ))
}

pub fn fn_mina(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args_a(args);
    if values.is_empty() {
        return Ok(FormulaValue::Number(0.0));
    }
    Ok(FormulaValue::Number(
        values.iter().copied().fold(f64::INFINITY, |a, b| a.min(b)),
    ))
}

pub fn fn_steyx(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (y, x) = match (args.first(), args.get(1)) {
        (Some(a), Some(b)) => match extra_paired_numbers(a, b) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let n = x.len();
    if n < 3 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }

    let (slope, intercept) = match extra_regression_slope_intercept(&x, &y) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let sse = x
        .iter()
        .zip(y.iter())
        .map(|(xi, yi)| {
            let y_hat = intercept + slope * xi;
            (yi - y_hat).powi(2)
        })
        .sum::<f64>();

    Ok(FormulaValue::Number((sse / (n as f64 - 2.0)).sqrt()))
}

pub fn fn_prob(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let lower = match extra_required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let upper = match extra_optional_number(args, 3, lower) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let (xv, pv) = match (args.first(), args.get(1)) {
        (Some(xr), Some(pr)) => match extra_paired_numbers(xr, pr) {
            Ok(v) => v,
            Err(e) => return Ok(e),
        },
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };

    if xv.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    if pv.iter().any(|p| !(*p >= 0.0 && *p <= 1.0)) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let p_sum = pv.iter().sum::<f64>();
    if (p_sum - 1.0).abs() > 1e-7 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let lo = lower.min(upper);
    let hi = lower.max(upper);
    let result = xv
        .iter()
        .zip(pv.iter())
        .filter(|(x, _)| **x >= lo && **x <= hi)
        .map(|(_, p)| *p)
        .sum::<f64>();

    Ok(FormulaValue::Number(result))
}

pub fn fn_permut(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let n = match extra_required_number(args, 0).and_then(|v| {
        extra_non_negative_integer(v)
            .ok_or(FormulaValue::Error(CellError::Num))
            .map(|n| n as f64)
    }) {
        Ok(v) => v as usize,
        Err(e) => return Ok(e),
    };
    let k = match extra_required_number(args, 1).and_then(|v| {
        extra_non_negative_integer(v)
            .ok_or(FormulaValue::Error(CellError::Num))
            .map(|k| k as f64)
    }) {
        Ok(v) => v as usize,
        Err(e) => return Ok(e),
    };
    if k > n {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(extra_permutation(n, k)))
}

pub fn fn_permutationa(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let n = match extra_required_number(args, 0).and_then(|v| {
        extra_non_negative_integer(v)
            .ok_or(FormulaValue::Error(CellError::Num))
            .map(|n| n as f64)
    }) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let k = match extra_required_number(args, 1).and_then(|v| {
        extra_non_negative_integer(v)
            .ok_or(FormulaValue::Error(CellError::Num))
            .map(|k| k as f64)
    }) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number(n.powf(k)))
}

pub fn fn_confidence_norm(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let alpha = match extra_required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let stdev = match extra_required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let size = match extra_required_number(args, 2).and_then(|v| {
        extra_non_negative_integer(v)
            .ok_or(FormulaValue::Error(CellError::Num))
            .map(|s| s as f64)
    }) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if !(0.0 < alpha && alpha < 1.0) || stdev <= 0.0 || size < 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let p = 1.0 - alpha / 2.0;
    let z = match extra_inverse_standard_normal(p) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    Ok(FormulaValue::Number(z * stdev / size.sqrt()))
}

pub fn fn_confidence_t(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let alpha = match extra_required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let stdev = match extra_required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let size = match extra_required_number(args, 2).and_then(|v| {
        extra_non_negative_integer(v)
            .ok_or(FormulaValue::Error(CellError::Num))
            .map(|s| s as f64)
    }) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if !(0.0 < alpha && alpha < 1.0) || stdev <= 0.0 || size < 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let df = size - 1.0;
    if df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let p = 1.0 - alpha / 2.0;
    let t = match extra_inverse_student_t(p, df) {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Num)),
    };

    Ok(FormulaValue::Number(t * stdev / size.sqrt()))
}

pub fn fn_gauss(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let z = match extra_required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(extra_normal_cdf(z) - 0.5))
}

pub fn fn_mode_mult(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args(args);
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    let mut uniques = values.clone();
    uniques.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    uniques.dedup_by(|a, b| (*a - *b).abs() < 1e-12);

    let mut best = 0usize;
    let mut modes = Vec::new();
    for u in uniques {
        let c = values.iter().filter(|v| (**v - u).abs() < 1e-12).count();
        if c > best {
            best = c;
            modes.clear();
            modes.push(u);
        } else if c == best {
            modes.push(u);
        }
    }

    if best <= 1 {
        return Ok(FormulaValue::Error(CellError::Na));
    }

    modes.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    Ok(FormulaValue::Array(
        modes
            .into_iter()
            .map(|v| vec![FormulaValue::Number(v)])
            .collect(),
    ))
}

pub fn fn_stdeva(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args_a(args);
    if values.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(extra_variance(&values, true).sqrt()))
}

pub fn fn_stdevpa(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args_a(args);
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(extra_variance(&values, false).sqrt()))
}

pub fn fn_vara(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args_a(args);
    if values.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(extra_variance(&values, true)))
}

pub fn fn_varpa(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = extra_collect_numeric_args_a(args);
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(extra_variance(&values, false)))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    #[test]
    fn test_stdev_s() {
        let result = eval("=STDEV.S(2,4,4,4,5,5,7,9)").unwrap();
        if let FormulaValue::Number(n) = result {
            assert!((n - 2.1380899353).abs() < 1e-9);
        } else {
            panic!("Expected Number");
        }
        assert_eq!(
            eval("=STDEV.S(1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_stdev_p() {
        let result = eval("=STDEV.P(2,4,4,4,5,5,7,9)").unwrap();
        if let FormulaValue::Number(n) = result {
            assert!((n - 2.0).abs() < 1e-12);
        } else {
            panic!("Expected Number");
        }
        assert_eq!(eval("=STDEV.P(5)").unwrap(), FormulaValue::Number(0.0));
    }

    #[test]
    fn test_var_s() {
        let result = eval("=VAR.S(2,4,4,4,5,5,7,9)").unwrap();
        if let FormulaValue::Number(n) = result {
            assert!((n - 4.5714285714).abs() < 1e-9);
        } else {
            panic!("Expected Number");
        }
        assert_eq!(
            eval("=VAR.S(1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_var_p() {
        assert_eq!(
            eval("=VAR.P(2,4,4,4,5,5,7,9)").unwrap(),
            FormulaValue::Number(4.0)
        );
        assert_eq!(eval("=VAR.P(5)").unwrap(), FormulaValue::Number(0.0));
    }

    #[test]
    fn test_mode_sngl() {
        assert_eq!(
            eval("=MODE.SNGL(1,2,2,3,3,3)").unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            eval("=MODE.SNGL(2,1,1,2)").unwrap(),
            FormulaValue::Number(2.0)
        );
    }

    #[test]
    fn test_maxifs() {
        assert_eq!(
            eval("=MAXIFS({10,20,30,40},{1,2,1,2},2)").unwrap(),
            FormulaValue::Number(40.0)
        );
        assert_eq!(
            eval("=MAXIFS({10,20,30},{1,2,3},99)").unwrap(),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_minifs() {
        assert_eq!(
            eval("=MINIFS({10,20,30,40},{1,2,1,2},2)").unwrap(),
            FormulaValue::Number(20.0)
        );
        assert_eq!(
            eval("=MINIFS({10,20,30},{1,2,3},99)").unwrap(),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_rank_eq() {
        assert_eq!(
            eval("=RANK.EQ(7,{10,9,7,7,4})").unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            eval("=RANK.EQ(7,{4,7,7,9,10},1)").unwrap(),
            FormulaValue::Number(2.0)
        );
    }

    #[test]
    fn test_rank_avg() {
        assert_eq!(
            eval("=RANK.AVG(7,{10,9,7,7,4})").unwrap(),
            FormulaValue::Number(3.5)
        );
        assert_eq!(
            eval("=RANK.AVG(7,{4,7,7,9,10},1)").unwrap(),
            FormulaValue::Number(2.5)
        );
    }

    #[test]
    fn test_percentile_inc() {
        assert_eq!(
            eval("=PERCENTILE.INC({1,2,3,4},0.25)").unwrap(),
            FormulaValue::Number(1.75)
        );
        assert_eq!(
            eval("=PERCENTILE.INC({1,2,3,4},1.1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_percentile_exc() {
        assert_eq!(
            eval("=PERCENTILE.EXC({1,2,3,4},0.25)").unwrap(),
            FormulaValue::Number(1.25)
        );
        assert_eq!(
            eval("=PERCENTILE.EXC({1,2,3,4},0.2)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_quartile_inc() {
        assert_eq!(
            eval("=QUARTILE.INC({1,2,3,4},2)").unwrap(),
            FormulaValue::Number(2.5)
        );
        assert_eq!(
            eval("=QUARTILE.INC({1,2,3,4},5)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_quartile_exc() {
        assert_eq!(
            eval("=QUARTILE.EXC({1,2,3,4},2)").unwrap(),
            FormulaValue::Number(2.5)
        );
        assert_eq!(
            eval("=QUARTILE.EXC({1,2,3,4},0)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_percentrank_inc() {
        assert_eq!(
            eval("=PERCENTRANK.INC({1,2,3,4},2)").unwrap(),
            FormulaValue::Number(0.333)
        );
        assert_eq!(
            eval("=PERCENTRANK.INC({1,2,3,4},0)").unwrap(),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_percentrank_exc() {
        assert_eq!(
            eval("=PERCENTRANK.EXC({1,2,3,4},2)").unwrap(),
            FormulaValue::Number(0.4)
        );
        assert_eq!(
            eval("=PERCENTRANK.EXC({1,2,3,4},0)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_stdev_alias() {
        assert_eq!(
            eval("=STDEV(1,2,3)").unwrap(),
            eval("=STDEV.S(1,2,3)").unwrap()
        );
        assert_eq!(
            eval("=STDEV(1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_stdevp_alias() {
        assert_eq!(
            eval("=STDEVP(1,2,3)").unwrap(),
            eval("=STDEV.P(1,2,3)").unwrap()
        );
        assert_eq!(eval("=STDEVP(5)").unwrap(), FormulaValue::Number(0.0));
    }

    #[test]
    fn test_var_alias() {
        assert_eq!(eval("=VAR(1,2,3)").unwrap(), eval("=VAR.S(1,2,3)").unwrap());
        assert_eq!(
            eval("=VAR(1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_varp_alias() {
        assert_eq!(
            eval("=VARP(1,2,3)").unwrap(),
            eval("=VAR.P(1,2,3)").unwrap()
        );
        assert_eq!(eval("=VARP(5)").unwrap(), FormulaValue::Number(0.0));
    }

    #[test]
    fn test_mode_alias() {
        assert_eq!(
            eval("=MODE(1,2,2,3)").unwrap(),
            eval("=MODE.SNGL(1,2,2,3)").unwrap()
        );
        assert_eq!(eval("=MODE(4,4,5)").unwrap(), FormulaValue::Number(4.0));
    }

    #[test]
    fn test_percentile_alias() {
        assert_eq!(
            eval("=PERCENTILE({1,2,3,4},0.5)").unwrap(),
            eval("=PERCENTILE.INC({1,2,3,4},0.5)").unwrap()
        );
        assert_eq!(
            eval("=PERCENTILE({1,2,3,4},-1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_quartile_alias() {
        assert_eq!(
            eval("=QUARTILE({1,2,3,4},1)").unwrap(),
            eval("=QUARTILE.INC({1,2,3,4},1)").unwrap()
        );
        assert_eq!(
            eval("=QUARTILE({1,2,3,4},9)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_rank_alias() {
        assert_eq!(
            eval("=RANK(5,{10,5,1})").unwrap(),
            eval("=RANK.EQ(5,{10,5,1})").unwrap()
        );
        assert_eq!(
            eval("=RANK(5,{1,5,10},1)").unwrap(),
            eval("=RANK.EQ(5,{1,5,10},1)").unwrap()
        );
    }

    #[test]
    fn test_percentrank_alias() {
        assert_eq!(
            eval("=PERCENTRANK({1,2,3,4},2)").unwrap(),
            eval("=PERCENTRANK.INC({1,2,3,4},2)").unwrap()
        );
        assert_eq!(
            eval("=PERCENTRANK({1,2,3,4},0)").unwrap(),
            FormulaValue::Number(0.0)
        );
    }

    mod distribution_tests {
        use super::*;

        fn ctx() -> EvaluationContext<'static> {
            EvaluationContext::simple()
        }

        fn num(v: FormulaValue) -> f64 {
            match v {
                FormulaValue::Number(n) => n,
                other => panic!("expected number, got {other:?}"),
            }
        }

        fn assert_close(actual: f64, expected: f64, tol: f64) {
            assert!(
                (actual - expected).abs() <= tol,
                "actual={actual}, expected={expected}, tol={tol}"
            );
        }

        #[test]
        fn test_norm_dist() {
            let c = ctx();
            let v = fn_norm_dist(
                &[
                    FormulaValue::Number(42.0),
                    FormulaValue::Number(40.0),
                    FormulaValue::Number(1.5),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.109_340_049_783_995_75, 1e-12);

            let v = fn_norm_dist(
                &[
                    FormulaValue::Number(42.0),
                    FormulaValue::Number(40.0),
                    FormulaValue::Number(1.5),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.908_788_761_035_648_2, 1e-7);
        }

        #[test]
        fn test_norm_s_dist() {
            let c = ctx();
            let v = fn_norm_s_dist(
                &[FormulaValue::Number(0.0), FormulaValue::Boolean(true)],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.5, 1e-12);

            let v = fn_norm_s_dist(
                &[FormulaValue::Number(1.96), FormulaValue::Boolean(true)],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.975_002_173_891_776_1, 1e-7);
        }

        #[test]
        fn test_norm_inv() {
            let c = ctx();
            let v = fn_norm_inv(
                &[
                    FormulaValue::Number(0.975),
                    FormulaValue::Number(0.0),
                    FormulaValue::Number(1.0),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 1.959_963_986_120_195, 1e-8);

            let e = fn_norm_inv(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(0.0),
                    FormulaValue::Number(1.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_norm_s_inv() {
            let c = ctx();
            let v = fn_norm_s_inv(&[FormulaValue::Number(0.5)], &c).unwrap();
            assert_close(num(v), 0.0, 1e-12);

            let v = fn_norm_s_inv(&[FormulaValue::Number(0.975)], &c).unwrap();
            assert_close(num(v), 1.959_963_986_120_195, 1e-8);
        }

        #[test]
        fn test_phi() {
            let c = ctx();
            let v = fn_phi(&[FormulaValue::Number(0.0)], &c).unwrap();
            assert_close(num(v), 0.398_942_280_401_432_7, 1e-12);

            let v = fn_phi(&[FormulaValue::Number(1.0)], &c).unwrap();
            assert_close(num(v), 0.241_970_724_519_143_37, 1e-12);
        }

        #[test]
        fn test_binom_dist() {
            let c = ctx();
            let v = fn_binom_dist(
                &[
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(0.5),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.117_187_5, 1e-12);

            let v = fn_binom_dist(
                &[
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(0.5),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.171_875, 1e-12);
        }

        #[test]
        fn test_binom_dist_range() {
            let c = ctx();
            let v = fn_binom_dist_range(
                &[
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(0.5),
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(5.0),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.568_359_375, 1e-12);

            let v = fn_binom_dist_range(
                &[
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(0.5),
                    FormulaValue::Number(3.0),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.117_187_5, 1e-12);
        }

        #[test]
        fn test_binom_inv() {
            let c = ctx();
            let v = fn_binom_inv(
                &[
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(0.5),
                    FormulaValue::Number(0.5),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 5.0, 1e-12);

            let e = fn_binom_inv(
                &[
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(0.5),
                    FormulaValue::Number(1.1),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_chisq_dist() {
            let c = ctx();
            let v = fn_chisq_dist(
                &[
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.632_120_558_828_557_7, 1e-12);

            let v = fn_chisq_dist(
                &[
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.183_939_720_585_721_17, 1e-12);
        }

        #[test]
        fn test_chisq_dist_rt() {
            let c = ctx();
            let v = fn_chisq_dist_rt(&[FormulaValue::Number(2.0), FormulaValue::Number(2.0)], &c)
                .unwrap();
            let rt = num(v);
            assert_close(rt, 0.367_879_441_171_442_33, 1e-12);

            let sum = num(fn_chisq_dist(
                &[
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap())
                + rt;
            assert_close(sum, 1.0, 1e-12);
        }

        #[test]
        fn test_chisq_inv() {
            let c = ctx();
            let v =
                fn_chisq_inv(&[FormulaValue::Number(0.95), FormulaValue::Number(2.0)], &c).unwrap();
            let x = num(v);
            assert_close(x, 5.991_464_547_107_981, 1e-8);

            let back = fn_chisq_dist(
                &[
                    FormulaValue::Number(x),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(back), 0.95, 1e-8);
        }

        #[test]
        fn test_chisq_inv_rt() {
            let c = ctx();
            let v = fn_chisq_inv_rt(&[FormulaValue::Number(0.05), FormulaValue::Number(2.0)], &c)
                .unwrap();
            let x = num(v);
            assert_close(x, 5.991_464_547_107_981, 1e-8);

            let rt = fn_chisq_dist_rt(&[FormulaValue::Number(x), FormulaValue::Number(2.0)], &c)
                .unwrap();
            assert_close(num(rt), 0.05, 1e-8);
        }

        #[test]
        fn test_chisq_test() {
            let c = ctx();
            let actual = FormulaValue::Array(vec![vec![
                FormulaValue::Number(10.0),
                FormulaValue::Number(20.0),
                FormulaValue::Number(30.0),
            ]]);
            let expected = FormulaValue::Array(vec![vec![
                FormulaValue::Number(12.0),
                FormulaValue::Number(18.0),
                FormulaValue::Number(30.0),
            ]]);
            let v = fn_chisq_test(&[actual, expected], &c).unwrap();
            assert_close(num(v), 0.757_465_128_396_966_4, 1e-12);

            let e = fn_chisq_test(
                &[
                    FormulaValue::Array(vec![vec![
                        FormulaValue::Number(1.0),
                        FormulaValue::Number(2.0),
                    ]]),
                    FormulaValue::Array(vec![vec![FormulaValue::Number(1.0)]]),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_t_dist() {
            let c = ctx();
            let v = fn_t_dist(
                &[
                    FormulaValue::Number(0.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.5, 1e-12);

            let v = fn_t_dist(
                &[
                    FormulaValue::Number(0.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.389_108_383_966_031, 1e-12);
        }

        #[test]
        fn test_t_dist_2t() {
            let c = ctx();
            let v =
                fn_t_dist_2t(&[FormulaValue::Number(2.0), FormulaValue::Number(10.0)], &c).unwrap();
            assert_close(num(v), 0.073_388_034_770_740_38, 1e-10);

            let e = fn_t_dist_2t(
                &[FormulaValue::Number(-1.0), FormulaValue::Number(10.0)],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_t_dist_rt() {
            let c = ctx();
            let v =
                fn_t_dist_rt(&[FormulaValue::Number(1.5), FormulaValue::Number(10.0)], &c).unwrap();
            let rt = num(v);
            assert_close(rt, 0.082_253_663_222_719_53, 1e-10);

            let sum = rt
                + num(fn_t_dist(
                    &[
                        FormulaValue::Number(1.5),
                        FormulaValue::Number(10.0),
                        FormulaValue::Boolean(true),
                    ],
                    &c,
                )
                .unwrap());
            assert_close(sum, 1.0, 1e-10);
        }

        #[test]
        fn test_t_inv() {
            let c = ctx();
            let v = fn_t_inv(
                &[FormulaValue::Number(0.975), FormulaValue::Number(10.0)],
                &c,
            )
            .unwrap();
            let x = num(v);
            assert_close(x, 2.228_138_851_964_938_5, 1e-6);

            let back = fn_t_dist(
                &[
                    FormulaValue::Number(x),
                    FormulaValue::Number(10.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(back), 0.975, 1e-6);
        }

        #[test]
        fn test_t_inv_2t() {
            let c = ctx();
            let v = fn_t_inv_2t(
                &[FormulaValue::Number(0.05), FormulaValue::Number(10.0)],
                &c,
            )
            .unwrap();
            assert_close(num(v), 2.228_138_851_964_938_5, 1e-6);

            let v =
                fn_t_inv_2t(&[FormulaValue::Number(1.0), FormulaValue::Number(10.0)], &c).unwrap();
            assert_close(num(v), 0.0, 1e-12);
        }

        #[test]
        fn test_t_test() {
            let c = ctx();
            let a = FormulaValue::Array(vec![vec![
                FormulaValue::Number(1.0),
                FormulaValue::Number(2.0),
                FormulaValue::Number(3.0),
                FormulaValue::Number(4.0),
                FormulaValue::Number(5.0),
            ]]);
            let b = FormulaValue::Array(vec![vec![
                FormulaValue::Number(1.1),
                FormulaValue::Number(1.9),
                FormulaValue::Number(3.2),
                FormulaValue::Number(3.8),
                FormulaValue::Number(5.1),
            ]]);
            let v = fn_t_test(
                &[a, b, FormulaValue::Number(2.0), FormulaValue::Number(1.0)],
                &c,
            )
            .unwrap();
            let p = num(v);
            assert!(p > 0.7 && p < 1.0);

            let e = fn_t_test(
                &[
                    FormulaValue::Array(vec![vec![FormulaValue::Number(1.0)]]),
                    FormulaValue::Array(vec![vec![FormulaValue::Number(2.0)]]),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_f_dist() {
            let c = ctx();
            let cdf = fn_f_dist(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(cdf), 0.534_880_573_462_199_5, 1e-10);

            let pdf = fn_f_dist(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert!(num(pdf) > 0.0);
        }

        #[test]
        fn test_f_dist_rt() {
            let c = ctx();
            let rt = fn_f_dist_rt(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                ],
                &c,
            )
            .unwrap();
            let cdf = fn_f_dist(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(rt) + num(cdf), 1.0, 1e-10);

            let e = fn_f_dist_rt(
                &[
                    FormulaValue::Number(-1.0),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_f_inv() {
            let c = ctx();
            let v = fn_f_inv(
                &[
                    FormulaValue::Number(0.95),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                ],
                &c,
            )
            .unwrap();
            let x = num(v);
            let back = fn_f_dist(
                &[
                    FormulaValue::Number(x),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(back), 0.95, 1e-7);

            let e = fn_f_inv(
                &[
                    FormulaValue::Number(0.0),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_f_inv_rt() {
            let c = ctx();
            let v = fn_f_inv_rt(
                &[
                    FormulaValue::Number(0.05),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                ],
                &c,
            )
            .unwrap();
            let x = num(v);
            let rt = fn_f_dist_rt(
                &[
                    FormulaValue::Number(x),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(rt), 0.05, 1e-7);

            let e = fn_f_inv_rt(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_f_test() {
            let c = ctx();
            let a = FormulaValue::Array(vec![vec![
                FormulaValue::Number(8.0),
                FormulaValue::Number(9.0),
                FormulaValue::Number(10.0),
                FormulaValue::Number(11.0),
                FormulaValue::Number(12.0),
            ]]);
            let b = FormulaValue::Array(vec![vec![
                FormulaValue::Number(1.0),
                FormulaValue::Number(2.0),
                FormulaValue::Number(3.0),
                FormulaValue::Number(4.0),
                FormulaValue::Number(5.0),
            ]]);
            let v = fn_f_test(&[a, b], &c).unwrap();
            let p = num(v);
            assert!(p > 0.0 && p <= 1.0);

            let e = fn_f_test(
                &[
                    FormulaValue::Array(vec![vec![FormulaValue::Number(1.0)]]),
                    FormulaValue::Array(vec![vec![FormulaValue::Number(2.0)]]),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_gamma() {
            let c = ctx();
            let v = fn_gamma(&[FormulaValue::Number(5.0)], &c).unwrap();
            assert_close(num(v), 24.0, 1e-10);

            let e = fn_gamma(&[FormulaValue::Number(0.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_gamma_dist() {
            let c = ctx();
            let v = fn_gamma_dist(
                &[
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.442_174_599_628_925_4, 1e-12);

            let v = fn_gamma_dist(
                &[
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.167_347_620_111_322_37, 1e-12);
        }

        #[test]
        fn test_gamma_inv() {
            let c = ctx();
            let v = fn_gamma_inv(
                &[
                    FormulaValue::Number(0.75),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                ],
                &c,
            )
            .unwrap();
            let x = num(v);
            let back = fn_gamma_dist(
                &[
                    FormulaValue::Number(x),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(back), 0.75, 1e-7);

            let e = fn_gamma_inv(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(2.0),
                ],
                &c,
            )
            .unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_gammaln() {
            let c = ctx();
            let v = fn_gammaln(&[FormulaValue::Number(5.0)], &c).unwrap();
            assert_close(num(v), 3.178_053_830_347_945_8, 1e-12);

            let e = fn_gammaln(&[FormulaValue::Number(-1.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_gammaln_precise() {
            let c = ctx();
            let v = fn_gammaln_precise(&[FormulaValue::Number(5.0)], &c).unwrap();
            assert_close(num(v), 3.178_053_830_347_945_8, 1e-12);

            let e = fn_gammaln_precise(&[FormulaValue::Number(0.0)], &c).unwrap();
            assert_eq!(e, FormulaValue::Error(CellError::Num));
        }

        #[test]
        fn test_beta_dist() {
            let c = ctx();
            let v = fn_beta_dist(
                &[
                    FormulaValue::Number(0.5),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(3.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 0.6875, 1e-12);

            let v = fn_beta_dist(
                &[
                    FormulaValue::Number(0.5),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(3.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(v), 1.5, 1e-12);
        }

        #[test]
        fn test_expon_dist() {
            let c = ctx();
            let cdf = fn_expon_dist(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(cdf), 0.864_664_716_763_387_3, 1e-12);

            let pdf = fn_expon_dist(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(pdf), 0.270_670_566_473_225_4, 1e-12);
        }

        #[test]
        fn test_hypgeom_dist() {
            let c = ctx();
            let pmf = fn_hypgeom_dist(
                &[
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(12.0),
                    FormulaValue::Number(7.0),
                    FormulaValue::Number(20.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            let pmf_num = num(pmf);
            assert_close(pmf_num, 0.286_068_111_455_108_34, 1e-12);

            let cdf = fn_hypgeom_dist(
                &[
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(12.0),
                    FormulaValue::Number(7.0),
                    FormulaValue::Number(20.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert!(num(cdf) > pmf_num);
        }

        #[test]
        fn test_negbinom_dist() {
            let c = ctx();
            let pmf = fn_negbinom_dist(
                &[
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(0.5),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            let pmf_num = num(pmf);
            assert_close(pmf_num, 0.061_096_191_406_25, 1e-12);

            let cdf = fn_negbinom_dist(
                &[
                    FormulaValue::Number(5.0),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(0.5),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert!(num(cdf) > pmf_num);
        }

        #[test]
        fn test_poisson_dist() {
            let c = ctx();
            let pmf = fn_poisson_dist(
                &[
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(pmf), 0.180_447_044_315_483_56, 1e-12);

            let cdf = fn_poisson_dist(
                &[
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(cdf), 0.857_123_460_498_547, 1e-12);
        }

        #[test]
        fn test_weibull_dist() {
            let c = ctx();
            let cdf = fn_weibull_dist(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(3.0),
                    FormulaValue::Boolean(true),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(cdf), 0.105_160_683_185_630_21, 1e-12);

            let pdf = fn_weibull_dist(
                &[
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(2.0),
                    FormulaValue::Number(3.0),
                    FormulaValue::Boolean(false),
                ],
                &c,
            )
            .unwrap();
            assert_close(num(pdf), 0.198_853_181_514_304_4, 1e-12);
        }
    }

    mod extra_tests {
        use super::*;

        fn ctx() -> EvaluationContext<'static> {
            EvaluationContext::simple()
        }

        fn n(v: f64) -> FormulaValue {
            FormulaValue::Number(v)
        }

        fn arr(v: &[f64]) -> FormulaValue {
            FormulaValue::Array(vec![v.iter().copied().map(FormulaValue::Number).collect()])
        }

        fn as_number(v: FormulaValue) -> f64 {
            match v {
                FormulaValue::Number(n) => n,
                _ => panic!("expected number"),
            }
        }

        #[test]
        fn test_avedev() {
            let c = ctx();
            let v = fn_avedev(&[n(2.0), n(4.0), n(6.0)], &c).unwrap();
            assert!((as_number(v) - 1.3333333333).abs() < 1e-9);
            assert_eq!(
                fn_avedev(&[], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_averagea() {
            let c = ctx();
            let v = fn_averagea(
                &[
                    n(2.0),
                    FormulaValue::Boolean(true),
                    FormulaValue::String("x".into()),
                ],
                &c,
            )
            .unwrap();
            assert!((as_number(v) - 1.0).abs() < 1e-12);
            assert_eq!(
                fn_averagea(&[FormulaValue::Empty], &c).unwrap(),
                FormulaValue::Error(CellError::Div0)
            );
        }

        #[test]
        fn test_devsq() {
            let c = ctx();
            assert!(
                (as_number(fn_devsq(&[n(1.0), n(2.0), n(3.0)], &c).unwrap()) - 2.0).abs() < 1e-12
            );
            assert_eq!(
                fn_devsq(&[], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_geomean() {
            let c = ctx();
            assert!(
                (as_number(fn_geomean(&[n(1.0), n(3.0), n(9.0)], &c).unwrap()) - 3.0).abs() < 1e-12
            );
            assert_eq!(
                fn_geomean(&[n(1.0), n(0.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_harmean() {
            let c = ctx();
            assert!(
                (as_number(fn_harmean(&[n(1.0), n(2.0), n(4.0)], &c).unwrap()) - 1.7142857143)
                    .abs()
                    < 1e-9
            );
            assert_eq!(
                fn_harmean(&[n(-1.0), n(2.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_kurt() {
            let c = ctx();
            let data = [n(1.0), n(2.0), n(3.0), n(4.0), n(5.0)];
            let v = as_number(fn_kurt(&data, &c).unwrap());
            assert!(v.is_finite());
            assert_eq!(
                fn_kurt(&[n(1.0), n(2.0), n(3.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_skew() {
            let c = ctx();
            assert!(as_number(fn_skew(&[n(1.0), n(2.0), n(4.0), n(9.0)], &c).unwrap()).is_finite());
            assert_eq!(
                fn_skew(&[n(1.0), n(2.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_skew_p() {
            let c = ctx();
            assert!(
                as_number(fn_skew_p(&[n(1.0), n(2.0), n(4.0), n(9.0)], &c).unwrap()).is_finite()
            );
            assert_eq!(
                fn_skew_p(&[n(1.0), n(1.0), n(1.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Div0)
            );
        }

        #[test]
        fn test_trimmean() {
            let c = ctx();
            let v = fn_trimmean(&[arr(&[1.0, 2.0, 3.0, 100.0]), n(0.5)], &c).unwrap();
            assert!((as_number(v) - 2.5).abs() < 1e-12);
            assert_eq!(
                fn_trimmean(&[arr(&[1.0, 2.0]), n(1.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_standardize() {
            let c = ctx();
            assert!(
                (as_number(fn_standardize(&[n(10.0), n(7.0), n(2.0)], &c).unwrap()) - 1.5).abs()
                    < 1e-12
            );
            assert_eq!(
                fn_standardize(&[n(1.0), n(0.0), n(0.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_correl() {
            let c = ctx();
            let v = fn_correl(&[arr(&[1.0, 2.0, 3.0]), arr(&[2.0, 4.0, 6.0])], &c).unwrap();
            assert!((as_number(v) - 1.0).abs() < 1e-12);
            assert_eq!(
                fn_correl(&[arr(&[1.0]), arr(&[2.0])], &c).unwrap(),
                FormulaValue::Error(CellError::Na)
            );
        }

        #[test]
        fn test_covariance_p() {
            let c = ctx();
            assert!(
                (as_number(
                    fn_covariance_p(&[arr(&[1.0, 2.0, 3.0]), arr(&[1.0, 5.0, 7.0])], &c).unwrap()
                ) - 2.0)
                    .abs()
                    < 1e-12
            );
            assert_eq!(
                fn_covariance_p(&[arr(&[1.0]), arr(&[2.0])], &c).unwrap(),
                FormulaValue::Error(CellError::Na)
            );
        }

        #[test]
        fn test_covariance_s() {
            let c = ctx();
            assert!(
                (as_number(
                    fn_covariance_s(&[arr(&[1.0, 2.0, 3.0]), arr(&[1.0, 5.0, 7.0])], &c).unwrap()
                ) - 3.0)
                    .abs()
                    < 1e-12
            );
            assert_eq!(
                fn_covariance_s(&[arr(&[1.0]), arr(&[2.0])], &c).unwrap(),
                FormulaValue::Error(CellError::Div0)
            );
        }

        #[test]
        fn test_pearson_and_rsq() {
            let c = ctx();
            let p =
                as_number(fn_pearson(&[arr(&[1.0, 2.0, 3.0]), arr(&[2.0, 4.0, 6.0])], &c).unwrap());
            let r2 =
                as_number(fn_rsq(&[arr(&[1.0, 2.0, 3.0]), arr(&[2.0, 4.0, 6.0])], &c).unwrap());
            assert!((p - 1.0).abs() < 1e-12);
            assert!((r2 - 1.0).abs() < 1e-12);
        }

        #[test]
        fn test_slope_and_intercept() {
            let c = ctx();
            let slope =
                as_number(fn_slope(&[arr(&[3.0, 5.0, 7.0]), arr(&[1.0, 2.0, 3.0])], &c).unwrap());
            let intercept = as_number(
                fn_intercept(&[arr(&[3.0, 5.0, 7.0]), arr(&[1.0, 2.0, 3.0])], &c).unwrap(),
            );
            assert!((slope - 2.0).abs() < 1e-12);
            assert!((intercept - 1.0).abs() < 1e-12);
        }

        #[test]
        fn test_fisher_and_inverse() {
            let c = ctx();
            let f = as_number(fn_fisher(&[n(0.5)], &c).unwrap());
            let inv = as_number(fn_fisherinv(&[n(f)], &c).unwrap());
            assert!((inv - 0.5).abs() < 1e-10);
            assert_eq!(
                fn_fisher(&[n(1.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_forecast_functions() {
            let c = ctx();
            let args = [n(4.0), arr(&[2.0, 4.0, 6.0]), arr(&[1.0, 2.0, 3.0])];
            let f1 = as_number(fn_forecast_linear(&args, &c).unwrap());
            let f2 = as_number(fn_forecast(&args, &c).unwrap());
            assert!((f1 - 8.0).abs() < 1e-12);
            assert!((f2 - 8.0).abs() < 1e-12);
        }

        #[test]
        fn test_frequency() {
            let c = ctx();
            let out = fn_frequency(&[arr(&[1.0, 2.0, 3.0, 10.0]), arr(&[2.0, 5.0])], &c).unwrap();
            assert_eq!(
                out,
                FormulaValue::Array(vec![vec![n(2.0)], vec![n(1.0)], vec![n(1.0)]])
            );
            let out2 = fn_frequency(&[arr(&[]), arr(&[1.0])], &c).unwrap();
            assert_eq!(out2, FormulaValue::Array(vec![vec![n(0.0)], vec![n(0.0)]]));
        }

        #[test]
        fn test_maxa_mina() {
            let c = ctx();
            let maxa = as_number(
                fn_maxa(
                    &[
                        FormulaValue::Boolean(true),
                        FormulaValue::String("x".into()),
                        n(3.0),
                    ],
                    &c,
                )
                .unwrap(),
            );
            let mina = as_number(
                fn_mina(
                    &[
                        FormulaValue::Boolean(false),
                        FormulaValue::String("x".into()),
                        n(3.0),
                    ],
                    &c,
                )
                .unwrap(),
            );
            assert!((maxa - 3.0).abs() < 1e-12);
            assert!((mina - 0.0).abs() < 1e-12);
        }

        #[test]
        fn test_steyx() {
            let c = ctx();
            let near_zero =
                as_number(fn_steyx(&[arr(&[2.0, 4.0, 6.0]), arr(&[1.0, 2.0, 3.0])], &c).unwrap());
            assert!(near_zero < 1e-10);
            assert_eq!(
                fn_steyx(&[arr(&[1.0, 2.0]), arr(&[1.0, 2.0])], &c).unwrap(),
                FormulaValue::Error(CellError::Div0)
            );
        }

        #[test]
        fn test_prob() {
            let c = ctx();
            let p = as_number(
                fn_prob(
                    &[arr(&[1.0, 2.0, 3.0]), arr(&[0.2, 0.3, 0.5]), n(2.0), n(3.0)],
                    &c,
                )
                .unwrap(),
            );
            assert!((p - 0.8).abs() < 1e-12);
            assert_eq!(
                fn_prob(&[arr(&[1.0, 2.0]), arr(&[0.4, 0.4]), n(1.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_permutation_functions() {
            let c = ctx();
            assert_eq!(fn_permut(&[n(5.0), n(3.0)], &c).unwrap(), n(60.0));
            assert_eq!(
                fn_permut(&[n(3.0), n(5.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
            assert_eq!(fn_permutationa(&[n(3.0), n(2.0)], &c).unwrap(), n(9.0));
            assert_eq!(
                fn_permutationa(&[n(-1.0), n(2.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_confidence_norm() {
            let c = ctx();
            let v = as_number(fn_confidence_norm(&[n(0.05), n(2.0), n(100.0)], &c).unwrap());
            assert!((v - 0.39199).abs() < 0.02);
            assert_eq!(
                fn_confidence_norm(&[n(1.0), n(1.0), n(10.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_confidence_t() {
            let c = ctx();
            let v = as_number(fn_confidence_t(&[n(0.05), n(2.0), n(25.0)], &c).unwrap());
            assert!(v > 0.7 && v < 0.95);
            assert_eq!(
                fn_confidence_t(&[n(0.05), n(2.0), n(1.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }

        #[test]
        fn test_gauss() {
            let c = ctx();
            let z0 = as_number(fn_gauss(&[n(0.0)], &c).unwrap());
            let z1 = as_number(fn_gauss(&[n(1.0)], &c).unwrap());
            assert!(z0.abs() < 1e-12);
            assert!((z1 - 0.3413).abs() < 1e-3);
        }

        #[test]
        fn test_mode_mult() {
            let c = ctx();
            let out = fn_mode_mult(&[arr(&[1.0, 2.0, 2.0, 3.0, 3.0])], &c).unwrap();
            assert_eq!(out, FormulaValue::Array(vec![vec![n(2.0)], vec![n(3.0)]]));
            assert_eq!(
                fn_mode_mult(&[arr(&[1.0, 2.0, 3.0])], &c).unwrap(),
                FormulaValue::Error(CellError::Na)
            );
        }

        #[test]
        fn test_a_variants_variance_stdev() {
            let c = ctx();
            let vals = [
                n(1.0),
                FormulaValue::Boolean(true),
                FormulaValue::String("x".into()),
            ];
            assert!(as_number(fn_stdeva(&vals, &c).unwrap()) > 0.0);
            assert!(as_number(fn_stdevpa(&vals, &c).unwrap()) > 0.0);
            assert!(as_number(fn_vara(&vals, &c).unwrap()) > 0.0);
            assert!(as_number(fn_varpa(&vals, &c).unwrap()) > 0.0);
            assert_eq!(
                fn_stdeva(&[n(1.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
            assert_eq!(
                fn_vara(&[n(1.0)], &c).unwrap(),
                FormulaValue::Error(CellError::Num)
            );
        }
    }
}
