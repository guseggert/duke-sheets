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
