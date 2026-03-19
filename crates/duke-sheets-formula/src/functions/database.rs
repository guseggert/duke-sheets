use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

#[derive(Clone, Copy)]
enum ComparisonOp {
    Equal,
    NotEqual,
    LessThan,
    LessEqual,
    GreaterThan,
    GreaterEqual,
}

fn database_filter(args: &[FormulaValue]) -> Result<Vec<FormulaValue>, FormulaValue> {
    if args.len() != 3 {
        return Err(FormulaValue::Error(CellError::Value));
    }

    let database = match &args[0] {
        FormulaValue::Array { data: arr, .. } => arr,
        FormulaValue::Error(e) => return Err(FormulaValue::Error(*e)),
        _ => return Err(FormulaValue::Error(CellError::Value)),
    };

    if database.is_empty() || database[0].is_empty() {
        return Err(FormulaValue::Error(CellError::Value));
    }

    let headers = &database[0];
    let field_idx = resolve_field_index(&args[1], headers)?;

    let criteria = match &args[2] {
        FormulaValue::Array { data: arr, .. } => arr,
        FormulaValue::Error(e) => return Err(FormulaValue::Error(*e)),
        _ => return Err(FormulaValue::Error(CellError::Value)),
    };

    if criteria.is_empty() {
        return Err(FormulaValue::Error(CellError::Value));
    }

    for row in criteria {
        for cell in row {
            if let FormulaValue::Error(e) = cell {
                return Err(FormulaValue::Error(*e));
            }
        }
    }

    let criteria_headers = &criteria[0];
    let criteria_rows = &criteria[1..];
    let criteria_cols = resolve_criteria_columns(criteria_headers, headers);

    let mut out = Vec::new();
    for row in database.iter().skip(1) {
        if row_matches_criteria(row, criteria_rows, &criteria_cols) {
            out.push(row.get(field_idx).cloned().unwrap_or(FormulaValue::Empty));
        }
    }

    Ok(out)
}

fn resolve_field_index(
    field: &FormulaValue,
    headers: &[FormulaValue],
) -> Result<usize, FormulaValue> {
    match field {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Number(n) => {
            let idx = n.trunc() as isize;
            if idx < 1 || (idx as usize) > headers.len() {
                Err(FormulaValue::Error(CellError::Value))
            } else {
                Ok((idx - 1) as usize)
            }
        }
        FormulaValue::String(s) => {
            let wanted = s.trim();
            if wanted.is_empty() {
                return Err(FormulaValue::Error(CellError::Value));
            }

            headers
                .iter()
                .position(|h| h.as_string().eq_ignore_ascii_case(wanted))
                .ok_or(FormulaValue::Error(CellError::Value))
        }
        _ => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn resolve_criteria_columns(
    criteria_headers: &[FormulaValue],
    database_headers: &[FormulaValue],
) -> Vec<Option<usize>> {
    criteria_headers
        .iter()
        .map(|h| {
            let header = h.as_string();
            if header.trim().is_empty() {
                None
            } else {
                database_headers
                    .iter()
                    .position(|db_h| db_h.as_string().eq_ignore_ascii_case(header.trim()))
            }
        })
        .collect()
}

fn row_matches_criteria(
    data_row: &[FormulaValue],
    criteria_rows: &[Vec<FormulaValue>],
    criteria_cols: &[Option<usize>],
) -> bool {
    if criteria_rows.is_empty() {
        return true;
    }

    for crit_row in criteria_rows {
        let mut row_match = true;

        for (col_idx, mapped_db_col) in criteria_cols.iter().enumerate() {
            let crit_cell = crit_row.get(col_idx).unwrap_or(&FormulaValue::Empty);
            if is_empty_criteria_cell(crit_cell) {
                continue;
            }

            let Some(db_col) = mapped_db_col else {
                row_match = false;
                break;
            };

            let value_cell = data_row.get(*db_col).unwrap_or(&FormulaValue::Empty);
            if !criteria_matches_value(crit_cell, value_cell) {
                row_match = false;
                break;
            }
        }

        if row_match {
            return true;
        }
    }

    false
}

fn is_empty_criteria_cell(value: &FormulaValue) -> bool {
    matches!(value, FormulaValue::Empty)
        || matches!(value, FormulaValue::String(s) if s.trim().is_empty())
}

fn criteria_matches_value(criteria: &FormulaValue, value: &FormulaValue) -> bool {
    match criteria {
        FormulaValue::Empty => true,
        FormulaValue::Error(_) | FormulaValue::Array { .. } => false,
        FormulaValue::Number(n) => {
            numeric_value_strict(value).is_some_and(|v| (v - *n).abs() < 1e-10)
        }
        FormulaValue::Boolean(b) => {
            let wanted = if *b { 1.0 } else { 0.0 };
            numeric_value_strict(value).is_some_and(|v| (v - wanted).abs() < 1e-10)
        }
        FormulaValue::String(s) => {
            let trimmed = s.trim();
            if trimmed.is_empty() {
                return true;
            }

            if let Some((op, operand)) = parse_comparison_criteria(trimmed) {
                return compare_with_operator(value, op, operand.trim());
            }

            let pattern = trimmed.to_lowercase();
            let text = value.as_string().to_lowercase();
            wildcard_match(&pattern, &text)
        }
    }
}

fn parse_comparison_criteria(s: &str) -> Option<(ComparisonOp, &str)> {
    if let Some(rest) = s.strip_prefix(">=") {
        return Some((ComparisonOp::GreaterEqual, rest));
    }
    if let Some(rest) = s.strip_prefix("<=") {
        return Some((ComparisonOp::LessEqual, rest));
    }
    if let Some(rest) = s.strip_prefix("<>") {
        return Some((ComparisonOp::NotEqual, rest));
    }
    if let Some(rest) = s.strip_prefix('>') {
        return Some((ComparisonOp::GreaterThan, rest));
    }
    if let Some(rest) = s.strip_prefix('<') {
        return Some((ComparisonOp::LessThan, rest));
    }
    if let Some(rest) = s.strip_prefix('=') {
        return Some((ComparisonOp::Equal, rest));
    }
    None
}

fn compare_with_operator(value: &FormulaValue, op: ComparisonOp, operand: &str) -> bool {
    if let Ok(operand_num) = operand.parse::<f64>() {
        if let Some(value_num) = numeric_value_strict(value) {
            return match op {
                ComparisonOp::Equal => (value_num - operand_num).abs() < 1e-10,
                ComparisonOp::NotEqual => (value_num - operand_num).abs() >= 1e-10,
                ComparisonOp::LessThan => value_num < operand_num,
                ComparisonOp::LessEqual => value_num <= operand_num,
                ComparisonOp::GreaterThan => value_num > operand_num,
                ComparisonOp::GreaterEqual => value_num >= operand_num,
            };
        }
    }

    let value_text = value.as_string().to_lowercase();
    let operand_text = operand.to_lowercase();

    match op {
        ComparisonOp::Equal => wildcard_match(&operand_text, &value_text),
        ComparisonOp::NotEqual => !wildcard_match(&operand_text, &value_text),
        ComparisonOp::LessThan => value_text < operand_text,
        ComparisonOp::LessEqual => value_text <= operand_text,
        ComparisonOp::GreaterThan => value_text > operand_text,
        ComparisonOp::GreaterEqual => value_text >= operand_text,
    }
}

fn wildcard_match(pattern: &str, text: &str) -> bool {
    if !pattern.contains('*') && !pattern.contains('?') {
        return pattern == text;
    }

    let pattern_chars: Vec<char> = pattern.chars().collect();
    let text_chars: Vec<char> = text.chars().collect();

    let mut pi = 0usize;
    let mut ti = 0usize;
    let mut star_pi = None;
    let mut star_ti = 0usize;

    while ti < text_chars.len() {
        if pi < pattern_chars.len()
            && (pattern_chars[pi] == '?' || pattern_chars[pi] == text_chars[ti])
        {
            pi += 1;
            ti += 1;
        } else if pi < pattern_chars.len() && pattern_chars[pi] == '*' {
            star_pi = Some(pi);
            star_ti = ti;
            pi += 1;
        } else if let Some(sp) = star_pi {
            pi = sp + 1;
            star_ti += 1;
            ti = star_ti;
        } else {
            return false;
        }
    }

    while pi < pattern_chars.len() && pattern_chars[pi] == '*' {
        pi += 1;
    }

    pi == pattern_chars.len()
}

fn numeric_value_strict(value: &FormulaValue) -> Option<f64> {
    match value {
        FormulaValue::Number(n) => Some(*n),
        FormulaValue::Boolean(true) => Some(1.0),
        FormulaValue::Boolean(false) => Some(0.0),
        _ => None,
    }
}

fn numeric_values(values: &[FormulaValue]) -> Result<Vec<f64>, FormulaValue> {
    let mut nums = Vec::new();
    for v in values {
        match v {
            FormulaValue::Error(e) => return Err(FormulaValue::Error(*e)),
            FormulaValue::Number(n) => nums.push(*n),
            _ => {}
        }
    }
    Ok(nums)
}

fn variance(values: &[f64], sample: bool) -> f64 {
    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;
    let sum_sq = values.iter().map(|x| (x - mean) * (x - mean)).sum::<f64>();
    if sample {
        sum_sq / (n - 1.0)
    } else {
        sum_sq / n
    }
}

pub fn fn_daverage(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if nums.is_empty() {
        Ok(FormulaValue::Error(CellError::Div0))
    } else {
        Ok(FormulaValue::Number(
            nums.iter().sum::<f64>() / nums.len() as f64,
        ))
    }
}

pub fn fn_dcount(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let count = match numeric_values(&values) {
        Ok(v) => v.len(),
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(count as f64))
}

pub fn fn_dcounta(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let mut count = 0usize;
    for v in values {
        match v {
            FormulaValue::Empty => {}
            FormulaValue::String(s) if s.is_empty() => {}
            _ => count += 1,
        }
    }
    Ok(FormulaValue::Number(count as f64))
}

pub fn fn_dget(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if values.len() == 1 {
        Ok(values[0].clone())
    } else if values.is_empty() {
        Ok(FormulaValue::Error(CellError::Value))
    } else {
        Ok(FormulaValue::Error(CellError::Num))
    }
}

pub fn fn_dmax(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let max = nums.iter().copied().reduce(f64::max).unwrap_or(0.0);
    Ok(FormulaValue::Number(max))
}

pub fn fn_dmin(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let min = nums.iter().copied().reduce(f64::min).unwrap_or(0.0);
    Ok(FormulaValue::Number(min))
}

pub fn fn_dproduct(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if nums.is_empty() {
        return Ok(FormulaValue::Number(0.0));
    }

    let product = nums.iter().product::<f64>();
    Ok(FormulaValue::Number(product))
}

pub fn fn_dstdev(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if nums.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(variance(&nums, true).sqrt()))
}

pub fn fn_dstdevp(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if nums.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(variance(&nums, false).sqrt()))
}

pub fn fn_dsum(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    Ok(FormulaValue::Number(nums.iter().sum::<f64>()))
}

pub fn fn_dvar(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if nums.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(variance(&nums, true)))
}

pub fn fn_dvarp(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let values = match database_filter(args) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let nums = match numeric_values(&values) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if nums.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    Ok(FormulaValue::Number(variance(&nums, false)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::evaluator::evaluate;
    use crate::parser::parse_formula;

    const DB: &str = r#"{"Tree","Height","Age","Yield";"Apple",18,20,14;"Pear",12,12,10;"Cherry",13,14,9;"Apple",14,15,11;"Pear",9,8,8;"Apple",8,9,6}"#;

    fn eval(formula: &str) -> Result<FormulaValue, crate::error::FormulaError> {
        let expr = parse_formula(formula)?;
        let ctx = EvaluationContext::simple();
        evaluate(&expr, &ctx)
    }

    fn eval_db(func: &str, field: &str, criteria: &str) -> FormulaValue {
        let formula = format!("={}({}, {}, {})", func, DB, field, criteria);
        eval(&formula).unwrap()
    }

    fn assert_close(actual: FormulaValue, expected: f64) {
        match actual {
            FormulaValue::Number(n) => {
                assert!((n - expected).abs() < 1e-9, "got {n}, expected {expected}")
            }
            other => panic!("Expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_daverage() {
        assert_close(
            eval_db("DAVERAGE", "\"Height\"", r#"{"Tree";"Apple"}"#),
            40.0 / 3.0,
        );
        assert_close(
            eval_db("DAVERAGE", "3", r#"{"Tree";"Apple";"Pear"}"#),
            64.0 / 5.0,
        );
        assert_eq!(
            eval_db("DAVERAGE", "\"Height\"", r#"{"Tree";"Banana"}"#),
            FormulaValue::Error(CellError::Div0)
        );
    }

    #[test]
    fn test_dcount() {
        assert_eq!(
            eval_db("DCOUNT", "\"Height\"", r#"{"Tree";"Apple"}"#),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            eval_db("DCOUNT", "2", r#"{"Age";">10"}"#),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            eval_db("DCOUNT", "\"Height\"", r#"{"Tree";"Banana"}"#),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_dcounta() {
        assert_eq!(
            eval_db("DCOUNTA", "\"Tree\"", r#"{"Age";">10"}"#),
            FormulaValue::Number(4.0)
        );
        assert_eq!(
            eval_db("DCOUNTA", "4", r#"{"Tree";"Apple";"Pear"}"#),
            FormulaValue::Number(5.0)
        );

        assert_eq!(
            eval(r#"=DCOUNTA({"A","Note";"x","";"y",1;"z",2},"Note",{"A";"*"})"#).unwrap(),
            FormulaValue::Number(2.0)
        );
    }

    #[test]
    fn test_dget() {
        assert_eq!(
            eval_db("DGET", "\"Height\"", r#"{"Tree";"Cherry"}"#),
            FormulaValue::Number(13.0)
        );
        assert_eq!(
            eval_db("DGET", "\"Height\"", r#"{"Tree";"Apple"}"#),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            eval_db("DGET", "\"Height\"", r#"{"Tree";"Banana"}"#),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_dmax() {
        assert_eq!(
            eval_db("DMAX", "\"Height\"", r#"{"Tree";"Apple"}"#),
            FormulaValue::Number(18.0)
        );
        assert_eq!(
            eval_db("DMAX", "4", r#"{"Tree";"Apple";"Pear"}"#),
            FormulaValue::Number(14.0)
        );
        assert_eq!(
            eval_db("DMAX", "\"Height\"", r#"{"Tree";"Banana"}"#),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_dmin() {
        assert_eq!(
            eval_db("DMIN", "\"Height\"", r#"{"Tree";"Apple"}"#),
            FormulaValue::Number(8.0)
        );
        assert_eq!(
            eval_db("DMIN", "3", r#"{"Tree";"Apple";"Pear"}"#),
            FormulaValue::Number(8.0)
        );
        assert_eq!(
            eval_db("DMIN", "\"Height\"", r#"{"Tree";"Banana"}"#),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_dproduct() {
        assert_eq!(
            eval_db("DPRODUCT", "\"Yield\"", r#"{"Tree";"Apple"}"#),
            FormulaValue::Number(924.0)
        );
        assert_eq!(
            eval_db("DPRODUCT", "2", r#"{"Tree";"A*"}"#),
            FormulaValue::Number(2016.0)
        );
        assert_eq!(
            eval_db("DPRODUCT", "\"Yield\"", r#"{"Tree";"Banana"}"#),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_dstdev() {
        assert_close(
            eval_db("DSTDEV", "\"Height\"", r#"{"Tree";"Apple"}"#),
            5.033222956847166,
        );
        assert_close(
            eval_db("DSTDEV", "2", r#"{"Age";">10"}"#),
            2.6299556396765835,
        );
        assert_eq!(
            eval_db("DSTDEV", "\"Height\"", r#"{"Tree";"Cherry"}"#),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_dstdevp() {
        assert_close(
            eval_db("DSTDEVP", "\"Height\"", r#"{"Tree";"Apple"}"#),
            4.109609335312651,
        );
        assert_close(eval_db("DSTDEVP", "2", r#"{"Tree";"Apple";"Pear"}"#), 3.6);
        assert_eq!(
            eval_db("DSTDEVP", "\"Height\"", r#"{"Tree";"Banana"}"#),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_dsum() {
        assert_eq!(
            eval_db("DSUM", "\"Height\"", r#"{"Tree";"Apple"}"#),
            FormulaValue::Number(40.0)
        );
        assert_eq!(
            eval_db("DSUM", "3", r#"{"Tree";"Apple";"Pear"}"#),
            FormulaValue::Number(64.0)
        );
        assert_eq!(
            eval_db("DSUM", "\"Height\"", r#"{"Age";">=15"}"#),
            FormulaValue::Number(32.0)
        );
    }

    #[test]
    fn test_dvar() {
        assert_close(
            eval_db("DVAR", "\"Height\"", r#"{"Tree";"Apple"}"#),
            25.333333333333332,
        );
        assert_close(eval_db("DVAR", "2", r#"{"Age";">10"}"#), 6.916666666666667);
        assert_eq!(
            eval_db("DVAR", "\"Height\"", r#"{"Tree";"Cherry"}"#),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_dvarp() {
        assert_close(
            eval_db("DVARP", "\"Height\"", r#"{"Tree";"Apple"}"#),
            16.88888888888889,
        );
        assert_close(eval_db("DVARP", "2", r#"{"Tree";"Apple";"Pear"}"#), 12.96);
        assert_eq!(
            eval_db("DVARP", "\"Height\"", r#"{"Tree";"Banana"}"#),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_criteria_and_logic() {
        assert_eq!(
            eval_db("DSUM", "\"Yield\"", r#"{"Tree","Age";"Apple",">10"}"#),
            FormulaValue::Number(25.0)
        );
        assert_eq!(
            eval_db("DSUM", "\"Yield\"", r#"{"Tree";"A?ple"}"#),
            FormulaValue::Number(31.0)
        );
        assert_eq!(
            eval_db("DSUM", "\"Yield\"", r#"{"Tree";"<>Apple"}"#),
            FormulaValue::Number(27.0)
        );
    }

    // Docs DB (includes Profit column, different Yield values)
    const DB_DOCS: &str = r#"{"Tree","Height","Age","Yield","Profit";"Apple",18,20,14,105;"Pear",12,12,10,96;"Cherry",13,14,9,105;"Apple",14,15,10,75;"Pear",9,8,8,76.8;"Apple",8,9,6,45}"#;

    #[test]
    fn test_daverage_docs() {
        // Docs Example 1: Criteria: Tree="Apple" AND Height>10
        // Matching Yield: 14, 10 → avg = 12
        let crit1 = r#"{"Tree","Height";"Apple",">10"}"#;
        assert_close(
            eval(&format!(r#"=DAVERAGE({}, "Yield", {})"#, DB_DOCS, crit1)).unwrap(),
            12.0,
        );

        // Docs Example 2: Criteria = database itself (all rows), field=3 (Age)
        // Average Age = (20+12+14+15+8+9)/6 = 13
        assert_close(
            eval(&format!("=DAVERAGE({}, 3, {})", DB_DOCS, DB_DOCS)).unwrap(),
            13.0,
        );
    }

    #[test]
    fn test_dcount_docs() {
        // Docs: =DCOUNT(db, "Age", criteria) = 1
        // DB has Apple row with "N/A" text in Age (not a number)
        // Criteria: Tree="Apple" AND Height>10 AND Height<16
        // Two Apples match height (14 and 12), but one has N/A Age → count 1
        assert_eq!(
            eval(
                r#"=DCOUNT({"Tree","Height","Age","Yield","Profit";"Apple",18,20,14,105;"Pear",12,12,10,96;"Cherry",13,14,9,105;"Apple",14,"N/A",10,75;"Pear",9,8,8,77;"Apple",12,11,6,45}, "Age", {"Tree","Height","Height";"Apple",">10","<16"})"#
            ).unwrap(),
            FormulaValue::Number(1.0)
        );
    }

    #[test]
    fn test_dget_docs() {
        // Docs: Multiple matches → #NUM!
        assert_eq!(
            eval_db("DGET", "\"Yield\"", r#"{"Tree";"Apple";"Pear"}"#),
            FormulaValue::Error(CellError::Num)
        );
        // Docs: Single match (Apple AND Height>10 AND Height<16) → Yield=11
        assert_eq!(
            eval_db(
                "DGET",
                "\"Yield\"",
                r#"{"Tree","Height","Height";"Apple",">10","<16";"Pear",">12",""}"#
            ),
            FormulaValue::Number(11.0)
        );
    }

    #[test]
    fn test_dmax_docs() {
        // Docs criteria: (Apple AND Height>10 AND Height<16) OR Pear
        // Matching Yield: 11, 10, 8 → max=11
        assert_eq!(
            eval_db(
                "DMAX",
                "\"Yield\"",
                r#"{"Tree","Height","Height";"Apple",">10","<16";"Pear","",""}"#
            ),
            FormulaValue::Number(11.0)
        );
    }

    #[test]
    fn test_dmin_docs() {
        // Same criteria as DMAX, Yield: 11, 10, 8 → min=8
        assert_eq!(
            eval_db(
                "DMIN",
                "\"Yield\"",
                r#"{"Tree","Height","Height";"Apple",">10","<16";"Pear","",""}"#
            ),
            FormulaValue::Number(8.0)
        );
    }

    #[test]
    fn test_dproduct_docs() {
        // Same criteria, Yield: 11 × 10 × 8 = 880
        assert_eq!(
            eval_db(
                "DPRODUCT",
                "\"Yield\"",
                r#"{"Tree","Height","Height";"Apple",">10","<16";"Pear","",""}"#,
            ),
            FormulaValue::Number(880.0),
        );
    }

    #[test]
    fn test_dstdev_docs() {
        // Docs criteria: Tree="Apple" OR Tree="Pear"
        // Our DB yields for Apple+Pear: [14, 10, 11, 8, 6]
        // Sample stdev = sqrt(9.2) ≈ 3.033150
        assert_close(
            eval_db("DSTDEV", "\"Yield\"", r#"{"Tree";"Apple";"Pear"}"#),
            9.2_f64.sqrt(),
        );
    }

    #[test]
    fn test_dstdevp_docs() {
        // Same criteria, population stdev
        // Pop variance = 36.8/5 = 7.36, stdev = sqrt(7.36) ≈ 2.71293
        assert_close(
            eval_db("DSTDEVP", "\"Yield\"", r#"{"Tree";"Apple";"Pear"}"#),
            7.36_f64.sqrt(),
        );
    }

    #[test]
    fn test_dsum_docs() {
        // Docs Example 1: All Apple Yield: 14+11+6 = 31
        assert_eq!(
            eval_db("DSUM", "\"Yield\"", r#"{"Tree";"Apple"}"#),
            FormulaValue::Number(31.0)
        );

        // Docs Example 2: (Apple AND Height>10 AND <16) OR Pear
        // Yield: 11 + 10 + 8 = 29
        assert_eq!(
            eval_db(
                "DSUM",
                "\"Yield\"",
                r#"{"Tree","Height","Height";"Apple",">10","<16";"Pear","",""}"#
            ),
            FormulaValue::Number(29.0)
        );
    }
    #[test]
    fn test_dvar_docs() {
        // Docs example: =DVAR(database, "Yield", A1:A3)
        // Criteria: Tree = "Apple" OR Tree = "Pear"
        // Matching yields: 14, 10, 11, 8, 6 → sample variance = 9.2
        assert_close(
            eval_db("DVAR", "\"Yield\"", r#"{"Tree";"Apple";"Pear"}"#),
            9.2,
        );
    }

    #[test]
    fn test_dvarp_docs() {
        // Docs example: DVARP(database, "Yield", {Tree; Apple; Pear})
        // Matching Yield values: {14, 10, 11, 8, 6}, population variance = 184/25
        assert_close(
            eval_db("DVARP", "\"Yield\"", r#"{"Tree";"Apple";"Pear"}"#),
            7.36,
        );
    }

    #[test]
    fn test_dcounta_docs() {
        // === Main example ===
        // DB with Profit column; Tree="Apple" AND Height>10 AND Height<16
        // Only Apple,14,15,10,75 matches → Profit=75 non-blank → 1
        let db_profit = r#"{"Tree","Height","Age","Yield","Profit";"Apple",18,20,14,105;"Pear",12,12,10,96;"Cherry",13,14,9,105;"Apple",14,15,10,75;"Pear",9,8,8,76.8;"Apple",8,9,6,45}"#;
        assert_eq!(
            eval(&format!(
                r#"=DCOUNTA({}, "Profit", {{"Tree","Height","Height";"Apple",">10","<16"}})"#,
                db_profit
            ))
            .unwrap(),
            FormulaValue::Number(1.0)
        );

        // Criteria examples share a 4-row sales database
        let db4 = r#"{"Category","Salesperson","Sales";"Beverages","Suyama",5122;"Meat","Davolio",450;"produce","Buchanan",6328;"Produce","Davolio",6544}"#;

        // === Multiple criteria in one column (OR) ===
        // Salesperson="Davolio" OR Salesperson="Buchanan" → 3
        assert_eq!(
            eval(&format!(
                r#"=DCOUNTA({}, 2, {{"Salesperson";"Davolio";"Buchanan"}})"#,
                db4
            ))
            .unwrap(),
            FormulaValue::Number(3.0)
        );

        // === Multiple criteria AND, field=1 ===
        // 6-row DB; Category="Produce" AND Sales>2000 → 2
        let db6 = r#"{"Category","Salesperson","Sales";"Beverages","Suyama",5122;"Meat","Davolio",450;"Produce","Buchanan",935;"Produce","Davolio",6544;"Beverages","Buchanan",3677;"Produce","Davolio",3186}"#;
        assert_eq!(
            eval(&format!(
                r#"=DCOUNTA({}, 1, {{"Category","Sales";"Produce",">2000"}})"#,
                db6
            ))
            .unwrap(),
            FormulaValue::Number(2.0)
        );

        // === Multiple sets with AND per row, OR between rows, field=1 ===
        // (Davolio AND Sales>3000) OR (Buchanan AND Sales>1500) → 2
        assert_eq!(
            eval(&format!(
                r#"=DCOUNTA({}, 1, {{"Salesperson","Sales";"Davolio",">3000";"Buchanan",">1500"}})"#,
                db4
            ))
            .unwrap(),
            FormulaValue::Number(2.0)
        );

        // === Duplicate column headers for range criteria, field=1 ===
        // (Sales>6000 AND Sales<6500) OR Sales<500 → 2
        assert_eq!(
            eval(&format!(
                r#"=DCOUNTA({}, 1, {{"Sales","Sales";">6000","<6500";"<500",""}})"#,
                db4
            ))
            .unwrap(),
            FormulaValue::Number(2.0)
        );

        // === Wildcard criteria, field=1 ===
        // Category starts with "Me" OR Salesperson matches ?u* → 3
        assert_eq!(
            eval(&format!(
                r#"=DCOUNTA({}, 1, {{"Category","Salesperson";"Me*","";"","?u*"}})"#,
                db4
            ))
            .unwrap(),
            FormulaValue::Number(3.0)
        );

        // === Formula-based criteria (pre-computed average = 4611) ===
        // field=1; Sales > 4611 → 3
        assert_eq!(
            eval(&format!(r#"=DCOUNTA({}, 1, {{"Sales";">4611"}})"#, db4)).unwrap(),
            FormulaValue::Number(3.0)
        );
    }
}
