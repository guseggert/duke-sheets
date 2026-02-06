//! Formula evaluator
//!
//! Evaluates formula ASTs to produce values.

use crate::ast::{BinaryOperator, FormulaExpr, UnaryOperator};
use crate::error::{FormulaError, FormulaResult};
use crate::functions::FunctionRegistry;
use duke_sheets_core::{CellError, CellValue, Workbook};
use std::sync::OnceLock;

/// Global function registry (lazily initialized)
static FUNCTION_REGISTRY: OnceLock<FunctionRegistry> = OnceLock::new();

fn get_function_registry() -> &'static FunctionRegistry {
    FUNCTION_REGISTRY.get_or_init(FunctionRegistry::new)
}

/// Value types during formula evaluation
#[derive(Debug, Clone, PartialEq)]
pub enum FormulaValue {
    Number(f64),
    String(String),
    Boolean(bool),
    Error(CellError),
    Array(Vec<Vec<FormulaValue>>),
    Empty,
}

impl FormulaValue {
    /// Convert to number, if possible
    pub fn as_number(&self) -> Option<f64> {
        match self {
            FormulaValue::Number(n) => Some(*n),
            FormulaValue::Boolean(true) => Some(1.0),
            FormulaValue::Boolean(false) => Some(0.0),
            FormulaValue::String(s) => s.parse().ok(),
            FormulaValue::Empty => Some(0.0),
            _ => None,
        }
    }

    /// Force conversion to number for arithmetic
    pub fn to_number(&self) -> FormulaResult<f64> {
        self.as_number()
            .ok_or_else(|| FormulaError::Evaluation(format!("Cannot convert {:?} to number", self)))
    }

    /// Convert to boolean
    pub fn as_bool(&self) -> Option<bool> {
        match self {
            FormulaValue::Boolean(b) => Some(*b),
            FormulaValue::Number(n) => Some(*n != 0.0),
            FormulaValue::String(s) => {
                let upper = s.to_uppercase();
                if upper == "TRUE" {
                    Some(true)
                } else if upper == "FALSE" {
                    Some(false)
                } else {
                    None
                }
            }
            _ => None,
        }
    }

    /// Convert to string
    pub fn as_string(&self) -> String {
        match self {
            FormulaValue::Number(n) => {
                // Format like Excel: no trailing zeros, but reasonable precision
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            FormulaValue::String(s) => s.clone(),
            FormulaValue::Boolean(true) => "TRUE".to_string(),
            FormulaValue::Boolean(false) => "FALSE".to_string(),
            FormulaValue::Error(e) => e.to_string(),
            FormulaValue::Empty => String::new(),
            FormulaValue::Array(_) => "#VALUE!".to_string(),
        }
    }

    /// Check if this is an error
    pub fn is_error(&self) -> bool {
        matches!(self, FormulaValue::Error(_))
    }

    /// Get the error if this is one
    pub fn get_error(&self) -> Option<CellError> {
        match self {
            FormulaValue::Error(e) => Some(*e),
            _ => None,
        }
    }
}

impl From<CellValue> for FormulaValue {
    fn from(value: CellValue) -> Self {
        match value {
            CellValue::Empty => FormulaValue::Empty,
            CellValue::Number(n) => FormulaValue::Number(n),
            CellValue::String(s) => FormulaValue::String(s.as_str().to_string()),
            CellValue::Boolean(b) => FormulaValue::Boolean(b),
            CellValue::Error(e) => FormulaValue::Error(e),
            CellValue::Formula { cached_value, .. } => cached_value
                .map(|v| (*v).into())
                .unwrap_or(FormulaValue::Empty),
            // SpillTarget values should be resolved by looking up the source cell
            // In this simple conversion, we return Empty - proper resolution
            // happens in the worksheet's get_value methods
            CellValue::SpillTarget { .. } => FormulaValue::Empty,
        }
    }
}

impl From<FormulaValue> for CellValue {
    fn from(value: FormulaValue) -> Self {
        match value {
            FormulaValue::Empty => CellValue::Empty,
            FormulaValue::Number(n) => CellValue::Number(n),
            FormulaValue::String(s) => CellValue::String(s.into()),
            FormulaValue::Boolean(b) => CellValue::Boolean(b),
            FormulaValue::Error(e) => CellValue::Error(e),
            FormulaValue::Array(_) => CellValue::Error(CellError::Value),
        }
    }
}

/// Context for formula evaluation
pub struct EvaluationContext<'a> {
    /// Reference to the workbook for cell lookups
    pub workbook: Option<&'a Workbook>,
    /// Current worksheet index
    pub current_sheet: usize,
    /// Current cell row (for relative references)
    pub current_row: u32,
    /// Current cell column (for relative references)
    pub current_col: u16,
}

impl<'a> EvaluationContext<'a> {
    /// Create a new evaluation context
    pub fn new(workbook: Option<&'a Workbook>, sheet: usize, row: u32, col: u16) -> Self {
        Self {
            workbook,
            current_sheet: sheet,
            current_row: row,
            current_col: col,
        }
    }

    /// Create a simple context without workbook (for testing)
    pub fn simple() -> Self {
        Self {
            workbook: None,
            current_sheet: 0,
            current_row: 0,
            current_col: 0,
        }
    }

    /// Get a cell value from the workbook
    pub fn get_cell_value(&self, sheet: Option<&str>, row: u32, col: u16) -> FormulaValue {
        let workbook = match self.workbook {
            Some(wb) => wb,
            None => return FormulaValue::Empty,
        };

        let sheet_idx = match sheet {
            Some(name) => match workbook.sheet_index(name) {
                Some(idx) => idx,
                None => return FormulaValue::Error(CellError::Ref),
            },
            None => self.current_sheet,
        };

        let worksheet = match workbook.worksheet(sheet_idx) {
            Some(ws) => ws,
            None => return FormulaValue::Error(CellError::Ref),
        };

        worksheet.get_value_at(row, col).into()
    }

    /// Get a range of cell values as an array
    pub fn get_range_values(
        &self,
        sheet: Option<&str>,
        start_row: u32,
        start_col: u16,
        end_row: u32,
        end_col: u16,
    ) -> FormulaValue {
        let workbook = match self.workbook {
            Some(wb) => wb,
            None => return FormulaValue::Array(vec![]),
        };

        let sheet_idx = match sheet {
            Some(name) => match workbook.sheet_index(name) {
                Some(idx) => idx,
                None => return FormulaValue::Error(CellError::Ref),
            },
            None => self.current_sheet,
        };

        let worksheet = match workbook.worksheet(sheet_idx) {
            Some(ws) => ws,
            None => return FormulaValue::Error(CellError::Ref),
        };

        let mut rows = Vec::new();
        for row in start_row..=end_row {
            let mut cols = Vec::new();
            for col in start_col..=end_col {
                cols.push(worksheet.get_value_at(row, col).into());
            }
            rows.push(cols);
        }

        FormulaValue::Array(rows)
    }

    /// Resolve a named range to its value
    ///
    /// This handles:
    /// - Cell references: "Sheet1!$A$1" -> cell value
    /// - Range references: "Sheet1!$A$1:$D$10" -> array of values
    /// - Constants: "0.0725" -> number
    /// - Formulas: "=SUM(A1:A10)" -> evaluated formula (recursive)
    pub fn resolve_named_range(&self, name: &str) -> Result<FormulaValue, FormulaError> {
        let workbook = self.workbook.ok_or_else(|| {
            FormulaError::InvalidReference("No workbook context for named range lookup".to_string())
        })?;

        let named_range = workbook
            .get_named_range(name, self.current_sheet)
            .ok_or_else(|| FormulaError::InvalidReference(format!("Unknown name: {}", name)))?;

        let refers_to = &named_range.refers_to;

        // If it's a formula, parse and evaluate it
        if refers_to.starts_with('=') {
            // Keep the '=' since the parser expects it
            let ast = crate::parser::parse_formula(refers_to)?;
            return crate::evaluator::evaluate(&ast, self);
        }

        // Try to parse as a number constant
        if let Ok(num) = refers_to.parse::<f64>() {
            return Ok(FormulaValue::Number(num));
        }

        // Try to parse as a boolean
        let upper = refers_to.to_uppercase();
        if upper == "TRUE" {
            return Ok(FormulaValue::Boolean(true));
        }
        if upper == "FALSE" {
            return Ok(FormulaValue::Boolean(false));
        }

        // Try to parse as a cell or range reference
        // This is a simplified parser - full implementation would reuse the main parser
        self.parse_and_resolve_reference(refers_to)
    }

    /// Parse a reference string and resolve it to a value
    fn parse_and_resolve_reference(&self, refers_to: &str) -> Result<FormulaValue, FormulaError> {
        // Handle sheet!reference format
        let (sheet_name, ref_part) = if let Some(pos) = refers_to.find('!') {
            let sheet = &refers_to[..pos];
            // Remove quotes if present (e.g., 'Sheet 1'!A1)
            let sheet = sheet.trim_matches('\'');
            (Some(sheet), &refers_to[pos + 1..])
        } else {
            (None, refers_to)
        };

        // Remove $ signs (absolute reference markers)
        let ref_clean = ref_part.replace('$', "");

        // Check if it's a range (contains :)
        if let Some(colon_pos) = ref_clean.find(':') {
            let start_ref = &ref_clean[..colon_pos];
            let end_ref = &ref_clean[colon_pos + 1..];

            let (start_row, start_col) = self.parse_cell_address(start_ref)?;
            let (end_row, end_col) = self.parse_cell_address(end_ref)?;

            return Ok(self.get_range_values(sheet_name, start_row, start_col, end_row, end_col));
        }

        // Single cell reference
        let (row, col) = self.parse_cell_address(&ref_clean)?;
        Ok(self.get_cell_value(sheet_name, row, col))
    }

    /// Parse a cell address like "A1" to (row, col)
    fn parse_cell_address(&self, addr: &str) -> Result<(u32, u16), FormulaError> {
        // Find where letters end and numbers begin
        let col_end = addr
            .find(|c: char| c.is_ascii_digit())
            .unwrap_or(addr.len());

        if col_end == 0 || col_end == addr.len() {
            return Err(FormulaError::InvalidReference(format!(
                "Invalid cell address: {}",
                addr
            )));
        }

        let col_str = &addr[..col_end];
        let row_str = &addr[col_end..];

        // Parse column (A=0, B=1, ..., Z=25, AA=26, etc.)
        let col = self.parse_column_letters(col_str)?;

        // Parse row (1-indexed in Excel, convert to 0-indexed)
        let row: u32 = row_str
            .parse()
            .map_err(|_| FormulaError::InvalidReference(format!("Invalid row: {}", row_str)))?;

        if row == 0 {
            return Err(FormulaError::InvalidReference(
                "Row number must be >= 1".to_string(),
            ));
        }

        Ok((row - 1, col))
    }

    /// Parse column letters (A=0, B=1, ..., Z=25, AA=26, etc.)
    fn parse_column_letters(&self, s: &str) -> Result<u16, FormulaError> {
        let s = s.to_uppercase();
        let mut col: u16 = 0;
        for c in s.chars() {
            if !c.is_ascii_uppercase() {
                return Err(FormulaError::InvalidReference(format!(
                    "Invalid column letter: {}",
                    c
                )));
            }
            col = col
                .checked_mul(26)
                .and_then(|v| v.checked_add((c as u16) - ('A' as u16) + 1))
                .ok_or_else(|| {
                    FormulaError::InvalidReference(format!("Column too large: {}", s))
                })?;
        }
        // Convert from 1-indexed to 0-indexed
        Ok(col - 1)
    }
}

/// Evaluate a formula expression
pub fn evaluate(expr: &FormulaExpr, ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    match expr {
        // === Literals ===
        FormulaExpr::Number(n) => Ok(FormulaValue::Number(*n)),
        FormulaExpr::String(s) => Ok(FormulaValue::String(s.clone())),
        FormulaExpr::Boolean(b) => Ok(FormulaValue::Boolean(*b)),
        FormulaExpr::Error(e) => Ok(FormulaValue::Error(*e)),

        // === References ===
        FormulaExpr::CellRef(cell_ref) => Ok(ctx.get_cell_value(
            cell_ref.sheet.as_deref(),
            cell_ref.address.row,
            cell_ref.address.col,
        )),

        FormulaExpr::RangeRef(range_ref) => Ok(ctx.get_range_values(
            range_ref.sheet.as_deref(),
            range_ref.range.start.row,
            range_ref.range.start.col,
            range_ref.range.end.row,
            range_ref.range.end.col,
        )),

        FormulaExpr::NameRef(name) => {
            // Resolve named range through the evaluation context
            ctx.resolve_named_range(name)
        }

        // === Operators ===
        FormulaExpr::BinaryOp { op, left, right } => evaluate_binary_op(*op, left, right, ctx),

        FormulaExpr::UnaryOp { op, operand } => evaluate_unary_op(*op, operand, ctx),

        // === Functions ===
        FormulaExpr::Function { name, args } => evaluate_function(name, args, ctx),

        // === Arrays ===
        FormulaExpr::Array(rows) => {
            let mut result_rows = Vec::new();
            for row in rows {
                let mut result_row = Vec::new();
                for expr in row {
                    result_row.push(evaluate(expr, ctx)?);
                }
                result_rows.push(result_row);
            }
            Ok(FormulaValue::Array(result_rows))
        }
    }
}

/// Evaluate a binary operation
fn evaluate_binary_op(
    op: BinaryOperator,
    left: &FormulaExpr,
    right: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    // Evaluate operands first
    let left_val = evaluate(left, ctx)?;
    let right_val = evaluate(right, ctx)?;

    // Propagate errors
    if let Some(e) = left_val.get_error() {
        return Ok(FormulaValue::Error(e));
    }
    if let Some(e) = right_val.get_error() {
        return Ok(FormulaValue::Error(e));
    }

    match op {
        // Arithmetic operators
        BinaryOperator::Add => {
            let l = left_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            let r = right_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            Ok(FormulaValue::Number(l + r))
        }
        BinaryOperator::Subtract => {
            let l = left_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            let r = right_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            Ok(FormulaValue::Number(l - r))
        }
        BinaryOperator::Multiply => {
            let l = left_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            let r = right_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            Ok(FormulaValue::Number(l * r))
        }
        BinaryOperator::Divide => {
            let l = left_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            let r = right_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            if r == 0.0 {
                Ok(FormulaValue::Error(CellError::Div0))
            } else {
                Ok(FormulaValue::Number(l / r))
            }
        }
        BinaryOperator::Power => {
            let l = left_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            let r = right_val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            let result = l.powf(r);
            if result.is_nan() || result.is_infinite() {
                Ok(FormulaValue::Error(CellError::Num))
            } else {
                Ok(FormulaValue::Number(result))
            }
        }

        // Comparison operators
        BinaryOperator::Equal => Ok(FormulaValue::Boolean(
            compare_values(&left_val, &right_val) == 0,
        )),
        BinaryOperator::NotEqual => Ok(FormulaValue::Boolean(
            compare_values(&left_val, &right_val) != 0,
        )),
        BinaryOperator::LessThan => Ok(FormulaValue::Boolean(
            compare_values(&left_val, &right_val) < 0,
        )),
        BinaryOperator::LessEqual => Ok(FormulaValue::Boolean(
            compare_values(&left_val, &right_val) <= 0,
        )),
        BinaryOperator::GreaterThan => Ok(FormulaValue::Boolean(
            compare_values(&left_val, &right_val) > 0,
        )),
        BinaryOperator::GreaterEqual => Ok(FormulaValue::Boolean(
            compare_values(&left_val, &right_val) >= 0,
        )),

        // Concatenation
        BinaryOperator::Concat => {
            let l = left_val.as_string();
            let r = right_val.as_string();
            Ok(FormulaValue::String(l + &r))
        }

        // Range operators (these shouldn't normally reach evaluation)
        BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect => Err(
            FormulaError::Evaluation("Range operators not supported in this context".into()),
        ),
    }
}

/// Compare two values for ordering (Excel-style comparison)
fn compare_values(left: &FormulaValue, right: &FormulaValue) -> i32 {
    // Empty values
    let left = match left {
        FormulaValue::Empty => &FormulaValue::Number(0.0),
        v => v,
    };
    let right = match right {
        FormulaValue::Empty => &FormulaValue::Number(0.0),
        v => v,
    };

    match (left, right) {
        // Numbers compare numerically
        (FormulaValue::Number(l), FormulaValue::Number(r)) => {
            if l < r {
                -1
            } else if l > r {
                1
            } else {
                0
            }
        }

        // Strings compare case-insensitively
        (FormulaValue::String(l), FormulaValue::String(r)) => {
            l.to_lowercase().cmp(&r.to_lowercase()) as i32
        }

        // Booleans: FALSE < TRUE
        (FormulaValue::Boolean(l), FormulaValue::Boolean(r)) => (*l as i32) - (*r as i32),

        // Mixed types: number < string < boolean
        // (In Excel, numbers are less than text which is less than boolean/logical)
        (FormulaValue::Number(_), FormulaValue::String(_)) => -1,
        (FormulaValue::String(_), FormulaValue::Number(_)) => 1,
        (FormulaValue::Number(_), FormulaValue::Boolean(_)) => -1,
        (FormulaValue::Boolean(_), FormulaValue::Number(_)) => 1,
        (FormulaValue::String(_), FormulaValue::Boolean(_)) => -1,
        (FormulaValue::Boolean(_), FormulaValue::String(_)) => 1,

        // Errors are equal to themselves
        (FormulaValue::Error(l), FormulaValue::Error(r)) => (l.code() as i32) - (r.code() as i32),

        // Other cases
        _ => 0,
    }
}

/// Evaluate a unary operation
fn evaluate_unary_op(
    op: UnaryOperator,
    operand: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let val = evaluate(operand, ctx)?;

    // Propagate errors
    if let Some(e) = val.get_error() {
        return Ok(FormulaValue::Error(e));
    }

    match op {
        UnaryOperator::Negate => {
            let n = val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            Ok(FormulaValue::Number(-n))
        }
        UnaryOperator::Percent => {
            let n = val
                .as_number()
                .ok_or_else(|| FormulaError::Evaluation("Expected number".into()))?;
            Ok(FormulaValue::Number(n / 100.0))
        }
    }
}

/// Evaluate a function call
fn evaluate_function(
    name: &str,
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let registry = get_function_registry();

    let func = registry
        .get(name)
        .ok_or_else(|| FormulaError::UnknownFunction(name.to_string()))?;

    // Check argument count
    if args.len() < func.min_args {
        return Err(FormulaError::ArgumentCount {
            function: name.to_string(),
            expected: format!("at least {}", func.min_args),
            actual: args.len(),
        });
    }

    if let Some(max) = func.max_args {
        if args.len() > max {
            return Err(FormulaError::ArgumentCount {
                function: name.to_string(),
                expected: format!("at most {}", max),
                actual: args.len(),
            });
        }
    }

    // Evaluate arguments
    let mut evaluated_args = Vec::with_capacity(args.len());
    for arg in args {
        evaluated_args.push(evaluate(arg, ctx)?);
    }

    // Call the function
    (func.implementation)(&evaluated_args, ctx)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::parse_formula;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = parse_formula(formula)?;
        let ctx = EvaluationContext::simple();
        evaluate(&ast, &ctx)
    }

    #[test]
    fn test_evaluate_number() {
        assert_eq!(eval("=42").unwrap(), FormulaValue::Number(42.0));
        assert_eq!(eval("=3.14").unwrap(), FormulaValue::Number(3.14));
    }

    #[test]
    fn test_evaluate_string() {
        assert_eq!(
            eval("=\"Hello\"").unwrap(),
            FormulaValue::String("Hello".into())
        );
    }

    #[test]
    fn test_evaluate_boolean() {
        assert_eq!(eval("=TRUE").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=FALSE").unwrap(), FormulaValue::Boolean(false));
    }

    #[test]
    fn test_evaluate_arithmetic() {
        assert_eq!(eval("=1+2").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(eval("=10-3").unwrap(), FormulaValue::Number(7.0));
        assert_eq!(eval("=4*5").unwrap(), FormulaValue::Number(20.0));
        assert_eq!(eval("=20/4").unwrap(), FormulaValue::Number(5.0));
        assert_eq!(eval("=2^10").unwrap(), FormulaValue::Number(1024.0));
    }

    #[test]
    fn test_evaluate_precedence() {
        assert_eq!(eval("=1+2*3").unwrap(), FormulaValue::Number(7.0));
        assert_eq!(eval("=(1+2)*3").unwrap(), FormulaValue::Number(9.0));
        assert_eq!(eval("=2+3*4-5").unwrap(), FormulaValue::Number(9.0));
    }

    #[test]
    fn test_evaluate_unary() {
        assert_eq!(eval("=-5").unwrap(), FormulaValue::Number(-5.0));
        assert_eq!(eval("=50%").unwrap(), FormulaValue::Number(0.5));
        assert_eq!(eval("=--5").unwrap(), FormulaValue::Number(5.0));
    }

    #[test]
    fn test_evaluate_comparison() {
        assert_eq!(eval("=1<2").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=1>2").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=5=5").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=5<>5").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(eval("=5<=5").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=5>=6").unwrap(), FormulaValue::Boolean(false));
    }

    #[test]
    fn test_evaluate_concatenation() {
        assert_eq!(
            eval("=\"Hello \"&\"World\"").unwrap(),
            FormulaValue::String("Hello World".into())
        );
        assert_eq!(
            eval("=\"Value: \"&42").unwrap(),
            FormulaValue::String("Value: 42".into())
        );
    }

    #[test]
    fn test_evaluate_division_by_zero() {
        assert_eq!(eval("=1/0").unwrap(), FormulaValue::Error(CellError::Div0));
    }

    #[test]
    fn test_evaluate_error() {
        assert_eq!(
            eval("=#VALUE!").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_evaluate_sum() {
        assert_eq!(eval("=SUM(1,2,3)").unwrap(), FormulaValue::Number(6.0));
        assert_eq!(eval("=SUM(1,2,3,4,5)").unwrap(), FormulaValue::Number(15.0));
    }

    #[test]
    fn test_evaluate_average() {
        assert_eq!(eval("=AVERAGE(2,4,6)").unwrap(), FormulaValue::Number(4.0));
    }

    #[test]
    fn test_evaluate_min_max() {
        assert_eq!(eval("=MIN(5,2,8,1)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(eval("=MAX(5,2,8,1)").unwrap(), FormulaValue::Number(8.0));
    }

    #[test]
    fn test_evaluate_count() {
        assert_eq!(
            eval("=COUNT(1,2,\"a\",3)").unwrap(),
            FormulaValue::Number(3.0)
        );
    }

    #[test]
    fn test_evaluate_if() {
        assert_eq!(eval("=IF(TRUE,1,2)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(eval("=IF(FALSE,1,2)").unwrap(), FormulaValue::Number(2.0));
        assert_eq!(
            eval("=IF(1>0,\"Yes\",\"No\")").unwrap(),
            FormulaValue::String("Yes".into())
        );
    }

    #[test]
    fn test_evaluate_and_or_not() {
        assert_eq!(
            eval("=AND(TRUE,TRUE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=AND(TRUE,FALSE)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=OR(TRUE,FALSE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=OR(FALSE,FALSE)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(eval("=NOT(TRUE)").unwrap(), FormulaValue::Boolean(false));
    }

    #[test]
    fn test_evaluate_nested_functions() {
        assert_eq!(
            eval("=SUM(1,IF(TRUE,10,20),3)").unwrap(),
            FormulaValue::Number(14.0)
        );
    }

    #[test]
    fn test_evaluate_array() {
        let result = eval("={1,2,3}").unwrap();
        if let FormulaValue::Array(rows) = result {
            assert_eq!(rows.len(), 1);
            assert_eq!(rows[0].len(), 3);
            assert_eq!(rows[0][0], FormulaValue::Number(1.0));
        } else {
            panic!("Expected array");
        }
    }

    #[test]
    fn test_evaluate_complex_formula() {
        // Test a complex real-world formula
        assert_eq!(
            eval("=IF(AND(1>0,2<3),SUM(1,2,3)*2,0)").unwrap(),
            FormulaValue::Number(12.0)
        );
    }

    #[test]
    fn test_text_functions() {
        assert_eq!(eval("=LEN(\"abc\")").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(
            eval("=LEFT(\"abcdef\",2)").unwrap(),
            FormulaValue::String("ab".into())
        );
        assert_eq!(
            eval("=RIGHT(\"abcdef\",3)").unwrap(),
            FormulaValue::String("def".into())
        );
        assert_eq!(
            eval("=MID(\"abcdef\",2,3)").unwrap(),
            FormulaValue::String("bcd".into())
        );
        assert_eq!(
            eval("=LOWER(\"AbC\")").unwrap(),
            FormulaValue::String("abc".into())
        );
        assert_eq!(
            eval("=UPPER(\"AbC\")").unwrap(),
            FormulaValue::String("ABC".into())
        );
        assert_eq!(
            eval("=TRIM(\"  a   b  \" )").unwrap(),
            FormulaValue::String("a b".into())
        );
        assert_eq!(
            eval("=CONCAT(\"a\",1,TRUE)").unwrap(),
            FormulaValue::String("a1TRUE".into())
        );
        assert_eq!(
            eval("=CONCAT({\"a\",\"b\";\"c\",\"d\"})").unwrap(),
            FormulaValue::String("abcd".into())
        );
    }

    #[test]
    fn test_info_functions() {
        assert_eq!(
            eval("=ISBLANK(\"\")").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(eval("=ISNUMBER(123)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISTEXT(\"x\")").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISERROR(1/0)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=ISNA(NA())").unwrap(), FormulaValue::Boolean(true));
    }

    #[test]
    fn test_date_functions_1900_system() {
        assert_eq!(
            eval("=DATE(1900,2,29)").unwrap(),
            FormulaValue::Number(60.0)
        );
        assert_eq!(eval("=DATE(1900,3,0)").unwrap(), FormulaValue::Number(60.0));
        assert_eq!(eval("=DATE(1900,3,1)").unwrap(), FormulaValue::Number(61.0));
        assert_eq!(eval("=YEAR(60)").unwrap(), FormulaValue::Number(1900.0));
        assert_eq!(eval("=MONTH(60)").unwrap(), FormulaValue::Number(2.0));
        assert_eq!(eval("=DAY(60)").unwrap(), FormulaValue::Number(29.0));
        // Year adjustment (0..1899 => +1900)
        assert_eq!(
            eval("=YEAR(DATE(108,1,2))").unwrap(),
            FormulaValue::Number(2008.0)
        );
    }

    #[test]
    fn test_date_functions_1904_system() {
        use duke_sheets_core::Workbook;

        // Create a workbook with 1904 date system
        let mut wb = Workbook::new();
        wb.settings_mut().date_1904 = true;

        fn eval_1904(formula: &str, wb: &Workbook) -> FormulaResult<FormulaValue> {
            let ast = parse_formula(formula)?;
            let ctx = EvaluationContext::new(Some(wb), 0, 0, 0);
            evaluate(&ast, &ctx)
        }

        // In 1904 system, 1904-01-01 = serial 0
        // So DATE(1904,1,1) should be 0
        assert_eq!(
            eval_1904("=DATE(1904,1,1)", &wb).unwrap(),
            FormulaValue::Number(0.0)
        );
        // DATE(1904,1,2) should be 1
        assert_eq!(
            eval_1904("=DATE(1904,1,2)", &wb).unwrap(),
            FormulaValue::Number(1.0)
        );
        // YEAR/MONTH/DAY should work correctly
        // Serial 0 in 1904 system = 1904-01-01
        assert_eq!(
            eval_1904("=YEAR(0)", &wb).unwrap(),
            FormulaValue::Number(1904.0)
        );
        assert_eq!(
            eval_1904("=MONTH(0)", &wb).unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval_1904("=DAY(0)", &wb).unwrap(),
            FormulaValue::Number(1.0)
        );
        // Serial 365 in 1904 system = 1905-01-01 (1904 is a leap year)
        assert_eq!(
            eval_1904("=YEAR(366)", &wb).unwrap(),
            FormulaValue::Number(1905.0)
        );
    }

    #[test]
    fn test_lookup_functions() {
        assert_eq!(
            eval("=INDEX({1,2;3,4},2,1)").unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            eval("=MATCH(2,{1,2,3},0)").unwrap(),
            FormulaValue::Number(2.0)
        );
        assert_eq!(
            eval("=VLOOKUP(2,{1,\"a\";2,\"b\";3,\"c\"},2,FALSE)").unwrap(),
            FormulaValue::String("b".into())
        );
    }

    #[test]
    fn test_abs_function() {
        // Basic positive/negative
        assert_eq!(eval("=ABS(-5)").unwrap(), FormulaValue::Number(5.0));
        assert_eq!(eval("=ABS(5)").unwrap(), FormulaValue::Number(5.0));
        assert_eq!(eval("=ABS(0)").unwrap(), FormulaValue::Number(0.0));
        // Decimal values
        assert_eq!(eval("=ABS(-3.14)").unwrap(), FormulaValue::Number(3.14));
        // Nested in expression
        assert_eq!(eval("=ABS(-2)+ABS(-3)").unwrap(), FormulaValue::Number(5.0));
    }

    #[test]
    fn test_round_function() {
        // Basic rounding
        assert_eq!(eval("=ROUND(2.5, 0)").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(eval("=ROUND(2.4, 0)").unwrap(), FormulaValue::Number(2.0));
        assert_eq!(eval("=ROUND(2.49, 0)").unwrap(), FormulaValue::Number(2.0));
        // Negative numbers - round half away from zero
        assert_eq!(eval("=ROUND(-2.5, 0)").unwrap(), FormulaValue::Number(-3.0));
        assert_eq!(eval("=ROUND(-2.4, 0)").unwrap(), FormulaValue::Number(-2.0));
        // Decimal places
        assert_eq!(
            eval("=ROUND(3.14159, 2)").unwrap(),
            FormulaValue::Number(3.14)
        );
        assert_eq!(
            eval("=ROUND(3.145, 2)").unwrap(),
            FormulaValue::Number(3.15)
        );
        // Negative digits (round to left of decimal)
        assert_eq!(
            eval("=ROUND(1234.5, -2)").unwrap(),
            FormulaValue::Number(1200.0)
        );
        assert_eq!(
            eval("=ROUND(1250, -2)").unwrap(),
            FormulaValue::Number(1300.0)
        );
        assert_eq!(
            eval("=ROUND(1249, -2)").unwrap(),
            FormulaValue::Number(1200.0)
        );
        // Default to 0 digits
        assert_eq!(eval("=ROUND(2.5)").unwrap(), FormulaValue::Number(3.0));
    }

    #[test]
    fn test_mod_function() {
        // Basic positive cases
        assert_eq!(eval("=MOD(3, 2)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(eval("=MOD(10, 3)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(eval("=MOD(6, 3)").unwrap(), FormulaValue::Number(0.0));
        // Negative dividend - result same sign as divisor (Excel behavior)
        assert_eq!(eval("=MOD(-3, 2)").unwrap(), FormulaValue::Number(1.0));
        // Negative divisor - result same sign as divisor
        assert_eq!(eval("=MOD(3, -2)").unwrap(), FormulaValue::Number(-1.0));
        assert_eq!(eval("=MOD(-3, -2)").unwrap(), FormulaValue::Number(-1.0));
        // Division by zero
        assert_eq!(
            eval("=MOD(5, 0)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
    }

    #[test]
    fn test_iferror_function() {
        // Error cases - should return second argument
        assert_eq!(eval("=IFERROR(1/0, 0)").unwrap(), FormulaValue::Number(0.0));
        assert_eq!(
            eval("=IFERROR(1/0, \"Error\")").unwrap(),
            FormulaValue::String("Error".into())
        );
        // Non-error cases - should return first argument
        assert_eq!(eval("=IFERROR(5, 0)").unwrap(), FormulaValue::Number(5.0));
        assert_eq!(
            eval("=IFERROR(\"ok\", 0)").unwrap(),
            FormulaValue::String("ok".into())
        );
        // NA error
        assert_eq!(
            eval("=IFERROR(NA(), 999)").unwrap(),
            FormulaValue::Number(999.0)
        );
    }

    #[test]
    fn test_ifna_function() {
        // NA error - should return second argument
        assert_eq!(
            eval("=IFNA(NA(), 999)").unwrap(),
            FormulaValue::Number(999.0)
        );
        // Other errors - should propagate (not caught by IFNA)
        assert_eq!(
            eval("=IFNA(1/0, 0)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
        // Non-error - should return first argument
        assert_eq!(eval("=IFNA(5, 0)").unwrap(), FormulaValue::Number(5.0));
    }

    #[test]
    fn test_counta_function() {
        // Array with mixed values
        // Note: {1, "a", TRUE} has 3 non-empty values
        assert_eq!(eval("=COUNTA({1,2,3})").unwrap(), FormulaValue::Number(3.0));
        // Single values
        assert_eq!(eval("=COUNTA(5)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(
            eval("=COUNTA(\"hello\")").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(eval("=COUNTA(TRUE)").unwrap(), FormulaValue::Number(1.0));
        // Multiple arguments
        assert_eq!(eval("=COUNTA(1, 2, 3)").unwrap(), FormulaValue::Number(3.0));
    }

    #[test]
    fn test_countblank_function() {
        // For now just test with non-blank single values
        assert_eq!(eval("=COUNTBLANK(5)").unwrap(), FormulaValue::Number(0.0));
        // Empty string counts as blank
        assert_eq!(
            eval("=COUNTBLANK(\"\")").unwrap(),
            FormulaValue::Number(1.0)
        );
    }

    #[test]
    fn test_int_function() {
        // Positive numbers
        assert_eq!(eval("=INT(3.7)").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(eval("=INT(3.2)").unwrap(), FormulaValue::Number(3.0));
        // Negative numbers - floors toward negative infinity
        assert_eq!(eval("=INT(-3.7)").unwrap(), FormulaValue::Number(-4.0));
        assert_eq!(eval("=INT(-3.2)").unwrap(), FormulaValue::Number(-4.0));
        // Integers unchanged
        assert_eq!(eval("=INT(5)").unwrap(), FormulaValue::Number(5.0));
    }

    #[test]
    fn test_trunc_function() {
        // Positive numbers - truncates toward zero
        assert_eq!(eval("=TRUNC(3.7)").unwrap(), FormulaValue::Number(3.0));
        // Negative numbers - truncates toward zero (not floor!)
        assert_eq!(eval("=TRUNC(-3.7)").unwrap(), FormulaValue::Number(-3.0));
        // With decimal places
        assert_eq!(
            eval("=TRUNC(3.14159, 2)").unwrap(),
            FormulaValue::Number(3.14)
        );
        // Negative decimal places
        assert_eq!(
            eval("=TRUNC(1234, -2)").unwrap(),
            FormulaValue::Number(1200.0)
        );
    }

    #[test]
    fn test_trig_functions() {
        // SIN
        assert_eq!(eval("=SIN(0)").unwrap(), FormulaValue::Number(0.0));
        assert_approx(eval("=SIN(PI()/2)").unwrap(), 1.0);
        assert_approx(eval("=SIN(PI())").unwrap(), 0.0);

        // COS
        assert_eq!(eval("=COS(0)").unwrap(), FormulaValue::Number(1.0));
        assert_approx(eval("=COS(PI()/2)").unwrap(), 0.0);
        assert_approx(eval("=COS(PI())").unwrap(), -1.0);

        // TAN
        assert_eq!(eval("=TAN(0)").unwrap(), FormulaValue::Number(0.0));
        assert_approx(eval("=TAN(PI()/4)").unwrap(), 1.0);

        // ASIN
        assert_eq!(eval("=ASIN(0)").unwrap(), FormulaValue::Number(0.0));
        assert_approx(eval("=ASIN(1)").unwrap(), std::f64::consts::FRAC_PI_2);
        // Out of range
        assert_eq!(
            eval("=ASIN(2)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );

        // ACOS
        assert_eq!(eval("=ACOS(1)").unwrap(), FormulaValue::Number(0.0));
        assert_approx(eval("=ACOS(0)").unwrap(), std::f64::consts::FRAC_PI_2);
        // Out of range
        assert_eq!(
            eval("=ACOS(2)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );

        // ATAN
        assert_eq!(eval("=ATAN(0)").unwrap(), FormulaValue::Number(0.0));
        assert_approx(eval("=ATAN(1)").unwrap(), std::f64::consts::FRAC_PI_4);

        // ATAN2
        assert_approx(eval("=ATAN2(1,1)").unwrap(), std::f64::consts::FRAC_PI_4);
        assert_approx(eval("=ATAN2(1,0)").unwrap(), 0.0);
        assert_approx(eval("=ATAN2(0,1)").unwrap(), std::f64::consts::FRAC_PI_2);
        // Both zero
        assert_eq!(
            eval("=ATAN2(0,0)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );

        // DEGREES
        assert_approx(eval("=DEGREES(PI())").unwrap(), 180.0);
        assert_approx(eval("=DEGREES(PI()/2)").unwrap(), 90.0);

        // RADIANS
        assert_approx(eval("=RADIANS(180)").unwrap(), std::f64::consts::PI);
        assert_approx(eval("=RADIANS(90)").unwrap(), std::f64::consts::FRAC_PI_2);
    }

    #[test]
    fn test_logical_true_false_xor() {
        // TRUE and FALSE
        assert_eq!(eval("=TRUE()").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=FALSE()").unwrap(), FormulaValue::Boolean(false));

        // XOR - true if odd number of TRUE values
        assert_eq!(eval("=XOR(TRUE)").unwrap(), FormulaValue::Boolean(true));
        assert_eq!(eval("=XOR(FALSE)").unwrap(), FormulaValue::Boolean(false));
        assert_eq!(
            eval("=XOR(TRUE, TRUE)").unwrap(),
            FormulaValue::Boolean(false)
        );
        assert_eq!(
            eval("=XOR(TRUE, FALSE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=XOR(TRUE, TRUE, TRUE)").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(eval("=XOR(1, 0, 1)").unwrap(), FormulaValue::Boolean(false));
    }

    #[test]
    fn test_char_code_functions() {
        // CHAR - convert number to character
        assert_eq!(eval("=CHAR(65)").unwrap(), FormulaValue::String("A".into()));
        assert_eq!(eval("=CHAR(97)").unwrap(), FormulaValue::String("a".into()));
        assert_eq!(eval("=CHAR(49)").unwrap(), FormulaValue::String("1".into()));

        // CODE - convert character to number
        assert_eq!(eval("=CODE(\"A\")").unwrap(), FormulaValue::Number(65.0));
        assert_eq!(eval("=CODE(\"a\")").unwrap(), FormulaValue::Number(97.0));
        assert_eq!(eval("=CODE(\"ABC\")").unwrap(), FormulaValue::Number(65.0)); // First char only

        // Round trip
        assert_eq!(
            eval("=CHAR(CODE(\"Z\"))").unwrap(),
            FormulaValue::String("Z".into())
        );
    }

    #[test]
    fn test_clean_value_functions() {
        // CLEAN - removes non-printable characters
        assert_eq!(
            eval("=CLEAN(\"Hello\")").unwrap(),
            FormulaValue::String("Hello".into())
        );

        // VALUE - convert text to number
        assert_eq!(
            eval("=VALUE(\"123\")").unwrap(),
            FormulaValue::Number(123.0)
        );
        assert_eq!(
            eval("=VALUE(\"3.14\")").unwrap(),
            FormulaValue::Number(3.14)
        );
        assert_eq!(
            eval("=VALUE(\"abc\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // T - returns text, empty for non-text
        assert_eq!(
            eval("=T(\"Hello\")").unwrap(),
            FormulaValue::String("Hello".into())
        );
        assert_eq!(eval("=T(123)").unwrap(), FormulaValue::String("".into()));

        // N - returns number, 0 for non-number
        assert_eq!(eval("=N(123)").unwrap(), FormulaValue::Number(123.0));
        assert_eq!(eval("=N(TRUE)").unwrap(), FormulaValue::Number(1.0));
        assert_eq!(eval("=N(\"text\")").unwrap(), FormulaValue::Number(0.0));
    }

    #[test]
    fn test_rounding_functions() {
        // ROUNDUP - away from zero
        assert_eq!(eval("=ROUNDUP(3.2, 0)").unwrap(), FormulaValue::Number(4.0));
        assert_eq!(eval("=ROUNDUP(3.7, 0)").unwrap(), FormulaValue::Number(4.0));
        assert_eq!(
            eval("=ROUNDUP(-3.2, 0)").unwrap(),
            FormulaValue::Number(-4.0)
        );
        assert_eq!(
            eval("=ROUNDUP(3.14159, 2)").unwrap(),
            FormulaValue::Number(3.15)
        );

        // ROUNDDOWN - toward zero
        assert_eq!(
            eval("=ROUNDDOWN(3.9, 0)").unwrap(),
            FormulaValue::Number(3.0)
        );
        assert_eq!(
            eval("=ROUNDDOWN(-3.9, 0)").unwrap(),
            FormulaValue::Number(-3.0)
        );
        assert_eq!(
            eval("=ROUNDDOWN(3.14159, 2)").unwrap(),
            FormulaValue::Number(3.14)
        );

        // CEILING.MATH
        assert_eq!(
            eval("=CEILING.MATH(4.3)").unwrap(),
            FormulaValue::Number(5.0)
        );
        assert_eq!(
            eval("=CEILING.MATH(-4.3)").unwrap(),
            FormulaValue::Number(-4.0)
        );
        assert_eq!(
            eval("=CEILING.MATH(6.7, 2)").unwrap(),
            FormulaValue::Number(8.0)
        );

        // FLOOR.MATH
        assert_eq!(eval("=FLOOR.MATH(4.7)").unwrap(), FormulaValue::Number(4.0));
        assert_eq!(
            eval("=FLOOR.MATH(-4.7)").unwrap(),
            FormulaValue::Number(-5.0)
        );
        assert_eq!(
            eval("=FLOOR.MATH(7.3, 2)").unwrap(),
            FormulaValue::Number(6.0)
        );

        // ODD - round to nearest odd integer away from zero
        assert_eq!(eval("=ODD(1.5)").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(eval("=ODD(2)").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(eval("=ODD(3)").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(eval("=ODD(-1.5)").unwrap(), FormulaValue::Number(-3.0));

        // EVEN - round to nearest even integer away from zero
        assert_eq!(eval("=EVEN(1.5)").unwrap(), FormulaValue::Number(2.0));
        assert_eq!(eval("=EVEN(2)").unwrap(), FormulaValue::Number(2.0));
        assert_eq!(eval("=EVEN(3)").unwrap(), FormulaValue::Number(4.0));
        assert_eq!(eval("=EVEN(-1.5)").unwrap(), FormulaValue::Number(-2.0));
    }

    #[test]
    fn test_sqrt_function() {
        assert_eq!(eval("=SQRT(4)").unwrap(), FormulaValue::Number(2.0));
        assert_eq!(eval("=SQRT(9)").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(eval("=SQRT(0)").unwrap(), FormulaValue::Number(0.0));
        // Negative numbers return error
        assert_eq!(
            eval("=SQRT(-1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_power_function() {
        assert_eq!(eval("=POWER(2, 3)").unwrap(), FormulaValue::Number(8.0));
        assert_eq!(eval("=POWER(10, 2)").unwrap(), FormulaValue::Number(100.0));
        assert_eq!(eval("=POWER(4, 0.5)").unwrap(), FormulaValue::Number(2.0)); // Square root
        assert_eq!(eval("=POWER(2, -1)").unwrap(), FormulaValue::Number(0.5));
    }

    // Helper for approximate floating point comparison in tests
    fn assert_approx(result: FormulaValue, expected: f64) {
        if let FormulaValue::Number(n) = result {
            assert!(
                (n - expected).abs() < 1e-9,
                "Expected {} but got {}",
                expected,
                n
            );
        } else {
            panic!("Expected Number but got {:?}", result);
        }
    }

    #[test]
    fn test_log_functions() {
        // LOG with default base 10
        assert_approx(eval("=LOG(100)").unwrap(), 2.0);
        assert_approx(eval("=LOG(1000)").unwrap(), 3.0);
        // LOG with custom base
        assert_approx(eval("=LOG(8, 2)").unwrap(), 3.0);
        // LOG10
        assert_approx(eval("=LOG10(100)").unwrap(), 2.0);
        // LN (natural log) - use actual e for precise test
        assert_approx(eval("=LN(EXP(1))").unwrap(), 1.0);
        // Negative inputs return error
        assert_eq!(
            eval("=LOG(-1)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_exp_function() {
        assert_eq!(eval("=EXP(0)").unwrap(), FormulaValue::Number(1.0));
        let exp1 = eval("=EXP(1)").unwrap();
        if let FormulaValue::Number(n) = exp1 {
            assert!((n - std::f64::consts::E).abs() < 0.0001);
        }
    }

    #[test]
    fn test_pi_function() {
        let pi = eval("=PI()").unwrap();
        if let FormulaValue::Number(n) = pi {
            assert!((n - std::f64::consts::PI).abs() < 0.0000001);
        }
    }

    #[test]
    fn test_find_function() {
        // Basic find
        assert_eq!(
            eval("=FIND(\"o\", \"Hello\")").unwrap(),
            FormulaValue::Number(5.0)
        );
        assert_eq!(
            eval("=FIND(\"l\", \"Hello\")").unwrap(),
            FormulaValue::Number(3.0)
        );
        // Case-sensitive
        assert_eq!(
            eval("=FIND(\"H\", \"Hello\")").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=FIND(\"h\", \"Hello\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        // With start position
        assert_eq!(
            eval("=FIND(\"l\", \"Hello\", 4)").unwrap(),
            FormulaValue::Number(4.0)
        );
        // Not found
        assert_eq!(
            eval("=FIND(\"z\", \"Hello\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_search_function() {
        // Basic search (case-insensitive)
        assert_eq!(
            eval("=SEARCH(\"o\", \"Hello\")").unwrap(),
            FormulaValue::Number(5.0)
        );
        assert_eq!(
            eval("=SEARCH(\"H\", \"Hello\")").unwrap(),
            FormulaValue::Number(1.0)
        );
        assert_eq!(
            eval("=SEARCH(\"h\", \"Hello\")").unwrap(),
            FormulaValue::Number(1.0)
        ); // Case insensitive!
           // With start position
        assert_eq!(
            eval("=SEARCH(\"l\", \"Hello\", 4)").unwrap(),
            FormulaValue::Number(4.0)
        );
    }

    #[test]
    fn test_exact_function() {
        assert_eq!(
            eval("=EXACT(\"Hello\", \"Hello\")").unwrap(),
            FormulaValue::Boolean(true)
        );
        assert_eq!(
            eval("=EXACT(\"Hello\", \"hello\")").unwrap(),
            FormulaValue::Boolean(false)
        ); // Case sensitive
        assert_eq!(
            eval("=EXACT(\"abc\", \"abc\")").unwrap(),
            FormulaValue::Boolean(true)
        );
    }

    #[test]
    fn test_rept_function() {
        assert_eq!(
            eval("=REPT(\"ab\", 3)").unwrap(),
            FormulaValue::String("ababab".into())
        );
        assert_eq!(
            eval("=REPT(\"x\", 5)").unwrap(),
            FormulaValue::String("xxxxx".into())
        );
        assert_eq!(
            eval("=REPT(\"test\", 0)").unwrap(),
            FormulaValue::String("".into())
        );
    }

    #[test]
    fn test_substitute_function() {
        // Replace all occurrences
        assert_eq!(
            eval("=SUBSTITUTE(\"Hello World\", \"o\", \"0\")").unwrap(),
            FormulaValue::String("Hell0 W0rld".into())
        );
        // Replace specific occurrence
        assert_eq!(
            eval("=SUBSTITUTE(\"Hello World\", \"o\", \"0\", 1)").unwrap(),
            FormulaValue::String("Hell0 World".into())
        );
        assert_eq!(
            eval("=SUBSTITUTE(\"Hello World\", \"o\", \"0\", 2)").unwrap(),
            FormulaValue::String("Hello W0rld".into())
        );
    }

    #[test]
    fn test_proper_function() {
        assert_eq!(
            eval("=PROPER(\"hello world\")").unwrap(),
            FormulaValue::String("Hello World".into())
        );
        assert_eq!(
            eval("=PROPER(\"HELLO WORLD\")").unwrap(),
            FormulaValue::String("Hello World".into())
        );
        assert_eq!(
            eval("=PROPER(\"hELLO wORLD\")").unwrap(),
            FormulaValue::String("Hello World".into())
        );
    }

    #[test]
    fn test_sumif_function() {
        // Basic numeric criteria - sum values equal to 5
        assert_eq!(
            eval("=SUMIF({1,5,3,5,2}, 5)").unwrap(),
            FormulaValue::Number(10.0) // 5 + 5
        );

        // Greater than criteria
        assert_eq!(
            eval("=SUMIF({1,5,3,8,2}, \">3\")").unwrap(),
            FormulaValue::Number(13.0) // 5 + 8
        );

        // Greater than or equal
        assert_eq!(
            eval("=SUMIF({1,5,3,8,2}, \">=3\")").unwrap(),
            FormulaValue::Number(16.0) // 5 + 3 + 8
        );

        // Less than
        assert_eq!(
            eval("=SUMIF({1,5,3,8,2}, \"<3\")").unwrap(),
            FormulaValue::Number(3.0) // 1 + 2
        );

        // Not equal
        assert_eq!(
            eval("=SUMIF({1,5,3,5,2}, \"<>5\")").unwrap(),
            FormulaValue::Number(6.0) // 1 + 3 + 2
        );

        // With separate sum_range (2D arrays for range and sum)
        // Range: check column 1, sum from column 2
        assert_eq!(
            eval("=SUMIF({1;5;3}, 5, {10;20;30})").unwrap(),
            FormulaValue::Number(20.0) // Row 2 matches, sum 20
        );

        // Multiple matches with sum_range
        assert_eq!(
            eval("=SUMIF({1;5;5}, 5, {10;20;30})").unwrap(),
            FormulaValue::Number(50.0) // Rows 2,3 match, sum 20+30
        );

        // String criteria as number
        assert_eq!(
            eval("=SUMIF({1,5,3}, \"5\")").unwrap(),
            FormulaValue::Number(5.0)
        );

        // Zero sum when nothing matches
        assert_eq!(
            eval("=SUMIF({1,2,3}, 99)").unwrap(),
            FormulaValue::Number(0.0)
        );
    }

    #[test]
    fn test_countif_function() {
        // Count values equal to 5
        assert_eq!(
            eval("=COUNTIF({1,5,3,5,2}, 5)").unwrap(),
            FormulaValue::Number(2.0)
        );

        // Count values greater than 3
        assert_eq!(
            eval("=COUNTIF({1,5,3,8,2}, \">3\")").unwrap(),
            FormulaValue::Number(2.0) // 5, 8
        );

        // Count values greater than or equal to 3
        assert_eq!(
            eval("=COUNTIF({1,5,3,8,2}, \">=3\")").unwrap(),
            FormulaValue::Number(3.0) // 5, 3, 8
        );

        // Count values less than 3
        assert_eq!(
            eval("=COUNTIF({1,5,3,8,2}, \"<3\")").unwrap(),
            FormulaValue::Number(2.0) // 1, 2
        );

        // Count values not equal to 5
        assert_eq!(
            eval("=COUNTIF({1,5,3,5,2}, \"<>5\")").unwrap(),
            FormulaValue::Number(3.0) // 1, 3, 2
        );

        // No matches
        assert_eq!(
            eval("=COUNTIF({1,2,3}, 99)").unwrap(),
            FormulaValue::Number(0.0)
        );

        // String criteria as number
        assert_eq!(
            eval("=COUNTIF({1,5,3}, \"5\")").unwrap(),
            FormulaValue::Number(1.0)
        );
    }

    #[test]
    fn test_averageif_function() {
        // Average of values equal to 5
        assert_eq!(
            eval("=AVERAGEIF({5,5,5}, 5)").unwrap(),
            FormulaValue::Number(5.0)
        );

        // Average of values greater than 3
        assert_eq!(
            eval("=AVERAGEIF({1,5,3,7,2}, \">3\")").unwrap(),
            FormulaValue::Number(6.0) // (5 + 7) / 2
        );

        // With separate average_range
        // Range: check column, average from different column
        assert_eq!(
            eval("=AVERAGEIF({1;5;3}, 5, {10;20;30})").unwrap(),
            FormulaValue::Number(20.0) // Row 2 matches, avg = 20
        );

        // Multiple matches with average_range
        assert_eq!(
            eval("=AVERAGEIF({5;5;3}, 5, {10;20;30})").unwrap(),
            FormulaValue::Number(15.0) // Rows 1,2 match, avg = (10+20)/2
        );

        // No matches - returns #DIV/0!
        assert_eq!(
            eval("=AVERAGEIF({1,2,3}, 99)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
    }

    #[test]
    fn test_median_function() {
        // Odd count - middle value
        assert_eq!(eval("=MEDIAN(1, 2, 3)").unwrap(), FormulaValue::Number(2.0));
        assert_eq!(
            eval("=MEDIAN(1, 5, 3, 9, 7)").unwrap(),
            FormulaValue::Number(5.0)
        );

        // Even count - average of two middle values
        assert_eq!(
            eval("=MEDIAN(1, 2, 3, 4)").unwrap(),
            FormulaValue::Number(2.5)
        );
        assert_eq!(
            eval("=MEDIAN(1, 2, 3, 4, 5, 6)").unwrap(),
            FormulaValue::Number(3.5)
        );

        // With array
        assert_eq!(
            eval("=MEDIAN({1, 5, 3, 9, 7})").unwrap(),
            FormulaValue::Number(5.0)
        );

        // Single value
        assert_eq!(eval("=MEDIAN(42)").unwrap(), FormulaValue::Number(42.0));
    }

    #[test]
    fn test_large_function() {
        // K-th largest value
        assert_eq!(
            eval("=LARGE({1,5,3,8,2}, 1)").unwrap(),
            FormulaValue::Number(8.0) // Largest
        );
        assert_eq!(
            eval("=LARGE({1,5,3,8,2}, 2)").unwrap(),
            FormulaValue::Number(5.0) // 2nd largest
        );
        assert_eq!(
            eval("=LARGE({1,5,3,8,2}, 5)").unwrap(),
            FormulaValue::Number(1.0) // 5th largest (smallest)
        );

        // K out of range
        assert_eq!(
            eval("=LARGE({1,2,3}, 0)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            eval("=LARGE({1,2,3}, 4)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_small_function() {
        // K-th smallest value
        assert_eq!(
            eval("=SMALL({1,5,3,8,2}, 1)").unwrap(),
            FormulaValue::Number(1.0) // Smallest
        );
        assert_eq!(
            eval("=SMALL({1,5,3,8,2}, 2)").unwrap(),
            FormulaValue::Number(2.0) // 2nd smallest
        );
        assert_eq!(
            eval("=SMALL({1,5,3,8,2}, 5)").unwrap(),
            FormulaValue::Number(8.0) // 5th smallest (largest)
        );

        // K out of range
        assert_eq!(
            eval("=SMALL({1,2,3}, 0)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
        assert_eq!(
            eval("=SMALL({1,2,3}, 4)").unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_rows_columns_functions() {
        // ROWS - count rows in array
        assert_eq!(
            eval("=ROWS({1,2,3})").unwrap(),
            FormulaValue::Number(1.0) // 1 row, 3 columns
        );
        assert_eq!(
            eval("=ROWS({1;2;3})").unwrap(),
            FormulaValue::Number(3.0) // 3 rows, 1 column
        );
        assert_eq!(
            eval("=ROWS({1,2;3,4;5,6})").unwrap(),
            FormulaValue::Number(3.0) // 3 rows
        );
        // Single value = 1 row
        assert_eq!(eval("=ROWS(5)").unwrap(), FormulaValue::Number(1.0));

        // COLUMNS - count columns in array
        assert_eq!(
            eval("=COLUMNS({1,2,3})").unwrap(),
            FormulaValue::Number(3.0) // 1 row, 3 columns
        );
        assert_eq!(
            eval("=COLUMNS({1;2;3})").unwrap(),
            FormulaValue::Number(1.0) // 3 rows, 1 column
        );
        assert_eq!(
            eval("=COLUMNS({1,2;3,4;5,6})").unwrap(),
            FormulaValue::Number(2.0) // 2 columns
        );
        // Single value = 1 column
        assert_eq!(eval("=COLUMNS(5)").unwrap(), FormulaValue::Number(1.0));
    }

    #[test]
    fn test_row_column_functions() {
        // ROW() with no args - returns current row (default context is row 0, so 1-indexed = 1)
        assert_eq!(eval("=ROW()").unwrap(), FormulaValue::Number(1.0));

        // COLUMN() with no args - returns current column (default context is col 0, so 1-indexed = 1)
        assert_eq!(eval("=COLUMN()").unwrap(), FormulaValue::Number(1.0));

        // ROW with single value - returns current row context
        assert_eq!(eval("=ROW(5)").unwrap(), FormulaValue::Number(1.0));

        // COLUMN with single value - returns current column context
        assert_eq!(eval("=COLUMN(5)").unwrap(), FormulaValue::Number(1.0));

        // ROW with array - returns column vector of row numbers
        // For {1;2;3} (3 rows), returns {1;2;3} since default current_row=0 -> 1,2,3
        let result = eval("=ROW({1;2;3})").unwrap();
        match result {
            FormulaValue::Array(arr) => {
                assert_eq!(arr.len(), 3); // 3 rows
                assert_eq!(arr[0].len(), 1); // 1 column each
                assert_eq!(arr[0][0], FormulaValue::Number(1.0));
                assert_eq!(arr[1][0], FormulaValue::Number(2.0));
                assert_eq!(arr[2][0], FormulaValue::Number(3.0));
            }
            _ => panic!("Expected array result"),
        }

        // COLUMN with array - returns row vector of column numbers
        // For {1,2,3} (3 columns), returns {1,2,3}
        let result = eval("=COLUMN({1,2,3})").unwrap();
        match result {
            FormulaValue::Array(arr) => {
                assert_eq!(arr.len(), 1); // 1 row
                assert_eq!(arr[0].len(), 3); // 3 columns
                assert_eq!(arr[0][0], FormulaValue::Number(1.0));
                assert_eq!(arr[0][1], FormulaValue::Number(2.0));
                assert_eq!(arr[0][2], FormulaValue::Number(3.0));
            }
            _ => panic!("Expected array result"),
        }
    }

    #[test]
    fn test_choose_function() {
        // Basic selection
        assert_eq!(
            eval("=CHOOSE(1, \"a\", \"b\", \"c\")").unwrap(),
            FormulaValue::String("a".into())
        );
        assert_eq!(
            eval("=CHOOSE(2, \"a\", \"b\", \"c\")").unwrap(),
            FormulaValue::String("b".into())
        );
        assert_eq!(
            eval("=CHOOSE(3, \"a\", \"b\", \"c\")").unwrap(),
            FormulaValue::String("c".into())
        );

        // With numbers
        assert_eq!(
            eval("=CHOOSE(2, 10, 20, 30)").unwrap(),
            FormulaValue::Number(20.0)
        );

        // Index floored
        assert_eq!(
            eval("=CHOOSE(2.9, 10, 20, 30)").unwrap(),
            FormulaValue::Number(20.0) // 2.9 -> 2
        );

        // Out of range
        assert_eq!(
            eval("=CHOOSE(0, \"a\", \"b\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            eval("=CHOOSE(4, \"a\", \"b\", \"c\")").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_ifs_function() {
        // First TRUE wins
        assert_eq!(
            eval("=IFS(FALSE, 1, TRUE, 2, TRUE, 3)").unwrap(),
            FormulaValue::Number(2.0)
        );

        // First condition TRUE
        assert_eq!(
            eval("=IFS(TRUE, \"yes\", FALSE, \"no\")").unwrap(),
            FormulaValue::String("yes".into())
        );

        // No TRUE condition = #N/A
        assert_eq!(
            eval("=IFS(FALSE, 1, FALSE, 2)").unwrap(),
            FormulaValue::Error(CellError::Na)
        );

        // Numeric conditions (0 = false, non-zero = true)
        assert_eq!(
            eval("=IFS(0, \"zero\", 1, \"one\")").unwrap(),
            FormulaValue::String("one".into())
        );
    }

    #[test]
    fn test_switch_function() {
        // Basic matching
        assert_eq!(
            eval("=SWITCH(2, 1, \"one\", 2, \"two\", 3, \"three\")").unwrap(),
            FormulaValue::String("two".into())
        );

        // With default (odd args after expression)
        assert_eq!(
            eval("=SWITCH(99, 1, \"one\", 2, \"two\", \"default\")").unwrap(),
            FormulaValue::String("default".into())
        );

        // No match, no default = #N/A
        assert_eq!(
            eval("=SWITCH(99, 1, \"one\", 2, \"two\")").unwrap(),
            FormulaValue::Error(CellError::Na)
        );

        // String matching (case insensitive)
        assert_eq!(
            eval("=SWITCH(\"B\", \"a\", 1, \"b\", 2, \"c\", 3)").unwrap(),
            FormulaValue::Number(2.0)
        );

        // First match wins
        assert_eq!(
            eval("=SWITCH(1, 1, \"first\", 1, \"second\")").unwrap(),
            FormulaValue::String("first".into())
        );
    }

    #[test]
    fn test_sumproduct_function() {
        // Basic: multiply corresponding elements and sum
        // {1,2,3} * {4,5,6} = {4,10,18} -> sum = 32
        assert_eq!(
            eval("=SUMPRODUCT({1,2,3}, {4,5,6})").unwrap(),
            FormulaValue::Number(32.0)
        );

        // Single array: just sum
        assert_eq!(
            eval("=SUMPRODUCT({1,2,3,4})").unwrap(),
            FormulaValue::Number(10.0)
        );

        // 2D array
        // {1,2;3,4} * {5,6;7,8} = {5,12;21,32} -> sum = 70
        assert_eq!(
            eval("=SUMPRODUCT({1,2;3,4}, {5,6;7,8})").unwrap(),
            FormulaValue::Number(70.0)
        );

        // Three arrays
        // {1,2} * {3,4} * {5,6} = {15,48} -> sum = 63
        assert_eq!(
            eval("=SUMPRODUCT({1,2}, {3,4}, {5,6})").unwrap(),
            FormulaValue::Number(63.0)
        );

        // Mismatched dimensions = #VALUE!
        assert_eq!(
            eval("=SUMPRODUCT({1,2,3}, {4,5})").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_sumifs_function() {
        // SUMIFS(sum_range, criteria_range1, criteria1, ...)
        // Sum where criteria matches
        // Sum {10,20,30,40} where {1,2,1,2} = 1 -> 10+30 = 40
        assert_eq!(
            eval("=SUMIFS({10,20,30,40}, {1,2,1,2}, 1)").unwrap(),
            FormulaValue::Number(40.0)
        );

        // Sum where value > 15: {10,20,30,40} where {10,20,30,40} > 15 -> 20+30+40 = 90
        assert_eq!(
            eval("=SUMIFS({10,20,30,40}, {10,20,30,40}, \">15\")").unwrap(),
            FormulaValue::Number(90.0)
        );

        // Multiple criteria: sum where A=1 AND B>2
        // {10,20,30,40} where {1,1,2,1}=1 AND {1,3,5,4}>2 -> 20+40 = 60
        assert_eq!(
            eval("=SUMIFS({10,20,30,40}, {1,1,2,1}, 1, {1,3,5,4}, \">2\")").unwrap(),
            FormulaValue::Number(60.0)
        );
    }

    #[test]
    fn test_countifs_function() {
        // COUNTIFS(criteria_range1, criteria1, ...)
        // Count where value = 1
        assert_eq!(
            eval("=COUNTIFS({1,2,1,2,1}, 1)").unwrap(),
            FormulaValue::Number(3.0)
        );

        // Count where value > 2
        assert_eq!(
            eval("=COUNTIFS({1,2,3,4,5}, \">2\")").unwrap(),
            FormulaValue::Number(3.0) // 3, 4, 5
        );

        // Multiple criteria: count where A=1 AND B>2
        assert_eq!(
            eval("=COUNTIFS({1,1,2,1}, 1, {1,3,5,4}, \">2\")").unwrap(),
            FormulaValue::Number(2.0) // positions 2 and 4
        );
    }

    #[test]
    fn test_averageifs_function() {
        // AVERAGEIFS(avg_range, criteria_range1, criteria1, ...)
        // Average where criteria matches
        // Average {10,20,30,40} where {1,2,1,2} = 1 -> (10+30)/2 = 20
        assert_eq!(
            eval("=AVERAGEIFS({10,20,30,40}, {1,2,1,2}, 1)").unwrap(),
            FormulaValue::Number(20.0)
        );

        // No matches = #DIV/0!
        assert_eq!(
            eval("=AVERAGEIFS({10,20,30}, {1,2,3}, 99)").unwrap(),
            FormulaValue::Error(CellError::Div0)
        );

        // Multiple criteria: sum where A=1 AND B>2
        // {10,20,30,40} where {1,1,2,1}=1 AND {5,3,5,4}>2
        // Index 0: 1=1 ✓ AND 5>2 ✓ -> 10
        // Index 1: 1=1 ✓ AND 3>2 ✓ -> 20
        // Index 2: 2=1 ✗ -> excluded
        // Index 3: 1=1 ✓ AND 4>2 ✓ -> 40
        // Average of {10,20,40} = 70/3 ≈ 23.33
        let result = eval("=AVERAGEIFS({10,20,30,40}, {1,1,2,1}, 1, {5,3,5,4}, \">2\")").unwrap();
        if let FormulaValue::Number(n) = result {
            assert!((n - 23.333333333333332).abs() < 1e-10);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_named_ranges() {
        // Test named range resolution
        use duke_sheets_core::Workbook;

        let mut workbook = Workbook::new();

        // Set up some cell values
        {
            let sheet = workbook.worksheet_mut(0).unwrap();
            sheet
                .set_cell_value_at(0, 0, duke_sheets_core::CellValue::Number(100.0))
                .unwrap(); // A1
            sheet
                .set_cell_value_at(0, 1, duke_sheets_core::CellValue::Number(200.0))
                .unwrap(); // B1
            sheet
                .set_cell_value_at(1, 0, duke_sheets_core::CellValue::Number(10.0))
                .unwrap(); // A2
            sheet
                .set_cell_value_at(1, 1, duke_sheets_core::CellValue::Number(20.0))
                .unwrap(); // B2
        }

        // Define named ranges
        workbook.define_name("Price", "Sheet1!$A$1").unwrap();
        workbook.define_name("TaxRate", "0.05").unwrap(); // Constant
        workbook
            .define_name("DataRange", "Sheet1!$A$1:$B$2")
            .unwrap();

        // Create evaluation context with the workbook
        let ctx = EvaluationContext::new(Some(&workbook), 0, 0, 0);

        // Test resolving a cell reference
        let result = ctx.resolve_named_range("Price").unwrap();
        assert_eq!(result, FormulaValue::Number(100.0));

        // Test resolving a constant
        let result = ctx.resolve_named_range("TaxRate").unwrap();
        assert_eq!(result, FormulaValue::Number(0.05));

        // Test resolving a range
        let result = ctx.resolve_named_range("DataRange").unwrap();
        match result {
            FormulaValue::Array(arr) => {
                assert_eq!(arr.len(), 2); // 2 rows
                assert_eq!(arr[0].len(), 2); // 2 columns
                assert_eq!(arr[0][0], FormulaValue::Number(100.0)); // A1
                assert_eq!(arr[0][1], FormulaValue::Number(200.0)); // B1
                assert_eq!(arr[1][0], FormulaValue::Number(10.0)); // A2
                assert_eq!(arr[1][1], FormulaValue::Number(20.0)); // B2
            }
            _ => panic!("Expected array result"),
        }

        // Test unknown name returns error
        let result = ctx.resolve_named_range("UnknownName");
        assert!(result.is_err());

        // Test case-insensitive lookup
        let result = ctx.resolve_named_range("price").unwrap();
        assert_eq!(result, FormulaValue::Number(100.0));
        let result = ctx.resolve_named_range("TAXRATE").unwrap();
        assert_eq!(result, FormulaValue::Number(0.05));
    }

    #[test]
    fn test_named_range_formula() {
        // Test named range that contains a formula
        use duke_sheets_core::Workbook;

        let mut workbook = Workbook::new();

        // Set up some cell values
        {
            let sheet = workbook.worksheet_mut(0).unwrap();
            sheet
                .set_cell_value_at(0, 0, duke_sheets_core::CellValue::Number(10.0))
                .unwrap(); // A1
            sheet
                .set_cell_value_at(0, 1, duke_sheets_core::CellValue::Number(20.0))
                .unwrap(); // B1
            sheet
                .set_cell_value_at(0, 2, duke_sheets_core::CellValue::Number(30.0))
                .unwrap(); // C1
        }

        // Define a named range that contains a formula
        workbook.define_name("MySum", "=10+20+30").unwrap();

        let ctx = EvaluationContext::new(Some(&workbook), 0, 0, 0);

        // Test resolving a formula
        let result = ctx.resolve_named_range("MySum").unwrap();
        assert_eq!(result, FormulaValue::Number(60.0));
    }

    #[test]
    fn test_sequence_function() {
        // SEQUENCE(rows) - basic column of numbers
        let result = eval("=SEQUENCE(5)").unwrap();
        match result {
            FormulaValue::Array(arr) => {
                assert_eq!(arr.len(), 5); // 5 rows
                assert_eq!(arr[0].len(), 1); // 1 column
                assert_eq!(arr[0][0], FormulaValue::Number(1.0));
                assert_eq!(arr[1][0], FormulaValue::Number(2.0));
                assert_eq!(arr[2][0], FormulaValue::Number(3.0));
                assert_eq!(arr[3][0], FormulaValue::Number(4.0));
                assert_eq!(arr[4][0], FormulaValue::Number(5.0));
            }
            _ => panic!("Expected array result"),
        }

        // SEQUENCE(rows, cols) - 2D array
        let result = eval("=SEQUENCE(3, 4)").unwrap();
        match result {
            FormulaValue::Array(arr) => {
                assert_eq!(arr.len(), 3); // 3 rows
                assert_eq!(arr[0].len(), 4); // 4 columns
                                             // Row 1: 1, 2, 3, 4
                assert_eq!(arr[0][0], FormulaValue::Number(1.0));
                assert_eq!(arr[0][3], FormulaValue::Number(4.0));
                // Row 2: 5, 6, 7, 8
                assert_eq!(arr[1][0], FormulaValue::Number(5.0));
                assert_eq!(arr[1][3], FormulaValue::Number(8.0));
                // Row 3: 9, 10, 11, 12
                assert_eq!(arr[2][0], FormulaValue::Number(9.0));
                assert_eq!(arr[2][3], FormulaValue::Number(12.0));
            }
            _ => panic!("Expected array result"),
        }

        // SEQUENCE(rows, cols, start) - custom start
        let result = eval("=SEQUENCE(3, 2, 10)").unwrap();
        match result {
            FormulaValue::Array(arr) => {
                assert_eq!(arr[0][0], FormulaValue::Number(10.0));
                assert_eq!(arr[0][1], FormulaValue::Number(11.0));
                assert_eq!(arr[1][0], FormulaValue::Number(12.0));
                assert_eq!(arr[2][1], FormulaValue::Number(15.0));
            }
            _ => panic!("Expected array result"),
        }

        // SEQUENCE(rows, cols, start, step) - custom step
        let result = eval("=SEQUENCE(4, 1, 2, 3)").unwrap();
        match result {
            FormulaValue::Array(arr) => {
                // 2, 5, 8, 11 (step of 3)
                assert_eq!(arr[0][0], FormulaValue::Number(2.0));
                assert_eq!(arr[1][0], FormulaValue::Number(5.0));
                assert_eq!(arr[2][0], FormulaValue::Number(8.0));
                assert_eq!(arr[3][0], FormulaValue::Number(11.0));
            }
            _ => panic!("Expected array result"),
        }

        // SEQUENCE with negative step (countdown)
        let result = eval("=SEQUENCE(5, 1, 10, -2)").unwrap();
        match result {
            FormulaValue::Array(arr) => {
                // 10, 8, 6, 4, 2
                assert_eq!(arr[0][0], FormulaValue::Number(10.0));
                assert_eq!(arr[1][0], FormulaValue::Number(8.0));
                assert_eq!(arr[4][0], FormulaValue::Number(2.0));
            }
            _ => panic!("Expected array result"),
        }

        // Error cases
        // rows < 1
        assert_eq!(
            eval("=SEQUENCE(0)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        // cols < 1
        assert_eq!(
            eval("=SEQUENCE(5, 0)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        // negative rows
        assert_eq!(
            eval("=SEQUENCE(-1)").unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }
}
