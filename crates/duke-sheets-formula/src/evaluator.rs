//! Formula evaluator
//!
//! Evaluates formula ASTs to produce values.

use crate::ast::{
    BinaryOperator, FormulaExpr, StructuredRefSpecifier, StructuredReference, UnaryOperator,
};
use crate::error::{FormulaError, FormulaResult};
use crate::functions::FunctionRegistry;
use duke_sheets_core::{CellError, CellValue, Table, Workbook, MAX_COLS, MAX_ROWS};
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
            // SpillTarget values should be resolved by looking up the source cell
            // In this simple conversion, we return Empty - proper resolution
            // happens in the worksheet's get_value methods
            CellValue::SpillTarget { .. } => FormulaValue::Empty,
            CellValue::RichText(runs) => {
                FormulaValue::String(duke_sheets_core::rich_text_to_plain(&runs))
            }
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

// Re-export ImageInfo/ImageSizing from core so existing imports keep working.
pub use duke_sheets_core::{ImageInfo, ImageSizing};

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
    /// Optional WEBSERVICE callback.
    pub web_service_fn: Option<&'a (dyn Fn(&str) -> Option<String> + Send + Sync)>,
    /// Optional RTD callback.
    pub rtd_fn: Option<&'a (dyn Fn(&str, &str, &[String]) -> Option<String> + Send + Sync)>,
    /// Optional IMAGE metadata sink.
    pub image_sink: Option<&'a (dyn Fn(usize, u32, u16, ImageInfo) + Send + Sync)>,
}

impl<'a> EvaluationContext<'a> {
    /// Create a new evaluation context
    pub fn new(workbook: Option<&'a Workbook>, sheet: usize, row: u32, col: u16) -> Self {
        Self {
            workbook,
            current_sheet: sheet,
            current_row: row,
            current_col: col,
            web_service_fn: None,
            rtd_fn: None,
            image_sink: None,
        }
    }

    /// Create a simple context without workbook (for testing)
    pub fn simple() -> Self {
        Self {
            workbook: None,
            current_sheet: 0,
            current_row: 0,
            current_col: 0,
            web_service_fn: None,
            rtd_fn: None,
            image_sink: None,
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

        let num_rows = (end_row - start_row + 1) as usize;
        let num_cols = (end_col - start_col + 1) as usize;
        let mut rows = Vec::with_capacity(num_rows);
        for row in start_row..=end_row {
            let mut cols = Vec::with_capacity(num_cols);
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

    /// Resolve a structured table reference to a formula value.
    ///
    /// Resolves references like `Table1[Revenue]`, `Table1[@Col]`,
    /// `Table1[[#Headers],[Col]]`, etc. by finding the named table
    /// across all worksheets and computing the target cell range.
    pub fn resolve_structured_ref(&self, sr: &StructuredReference) -> FormulaResult<FormulaValue> {
        let workbook = self.workbook.ok_or_else(|| {
            FormulaError::InvalidReference(
                "No workbook context for structured reference".to_string(),
            )
        })?;

        // Find the table by name across all worksheets.
        // If no table name, search the current worksheet for a table
        // that contains the current cell (unqualified ref like [Column]).
        let (table, sheet_idx) = self.find_table(workbook, sr)?;

        // Determine which column(s) to return.
        let col_idx = match &sr.column {
            Some(col_name) => {
                let idx = table
                    .columns
                    .iter()
                    .position(|c| c.name.eq_ignore_ascii_case(col_name))
                    .ok_or_else(|| {
                        FormulaError::InvalidReference(format!(
                            "Column '{}' not found in table '{}'",
                            col_name, table.name
                        ))
                    })?;
                Some(idx)
            }
            None => None,
        };

        // Compute the spreadsheet column for this table column.
        let abs_col = col_idx.map(|i| table.reference.start.col + i as u16);

        // Determine row range based on specifiers.
        let ref_start_row = table.reference.start.row;
        let ref_end_row = table.reference.end.row;
        let header_rows = table.header_row_count;
        let totals_rows = table.totals_row_count;

        // Data row boundaries (excluding header and totals).
        let data_start = ref_start_row + header_rows;
        let data_end = if totals_rows > 0 {
            ref_end_row - totals_rows
        } else {
            ref_end_row
        };

        // Determine effective specifiers. If none provided:
        // - With a column and no specifiers → #Data (implicit)
        // - No column and no specifiers → #Data (implicit)
        let has_this_row = sr.specifiers.contains(&StructuredRefSpecifier::ThisRow);
        let has_all = sr.specifiers.contains(&StructuredRefSpecifier::All);
        let has_data = sr.specifiers.contains(&StructuredRefSpecifier::Data);
        let has_headers = sr.specifiers.contains(&StructuredRefSpecifier::Headers);
        let has_totals = sr.specifiers.contains(&StructuredRefSpecifier::Totals);
        let no_specifiers = sr.specifiers.is_empty();

        // #This Row: return a single cell value from the current row.
        if has_this_row {
            let col = abs_col.ok_or_else(|| {
                FormulaError::InvalidReference("#This Row requires a column specifier".to_string())
            })?;
            // The current_row from context must be within the table data range.
            let row = self.current_row;
            return Ok(self.get_cell_value(
                self.sheet_name_for(workbook, sheet_idx).as_deref(),
                row,
                col,
            ));
        }

        // #All: entire table range including headers and totals.
        if has_all {
            return match abs_col {
                Some(col) => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    ref_start_row,
                    col,
                    ref_end_row,
                    col,
                )),
                None => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    ref_start_row,
                    table.reference.start.col,
                    ref_end_row,
                    table.reference.end.col,
                )),
            };
        }

        // #Headers: header row only.
        if has_headers && !has_data && !has_totals {
            if header_rows == 0 {
                return Ok(FormulaValue::Error(CellError::Ref));
            }
            let header_end = ref_start_row + header_rows - 1;
            return match abs_col {
                Some(col) => Ok(self.get_cell_value(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    ref_start_row,
                    col,
                )),
                None => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    ref_start_row,
                    table.reference.start.col,
                    header_end,
                    table.reference.end.col,
                )),
            };
        }

        // #Totals: totals row only.
        if has_totals && !has_data && !has_headers {
            if totals_rows == 0 {
                return Ok(FormulaValue::Error(CellError::Ref));
            }
            let totals_start = ref_end_row - totals_rows + 1;
            return match abs_col {
                Some(col) => Ok(self.get_cell_value(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    totals_start,
                    col,
                )),
                None => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    totals_start,
                    table.reference.start.col,
                    ref_end_row,
                    table.reference.end.col,
                )),
            };
        }

        // Combined specifiers: #Headers + #Data, #Data + #Totals, or
        // #Headers + #Data + #Totals (same as #All for a column).
        if has_headers && has_data {
            let end = if has_totals { ref_end_row } else { data_end };
            return match abs_col {
                Some(col) => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    ref_start_row,
                    col,
                    end,
                    col,
                )),
                None => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    ref_start_row,
                    table.reference.start.col,
                    end,
                    table.reference.end.col,
                )),
            };
        }

        if has_data && has_totals {
            return match abs_col {
                Some(col) => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    data_start,
                    col,
                    ref_end_row,
                    col,
                )),
                None => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    data_start,
                    table.reference.start.col,
                    ref_end_row,
                    table.reference.end.col,
                )),
            };
        }

        // Default: #Data (implicit when no specifiers or explicit #Data).
        // Column specified → single column data range.
        // No column → entire data range.
        if no_specifiers || has_data {
            return match abs_col {
                Some(col) => {
                    if data_start == data_end {
                        // Single data row → return as cell value.
                        Ok(self.get_cell_value(
                            self.sheet_name_for(workbook, sheet_idx).as_deref(),
                            data_start,
                            col,
                        ))
                    } else {
                        Ok(self.get_range_values(
                            self.sheet_name_for(workbook, sheet_idx).as_deref(),
                            data_start,
                            col,
                            data_end,
                            col,
                        ))
                    }
                }
                None => Ok(self.get_range_values(
                    self.sheet_name_for(workbook, sheet_idx).as_deref(),
                    data_start,
                    table.reference.start.col,
                    data_end,
                    table.reference.end.col,
                )),
            };
        }

        // Fallback — unrecognized specifier combination.
        Ok(FormulaValue::Error(CellError::Ref))
    }

    /// Find a table by name across all worksheets.
    ///
    /// If the structured reference has no table name (unqualified ref),
    /// searches the current worksheet for a table containing the current cell.
    fn find_table<'b>(
        &self,
        workbook: &'b Workbook,
        sr: &StructuredReference,
    ) -> FormulaResult<(&'b Table, usize)> {
        match &sr.table {
            Some(table_name) => {
                // Search all worksheets for the named table.
                for idx in 0..workbook.sheet_count() {
                    if let Some(ws) = workbook.worksheet(idx) {
                        if let Some(t) = ws.table_by_name(table_name) {
                            return Ok((t, idx));
                        }
                    }
                }
                Err(FormulaError::InvalidReference(format!(
                    "Table '{}' not found",
                    table_name
                )))
            }
            None => {
                // Unqualified ref: find a table on the current sheet
                // that contains the current cell.
                let ws = workbook.worksheet(self.current_sheet).ok_or_else(|| {
                    FormulaError::InvalidReference("Current worksheet not found".to_string())
                })?;
                for t in ws.tables() {
                    let r = &t.reference;
                    if self.current_row >= r.start.row
                        && self.current_row <= r.end.row
                        && self.current_col >= r.start.col
                        && self.current_col <= r.end.col
                    {
                        return Ok((t, self.current_sheet));
                    }
                }
                Err(FormulaError::InvalidReference(
                    "No table found containing current cell".to_string(),
                ))
            }
        }
    }

    /// Get the sheet name for a given worksheet index, or None if it's the
    /// current sheet (to pass to get_cell_value / get_range_values).
    fn sheet_name_for(&self, workbook: &Workbook, sheet_idx: usize) -> Option<String> {
        if sheet_idx == self.current_sheet {
            None
        } else {
            workbook
                .worksheet(sheet_idx)
                .map(|ws| ws.name().to_string())
        }
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
        FormulaExpr::Empty => Ok(FormulaValue::Empty),

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

        // Structured table references (e.g., Table1[Revenue], Table1[@Col]).
        FormulaExpr::StructuredRef(sr) => ctx.resolve_structured_ref(sr),

        // External workbook references — external books are not loaded
        FormulaExpr::ExternalRef(_) => Ok(FormulaValue::Error(CellError::Ref)),

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

fn apply_binary_op_array(
    op: BinaryOperator,
    left: &FormulaValue,
    right: &FormulaValue,
) -> FormulaResult<FormulaValue> {
    let left_arr = to_array(left);
    let right_arr = to_array(right);
    let left_is_array = matches!(left, FormulaValue::Array(_));
    let right_is_array = matches!(right, FormulaValue::Array(_));

    let rows = left_arr.len().max(right_arr.len());
    let cols = left_arr
        .iter()
        .map(|r| r.len())
        .max()
        .unwrap_or(0)
        .max(right_arr.iter().map(|r| r.len()).max().unwrap_or(0));

    let mut result = Vec::with_capacity(rows);
    for r in 0..rows {
        let mut row = Vec::with_capacity(cols);
        for c in 0..cols {
            let l = if left_is_array {
                left_arr.get(r).and_then(|row| row.get(c))
            } else {
                left_arr.first().and_then(|row| row.first())
            };
            let r_val = if right_is_array {
                right_arr.get(r).and_then(|row| row.get(c))
            } else {
                right_arr.first().and_then(|row| row.first())
            };

            let cell_result = match (l, r_val) {
                (Some(lv), Some(rv)) => apply_scalar_binary_op(op, lv, rv)
                    .unwrap_or(FormulaValue::Error(CellError::Value)),
                _ => FormulaValue::Error(CellError::Na),
            };
            row.push(cell_result);
        }
        result.push(row);
    }

    Ok(FormulaValue::Array(result))
}

fn to_array(val: &FormulaValue) -> Vec<Vec<FormulaValue>> {
    match val {
        FormulaValue::Array(rows) => rows.clone(),
        other => vec![vec![other.clone()]],
    }
}

fn apply_scalar_binary_op(
    op: BinaryOperator,
    left_val: &FormulaValue,
    right_val: &FormulaValue,
) -> FormulaResult<FormulaValue> {
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
            compare_values(left_val, right_val) == 0,
        )),
        BinaryOperator::NotEqual => Ok(FormulaValue::Boolean(
            compare_values(left_val, right_val) != 0,
        )),
        BinaryOperator::LessThan => Ok(FormulaValue::Boolean(
            compare_values(left_val, right_val) < 0,
        )),
        BinaryOperator::LessEqual => Ok(FormulaValue::Boolean(
            compare_values(left_val, right_val) <= 0,
        )),
        BinaryOperator::GreaterThan => Ok(FormulaValue::Boolean(
            compare_values(left_val, right_val) > 0,
        )),
        BinaryOperator::GreaterEqual => Ok(FormulaValue::Boolean(
            compare_values(left_val, right_val) >= 0,
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

/// Evaluate a binary operation
fn evaluate_binary_op(
    op: BinaryOperator,
    left: &FormulaExpr,
    right: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let left_val = evaluate(left, ctx)?;
    let right_val = evaluate(right, ctx)?;

    if matches!(left_val, FormulaValue::Array(_)) || matches!(right_val, FormulaValue::Array(_)) {
        return apply_binary_op_array(op, &left_val, &right_val);
    }

    apply_scalar_binary_op(op, &left_val, &right_val)
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
    // The # (SpillRange) operator must inspect the operand AST to get the
    // cell reference, not the evaluated value — handle it before evaluating.
    if op == UnaryOperator::SpillRange {
        return evaluate_spill_range(operand, ctx);
    }

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
        UnaryOperator::ImplicitIntersection => {
            // @ operator: reduce a multi-value result to a single value.
            // If the value is an array, pick the element on the same row/column
            // as the formula cell. If already scalar, pass through.
            match val {
                FormulaValue::Array(ref rows) => {
                    if rows.is_empty() {
                        return Ok(FormulaValue::Error(CellError::Value));
                    }
                    let num_rows = rows.len();
                    let num_cols = rows[0].len();
                    if num_rows == 1 && num_cols == 1 {
                        // 1x1 array — return the single element
                        return Ok(rows[0][0].clone());
                    }
                    if num_rows == 1 {
                        // Single row — select by column (use formula's column offset)
                        // For simplicity, return the first element
                        return Ok(rows[0].first().cloned().unwrap_or(FormulaValue::Empty));
                    }
                    if num_cols == 1 {
                        // Single column — select by row (use formula's row)
                        // For simplicity, return the first element
                        return Ok(rows
                            .first()
                            .and_then(|r| r.first())
                            .cloned()
                            .unwrap_or(FormulaValue::Empty));
                    }
                    // Multi-row, multi-column: return top-left element
                    Ok(rows[0][0].clone())
                }
                _ => Ok(val), // scalars pass through unchanged
            }
        }
        UnaryOperator::SpillRange => {
            unreachable!("SpillRange handled above")
        }
    }
}

/// Evaluate the # (spill range) operator.
///
/// `A1#` resolves to the full spill range of the formula in A1.
/// If A1 is not a spill source, returns just the single cell value.
fn evaluate_spill_range(
    operand: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    // Extract the cell reference from the operand AST
    let (sheet, row, col) = match operand {
        FormulaExpr::CellRef(cell_ref) => (
            cell_ref.sheet.as_deref(),
            cell_ref.address.row,
            cell_ref.address.col,
        ),
        _ => {
            // # applied to non-cell-ref: evaluate normally and return
            return evaluate(operand, ctx);
        }
    };

    let workbook = match ctx.workbook {
        Some(wb) => wb,
        None => return Ok(FormulaValue::Error(CellError::Ref)),
    };

    let sheet_idx = match sheet {
        Some(name) => match workbook.sheet_index(name) {
            Some(idx) => idx,
            None => return Ok(FormulaValue::Error(CellError::Ref)),
        },
        None => ctx.current_sheet,
    };

    let worksheet = match workbook.worksheet(sheet_idx) {
        Some(ws) => ws,
        None => return Ok(FormulaValue::Error(CellError::Ref)),
    };

    // Check if this cell is a spill source
    if let Some(spill_info) = worksheet.get_spill_info(row, col) {
        let num_rows = spill_info.rows;
        let num_cols = spill_info.cols;
        // Build an array from the spill range
        let mut result_rows = Vec::with_capacity(num_rows as usize);
        for r in 0..num_rows {
            let mut result_cols = Vec::with_capacity(num_cols as usize);
            for c in 0..num_cols {
                let val = worksheet.get_value_at(row + r, col + c);
                result_cols.push(FormulaValue::from(val));
            }
            result_rows.push(result_cols);
        }
        Ok(FormulaValue::Array(result_rows))
    } else {
        // Not a spill source — return just the single cell value (1x1)
        Ok(FormulaValue::from(worksheet.get_value_at(row, col)))
    }
}

/// Evaluate a function call
fn evaluate_function(
    name: &str,
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let registry = get_function_registry();

    // Strip Excel future function prefixes — Excel stores newer functions
    // like IFNA, IFS, SWITCH, TEXTJOIN with _xlfn. (or _xlws.) in XML.
    // Note: function names are already uppercased by the parser.
    let lookup_name = name
        .strip_prefix("_XLFN.")
        .or_else(|| name.strip_prefix("_XLWS."))
        .unwrap_or(name);

    let func = registry
        .get(lookup_name)
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

    // Special case: AREAS needs the raw AST to count union branches.
    if lookup_name == "AREAS" {
        return evaluate_areas(args, ctx);
    }

    if lookup_name == "OFFSET" {
        return evaluate_offset(args, ctx);
    }

    // Special case: FORMULATEXT needs the raw cell reference, not the evaluated value.
    if lookup_name == "FORMULATEXT" {
        return evaluate_formulatext(args, ctx);
    }
    // Evaluate arguments
    let mut evaluated_args = Vec::with_capacity(args.len());
    for arg in args {
        evaluated_args.push(evaluate(arg, ctx)?);
    }

    // Call the function
    (func.implementation)(&evaluated_args, ctx)
}

/// FORMULATEXT requires special handling: it needs the raw cell reference
/// expression to look up the formula text, not the evaluated cell value.
fn evaluate_formulatext(
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let arg = &args[0];

    // Extract cell coordinates from the reference expression
    let (sheet_name, row, col) = match arg {
        FormulaExpr::CellRef(cell_ref) => (
            cell_ref.sheet.as_deref(),
            cell_ref.address.row,
            cell_ref.address.col,
        ),
        FormulaExpr::RangeRef(range_ref) => {
            // For ranges, use the upper-left cell (per Excel docs)
            (
                range_ref.sheet.as_deref(),
                range_ref.range.start.row,
                range_ref.range.start.col,
            )
        }
        FormulaExpr::Error(e) => return Ok(FormulaValue::Error(*e)),
        // For non-reference expressions, evaluate and propagate errors or return #N/A
        _ => {
            return match evaluate(arg, ctx)? {
                FormulaValue::Error(e) => Ok(FormulaValue::Error(e)),
                _ => Ok(FormulaValue::Error(CellError::Na)),
            };
        }
    };

    let workbook = match ctx.workbook {
        Some(wb) => wb,
        None => return Ok(FormulaValue::Error(CellError::Na)),
    };

    let sheet_idx = match sheet_name {
        Some(name) => match workbook.sheet_index(name) {
            Some(idx) => idx,
            None => return Ok(FormulaValue::Error(CellError::Na)),
        },
        None => ctx.current_sheet,
    };

    let worksheet = match workbook.worksheet(sheet_idx) {
        Some(ws) => ws,
        None => return Ok(FormulaValue::Error(CellError::Na)),
    };

    match worksheet.get_formula_at(row, col) {
        Some(formula) => Ok(FormulaValue::String(formula.to_string().into())),
        None => Ok(FormulaValue::Error(CellError::Na)),
    }
}

/// AREAS needs the raw AST to count how many separate reference areas the
/// argument contains. A Union binary-op (`(A1:B2,C3:D4)`) contributes the sum
/// of its children; all other reference-like nodes count as 1 area.
fn evaluate_areas(args: &[FormulaExpr], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn count_areas(expr: &FormulaExpr, ctx: &EvaluationContext) -> Result<u32, CellError> {
        match expr {
            FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_) | FormulaExpr::NameRef(_) => {
                Ok(1)
            }
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Union,
                left,
                right,
            } => Ok(count_areas(left, ctx)? + count_areas(right, ctx)?),
            // Intersect and Range each produce a single contiguous area
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Intersect | BinaryOperator::Range,
                ..
            } => Ok(1),
            FormulaExpr::Error(e) => Err(*e),
            other => {
                // Evaluate and propagate errors; any non-error result is 1 area.
                match evaluate(other, ctx) {
                    Ok(FormulaValue::Error(e)) => Err(e),
                    Err(_) => Ok(1),
                    _ => Ok(1),
                }
            }
        }
    }

    let arg = args.first().ok_or(FormulaError::ArgumentCount {
        function: "AREAS".into(),
        expected: "1".into(),
        actual: 0,
    })?;

    match count_areas(arg, ctx) {
        Ok(count) => Ok(FormulaValue::Number(f64::from(count))),
        Err(e) => Ok(FormulaValue::Error(e)),
    }
}


fn evaluate_offset(args: &[FormulaExpr], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let to_i64_trunc = |value: &FormulaValue| value.as_number().map(|n| n.trunc() as i64);

    let (sheet_name, base_row, base_col, base_height, base_width) = match &args[0] {
        FormulaExpr::CellRef(cell_ref) => (
            cell_ref.sheet.as_deref(),
            i64::from(cell_ref.address.row),
            i64::from(cell_ref.address.col),
            1_i64,
            1_i64,
        ),
        FormulaExpr::RangeRef(range_ref) => (
            range_ref.sheet.as_deref(),
            i64::from(range_ref.range.start.row),
            i64::from(range_ref.range.start.col),
            i64::from(range_ref.range.end.row - range_ref.range.start.row + 1),
            i64::from(range_ref.range.end.col - range_ref.range.start.col + 1),
        ),
        FormulaExpr::Error(e) => return Ok(FormulaValue::Error(*e)),
        arg => {
            let mut evaluated_args = Vec::with_capacity(args.len());
            evaluated_args.push(evaluate(arg, ctx)?);
            for other_arg in &args[1..] {
                evaluated_args.push(evaluate(other_arg, ctx)?);
            }
            return crate::functions::lookup::fn_offset(&evaluated_args, ctx);
        }
    };

    let rows_offset = match evaluate(&args[1], ctx)? {
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(e)),
        value => match to_i64_trunc(&value) {
            Some(offset) => offset,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
    };
    let cols_offset = match evaluate(&args[2], ctx)? {
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(e)),
        value => match to_i64_trunc(&value) {
            Some(offset) => offset,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
    };

    let height = match args.get(3) {
        None | Some(FormulaExpr::Empty) => base_height,
        Some(arg) => match evaluate(arg, ctx)? {
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(e)),
            value => match to_i64_trunc(&value) {
                Some(height) if height >= 1 => height,
                Some(_) => return Ok(FormulaValue::Error(CellError::Ref)),
                None => return Ok(FormulaValue::Error(CellError::Value)),
            },
        },
    };
    let width = match args.get(4) {
        None | Some(FormulaExpr::Empty) => base_width,
        Some(arg) => match evaluate(arg, ctx)? {
            FormulaValue::Error(e) => return Ok(FormulaValue::Error(e)),
            value => match to_i64_trunc(&value) {
                Some(width) if width >= 1 => width,
                Some(_) => return Ok(FormulaValue::Error(CellError::Ref)),
                None => return Ok(FormulaValue::Error(CellError::Value)),
            },
        },
    };

    let start_row = match base_row.checked_add(rows_offset) {
        Some(row) if row >= 0 => row,
        _ => return Ok(FormulaValue::Error(CellError::Ref)),
    };
    let start_col = match base_col.checked_add(cols_offset) {
        Some(col) if col >= 0 => col,
        _ => return Ok(FormulaValue::Error(CellError::Ref)),
    };

    let end_row = match start_row.checked_add(height - 1) {
        Some(row) => row,
        None => return Ok(FormulaValue::Error(CellError::Ref)),
    };
    let end_col = match start_col.checked_add(width - 1) {
        Some(col) => col,
        None => return Ok(FormulaValue::Error(CellError::Ref)),
    };

    if start_row >= i64::from(MAX_ROWS)
        || start_col >= i64::from(MAX_COLS)
        || end_row >= i64::from(MAX_ROWS)
        || end_col >= i64::from(MAX_COLS)
    {
        return Ok(FormulaValue::Error(CellError::Ref));
    }

    let start_row = start_row as u32;
    let start_col = start_col as u16;

    if height == 1 && width == 1 {
        return Ok(ctx.get_cell_value(sheet_name, start_row, start_col));
    }

    Ok(ctx.get_range_values(
        sheet_name,
        start_row,
        start_col,
        end_row as u32,
        end_col as u16,
    ))
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
