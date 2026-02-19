//! Formula Abstract Syntax Tree types

use duke_sheets_core::{CellAddress, CellError, CellRange};

/// Formula expression AST
#[derive(Debug, Clone, PartialEq)]
pub enum FormulaExpr {
    // === Literals ===
    /// Numeric literal
    Number(f64),
    /// String literal
    String(String),
    /// Boolean literal
    Boolean(bool),
    /// Error literal
    Error(CellError),

    // === References ===
    /// Single cell reference
    CellRef(CellReference),
    /// Range reference
    RangeRef(RangeReference),
    /// Named range or defined name
    NameRef(String),
    /// Structured table reference (e.g., Table1[Column1], Table1[[#Headers],[Col]])
    StructuredRef(StructuredReference),
    /// External workbook reference (e.g., [Book.xlsx]Sheet1!A1)
    ExternalRef(ExternalReference),

    // === Operators ===
    /// Binary operation
    BinaryOp {
        op: BinaryOperator,
        left: Box<FormulaExpr>,
        right: Box<FormulaExpr>,
    },
    /// Unary operation
    UnaryOp {
        op: UnaryOperator,
        operand: Box<FormulaExpr>,
    },

    // === Function call ===
    Function {
        name: String,
        args: Vec<FormulaExpr>,
    },

    // === Array ===
    Array(Vec<Vec<FormulaExpr>>),
}

/// Cell reference with optional sheet
#[derive(Debug, Clone, PartialEq)]
pub struct CellReference {
    pub sheet: Option<String>,
    pub address: CellAddress,
}

/// Range reference with optional sheet
#[derive(Debug, Clone, PartialEq)]
pub struct RangeReference {
    pub sheet: Option<String>,
    pub range: CellRange,
}

/// Binary operators
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOperator {
    // Arithmetic
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,

    // Comparison
    Equal,
    NotEqual,
    LessThan,
    LessEqual,
    GreaterThan,
    GreaterEqual,

    // Text
    Concat,

    // Range
    Range,
    Union,
    Intersect,
}

/// Structured table reference
#[derive(Debug, Clone, PartialEq)]
pub struct StructuredReference {
    /// Table name (e.g., "Table1"). None for unqualified refs like [Column]
    pub table: Option<String>,
    /// Column name (e.g., "Column1"). None when only specifiers are used
    pub column: Option<String>,
    /// Special item specifiers (#All, #Data, #Headers, #Totals, #This Row)
    pub specifiers: Vec<StructuredRefSpecifier>,
}

/// Structured reference specifier keywords
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StructuredRefSpecifier {
    All,
    Data,
    Headers,
    Totals,
    ThisRow,
}

/// External workbook reference
#[derive(Debug, Clone, PartialEq)]
pub struct ExternalReference {
    /// Workbook filename (e.g., "Book1.xlsx")
    pub book: String,
    /// Sheet name (optional, e.g., "Sheet1")
    pub sheet: Option<String>,
    /// Cell address
    pub address: CellAddress,
}

/// Unary operators
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOperator {
    Negate,
    Percent,
    /// Implicit intersection operator (@) — selects a single value from a range
    ImplicitIntersection,
    /// Spill range operator (#) — references the entire spill range of a cell
    SpillRange,
}
