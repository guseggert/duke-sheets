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

/// Unary operators
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOperator {
    Negate,
    Percent,
}
