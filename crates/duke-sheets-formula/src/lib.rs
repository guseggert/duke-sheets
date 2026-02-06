//! # duke-sheets-formula
//!
//! Formula parser and evaluator for duke-sheets.
//!
//! This crate provides:
//! - Formula parsing (text → AST)
//! - Formula evaluation (AST → value)
//! - Built-in Excel functions (~450)
//! - Dependency tracking for calculation chains
//!
//! ## Example
//!
//! ```rust,ignore
//! use duke_sheets_formula::{parse_formula, evaluate};
//!
//! let ast = parse_formula("=SUM(A1:A10)")?;
//! let result = evaluate(&ast, &context)?;
//! ```

pub mod ast;
pub mod dependency;
pub mod error;
pub mod evaluator;
pub mod functions;
pub mod parser;

pub use ast::{BinaryOperator, CellReference, FormulaExpr, RangeReference, UnaryOperator};
pub use error::{FormulaError, FormulaResult};
pub use evaluator::{evaluate, EvaluationContext, FormulaValue};
pub use parser::parse_formula;
