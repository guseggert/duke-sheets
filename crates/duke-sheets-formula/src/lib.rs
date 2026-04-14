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
pub mod decompile;
pub mod dependency;
pub mod error;
pub mod eval_cache;
pub mod evaluator;
pub mod functions;
pub mod parser;

pub use ast::{
    BinaryOperator, CellReference, ExternalReference, FormulaExpr, RangeReference,
    StructuredRefSpecifier, StructuredReference, UnaryOperator,
};
pub use error::{FormulaError, FormulaResult};
pub use eval_cache::{EvalCache, LookupIndex, RangeKey};
pub use evaluator::{
    evaluate, EvaluationContext, FormulaValue, ImageInfo, ImageSizing, RangeSource,
};
pub use parser::parse_formula;
