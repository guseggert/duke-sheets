//! BIFF8 formula token decompiler.
//!
//! Decompiles the RPN token byte stream stored in FORMULA records into
//! human-readable formula text (e.g., `SUM(A1:A10)`).
//!
//! # Usage
//!
//! ```ignore
//! use duke_sheets_xls::biff::formula;
//!
//! let formula_text = formula::decompile(token_bytes, &sheet_names);
//! // formula_text might be "SUM(A1:A10)"
//! ```

pub mod decompiler;
pub mod function_table;
pub mod ptg;
pub mod token_parser;

/// Decompile BIFF8 formula token bytes into a human-readable formula string.
///
/// The `data` slice should be the raw RPN token array (`cce` bytes from the
/// FORMULA record, starting at offset 22). The `sheet_names` are used for
/// 3D references (cross-sheet refs like `Sheet2!A1`).
///
/// Returns the formula text **without** a leading `=` sign. Returns an empty
/// string if the token stream is empty or decompilation fails.
pub fn decompile(data: &[u8], sheet_names: &[String]) -> String {
    if data.is_empty() {
        return String::new();
    }
    let tokens = token_parser::parse_tokens(data);
    decompiler::decompile(&tokens, sheet_names)
}
