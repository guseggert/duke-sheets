//! BIFF8 formula token decompiler.
//!
//! Decompiles the RPN token byte stream stored in FORMULA records into
//! human-readable formula text (e.g., `SUM(A1:A10)`).
//!
//! # Usage
//!
//! ```ignore
//! use duke_sheets_xls::biff::formula::{self, FormulaContext};
//!
//! let ctx = FormulaContext {
//!     sheet_names: &sheet_names,
//!     extern_sheet: &extern_sheet,
//!     supbooks: &supbooks,
//!     names: &names,
//! };
//! let formula_text = formula::decompile(token_bytes, &ctx);
//! // formula_text might be "SUM(Sheet1!A1:A10)"
//! ```

pub mod decompiler;
pub mod function_table;
pub mod ptg;
pub mod token_parser;

// ---------------------------------------------------------------------------
// Formula context types — parsed from workbook globals records
// ---------------------------------------------------------------------------

/// A SUPBOOK record: describes a supporting workbook reference.
#[derive(Debug, Clone)]
pub enum SupBook {
    /// Self-reference (cch == 0x0401): refers to the current workbook.
    /// Sheet indices in EXTERNSHEET entries map to `sheet_names[]`.
    SelfRef { sheet_count: u16 },
    /// Add-in functions sentinel (ctab == 1, cch == 0x003A).
    AddIn,
    /// External workbook reference.
    External {
        /// Encoded file path (with special byte prefixes stripped).
        path: String,
        /// Sheet names in the external workbook.
        sheets: Vec<String>,
    },
}

/// An entry from the EXTERNSHEET record: maps an index to a SUPBOOK + sheet range.
#[derive(Debug, Clone, Copy)]
pub struct ExternSheetEntry {
    /// 0-based index into the `supbooks` array.
    pub sup_book_idx: u16,
    /// 0-based index of the first referenced sheet (0xFFFE = workbook-level).
    pub first_sheet: u16,
    /// 0-based index of the last referenced sheet (0xFFFE = workbook-level).
    pub last_sheet: u16,
}

/// A defined name record (NAME / Lbl).
#[derive(Debug, Clone)]
pub struct NameRecord {
    /// The defined name string (e.g., "MyRange", "Print_Area").
    pub name: String,
    /// Sheet scope: 0 = global/workbook, 1+ = sheet-scoped (1-based index).
    pub sheet_idx: u16,
    /// Whether this is a built-in name (Print_Area, _FilterDatabase, etc.).
    pub is_builtin: bool,
}

/// Built-in name indices (when the `fBuiltin` flag is set in the NAME record).
pub const BUILTIN_NAMES: &[&str] = &[
    "Consolidate_Area", // 0x00
    "Auto_Open",        // 0x01
    "Auto_Close",       // 0x02
    "Extract",          // 0x03
    "Database",         // 0x04
    "Criteria",         // 0x05
    "Print_Area",       // 0x06
    "Print_Titles",     // 0x07
    "Recorder",         // 0x08
    "Data_Form",        // 0x09
    "Auto_Activate",    // 0x0A
    "Auto_Deactivate",  // 0x0B
    "Sheet_Title",      // 0x0C
    "_FilterDatabase",  // 0x0D
];

/// Context for formula decompilation, built from workbook globals.
#[derive(Debug)]
pub struct FormulaContext {
    /// Sheet names from BOUNDSHEET records.
    pub sheet_names: Vec<String>,
    /// EXTERNSHEET index table (maps extern_sheet_idx → SUPBOOK + sheet range).
    pub extern_sheet: Vec<ExternSheetEntry>,
    /// Supporting workbook references.
    pub supbooks: Vec<SupBook>,
    /// Defined name records.
    pub names: Vec<NameRecord>,
}

impl FormulaContext {
    /// Create an empty context (no EXTERNSHEET/SUPBOOK/NAME data).
    pub fn new(sheet_names: Vec<String>) -> Self {
        Self {
            sheet_names,
            extern_sheet: Vec::new(),
            supbooks: Vec::new(),
            names: Vec::new(),
        }
    }
}

/// Decompile BIFF8 formula token bytes into a human-readable formula string.
///
/// The `data` slice should be the raw RPN token array (`cce` bytes from the
/// FORMULA record, starting at offset 22). The `ctx` provides sheet names,
/// EXTERNSHEET mappings, and defined names for resolving 3D references.
///
/// Returns the formula text **without** a leading `=` sign. Returns an empty
/// string if the token stream is empty or decompilation fails.
pub fn decompile(data: &[u8], ctx: &FormulaContext) -> String {
    if data.is_empty() {
        return String::new();
    }
    let tokens = token_parser::parse_tokens(data);
    decompiler::decompile(&tokens, ctx)
}
