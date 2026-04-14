//! Shared formula decompiler infrastructure.
//!
//! Converts parsed formula tokens (RPN) into human-readable infix formula
//! strings. This module is format-agnostic — it works with `ParsedToken`
//! values produced by any binary parser (BIFF8, BIFF12, etc.).

pub mod decompiler;
pub mod function_table;
pub mod parsed_token;

pub use decompiler::decompile;
pub use function_table::{function_argc, function_index, function_name};
pub use parsed_token::ParsedToken;

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
    /// Sheet scope: 0xFFFFFFFF = workbook, 0+ = sheet-scoped (0-based index).
    pub sheet_idx: u32,
    /// Whether this is a built-in name (Print_Area, _FilterDatabase, etc.).
    pub is_builtin: bool,
    /// Raw formula token bytes from the NAME record body.
    pub formula_body: Vec<u8>,
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
    /// Base cell position for shared formula offset resolution.
    /// When set, tRefN/tAreaN offsets are adjusted relative to this cell.
    pub base_cell: Option<(u32, u16)>,
}

impl FormulaContext {
    /// Create an empty context (no EXTERNSHEET/SUPBOOK/NAME data).
    pub fn new(sheet_names: Vec<String>) -> Self {
        Self {
            sheet_names,
            extern_sheet: Vec::new(),
            supbooks: Vec::new(),
            names: Vec::new(),
            base_cell: None,
        }
    }

    /// Set the base cell for shared formula offset resolution.
    pub fn set_base_cell(&mut self, row: u32, col: u16) {
        self.base_cell = Some((row, col));
    }

    /// Clear the base cell.
    pub fn clear_base_cell(&mut self) {
        self.base_cell = None;
    }
}
