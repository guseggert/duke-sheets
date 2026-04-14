//! Parsed formula token enum shared across binary format readers.
//!
//! Row fields use u32 to support both BIFF8 (65 536 rows) and
//! BIFF12/XLSB (1 048 576 rows). Column fields stay u16 (max 16 384).

/// A parsed formula token with its associated data.
#[derive(Debug, Clone, PartialEq)]
pub enum ParsedToken {
    // --- Binary operators ---
    Add,
    Sub,
    Mul,
    Div,
    Power,
    Concat,
    Lt,
    Le,
    Eq,
    Ge,
    Gt,
    Ne,
    Isect, // intersection (space)
    List,  // union (comma)
    Range, // colon

    // --- Unary operators ---
    Uplus,
    Uminus,
    Percent,
    Paren,

    // --- Constants ---
    MissArg,
    Str(String),
    Err(u8), // error code byte
    Bool(bool),
    Int(u16),
    Num(f64),

    // --- Cell references ---
    Ref {
        row: u32,
        col: u16,
        row_relative: bool,
        col_relative: bool,
    },
    Area {
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
        first_row_rel: bool,
        first_col_rel: bool,
        last_row_rel: bool,
        last_col_rel: bool,
    },
    RefErr,
    AreaErr,

    // --- Functions ---
    Func {
        /// Function index in the BIFF8 function table.
        func_idx: u16,
    },
    FuncVar {
        /// Number of arguments actually passed.
        argc: u8,
        /// Function index (bits 0-14). Bit 15 = CE (command-equivalent) flag.
        func_idx: u16,
    },

    // --- tAttr sub-types ---
    AttrVolatile,
    AttrIf {
        offset: u16,
    },
    AttrChoose {
        count: u16,
        offsets: Vec<u16>,
    },
    AttrSkip {
        offset: u16,
    },
    AttrSum,
    AttrAssign,
    AttrSpace {
        space_type: u8,
        count: u8,
    },

    // --- Phase 2/3 stubs (skip data, emit placeholder) ---
    /// Named range reference (Phase 2).
    Name {
        name_idx: u16,
    },
    /// External name reference (Phase 2).
    NameX {
        extern_sheet_idx: u16,
        name_idx: u16,
    },
    /// 3D cell reference (Phase 2).
    Ref3d {
        extern_sheet_idx: u16,
        row: u32,
        col: u16,
        row_relative: bool,
        col_relative: bool,
    },
    /// 3D area reference (Phase 2).
    Area3d {
        extern_sheet_idx: u16,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
        first_row_rel: bool,
        first_col_rel: bool,
        last_row_rel: bool,
        last_col_rel: bool,
    },
    /// Deleted 3D ref (Phase 2).
    RefErr3d {
        extern_sheet_idx: u16,
    },
    /// Deleted 3D area (Phase 2).
    AreaErr3d {
        extern_sheet_idx: u16,
    },
    /// Relative cell reference for shared formulas (tRefN).
    /// Offsets are signed when used with a base cell.
    RefN {
        /// Signed row offset (relative to shared formula origin).
        row_offset: i32,
        /// Signed column offset (relative to shared formula origin).
        col_offset: i16,
        row_relative: bool,
        col_relative: bool,
    },
    /// Relative area reference for shared formulas (tAreaN).
    AreaN {
        first_row_offset: i32,
        last_row_offset: i32,
        first_col_offset: i16,
        last_col_offset: i16,
        first_row_rel: bool,
        first_col_rel: bool,
        last_row_rel: bool,
        last_col_rel: bool,
    },

    /// Array constant (tArray) with pre-formatted text like `{1,2,3;4,5,6}`.
    Array {
        text: String,
    },

    /// Array/shared formula indicator (tExp).
    Exp {
        row: u32,
        col: u16,
    },
    Table {
        row: u32,
        col: u16,
    },
    /// Memory function — the decompiler treats this as a no-op; the
    /// sub-expression tokens that follow produce the actual reference.
    MemFunc {
        subexpr_len: u16,
    },

    /// Unknown token — skipped. Carries original byte for debugging.
    Unknown(u8),
}
