//! BIFF8 formula token (Ptg) byte constants and helpers.
//!
//! Token bytes 0x00–0x1F are unclassified (operators, constants, tAttr).
//! Token bytes 0x20–0x7F are classified: the base type is `byte & 0x1F`
//! (when base >= 0x20), with class bits in bits 5-6:
//!   - 0x00 = Reference (R)
//!   - 0x20 = Value (V)
//!   - 0x40 = Array (A)
//!
//! For decompilation the class does not affect the output string.

// ---------------------------------------------------------------------------
// Unclassified operators (0x01–0x15) — all 1 byte, no data
// ---------------------------------------------------------------------------
pub const PTG_EXP: u8 = 0x01; // Array/shared formula indicator
pub const PTG_TBL: u8 = 0x02; // Data table indicator
pub const PTG_ADD: u8 = 0x03;
pub const PTG_SUB: u8 = 0x04;
pub const PTG_MUL: u8 = 0x05;
pub const PTG_DIV: u8 = 0x06;
pub const PTG_POWER: u8 = 0x07;
pub const PTG_CONCAT: u8 = 0x08;
pub const PTG_LT: u8 = 0x09;
pub const PTG_LE: u8 = 0x0A;
pub const PTG_EQ: u8 = 0x0B;
pub const PTG_GE: u8 = 0x0C;
pub const PTG_GT: u8 = 0x0D;
pub const PTG_NE: u8 = 0x0E;
pub const PTG_ISECT: u8 = 0x0F;
pub const PTG_LIST: u8 = 0x10;
pub const PTG_RANGE: u8 = 0x11;
pub const PTG_UPLUS: u8 = 0x12;
pub const PTG_UMINUS: u8 = 0x13;
pub const PTG_PERCENT: u8 = 0x14;
pub const PTG_PAREN: u8 = 0x15;

// ---------------------------------------------------------------------------
// Unclassified constants (0x16–0x1F)
// ---------------------------------------------------------------------------
pub const PTG_MISS_ARG: u8 = 0x16; // 1 byte (no data)
pub const PTG_STR: u8 = 0x17; // variable: 1-byte len + BIFF8 string
pub const PTG_ATTR: u8 = 0x19; // variable (sub-types below)
pub const PTG_ERR: u8 = 0x1C; // 2 bytes: error code
pub const PTG_BOOL: u8 = 0x1D; // 2 bytes: 0 or 1
pub const PTG_INT: u8 = 0x1E; // 3 bytes: u16 value
pub const PTG_NUM: u8 = 0x1F; // 9 bytes: f64 value

// ---------------------------------------------------------------------------
// tAttr sub-type flags (byte at offset +1 after 0x19)
// ---------------------------------------------------------------------------
pub const ATTR_VOLATILE: u8 = 0x01;
pub const ATTR_IF: u8 = 0x02;
pub const ATTR_CHOOSE: u8 = 0x04;
pub const ATTR_SKIP: u8 = 0x08;
pub const ATTR_SUM: u8 = 0x10;
pub const ATTR_ASSIGN: u8 = 0x20;
pub const ATTR_SPACE: u8 = 0x40;

// ---------------------------------------------------------------------------
// Classified operand tokens — base values (before R/V/A class offset)
// Strip class with: base = byte & 0x1F  (only when byte >= 0x20)
//
// The actual byte in the stream is base + 0x00 (R), + 0x20 (V), or + 0x40 (A).
// ---------------------------------------------------------------------------
pub const PTG_ARRAY: u8 = 0x20; // 7 bytes data (extra data after token stream)
pub const PTG_FUNC: u8 = 0x21; // 2 bytes: function index (u16)
pub const PTG_FUNC_VAR: u8 = 0x22; // 3 bytes: argc(u8) + func_idx(u16)
pub const PTG_NAME: u8 = 0x23; // 4 bytes: name_idx(u16) + 2 reserved
pub const PTG_REF: u8 = 0x24; // 4 bytes: row(u16) + col_rw(u16)
pub const PTG_AREA: u8 = 0x25; // 8 bytes: row1, row2, col_rw1, col_rw2
pub const PTG_MEM_AREA: u8 = 0x26; // 6 bytes: reserved(4) + subexpr_len(2)
pub const PTG_MEM_ERR: u8 = 0x27; // 6 bytes
pub const PTG_MEM_NO_MEM: u8 = 0x28; // 6 bytes
pub const PTG_MEM_FUNC: u8 = 0x29; // 2 bytes: subexpr_len(2)
pub const PTG_REF_ERR: u8 = 0x2A; // 4 bytes (deleted ref)
pub const PTG_AREA_ERR: u8 = 0x2B; // 8 bytes (deleted area)
pub const PTG_REF_N: u8 = 0x2C; // 4 bytes (shared formula relative ref)
pub const PTG_AREA_N: u8 = 0x2D; // 8 bytes (shared formula relative area)
pub const PTG_NAME_X: u8 = 0x39; // 6 bytes: extern_sheet(u16) + name_idx(u16) + 2 reserved
pub const PTG_REF_3D: u8 = 0x3A; // 6 bytes: extern_sheet(u16) + ref data(4)
pub const PTG_AREA_3D: u8 = 0x3B; // 10 bytes: extern_sheet(u16) + area data(8)
pub const PTG_REF_ERR_3D: u8 = 0x3C; // 6 bytes
pub const PTG_AREA_ERR_3D: u8 = 0x3D; // 10 bytes

/// Strip the R/V/A class bits from a classified token byte.
///
/// For bytes >= 0x20, returns the base token type (0x20–0x3F range).
/// For bytes < 0x20 (unclassified), returns the byte unchanged.
#[inline]
pub fn base_ptg(byte: u8) -> u8 {
    if byte >= 0x20 {
        // Classified token: base is in bits 0-4, class in bits 5-6.
        // But the base range is 0x20..0x3F, so we mask with 0x1F and
        // add 0x20 back to keep it in the classified range.
        (byte & 0x1F) | 0x20
    } else {
        byte
    }
}

/// Return the data size (in bytes) that follows a given base token byte.
/// Returns `None` for variable-length tokens (tStr, tAttr) that need
/// special handling, or for completely unknown tokens.
pub fn token_data_size(base: u8) -> Option<usize> {
    match base {
        // Unclassified operators — no data bytes
        PTG_ADD | PTG_SUB | PTG_MUL | PTG_DIV | PTG_POWER | PTG_CONCAT | PTG_LT | PTG_LE
        | PTG_EQ | PTG_GE | PTG_GT | PTG_NE | PTG_ISECT | PTG_LIST | PTG_RANGE | PTG_UPLUS
        | PTG_UMINUS | PTG_PERCENT | PTG_PAREN => Some(0),

        // Constants
        PTG_MISS_ARG => Some(0),
        PTG_STR => None,  // variable: short string
        PTG_ATTR => None, // variable: depends on sub-type
        PTG_ERR => Some(1),
        PTG_BOOL => Some(1),
        PTG_INT => Some(2),
        PTG_NUM => Some(8),

        // Classified operands
        PTG_ARRAY => Some(7),
        PTG_FUNC => Some(2),
        PTG_FUNC_VAR => Some(3),
        PTG_NAME => Some(4),
        PTG_REF => Some(4),
        PTG_AREA => Some(8),
        PTG_MEM_AREA => Some(6),
        PTG_MEM_ERR => Some(6),
        PTG_MEM_NO_MEM => Some(6),
        PTG_MEM_FUNC => Some(2),
        PTG_REF_ERR => Some(4),
        PTG_AREA_ERR => Some(8),
        PTG_REF_N => Some(4),
        PTG_AREA_N => Some(8),
        PTG_NAME_X => Some(6),
        PTG_REF_3D => Some(6),
        PTG_AREA_3D => Some(10),
        PTG_REF_ERR_3D => Some(6),
        PTG_AREA_ERR_3D => Some(10),

        // tExp / tTbl
        PTG_EXP => Some(4),
        PTG_TBL => Some(4),

        _ => None,
    }
}
