//! BIFF12 formula token (Ptg) byte constants and helpers.
//!
//! The Ptg opcodes are identical to BIFF8. The data sizes after each
//! token byte differ (wider row/col/index fields) but the token type
//! IDs themselves are the same.
//!
//! Token bytes 0x00-0x1F are unclassified (operators, constants, tAttr).
//! Token bytes 0x20-0x7F are classified: the base type is `byte & 0x1F`
//! (when base >= 0x20), with class bits in bits 5-6:
//!   - 0x00 = Reference (R)
//!   - 0x20 = Value (V)
//!   - 0x40 = Array (A)

pub const PTG_EXP: u8 = 0x01;
pub const PTG_TBL: u8 = 0x02;
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

pub const PTG_MISS_ARG: u8 = 0x16;
pub const PTG_STR: u8 = 0x17;
pub const PTG_ATTR: u8 = 0x19;
pub const PTG_ERR: u8 = 0x1C;
pub const PTG_BOOL: u8 = 0x1D;
pub const PTG_INT: u8 = 0x1E;
pub const PTG_NUM: u8 = 0x1F;

pub const ATTR_VOLATILE: u8 = 0x01;
pub const ATTR_IF: u8 = 0x02;
pub const ATTR_CHOOSE: u8 = 0x04;
pub const ATTR_SKIP: u8 = 0x08;
pub const ATTR_SUM: u8 = 0x10;
pub const ATTR_ASSIGN: u8 = 0x20;
pub const ATTR_SPACE: u8 = 0x40;

pub const PTG_ARRAY: u8 = 0x20;
pub const PTG_FUNC: u8 = 0x21;
pub const PTG_FUNC_VAR: u8 = 0x22;
pub const PTG_NAME: u8 = 0x23;
pub const PTG_REF: u8 = 0x24;
pub const PTG_AREA: u8 = 0x25;
pub const PTG_MEM_AREA: u8 = 0x26;
pub const PTG_MEM_ERR: u8 = 0x27;
pub const PTG_MEM_NO_MEM: u8 = 0x28;
pub const PTG_MEM_FUNC: u8 = 0x29;
pub const PTG_REF_ERR: u8 = 0x2A;
pub const PTG_AREA_ERR: u8 = 0x2B;
pub const PTG_REF_N: u8 = 0x2C;
pub const PTG_AREA_N: u8 = 0x2D;
pub const PTG_NAME_X: u8 = 0x39;
pub const PTG_REF_3D: u8 = 0x3A;
pub const PTG_AREA_3D: u8 = 0x3B;
pub const PTG_REF_ERR_3D: u8 = 0x3C;
pub const PTG_AREA_ERR_3D: u8 = 0x3D;

/// Strip the R/V/A class bits from a classified token byte.
#[inline]
pub fn base_ptg(byte: u8) -> u8 {
    if byte >= 0x20 {
        (byte & 0x1F) | 0x20
    } else {
        byte
    }
}
