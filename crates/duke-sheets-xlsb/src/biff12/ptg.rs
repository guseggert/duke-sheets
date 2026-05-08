//! BIFF12 formula token (Ptg) byte constants and helpers.
//!
//! The Ptg opcodes are identical to BIFF8. The data sizes after each
//! token byte differ (wider row/col/index fields) but the token type
//! IDs themselves are the same.
//!
//! Token bytes 0x00-0x1F are unclassified (operators, constants, tAttr).
//!
//! Token bytes >= 0x20 are classified: the byte packs a 5-bit ptg id
//! in bits 0-4, a 2-bit PtgDataType in bits 5-6, and a reserved bit
//! in bit 7. The PtgDataType encoding is per [MS-XLS] §2.5.198.16:
//!
//!   - R-class (Reference): bits[6:5] = 0b01, byte += 0x20
//!   - V-class (Value):     bits[6:5] = 0b10, byte += 0x40
//!   - A-class (Array):     bits[6:5] = 0b11, byte += 0x60
//!
//! Pitfall: the constants in this module store the R-class form. To
//! convert R→V you CANNOT use `r | 0x20`. R-class already has bit 5
//! set, so the OR is a no-op and the byte stays R-class. Excel reads
//! the unchanged byte as a reference and #VALUE!s on evaluation.
//!
//! Use `v_class(r)` (or `r + 0x20`, which works because the carry
//! from bit 5 propagates correctly) to convert R→V.

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

/// Convert an R-class classified ptg byte to its V-class (value) form.
///
/// Per [MS-XLS] §2.5.198.16 PtgDataType, the type is a 2-bit field at
/// bit positions 5-6: R = 0b01, V = 0b10, A = 0b11. The R-class form
/// of a token always has bit 5 set; V-class has bit 5 clear and bit 6
/// set.
///
/// This helper exists because `r | 0x20` does NOT convert R→V — it's
/// a no-op since R-class bytes already have bit 5 set. Use this
/// helper or the `+ 0x20` arithmetic (which carries) instead.
#[inline]
pub const fn v_class(r: u8) -> u8 {
    (r & !0x60) | 0x40
}

/// Convert an R-class classified ptg byte to its A-class (array) form.
#[inline]
pub const fn a_class(r: u8) -> u8 {
    r | 0x40
}
