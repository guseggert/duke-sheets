//! Parse BIFF8 formula token bytes into structured `ParsedToken` values.
//!
//! The input is the raw RPN byte array from a FORMULA record (the `cce` bytes
//! starting at offset 22 in the record body). The output is a `Vec<ParsedToken>`
//! that the decompiler can process via a stack machine.

use super::ptg;

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
        row: u16,
        col: u16,
        row_relative: bool,
        col_relative: bool,
    },
    Area {
        first_row: u16,
        last_row: u16,
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
        row: u16,
        col: u16,
        row_relative: bool,
        col_relative: bool,
    },
    /// 3D area reference (Phase 2).
    Area3d {
        extern_sheet_idx: u16,
        first_row: u16,
        last_row: u16,
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
        row_offset: i16,
        /// Signed column offset (relative to shared formula origin).
        col_offset: i16,
        row_relative: bool,
        col_relative: bool,
    },
    /// Relative area reference for shared formulas (tAreaN).
    AreaN {
        first_row_offset: i16,
        last_row_offset: i16,
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
        row: u16,
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

/// Parse a BIFF8 formula token byte stream into structured tokens.
///
/// The `data` slice should be exactly `cce` bytes from the FORMULA record
/// (the RPN token array). The `extra_data` slice contains any extra data
/// appended after the token stream (used by tArray for array constants).
pub fn parse_tokens(data: &[u8]) -> Vec<ParsedToken> {
    parse_tokens_with_extra(data, &[])
}

/// Parse tokens with an optional extra-data section for tArray constants.
pub fn parse_tokens_with_extra(data: &[u8], extra_data: &[u8]) -> Vec<ParsedToken> {
    let mut tokens = Vec::new();
    let mut pos = 0;
    let mut extra_pos = 0usize; // tracks position within extra_data for tArray

    while pos < data.len() {
        let raw_byte = data[pos];
        pos += 1;

        let base = ptg::base_ptg(raw_byte);

        match base {
            // ---- Binary operators (1 byte each, no data) ----
            ptg::PTG_ADD => tokens.push(ParsedToken::Add),
            ptg::PTG_SUB => tokens.push(ParsedToken::Sub),
            ptg::PTG_MUL => tokens.push(ParsedToken::Mul),
            ptg::PTG_DIV => tokens.push(ParsedToken::Div),
            ptg::PTG_POWER => tokens.push(ParsedToken::Power),
            ptg::PTG_CONCAT => tokens.push(ParsedToken::Concat),
            ptg::PTG_LT => tokens.push(ParsedToken::Lt),
            ptg::PTG_LE => tokens.push(ParsedToken::Le),
            ptg::PTG_EQ => tokens.push(ParsedToken::Eq),
            ptg::PTG_GE => tokens.push(ParsedToken::Ge),
            ptg::PTG_GT => tokens.push(ParsedToken::Gt),
            ptg::PTG_NE => tokens.push(ParsedToken::Ne),
            ptg::PTG_ISECT => tokens.push(ParsedToken::Isect),
            ptg::PTG_LIST => tokens.push(ParsedToken::List),
            ptg::PTG_RANGE => tokens.push(ParsedToken::Range),

            // ---- Unary operators ----
            ptg::PTG_UPLUS => tokens.push(ParsedToken::Uplus),
            ptg::PTG_UMINUS => tokens.push(ParsedToken::Uminus),
            ptg::PTG_PERCENT => tokens.push(ParsedToken::Percent),
            ptg::PTG_PAREN => tokens.push(ParsedToken::Paren),

            // ---- Constants ----
            ptg::PTG_MISS_ARG => tokens.push(ParsedToken::MissArg),

            ptg::PTG_STR => {
                // BIFF8 short string: 1-byte length, then flags byte, then chars
                if pos >= data.len() {
                    break;
                }
                let str_len = data[pos] as usize;
                pos += 1;
                if pos >= data.len() {
                    break;
                }
                let flags = data[pos];
                pos += 1;
                let wide = (flags & 0x01) != 0;
                let s = if wide {
                    let byte_len = str_len * 2;
                    if pos + byte_len > data.len() {
                        break;
                    }
                    let chars: Vec<u16> = data[pos..pos + byte_len]
                        .chunks_exact(2)
                        .map(|c| u16::from_le_bytes([c[0], c[1]]))
                        .collect();
                    pos += byte_len;
                    String::from_utf16_lossy(&chars)
                } else {
                    if pos + str_len > data.len() {
                        break;
                    }
                    let s = data[pos..pos + str_len]
                        .iter()
                        .map(|&b| b as char)
                        .collect::<String>();
                    pos += str_len;
                    s
                };
                tokens.push(ParsedToken::Str(s));
            }

            ptg::PTG_ERR => {
                if pos >= data.len() {
                    break;
                }
                tokens.push(ParsedToken::Err(data[pos]));
                pos += 1;
            }

            ptg::PTG_BOOL => {
                if pos >= data.len() {
                    break;
                }
                tokens.push(ParsedToken::Bool(data[pos] != 0));
                pos += 1;
            }

            ptg::PTG_INT => {
                if pos + 2 > data.len() {
                    break;
                }
                let val = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 2;
                tokens.push(ParsedToken::Int(val));
            }

            ptg::PTG_NUM => {
                if pos + 8 > data.len() {
                    break;
                }
                let val = f64::from_le_bytes(data[pos..pos + 8].try_into().unwrap());
                pos += 8;
                tokens.push(ParsedToken::Num(val));
            }

            // ---- tAttr (0x19) — sub-types ----
            ptg::PTG_ATTR => {
                if pos + 2 > data.len() {
                    break;
                }
                let flags = data[pos];
                pos += 1;
                let attr_data = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 2;

                if (flags & ptg::ATTR_SPACE) != 0 {
                    // tAttrSpace: type(1) + count(1) encoded in the 2-byte data
                    tokens.push(ParsedToken::AttrSpace {
                        space_type: (attr_data & 0xFF) as u8,
                        count: (attr_data >> 8) as u8,
                    });
                } else if (flags & ptg::ATTR_SUM) != 0 {
                    tokens.push(ParsedToken::AttrSum);
                } else if (flags & ptg::ATTR_IF) != 0 {
                    tokens.push(ParsedToken::AttrIf { offset: attr_data });
                } else if (flags & ptg::ATTR_CHOOSE) != 0 {
                    // tAttrChoose: attr_data = number of choices (nc)
                    // Followed by (nc+1) u16 jump offsets
                    let nc = attr_data as usize;
                    let jump_count = nc + 1;
                    let mut offsets = Vec::with_capacity(jump_count);
                    for _ in 0..jump_count {
                        if pos + 2 > data.len() {
                            break;
                        }
                        let off = u16::from_le_bytes([data[pos], data[pos + 1]]);
                        pos += 2;
                        offsets.push(off);
                    }
                    tokens.push(ParsedToken::AttrChoose {
                        count: attr_data,
                        offsets,
                    });
                } else if (flags & ptg::ATTR_SKIP) != 0 {
                    tokens.push(ParsedToken::AttrSkip { offset: attr_data });
                } else if (flags & ptg::ATTR_VOLATILE) != 0 {
                    tokens.push(ParsedToken::AttrVolatile);
                } else if (flags & ptg::ATTR_ASSIGN) != 0 {
                    tokens.push(ParsedToken::AttrAssign);
                } else {
                    // Unknown attr sub-type, data already consumed
                }
            }

            // ---- Cell references ----
            ptg::PTG_REF => {
                if pos + 4 > data.len() {
                    break;
                }
                let (row, col, row_rel, col_rel) = parse_ref_fields(&data[pos..]);
                pos += 4;
                tokens.push(ParsedToken::Ref {
                    row,
                    col,
                    row_relative: row_rel,
                    col_relative: col_rel,
                });
            }

            ptg::PTG_AREA => {
                if pos + 8 > data.len() {
                    break;
                }
                let (fr, lr, fc, lc, frr, fcr, lrr, lcr) = parse_area_fields(&data[pos..]);
                pos += 8;
                tokens.push(ParsedToken::Area {
                    first_row: fr,
                    last_row: lr,
                    first_col: fc,
                    last_col: lc,
                    first_row_rel: frr,
                    first_col_rel: fcr,
                    last_row_rel: lrr,
                    last_col_rel: lcr,
                });
            }

            ptg::PTG_REF_ERR => {
                // 4 bytes of deleted ref data — skip
                if pos + 4 > data.len() {
                    break;
                }
                pos += 4;
                tokens.push(ParsedToken::RefErr);
            }

            ptg::PTG_AREA_ERR => {
                // 8 bytes of deleted area data — skip
                if pos + 8 > data.len() {
                    break;
                }
                pos += 8;
                tokens.push(ParsedToken::AreaErr);
            }

            // ---- Functions ----
            ptg::PTG_FUNC => {
                if pos + 2 > data.len() {
                    break;
                }
                let func_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 2;
                tokens.push(ParsedToken::Func { func_idx });
            }

            ptg::PTG_FUNC_VAR => {
                if pos + 3 > data.len() {
                    break;
                }
                let argc = data[pos] & 0x7F; // bits 0-6 = argument count
                pos += 1;
                let func_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 2;
                tokens.push(ParsedToken::FuncVar { argc, func_idx });
            }

            // ---- Phase 2 tokens: parse data, emit stubs ----
            ptg::PTG_NAME => {
                if pos + 4 > data.len() {
                    break;
                }
                let name_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 4; // 2 bytes name_idx + 2 reserved
                tokens.push(ParsedToken::Name { name_idx });
            }

            ptg::PTG_NAME_X => {
                if pos + 6 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                let name_idx = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
                pos += 6; // 2+2+2 reserved
                tokens.push(ParsedToken::NameX {
                    extern_sheet_idx,
                    name_idx,
                });
            }

            ptg::PTG_REF_3D => {
                if pos + 6 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                let (row, col, row_rel, col_rel) = parse_ref_fields(&data[pos + 2..]);
                pos += 6;
                tokens.push(ParsedToken::Ref3d {
                    extern_sheet_idx,
                    row,
                    col,
                    row_relative: row_rel,
                    col_relative: col_rel,
                });
            }

            ptg::PTG_AREA_3D => {
                if pos + 10 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                let (fr, lr, fc, lc, frr, fcr, lrr, lcr) = parse_area_fields(&data[pos + 2..]);
                pos += 10;
                tokens.push(ParsedToken::Area3d {
                    extern_sheet_idx,
                    first_row: fr,
                    last_row: lr,
                    first_col: fc,
                    last_col: lc,
                    first_row_rel: frr,
                    first_col_rel: fcr,
                    last_row_rel: lrr,
                    last_col_rel: lcr,
                });
            }

            ptg::PTG_REF_ERR_3D => {
                if pos + 6 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 6;
                tokens.push(ParsedToken::RefErr3d { extern_sheet_idx });
            }

            ptg::PTG_AREA_ERR_3D => {
                if pos + 10 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 10;
                tokens.push(ParsedToken::AreaErr3d { extern_sheet_idx });
            }

            // ---- tExp: array/shared formula indicator ----
            ptg::PTG_EXP => {
                if pos + 4 > data.len() {
                    break;
                }
                let row = u16::from_le_bytes([data[pos], data[pos + 1]]);
                let col = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
                pos += 4;
                tokens.push(ParsedToken::Exp { row, col });
            }

            // ---- Memory tokens (Phase 3 — skip sub-expression bytes) ----
            ptg::PTG_MEM_FUNC => {
                if pos + 2 > data.len() {
                    break;
                }
                let subexpr_len = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 2;
                // Don't skip the sub-expression — it contains real tokens
                // that the decompiler needs to process. MemFunc is just a
                // hint to Excel's evaluator.
                tokens.push(ParsedToken::MemFunc { subexpr_len });
            }

            ptg::PTG_MEM_AREA | ptg::PTG_MEM_ERR | ptg::PTG_MEM_NO_MEM => {
                // 6 bytes: reserved(4) + subexpr_len(2)
                if pos + 6 > data.len() {
                    break;
                }
                let subexpr_len = u16::from_le_bytes([data[pos + 4], data[pos + 5]]);
                pos += 6;
                // Like MemFunc, sub-expression tokens follow in the stream.
                tokens.push(ParsedToken::MemFunc { subexpr_len });
            }

            // ---- Phase 3 stubs: relative refs, array, table ----
            ptg::PTG_REF_N => {
                // tRefN: 4 bytes with signed offsets for shared formulas.
                if pos + 4 > data.len() {
                    break;
                }
                let (row_off, col_off, row_rel, col_rel) = parse_refn_fields(&data[pos..]);
                pos += 4;
                tokens.push(ParsedToken::RefN {
                    row_offset: row_off,
                    col_offset: col_off,
                    row_relative: row_rel,
                    col_relative: col_rel,
                });
            }

            ptg::PTG_AREA_N => {
                // tAreaN: 8 bytes with signed offsets for shared formulas.
                if pos + 8 > data.len() {
                    break;
                }
                let (fro, lro, fco, lco, frr, fcr, lrr, lcr) = parse_arean_fields(&data[pos..]);
                pos += 8;
                tokens.push(ParsedToken::AreaN {
                    first_row_offset: fro,
                    last_row_offset: lro,
                    first_col_offset: fco,
                    last_col_offset: lco,
                    first_row_rel: frr,
                    first_col_rel: fcr,
                    last_row_rel: lrr,
                    last_col_rel: lcr,
                });
            }

            ptg::PTG_ARRAY => {
                // 7 bytes header in token stream; actual data in extra section
                if pos + 7 > data.len() {
                    break;
                }
                pos += 7;
                // Parse array constant from extra_data at current extra_pos
                let text = parse_array_constant(extra_data, &mut extra_pos);
                tokens.push(ParsedToken::Array { text });
            }

            ptg::PTG_TBL => {
                if pos + 4 > data.len() {
                    break;
                }
                pos += 4;
                tokens.push(ParsedToken::Unknown(raw_byte));
            }

            _ => {
                // Truly unknown token — can't determine size, stop parsing
                tokens.push(ParsedToken::Unknown(raw_byte));
                break;
            }
        }
    }

    tokens
}

/// Parse an array constant from the extra-data section of a FORMULA record.
///
/// Format: nc(1) + nr(2) + elements(nc*nr), where:
///   - nc = number of columns minus 1 (so actual = nc + 1)
///   - nr = number of rows minus 1 (so actual = nr + 1)
///   - Each element: type_byte + data
///     - 0x00: empty (0 extra bytes, some writers pad with 8 bytes)
///     - 0x01: f64 (8 bytes)
///     - 0x02: string (BIFF8 unicode: len(u16) + flags(1) + chars)
///     - 0x04: bool (8 bytes: val(1) + 7 padding)
///     - 0x10: error (8 bytes: code(1) + 7 padding)
///
/// Returns formatted text like `{1,2,3;4,5,6}` (commas between columns,
/// semicolons between rows).
fn parse_array_constant(extra: &[u8], epos: &mut usize) -> String {
    if *epos + 3 > extra.len() {
        return "{<?>}".to_string();
    }
    let nc = extra[*epos] as usize + 1;
    let nr = u16::from_le_bytes([extra[*epos + 1], extra[*epos + 2]]) as usize + 1;
    *epos += 3;

    let mut rows: Vec<String> = Vec::with_capacity(nr);
    for _r in 0..nr {
        let mut cols: Vec<String> = Vec::with_capacity(nc);
        for _c in 0..nc {
            if *epos >= extra.len() {
                cols.push("<?>".to_string());
                continue;
            }
            let type_byte = extra[*epos];
            *epos += 1;
            let val_str = match type_byte {
                0x00 => {
                    // Empty — some implementations write 8 padding bytes
                    if *epos + 8 <= extra.len() {
                        *epos += 8;
                    }
                    String::new()
                }
                0x01 => {
                    // IEEE 754 double
                    if *epos + 8 > extra.len() {
                        "<?>".to_string()
                    } else {
                        let val = f64::from_le_bytes(extra[*epos..*epos + 8].try_into().unwrap());
                        *epos += 8;
                        if val == val.trunc() && val.abs() < 1e15 {
                            format!("{}", val as i64)
                        } else {
                            format!("{}", val)
                        }
                    }
                }
                0x02 => {
                    // String: len(u16) + flags(1) + chars
                    if *epos + 3 > extra.len() {
                        "<?>".to_string()
                    } else {
                        let slen = u16::from_le_bytes([extra[*epos], extra[*epos + 1]]) as usize;
                        let flags = extra[*epos + 2];
                        *epos += 3;
                        let wide = (flags & 0x01) != 0;
                        let s = if wide {
                            let byte_count = slen * 2;
                            if *epos + byte_count > extra.len() {
                                *epos = extra.len();
                                "<?>".to_string()
                            } else {
                                let chars: Vec<u16> = (0..slen)
                                    .map(|i| {
                                        u16::from_le_bytes([
                                            extra[*epos + i * 2],
                                            extra[*epos + i * 2 + 1],
                                        ])
                                    })
                                    .collect();
                                *epos += byte_count;
                                String::from_utf16_lossy(&chars)
                            }
                        } else if *epos + slen > extra.len() {
                            *epos = extra.len();
                            "<?>".to_string()
                        } else {
                            let s: String = extra[*epos..*epos + slen]
                                .iter()
                                .map(|&b| b as char)
                                .collect();
                            *epos += slen;
                            s
                        };
                        format!("\"{}\"", s.replace('"', "\"\""))
                    }
                }
                0x04 => {
                    // Boolean: val(1) + 7 padding bytes
                    if *epos + 8 > extra.len() {
                        "<?>".to_string()
                    } else {
                        let val = extra[*epos] != 0;
                        *epos += 8;
                        if val { "TRUE" } else { "FALSE" }.to_string()
                    }
                }
                0x10 => {
                    // Error: code(1) + 7 padding bytes
                    if *epos + 8 > extra.len() {
                        "<?>".to_string()
                    } else {
                        let code = extra[*epos];
                        *epos += 8;
                        match code {
                            0x00 => "#NULL!",
                            0x07 => "#DIV/0!",
                            0x0F => "#VALUE!",
                            0x17 => "#REF!",
                            0x1D => "#NAME?",
                            0x24 => "#NUM!",
                            0x2A => "#N/A",
                            _ => "#UNKNOWN!",
                        }
                        .to_string()
                    }
                }
                _ => {
                    // Unknown element type — skip 8 bytes (common padding)
                    if *epos + 8 <= extra.len() {
                        *epos += 8;
                    }
                    "<?>".to_string()
                }
            };
            cols.push(val_str);
        }
        rows.push(cols.join(","));
    }
    format!("{{{}}}", rows.join(";"))
}

/// Parse a BIFF8 cell reference from 4 bytes: row(u16) + col_rw(u16).
///
/// Returns (row, col, row_relative, col_relative).
fn parse_ref_fields(data: &[u8]) -> (u16, u16, bool, bool) {
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col_rw = u16::from_le_bytes([data[2], data[3]]);
    let col = col_rw & 0x00FF; // bits 0-7
    let row_rel = (col_rw & 0x4000) != 0; // bit 14
    let col_rel = (col_rw & 0x8000) != 0; // bit 15
    (row, col, row_rel, col_rel)
}

/// Parse a tRefN (shared formula relative ref) from 4 bytes.
///
/// Row field is a signed 16-bit offset. Column field is in bits 0-7
/// of the col_rw word, sign-extended from 8 bits to i16.
///
/// Returns (row_offset, col_offset, row_relative, col_relative).
fn parse_refn_fields(data: &[u8]) -> (i16, i16, bool, bool) {
    let row_off = i16::from_le_bytes([data[0], data[1]]);
    let col_rw = u16::from_le_bytes([data[2], data[3]]);
    // Column offset is an 8-bit signed value in bits 0-7
    let col_off = (col_rw & 0x00FF) as u8 as i8 as i16;
    let row_rel = (col_rw & 0x4000) != 0;
    let col_rel = (col_rw & 0x8000) != 0;
    (row_off, col_off, row_rel, col_rel)
}

/// Parse a tAreaN (shared formula relative area) from 8 bytes.
///
/// Row fields are signed 16-bit offsets. Column fields are 8-bit
/// signed offsets in bits 0-7 of the col_rw words.
fn parse_arean_fields(data: &[u8]) -> (i16, i16, i16, i16, bool, bool, bool, bool) {
    let first_row_off = i16::from_le_bytes([data[0], data[1]]);
    let last_row_off = i16::from_le_bytes([data[2], data[3]]);
    let first_col_rw = u16::from_le_bytes([data[4], data[5]]);
    let last_col_rw = u16::from_le_bytes([data[6], data[7]]);

    let first_col_off = (first_col_rw & 0x00FF) as u8 as i8 as i16;
    let first_row_rel = (first_col_rw & 0x4000) != 0;
    let first_col_rel = (first_col_rw & 0x8000) != 0;

    let last_col_off = (last_col_rw & 0x00FF) as u8 as i8 as i16;
    let last_row_rel = (last_col_rw & 0x4000) != 0;
    let last_col_rel = (last_col_rw & 0x8000) != 0;

    (
        first_row_off,
        last_row_off,
        first_col_off,
        last_col_off,
        first_row_rel,
        first_col_rel,
        last_row_rel,
        last_col_rel,
    )
}

/// Parse a BIFF8 area reference from 8 bytes.
///
/// Returns (first_row, last_row, first_col, last_col,
///          first_row_rel, first_col_rel, last_row_rel, last_col_rel).
fn parse_area_fields(data: &[u8]) -> (u16, u16, u16, u16, bool, bool, bool, bool) {
    let first_row = u16::from_le_bytes([data[0], data[1]]);
    let last_row = u16::from_le_bytes([data[2], data[3]]);
    let first_col_rw = u16::from_le_bytes([data[4], data[5]]);
    let last_col_rw = u16::from_le_bytes([data[6], data[7]]);

    let first_col = first_col_rw & 0x00FF;
    let first_row_rel = (first_col_rw & 0x4000) != 0;
    let first_col_rel = (first_col_rw & 0x8000) != 0;

    let last_col = last_col_rw & 0x00FF;
    let last_row_rel = (last_col_rw & 0x4000) != 0;
    let last_col_rel = (last_col_rw & 0x8000) != 0;

    (
        first_row,
        last_row,
        first_col,
        last_col,
        first_row_rel,
        first_col_rel,
        last_row_rel,
        last_col_rel,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_int_constant() {
        // tInt(0x1E) + value 42 (u16 LE)
        let data = [0x1E, 0x2A, 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Int(42)]);
    }

    #[test]
    fn test_parse_num_constant() {
        // tNum(0x1F) + 3.14 as f64 LE
        let val: f64 = 3.14;
        let mut data = vec![0x1F];
        data.extend_from_slice(&val.to_le_bytes());
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Num(3.14)]);
    }

    #[test]
    fn test_parse_bool_true() {
        let data = [0x1D, 0x01];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Bool(true)]);
    }

    #[test]
    fn test_parse_bool_false() {
        let data = [0x1D, 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Bool(false)]);
    }

    #[test]
    fn test_parse_err() {
        // tErr(0x1C) + #VALUE! (0x0F)
        let data = [0x1C, 0x0F];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Err(0x0F)]);
    }

    #[test]
    fn test_parse_str_compressed() {
        // tStr: len=5, flags=0x00 (compressed), "hello"
        let data = [0x17, 0x05, 0x00, b'h', b'e', b'l', b'l', b'o'];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Str("hello".to_string())]);
    }

    #[test]
    fn test_parse_str_wide() {
        // tStr: len=2, flags=0x01 (wide), "AB"
        let data = [0x17, 0x02, 0x01, b'A', 0x00, b'B', 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Str("AB".to_string())]);
    }

    #[test]
    fn test_parse_ref() {
        // tRefV (0x44 = 0x24 + 0x20 V-class): row=0, col=0 (A1, absolute)
        let data = [0x44, 0x00, 0x00, 0x00, 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Ref {
                row: 0,
                col: 0,
                row_relative: false,
                col_relative: false,
            }]
        );
    }

    #[test]
    fn test_parse_ref_relative() {
        // tRefV: row=4, col=2, both relative (bits 14+15 set)
        // col_rw = 2 | 0x4000 | 0x8000 = 0xC002
        let data = [0x44, 0x04, 0x00, 0x02, 0xC0];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Ref {
                row: 4,
                col: 2,
                row_relative: true,
                col_relative: true,
            }]
        );
    }

    #[test]
    fn test_parse_area() {
        // tAreaV (0x45 = 0x25+0x20): A1:C3 (rows 0-2, cols 0-2)
        let data = [
            0x45, 0x00, 0x00, // first_row=0
            0x02, 0x00, // last_row=2
            0x00, 0x00, // first_col=0
            0x02, 0x00, // last_col=2
        ];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Area {
                first_row: 0,
                last_row: 2,
                first_col: 0,
                last_col: 2,
                first_row_rel: false,
                first_col_rel: false,
                last_row_rel: false,
                last_col_rel: false,
            }]
        );
    }

    #[test]
    fn test_parse_func_sum() {
        // tFuncV (0x41 = 0x21+0x20): SUM = index 4
        let data = [0x41, 0x04, 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Func { func_idx: 4 }]);
    }

    #[test]
    fn test_parse_funcvar_if() {
        // tFuncVarV (0x42 = 0x22+0x20): IF = index 1, 3 args
        let data = [0x42, 0x03, 0x01, 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::FuncVar {
                argc: 3,
                func_idx: 1,
            }]
        );
    }

    #[test]
    fn test_parse_add_two_ints() {
        // 1 + 2: tInt(1) tInt(2) tAdd
        let data = [0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Int(1), ParsedToken::Int(2), ParsedToken::Add,]
        );
    }

    #[test]
    fn test_parse_miss_arg() {
        let data = [0x16];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::MissArg]);
    }

    #[test]
    fn test_parse_attr_sum() {
        // tAttr(0x19) + flags=0x10 (SUM) + 2 bytes data
        let data = [0x19, 0x10, 0x00, 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::AttrSum]);
    }

    #[test]
    fn test_parse_attr_volatile() {
        let data = [0x19, 0x01, 0x00, 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::AttrVolatile]);
    }

    #[test]
    fn test_parse_classified_variants() {
        // tRefR (0x24), tRefV (0x44), tRefA (0x64) should all produce same Ref
        for ptg_byte in [0x24, 0x44, 0x64] {
            let data = [ptg_byte, 0x00, 0x00, 0x00, 0x00];
            let tokens = parse_tokens(&data);
            assert_eq!(
                tokens,
                vec![ParsedToken::Ref {
                    row: 0,
                    col: 0,
                    row_relative: false,
                    col_relative: false,
                }],
                "failed for ptg byte 0x{:02X}",
                ptg_byte
            );
        }
    }

    #[test]
    fn test_parse_refn_signed_offsets() {
        // tRefNV (0x4C = 0x2C + 0x20): row_off=-1 (0xFFFF), col_off=-2 (0xFE), both relative
        // col_rw = 0xFE | 0x4000 | 0x8000 = 0xC0FE
        let data = [0x4C, 0xFF, 0xFF, 0xFE, 0xC0];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::RefN {
                row_offset: -1,
                col_offset: -2,
                row_relative: true,
                col_relative: true,
            }]
        );
    }

    #[test]
    fn test_parse_refn_positive_offsets() {
        // tRefNV: row_off=3, col_off=5, row absolute, col relative
        // col_rw = 5 | 0x8000 = 0x8005
        let data = [0x4C, 0x03, 0x00, 0x05, 0x80];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::RefN {
                row_offset: 3,
                col_offset: 5,
                row_relative: false,
                col_relative: true,
            }]
        );
    }

    #[test]
    fn test_parse_arean_signed_offsets() {
        // tAreaNV (0x4D = 0x2D + 0x20):
        // first_row_off=0, last_row_off=9, first_col_off=-1, last_col_off=-1
        // first_col_rw = 0xFF | 0xC000 = 0xC0FF, last_col_rw = 0xFF | 0xC000 = 0xC0FF
        let data = [
            0x4D, 0x00, 0x00, // first_row_off = 0
            0x09, 0x00, // last_row_off = 9
            0xFF, 0xC0, // first_col: off=-1, both rel
            0xFF, 0xC0, // last_col: off=-1, both rel
        ];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::AreaN {
                first_row_offset: 0,
                last_row_offset: 9,
                first_col_offset: -1,
                last_col_offset: -1,
                first_row_rel: true,
                first_col_rel: true,
                last_row_rel: true,
                last_col_rel: true,
            }]
        );
    }

    #[test]
    fn test_parse_array_constant_simple() {
        // tArrayV (0x60 = 0x20 + 0x40): 7-byte header in token stream
        // Extra data: nc=2 (3 cols), nr=0 (1 row), three f64 values: 1.0, 2.0, 3.0
        let token_data = [0x60, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let mut extra = Vec::new();
        extra.push(2u8); // nc = 2 → 3 columns
        extra.extend_from_slice(&0u16.to_le_bytes()); // nr = 0 → 1 row
        for val in [1.0f64, 2.0, 3.0] {
            extra.push(0x01); // type = f64
            extra.extend_from_slice(&val.to_le_bytes());
        }
        let tokens = parse_tokens_with_extra(&token_data, &extra);
        assert_eq!(
            tokens,
            vec![ParsedToken::Array {
                text: "{1,2,3}".to_string()
            }]
        );
    }

    #[test]
    fn test_parse_array_constant_2d() {
        // 2x2 array: {1,2;3,4}
        let token_data = [0x60, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let mut extra = Vec::new();
        extra.push(1u8); // nc = 1 → 2 columns
        extra.extend_from_slice(&1u16.to_le_bytes()); // nr = 1 → 2 rows
        for val in [1.0f64, 2.0, 3.0, 4.0] {
            extra.push(0x01); // type = f64
            extra.extend_from_slice(&val.to_le_bytes());
        }
        let tokens = parse_tokens_with_extra(&token_data, &extra);
        assert_eq!(
            tokens,
            vec![ParsedToken::Array {
                text: "{1,2;3,4}".to_string()
            }]
        );
    }

    #[test]
    fn test_parse_array_constant_mixed_types() {
        // 1x3 array with string, bool, error: {"hello",TRUE,#N/A}
        let token_data = [0x60, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let mut extra = Vec::new();
        extra.push(2u8); // nc = 2 → 3 columns
        extra.extend_from_slice(&0u16.to_le_bytes()); // nr = 0 → 1 row

        // String "hello": type=0x02, len=5, flags=0x00 (compressed), "hello"
        extra.push(0x02);
        extra.extend_from_slice(&5u16.to_le_bytes());
        extra.push(0x00); // flags: compressed
        extra.extend_from_slice(b"hello");

        // Bool TRUE: type=0x04, val=1, + 7 padding
        extra.push(0x04);
        extra.push(0x01);
        extra.extend_from_slice(&[0; 7]);

        // Error #N/A: type=0x10, code=0x2A, + 7 padding
        extra.push(0x10);
        extra.push(0x2A);
        extra.extend_from_slice(&[0; 7]);

        let tokens = parse_tokens_with_extra(&token_data, &extra);
        assert_eq!(
            tokens,
            vec![ParsedToken::Array {
                text: "{\"hello\",TRUE,#N/A}".to_string()
            }]
        );
    }
}
