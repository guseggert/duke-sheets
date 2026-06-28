//! Parse BIFF12 formula token bytes into structured `ParsedToken` values.
//!
//! BIFF12 tokens use the same Ptg opcodes as BIFF8 but with wider row fields:
//! rows are u32, columns are 14-bit in a u16 word.

use super::ptg;
pub use duke_sheets_formula::decompile::ParsedToken;

pub fn parse_tokens(data: &[u8]) -> Vec<ParsedToken> {
    parse_tokens_with_extra(data, &[])
}

pub fn parse_tokens_with_extra(data: &[u8], extra_data: &[u8]) -> Vec<ParsedToken> {
    let mut tokens = Vec::new();
    let mut pos = 0;
    let mut extra_pos = 0usize;

    while pos < data.len() {
        let raw_byte = data[pos];
        pos += 1;

        let base = ptg::base_ptg(raw_byte);

        match base {
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

            ptg::PTG_UPLUS => tokens.push(ParsedToken::Uplus),
            ptg::PTG_UMINUS => tokens.push(ParsedToken::Uminus),
            ptg::PTG_PERCENT => tokens.push(ParsedToken::Percent),
            ptg::PTG_PAREN => tokens.push(ParsedToken::Paren),

            ptg::PTG_MISS_ARG => tokens.push(ParsedToken::MissArg),

            ptg::PTG_STR => {
                // BIFF12 tStr: ShortXLUnicodeString (u16 char count + UTF-16LE)
                if pos + 2 > data.len() {
                    break;
                }
                let char_count = u16::from_le_bytes([data[pos], data[pos + 1]]) as usize;
                pos += 2;
                let byte_len = char_count * 2;
                if pos + byte_len > data.len() {
                    break;
                }
                let chars: Vec<u16> = data[pos..pos + byte_len]
                    .chunks_exact(2)
                    .map(|c| u16::from_le_bytes([c[0], c[1]]))
                    .collect();
                pos += byte_len;
                tokens.push(ParsedToken::Str(String::from_utf16_lossy(&chars)));
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

            ptg::PTG_ATTR => {
                if pos + 3 > data.len() {
                    break;
                }
                let flags = data[pos];
                pos += 1;
                let attr_data = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 2;

                if (flags & ptg::ATTR_SPACE) != 0 {
                    tokens.push(ParsedToken::AttrSpace {
                        space_type: (attr_data & 0xFF) as u8,
                        count: ((attr_data >> 8) & 0xFF) as u8,
                    });
                } else if (flags & ptg::ATTR_SUM) != 0 {
                    tokens.push(ParsedToken::AttrSum);
                } else if (flags & ptg::ATTR_IF) != 0 {
                    tokens.push(ParsedToken::AttrIf { offset: attr_data });
                } else if (flags & ptg::ATTR_CHOOSE) != 0 {
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
                }
            }

            ptg::PTG_REF => {
                if pos + 6 > data.len() {
                    break;
                }
                let (row, col, row_rel, col_rel) = parse_ref_fields(&data[pos..]);
                pos += 6;
                tokens.push(ParsedToken::Ref {
                    row,
                    col,
                    row_relative: row_rel,
                    col_relative: col_rel,
                });
            }

            ptg::PTG_AREA => {
                if pos + 12 > data.len() {
                    break;
                }
                let (fr, lr, fc, lc, frr, fcr, lrr, lcr) = parse_area_fields(&data[pos..]);
                pos += 12;
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
                if pos + 6 > data.len() {
                    break;
                }
                pos += 6;
                tokens.push(ParsedToken::RefErr);
            }

            ptg::PTG_AREA_ERR => {
                if pos + 12 > data.len() {
                    break;
                }
                pos += 12;
                tokens.push(ParsedToken::AreaErr);
            }

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
                // BIFF12 cparams is a full unsigned byte (LO reads it
                // unmasked); BIFF8's fPrompt bit does not exist here.
                // UDF calls are identified by iftab 0xFF alone.
                let argc = data[pos];
                pos += 1;
                let func_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 2;
                tokens.push(ParsedToken::FuncVar { argc, func_idx });
            }

            ptg::PTG_NAME => {
                if pos + 4 > data.len() {
                    break;
                }
                let name_idx =
                    u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
                pos += 4;
                tokens.push(ParsedToken::Name { name_idx });
            }

            ptg::PTG_NAME_X => {
                if pos + 6 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                let name_idx = u32::from_le_bytes([
                    data[pos + 2],
                    data[pos + 3],
                    data[pos + 4],
                    data[pos + 5],
                ]);
                pos += 6;
                tokens.push(ParsedToken::NameX {
                    extern_sheet_idx,
                    name_idx,
                });
            }

            ptg::PTG_REF_3D => {
                if pos + 8 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                let (row, col, row_rel, col_rel) = parse_ref_fields(&data[pos + 2..]);
                pos += 8;
                tokens.push(ParsedToken::Ref3d {
                    extern_sheet_idx,
                    row,
                    col,
                    row_relative: row_rel,
                    col_relative: col_rel,
                });
            }

            ptg::PTG_AREA_3D => {
                if pos + 14 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                let (fr, lr, fc, lc, frr, fcr, lrr, lcr) = parse_area_fields(&data[pos + 2..]);
                pos += 14;
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
                if pos + 8 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 8;
                tokens.push(ParsedToken::RefErr3d { extern_sheet_idx });
            }

            ptg::PTG_AREA_ERR_3D => {
                if pos + 14 > data.len() {
                    break;
                }
                let extern_sheet_idx = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 14;
                tokens.push(ParsedToken::AreaErr3d { extern_sheet_idx });
            }

            ptg::PTG_EXP => {
                if pos + 6 > data.len() {
                    break;
                }
                let row =
                    u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
                let col = u16::from_le_bytes([data[pos + 4], data[pos + 5]]);
                pos += 6;
                tokens.push(ParsedToken::Exp { row, col });
            }

            ptg::PTG_TBL => {
                if pos + 6 > data.len() {
                    break;
                }
                let row =
                    u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
                let col = u16::from_le_bytes([data[pos + 4], data[pos + 5]]);
                pos += 6;
                tokens.push(ParsedToken::Table { row, col });
            }

            ptg::PTG_MEM_FUNC => {
                if pos + 2 > data.len() {
                    break;
                }
                let subexpr_len = u16::from_le_bytes([data[pos], data[pos + 1]]);
                pos += 2;
                tokens.push(ParsedToken::MemFunc { subexpr_len });
            }

            ptg::PTG_MEM_AREA | ptg::PTG_MEM_ERR | ptg::PTG_MEM_NO_MEM => {
                if pos + 6 > data.len() {
                    break;
                }
                let subexpr_len = u16::from_le_bytes([data[pos + 4], data[pos + 5]]);
                pos += 6;
                tokens.push(ParsedToken::MemFunc { subexpr_len });
            }

            ptg::PTG_REF_N => {
                if pos + 6 > data.len() {
                    break;
                }
                let (row_off, col_off, row_rel, col_rel) = parse_refn_fields(&data[pos..]);
                pos += 6;
                tokens.push(ParsedToken::RefN {
                    row_offset: row_off,
                    col_offset: col_off,
                    row_relative: row_rel,
                    col_relative: col_rel,
                });
            }

            ptg::PTG_AREA_N => {
                if pos + 12 > data.len() {
                    break;
                }
                let (fro, lro, fco, lco, frr, fcr, lrr, lcr) = parse_arean_fields(&data[pos..]);
                pos += 12;
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
                // BIFF12: 14 reserved bytes after the ptg byte (15 total)
                if pos + 14 > data.len() {
                    break;
                }
                pos += 14;
                let text = parse_array_constant(extra_data, &mut extra_pos);
                tokens.push(ParsedToken::Array { text });
            }

            _ => {
                tokens.push(ParsedToken::Unknown(raw_byte));
                break;
            }
        }
    }

    tokens
}

/// Parse BIFF12 ref fields: row(u32) + col_rw(u16) = 6 bytes.
/// col_rw bits 0-13 = column, bit 14 = fRwRel, bit 15 = fColRel.
fn parse_ref_fields(data: &[u8]) -> (u32, u16, bool, bool) {
    let row = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let col_rw = u16::from_le_bytes([data[4], data[5]]);
    let col = col_rw & 0x3FFF;
    let row_rel = (col_rw & 0x4000) != 0;
    let col_rel = (col_rw & 0x8000) != 0;
    (row, col, row_rel, col_rel)
}

/// Parse BIFF12 area fields: fr(u32) + lr(u32) + fc_rw(u16) + lc_rw(u16) = 12 bytes.
fn parse_area_fields(data: &[u8]) -> (u32, u32, u16, u16, bool, bool, bool, bool) {
    let first_row = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let last_row = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
    let first_col_rw = u16::from_le_bytes([data[8], data[9]]);
    let last_col_rw = u16::from_le_bytes([data[10], data[11]]);

    let first_col = first_col_rw & 0x3FFF;
    let first_row_rel = (first_col_rw & 0x4000) != 0;
    let first_col_rel = (first_col_rw & 0x8000) != 0;

    let last_col = last_col_rw & 0x3FFF;
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

/// Parse BIFF12 tRefN: row_off(i32) + col_rw(u16) = 6 bytes.
/// Column offset is sign-extended from 14 bits.
fn parse_refn_fields(data: &[u8]) -> (i32, i16, bool, bool) {
    let row_off = i32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let col_rw = u16::from_le_bytes([data[4], data[5]]);
    let raw_col = col_rw & 0x3FFF;
    let col_off = if raw_col & 0x2000 != 0 {
        (raw_col | 0xC000) as i16
    } else {
        raw_col as i16
    };
    let row_rel = (col_rw & 0x4000) != 0;
    let col_rel = (col_rw & 0x8000) != 0;
    (row_off, col_off, row_rel, col_rel)
}

/// Parse BIFF12 tAreaN: fro(i32) + lro(i32) + fc_rw(u16) + lc_rw(u16) = 12 bytes.
fn parse_arean_fields(data: &[u8]) -> (i32, i32, i16, i16, bool, bool, bool, bool) {
    let first_row_off = i32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let last_row_off = i32::from_le_bytes([data[4], data[5], data[6], data[7]]);
    let first_col_rw = u16::from_le_bytes([data[8], data[9]]);
    let last_col_rw = u16::from_le_bytes([data[10], data[11]]);

    let first_raw = first_col_rw & 0x3FFF;
    let first_col_off = if first_raw & 0x2000 != 0 {
        (first_raw | 0xC000) as i16
    } else {
        first_raw as i16
    };
    let first_row_rel = (first_col_rw & 0x4000) != 0;
    let first_col_rel = (first_col_rw & 0x8000) != 0;

    let last_raw = last_col_rw & 0x3FFF;
    let last_col_off = if last_raw & 0x2000 != 0 {
        (last_raw | 0xC000) as i16
    } else {
        last_raw as i16
    };
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

/// Parse a BIFF12 array constant from the extra-data section.
///
/// BIFF12 format: rows(u32) + cols(u32) + row-major elements. Verified
/// against native Excel authoring (e.g. {1,2,3} (1x3) → 01 00 00 00 03 00 00
/// 00). SerAr types: 0x00=number, 0x01=string, 0x02=bool, 0x04=error,
/// 0x10=empty.
fn parse_array_constant(extra: &[u8], epos: &mut usize) -> String {
    if *epos + 8 > extra.len() {
        return "{<?>}".to_string();
    }
    let nr = u32::from_le_bytes([
        extra[*epos],
        extra[*epos + 1],
        extra[*epos + 2],
        extra[*epos + 3],
    ]) as usize;
    let nc = u32::from_le_bytes([
        extra[*epos + 4],
        extra[*epos + 5],
        extra[*epos + 6],
        extra[*epos + 7],
    ]) as usize;
    *epos += 8;

    // Each element occupies at least one byte (its type tag), so the
    // claimed dimensions must fit the remaining data. Untrusted counts
    // would otherwise drive a rows*cols loop (up to ~1.8e19) or a
    // multi-gigabyte Vec reservation.
    let remaining = extra.len() - *epos;
    if nc == 0 || nr == 0 || nr.checked_mul(nc).is_none_or(|n| n > remaining) {
        *epos = extra.len();
        return "{<?>}".to_string();
    }

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
                    // number: f64 (8 bytes)
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
                0x01 => {
                    // string: u16 char_count + UTF-16LE (ShortXLUnicodeString)
                    if *epos + 2 > extra.len() {
                        "<?>".to_string()
                    } else {
                        let slen = u16::from_le_bytes([extra[*epos], extra[*epos + 1]]) as usize;
                        *epos += 2;
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
                            let s = String::from_utf16_lossy(&chars);
                            format!("\"{}\"", s.replace('"', "\"\""))
                        }
                    }
                }
                0x02 => {
                    if *epos >= extra.len() {
                        "<?>".to_string()
                    } else {
                        let val = extra[*epos] != 0;
                        *epos += 1;
                        if val { "TRUE" } else { "FALSE" }.to_string()
                    }
                }
                0x04 => {
                    if *epos >= extra.len() {
                        "<?>".to_string()
                    } else {
                        let code = extra[*epos];
                        // error byte + 3 reserved bytes (LO skips 3).
                        *epos += 1;
                        if *epos + 3 <= extra.len() {
                            *epos += 3;
                        }
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
                0x10 => {
                    // empty: 8 bytes padding
                    if *epos + 8 <= extra.len() {
                        *epos += 8;
                    }
                    String::new()
                }
                _ => {
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

#[cfg(test)]
mod tests {
    use super::*;

    fn make_ref(row: u32, col: u16, row_rel: bool, col_rel: bool) -> Vec<u8> {
        let mut out = row.to_le_bytes().to_vec();
        let mut col_word: u16 = col;
        if row_rel {
            col_word |= 0x4000;
        }
        if col_rel {
            col_word |= 0x8000;
        }
        out.extend_from_slice(&col_word.to_le_bytes());
        out
    }

    #[test]
    fn test_parse_int_constant() {
        let data = [0x1E, 0x2A, 0x00];
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Int(42)]);
    }

    #[test]
    fn test_parse_num_constant() {
        let val: f64 = 3.14;
        let mut data = vec![0x1F];
        data.extend_from_slice(&val.to_le_bytes());
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Num(3.14)]);
    }

    #[test]
    fn test_parse_bool() {
        assert_eq!(parse_tokens(&[0x1D, 0x01]), vec![ParsedToken::Bool(true)]);
        assert_eq!(parse_tokens(&[0x1D, 0x00]), vec![ParsedToken::Bool(false)]);
    }

    #[test]
    fn test_parse_err() {
        assert_eq!(parse_tokens(&[0x1C, 0x0F]), vec![ParsedToken::Err(0x0F)]);
    }

    #[test]
    fn test_parse_str_wide() {
        let mut data = vec![0x17];
        data.extend_from_slice(&3u16.to_le_bytes()); // 3 chars (u16)
        for &ch in &[b'A', b'B', b'C'] {
            data.push(ch);
            data.push(0x00);
        }
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Str("ABC".to_string())]);
    }

    #[test]
    fn test_parse_ref() {
        // tRefV (0x44): row=5, col=2, absolute
        let mut data = vec![0x44];
        data.extend_from_slice(&make_ref(5, 2, false, false));
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Ref {
                row: 5,
                col: 2,
                row_relative: false,
                col_relative: false,
            }]
        );
    }

    #[test]
    fn test_parse_ref_relative() {
        // tRefV: row=4, col=2, both relative
        let mut data = vec![0x44];
        data.extend_from_slice(&make_ref(4, 2, true, true));
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
    fn test_parse_ref_large_row() {
        // Row 100000 - beyond BIFF8 u16 range
        let mut data = vec![0x44];
        data.extend_from_slice(&make_ref(100_000, 5, true, true));
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Ref {
                row: 100_000,
                col: 5,
                row_relative: true,
                col_relative: true,
            }]
        );
    }

    #[test]
    fn test_parse_area() {
        let mut data = vec![0x45];
        data.extend_from_slice(&0u32.to_le_bytes()); // first_row
        data.extend_from_slice(&9u32.to_le_bytes()); // last_row
        data.extend_from_slice(&0u16.to_le_bytes()); // first_col word
        data.extend_from_slice(&2u16.to_le_bytes()); // last_col word
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Area {
                first_row: 0,
                last_row: 9,
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
    fn test_parse_func() {
        let mut data = vec![0x41];
        data.extend_from_slice(&4u16.to_le_bytes());
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Func { func_idx: 4 }]);
    }

    #[test]
    fn test_parse_funcvar() {
        let mut data = vec![0x42, 0x03];
        data.extend_from_slice(&1u16.to_le_bytes());
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
        let data = [0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03];
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Int(1), ParsedToken::Int(2), ParsedToken::Add,]
        );
    }

    #[test]
    fn test_parse_attr_sum() {
        let mut data = vec![0x19, 0x10];
        data.extend_from_slice(&0u16.to_le_bytes());
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::AttrSum]);
    }

    #[test]
    fn test_parse_ref3d() {
        // tRef3dV (0x5A = 0x3A + 0x20): extern_sheet=0, row=0, col=0
        let mut data = vec![0x5A];
        data.extend_from_slice(&0u16.to_le_bytes()); // extern_sheet
        data.extend_from_slice(&make_ref(0, 0, false, false));
        let tokens = parse_tokens(&data);
        assert_eq!(
            tokens,
            vec![ParsedToken::Ref3d {
                extern_sheet_idx: 0,
                row: 0,
                col: 0,
                row_relative: false,
                col_relative: false,
            }]
        );
    }

    #[test]
    fn test_parse_name() {
        let mut data = vec![0x43];
        data.extend_from_slice(&1u32.to_le_bytes());
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Name { name_idx: 1 }]);
    }

    #[test]
    fn test_parse_exp() {
        let mut data = vec![0x01];
        data.extend_from_slice(&10u32.to_le_bytes());
        data.extend_from_slice(&5u16.to_le_bytes());
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Exp { row: 10, col: 5 }]);
    }

    #[test]
    fn test_parse_tbl() {
        let mut data = vec![0x02];
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::Table { row: 2, col: 3 }]);
    }

    #[test]
    fn test_parse_refn_signed() {
        let mut data = vec![0x4C];
        data.extend_from_slice(&(-1i32).to_le_bytes());
        // col_off = -2 as 14-bit = 0x3FFE, flags: row_rel(0x4000) + col_rel(0x8000)
        let col_word: u16 = ((-2i16 as u16) & 0x3FFF) | 0x4000 | 0x8000;
        data.extend_from_slice(&col_word.to_le_bytes());
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
    fn test_parse_classified_variants() {
        // tRefR (0x24), tRefV (0x44), tRefA (0x64) all produce same Ref
        for ptg_byte in [0x24, 0x44, 0x64] {
            let mut data = vec![ptg_byte];
            data.extend_from_slice(&make_ref(0, 0, false, false));
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
    fn test_parse_memfunc() {
        let mut data = vec![0x49];
        data.extend_from_slice(&10u16.to_le_bytes());
        let tokens = parse_tokens(&data);
        assert_eq!(tokens, vec![ParsedToken::MemFunc { subexpr_len: 10 }]);
    }

    #[test]
    fn test_parse_array_constant_simple() {
        // BIFF12: 1 ptg byte (0x60 = tArray A-class) + 14 reserved bytes = 15
        let mut token_data = vec![0x60];
        token_data.extend_from_slice(&[0u8; 14]);
        let mut extra = Vec::new();
        // BIFF12: rows(u32) + cols(u32) — rows first, matching native Excel.
        extra.extend_from_slice(&1u32.to_le_bytes()); // 1 row
        extra.extend_from_slice(&3u32.to_le_bytes()); // 3 cols
                                                      // BIFF12 SerAr: type 0x00 = number (f64)
        for val in [1.0f64, 2.0, 3.0] {
            extra.push(0x00);
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
    fn test_parse_array_constant_rejects_implausible_counts() {
        // A malicious rgcb header can claim huge dimensions; each
        // element needs at least one byte, so the parser must bail
        // instead of looping rows*cols times or reserving gigabytes.
        let mut token_data = vec![0x60];
        token_data.extend_from_slice(&[0u8; 14]);
        let mut extra = Vec::new();
        extra.extend_from_slice(&1000u32.to_le_bytes()); // claimed rows
        extra.extend_from_slice(&1000u32.to_le_bytes()); // claimed cols
        extra.push(0x00); // one real element, not a million
        extra.extend_from_slice(&1.0f64.to_le_bytes());

        let tokens = parse_tokens_with_extra(&token_data, &extra);
        assert_eq!(tokens.len(), 1);
        let ParsedToken::Array { text } = &tokens[0] else {
            panic!("expected Array token, got {:?}", tokens[0]);
        };
        assert!(
            text.len() < 64,
            "implausible counts must not synthesize a giant literal; got {} chars",
            text.len()
        );
    }

    #[test]
    fn test_ptg_name_index_is_not_truncated() {
        // BIFF12 PtgName carries a 4-byte nameindex; values above
        // 0xFFFF must survive parsing untruncated.
        let mut data = vec![0x43u8]; // PtgName V-class
        data.extend_from_slice(&0x0001_2345u32.to_le_bytes());
        let tokens = parse_tokens_with_extra(&data, &[]);
        assert_eq!(
            tokens,
            vec![ParsedToken::Name {
                name_idx: 0x0001_2345
            }]
        );
    }
}
