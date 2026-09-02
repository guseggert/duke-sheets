//! Decompilation of the BIFF token subset used by pivot formulas.

use super::{function_argc, function_name};

/// How a format interprets the argument-count byte in `PtgFuncVar`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotVariableArgCount {
    /// BIFF8 uses bits 0-6 for the count and bit 7 for `fPrompt`.
    Biff8,
    /// BIFF12 uses the complete byte as the argument count.
    Biff12,
}

impl PivotVariableArgCount {
    fn decode(self, encoded: u8) -> usize {
        match self {
            Self::Biff8 => (encoded & 0x7f) as usize,
            Self::Biff12 => encoded as usize,
        }
    }
}

/// A resolved function used while decompiling a pivot formula.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotFormulaFunction {
    /// Formula function name.
    pub name: String,
    /// Number of operands consumed from the formula stack.
    pub argc: usize,
}

/// Format-specific decoding and name-resolution hooks for pivot formulas.
///
/// The defaults implement the common BIFF8/BIFF12 layouts. Implementors only
/// need to override methods where their container uses a different layout.
pub trait PivotFormulaHooks {
    /// Resolve a zero-based `PtgSxName` index to formula text.
    fn resolve_name(&mut self, index: u32) -> Option<String>;

    /// Interpret the `PtgFuncVar` argument-count byte.
    fn variable_arg_count(&self) -> PivotVariableArgCount;

    /// Read a `PtgStr` payload and advance `offset` past it.
    fn read_string(
        &mut self,
        data: &[u8],
        offset: &mut usize,
    ) -> Result<String, PivotFormulaError> {
        read_biff_short_string(data, offset)
    }

    /// Read a `PtgSxName` payload and advance `offset` past it.
    fn read_name_index(
        &mut self,
        data: &[u8],
        offset: &mut usize,
    ) -> Result<u32, PivotFormulaError> {
        let subtype = take_u8(data, offset)?;
        if subtype != 0x1d {
            return Err(PivotFormulaError::UnsupportedExtension(subtype));
        }
        take_u32(data, offset)
    }

    /// Resolve a function id and its encoded argument count.
    fn resolve_function(
        &mut self,
        function_id: u16,
        variable_argc: Option<u8>,
    ) -> Option<PivotFormulaFunction> {
        let name = function_name(function_id);
        if name.is_empty() {
            return None;
        }
        let argc = match variable_argc {
            Some(encoded) => self.variable_arg_count().decode(encoded),
            None => {
                let argc = function_argc(function_id);
                if argc > 253 {
                    return None;
                }
                argc as usize
            }
        };
        Some(PivotFormulaFunction {
            name: name.to_string(),
            argc,
        })
    }
}

/// Failure encountered while decoding or reducing a pivot formula token stream.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub enum PivotFormulaError {
    /// A token payload ended before all required bytes were available.
    #[error("truncated pivot formula token payload")]
    UnexpectedEnd,
    /// The token is outside the pivot-formula subset understood by the decoder.
    #[error("unsupported pivot formula token 0x{0:02X}")]
    UnsupportedToken(u8),
    /// The `PtgElf` extension subtype is not `PtgSxName`.
    #[error("unsupported pivot formula extension 0x{0:02X}")]
    UnsupportedExtension(u8),
    /// A `PtgSxName` index could not be resolved by the caller.
    #[error("unresolved pivot formula name index {0}")]
    UnresolvedName(u32),
    /// A function id is unknown or invalid for its token kind.
    #[error("unsupported pivot formula function id {0}")]
    UnsupportedFunction(u16),
    /// An operator or function did not have enough operands.
    #[error("malformed pivot formula stack")]
    StackUnderflow,
    /// More than one expression remained after consuming all tokens.
    #[error("pivot formula left {0} operands on the stack")]
    TrailingOperands(usize),
}

#[derive(Debug, Clone)]
struct FormulaText {
    text: String,
    precedence: u8,
}

impl FormulaText {
    fn atom(text: String) -> Self {
        Self {
            text,
            precedence: 8,
        }
    }
}

/// Decompile a BIFF pivot formula token stream into infix formula text.
///
/// `hooks` supplies format-specific payload decoding, `PtgSxName` resolution,
/// and the BIFF8-versus-BIFF12 `PtgFuncVar` argument-count behavior.
pub fn decompile_pivot_formula(
    tokens: &[u8],
    hooks: &mut impl PivotFormulaHooks,
) -> Result<String, PivotFormulaError> {
    let mut stack = Vec::new();
    let mut pos = 0;
    while pos < tokens.len() {
        let token = take_u8(tokens, &mut pos)?;
        match token {
            0x03 => push_binary(&mut stack, "+", 3, false)?,
            0x04 => push_binary(&mut stack, "-", 3, true)?,
            0x05 => push_binary(&mut stack, "*", 4, false)?,
            0x06 => push_binary(&mut stack, "/", 4, true)?,
            0x07 => push_binary(&mut stack, "^", 5, true)?,
            0x08 => push_binary(&mut stack, "&", 2, false)?,
            0x09 => push_binary(&mut stack, "<", 1, true)?,
            0x0a => push_binary(&mut stack, "<=", 1, true)?,
            0x0b => push_binary(&mut stack, "=", 1, true)?,
            0x0c => push_binary(&mut stack, ">=", 1, true)?,
            0x0d => push_binary(&mut stack, ">", 1, true)?,
            0x0e => push_binary(&mut stack, "<>", 1, true)?,
            0x12 => push_prefix(&mut stack, "+")?,
            0x13 => push_prefix(&mut stack, "-")?,
            0x14 => push_percent(&mut stack)?,
            0x15 => push_paren(&mut stack)?,
            0x17 => {
                let value = hooks.read_string(tokens, &mut pos)?;
                stack.push(FormulaText::atom(format!(
                    "\"{}\"",
                    value.replace('"', "\"\"")
                )));
            }
            0x18 => {
                let index = hooks.read_name_index(tokens, &mut pos)?;
                let name = hooks
                    .resolve_name(index)
                    .ok_or(PivotFormulaError::UnresolvedName(index))?;
                stack.push(FormulaText::atom(name));
            }
            token if base_ptg(token) == 0x21 => {
                let function_id = take_u16(tokens, &mut pos)?;
                let function = hooks
                    .resolve_function(function_id, None)
                    .ok_or(PivotFormulaError::UnsupportedFunction(function_id))?;
                push_function(&mut stack, function)?;
            }
            token if base_ptg(token) == 0x22 => {
                let argc = take_u8(tokens, &mut pos)?;
                let function_id = take_u16(tokens, &mut pos)?;
                let function = hooks
                    .resolve_function(function_id, Some(argc))
                    .ok_or(PivotFormulaError::UnsupportedFunction(function_id))?;
                push_function(&mut stack, function)?;
            }
            0x1c => {
                let code = take_u8(tokens, &mut pos)?;
                stack.push(FormulaText::atom(format_formula_error(code).to_string()));
            }
            0x1d => {
                let value = take_u8(tokens, &mut pos)?;
                stack.push(FormulaText::atom(
                    if value == 0 { "FALSE" } else { "TRUE" }.into(),
                ));
            }
            0x1e => stack.push(FormulaText::atom(take_u16(tokens, &mut pos)?.to_string())),
            0x1f => {
                let bytes = take(tokens, &mut pos, 8)?;
                let bytes: [u8; 8] = bytes
                    .try_into()
                    .map_err(|_| PivotFormulaError::UnexpectedEnd)?;
                let value = f64::from_le_bytes(bytes);
                stack.push(FormulaText::atom(format_formula_number(value)));
            }
            _ => return Err(PivotFormulaError::UnsupportedToken(token)),
        }
    }

    if stack.len() != 1 {
        return Err(if stack.is_empty() {
            PivotFormulaError::StackUnderflow
        } else {
            PivotFormulaError::TrailingOperands(stack.len())
        });
    }
    Ok(stack.pop().ok_or(PivotFormulaError::StackUnderflow)?.text)
}

fn read_biff_short_string(data: &[u8], offset: &mut usize) -> Result<String, PivotFormulaError> {
    let char_count = take_u8(data, offset)? as usize;
    let flags = take_u8(data, offset)?;
    if flags & 1 != 0 {
        let bytes = take(
            data,
            offset,
            char_count
                .checked_mul(2)
                .ok_or(PivotFormulaError::UnexpectedEnd)?,
        )?;
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        Ok(String::from_utf16_lossy(&units))
    } else {
        Ok(take(data, offset, char_count)?
            .iter()
            .map(|&byte| byte as char)
            .collect())
    }
}

fn take<'a>(data: &'a [u8], offset: &mut usize, len: usize) -> Result<&'a [u8], PivotFormulaError> {
    let end = offset
        .checked_add(len)
        .ok_or(PivotFormulaError::UnexpectedEnd)?;
    let value = data
        .get(*offset..end)
        .ok_or(PivotFormulaError::UnexpectedEnd)?;
    *offset = end;
    Ok(value)
}

fn take_u8(data: &[u8], offset: &mut usize) -> Result<u8, PivotFormulaError> {
    Ok(take(data, offset, 1)?[0])
}

fn take_u16(data: &[u8], offset: &mut usize) -> Result<u16, PivotFormulaError> {
    Ok(u16::from_le_bytes(
        take(data, offset, 2)?.try_into().unwrap(),
    ))
}

fn take_u32(data: &[u8], offset: &mut usize) -> Result<u32, PivotFormulaError> {
    Ok(u32::from_le_bytes(
        take(data, offset, 4)?.try_into().unwrap(),
    ))
}

fn base_ptg(token: u8) -> u8 {
    if token >= 0x20 {
        (token & 0x1f) | 0x20
    } else {
        token
    }
}

fn push_binary(
    stack: &mut Vec<FormulaText>,
    op: &str,
    precedence: u8,
    parenthesize_equal_right: bool,
) -> Result<(), PivotFormulaError> {
    let right = stack.pop().ok_or(PivotFormulaError::StackUnderflow)?;
    let left = stack.pop().ok_or(PivotFormulaError::StackUnderflow)?;
    let left = if left.precedence < precedence {
        format!("({})", left.text)
    } else {
        left.text
    };
    let right = if right.precedence < precedence
        || (parenthesize_equal_right && right.precedence == precedence)
    {
        format!("({})", right.text)
    } else {
        right.text
    };
    stack.push(FormulaText {
        text: format!("{left}{op}{right}"),
        precedence,
    });
    Ok(())
}

fn push_prefix(stack: &mut Vec<FormulaText>, op: &str) -> Result<(), PivotFormulaError> {
    let operand = stack.pop().ok_or(PivotFormulaError::StackUnderflow)?;
    let text = if operand.precedence < 6 {
        format!("{op}({})", operand.text)
    } else {
        format!("{op}{}", operand.text)
    };
    stack.push(FormulaText {
        text,
        precedence: 6,
    });
    Ok(())
}

fn push_percent(stack: &mut Vec<FormulaText>) -> Result<(), PivotFormulaError> {
    let operand = stack.pop().ok_or(PivotFormulaError::StackUnderflow)?;
    let text = if operand.precedence < 7 {
        format!("({})%", operand.text)
    } else {
        format!("{}%", operand.text)
    };
    stack.push(FormulaText {
        text,
        precedence: 7,
    });
    Ok(())
}

fn push_paren(stack: &mut Vec<FormulaText>) -> Result<(), PivotFormulaError> {
    let operand = stack.pop().ok_or(PivotFormulaError::StackUnderflow)?;
    stack.push(FormulaText::atom(format!("({})", operand.text)));
    Ok(())
}

fn push_function(
    stack: &mut Vec<FormulaText>,
    function: PivotFormulaFunction,
) -> Result<(), PivotFormulaError> {
    if stack.len() < function.argc {
        return Err(PivotFormulaError::StackUnderflow);
    }
    let start = stack.len() - function.argc;
    let args = stack.drain(start..).map(|arg| arg.text).collect::<Vec<_>>();
    stack.push(FormulaText::atom(format!(
        "{}({})",
        function.name,
        args.join(",")
    )));
    Ok(())
}

fn format_formula_error(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0f => "#VALUE!",
        0x17 => "#REF!",
        0x1d => "#NAME?",
        0x24 => "#NUM!",
        0x2a => "#N/A",
        0x2b => "#GETTING_DATA",
        _ => "#VALUE!",
    }
}

fn format_formula_number(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    struct Hooks {
        argc: PivotVariableArgCount,
        names: Vec<String>,
    }

    impl PivotFormulaHooks for Hooks {
        fn resolve_name(&mut self, index: u32) -> Option<String> {
            self.names.get(index as usize).cloned()
        }

        fn variable_arg_count(&self) -> PivotVariableArgCount {
            self.argc
        }
    }

    fn decompile(tokens: &[u8]) -> Result<String, PivotFormulaError> {
        decompile_pivot_formula(
            tokens,
            &mut Hooks {
                argc: PivotVariableArgCount::Biff12,
                names: vec!["Sales Total".into(), "East".into()],
            },
        )
    }

    #[test]
    fn decompiles_literals() {
        assert_eq!(
            decompile(&[0x17, 3, 0, b'a', b'"', b'b']).unwrap(),
            "\"a\"\"b\""
        );
        assert_eq!(decompile(&[0x1d, 1]).unwrap(), "TRUE");
        assert_eq!(decompile(&[0x1c, 7]).unwrap(), "#DIV/0!");
        assert_eq!(
            decompile(&[0x1f, 0, 0, 0, 0, 0, 0, 4, 0x40]).unwrap(),
            "2.5"
        );
    }

    #[test]
    fn decompiles_unary_and_binary_operations_with_precedence() {
        let tokens = [0x1e, 1, 0, 0x1e, 2, 0, 0x03, 0x13, 0x1e, 3, 0, 0x05, 0x14];
        assert_eq!(decompile(&tokens).unwrap(), "(-(1+2)*3)%");
    }

    #[test]
    fn decompiles_fixed_and_variable_functions() {
        assert_eq!(decompile(&[0x1e, 2, 0, 0x21, 24, 0]).unwrap(), "ABS(2)");
        let tokens = [0x1e, 1, 0, 0x1e, 2, 0, 0x22, 2, 4, 0];
        assert_eq!(decompile(&tokens).unwrap(), "SUM(1,2)");
    }

    #[test]
    fn resolves_field_and_item_names() {
        let field = [0x18, 0x1d, 0, 0, 0, 0];
        let item = [0x18, 0x1d, 1, 0, 0, 0];
        assert_eq!(decompile(&field).unwrap(), "Sales Total");
        assert_eq!(decompile(&item).unwrap(), "East");
    }

    #[test]
    fn rejects_malformed_stacks_and_unsupported_tokens() {
        assert_eq!(decompile(&[0x03]), Err(PivotFormulaError::StackUnderflow));
        assert_eq!(
            decompile(&[0x1e, 1, 0, 0x1e, 2, 0]),
            Err(PivotFormulaError::TrailingOperands(2))
        );
        assert_eq!(
            decompile(&[0xff]),
            Err(PivotFormulaError::UnsupportedToken(0xff))
        );
    }

    #[test]
    fn variable_arg_count_preserves_biff8_prompt_bit_semantics() {
        let tokens = [0x1e, 1, 0, 0x1e, 2, 0, 0x22, 0x82, 4, 0];

        let mut biff8 = Hooks {
            argc: PivotVariableArgCount::Biff8,
            names: vec![],
        };
        assert_eq!(
            decompile_pivot_formula(&tokens, &mut biff8).unwrap(),
            "SUM(1,2)"
        );

        let mut biff12 = Hooks {
            argc: PivotVariableArgCount::Biff12,
            names: vec![],
        };
        assert_eq!(
            decompile_pivot_formula(&tokens, &mut biff12),
            Err(PivotFormulaError::StackUnderflow)
        );
    }
}
