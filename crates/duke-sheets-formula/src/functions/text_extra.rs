//! Additional text functions

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

fn to_i64_trunc(v: &FormulaValue) -> Option<i64> {
    v.as_number().map(|n| n.trunc() as i64)
}

fn scalar_string(v: &FormulaValue) -> Result<String, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => Ok(v.as_string()),
    }
}

fn scalar_bool(v: &FormulaValue) -> Result<bool, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => v.as_bool().ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn scalar_i64(v: &FormulaValue) -> Result<i64, FormulaValue> {
    match v {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array(_) => Err(FormulaValue::Error(CellError::Value)),
        _ => to_i64_trunc(v).ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn split_text(text: &str, delim: &str, ignore_empty: bool, case_insensitive: bool) -> Vec<String> {
    if delim.is_empty() {
        let mut out: Vec<String> = text.chars().map(|c| c.to_string()).collect();
        if ignore_empty {
            out.retain(|s| !s.is_empty());
        }
        return out;
    }

    let mut out = Vec::new();
    let mut start = 0usize;
    let (search_text, search_delim) = if case_insensitive {
        (text.to_ascii_lowercase(), delim.to_ascii_lowercase())
    } else {
        (text.to_string(), delim.to_string())
    };

    while let Some(rel) = search_text[start..].find(&search_delim) {
        let pos = start + rel;
        let piece = &text[start..pos];
        if !ignore_empty || !piece.is_empty() {
            out.push(piece.to_string());
        }
        start = pos + delim.len();
    }

    let tail = &text[start..];
    if !ignore_empty || !tail.is_empty() {
        out.push(tail.to_string());
    }
    out
}

fn parse_match_mode(v: Option<&FormulaValue>) -> Result<bool, FormulaValue> {
    let Some(v) = v else {
        return Ok(false);
    };
    let mode = scalar_i64(v)?;
    match mode {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn delimiter_positions(text: &str, delimiter: &str, case_insensitive: bool) -> Vec<usize> {
    if delimiter.is_empty() {
        return vec![0];
    }

    let (search_text, search_delim) = if case_insensitive {
        (text.to_ascii_lowercase(), delimiter.to_ascii_lowercase())
    } else {
        (text.to_string(), delimiter.to_string())
    };

    search_text
        .match_indices(&search_delim)
        .map(|(idx, _)| idx)
        .collect()
}

fn text_before_after(
    text: &str,
    delimiter: &str,
    instance_num: i64,
    case_insensitive: bool,
    match_end: bool,
    if_not_found: FormulaValue,
    before: bool,
) -> FormulaValue {
    if instance_num == 0 {
        return FormulaValue::Error(CellError::Value);
    }

    if delimiter.is_empty() {
        if before {
            return if instance_num > 0 {
                FormulaValue::String(String::new())
            } else {
                FormulaValue::String(text.to_string())
            };
        }
        return if instance_num > 0 {
            FormulaValue::String(text.to_string())
        } else {
            FormulaValue::String(String::new())
        };
    }

    let mut positions = delimiter_positions(text, delimiter, case_insensitive);
    if match_end {
        positions.push(text.len());
    }

    if positions.is_empty() {
        return if_not_found;
    }

    let idx_opt = if instance_num > 0 {
        let idx = (instance_num - 1) as usize;
        if idx < positions.len() {
            Some(idx)
        } else {
            None
        }
    } else {
        let back = (-instance_num) as usize;
        if back <= positions.len() {
            Some(positions.len() - back)
        } else {
            None
        }
    };

    let Some(idx) = idx_opt else {
        return if_not_found;
    };

    let pos = positions[idx];
    if before {
        FormulaValue::String(text[..pos].to_string())
    } else if pos == text.len() {
        FormulaValue::String(String::new())
    } else {
        FormulaValue::String(text[pos + delimiter.len()..].to_string())
    }
}

/// REPLACE(old_text, start_num, num_chars, new_text)
pub fn fn_replace(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let old_text = match scalar_string(args.get(0).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let start_num = match scalar_i64(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let num_chars = match scalar_i64(args.get(2).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let new_text = match scalar_string(args.get(3).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if start_num < 1 || num_chars < 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    let chars: Vec<char> = old_text.chars().collect();
    let len = chars.len();
    let start0 = (start_num - 1) as usize;
    let prefix_end = start0.min(len);
    let suffix_start = (start0.saturating_add(num_chars as usize)).min(len);

    let mut out = String::new();
    out.extend(chars[..prefix_end].iter());
    out.push_str(&new_text);
    out.extend(chars[suffix_start..].iter());

    Ok(FormulaValue::String(out))
}

/// REPLACEB(old_text, start_num, num_chars, new_text)
pub fn fn_replaceb(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_replace(args, _ctx)
}

/// TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
pub fn fn_textbefore(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let text = match scalar_string(args.get(0).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let delimiter = match scalar_string(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let instance_num = match args.get(2) {
        Some(v) => match scalar_i64(v) {
            Ok(n) => n,
            Err(e) => return Ok(e),
        },
        None => 1,
    };

    let case_insensitive = match parse_match_mode(args.get(3)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let match_end = match args.get(4) {
        Some(v) => match scalar_bool(v) {
            Ok(b) => b,
            Err(e) => return Ok(e),
        },
        None => false,
    };

    let if_not_found = args
        .get(5)
        .cloned()
        .unwrap_or(FormulaValue::Error(CellError::Na));

    Ok(text_before_after(
        &text,
        &delimiter,
        instance_num,
        case_insensitive,
        match_end,
        if_not_found,
        true,
    ))
}

/// TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
pub fn fn_textafter(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let text = match scalar_string(args.get(0).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let delimiter = match scalar_string(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let instance_num = match args.get(2) {
        Some(v) => match scalar_i64(v) {
            Ok(n) => n,
            Err(e) => return Ok(e),
        },
        None => 1,
    };

    let case_insensitive = match parse_match_mode(args.get(3)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let match_end = match args.get(4) {
        Some(v) => match scalar_bool(v) {
            Ok(b) => b,
            Err(e) => return Ok(e),
        },
        None => false,
    };

    let if_not_found = args
        .get(5)
        .cloned()
        .unwrap_or(FormulaValue::Error(CellError::Na));

    Ok(text_before_after(
        &text,
        &delimiter,
        instance_num,
        case_insensitive,
        match_end,
        if_not_found,
        false,
    ))
}

/// TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])
pub fn fn_textsplit(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let text = match scalar_string(args.get(0).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let col_delimiter = match scalar_string(args.get(1).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let row_delimiter = match args.get(2) {
        Some(v) if !matches!(v, FormulaValue::Empty) => match scalar_string(v) {
            Ok(s) => Some(s),
            Err(e) => return Ok(e),
        },
        _ => None,
    };

    let ignore_empty = match args.get(3) {
        Some(v) if !matches!(v, FormulaValue::Empty) => match scalar_bool(v) {
            Ok(b) => b,
            Err(e) => return Ok(e),
        },
        _ => false,
    };

    let case_insensitive = match parse_match_mode(args.get(4)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let pad_with = args
        .get(5)
        .cloned()
        .unwrap_or(FormulaValue::Error(CellError::Na));

    let row_strings = match row_delimiter {
        Some(ref rd) => split_text(&text, rd, ignore_empty, case_insensitive),
        None => vec![text],
    };

    let mut out: Vec<Vec<FormulaValue>> = row_strings
        .iter()
        .map(|row| {
            split_text(row, &col_delimiter, ignore_empty, case_insensitive)
                .into_iter()
                .map(FormulaValue::String)
                .collect::<Vec<_>>()
        })
        .collect();

    let max_cols = out.iter().map(|r| r.len()).max().unwrap_or(0);
    for row in &mut out {
        while row.len() < max_cols {
            row.push(pad_with.clone());
        }
    }

    Ok(FormulaValue::Array(out))
}

/// UNICHAR(number)
pub fn fn_unichar(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match scalar_i64(args.get(0).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if number <= 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    match char::from_u32(number as u32) {
        Some(ch) => Ok(FormulaValue::String(ch.to_string())),
        None => Ok(FormulaValue::Error(CellError::Value)),
    }
}

/// UNICODE(text)
pub fn fn_unicode(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match scalar_string(args.get(0).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match text.chars().next() {
        Some(ch) => Ok(FormulaValue::Number(ch as u32 as f64)),
        None => Ok(FormulaValue::Error(CellError::Value)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    fn s(v: &str) -> FormulaValue {
        FormulaValue::String(v.to_string())
    }

    #[test]
    fn test_replace_and_replaceb() {
        assert_eq!(
            fn_replace(
                &[
                    s("abcdefgh"),
                    FormulaValue::Number(3.0),
                    FormulaValue::Number(4.0),
                    s("Z")
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("abZgh")
        );

        assert_eq!(
            fn_replaceb(
                &[
                    s("abc"),
                    FormulaValue::Number(10.0),
                    FormulaValue::Number(2.0),
                    s("X")
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("abcX")
        );

        assert_eq!(
            fn_replace(
                &[
                    s("abc"),
                    FormulaValue::Number(0.0),
                    FormulaValue::Number(1.0),
                    s("x")
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_textbefore_cases() {
        assert_eq!(
            fn_textbefore(&[s("a-b-c"), s("-")], &EvaluationContext::simple()).unwrap(),
            s("a")
        );
        assert_eq!(
            fn_textbefore(
                &[
                    s("a-b-c"),
                    s("-"),
                    FormulaValue::Number(-1.0),
                    FormulaValue::Number(0.0),
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("a-b")
        );
        assert_eq!(
            fn_textbefore(
                &[
                    s("Alpha-Beta"),
                    s("-beta"),
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(1.0),
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("Alpha")
        );
        assert_eq!(
            fn_textbefore(
                &[
                    s("abc"),
                    s("/"),
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(0.0),
                    FormulaValue::Boolean(false),
                    s("NF"),
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("NF")
        );
    }

    #[test]
    fn test_textafter_cases() {
        assert_eq!(
            fn_textafter(&[s("a-b-c"), s("-")], &EvaluationContext::simple()).unwrap(),
            s("b-c")
        );
        assert_eq!(
            fn_textafter(
                &[
                    s("a-b-c"),
                    s("-"),
                    FormulaValue::Number(-1.0),
                    FormulaValue::Number(0.0),
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("c")
        );
        assert_eq!(
            fn_textafter(
                &[
                    s("abc"),
                    s("/"),
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(0.0),
                    FormulaValue::Boolean(false),
                    s("none"),
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("none")
        );
        assert_eq!(
            fn_textafter(
                &[
                    s("abc"),
                    s("/"),
                    FormulaValue::Number(1.0),
                    FormulaValue::Number(0.0),
                    FormulaValue::Boolean(true),
                ],
                &EvaluationContext::simple()
            )
            .unwrap(),
            s("")
        );
    }

    #[test]
    fn test_textsplit() {
        let out = fn_textsplit(
            &[s("a,b;c"), s(","), s(";"), FormulaValue::Boolean(false)],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(
            out,
            FormulaValue::Array(vec![
                vec![s("a"), s("b")],
                vec![s("c"), FormulaValue::Error(CellError::Na)]
            ])
        );

        let out2 = fn_textsplit(
            &[
                s("a||b"),
                s("|"),
                FormulaValue::Empty,
                FormulaValue::Boolean(true),
            ],
            &EvaluationContext::simple(),
        )
        .unwrap();
        assert_eq!(out2, FormulaValue::Array(vec![vec![s("a"), s("b")]]));

        let arr = eval("={\"X\",\"Y\"}").unwrap();
        if let FormulaValue::Array(rows) = arr {
            assert_eq!(rows.len(), 1);
        } else {
            panic!("expected array");
        }
    }

    #[test]
    fn test_unichar_and_unicode() {
        assert_eq!(
            fn_unichar(&[FormulaValue::Number(65.0)], &EvaluationContext::simple()).unwrap(),
            s("A")
        );
        assert_eq!(
            fn_unicode(&[s("A")], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Number(65.0)
        );
        assert_eq!(
            fn_unichar(&[FormulaValue::Number(0.0)], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
        assert_eq!(
            fn_unicode(&[s("")], &EvaluationContext::simple()).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }
}
