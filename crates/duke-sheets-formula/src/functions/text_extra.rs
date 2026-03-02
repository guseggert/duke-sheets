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

// ---------- Half-width / Full-width character conversion ----------

/// Map a single full-width character to its half-width equivalent(s).
/// Returns None if the character should pass through unchanged.
fn fullwidth_to_halfwidth(ch: char) -> Option<Vec<char>> {
    match ch {
        // Full-width ASCII (U+FF01..=U+FF5E) → ASCII (U+0021..=U+007E)
        '\u{FF01}'..='\u{FF5E}' => {
            let ascii = (ch as u32 - 0xFF01 + 0x0021) as u8 as char;
            Some(vec![ascii])
        }
        // Full-width space → ASCII space
        '\u{3000}' => Some(vec![' ']),
        // Full-width katakana → half-width katakana
        // Voiced pairs (dakuten): カ→ｶﾞ etc.
        'ガ' => Some(vec!['\u{FF76}', '\u{FF9E}']),
        'ギ' => Some(vec!['\u{FF77}', '\u{FF9E}']),
        'グ' => Some(vec!['\u{FF78}', '\u{FF9E}']),
        'ゲ' => Some(vec!['\u{FF79}', '\u{FF9E}']),
        'ゴ' => Some(vec!['\u{FF7A}', '\u{FF9E}']),
        'ザ' => Some(vec!['\u{FF7B}', '\u{FF9E}']),
        'ジ' => Some(vec!['\u{FF7C}', '\u{FF9E}']),
        'ズ' => Some(vec!['\u{FF7D}', '\u{FF9E}']),
        'ゼ' => Some(vec!['\u{FF7E}', '\u{FF9E}']),
        'ゾ' => Some(vec!['\u{FF7F}', '\u{FF9E}']),
        'ダ' => Some(vec!['\u{FF80}', '\u{FF9E}']),
        'ヂ' => Some(vec!['\u{FF81}', '\u{FF9E}']),
        'ヅ' => Some(vec!['\u{FF82}', '\u{FF9E}']),
        'デ' => Some(vec!['\u{FF83}', '\u{FF9E}']),
        'ド' => Some(vec!['\u{FF84}', '\u{FF9E}']),
        'バ' => Some(vec!['\u{FF8A}', '\u{FF9E}']),
        'ビ' => Some(vec!['\u{FF8B}', '\u{FF9E}']),
        'ブ' => Some(vec!['\u{FF8C}', '\u{FF9E}']),
        'ベ' => Some(vec!['\u{FF8D}', '\u{FF9E}']),
        'ボ' => Some(vec!['\u{FF8E}', '\u{FF9E}']),
        'ヴ' => Some(vec!['\u{FF73}', '\u{FF9E}']),
        // Semi-voiced pairs (handakuten): パ→ﾊﾟ etc.
        'パ' => Some(vec!['\u{FF8A}', '\u{FF9F}']),
        'ピ' => Some(vec!['\u{FF8B}', '\u{FF9F}']),
        'プ' => Some(vec!['\u{FF8C}', '\u{FF9F}']),
        'ペ' => Some(vec!['\u{FF8D}', '\u{FF9F}']),
        'ポ' => Some(vec!['\u{FF8E}', '\u{FF9F}']),
        // Unvoiced katakana
        'ア' => Some(vec!['\u{FF71}']),
        'イ' => Some(vec!['\u{FF72}']),
        'ウ' => Some(vec!['\u{FF73}']),
        'エ' => Some(vec!['\u{FF74}']),
        'オ' => Some(vec!['\u{FF75}']),
        'カ' => Some(vec!['\u{FF76}']),
        'キ' => Some(vec!['\u{FF77}']),
        'ク' => Some(vec!['\u{FF78}']),
        'ケ' => Some(vec!['\u{FF79}']),
        'コ' => Some(vec!['\u{FF7A}']),
        'サ' => Some(vec!['\u{FF7B}']),
        'シ' => Some(vec!['\u{FF7C}']),
        'ス' => Some(vec!['\u{FF7D}']),
        'セ' => Some(vec!['\u{FF7E}']),
        'ソ' => Some(vec!['\u{FF7F}']),
        'タ' => Some(vec!['\u{FF80}']),
        'チ' => Some(vec!['\u{FF81}']),
        'ツ' => Some(vec!['\u{FF82}']),
        'テ' => Some(vec!['\u{FF83}']),
        'ト' => Some(vec!['\u{FF84}']),
        'ナ' => Some(vec!['\u{FF85}']),
        'ニ' => Some(vec!['\u{FF86}']),
        'ヌ' => Some(vec!['\u{FF87}']),
        'ネ' => Some(vec!['\u{FF88}']),
        'ノ' => Some(vec!['\u{FF89}']),
        'ハ' => Some(vec!['\u{FF8A}']),
        'ヒ' => Some(vec!['\u{FF8B}']),
        'フ' => Some(vec!['\u{FF8C}']),
        'ヘ' => Some(vec!['\u{FF8D}']),
        'ホ' => Some(vec!['\u{FF8E}']),
        'マ' => Some(vec!['\u{FF8F}']),
        'ミ' => Some(vec!['\u{FF90}']),
        'ム' => Some(vec!['\u{FF91}']),
        'メ' => Some(vec!['\u{FF92}']),
        'モ' => Some(vec!['\u{FF93}']),
        'ヤ' => Some(vec!['\u{FF94}']),
        'ユ' => Some(vec!['\u{FF95}']),
        'ヨ' => Some(vec!['\u{FF96}']),
        'ラ' => Some(vec!['\u{FF97}']),
        'リ' => Some(vec!['\u{FF98}']),
        'ル' => Some(vec!['\u{FF99}']),
        'レ' => Some(vec!['\u{FF9A}']),
        'ロ' => Some(vec!['\u{FF9B}']),
        'ワ' => Some(vec!['\u{FF9C}']),
        'ヲ' => Some(vec!['\u{FF66}']),
        'ン' => Some(vec!['\u{FF9D}']),
        // Small katakana
        'ァ' => Some(vec!['\u{FF67}']),
        'ィ' => Some(vec!['\u{FF68}']),
        'ゥ' => Some(vec!['\u{FF69}']),
        'ェ' => Some(vec!['\u{FF6A}']),
        'ォ' => Some(vec!['\u{FF6B}']),
        'ッ' => Some(vec!['\u{FF6F}']),
        'ャ' => Some(vec!['\u{FF6C}']),
        'ュ' => Some(vec!['\u{FF6D}']),
        'ョ' => Some(vec!['\u{FF6E}']),
        // Punctuation
        '。' => Some(vec!['\u{FF61}']),
        '「' => Some(vec!['\u{FF62}']),
        '」' => Some(vec!['\u{FF63}']),
        '、' => Some(vec!['\u{FF64}']),
        '・' => Some(vec!['\u{FF65}']),
        'ー' => Some(vec!['\u{FF70}']),
        '゛' => Some(vec!['\u{FF9E}']),
        '゜' => Some(vec!['\u{FF9F}']),
        _ => None,
    }
}

/// Map half-width character(s) to a full-width character.
/// `next` is the following char for checking dakuten/handakuten combining marks.
/// Returns (full-width char, whether `next` was consumed).
fn halfwidth_to_fullwidth(ch: char, next: Option<char>) -> (Option<char>, bool) {
    // Half-width ASCII (U+0021..=U+007E) → Full-width (U+FF01..=U+FF5E)
    if ('!'..='~').contains(&ch) {
        let fw = char::from_u32(ch as u32 - 0x0021 + 0xFF01).unwrap();
        return (Some(fw), false);
    }
    // ASCII space → full-width space
    if ch == ' ' {
        return (Some('\u{3000}'), false);
    }
    let is_dakuten = next == Some('\u{FF9E}');
    let is_handakuten = next == Some('\u{FF9F}');
    match ch {
        // Half-width katakana with potential combining marks
        '\u{FF76}' if is_dakuten => (Some('ガ'), true),
        '\u{FF77}' if is_dakuten => (Some('ギ'), true),
        '\u{FF78}' if is_dakuten => (Some('グ'), true),
        '\u{FF79}' if is_dakuten => (Some('ゲ'), true),
        '\u{FF7A}' if is_dakuten => (Some('ゴ'), true),
        '\u{FF7B}' if is_dakuten => (Some('ザ'), true),
        '\u{FF7C}' if is_dakuten => (Some('ジ'), true),
        '\u{FF7D}' if is_dakuten => (Some('ズ'), true),
        '\u{FF7E}' if is_dakuten => (Some('ゼ'), true),
        '\u{FF7F}' if is_dakuten => (Some('ゾ'), true),
        '\u{FF80}' if is_dakuten => (Some('ダ'), true),
        '\u{FF81}' if is_dakuten => (Some('ヂ'), true),
        '\u{FF82}' if is_dakuten => (Some('ヅ'), true),
        '\u{FF83}' if is_dakuten => (Some('デ'), true),
        '\u{FF84}' if is_dakuten => (Some('ド'), true),
        '\u{FF8A}' if is_dakuten => (Some('バ'), true),
        '\u{FF8B}' if is_dakuten => (Some('ビ'), true),
        '\u{FF8C}' if is_dakuten => (Some('ブ'), true),
        '\u{FF8D}' if is_dakuten => (Some('ベ'), true),
        '\u{FF8E}' if is_dakuten => (Some('ボ'), true),
        '\u{FF73}' if is_dakuten => (Some('ヴ'), true),
        '\u{FF8A}' if is_handakuten => (Some('パ'), true),
        '\u{FF8B}' if is_handakuten => (Some('ピ'), true),
        '\u{FF8C}' if is_handakuten => (Some('プ'), true),
        '\u{FF8D}' if is_handakuten => (Some('ペ'), true),
        '\u{FF8E}' if is_handakuten => (Some('ポ'), true),
        // Unvoiced half-width katakana
        '\u{FF71}' => (Some('ア'), false),
        '\u{FF72}' => (Some('イ'), false),
        '\u{FF73}' => (Some('ウ'), false),
        '\u{FF74}' => (Some('エ'), false),
        '\u{FF75}' => (Some('オ'), false),
        '\u{FF76}' => (Some('カ'), false),
        '\u{FF77}' => (Some('キ'), false),
        '\u{FF78}' => (Some('ク'), false),
        '\u{FF79}' => (Some('ケ'), false),
        '\u{FF7A}' => (Some('コ'), false),
        '\u{FF7B}' => (Some('サ'), false),
        '\u{FF7C}' => (Some('シ'), false),
        '\u{FF7D}' => (Some('ス'), false),
        '\u{FF7E}' => (Some('セ'), false),
        '\u{FF7F}' => (Some('ソ'), false),
        '\u{FF80}' => (Some('タ'), false),
        '\u{FF81}' => (Some('チ'), false),
        '\u{FF82}' => (Some('ツ'), false),
        '\u{FF83}' => (Some('テ'), false),
        '\u{FF84}' => (Some('ト'), false),
        '\u{FF85}' => (Some('ナ'), false),
        '\u{FF86}' => (Some('ニ'), false),
        '\u{FF87}' => (Some('ヌ'), false),
        '\u{FF88}' => (Some('ネ'), false),
        '\u{FF89}' => (Some('ノ'), false),
        '\u{FF8A}' => (Some('ハ'), false),
        '\u{FF8B}' => (Some('ヒ'), false),
        '\u{FF8C}' => (Some('フ'), false),
        '\u{FF8D}' => (Some('ヘ'), false),
        '\u{FF8E}' => (Some('ホ'), false),
        '\u{FF8F}' => (Some('マ'), false),
        '\u{FF90}' => (Some('ミ'), false),
        '\u{FF91}' => (Some('ム'), false),
        '\u{FF92}' => (Some('メ'), false),
        '\u{FF93}' => (Some('モ'), false),
        '\u{FF94}' => (Some('ヤ'), false),
        '\u{FF95}' => (Some('ユ'), false),
        '\u{FF96}' => (Some('ヨ'), false),
        '\u{FF97}' => (Some('ラ'), false),
        '\u{FF98}' => (Some('リ'), false),
        '\u{FF99}' => (Some('ル'), false),
        '\u{FF9A}' => (Some('レ'), false),
        '\u{FF9B}' => (Some('ロ'), false),
        '\u{FF9C}' => (Some('ワ'), false),
        '\u{FF66}' => (Some('ヲ'), false),
        '\u{FF9D}' => (Some('ン'), false),
        // Small half-width katakana
        '\u{FF67}' => (Some('ァ'), false),
        '\u{FF68}' => (Some('ィ'), false),
        '\u{FF69}' => (Some('ゥ'), false),
        '\u{FF6A}' => (Some('ェ'), false),
        '\u{FF6B}' => (Some('ォ'), false),
        '\u{FF6F}' => (Some('ッ'), false),
        '\u{FF6C}' => (Some('ャ'), false),
        '\u{FF6D}' => (Some('ュ'), false),
        '\u{FF6E}' => (Some('ョ'), false),
        // Punctuation
        '\u{FF61}' => (Some('。'), false),
        '\u{FF62}' => (Some('「'), false),
        '\u{FF63}' => (Some('」'), false),
        '\u{FF64}' => (Some('、'), false),
        '\u{FF65}' => (Some('・'), false),
        '\u{FF70}' => (Some('ー'), false),
        '\u{FF9E}' => (Some('゛'), false),
        '\u{FF9F}' => (Some('゜'), false),
        _ => (None, false),
    }
}

/// ASC(text) — Convert full-width characters to half-width.
pub fn fn_asc(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match scalar_string(args.get(0).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let mut out = String::new();
    for ch in text.chars() {
        match fullwidth_to_halfwidth(ch) {
            Some(chars) => {
                for c in chars {
                    out.push(c);
                }
            }
            None => out.push(ch),
        }
    }
    Ok(FormulaValue::String(out))
}

/// JIS(text) — Convert half-width characters to full-width.
pub fn fn_jis(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let text = match scalar_string(args.get(0).unwrap()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let chars: Vec<char> = text.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < chars.len() {
        let next = chars.get(i + 1).copied();
        let (fw, consumed) = halfwidth_to_fullwidth(chars[i], next);
        match fw {
            Some(c) => {
                out.push(c);
                if consumed {
                    i += 1;
                }
            }
            None => out.push(chars[i]),
        }
        i += 1;
    }
    Ok(FormulaValue::String(out))
}

/// DBCS(text) — Alias for JIS. Convert half-width characters to full-width.
pub fn fn_dbcs(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    fn_jis(args, _ctx)
}

// ---------- BAHTTEXT ----------

/// Thai digit words.
const THAI_DIGITS: [&str; 10] = [
    "ศูนย์", // 0
    "หนึ่ง", // 1
    "สอง", // 2
    "สาม", // 3
    "สี่",   // 4
    "ห้า",  // 5
    "หก",  // 6
    "เจ็ด", // 7
    "แปด", // 8
    "เก้า", // 9
];

/// Thai place value words.
const THAI_PLACES: [&str; 6] = [
    "",    // ones (no place word)
    "สิบ",  // tens
    "ร้อย", // hundreds
    "พัน",  // thousands
    "หมื่น", // ten-thousands
    "แสน", // hundred-thousands
];

/// Convert an integer (0..999_999) to Thai text for one group of up to 6 digits.
fn thai_group(mut n: u64) -> String {
    if n == 0 {
        return String::new();
    }
    let mut parts = Vec::new();
    let mut digits = Vec::new();
    while n > 0 {
        digits.push((n % 10) as usize);
        n /= 10;
    }
    // digits[0] = ones, digits[1] = tens, etc.
    for (place, &digit) in digits.iter().enumerate().rev() {
        if digit == 0 {
            continue;
        }
        // Special case: 1 in tens place → just "สิบ" (not "หนึ่งสิบ")
        if place == 1 && digit == 1 {
            parts.push("สิบ".to_string());
            continue;
        }
        // Special case: 2 in tens place → "ยี่สิบ" (not "สองสิบ")
        if place == 1 && digit == 2 {
            parts.push("ยี่สิบ".to_string());
            continue;
        }
        // Special case: 1 in ones place and there are higher digits → "เอ็ด" (not "หนึ่ง")
        if place == 0 && digit == 1 && digits.len() > 1 {
            parts.push("เอ็ด".to_string());
            continue;
        }
        let mut s = THAI_DIGITS[digit].to_string();
        s.push_str(THAI_PLACES[place]);
        parts.push(s);
    }
    parts.join("")
}

/// Convert a non-negative number to Thai Baht text.
fn number_to_bahttext(value: f64) -> String {
    // Round to 2 decimal places (satang)
    let rounded = (value * 100.0).round() as u64;
    let baht_part = rounded / 100;
    let satang_part = rounded % 100;

    // Build baht portion: split into groups of 6 digits (ล้าน = million)
    let baht_text = if baht_part == 0 {
        String::new()
    } else {
        let mut groups = Vec::new();
        let mut remaining = baht_part;
        while remaining > 0 {
            groups.push(remaining % 1_000_000);
            remaining /= 1_000_000;
        }
        let mut parts = Vec::new();
        for (i, &group) in groups.iter().enumerate().rev() {
            let group_text = thai_group(group);
            if !group_text.is_empty() {
                parts.push(group_text);
                // Append ล้าน for each group above the first
                for _ in 0..i {
                    parts.push("ล้าน".to_string());
                }
            }
        }
        parts.join("")
    };

    let satang_text = if satang_part == 0 {
        String::new()
    } else {
        thai_group(satang_part)
    };

    // Combine
    if baht_part == 0 && satang_part == 0 {
        "ศูนย์บาทถ้วน".to_string()
    } else if satang_part == 0 {
        format!("{}บาทถ้วน", baht_text)
    } else if baht_part == 0 {
        format!("{}สตางค์", satang_text)
    } else {
        format!("{}บาท{}สตางค์", baht_text, satang_text)
    }
}

/// BAHTTEXT(number) — Convert a number to Thai Baht text.
pub fn fn_bahttext(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let value = match args.get(0).unwrap() {
        FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
        v => match v.as_number() {
            Some(n) => n,
            None => return Ok(FormulaValue::Error(CellError::Value)),
        },
    };
    if value < 0.0 {
        // Excel returns "ลบ" prefix for negative
        let text = number_to_bahttext(value.abs());
        Ok(FormulaValue::String(format!("ลบ{}", text)))
    } else {
        Ok(FormulaValue::String(number_to_bahttext(value)))
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

    // ---------- ASC tests ----------

    #[test]
    fn test_asc_fullwidth_ascii() {
        // Full-width A (U+FF21) -> ASCII A
        assert_eq!(eval("=ASC(\"\u{FF21}\")").unwrap(), s("A"));
        // Full-width digits
        assert_eq!(
            eval("=ASC(\"\u{FF11}\u{FF12}\u{FF13}\")").unwrap(),
            s("123")
        );
        // Full-width punctuation
        assert_eq!(eval("=ASC(\"\u{FF01}\")").unwrap(), s("!"));
    }

    #[test]
    fn test_asc_fullwidth_space() {
        // Ideographic space (U+3000) -> ASCII space
        assert_eq!(eval("=ASC(\"\u{3000}\")").unwrap(), s(" "));
    }

    #[test]
    fn test_asc_fullwidth_katakana() {
        // Full-width katakana ア (U+30A2) -> half-width ｱ (U+FF71)
        assert_eq!(eval("=ASC(\"\u{30A2}\")").unwrap(), s("\u{FF71}"));
        // Full-width カ (U+30AB) -> half-width ｶ (U+FF76)
        assert_eq!(eval("=ASC(\"\u{30AB}\")").unwrap(), s("\u{FF76}"));
    }

    #[test]
    fn test_asc_voiced_katakana() {
        // Full-width ガ (U+30AC) -> half-width ｶﾞ (U+FF76 + U+FF9E)
        assert_eq!(eval("=ASC(\"\u{30AC}\")").unwrap(), s("\u{FF76}\u{FF9E}"));
    }

    #[test]
    fn test_asc_passthrough() {
        // ASCII and non-CJK pass through unchanged
        assert_eq!(eval("=ASC(\"Hello\")").unwrap(), s("Hello"));
        assert_eq!(eval("=ASC(\"\")").unwrap(), s(""));
    }

    #[test]
    fn test_asc_mixed() {
        // Mixed full-width ASCII and regular ASCII
        assert_eq!(eval("=ASC(\"\u{FF21}B\u{FF23}\")").unwrap(), s("ABC"));
    }

    // ---------- JIS tests ----------

    #[test]
    fn test_jis_halfwidth_ascii() {
        // ASCII A -> full-width Ａ (U+FF21)
        assert_eq!(eval("=JIS(\"A\")").unwrap(), s("\u{FF21}"));
        // ASCII digits
        assert_eq!(
            eval("=JIS(\"123\")").unwrap(),
            s("\u{FF11}\u{FF12}\u{FF13}")
        );
    }

    #[test]
    fn test_jis_space() {
        // ASCII space -> ideographic space (U+3000)
        assert_eq!(eval("=JIS(\" \")").unwrap(), s("\u{3000}"));
    }

    #[test]
    fn test_jis_halfwidth_katakana() {
        // Half-width ｱ (U+FF71) -> full-width ア (U+30A2)
        assert_eq!(eval("=JIS(\"\u{FF71}\")").unwrap(), s("\u{30A2}"));
    }

    #[test]
    fn test_jis_voiced_combining() {
        // Half-width ｶﾞ (U+FF76 + U+FF9E) -> full-width ガ (U+30AC)
        assert_eq!(eval("=JIS(\"\u{FF76}\u{FF9E}\")").unwrap(), s("\u{30AC}"));
    }

    #[test]
    fn test_jis_semi_voiced_combining() {
        // Half-width ﾊﾟ (U+FF8A + U+FF9F) -> full-width パ (U+30D1)
        assert_eq!(eval("=JIS(\"\u{FF8A}\u{FF9F}\")").unwrap(), s("\u{30D1}"));
    }

    #[test]
    fn test_jis_passthrough() {
        // Full-width katakana passes through unchanged
        assert_eq!(eval("=JIS(\"\u{30A2}\")").unwrap(), s("\u{30A2}"));
        assert_eq!(eval("=JIS(\"\")").unwrap(), s(""));
    }

    // ---------- DBCS tests ----------

    #[test]
    fn test_dbcs_same_as_jis() {
        // DBCS should behave identically to JIS
        assert_eq!(eval("=DBCS(\"A\")").unwrap(), s("\u{FF21}"));
        assert_eq!(
            eval("=DBCS(\"123\")").unwrap(),
            s("\u{FF11}\u{FF12}\u{FF13}")
        );
        assert_eq!(eval("=DBCS(\" \")").unwrap(), s("\u{3000}"));
    }

    // ---------- ASC/JIS roundtrip ----------

    #[test]
    fn test_asc_jis_roundtrip() {
        // JIS(ASC(text)) should return original for full-width content
        let original = "\u{FF21}\u{FF22}\u{FF23}"; // ＡＢＣ
        let asc_result = eval(&format!("=ASC(\"{}\")", original)).unwrap();
        assert_eq!(asc_result, s("ABC"));
        let jis_result = eval("=JIS(\"ABC\")").unwrap();
        assert_eq!(jis_result, s(original));
    }

    // ---------- BAHTTEXT tests ----------

    #[test]
    fn test_bahttext_zero() {
        // 0 -> "ศูนย์บาทถ้วน"
        assert_eq!(eval("=BAHTTEXT(0)").unwrap(), s("ศูนย์บาทถ้วน"));
    }

    #[test]
    fn test_bahttext_integer() {
        // 1 -> "หนึ่งบาทถ้วน"
        assert_eq!(eval("=BAHTTEXT(1)").unwrap(), s("หนึ่งบาทถ้วน"));
    }

    #[test]
    fn test_bahttext_tens() {
        // 10 -> "สิบบาทถ้วน" (not "หนึ่งสิบ")
        assert_eq!(eval("=BAHTTEXT(10)").unwrap(), s("สิบบาทถ้วน"));
    }

    #[test]
    fn test_bahttext_twenty() {
        // 20 -> "ยี่สิบบาทถ้วน" (not "สองสิบ")
        assert_eq!(eval("=BAHTTEXT(20)").unwrap(), s("ยี่สิบบาทถ้วน"));
    }

    #[test]
    fn test_bahttext_eleven() {
        // 11 -> "สิบเอ็ดบาทถ้วน" (not "สิบหนึ่ง")
        assert_eq!(eval("=BAHTTEXT(11)").unwrap(), s("สิบเอ็ดบาทถ้วน"));
    }

    #[test]
    fn test_bahttext_twenty_one() {
        // 21 -> "ยี่สิบเอ็ดบาทถ้วน"
        assert_eq!(eval("=BAHTTEXT(21)").unwrap(), s("ยี่สิบเอ็ดบาทถ้วน"));
    }

    #[test]
    fn test_bahttext_hundreds() {
        // 100 -> "หนึ่งร้อยบาทถ้วน"
        assert_eq!(eval("=BAHTTEXT(100)").unwrap(), s("หนึ่งร้อยบาทถ้วน"));
    }

    #[test]
    fn test_bahttext_with_satang() {
        // 1.50 -> "หนึ่งบาทห้าสิบสตางค์"
        assert_eq!(eval("=BAHTTEXT(1.5)").unwrap(), s("หนึ่งบาทห้าสิบสตางค์"));
    }

    #[test]
    fn test_bahttext_satang_only() {
        // 0.25 -> "ยี่สิบห้าสตางค์"
        assert_eq!(eval("=BAHTTEXT(0.25)").unwrap(), s("ยี่สิบห้าสตางค์"));
    }

    #[test]
    fn test_bahttext_negative() {
        // -1 -> "ลบหนึ่งบาทถ้วน"
        assert_eq!(eval("=BAHTTEXT(-1)").unwrap(), s("ลบหนึ่งบาทถ้วน"));
    }

    #[test]
    fn test_bahttext_million() {
        // 1000000 -> "หนึ่งล้านบาทถ้วน"
        assert_eq!(eval("=BAHTTEXT(1000000)").unwrap(), s("หนึ่งล้านบาทถ้วน"));
    }

    #[test]
    fn test_bahttext_complex() {
        // 5678 -> "ห้าพันหกร้อยเจ็ดสิบแปดบาทถ้วน"
        assert_eq!(
            eval("=BAHTTEXT(5678)").unwrap(),
            s("ห้าพันหกร้อยเจ็ดสิบแปดบาทถ้วน")
        );
    }
}
