//! Tests for the format code lexer.

use ssfmt::parser::lexer::Lexer;
use ssfmt::parser::tokens::Token;

#[test]
fn test_lex_simple_number_format() {
    let mut lexer = Lexer::new("#,##0.00");
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::ThousandsSep);
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::DecimalPoint);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_date_format() {
    let mut lexer = Lexer::new("yyyy-mm-dd");
    assert_eq!(lexer.next_token().unwrap().token, Token::Year);
    assert_eq!(lexer.next_token().unwrap().token, Token::Year);
    assert_eq!(lexer.next_token().unwrap().token, Token::Year);
    assert_eq!(lexer.next_token().unwrap().token, Token::Year);
    // Note: '-' is lexed as Minus; the parser determines if it's a literal separator
    assert_eq!(lexer.next_token().unwrap().token, Token::Minus);
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Minus);
    assert_eq!(lexer.next_token().unwrap().token, Token::Day);
    assert_eq!(lexer.next_token().unwrap().token, Token::Day);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_quoted_string() {
    let mut lexer = Lexer::new("\"USD\"0.00");
    assert_eq!(
        lexer.next_token().unwrap().token,
        Token::QuotedString("USD".into())
    );
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::DecimalPoint);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_escaped_char() {
    let mut lexer = Lexer::new("\\$0.00");
    assert_eq!(lexer.next_token().unwrap().token, Token::EscapedChar('$'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::DecimalPoint);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_bracket() {
    let mut lexer = Lexer::new("[Red]0");
    assert_eq!(lexer.next_token().unwrap().token, Token::OpenBracket);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('R'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('e'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('d'));
    assert_eq!(lexer.next_token().unwrap().token, Token::CloseBracket);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_sections() {
    let mut lexer = Lexer::new("0;-0");
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::SectionSep);
    assert_eq!(lexer.next_token().unwrap().token, Token::Minus);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_time_format() {
    let mut lexer = Lexer::new("hh:mm:ss");
    assert_eq!(lexer.next_token().unwrap().token, Token::Hour);
    assert_eq!(lexer.next_token().unwrap().token, Token::Hour);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(':'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Month); // mm is Month in non-time context, but will be resolved by parser
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(':'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Second);
    assert_eq!(lexer.next_token().unwrap().token, Token::Second);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_percent() {
    let mut lexer = Lexer::new("0.00%");
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::DecimalPoint);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Percent);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_scientific_notation() {
    let mut lexer = Lexer::new("0.00E+00");
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::DecimalPoint);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::ExponentUpper);
    assert_eq!(lexer.next_token().unwrap().token, Token::Plus);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_fraction() {
    let mut lexer = Lexer::new("# ?/?");
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(' '));
    assert_eq!(lexer.next_token().unwrap().token, Token::Question);
    assert_eq!(lexer.next_token().unwrap().token, Token::Slash);
    assert_eq!(lexer.next_token().unwrap().token, Token::Question);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_text_placeholder() {
    let mut lexer = Lexer::new("@");
    assert_eq!(lexer.next_token().unwrap().token, Token::At);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_fill_character() {
    let mut lexer = Lexer::new("*-0");
    assert_eq!(lexer.next_token().unwrap().token, Token::Asterisk);
    assert_eq!(lexer.next_token().unwrap().token, Token::Minus);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_skip_character() {
    let mut lexer = Lexer::new("_)0");
    assert_eq!(lexer.next_token().unwrap().token, Token::Underscore);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(')'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_am_pm() {
    let mut lexer = Lexer::new("h:mm AM/PM");
    assert_eq!(lexer.next_token().unwrap().token, Token::Hour);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(':'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(' '));
    assert_eq!(
        lexer.next_token().unwrap().token,
        Token::AmPm("AM/PM".into())
    );
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_am_pm_lowercase() {
    let mut lexer = Lexer::new("h:mm am/pm");
    assert_eq!(lexer.next_token().unwrap().token, Token::Hour);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(':'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(' '));
    assert_eq!(
        lexer.next_token().unwrap().token,
        Token::AmPm("am/pm".into())
    );
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_am_pm_short() {
    let mut lexer = Lexer::new("h:mm A/P");
    assert_eq!(lexer.next_token().unwrap().token, Token::Hour);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(':'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(' '));
    assert_eq!(lexer.next_token().unwrap().token, Token::AmPm("A/P".into()));
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_elapsed_hours_in_bracket() {
    // Inside brackets, h is treated as literal
    let mut lexer = Lexer::new("[h]");
    assert_eq!(lexer.next_token().unwrap().token, Token::OpenBracket);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('h'));
    assert_eq!(lexer.next_token().unwrap().token, Token::CloseBracket);
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_unterminated_quote() {
    let mut lexer = Lexer::new("\"USD");
    let result = lexer.next_token();
    assert!(result.is_err());
}

#[test]
fn test_lex_empty_input() {
    let mut lexer = Lexer::new("");
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_complex_format() {
    // Test a complex format: #,##0.00_);[Red](#,##0.00)
    let mut lexer = Lexer::new("#,##0.00_);[Red](#,##0.00)");

    // First section: #,##0.00_)
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::ThousandsSep);
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::DecimalPoint);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Underscore);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(')'));

    // Section separator
    assert_eq!(lexer.next_token().unwrap().token, Token::SectionSep);

    // Second section: [Red](#,##0.00)
    assert_eq!(lexer.next_token().unwrap().token, Token::OpenBracket);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('R'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('e'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('d'));
    assert_eq!(lexer.next_token().unwrap().token, Token::CloseBracket);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('('));
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::ThousandsSep);
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::Hash);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::DecimalPoint);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal(')'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Eof);
}

#[test]
fn test_lex_position_tracking() {
    let mut lexer = Lexer::new("0.00");
    let tok1 = lexer.next_token().unwrap();
    assert_eq!(tok1.start, 0);
    assert_eq!(tok1.end, 1);

    let tok2 = lexer.next_token().unwrap();
    assert_eq!(tok2.start, 1);
    assert_eq!(tok2.end, 2);

    let tok3 = lexer.next_token().unwrap();
    assert_eq!(tok3.start, 2);
    assert_eq!(tok3.end, 3);
}

#[test]
fn test_lex_quoted_string_position() {
    let mut lexer = Lexer::new("\"AB\"0");
    let tok1 = lexer.next_token().unwrap();
    assert_eq!(tok1.token, Token::QuotedString("AB".into()));
    assert_eq!(tok1.start, 0);
    assert_eq!(tok1.end, 4); // Includes the quotes

    let tok2 = lexer.next_token().unwrap();
    assert_eq!(tok2.token, Token::Zero);
    assert_eq!(tok2.start, 4);
    assert_eq!(tok2.end, 5);
}
