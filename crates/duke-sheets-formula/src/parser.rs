//! Formula parser
//!
//! A recursive descent parser for Excel formulas with proper operator precedence.

use crate::ast::{BinaryOperator, CellReference, FormulaExpr, RangeReference, UnaryOperator};
use crate::error::{FormulaError, FormulaResult};
use duke_sheets_core::{CellAddress, CellError, CellRange};

/// Parse a formula string into an AST
///
/// # Example
/// ```rust
/// use duke_sheets_formula::parse_formula;
///
/// let ast = parse_formula("=1+2").unwrap();
/// let ast = parse_formula("=SUM(A1:A10)").unwrap();
/// let ast = parse_formula("=IF(A1>0,\"Yes\",\"No\")").unwrap();
/// ```
pub fn parse_formula(formula: &str) -> FormulaResult<FormulaExpr> {
    let formula = formula.trim();

    // Formula must start with '='
    let formula = formula
        .strip_prefix('=')
        .ok_or_else(|| FormulaError::Parse("Formula must start with '='".into()))?;

    let mut parser = FormulaParser::new(formula);
    let expr = parser.parse_expression()?;

    // Make sure we consumed all input
    parser.skip_whitespace();
    if !parser.is_at_end() {
        return Err(FormulaError::Parse(format!(
            "Unexpected characters after expression: '{}'",
            &parser.input[parser.pos..]
        )));
    }

    Ok(expr)
}

/// Token types
#[derive(Debug, Clone, PartialEq)]
enum Token {
    // Literals
    Number(f64),
    String(String),
    Boolean(bool),
    Error(CellError),

    // Identifiers and references
    Identifier(String), // Function name or named range
    CellRef(String),    // Cell reference like A1, $A$1
    SheetRef(String),   // Sheet reference like Sheet1!

    // Operators
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Percent,
    Ampersand,
    Equal,
    NotEqual,
    LessThan,
    LessEqual,
    GreaterThan,
    GreaterEqual,
    Colon,
    Comma,
    Semicolon,

    // Delimiters
    LeftParen,
    RightParen,
    LeftBrace,
    RightBrace,

    // End of input
    Eof,
}

/// Formula parser
struct FormulaParser<'a> {
    input: &'a str,
    pos: usize,
    current_token: Option<Token>,
}

impl<'a> FormulaParser<'a> {
    fn new(input: &'a str) -> Self {
        let mut parser = Self {
            input,
            pos: 0,
            current_token: None,
        };
        parser.advance_token();
        parser
    }

    // === Token scanning ===

    fn advance_token(&mut self) {
        self.skip_whitespace();
        self.current_token = Some(self.scan_token());
    }

    fn scan_token(&mut self) -> Token {
        self.skip_whitespace();

        if self.is_at_end() {
            return Token::Eof;
        }

        let c = self.peek_char().unwrap();

        // Single-character tokens
        match c {
            '+' => {
                self.advance();
                return Token::Plus;
            }
            '-' => {
                self.advance();
                return Token::Minus;
            }
            '*' => {
                self.advance();
                return Token::Star;
            }
            '/' => {
                self.advance();
                return Token::Slash;
            }
            '^' => {
                self.advance();
                return Token::Caret;
            }
            '%' => {
                self.advance();
                return Token::Percent;
            }
            '&' => {
                self.advance();
                return Token::Ampersand;
            }
            ':' => {
                self.advance();
                return Token::Colon;
            }
            ',' => {
                self.advance();
                return Token::Comma;
            }
            ';' => {
                self.advance();
                return Token::Semicolon;
            }
            '(' => {
                self.advance();
                return Token::LeftParen;
            }
            ')' => {
                self.advance();
                return Token::RightParen;
            }
            '{' => {
                self.advance();
                return Token::LeftBrace;
            }
            '}' => {
                self.advance();
                return Token::RightBrace;
            }
            _ => {}
        }

        // Two-character operators
        if c == '<' {
            self.advance();
            if self.peek_char() == Some('=') {
                self.advance();
                return Token::LessEqual;
            } else if self.peek_char() == Some('>') {
                self.advance();
                return Token::NotEqual;
            }
            return Token::LessThan;
        }

        if c == '>' {
            self.advance();
            if self.peek_char() == Some('=') {
                self.advance();
                return Token::GreaterEqual;
            }
            return Token::GreaterThan;
        }

        if c == '=' {
            self.advance();
            return Token::Equal;
        }

        // String literal
        if c == '"' {
            return self.scan_string();
        }

        // Number
        if c.is_ascii_digit()
            || (c == '.' && self.peek_char_at(1).map_or(false, |c| c.is_ascii_digit()))
        {
            return self.scan_number();
        }

        // Identifier, cell reference, or boolean/error
        if c.is_ascii_alphabetic() || c == '_' || c == '$' || c == '#' {
            return self.scan_identifier_or_ref();
        }

        // Unknown character
        self.advance();
        Token::Eof
    }

    fn scan_string(&mut self) -> Token {
        self.advance(); // Skip opening quote

        let mut s = String::new();
        while let Some(c) = self.peek_char() {
            if c == '"' {
                // Check for escaped quote ("")
                if self.peek_char_at(1) == Some('"') {
                    s.push('"');
                    self.advance();
                    self.advance();
                } else {
                    break;
                }
            } else {
                s.push(c);
                self.advance();
            }
        }

        // Skip closing quote
        if self.peek_char() == Some('"') {
            self.advance();
        }

        Token::String(s)
    }

    fn scan_number(&mut self) -> Token {
        let start = self.pos;

        // Integer part
        while self.peek_char().map_or(false, |c| c.is_ascii_digit()) {
            self.advance();
        }

        // Decimal part
        if self.peek_char() == Some('.') {
            self.advance();
            while self.peek_char().map_or(false, |c| c.is_ascii_digit()) {
                self.advance();
            }
        }

        // Exponent part
        if self.peek_char().map_or(false, |c| c == 'e' || c == 'E') {
            self.advance();
            if self.peek_char().map_or(false, |c| c == '+' || c == '-') {
                self.advance();
            }
            while self.peek_char().map_or(false, |c| c.is_ascii_digit()) {
                self.advance();
            }
        }

        let num_str = &self.input[start..self.pos];
        let num: f64 = num_str.parse().unwrap_or(0.0);
        Token::Number(num)
    }

    fn scan_identifier_or_ref(&mut self) -> Token {
        // Check for error values first (#VALUE!, #REF!, etc.)
        if self.peek_char() == Some('#') {
            let start = self.pos;
            self.advance();
            while self.peek_char().map_or(false, |c| {
                c.is_ascii_alphanumeric() || c == '!' || c == '/' || c == '?'
            }) {
                self.advance();
            }
            let error_str = &self.input[start..self.pos];
            if let Some(err) = CellError::from_str(error_str) {
                return Token::Error(err);
            }
            // If not a valid error, treat as identifier
            return Token::Identifier(error_str.to_string());
        }

        let start = self.pos;

        // Scan identifier/reference
        while self.peek_char().map_or(false, |c| {
            c.is_ascii_alphanumeric() || c == '_' || c == '$' || c == '.'
        }) {
            self.advance();
        }

        let text = &self.input[start..self.pos];

        // Check for sheet reference (ends with !)
        if self.peek_char() == Some('!') {
            self.advance();
            let sheet_name = text.trim_matches('\'').to_string();
            return Token::SheetRef(sheet_name);
        }

        // Check for boolean literals (but not if followed by '(' - then it's a function call)
        let upper = text.to_uppercase();
        if upper == "TRUE" && self.peek_char() != Some('(') {
            return Token::Boolean(true);
        }
        if upper == "FALSE" && self.peek_char() != Some('(') {
            return Token::Boolean(false);
        }

        // Check if it looks like a cell reference (letter(s) followed by number(s))
        // BUT if followed by '(' it's a function call (e.g., LOG10(100) is function, not cell ref)
        if Self::is_cell_reference(text) && self.peek_char() != Some('(') {
            return Token::CellRef(text.to_string());
        }

        // Otherwise it's an identifier (function name or named range)
        Token::Identifier(text.to_string())
    }

    fn is_cell_reference(text: &str) -> bool {
        // Cell reference pattern: [$]A-XFD[$]1-1048576
        // Simplified check: starts with optional $, then letters, then optional $, then digits
        let chars: Vec<char> = text.chars().collect();
        let mut i = 0;

        // Skip leading $
        if chars.get(i) == Some(&'$') {
            i += 1;
        }

        // Must have letters
        let letter_start = i;
        while i < chars.len() && chars[i].is_ascii_alphabetic() {
            i += 1;
        }
        if i == letter_start {
            return false;
        }

        // Skip optional $
        if chars.get(i) == Some(&'$') {
            i += 1;
        }

        // Must have digits
        let digit_start = i;
        while i < chars.len() && chars[i].is_ascii_digit() {
            i += 1;
        }
        if i == digit_start {
            return false;
        }

        // Must have consumed everything
        i == chars.len()
    }

    // === Helper methods ===

    fn peek_char(&self) -> Option<char> {
        self.input[self.pos..].chars().next()
    }

    fn peek_char_at(&self, offset: usize) -> Option<char> {
        self.input[self.pos..].chars().nth(offset)
    }

    fn advance(&mut self) {
        if let Some(c) = self.peek_char() {
            self.pos += c.len_utf8();
        }
    }

    fn skip_whitespace(&mut self) {
        while self.peek_char().map_or(false, |c| c.is_whitespace()) {
            self.advance();
        }
    }

    fn is_at_end(&self) -> bool {
        self.pos >= self.input.len()
    }

    fn current_token(&self) -> &Token {
        self.current_token.as_ref().unwrap_or(&Token::Eof)
    }

    fn consume(&mut self) -> Token {
        let token = self.current_token.take().unwrap_or(Token::Eof);
        self.advance_token();
        token
    }

    fn expect(&mut self, expected: &Token) -> FormulaResult<()> {
        if self.current_token() == expected {
            self.consume();
            Ok(())
        } else {
            Err(FormulaError::Parse(format!(
                "Expected {:?}, got {:?}",
                expected,
                self.current_token()
            )))
        }
    }

    // === Expression parsing with precedence ===
    // Precedence (lowest to highest):
    // 1. Comparison: =, <>, <, <=, >, >=
    // 2. Concatenation: &
    // 3. Addition/Subtraction: +, -
    // 4. Multiplication/Division: *, /
    // 5. Exponentiation: ^
    // 6. Unary: -, %
    // 7. Range/Union: :, , (space)
    // 8. Primary: literals, references, function calls, parentheses

    fn parse_expression(&mut self) -> FormulaResult<FormulaExpr> {
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> FormulaResult<FormulaExpr> {
        let mut left = self.parse_concatenation()?;

        loop {
            let op = match self.current_token() {
                Token::Equal => BinaryOperator::Equal,
                Token::NotEqual => BinaryOperator::NotEqual,
                Token::LessThan => BinaryOperator::LessThan,
                Token::LessEqual => BinaryOperator::LessEqual,
                Token::GreaterThan => BinaryOperator::GreaterThan,
                Token::GreaterEqual => BinaryOperator::GreaterEqual,
                _ => break,
            };

            self.consume();
            let right = self.parse_concatenation()?;
            left = FormulaExpr::BinaryOp {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }

        Ok(left)
    }

    fn parse_concatenation(&mut self) -> FormulaResult<FormulaExpr> {
        let mut left = self.parse_additive()?;

        while matches!(self.current_token(), Token::Ampersand) {
            self.consume();
            let right = self.parse_additive()?;
            left = FormulaExpr::BinaryOp {
                op: BinaryOperator::Concat,
                left: Box::new(left),
                right: Box::new(right),
            };
        }

        Ok(left)
    }

    fn parse_additive(&mut self) -> FormulaResult<FormulaExpr> {
        let mut left = self.parse_multiplicative()?;

        loop {
            let op = match self.current_token() {
                Token::Plus => BinaryOperator::Add,
                Token::Minus => BinaryOperator::Subtract,
                _ => break,
            };

            self.consume();
            let right = self.parse_multiplicative()?;
            left = FormulaExpr::BinaryOp {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }

        Ok(left)
    }

    fn parse_multiplicative(&mut self) -> FormulaResult<FormulaExpr> {
        let mut left = self.parse_exponent()?;

        loop {
            let op = match self.current_token() {
                Token::Star => BinaryOperator::Multiply,
                Token::Slash => BinaryOperator::Divide,
                _ => break,
            };

            self.consume();
            let right = self.parse_exponent()?;
            left = FormulaExpr::BinaryOp {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }

        Ok(left)
    }

    fn parse_exponent(&mut self) -> FormulaResult<FormulaExpr> {
        let left = self.parse_unary()?;

        if matches!(self.current_token(), Token::Caret) {
            self.consume();
            let right = self.parse_exponent()?; // Right associative
            return Ok(FormulaExpr::BinaryOp {
                op: BinaryOperator::Power,
                left: Box::new(left),
                right: Box::new(right),
            });
        }

        Ok(left)
    }

    fn parse_unary(&mut self) -> FormulaResult<FormulaExpr> {
        // Prefix unary minus
        if matches!(self.current_token(), Token::Minus) {
            self.consume();
            let operand = self.parse_unary()?;
            return Ok(FormulaExpr::UnaryOp {
                op: UnaryOperator::Negate,
                operand: Box::new(operand),
            });
        }

        // Prefix plus (no-op)
        if matches!(self.current_token(), Token::Plus) {
            self.consume();
            return self.parse_unary();
        }

        // Parse primary, then check for postfix percent
        let mut expr = self.parse_range()?;

        while matches!(self.current_token(), Token::Percent) {
            self.consume();
            expr = FormulaExpr::UnaryOp {
                op: UnaryOperator::Percent,
                operand: Box::new(expr),
            };
        }

        Ok(expr)
    }

    fn parse_range(&mut self) -> FormulaResult<FormulaExpr> {
        let left = self.parse_primary()?;

        // Check for range operator (:)
        if matches!(self.current_token(), Token::Colon) {
            self.consume();
            let right = self.parse_primary()?;

            // Try to convert to a RangeRef if both are cell references
            if let (FormulaExpr::CellRef(start_ref), FormulaExpr::CellRef(end_ref)) =
                (&left, &right)
            {
                // Make sure sheets match
                if start_ref.sheet != end_ref.sheet {
                    return Err(FormulaError::Parse(
                        "Range references must be on the same sheet".into(),
                    ));
                }

                let range = CellRange::new(start_ref.address, end_ref.address);
                return Ok(FormulaExpr::RangeRef(RangeReference {
                    sheet: start_ref.sheet.clone(),
                    range,
                }));
            }

            return Ok(FormulaExpr::BinaryOp {
                op: BinaryOperator::Range,
                left: Box::new(left),
                right: Box::new(right),
            });
        }

        Ok(left)
    }

    fn parse_primary(&mut self) -> FormulaResult<FormulaExpr> {
        match self.current_token().clone() {
            Token::Number(n) => {
                self.consume();
                Ok(FormulaExpr::Number(n))
            }

            Token::String(s) => {
                self.consume();
                Ok(FormulaExpr::String(s))
            }

            Token::Boolean(b) => {
                self.consume();
                Ok(FormulaExpr::Boolean(b))
            }

            Token::Error(e) => {
                self.consume();
                Ok(FormulaExpr::Error(e))
            }

            Token::LeftParen => {
                self.consume();
                let expr = self.parse_expression()?;
                self.expect(&Token::RightParen)?;
                Ok(expr)
            }

            Token::LeftBrace => self.parse_array(),

            Token::SheetRef(sheet) => {
                self.consume();
                self.parse_sheet_reference(sheet)
            }

            Token::CellRef(ref_str) => {
                self.consume();
                self.parse_cell_reference(None, &ref_str)
            }

            Token::Identifier(name) => {
                self.consume();
                // Check if it's a function call
                if matches!(self.current_token(), Token::LeftParen) {
                    self.parse_function_call(name)
                } else {
                    // Named range
                    Ok(FormulaExpr::NameRef(name))
                }
            }

            _ => Err(FormulaError::Parse(format!(
                "Unexpected token: {:?}",
                self.current_token()
            ))),
        }
    }

    fn parse_array(&mut self) -> FormulaResult<FormulaExpr> {
        self.expect(&Token::LeftBrace)?;

        let mut rows = Vec::new();
        let mut current_row = Vec::new();

        // Parse first element
        if !matches!(self.current_token(), Token::RightBrace) {
            current_row.push(self.parse_expression()?);

            loop {
                match self.current_token() {
                    Token::Comma => {
                        self.consume();
                        current_row.push(self.parse_expression()?);
                    }
                    Token::Semicolon => {
                        self.consume();
                        rows.push(current_row);
                        current_row = Vec::new();
                        current_row.push(self.parse_expression()?);
                    }
                    Token::RightBrace => break,
                    _ => {
                        return Err(FormulaError::Parse(
                            "Expected ',' ';' or '}' in array".into(),
                        ))
                    }
                }
            }
        }

        if !current_row.is_empty() {
            rows.push(current_row);
        }

        self.expect(&Token::RightBrace)?;
        Ok(FormulaExpr::Array(rows))
    }

    fn parse_function_call(&mut self, name: String) -> FormulaResult<FormulaExpr> {
        self.expect(&Token::LeftParen)?;

        let mut args = Vec::new();

        // Parse arguments
        if !matches!(self.current_token(), Token::RightParen) {
            args.push(self.parse_expression()?);

            while matches!(self.current_token(), Token::Comma) {
                self.consume();
                args.push(self.parse_expression()?);
            }
        }

        self.expect(&Token::RightParen)?;

        Ok(FormulaExpr::Function {
            name: name.to_uppercase(),
            args,
        })
    }

    fn parse_sheet_reference(&mut self, sheet: String) -> FormulaResult<FormulaExpr> {
        // After Sheet1!, we expect a cell reference
        match self.current_token().clone() {
            Token::CellRef(ref_str) => {
                self.consume();
                self.parse_cell_reference(Some(sheet), &ref_str)
            }
            _ => Err(FormulaError::Parse(
                "Expected cell reference after sheet name".into(),
            )),
        }
    }

    fn parse_cell_reference(
        &mut self,
        sheet: Option<String>,
        ref_str: &str,
    ) -> FormulaResult<FormulaExpr> {
        // Parse the cell address, stripping $ signs
        let clean_ref = ref_str.replace('$', "");
        let address = CellAddress::parse(&clean_ref).map_err(|e| {
            FormulaError::Parse(format!("Invalid cell reference '{}': {}", ref_str, e))
        })?;

        Ok(FormulaExpr::CellRef(CellReference { sheet, address }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_number() {
        let ast = parse_formula("=42").unwrap();
        assert_eq!(ast, FormulaExpr::Number(42.0));

        let ast = parse_formula("=3.14").unwrap();
        assert_eq!(ast, FormulaExpr::Number(3.14));

        let ast = parse_formula("=1e10").unwrap();
        assert_eq!(ast, FormulaExpr::Number(1e10));
    }

    #[test]
    fn test_parse_string() {
        let ast = parse_formula("=\"Hello\"").unwrap();
        assert_eq!(ast, FormulaExpr::String("Hello".into()));

        let ast = parse_formula("=\"Hello \"\"World\"\"\"").unwrap();
        assert_eq!(ast, FormulaExpr::String("Hello \"World\"".into()));
    }

    #[test]
    fn test_parse_boolean() {
        let ast = parse_formula("=TRUE").unwrap();
        assert_eq!(ast, FormulaExpr::Boolean(true));

        let ast = parse_formula("=FALSE").unwrap();
        assert_eq!(ast, FormulaExpr::Boolean(false));
    }

    #[test]
    fn test_parse_arithmetic() {
        let ast = parse_formula("=1+2").unwrap();
        assert!(matches!(
            ast,
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Add,
                ..
            }
        ));

        let ast = parse_formula("=1+2*3").unwrap();
        // Should parse as 1+(2*3) due to precedence
        if let FormulaExpr::BinaryOp { op, left, right } = ast {
            assert_eq!(op, BinaryOperator::Add);
            assert_eq!(*left, FormulaExpr::Number(1.0));
            assert!(matches!(
                *right,
                FormulaExpr::BinaryOp {
                    op: BinaryOperator::Multiply,
                    ..
                }
            ));
        } else {
            panic!("Expected BinaryOp");
        }
    }

    #[test]
    fn test_parse_comparison() {
        let ast = parse_formula("=A1>5").unwrap();
        assert!(matches!(
            ast,
            FormulaExpr::BinaryOp {
                op: BinaryOperator::GreaterThan,
                ..
            }
        ));

        let ast = parse_formula("=A1<>B1").unwrap();
        assert!(matches!(
            ast,
            FormulaExpr::BinaryOp {
                op: BinaryOperator::NotEqual,
                ..
            }
        ));
    }

    #[test]
    fn test_parse_unary() {
        let ast = parse_formula("=-5").unwrap();
        assert!(matches!(
            ast,
            FormulaExpr::UnaryOp {
                op: UnaryOperator::Negate,
                ..
            }
        ));

        let ast = parse_formula("=50%").unwrap();
        assert!(matches!(
            ast,
            FormulaExpr::UnaryOp {
                op: UnaryOperator::Percent,
                ..
            }
        ));
    }

    #[test]
    fn test_parse_cell_reference() {
        let ast = parse_formula("=A1").unwrap();
        if let FormulaExpr::CellRef(cell_ref) = ast {
            assert_eq!(cell_ref.address.row, 0);
            assert_eq!(cell_ref.address.col, 0);
            assert!(cell_ref.sheet.is_none());
        } else {
            panic!("Expected CellRef");
        }

        let ast = parse_formula("=$B$2").unwrap();
        if let FormulaExpr::CellRef(cell_ref) = ast {
            assert_eq!(cell_ref.address.row, 1);
            assert_eq!(cell_ref.address.col, 1);
        } else {
            panic!("Expected CellRef");
        }
    }

    #[test]
    fn test_parse_range_reference() {
        let ast = parse_formula("=A1:B10").unwrap();
        if let FormulaExpr::RangeRef(range_ref) = ast {
            assert_eq!(range_ref.range.start.row, 0);
            assert_eq!(range_ref.range.start.col, 0);
            assert_eq!(range_ref.range.end.row, 9);
            assert_eq!(range_ref.range.end.col, 1);
        } else {
            panic!("Expected RangeRef");
        }
    }

    #[test]
    fn test_parse_function() {
        let ast = parse_formula("=SUM(1,2,3)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "SUM");
            assert_eq!(args.len(), 3);
        } else {
            panic!("Expected Function");
        }

        let ast = parse_formula("=SUM(A1:A10)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "SUM");
            assert_eq!(args.len(), 1);
            assert!(matches!(&args[0], FormulaExpr::RangeRef(_)));
        } else {
            panic!("Expected Function");
        }
    }

    #[test]
    fn test_parse_nested_function() {
        let ast = parse_formula("=IF(A1>0,SUM(B1:B10),0)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "IF");
            assert_eq!(args.len(), 3);
        } else {
            panic!("Expected Function");
        }
    }

    #[test]
    fn test_parse_parentheses() {
        let ast = parse_formula("=(1+2)*3").unwrap();
        if let FormulaExpr::BinaryOp { op, left, right } = ast {
            assert_eq!(op, BinaryOperator::Multiply);
            assert!(matches!(
                *left,
                FormulaExpr::BinaryOp {
                    op: BinaryOperator::Add,
                    ..
                }
            ));
            assert_eq!(*right, FormulaExpr::Number(3.0));
        } else {
            panic!("Expected BinaryOp");
        }
    }

    #[test]
    fn test_parse_array() {
        let ast = parse_formula("={1,2,3}").unwrap();
        if let FormulaExpr::Array(rows) = ast {
            assert_eq!(rows.len(), 1);
            assert_eq!(rows[0].len(), 3);
        } else {
            panic!("Expected Array");
        }

        let ast = parse_formula("={1,2;3,4}").unwrap();
        if let FormulaExpr::Array(rows) = ast {
            assert_eq!(rows.len(), 2);
            assert_eq!(rows[0].len(), 2);
            assert_eq!(rows[1].len(), 2);
        } else {
            panic!("Expected Array");
        }
    }

    #[test]
    fn test_parse_concatenation() {
        let ast = parse_formula("=\"Hello \"&\"World\"").unwrap();
        if let FormulaExpr::BinaryOp { op, .. } = ast {
            assert_eq!(op, BinaryOperator::Concat);
        } else {
            panic!("Expected BinaryOp");
        }
    }

    #[test]
    fn test_parse_error() {
        let ast = parse_formula("=#VALUE!").unwrap();
        assert_eq!(ast, FormulaExpr::Error(CellError::Value));

        let ast = parse_formula("=#DIV/0!").unwrap();
        assert_eq!(ast, FormulaExpr::Error(CellError::Div0));
    }

    #[test]
    fn test_complex_formula() {
        // A complex real-world formula
        let ast = parse_formula("=IF(AND(A1>0,B1<100),A1*B1/100,0)").unwrap();
        assert!(matches!(ast, FormulaExpr::Function { .. }));
    }
}
