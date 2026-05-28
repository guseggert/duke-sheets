//! Formula parser
//!
//! A recursive descent parser for Excel formulas with proper operator precedence.

use crate::ast::{
    BinaryOperator, CellReference, ExternalReference, FormulaExpr, RangeReference,
    StructuredRefSpecifier, StructuredReference, UnaryOperator,
};
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
    At,   // @ implicit intersection
    Hash, // # spill range
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

    /// Content between [ and ] (including nested brackets preserved)
    BracketExpr(String),

    /// Synthetic token: whitespace between two value-producing tokens
    /// (the implicit intersection operator). Distinct from skip_whitespace
    /// which silently swallows non-meaningful spaces.
    Space,

    // Unknown character (for better error reporting)
    Unknown(char),

    // End of input
    Eof,
}

/// Formula parser
struct FormulaParser<'a> {
    input: &'a str,
    pos: usize,
    current_token: Option<Token>,
    depth: usize,
}

impl<'a> FormulaParser<'a> {
    fn new(input: &'a str) -> Self {
        let mut parser = Self {
            input,
            pos: 0,
            current_token: None,
            depth: 0,
        };
        parser.advance_token();
        parser
    }

    fn advance_token(&mut self) {
        // The implicit intersection operator in Excel is whitespace
        // between two value-producing tokens. Detect it here so the
        // parser can consume Token::Space rather than silently
        // collapsing the gap.
        let prev_value_producing = matches!(
            self.current_token,
            Some(Token::CellRef(_))
                | Some(Token::Number(_))
                | Some(Token::String(_))
                | Some(Token::Boolean(_))
                | Some(Token::Error(_))
                | Some(Token::Identifier(_))
                | Some(Token::RightParen)
                | Some(Token::RightBrace)
                | Some(Token::BracketExpr(_))
                | Some(Token::Percent)
                | Some(Token::Hash)
        );
        let had_whitespace = self
            .peek_char()
            .is_some_and(|c| c.is_whitespace() && c != '\n');
        if prev_value_producing && had_whitespace {
            self.skip_whitespace();
            if let Some(next) = self.peek_char() {
                if matches!(
                    next,
                    'A'..='Z' | 'a'..='z' | '_' | '\'' | '$' | '(' | '0'..='9'
                ) {
                    self.current_token = Some(Token::Space);
                    return;
                }
            }
            // Whitespace was non-meaningful (trailing); fall through
            // to scan as usual. self.skip_whitespace() above already
            // moved pos past it.
            self.current_token = Some(self.scan_token());
            return;
        }
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
            '[' => {
                return self.scan_bracket_expr();
            }
            ']' => {
                // Stray ']' without matching '[' - treat as unknown
                self.advance();
                return Token::Unknown(']');
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
            || (c == '.' && self.peek_char_at(1).is_some_and(|c| c.is_ascii_digit()))
        {
            return self.scan_number();
        }

        // Implicit intersection operator
        if c == '@' {
            self.advance();
            return Token::At;
        }

        // Quoted sheet reference: 'Sheet Name'!
        if c == '\'' {
            return self.scan_quoted_sheet_ref();
        }

        // Identifier, cell reference, or boolean/error
        if c.is_ascii_alphabetic() || c == '_' || c == '$' || c == '#' {
            return self.scan_identifier_or_ref();
        }

        // Unknown character - emit Token::Unknown for better error messages
        let unknown = c;
        self.advance();
        Token::Unknown(unknown)
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

    fn scan_quoted_sheet_ref(&mut self) -> Token {
        self.advance(); // Skip opening apostrophe

        let mut name = String::new();
        loop {
            match self.peek_char() {
                Some('\'') => {
                    self.advance();
                    // Doubled apostrophe '' is an escaped literal apostrophe
                    if self.peek_char() == Some('\'') {
                        name.push('\'');
                        self.advance();
                    } else {
                        // End of quoted name - expect '!' next
                        break;
                    }
                }
                Some(c) => {
                    name.push(c);
                    self.advance();
                }
                None => {
                    // Unterminated quoted sheet name - return as unknown
                    return Token::Unknown('\'');
                }
            }
        }

        // After closing apostrophe, expect '!'
        if self.peek_char() == Some('!') {
            self.advance();
            Token::SheetRef(name)
        } else {
            // Quoted string without '!' - not a sheet ref
            // Return as unknown since we've consumed the content
            Token::Unknown('\'')
        }
    }

    fn scan_bracket_expr(&mut self) -> Token {
        self.advance(); // Skip opening '['
        let start = self.pos;
        let mut depth = 0;
        while let Some(c) = self.peek_char() {
            match c {
                '[' => {
                    depth += 1;
                    self.advance();
                }
                ']' => {
                    if depth == 0 {
                        let content = self.input[start..self.pos].to_string();
                        self.advance(); // Skip closing ']'
                        return Token::BracketExpr(content);
                    }
                    depth -= 1;
                    self.advance();
                }
                _ => self.advance(),
            }
        }
        // Unterminated bracket expression
        Token::Unknown('[')
    }

    fn scan_number(&mut self) -> Token {
        let start = self.pos;

        // Integer part
        while self.peek_char().is_some_and(|c| c.is_ascii_digit()) {
            self.advance();
        }

        // Decimal part
        if self.peek_char() == Some('.') {
            self.advance();
            while self.peek_char().is_some_and(|c| c.is_ascii_digit()) {
                self.advance();
            }
        }

        // Exponent part
        if self.peek_char().is_some_and(|c| c == 'e' || c == 'E') {
            self.advance();
            if self.peek_char().is_some_and(|c| c == '+' || c == '-') {
                self.advance();
            }
            while self.peek_char().is_some_and(|c| c.is_ascii_digit()) {
                self.advance();
            }
        }

        let num_str = &self.input[start..self.pos];
        let num: f64 = num_str.parse().unwrap_or(0.0);
        Token::Number(num)
    }

    fn scan_identifier_or_ref(&mut self) -> Token {
        // Check for error values first (#VALUE!, #REF!, etc.)
        // or standalone # (spill range operator)
        if self.peek_char() == Some('#') {
            // Peek ahead: if next char after # isn't alphanumeric, it's the
            // spill range operator (e.g., A1#)
            if !self
                .peek_char_at(1)
                .is_some_and(|c| c.is_ascii_alphabetic())
            {
                self.advance();
                return Token::Hash;
            }
            let start = self.pos;
            self.advance();
            while self
                .peek_char()
                .is_some_and(|c| c.is_ascii_alphanumeric() || c == '!' || c == '/' || c == '?')
            {
                self.advance();
            }
            let error_str = &self.input[start..self.pos];
            if let Some(err) = CellError::parse(error_str) {
                return Token::Error(err);
            }
            // If not a valid error, treat as identifier
            return Token::Identifier(error_str.to_string());
        }

        let start = self.pos;

        // Scan identifier/reference
        while self
            .peek_char()
            .is_some_and(|c| c.is_ascii_alphanumeric() || c == '_' || c == '$' || c == '.')
        {
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
        // AND if followed by '[' it's a table name for structured refs (e.g., Table1[Column])
        if Self::is_cell_reference(text)
            && self.peek_char() != Some('(')
            && self.peek_char() != Some('[')
        {
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
        while self.peek_char().is_some_and(|c| c.is_whitespace()) {
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
        // advance_token inspects self.current_token to detect the
        // implicit-intersection space between value-producing tokens.
        // Use clone() rather than take() so it can do that lookup.
        let token = self.current_token.clone().unwrap_or(Token::Eof);
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
        self.depth += 1;
        if self.depth > 256 {
            return Err(FormulaError::Parse("formula nesting too deep".into()));
        }
        let result = self.parse_comparison();
        self.depth -= 1;
        result
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

        // Prefix plus. Excel preserves it as a distinct PtgUplus token, so we
        // keep it in the AST rather than dropping it.
        if matches!(self.current_token(), Token::Plus) {
            self.consume();
            let operand = self.parse_unary()?;
            return Ok(FormulaExpr::UnaryOp {
                op: UnaryOperator::Plus,
                operand: Box::new(operand),
            });
        }

        // Prefix implicit intersection (@)
        if matches!(self.current_token(), Token::At) {
            self.consume();
            let operand = self.parse_unary()?;
            return Ok(FormulaExpr::UnaryOp {
                op: UnaryOperator::ImplicitIntersection,
                operand: Box::new(operand),
            });
        }

        // Parse primary (via intersection layer, which folds Space
        // tokens into BinaryOp(Intersect)), then check for postfix
        // operators (%, #).
        let mut expr = self.parse_intersect()?;

        loop {
            match self.current_token() {
                Token::Percent => {
                    self.consume();
                    expr = FormulaExpr::UnaryOp {
                        op: UnaryOperator::Percent,
                        operand: Box::new(expr),
                    };
                }
                Token::Hash => {
                    self.consume();
                    expr = FormulaExpr::UnaryOp {
                        op: UnaryOperator::SpillRange,
                        operand: Box::new(expr),
                    };
                }
                _ => break,
            }
        }

        Ok(expr)
    }

    /// Implicit intersection: a single space between two value-
    /// producing expressions, e.g. `A1:B3 B2:C3` (cells in both
    /// ranges). Excel precedence: range (:) > space > union (,).
    fn parse_intersect(&mut self) -> FormulaResult<FormulaExpr> {
        let mut left = self.parse_range()?;
        while matches!(self.current_token(), Token::Space) {
            self.consume();
            let right = self.parse_range()?;
            left = FormulaExpr::BinaryOp {
                op: BinaryOperator::Intersect,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
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
                // Resolve sheet: if right has no sheet, inherit from left
                // (e.g., Sheet1!A1:B10 means both A1 and B10 are on Sheet1)
                let left_sheet = &start_ref.sheet;
                let right_sheet = &end_ref.sheet;
                let sheet = match (left_sheet, right_sheet) {
                    (s, None) => s.clone(),
                    (None, s) => s.clone(),
                    (Some(a), Some(b)) if a == b => Some(a.clone()),
                    (Some(a), Some(b)) => {
                        return Err(FormulaError::Parse(format!(
                            "Range references must be on the same sheet: '{}' vs '{}'",
                            a, b
                        )));
                    }
                };

                let range = CellRange::new(start_ref.address, end_ref.address);
                return Ok(FormulaExpr::RangeRef(RangeReference { sheet, range }));
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
                let mut expr = self.parse_expression()?;
                // Bare parens accept comma-as-union per Excel: an
                // unparenthesised list of refs joined with commas
                // forms a Union expression. Function-call commas
                // are handled separately in parse_function_call,
                // which never reaches this branch.
                let mut is_union = false;
                while matches!(self.current_token(), Token::Comma) {
                    self.consume();
                    let right = self.parse_expression()?;
                    expr = FormulaExpr::BinaryOp {
                        op: BinaryOperator::Union,
                        left: Box::new(right),
                        right: Box::new(expr),
                    };
                    is_union = true;
                }
                self.expect(&Token::RightParen)?;
                // Excel preserves the parentheses as a postfix PtgParen token
                // (even when redundant or nested). The union-in-parens form
                // carries its own parenthesis handling in the writers'
                // PtgMemFunc path, so don't double-wrap it here.
                if is_union {
                    Ok(expr)
                } else {
                    Ok(FormulaExpr::UnaryOp {
                        op: UnaryOperator::Paren,
                        operand: Box::new(expr),
                    })
                }
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
                } else if let Token::BracketExpr(_) = self.current_token() {
                    // Structured table reference: Table1[Column1]
                    let Token::BracketExpr(content) = self.consume() else {
                        unreachable!()
                    };
                    self.parse_structured_ref_content(Some(name), &content)
                } else {
                    // Named range
                    Ok(FormulaExpr::NameRef(name))
                }
            }

            Token::BracketExpr(content) => {
                self.consume();
                // Could be:
                // 1. External workbook ref: [Book.xlsx]Sheet1!A1
                // 2. Unqualified structured ref: [Column1] or [#Headers]
                self.parse_bracket_expression(content)
            }

            Token::Unknown(c) => Err(FormulaError::Parse(format!(
                "Unexpected character: '{}'",
                c
            ))),

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
                        current_row = vec![self.parse_expression()?];
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

        // Parse arguments, supporting empty/omitted args (e.g., XLOOKUP(x,a,b,,1))
        if !matches!(self.current_token(), Token::RightParen) {
            // First argument: empty if immediately followed by comma
            if matches!(self.current_token(), Token::Comma) {
                args.push(FormulaExpr::Empty);
            } else {
                args.push(self.parse_expression()?);
            }

            while matches!(self.current_token(), Token::Comma) {
                self.consume();
                // Empty argument if next token is comma or rparen
                if matches!(self.current_token(), Token::Comma | Token::RightParen) {
                    args.push(FormulaExpr::Empty);
                } else {
                    args.push(self.parse_expression()?);
                }
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
        let address = CellAddress::parse(ref_str).map_err(|e| {
            FormulaError::Parse(format!("Invalid cell reference '{}': {}", ref_str, e))
        })?;

        Ok(FormulaExpr::CellRef(CellReference { sheet, address }))
    }

    /// Parse a bracket expression at expression start (no preceding table name).
    /// Could be an external workbook ref or an unqualified structured ref.
    /// `content` is the text that was between [ and ].
    fn parse_bracket_expression(&mut self, content: String) -> FormulaResult<FormulaExpr> {
        // Check if what follows is a sheet reference or cell reference
        // → external workbook reference: [Book.xlsx]Sheet1!A1
        match self.current_token().clone() {
            Token::SheetRef(sheet) => {
                self.consume();
                match self.current_token().clone() {
                    Token::CellRef(ref_str) => {
                        self.consume();
                        let address = CellAddress::parse(&ref_str).map_err(|e| {
                            FormulaError::Parse(format!(
                                "Invalid cell reference '{}': {}",
                                ref_str, e
                            ))
                        })?;
                        Ok(FormulaExpr::ExternalRef(ExternalReference {
                            book: content,
                            sheet: Some(sheet),
                            address,
                        }))
                    }
                    _ => Err(FormulaError::Parse(
                        "Expected cell reference after external sheet name".into(),
                    )),
                }
            }
            Token::CellRef(ref_str) => {
                self.consume();
                let clean_ref = ref_str.replace('$', "");
                let address = CellAddress::parse(&clean_ref).map_err(|e| {
                    FormulaError::Parse(format!("Invalid cell reference '{}': {}", ref_str, e))
                })?;
                Ok(FormulaExpr::ExternalRef(ExternalReference {
                    book: content,
                    sheet: None,
                    address,
                }))
            }
            Token::Identifier(name) => {
                self.consume();
                // Named range in external workbook
                Ok(FormulaExpr::NameRef(format!("[{}]{}", content, name)))
            }
            _ => {
                // Not followed by a reference - unqualified structured ref
                self.parse_structured_ref_content(None, &content)
            }
        }
    }

    /// Parse the content of a structured reference bracket.
    /// `table` is the optional table name. `content` is the text between [ and ].
    fn parse_structured_ref_content(
        &self,
        table: Option<String>,
        content: &str,
    ) -> FormulaResult<FormulaExpr> {
        let content = content.trim();

        // Check for nested brackets: [[#Headers],[Column1]]
        if content.starts_with('[') && content.ends_with(']') {
            // Complex structured ref with multiple bracket groups
            let inner = &content[1..content.len() - 1];
            let mut specifiers = Vec::new();
            let mut column = None;

            // Split on "],[" to separate bracket groups
            for part in Self::split_bracket_groups(inner) {
                let part = part.trim();
                if part.starts_with('#') {
                    specifiers.push(Self::parse_specifier_keyword(part)?);
                } else if let Some(stripped) = part.strip_prefix('@') {
                    specifiers.push(StructuredRefSpecifier::ThisRow);
                    let col_name = stripped.trim();
                    if !col_name.is_empty() {
                        column = Some(col_name.to_string());
                    }
                } else {
                    column = Some(part.to_string());
                }
            }

            Ok(FormulaExpr::StructuredRef(StructuredReference {
                table,
                column,
                specifiers,
            }))
        } else if content.starts_with('#') {
            // Simple specifier: [#All], [#Headers], etc.
            let spec = Self::parse_specifier_keyword(content)?;
            Ok(FormulaExpr::StructuredRef(StructuredReference {
                table,
                column: None,
                specifiers: vec![spec],
            }))
        } else if let Some(stripped) = content.strip_prefix('@') {
            // This-row shorthand: [@Column1]
            let col_name = stripped.trim();
            Ok(FormulaExpr::StructuredRef(StructuredReference {
                table,
                column: if col_name.is_empty() {
                    None
                } else {
                    Some(col_name.to_string())
                },
                specifiers: vec![StructuredRefSpecifier::ThisRow],
            }))
        } else {
            // Plain column name: [Column1]
            Ok(FormulaExpr::StructuredRef(StructuredReference {
                table,
                column: Some(content.to_string()),
                specifiers: vec![],
            }))
        }
    }

    /// Split bracket group content on "],["  boundaries.
    /// E.g., "#Headers],[Column1" → ["#Headers", "Column1"]
    fn split_bracket_groups(inner: &str) -> Vec<&str> {
        let mut parts = Vec::new();
        let mut start = 0;
        let bytes = inner.as_bytes();
        let mut i = 0;
        while i < bytes.len() {
            if bytes[i] == b']'
                && i + 2 < bytes.len()
                && bytes[i + 1] == b','
                && bytes[i + 2] == b'['
            {
                parts.push(&inner[start..i]);
                i += 3; // skip ],[
                start = i;
            } else {
                i += 1;
            }
        }
        parts.push(&inner[start..]);
        parts
    }

    fn parse_specifier_keyword(s: &str) -> FormulaResult<StructuredRefSpecifier> {
        match s.to_lowercase().as_str() {
            "#all" => Ok(StructuredRefSpecifier::All),
            "#data" => Ok(StructuredRefSpecifier::Data),
            "#headers" => Ok(StructuredRefSpecifier::Headers),
            "#totals" => Ok(StructuredRefSpecifier::Totals),
            "#this row" => Ok(StructuredRefSpecifier::ThisRow),
            _ => Err(FormulaError::Parse(format!(
                "Unknown structured reference specifier: '{}'",
                s
            ))),
        }
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
        // Parentheses are preserved as a UnaryOp(Paren) wrapper so they
        // round-trip to Excel's PtgParen token byte-for-byte.
        let ast = parse_formula("=(1+2)*3").unwrap();
        if let FormulaExpr::BinaryOp { op, left, right } = ast {
            assert_eq!(op, BinaryOperator::Multiply);
            let FormulaExpr::UnaryOp {
                op: UnaryOperator::Paren,
                operand,
            } = *left
            else {
                panic!("Expected left to be UnaryOp(Paren), got {left:?}");
            };
            assert!(matches!(
                *operand,
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

    #[test]
    fn test_parse_quoted_sheet_ref() {
        // Basic quoted sheet name with space
        let ast = parse_formula("='Sheet 1'!A1").unwrap();
        if let FormulaExpr::CellRef(cell_ref) = ast {
            assert_eq!(cell_ref.sheet, Some("Sheet 1".to_string()));
            assert_eq!(cell_ref.address.row, 0);
            assert_eq!(cell_ref.address.col, 0);
        } else {
            panic!("Expected CellRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_quoted_sheet_ref_hyphen() {
        // Sheet name with hyphen
        let ast = parse_formula("='Data-2025'!B2").unwrap();
        if let FormulaExpr::CellRef(cell_ref) = ast {
            assert_eq!(cell_ref.sheet, Some("Data-2025".to_string()));
            assert_eq!(cell_ref.address.row, 1);
            assert_eq!(cell_ref.address.col, 1);
        } else {
            panic!("Expected CellRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_quoted_sheet_ref_escaped_apostrophe() {
        // Doubled apostrophe escape: 'It''s A Sheet'!A1
        let ast = parse_formula("='It''s A Sheet'!A1").unwrap();
        if let FormulaExpr::CellRef(cell_ref) = ast {
            assert_eq!(cell_ref.sheet, Some("It's A Sheet".to_string()));
        } else {
            panic!("Expected CellRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_quoted_sheet_ref_range() {
        // Quoted sheet ref with range
        let ast = parse_formula("='Sheet 1'!B2:C10").unwrap();
        if let FormulaExpr::RangeRef(range_ref) = ast {
            assert_eq!(range_ref.sheet, Some("Sheet 1".to_string()));
            assert_eq!(range_ref.range.start.row, 1);
            assert_eq!(range_ref.range.start.col, 1);
            assert_eq!(range_ref.range.end.row, 9);
            assert_eq!(range_ref.range.end.col, 2);
        } else {
            panic!("Expected RangeRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_quoted_sheet_ref_in_function() {
        // Quoted sheet ref inside a function call
        let ast = parse_formula("=SUM('Sheet 1'!A1:A10)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "SUM");
            assert_eq!(args.len(), 1);
            if let FormulaExpr::RangeRef(range_ref) = &args[0] {
                assert_eq!(range_ref.sheet, Some("Sheet 1".to_string()));
            } else {
                panic!("Expected RangeRef arg, got {:?}", args[0]);
            }
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_multiple_quoted_sheet_refs() {
        // Two quoted sheet refs combined with operator
        let ast = parse_formula("='Sheet1'!A1+'Sheet2'!B1").unwrap();
        if let FormulaExpr::BinaryOp {
            op: BinaryOperator::Add,
            left,
            right,
        } = ast
        {
            if let FormulaExpr::CellRef(r) = left.as_ref() {
                assert_eq!(r.sheet, Some("Sheet1".to_string()));
            } else {
                panic!("Expected CellRef left");
            }
            if let FormulaExpr::CellRef(r) = right.as_ref() {
                assert_eq!(r.sheet, Some("Sheet2".to_string()));
            } else {
                panic!("Expected CellRef right");
            }
        } else {
            panic!("Expected BinaryOp Add");
        }
    }

    #[test]
    fn test_unknown_char_error_message() {
        // Unknown character should produce a clear error, not "Unexpected token: Eof"
        let err = parse_formula("=~A1").unwrap_err();
        let msg = err.to_string();
        assert!(
            msg.contains("Unexpected character: '~'"),
            "Expected descriptive error, got: {}",
            msg
        );
        // Verify it does NOT say "Eof"
        assert!(!msg.contains("Eof"), "Should not mention Eof: {}", msg);
    }

    #[test]
    fn test_unterminated_quoted_sheet_ref() {
        // Unterminated quoted sheet name should produce error
        let err = parse_formula("='Sheet 1!A1").unwrap_err();
        assert!(err.to_string().contains("Unexpected character"));
    }

    #[test]
    fn test_parse_implicit_intersection() {
        // @ as prefix operator
        let ast = parse_formula("=@A1").unwrap();
        if let FormulaExpr::UnaryOp { op, operand } = ast {
            assert_eq!(op, UnaryOperator::ImplicitIntersection);
            assert!(matches!(*operand, FormulaExpr::CellRef(_)));
        } else {
            panic!("Expected UnaryOp, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_implicit_intersection_range() {
        // @ with a range reference
        let ast = parse_formula("=@A1:A10").unwrap();
        if let FormulaExpr::UnaryOp { op, operand } = ast {
            assert_eq!(op, UnaryOperator::ImplicitIntersection);
            assert!(matches!(*operand, FormulaExpr::RangeRef(_)));
        } else {
            panic!("Expected UnaryOp, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_implicit_intersection_in_function() {
        // @ inside a function call
        let ast = parse_formula("=SUM(@A1:A10)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "SUM");
            assert_eq!(args.len(), 1);
            assert!(matches!(
                &args[0],
                FormulaExpr::UnaryOp {
                    op: UnaryOperator::ImplicitIntersection,
                    ..
                }
            ));
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_spill_range() {
        // # as postfix operator on a cell reference
        let ast = parse_formula("=A1#").unwrap();
        if let FormulaExpr::UnaryOp { op, operand } = ast {
            assert_eq!(op, UnaryOperator::SpillRange);
            assert!(matches!(*operand, FormulaExpr::CellRef(_)));
        } else {
            panic!("Expected UnaryOp, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_spill_range_in_function() {
        // # inside SUM
        let ast = parse_formula("=SUM(A1#)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "SUM");
            assert_eq!(args.len(), 1);
            assert!(matches!(
                &args[0],
                FormulaExpr::UnaryOp {
                    op: UnaryOperator::SpillRange,
                    ..
                }
            ));
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_error_values_still_work() {
        // Make sure error values like #VALUE! still parse correctly
        // (they also start with # so we need to verify no regression)
        let ast = parse_formula("=#VALUE!").unwrap();
        assert_eq!(ast, FormulaExpr::Error(CellError::Value));

        let ast = parse_formula("=#N/A").unwrap();
        assert_eq!(ast, FormulaExpr::Error(CellError::Na));

        let ast = parse_formula("=#REF!").unwrap();
        assert_eq!(ast, FormulaExpr::Error(CellError::Ref));
    }

    #[test]
    fn test_parse_structured_ref_simple() {
        // Table1[Column1]
        let ast = parse_formula("=Table1[Column1]").unwrap();
        if let FormulaExpr::StructuredRef(sr) = ast {
            assert_eq!(sr.table, Some("Table1".to_string()));
            assert_eq!(sr.column, Some("Column1".to_string()));
            assert!(sr.specifiers.is_empty());
        } else {
            panic!("Expected StructuredRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_structured_ref_this_row() {
        // Table1[@Column1] - shorthand for this-row
        let ast = parse_formula("=Table1[@Column1]").unwrap();
        if let FormulaExpr::StructuredRef(sr) = ast {
            assert_eq!(sr.table, Some("Table1".to_string()));
            assert_eq!(sr.column, Some("Column1".to_string()));
            assert_eq!(sr.specifiers, vec![StructuredRefSpecifier::ThisRow]);
        } else {
            panic!("Expected StructuredRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_structured_ref_specifier() {
        // Table1[#All]
        let ast = parse_formula("=Table1[#All]").unwrap();
        if let FormulaExpr::StructuredRef(sr) = ast {
            assert_eq!(sr.table, Some("Table1".to_string()));
            assert!(sr.column.is_none());
            assert_eq!(sr.specifiers, vec![StructuredRefSpecifier::All]);
        } else {
            panic!("Expected StructuredRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_structured_ref_complex() {
        // Table1[[#Headers],[Column1]]
        let ast = parse_formula("=Table1[[#Headers],[Column1]]").unwrap();
        if let FormulaExpr::StructuredRef(sr) = ast {
            assert_eq!(sr.table, Some("Table1".to_string()));
            assert_eq!(sr.column, Some("Column1".to_string()));
            assert_eq!(sr.specifiers, vec![StructuredRefSpecifier::Headers]);
        } else {
            panic!("Expected StructuredRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_structured_ref_in_function() {
        // SUM(Table1[Sales])
        let ast = parse_formula("=SUM(Table1[Sales])").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "SUM");
            assert_eq!(args.len(), 1);
            assert!(matches!(&args[0], FormulaExpr::StructuredRef(_)));
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_unqualified_structured_ref() {
        // [Column1] - no table name
        let ast = parse_formula("=[Column1]").unwrap();
        if let FormulaExpr::StructuredRef(sr) = ast {
            assert!(sr.table.is_none());
            assert_eq!(sr.column, Some("Column1".to_string()));
            assert!(sr.specifiers.is_empty());
        } else {
            panic!("Expected StructuredRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_external_ref() {
        // [Book.xlsx]Sheet1!A1
        let ast = parse_formula("=[Book.xlsx]Sheet1!A1").unwrap();
        if let FormulaExpr::ExternalRef(ext) = ast {
            assert_eq!(ext.book, "Book.xlsx");
            assert_eq!(ext.sheet, Some("Sheet1".to_string()));
            assert_eq!(ext.address.row, 0);
            assert_eq!(ext.address.col, 0);
        } else {
            panic!("Expected ExternalRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_external_ref_quoted_sheet() {
        // [Data.xlsx]'Sheet 1'!B5
        let ast = parse_formula("=[Data.xlsx]'Sheet 1'!B5").unwrap();
        if let FormulaExpr::ExternalRef(ext) = ast {
            assert_eq!(ext.book, "Data.xlsx");
            assert_eq!(ext.sheet, Some("Sheet 1".to_string()));
            assert_eq!(ext.address.row, 4);
            assert_eq!(ext.address.col, 1);
        } else {
            panic!("Expected ExternalRef, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_empty_args_middle() {
        // XLOOKUP(x,a,b,,1) - 4th arg is empty
        let ast = parse_formula("=FUNC(1,2,,4)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "FUNC");
            assert_eq!(args.len(), 4);
            assert_eq!(args[0], FormulaExpr::Number(1.0));
            assert_eq!(args[1], FormulaExpr::Number(2.0));
            assert_eq!(args[2], FormulaExpr::Empty);
            assert_eq!(args[3], FormulaExpr::Number(4.0));
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_empty_args_leading() {
        // FUNC(,1) - 1st arg is empty
        let ast = parse_formula("=FUNC(,1)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "FUNC");
            assert_eq!(args.len(), 2);
            assert_eq!(args[0], FormulaExpr::Empty);
            assert_eq!(args[1], FormulaExpr::Number(1.0));
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_empty_args_trailing() {
        // FUNC(1,) - 2nd arg is empty
        let ast = parse_formula("=FUNC(1,)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "FUNC");
            assert_eq!(args.len(), 2);
            assert_eq!(args[0], FormulaExpr::Number(1.0));
            assert_eq!(args[1], FormulaExpr::Empty);
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_empty_args_multiple_consecutive() {
        // FUNC(1,,,4) - 2nd and 3rd args are empty
        let ast = parse_formula("=FUNC(1,,,4)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "FUNC");
            assert_eq!(args.len(), 4);
            assert_eq!(args[0], FormulaExpr::Number(1.0));
            assert_eq!(args[1], FormulaExpr::Empty);
            assert_eq!(args[2], FormulaExpr::Empty);
            assert_eq!(args[3], FormulaExpr::Number(4.0));
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_empty_args_all_empty() {
        // FUNC(,,) - all 3 args empty
        let ast = parse_formula("=FUNC(,,)").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "FUNC");
            assert_eq!(args.len(), 3);
            assert_eq!(args[0], FormulaExpr::Empty);
            assert_eq!(args[1], FormulaExpr::Empty);
            assert_eq!(args[2], FormulaExpr::Empty);
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_empty_args_no_args_still_works() {
        // FUNC() - zero args, no empties
        let ast = parse_formula("=FUNC()").unwrap();
        if let FormulaExpr::Function { name, args } = ast {
            assert_eq!(name, "FUNC");
            assert_eq!(args.len(), 0);
        } else {
            panic!("Expected Function, got {:?}", ast);
        }
    }

    #[test]
    fn test_parse_intersection() {
        let ast = parse_formula("=A1:B3 B2:C3").unwrap();
        match ast {
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Intersect,
                ..
            } => {}
            other => panic!("Expected Intersect, got {other:?}"),
        }
    }

    #[test]
    fn test_parse_intersection_in_function() {
        let ast = parse_formula("=SUM(A1:B3 B2:C3)").unwrap();
        let FormulaExpr::Function { name, args } = ast else {
            panic!("Expected SUM");
        };
        assert_eq!(name, "SUM");
        assert_eq!(args.len(), 1);
        assert!(matches!(
            &args[0],
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Intersect,
                ..
            }
        ));
    }

    #[test]
    fn test_parse_union_in_parens() {
        let ast = parse_formula("=(A1:A2,C2:C3)").unwrap();
        match ast {
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Union,
                ..
            } => {}
            other => panic!("Expected Union, got {other:?}"),
        }
    }

    #[test]
    fn test_parse_function_call_still_uses_comma_as_arg_separator() {
        let ast = parse_formula("=SUM(A1,B1)").unwrap();
        let FormulaExpr::Function { args, .. } = ast else {
            panic!("Expected Function");
        };
        assert_eq!(args.len(), 2, "comma should be arg separator inside SUM");
    }
}
