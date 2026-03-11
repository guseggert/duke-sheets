# ssfmt Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a Rust crate for Excel-compatible ECMA-376 number format codes.

**Architecture:** Parse format codes into an AST, then walk the AST to format values. Compile-once, format-many pattern for performance. ICU4X for locale support behind feature flag.

**Tech Stack:** Rust, chrono (optional), icu (optional), lru for caching

---

## Phase 1: Project Foundation

### Task 1.1: Set Up Cargo.toml

**Files:**
- Modify: `Cargo.toml`

**Step 1: Update Cargo.toml with dependencies and features**

```toml
[package]
name = "ssfmt"
version = "0.1.0"
edition = "2021"
description = "Excel-compatible ECMA-376 number format codes"
license = "MIT OR Apache-2.0"
repository = "https://github.com/yourusername/ssfmt"
keywords = ["excel", "spreadsheet", "formatting", "ecma-376"]
categories = ["parsing", "text-processing"]

[dependencies]
chrono = { version = "0.4", optional = true, default-features = false, features = ["std"] }
lru = "0.12"
thiserror = "1.0"

[dev-dependencies]
serde = { version = "1.0", features = ["derive"] }
serde_json = "1.0"

[features]
default = ["chrono"]
chrono = ["dep:chrono"]
```

**Step 2: Verify it compiles**

Run: `cargo check`
Expected: Success

**Step 3: Commit**

```bash
git add Cargo.toml
git commit -m "feat: configure Cargo.toml with dependencies and features"
```

---

### Task 1.2: Create Module Structure

**Files:**
- Create: `src/error.rs`
- Create: `src/ast.rs`
- Create: `src/value.rs`
- Create: `src/options.rs`
- Create: `src/date_serial.rs`
- Create: `src/parser/mod.rs`
- Create: `src/parser/tokens.rs`
- Create: `src/parser/lexer.rs`
- Create: `src/formatter/mod.rs`
- Create: `src/formatter/number.rs`
- Create: `src/formatter/date.rs`
- Create: `src/formatter/text.rs`
- Create: `src/formatter/fraction.rs`
- Create: `src/locale/mod.rs`
- Create: `src/locale/builtin.rs`
- Create: `src/cache.rs`
- Modify: `src/lib.rs`

**Step 1: Create empty module files**

Create each file with a placeholder comment:

```rust
//! Module description here
```

**Step 2: Set up lib.rs with module declarations**

```rust
//! ssfmt - Excel-compatible ECMA-376 number format codes
//!
//! This crate provides parsing and formatting of spreadsheet number format codes,
//! matching Excel's actual behavior including undocumented quirks.

pub mod ast;
pub mod error;
pub mod options;
pub mod value;

mod cache;
mod date_serial;
mod formatter;
mod locale;
mod parser;

// Re-exports
pub use ast::NumberFormat;
pub use error::{FormatError, ParseError};
pub use options::{DateSystem, FormatOptions};
pub use value::Value;
```

**Step 3: Verify it compiles**

Run: `cargo check`
Expected: Success (with warnings about unused modules)

**Step 4: Commit**

```bash
git add src/
git commit -m "feat: create module structure skeleton"
```

---

## Phase 2: Error Types

### Task 2.1: Implement ParseError

**Files:**
- Modify: `src/error.rs`
- Create: `tests/error_tests.rs`

**Step 1: Write the failing test**

Create `tests/error_tests.rs`:

```rust
use ssfmt::ParseError;

#[test]
fn test_parse_error_display() {
    let err = ParseError::UnexpectedToken { position: 5, found: 'x' };
    let msg = format!("{}", err);
    assert!(msg.contains("position 5"));
    assert!(msg.contains("'x'"));
}

#[test]
fn test_parse_error_too_many_sections() {
    let err = ParseError::TooManySections;
    let msg = format!("{}", err);
    assert!(msg.contains("4"));
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test error_tests`
Expected: FAIL - ParseError not found

**Step 3: Implement ParseError**

```rust
//! Error types for parsing and formatting.

use thiserror::Error;

/// Errors that can occur when parsing a format code.
#[derive(Debug, Clone, PartialEq, Error)]
pub enum ParseError {
    #[error("unexpected token at position {position}: found '{found}'")]
    UnexpectedToken { position: usize, found: char },

    #[error("unterminated bracket at position {position}")]
    UnterminatedBracket { position: usize },

    #[error("invalid condition at position {position}: {reason}")]
    InvalidCondition { position: usize, reason: String },

    #[error("invalid locale code at position {position}")]
    InvalidLocaleCode { position: usize },

    #[error("too many sections (maximum 4 allowed)")]
    TooManySections,

    #[error("empty format code")]
    EmptyFormat,
}

/// Errors that can occur when formatting a value.
#[derive(Debug, Clone, PartialEq, Error)]
pub enum FormatError {
    #[error("type mismatch: expected {expected}, got {got}")]
    TypeMismatch {
        expected: &'static str,
        got: &'static str,
    },

    #[error("date out of range: serial number {serial}")]
    DateOutOfRange { serial: f64 },

    #[error("invalid serial number: {value}")]
    InvalidSerialNumber { value: f64 },
}
```

**Step 4: Run test to verify it passes**

Run: `cargo test --test error_tests`
Expected: PASS

**Step 5: Commit**

```bash
git add src/error.rs tests/error_tests.rs
git commit -m "feat: implement ParseError and FormatError types"
```

---

## Phase 3: AST Types

### Task 3.1: Implement Color and Condition Enums

**Files:**
- Modify: `src/ast.rs`
- Create: `tests/ast_tests.rs`

**Step 1: Write the failing test**

Create `tests/ast_tests.rs`:

```rust
use ssfmt::ast::{Color, Condition, NamedColor};

#[test]
fn test_named_color_from_str() {
    assert_eq!("Red".parse::<NamedColor>().unwrap(), NamedColor::Red);
    assert_eq!("blue".parse::<NamedColor>().unwrap(), NamedColor::Blue);
    assert!("invalid".parse::<NamedColor>().is_err());
}

#[test]
fn test_condition_evaluate() {
    let cond = Condition::GreaterThan(100.0);
    assert!(cond.evaluate(150.0));
    assert!(!cond.evaluate(50.0));
    assert!(!cond.evaluate(100.0));
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test ast_tests`
Expected: FAIL - types not found

**Step 3: Implement Color and Condition**

```rust
//! AST types for parsed format codes.

use std::str::FromStr;

/// Named colors supported in format codes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NamedColor {
    Black,
    Blue,
    Cyan,
    Green,
    Magenta,
    Red,
    White,
    Yellow,
}

impl FromStr for NamedColor {
    type Err = ();

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s.to_lowercase().as_str() {
            "black" => Ok(NamedColor::Black),
            "blue" => Ok(NamedColor::Blue),
            "cyan" => Ok(NamedColor::Cyan),
            "green" => Ok(NamedColor::Green),
            "magenta" => Ok(NamedColor::Magenta),
            "red" => Ok(NamedColor::Red),
            "white" => Ok(NamedColor::White),
            "yellow" => Ok(NamedColor::Yellow),
            _ => Err(()),
        }
    }
}

/// Color specification in a format section.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Color {
    Named(NamedColor),
    Indexed(u8),
}

/// Conditional expression for section selection.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Condition {
    GreaterThan(f64),
    LessThan(f64),
    Equal(f64),
    GreaterOrEqual(f64),
    LessOrEqual(f64),
    NotEqual(f64),
}

impl Condition {
    /// Evaluate this condition against a value.
    pub fn evaluate(&self, value: f64) -> bool {
        match self {
            Condition::GreaterThan(n) => value > *n,
            Condition::LessThan(n) => value < *n,
            Condition::Equal(n) => (value - n).abs() < f64::EPSILON,
            Condition::GreaterOrEqual(n) => value >= *n,
            Condition::LessOrEqual(n) => value <= *n,
            Condition::NotEqual(n) => (value - n).abs() >= f64::EPSILON,
        }
    }
}
```

**Step 4: Update lib.rs to expose ast module contents**

In `src/lib.rs`, the `pub mod ast;` already exposes it.

**Step 5: Run test to verify it passes**

Run: `cargo test --test ast_tests`
Expected: PASS

**Step 6: Commit**

```bash
git add src/ast.rs tests/ast_tests.rs
git commit -m "feat: implement Color and Condition AST types"
```

---

### Task 3.2: Implement FormatPart Enum

**Files:**
- Modify: `src/ast.rs`
- Modify: `tests/ast_tests.rs`

**Step 1: Write the failing test**

Add to `tests/ast_tests.rs`:

```rust
use ssfmt::ast::{DatePart, DigitPlaceholder, FormatPart};

#[test]
fn test_digit_placeholder_properties() {
    assert!(DigitPlaceholder::Zero.is_required());
    assert!(!DigitPlaceholder::Hash.is_required());
    assert!(!DigitPlaceholder::Question.is_required());
}

#[test]
fn test_format_part_is_date_part() {
    let year = FormatPart::DatePart(DatePart::Year4);
    let digit = FormatPart::Digit(DigitPlaceholder::Zero);

    assert!(year.is_date_part());
    assert!(!digit.is_date_part());
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test ast_tests`
Expected: FAIL - types not found

**Step 3: Implement FormatPart and related types**

Add to `src/ast.rs`:

```rust
/// Digit placeholder type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DigitPlaceholder {
    /// `0` - Display digit or zero
    Zero,
    /// `#` - Display digit or nothing
    Hash,
    /// `?` - Display digit or space
    Question,
}

impl DigitPlaceholder {
    /// Returns true if this placeholder requires a digit (shows 0 for missing).
    pub fn is_required(&self) -> bool {
        matches!(self, DigitPlaceholder::Zero)
    }

    /// Returns the character to display when no digit is present.
    pub fn empty_char(&self) -> Option<char> {
        match self {
            DigitPlaceholder::Zero => Some('0'),
            DigitPlaceholder::Hash => None,
            DigitPlaceholder::Question => Some(' '),
        }
    }
}

/// Date/time format parts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DatePart {
    /// `yy` - Two-digit year
    Year2,
    /// `yyyy` - Four-digit year
    Year4,
    /// `m` - Month (1-12)
    Month,
    /// `mm` - Month with leading zero (01-12)
    Month2,
    /// `mmm` - Abbreviated month name
    MonthAbbr,
    /// `mmmm` - Full month name
    MonthFull,
    /// `mmmmm` - First letter of month
    MonthLetter,
    /// `d` - Day (1-31)
    Day,
    /// `dd` - Day with leading zero (01-31)
    Day2,
    /// `ddd` - Abbreviated day name
    DayAbbr,
    /// `dddd` - Full day name
    DayFull,
    /// `h` - Hour (0-23 or 1-12)
    Hour,
    /// `hh` - Hour with leading zero
    Hour2,
    /// `m` (after hour) - Minute
    Minute,
    /// `mm` (after hour) - Minute with leading zero
    Minute2,
    /// `s` - Second
    Second,
    /// `ss` - Second with leading zero
    Second2,
    /// `.0`, `.00`, etc. - Fractional seconds
    SubSecond(u8),
}

/// AM/PM format style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AmPmStyle {
    /// `AM/PM`
    Upper,
    /// `am/pm`
    Lower,
    /// `A/P`
    ShortUpper,
    /// `a/p`
    ShortLower,
}

/// Elapsed time format part (for durations).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ElapsedPart {
    /// `[h]` - Total hours
    Hours,
    /// `[m]` - Total minutes
    Minutes,
    /// `[s]` - Total seconds
    Seconds,
}

/// Fraction denominator specification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FractionDenom {
    /// `?/?` - Up to N digits
    UpToDigits(u8),
    /// `?/8` - Fixed denominator
    Fixed(u32),
}

/// Locale code from format string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LocaleCode {
    /// Currency symbol override (e.g., "€")
    pub currency: Option<String>,
    /// Locale identifier (e.g., 0x409 for en-US)
    pub lcid: Option<u32>,
}

/// A single part of a format section.
#[derive(Debug, Clone, PartialEq)]
pub enum FormatPart {
    /// Literal text
    Literal(String),
    /// Digit placeholder (0, #, ?)
    Digit(DigitPlaceholder),
    /// Decimal point
    DecimalPoint,
    /// Thousands separator (also scaling indicator)
    ThousandsSeparator,
    /// Percent sign (multiplies by 100)
    Percent,
    /// Scientific notation (E+, E-, e+, e-)
    Scientific { upper: bool, show_plus: bool },
    /// Fraction format
    Fraction {
        integer_digits: Vec<DigitPlaceholder>,
        numerator_digits: Vec<DigitPlaceholder>,
        denominator: FractionDenom,
    },
    /// Date/time part
    DatePart(DatePart),
    /// AM/PM indicator
    AmPm(AmPmStyle),
    /// Elapsed time part
    Elapsed(ElapsedPart),
    /// Text placeholder (@)
    TextPlaceholder,
    /// Fill character (*x)
    Fill(char),
    /// Skip width (_x)
    Skip(char),
    /// Locale code ([$...])
    Locale(LocaleCode),
}

impl FormatPart {
    /// Returns true if this is a date/time part.
    pub fn is_date_part(&self) -> bool {
        matches!(
            self,
            FormatPart::DatePart(_) | FormatPart::AmPm(_) | FormatPart::Elapsed(_)
        )
    }

    /// Returns true if this is a numeric part.
    pub fn is_numeric_part(&self) -> bool {
        matches!(
            self,
            FormatPart::Digit(_)
                | FormatPart::DecimalPoint
                | FormatPart::ThousandsSeparator
                | FormatPart::Percent
                | FormatPart::Scientific { .. }
                | FormatPart::Fraction { .. }
        )
    }
}
```

**Step 4: Run test to verify it passes**

Run: `cargo test --test ast_tests`
Expected: PASS

**Step 5: Commit**

```bash
git add src/ast.rs tests/ast_tests.rs
git commit -m "feat: implement FormatPart and related AST types"
```

---

### Task 3.3: Implement Section and NumberFormat

**Files:**
- Modify: `src/ast.rs`
- Modify: `tests/ast_tests.rs`

**Step 1: Write the failing test**

Add to `tests/ast_tests.rs`:

```rust
use ssfmt::NumberFormat;
use ssfmt::ast::Section;

#[test]
fn test_number_format_is_date_format() {
    // A format with date parts should be detected as date format
    let section = Section {
        condition: None,
        color: None,
        parts: vec![
            FormatPart::DatePart(DatePart::Year4),
            FormatPart::Literal("-".into()),
            FormatPart::DatePart(DatePart::Month2),
        ],
    };
    let format = NumberFormat::from_sections(vec![section]);
    assert!(format.is_date_format());
}

#[test]
fn test_number_format_sections_limit() {
    let sections: Vec<Section> = (0..5)
        .map(|_| Section {
            condition: None,
            color: None,
            parts: vec![],
        })
        .collect();
    // Should only keep first 4 sections
    let format = NumberFormat::from_sections(sections);
    assert_eq!(format.sections().len(), 4);
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test ast_tests`
Expected: FAIL - Section and NumberFormat not found

**Step 3: Implement Section and NumberFormat**

Add to `src/ast.rs`:

```rust
/// A single section of a format code.
///
/// Format codes can have up to 4 sections:
/// 1. Positive numbers (or all numbers if only one section)
/// 2. Negative numbers
/// 3. Zero
/// 4. Text
#[derive(Debug, Clone, PartialEq)]
pub struct Section {
    /// Optional condition for this section (e.g., [>100])
    pub condition: Option<Condition>,
    /// Optional color for this section (e.g., [Red])
    pub color: Option<Color>,
    /// The format parts that make up this section
    pub parts: Vec<FormatPart>,
}

impl Section {
    /// Returns true if this section contains any date/time parts.
    pub fn has_date_parts(&self) -> bool {
        self.parts.iter().any(|p| p.is_date_part())
    }

    /// Returns true if this section contains a text placeholder.
    pub fn has_text_placeholder(&self) -> bool {
        self.parts.iter().any(|p| matches!(p, FormatPart::TextPlaceholder))
    }

    /// Returns true if this section contains a percent sign.
    pub fn has_percent(&self) -> bool {
        self.parts.iter().any(|p| matches!(p, FormatPart::Percent))
    }
}

/// A parsed number format code.
///
/// This is the main type returned by parsing. It can be reused to format
/// multiple values efficiently.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberFormat {
    sections: Vec<Section>,
}

impl NumberFormat {
    /// Create a NumberFormat from parsed sections.
    ///
    /// Limits to 4 sections maximum per Excel spec.
    pub fn from_sections(sections: Vec<Section>) -> Self {
        let sections = if sections.len() > 4 {
            sections.into_iter().take(4).collect()
        } else {
            sections
        };
        NumberFormat { sections }
    }

    /// Get the sections of this format.
    pub fn sections(&self) -> &[Section] {
        &self.sections
    }

    /// Returns true if this format contains date/time parts.
    pub fn is_date_format(&self) -> bool {
        self.sections.iter().any(|s| s.has_date_parts())
    }

    /// Returns true if this is a text-only format.
    pub fn is_text_format(&self) -> bool {
        self.sections.len() == 1 && self.sections[0].has_text_placeholder()
    }

    /// Returns true if this format contains a percent sign.
    pub fn is_percentage(&self) -> bool {
        self.sections.iter().any(|s| s.has_percent())
    }

    /// Returns true if any section has a color.
    pub fn has_color(&self) -> bool {
        self.sections.iter().any(|s| s.color.is_some())
    }

    /// Returns true if any section has a condition.
    pub fn has_condition(&self) -> bool {
        self.sections.iter().any(|s| s.condition.is_some())
    }
}
```

**Step 4: Update lib.rs to re-export NumberFormat**

The `pub use ast::NumberFormat;` is already in lib.rs.

**Step 5: Run test to verify it passes**

Run: `cargo test --test ast_tests`
Expected: PASS

**Step 6: Commit**

```bash
git add src/ast.rs tests/ast_tests.rs
git commit -m "feat: implement Section and NumberFormat AST types"
```

---

## Phase 4: Value Types and Options

### Task 4.1: Implement Value Enum

**Files:**
- Modify: `src/value.rs`
- Create: `tests/value_tests.rs`

**Step 1: Write the failing test**

Create `tests/value_tests.rs`:

```rust
use ssfmt::Value;

#[test]
fn test_value_from_f64() {
    let v: Value = 42.5.into();
    assert!(matches!(v, Value::Number(n) if (n - 42.5).abs() < f64::EPSILON));
}

#[test]
fn test_value_from_i64() {
    let v: Value = 42i64.into();
    assert!(matches!(v, Value::Number(n) if (n - 42.0).abs() < f64::EPSILON));
}

#[test]
fn test_value_from_str() {
    let v: Value = "hello".into();
    assert!(matches!(v, Value::Text(s) if s == "hello"));
}

#[test]
fn test_value_from_bool() {
    let v: Value = true.into();
    assert!(matches!(v, Value::Bool(true)));
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test value_tests`
Expected: FAIL - Value not found

**Step 3: Implement Value**

```rust
//! Value types that can be formatted.

/// A value that can be formatted using a number format code.
#[derive(Debug, Clone, PartialEq)]
pub enum Value<'a> {
    /// A numeric value (including Excel serial dates)
    Number(f64),
    /// A text value
    Text(&'a str),
    /// A boolean value
    Bool(bool),
    /// An empty cell
    Empty,
    /// A chrono DateTime (requires `chrono` feature)
    #[cfg(feature = "chrono")]
    DateTime(chrono::NaiveDateTime),
    /// A chrono Date (requires `chrono` feature)
    #[cfg(feature = "chrono")]
    Date(chrono::NaiveDate),
    /// A chrono Time (requires `chrono` feature)
    #[cfg(feature = "chrono")]
    Time(chrono::NaiveTime),
}

impl<'a> From<f64> for Value<'a> {
    fn from(n: f64) -> Self {
        Value::Number(n)
    }
}

impl<'a> From<f32> for Value<'a> {
    fn from(n: f32) -> Self {
        Value::Number(n as f64)
    }
}

impl<'a> From<i64> for Value<'a> {
    fn from(n: i64) -> Self {
        Value::Number(n as f64)
    }
}

impl<'a> From<i32> for Value<'a> {
    fn from(n: i32) -> Self {
        Value::Number(n as f64)
    }
}

impl<'a> From<&'a str> for Value<'a> {
    fn from(s: &'a str) -> Self {
        Value::Text(s)
    }
}

impl<'a> From<bool> for Value<'a> {
    fn from(b: bool) -> Self {
        Value::Bool(b)
    }
}

impl<'a> From<()> for Value<'a> {
    fn from(_: ()) -> Self {
        Value::Empty
    }
}

#[cfg(feature = "chrono")]
impl<'a> From<chrono::NaiveDateTime> for Value<'a> {
    fn from(dt: chrono::NaiveDateTime) -> Self {
        Value::DateTime(dt)
    }
}

#[cfg(feature = "chrono")]
impl<'a> From<chrono::NaiveDate> for Value<'a> {
    fn from(d: chrono::NaiveDate) -> Self {
        Value::Date(d)
    }
}

#[cfg(feature = "chrono")]
impl<'a> From<chrono::NaiveTime> for Value<'a> {
    fn from(t: chrono::NaiveTime) -> Self {
        Value::Time(t)
    }
}

impl<'a> Value<'a> {
    /// Returns the value as a number if possible.
    pub fn as_number(&self) -> Option<f64> {
        match self {
            Value::Number(n) => Some(*n),
            Value::Bool(true) => Some(1.0),
            Value::Bool(false) => Some(0.0),
            _ => None,
        }
    }

    /// Returns the value as text if it is text.
    pub fn as_text(&self) -> Option<&'a str> {
        match self {
            Value::Text(s) => Some(s),
            _ => None,
        }
    }

    /// Returns true if this value is empty.
    pub fn is_empty(&self) -> bool {
        matches!(self, Value::Empty)
    }

    /// Returns a type name for error messages.
    pub fn type_name(&self) -> &'static str {
        match self {
            Value::Number(_) => "number",
            Value::Text(_) => "text",
            Value::Bool(_) => "boolean",
            Value::Empty => "empty",
            #[cfg(feature = "chrono")]
            Value::DateTime(_) => "datetime",
            #[cfg(feature = "chrono")]
            Value::Date(_) => "date",
            #[cfg(feature = "chrono")]
            Value::Time(_) => "time",
        }
    }
}
```

**Step 4: Run test to verify it passes**

Run: `cargo test --test value_tests`
Expected: PASS

**Step 5: Commit**

```bash
git add src/value.rs tests/value_tests.rs
git commit -m "feat: implement Value enum with From conversions"
```

---

### Task 4.2: Implement FormatOptions

**Files:**
- Modify: `src/options.rs`
- Create: `tests/options_tests.rs`

**Step 1: Write the failing test**

Create `tests/options_tests.rs`:

```rust
use ssfmt::{DateSystem, FormatOptions};

#[test]
fn test_default_options() {
    let opts = FormatOptions::default();
    assert_eq!(opts.date_system, DateSystem::Date1900);
}

#[test]
fn test_date_system_epoch() {
    // 1900 system: day 1 = Jan 1, 1900
    // 1904 system: day 0 = Jan 1, 1904
    assert_eq!(DateSystem::Date1900.epoch_year(), 1900);
    assert_eq!(DateSystem::Date1904.epoch_year(), 1904);
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test options_tests`
Expected: FAIL - types not found

**Step 3: Implement FormatOptions**

```rust
//! Formatting options and configuration.

use crate::locale::Locale;

/// The date system used for serial number conversion.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DateSystem {
    /// Windows Excel default (1900-based, includes leap year bug)
    #[default]
    Date1900,
    /// Mac Excel legacy (1904-based)
    Date1904,
}

impl DateSystem {
    /// Returns the epoch year for this date system.
    pub fn epoch_year(&self) -> i32 {
        match self {
            DateSystem::Date1900 => 1900,
            DateSystem::Date1904 => 1904,
        }
    }
}

/// Options for formatting values.
#[derive(Debug, Clone)]
pub struct FormatOptions {
    /// The date system to use for serial number conversion.
    pub date_system: DateSystem,
    /// The locale for formatting.
    pub locale: Locale,
}

impl Default for FormatOptions {
    fn default() -> Self {
        FormatOptions {
            date_system: DateSystem::default(),
            locale: Locale::default(),
        }
    }
}
```

**Step 4: Create minimal Locale in locale/mod.rs**

```rust
//! Locale support for formatting.

mod builtin;

pub use builtin::Locale;
```

And in `src/locale/builtin.rs`:

```rust
//! Built-in locale data.

/// Locale settings for formatting.
#[derive(Debug, Clone)]
pub struct Locale {
    /// Decimal separator character
    pub decimal_separator: char,
    /// Thousands separator character
    pub thousands_separator: char,
    /// Currency symbol
    pub currency_symbol: &'static str,
    /// AM string
    pub am_string: &'static str,
    /// PM string
    pub pm_string: &'static str,
    /// Short month names
    pub month_names_short: [&'static str; 12],
    /// Full month names
    pub month_names_full: [&'static str; 12],
    /// Short day names
    pub day_names_short: [&'static str; 7],
    /// Full day names
    pub day_names_full: [&'static str; 7],
}

impl Default for Locale {
    fn default() -> Self {
        Self::en_us()
    }
}

impl Locale {
    /// US English locale.
    pub fn en_us() -> Self {
        Locale {
            decimal_separator: '.',
            thousands_separator: ',',
            currency_symbol: "$",
            am_string: "AM",
            pm_string: "PM",
            month_names_short: [
                "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            ],
            month_names_full: [
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December",
            ],
            day_names_short: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
            day_names_full: [
                "Sunday", "Monday", "Tuesday", "Wednesday",
                "Thursday", "Friday", "Saturday",
            ],
        }
    }
}
```

**Step 5: Update lib.rs exports**

Add to lib.rs:

```rust
pub use options::DateSystem;
```

**Step 6: Run test to verify it passes**

Run: `cargo test --test options_tests`
Expected: PASS

**Step 7: Commit**

```bash
git add src/options.rs src/locale/mod.rs src/locale/builtin.rs tests/options_tests.rs
git commit -m "feat: implement FormatOptions, DateSystem, and basic Locale"
```

---

## Phase 5: Date Serial Conversion

### Task 5.1: Implement Excel Serial to Date Conversion

**Files:**
- Modify: `src/date_serial.rs`
- Create: `tests/date_serial_tests.rs`

**Step 1: Write the failing test**

Create `tests/date_serial_tests.rs`:

```rust
use ssfmt::date_serial::{serial_to_date, date_to_serial};
use ssfmt::DateSystem;

#[test]
fn test_serial_to_date_1900_basic() {
    // Day 1 = January 1, 1900
    let (y, m, d) = serial_to_date(1.0, DateSystem::Date1900).unwrap();
    assert_eq!((y, m, d), (1900, 1, 1));
}

#[test]
fn test_serial_to_date_1900_day_60() {
    // Day 60 = February 29, 1900 (Excel's bug - this date doesn't exist)
    let (y, m, d) = serial_to_date(60.0, DateSystem::Date1900).unwrap();
    assert_eq!((y, m, d), (1900, 2, 29));
}

#[test]
fn test_serial_to_date_1900_day_61() {
    // Day 61 = March 1, 1900
    let (y, m, d) = serial_to_date(61.0, DateSystem::Date1900).unwrap();
    assert_eq!((y, m, d), (1900, 3, 1));
}

#[test]
fn test_serial_to_date_known_date() {
    // January 9, 2026 should be serial 46031 in 1900 system
    let (y, m, d) = serial_to_date(46031.0, DateSystem::Date1900).unwrap();
    assert_eq!((y, m, d), (2026, 1, 9));
}

#[test]
fn test_serial_to_time() {
    // 0.5 = 12:00:00 (noon)
    let (h, m, s) = ssfmt::date_serial::serial_to_time(0.5);
    assert_eq!((h, m, s), (12, 0, 0));

    // 0.75 = 18:00:00 (6 PM)
    let (h, m, s) = ssfmt::date_serial::serial_to_time(0.75);
    assert_eq!((h, m, s), (18, 0, 0));
}

#[test]
fn test_date_to_serial() {
    let serial = date_to_serial(2026, 1, 9, DateSystem::Date1900);
    assert!((serial - 46031.0).abs() < 0.0001);
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test date_serial_tests`
Expected: FAIL - module not found

**Step 3: Implement date_serial module**

```rust
//! Excel serial number to date/time conversion.
//!
//! Excel stores dates as serial numbers:
//! - 1900 system: Day 1 = January 1, 1900
//! - 1904 system: Day 0 = January 1, 1904
//!
//! Note: The 1900 system has a bug where February 29, 1900 exists (day 60),
//! even though 1900 was not a leap year. We replicate this bug for compatibility.

use crate::options::DateSystem;

/// Convert an Excel serial number to a date (year, month, day).
///
/// Returns None if the serial number is out of range.
pub fn serial_to_date(serial: f64, system: DateSystem) -> Option<(i32, u32, u32)> {
    if serial < 0.0 {
        return None;
    }

    let day_num = serial.trunc() as i64;

    match system {
        DateSystem::Date1900 => serial_to_date_1900(day_num),
        DateSystem::Date1904 => serial_to_date_1904(day_num),
    }
}

fn serial_to_date_1900(day_num: i64) -> Option<(i32, u32, u32)> {
    if day_num < 1 {
        // Day 0 is sometimes used for time-only values
        return Some((1900, 1, 0)); // Special case
    }

    // Handle Excel's leap year bug: day 60 = Feb 29, 1900
    if day_num == 60 {
        return Some((1900, 2, 29));
    }

    // Adjust for the bug: days after 60 are shifted by 1
    let adjusted = if day_num > 60 { day_num - 1 } else { day_num };

    // Convert to days since March 1, year 0 (a common algorithm base)
    // Day 1 in Excel = Jan 1, 1900
    // We use the algorithm from Howard Hinnant's date library

    // Days from year 0 to Jan 1, 1900
    const DAYS_TO_1900: i64 = 693960; // Calculated: days from 0000-03-01 to 1900-01-01

    let z = adjusted + DAYS_TO_1900 - 1; // -1 because Excel day 1 = Jan 1
    let era = (if z >= 0 { z } else { z - 146096 }) / 146097;
    let doe = (z - era * 146097) as u32;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };

    Some((y as i32, m, d))
}

fn serial_to_date_1904(day_num: i64) -> Option<(i32, u32, u32)> {
    if day_num < 0 {
        return None;
    }

    // Day 0 in 1904 system = Jan 1, 1904
    // No leap year bug in this system

    const DAYS_TO_1904: i64 = 695422; // Days from 0000-03-01 to 1904-01-01

    let z = day_num + DAYS_TO_1904;
    let era = (if z >= 0 { z } else { z - 146096 }) / 146097;
    let doe = (z - era * 146097) as u32;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };

    Some((y as i32, m, d))
}

/// Extract time components from an Excel serial number.
///
/// Returns (hour, minute, second) where the fractional part represents time.
pub fn serial_to_time(serial: f64) -> (u32, u32, u32) {
    let frac = serial.fract().abs();
    let total_seconds = (frac * 86400.0).round() as u32;

    let hours = total_seconds / 3600;
    let minutes = (total_seconds % 3600) / 60;
    let seconds = total_seconds % 60;

    (hours, minutes, seconds)
}

/// Extract subsecond fraction from an Excel serial number.
pub fn serial_to_subseconds(serial: f64) -> f64 {
    let frac = serial.fract().abs();
    let total_seconds = frac * 86400.0;
    total_seconds.fract()
}

/// Convert a date to an Excel serial number.
pub fn date_to_serial(year: i32, month: u32, day: u32, system: DateSystem) -> f64 {
    match system {
        DateSystem::Date1900 => date_to_serial_1900(year, month, day),
        DateSystem::Date1904 => date_to_serial_1904(year, month, day),
    }
}

fn date_to_serial_1900(year: i32, month: u32, day: u32) -> f64 {
    // Handle the special Feb 29, 1900 case
    if year == 1900 && month == 2 && day == 29 {
        return 60.0;
    }

    let y = if month <= 2 {
        year as i64 - 1
    } else {
        year as i64
    };
    let m = if month <= 2 { month + 12 } else { month } as i64;
    let d = day as i64;

    let era = (if y >= 0 { y } else { y - 399 }) / 400;
    let yoe = (y - era * 400) as u32;
    let doy = (153 * (m - 3) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy as u32;
    let days = era * 146097 + doe as i64;

    const DAYS_TO_1900: i64 = 693960;
    let serial = days - DAYS_TO_1900 + 1; // +1 because Excel day 1 = Jan 1

    // Adjust for leap year bug
    let serial = if serial > 60 { serial + 1 } else { serial };

    serial as f64
}

fn date_to_serial_1904(year: i32, month: u32, day: u32) -> f64 {
    let y = if month <= 2 {
        year as i64 - 1
    } else {
        year as i64
    };
    let m = if month <= 2 { month + 12 } else { month } as i64;
    let d = day as i64;

    let era = (if y >= 0 { y } else { y - 399 }) / 400;
    let yoe = (y - era * 400) as u32;
    let doy = (153 * (m - 3) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy as u32;
    let days = era * 146097 + doe as i64;

    const DAYS_TO_1904: i64 = 695422;
    (days - DAYS_TO_1904) as f64
}

/// Get the day of week (0 = Sunday, 6 = Saturday) from a serial number.
pub fn serial_to_weekday(serial: f64, system: DateSystem) -> u32 {
    let day_num = serial.trunc() as i64;

    match system {
        // In 1900 system, day 1 (Jan 1, 1900) was a Sunday
        // But due to the leap year bug, we need to adjust after day 60
        DateSystem::Date1900 => {
            if day_num <= 60 {
                ((day_num - 1) % 7) as u32
            } else {
                ((day_num - 2) % 7) as u32
            }
        }
        // In 1904 system, day 0 (Jan 1, 1904) was a Friday
        DateSystem::Date1904 => ((day_num + 5) % 7) as u32,
    }
}
```

**Step 4: Make date_serial module public**

In `src/lib.rs`, change:
```rust
mod date_serial;
```
to:
```rust
pub mod date_serial;
```

**Step 5: Run test to verify it passes**

Run: `cargo test --test date_serial_tests`
Expected: PASS

**Step 6: Commit**

```bash
git add src/date_serial.rs tests/date_serial_tests.rs src/lib.rs
git commit -m "feat: implement Excel serial number to date/time conversion"
```

---

## Phase 6: Lexer and Tokens

### Task 6.1: Implement Token Types

**Files:**
- Modify: `src/parser/tokens.rs`

**Step 1: Implement token types**

```rust
//! Token types for the format code lexer.

/// A token in a format code string.
#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    // Literals
    Literal(char),
    EscapedChar(char),
    QuotedString(String),

    // Digit placeholders
    Zero,       // 0
    Hash,       // #
    Question,   // ?

    // Separators
    DecimalPoint,      // .
    ThousandsSep,      // ,
    SectionSep,        // ;

    // Special characters
    Percent,           // %
    At,                // @
    Asterisk,          // *
    Underscore,        // _

    // Scientific notation
    ExponentUpper,     // E
    ExponentLower,     // e
    Plus,              // +
    Minus,             // -

    // Fraction
    Slash,             // /

    // Date/time
    Year,              // y
    Month,             // m
    Day,               // d
    Hour,              // h
    Second,            // s

    // Brackets
    OpenBracket,       // [
    CloseBracket,      // ]

    // AM/PM markers
    AmPm(String),      // AM/PM, am/pm, A/P, a/p

    // End of input
    Eof,
}

/// A token with its position in the source.
#[derive(Debug, Clone)]
pub struct SpannedToken {
    pub token: Token,
    pub start: usize,
    pub end: usize,
}
```

**Step 2: Commit**

```bash
git add src/parser/tokens.rs
git commit -m "feat: define token types for format code lexer"
```

---

### Task 6.2: Implement Lexer

**Files:**
- Modify: `src/parser/lexer.rs`
- Create: `tests/lexer_tests.rs`

**Step 1: Write the failing test**

Create `tests/lexer_tests.rs`:

```rust
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
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('-'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Month);
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('-'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Day);
    assert_eq!(lexer.next_token().unwrap().token, Token::Day);
}

#[test]
fn test_lex_quoted_string() {
    let mut lexer = Lexer::new("\"USD\"0.00");

    assert_eq!(lexer.next_token().unwrap().token, Token::QuotedString("USD".into()));
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
}

#[test]
fn test_lex_escaped_char() {
    let mut lexer = Lexer::new("\\$0.00");

    assert_eq!(lexer.next_token().unwrap().token, Token::EscapedChar('$'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
}

#[test]
fn test_lex_bracket() {
    let mut lexer = Lexer::new("[Red]0");

    assert_eq!(lexer.next_token().unwrap().token, Token::OpenBracket);
    // Inside bracket, letters are literals
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('R'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('e'));
    assert_eq!(lexer.next_token().unwrap().token, Token::Literal('d'));
    assert_eq!(lexer.next_token().unwrap().token, Token::CloseBracket);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
}

#[test]
fn test_lex_sections() {
    let mut lexer = Lexer::new("0;-0");

    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
    assert_eq!(lexer.next_token().unwrap().token, Token::SectionSep);
    assert_eq!(lexer.next_token().unwrap().token, Token::Minus);
    assert_eq!(lexer.next_token().unwrap().token, Token::Zero);
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test lexer_tests`
Expected: FAIL - module not found

**Step 3: Implement Lexer**

```rust
//! Lexer for format code strings.

use crate::error::ParseError;
use super::tokens::{Token, SpannedToken};

/// Lexer for format code strings.
pub struct Lexer<'a> {
    input: &'a str,
    chars: std::iter::Peekable<std::str::CharIndices<'a>>,
    position: usize,
    in_bracket: bool,
}

impl<'a> Lexer<'a> {
    /// Create a new lexer for the given input.
    pub fn new(input: &'a str) -> Self {
        Lexer {
            input,
            chars: input.char_indices().peekable(),
            position: 0,
            in_bracket: false,
        }
    }

    /// Get the next token.
    pub fn next_token(&mut self) -> Result<SpannedToken, ParseError> {
        let start = self.position;

        let Some((pos, ch)) = self.chars.next() else {
            return Ok(SpannedToken {
                token: Token::Eof,
                start,
                end: start,
            });
        };

        self.position = pos + ch.len_utf8();

        let token = match ch {
            // Escape sequences
            '\\' => {
                if let Some((_, escaped)) = self.chars.next() {
                    self.position += escaped.len_utf8();
                    Token::EscapedChar(escaped)
                } else {
                    return Err(ParseError::UnexpectedToken {
                        position: pos,
                        found: ch,
                    });
                }
            }

            // Quoted strings
            '"' => self.lex_quoted_string(pos)?,

            // Digit placeholders
            '0' => Token::Zero,
            '#' => Token::Hash,
            '?' => Token::Question,

            // Separators
            '.' => Token::DecimalPoint,
            ',' => Token::ThousandsSep,
            ';' => Token::SectionSep,

            // Special
            '%' => Token::Percent,
            '@' => Token::At,
            '*' => Token::Asterisk,
            '_' => Token::Underscore,

            // Scientific/math
            '+' => Token::Plus,
            '-' => Token::Minus,
            '/' => Token::Slash,

            // Exponent markers
            'E' => Token::ExponentUpper,
            'e' => Token::ExponentLower,

            // Brackets
            '[' => {
                self.in_bracket = true;
                Token::OpenBracket
            }
            ']' => {
                self.in_bracket = false;
                Token::CloseBracket
            }

            // Date/time (only outside brackets as format specifiers)
            'y' | 'Y' if !self.in_bracket => Token::Year,
            'm' | 'M' if !self.in_bracket => Token::Month,
            'd' | 'D' if !self.in_bracket => Token::Day,
            'h' | 'H' if !self.in_bracket => Token::Hour,
            's' | 'S' if !self.in_bracket => Token::Second,

            // AM/PM detection
            'A' | 'a' if !self.in_bracket => self.try_lex_ampm(ch),

            // Everything else is a literal
            _ => Token::Literal(ch),
        };

        Ok(SpannedToken {
            token,
            start,
            end: self.position,
        })
    }

    fn lex_quoted_string(&mut self, start: usize) -> Result<Token, ParseError> {
        let mut s = String::new();

        loop {
            match self.chars.next() {
                Some((_, '"')) => {
                    self.position += 1;
                    break;
                }
                Some((_, ch)) => {
                    self.position += ch.len_utf8();
                    s.push(ch);
                }
                None => {
                    return Err(ParseError::UnterminatedBracket { position: start });
                }
            }
        }

        Ok(Token::QuotedString(s))
    }

    fn try_lex_ampm(&mut self, first: char) -> Token {
        // Check for AM/PM, A/P patterns
        let remaining: String = self.input[self.position..].chars().take(4).collect();

        if remaining.to_lowercase().starts_with("m/pm") {
            // AM/PM or am/pm
            for _ in 0..4 {
                if let Some((_, ch)) = self.chars.next() {
                    self.position += ch.len_utf8();
                }
            }
            let full = format!("{}{}", first, &remaining[..4]);
            Token::AmPm(full)
        } else if remaining.to_lowercase().starts_with("/p") {
            // A/P or a/p
            for _ in 0..2 {
                if let Some((_, ch)) = self.chars.next() {
                    self.position += ch.len_utf8();
                }
            }
            let full = format!("{}{}", first, &remaining[..2]);
            Token::AmPm(full)
        } else {
            Token::Literal(first)
        }
    }

    /// Peek at the next character without consuming it.
    pub fn peek(&mut self) -> Option<char> {
        self.chars.peek().map(|(_, ch)| *ch)
    }

    /// Get the current position in the input.
    pub fn position(&self) -> usize {
        self.position
    }
}
```

**Step 4: Make parser module public and update mod.rs**

In `src/parser/mod.rs`:

```rust
//! Parser for format code strings.

pub mod lexer;
pub mod tokens;
```

In `src/lib.rs`, change:
```rust
mod parser;
```
to:
```rust
pub mod parser;
```

**Step 5: Run test to verify it passes**

Run: `cargo test --test lexer_tests`
Expected: PASS

**Step 6: Commit**

```bash
git add src/parser/ tests/lexer_tests.rs src/lib.rs
git commit -m "feat: implement lexer for format codes"
```

---

## Phase 7: Parser

### Task 7.1: Implement Format Code Parser

**Files:**
- Modify: `src/parser/mod.rs`
- Create: `tests/parser_tests.rs`

**Step 1: Write the failing test**

Create `tests/parser_tests.rs`:

```rust
use ssfmt::NumberFormat;
use ssfmt::ast::{FormatPart, DigitPlaceholder, DatePart, Color, NamedColor};

#[test]
fn test_parse_simple_number() {
    let fmt = NumberFormat::parse("#,##0.00").unwrap();
    assert_eq!(fmt.sections().len(), 1);
    assert!(!fmt.is_date_format());
}

#[test]
fn test_parse_date_format() {
    let fmt = NumberFormat::parse("yyyy-mm-dd").unwrap();
    assert_eq!(fmt.sections().len(), 1);
    assert!(fmt.is_date_format());
}

#[test]
fn test_parse_multiple_sections() {
    let fmt = NumberFormat::parse("#,##0;-#,##0;0").unwrap();
    assert_eq!(fmt.sections().len(), 3);
}

#[test]
fn test_parse_color() {
    let fmt = NumberFormat::parse("[Red]0").unwrap();
    assert!(fmt.has_color());
    assert_eq!(fmt.sections()[0].color, Some(Color::Named(NamedColor::Red)));
}

#[test]
fn test_parse_percentage() {
    let fmt = NumberFormat::parse("0%").unwrap();
    assert!(fmt.is_percentage());
}

#[test]
fn test_parse_text_format() {
    let fmt = NumberFormat::parse("@").unwrap();
    assert!(fmt.is_text_format());
}

#[test]
fn test_parse_too_many_sections() {
    let result = NumberFormat::parse("0;0;0;0;0");
    // Should succeed but truncate to 4 sections
    let fmt = result.unwrap();
    assert_eq!(fmt.sections().len(), 4);
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test parser_tests`
Expected: FAIL - NumberFormat::parse not found

**Step 3: Implement parser**

Update `src/parser/mod.rs`:

```rust
//! Parser for format code strings.

pub mod lexer;
pub mod tokens;

use crate::ast::*;
use crate::error::ParseError;
use lexer::Lexer;
use tokens::Token;

/// Parse a format code string into a NumberFormat.
pub fn parse(format_code: &str) -> Result<NumberFormat, ParseError> {
    if format_code.is_empty() {
        return Err(ParseError::EmptyFormat);
    }

    let mut lexer = Lexer::new(format_code);
    let mut sections = Vec::new();
    let mut current_section = SectionBuilder::new();

    loop {
        let spanned = lexer.next_token()?;

        match spanned.token {
            Token::Eof => {
                sections.push(current_section.build());
                break;
            }
            Token::SectionSep => {
                sections.push(current_section.build());
                current_section = SectionBuilder::new();
            }
            Token::OpenBracket => {
                parse_bracket_content(&mut lexer, &mut current_section)?;
            }
            token => {
                parse_format_token(token, &mut lexer, &mut current_section)?;
            }
        }
    }

    if sections.len() > 4 {
        sections.truncate(4);
    }

    Ok(NumberFormat::from_sections(sections))
}

fn parse_bracket_content(
    lexer: &mut Lexer,
    section: &mut SectionBuilder,
) -> Result<(), ParseError> {
    let start = lexer.position();
    let mut content = String::new();

    loop {
        let spanned = lexer.next_token()?;
        match spanned.token {
            Token::CloseBracket => break,
            Token::Literal(ch) => content.push(ch),
            Token::Eof => {
                return Err(ParseError::UnterminatedBracket { position: start });
            }
            _ => {
                // Other tokens inside brackets become literals
                content.push_str(&format!("{:?}", spanned.token));
            }
        }
    }

    // Parse bracket content
    if let Some(color) = try_parse_color(&content) {
        section.color = Some(color);
    } else if let Some(condition) = try_parse_condition(&content) {
        section.condition = Some(condition);
    } else if let Some(elapsed) = try_parse_elapsed(&content) {
        section.parts.push(FormatPart::Elapsed(elapsed));
    } else if let Some(locale) = try_parse_locale(&content) {
        section.parts.push(FormatPart::Locale(locale));
    }

    Ok(())
}

fn try_parse_color(content: &str) -> Option<Color> {
    // Try named colors
    if let Ok(named) = content.parse::<NamedColor>() {
        return Some(Color::Named(named));
    }

    // Try indexed colors (Color1 - Color56)
    if content.to_lowercase().starts_with("color") {
        if let Ok(n) = content[5..].parse::<u8>() {
            if (1..=56).contains(&n) {
                return Some(Color::Indexed(n));
            }
        }
    }

    None
}

fn try_parse_condition(content: &str) -> Option<Condition> {
    let content = content.trim();

    if let Some(rest) = content.strip_prefix(">=") {
        rest.trim().parse::<f64>().ok().map(Condition::GreaterOrEqual)
    } else if let Some(rest) = content.strip_prefix("<=") {
        rest.trim().parse::<f64>().ok().map(Condition::LessOrEqual)
    } else if let Some(rest) = content.strip_prefix("<>") {
        rest.trim().parse::<f64>().ok().map(Condition::NotEqual)
    } else if let Some(rest) = content.strip_prefix('>') {
        rest.trim().parse::<f64>().ok().map(Condition::GreaterThan)
    } else if let Some(rest) = content.strip_prefix('<') {
        rest.trim().parse::<f64>().ok().map(Condition::LessThan)
    } else if let Some(rest) = content.strip_prefix('=') {
        rest.trim().parse::<f64>().ok().map(Condition::Equal)
    } else {
        None
    }
}

fn try_parse_elapsed(content: &str) -> Option<ElapsedPart> {
    match content.to_lowercase().as_str() {
        "h" => Some(ElapsedPart::Hours),
        "m" => Some(ElapsedPart::Minutes),
        "s" => Some(ElapsedPart::Seconds),
        _ => None,
    }
}

fn try_parse_locale(content: &str) -> Option<LocaleCode> {
    if !content.starts_with('$') {
        return None;
    }

    let content = &content[1..]; // Skip $

    // Format: [$<currency>-<lcid>] or [$-<lcid>]
    if let Some(dash_pos) = content.rfind('-') {
        let currency_part = &content[..dash_pos];
        let lcid_part = &content[dash_pos + 1..];

        let currency = if currency_part.is_empty() {
            None
        } else {
            Some(currency_part.to_string())
        };

        let lcid = u32::from_str_radix(lcid_part, 16).ok();

        Some(LocaleCode { currency, lcid })
    } else {
        // Just currency, no LCID
        Some(LocaleCode {
            currency: Some(content.to_string()),
            lcid: None,
        })
    }
}

fn parse_format_token(
    token: Token,
    lexer: &mut Lexer,
    section: &mut SectionBuilder,
) -> Result<(), ParseError> {
    match token {
        Token::Zero => section.parts.push(FormatPart::Digit(DigitPlaceholder::Zero)),
        Token::Hash => section.parts.push(FormatPart::Digit(DigitPlaceholder::Hash)),
        Token::Question => section.parts.push(FormatPart::Digit(DigitPlaceholder::Question)),

        Token::DecimalPoint => section.parts.push(FormatPart::DecimalPoint),
        Token::ThousandsSep => section.parts.push(FormatPart::ThousandsSeparator),
        Token::Percent => section.parts.push(FormatPart::Percent),
        Token::At => section.parts.push(FormatPart::TextPlaceholder),

        Token::Literal(ch) => section.push_literal(ch),
        Token::EscapedChar(ch) => section.push_literal(ch),
        Token::QuotedString(s) => {
            for ch in s.chars() {
                section.push_literal(ch);
            }
        }

        Token::Asterisk => {
            // Fill character: next char is the fill
            if let Some(ch) = lexer.peek() {
                let _ = lexer.next_token();
                section.parts.push(FormatPart::Fill(ch));
            }
        }

        Token::Underscore => {
            // Skip width: next char determines width
            if let Some(ch) = lexer.peek() {
                let _ = lexer.next_token();
                section.parts.push(FormatPart::Skip(ch));
            }
        }

        Token::Year => {
            let count = 1 + count_consecutive(lexer, Token::Year);
            let part = if count >= 4 {
                DatePart::Year4
            } else {
                DatePart::Year2
            };
            section.parts.push(FormatPart::DatePart(part));
        }

        Token::Month => {
            let count = 1 + count_consecutive(lexer, Token::Month);
            let part = match count {
                1 => DatePart::Month,
                2 => DatePart::Month2,
                3 => DatePart::MonthAbbr,
                4 => DatePart::MonthFull,
                _ => DatePart::MonthLetter,
            };
            // Check context: if preceded by hour, this is minute
            if section.last_was_hour() {
                let part = if count >= 2 {
                    DatePart::Minute2
                } else {
                    DatePart::Minute
                };
                section.parts.push(FormatPart::DatePart(part));
            } else {
                section.parts.push(FormatPart::DatePart(part));
            }
        }

        Token::Day => {
            let count = 1 + count_consecutive(lexer, Token::Day);
            let part = match count {
                1 => DatePart::Day,
                2 => DatePart::Day2,
                3 => DatePart::DayAbbr,
                _ => DatePart::DayFull,
            };
            section.parts.push(FormatPart::DatePart(part));
        }

        Token::Hour => {
            let count = 1 + count_consecutive(lexer, Token::Hour);
            let part = if count >= 2 {
                DatePart::Hour2
            } else {
                DatePart::Hour
            };
            section.parts.push(FormatPart::DatePart(part));
        }

        Token::Second => {
            let count = 1 + count_consecutive(lexer, Token::Second);
            let part = if count >= 2 {
                DatePart::Second2
            } else {
                DatePart::Second
            };
            section.parts.push(FormatPart::DatePart(part));
        }

        Token::AmPm(s) => {
            let style = match s.as_str() {
                "AM/PM" => AmPmStyle::Upper,
                "am/pm" => AmPmStyle::Lower,
                "A/P" => AmPmStyle::ShortUpper,
                "a/p" => AmPmStyle::ShortLower,
                _ => AmPmStyle::Upper,
            };
            section.parts.push(FormatPart::AmPm(style));
        }

        Token::ExponentUpper | Token::ExponentLower => {
            let upper = matches!(token, Token::ExponentUpper);
            // Look for + or -
            if let Some(spanned) = peek_token(lexer) {
                let show_plus = matches!(spanned.token, Token::Plus);
                if matches!(spanned.token, Token::Plus | Token::Minus) {
                    let _ = lexer.next_token();
                }
                section.parts.push(FormatPart::Scientific { upper, show_plus });
            }
        }

        Token::Plus | Token::Minus => {
            // As literal when not part of exponent
            let ch = if matches!(token, Token::Plus) { '+' } else { '-' };
            section.push_literal(ch);
        }

        Token::Slash => {
            // Could be part of fraction - for now, just literal
            section.push_literal('/');
        }

        _ => {}
    }

    Ok(())
}

fn count_consecutive(lexer: &mut Lexer, target: Token) -> usize {
    let mut count = 0;
    while let Some(spanned) = peek_token(lexer) {
        if std::mem::discriminant(&spanned.token) == std::mem::discriminant(&target) {
            let _ = lexer.next_token();
            count += 1;
        } else {
            break;
        }
    }
    count
}

fn peek_token(lexer: &mut Lexer) -> Option<tokens::SpannedToken> {
    // This is a simplified peek - in a real implementation we'd want proper lookahead
    None // Simplified for now
}

/// Helper for building sections.
struct SectionBuilder {
    condition: Option<Condition>,
    color: Option<Color>,
    parts: Vec<FormatPart>,
}

impl SectionBuilder {
    fn new() -> Self {
        SectionBuilder {
            condition: None,
            color: None,
            parts: Vec::new(),
        }
    }

    fn build(self) -> Section {
        Section {
            condition: self.condition,
            color: self.color,
            parts: self.parts,
        }
    }

    fn push_literal(&mut self, ch: char) {
        // Merge consecutive literals
        if let Some(FormatPart::Literal(s)) = self.parts.last_mut() {
            s.push(ch);
        } else {
            self.parts.push(FormatPart::Literal(ch.to_string()));
        }
    }

    fn last_was_hour(&self) -> bool {
        self.parts.iter().rev().find_map(|p| {
            match p {
                FormatPart::DatePart(DatePart::Hour) | FormatPart::DatePart(DatePart::Hour2) => Some(true),
                FormatPart::DatePart(_) => Some(false),
                FormatPart::Literal(_) => None, // Skip literals
                _ => Some(false),
            }
        }).unwrap_or(false)
    }
}

// Add parse method to NumberFormat
impl NumberFormat {
    /// Parse a format code string.
    pub fn parse(format_code: &str) -> Result<Self, ParseError> {
        parse(format_code)
    }
}
```

**Step 4: Run test to verify it passes**

Run: `cargo test --test parser_tests`
Expected: PASS

**Step 5: Commit**

```bash
git add src/parser/mod.rs src/ast.rs tests/parser_tests.rs
git commit -m "feat: implement format code parser"
```

---

## Phase 8: Basic Number Formatter

### Task 8.1: Implement Number Formatting

**Files:**
- Modify: `src/formatter/mod.rs`
- Modify: `src/formatter/number.rs`
- Create: `tests/format_numbers.rs`

**Step 1: Write the failing test**

Create `tests/format_numbers.rs`:

```rust
use ssfmt::{NumberFormat, FormatOptions};

#[test]
fn test_format_integer() {
    let fmt = NumberFormat::parse("0").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(42.0, &opts), "42");
    assert_eq!(fmt.format(42.7, &opts), "43"); // Rounds
}

#[test]
fn test_format_decimal() {
    let fmt = NumberFormat::parse("0.00").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(42.0, &opts), "42.00");
    assert_eq!(fmt.format(42.567, &opts), "42.57");
}

#[test]
fn test_format_thousands() {
    let fmt = NumberFormat::parse("#,##0").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(1234567.0, &opts), "1,234,567");
    assert_eq!(fmt.format(123.0, &opts), "123");
}

#[test]
fn test_format_percentage() {
    let fmt = NumberFormat::parse("0%").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(0.42, &opts), "42%");
    assert_eq!(fmt.format(1.5, &opts), "150%");
}

#[test]
fn test_format_hash_placeholder() {
    let fmt = NumberFormat::parse("#.##").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(42.5, &opts), "42.5");
    assert_eq!(fmt.format(42.0, &opts), "42.");
}

#[test]
fn test_format_negative_section() {
    let fmt = NumberFormat::parse("0;-0").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(42.0, &opts), "42");
    assert_eq!(fmt.format(-42.0, &opts), "-42");
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test format_numbers`
Expected: FAIL - format method not found

**Step 3: Implement formatter**

In `src/formatter/mod.rs`:

```rust
//! Formatting implementation.

mod number;
mod date;
mod text;
mod fraction;

use crate::ast::{NumberFormat, Section, FormatPart, DigitPlaceholder};
use crate::error::FormatError;
use crate::options::FormatOptions;
use crate::value::Value;

impl NumberFormat {
    /// Format a value using this format code.
    ///
    /// This method never fails - it falls back to a general format on errors.
    pub fn format<'a>(&self, value: impl Into<Value<'a>>, opts: &FormatOptions) -> String {
        let value = value.into();
        self.try_format(value.clone(), opts)
            .unwrap_or_else(|_| fallback_format(&value))
    }

    /// Format a value, returning an error on type mismatch.
    pub fn try_format<'a>(
        &self,
        value: impl Into<Value<'a>>,
        opts: &FormatOptions,
    ) -> Result<String, FormatError> {
        let value = value.into();

        // Handle empty values
        if value.is_empty() {
            return Ok(String::new());
        }

        // Handle text values
        if let Value::Text(s) = &value {
            return Ok(self.format_text(s, opts));
        }

        // Get numeric value
        let num = value.as_number().ok_or(FormatError::TypeMismatch {
            expected: "number",
            got: value.type_name(),
        })?;

        // Select section based on value
        let section = self.select_section(num);

        // Format based on section contents
        if section.has_date_parts() {
            date::format_date(num, section, opts)
        } else {
            Ok(number::format_number(num, section, opts))
        }
    }

    fn select_section(&self, value: f64) -> &Section {
        match self.sections().len() {
            0 => unreachable!("NumberFormat always has at least one section"),
            1 => &self.sections()[0],
            2 => {
                if value >= 0.0 {
                    &self.sections()[0]
                } else {
                    &self.sections()[1]
                }
            }
            3 => {
                if value > 0.0 {
                    &self.sections()[0]
                } else if value < 0.0 {
                    &self.sections()[1]
                } else {
                    &self.sections()[2]
                }
            }
            _ => {
                if value > 0.0 {
                    &self.sections()[0]
                } else if value < 0.0 {
                    &self.sections()[1]
                } else {
                    &self.sections()[2]
                }
            }
        }
    }

    fn format_text(&self, s: &str, _opts: &FormatOptions) -> String {
        // Find text section (4th section or section with @)
        let section = if self.sections().len() == 4 {
            &self.sections()[3]
        } else {
            self.sections()
                .iter()
                .find(|s| s.has_text_placeholder())
                .unwrap_or(&self.sections()[0])
        };

        let mut result = String::new();
        for part in &section.parts {
            match part {
                FormatPart::TextPlaceholder => result.push_str(s),
                FormatPart::Literal(lit) => result.push_str(lit),
                _ => {}
            }
        }

        if result.is_empty() {
            s.to_string()
        } else {
            result
        }
    }
}

fn fallback_format(value: &Value) -> String {
    match value {
        Value::Number(n) => {
            if n.fract() == 0.0 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        Value::Text(s) => s.to_string(),
        Value::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        Value::Empty => String::new(),
        #[cfg(feature = "chrono")]
        Value::DateTime(dt) => dt.to_string(),
        #[cfg(feature = "chrono")]
        Value::Date(d) => d.to_string(),
        #[cfg(feature = "chrono")]
        Value::Time(t) => t.to_string(),
    }
}
```

In `src/formatter/number.rs`:

```rust
//! Number formatting implementation.

use crate::ast::{Section, FormatPart, DigitPlaceholder};
use crate::options::FormatOptions;

pub fn format_number(value: f64, section: &Section, opts: &FormatOptions) -> String {
    let mut value = value.abs(); // Section selection already handled sign

    // Count percent signs (each multiplies by 100)
    let percent_count = section.parts.iter().filter(|p| matches!(p, FormatPart::Percent)).count();
    for _ in 0..percent_count {
        value *= 100.0;
    }

    // Analyze format structure
    let analysis = analyze_format(section);

    // Round to appropriate decimal places
    let multiplier = 10_f64.powi(analysis.decimal_places as i32);
    value = (value * multiplier).round() / multiplier;

    // Format the number
    let formatted = format_with_placeholders(value, &analysis, opts);

    // Build result with literals
    build_result(section, &formatted, opts)
}

struct FormatAnalysis {
    integer_placeholders: Vec<DigitPlaceholder>,
    decimal_placeholders: Vec<DigitPlaceholder>,
    decimal_places: usize,
    has_thousands_sep: bool,
}

fn analyze_format(section: &Section) -> FormatAnalysis {
    let mut integer_placeholders = Vec::new();
    let mut decimal_placeholders = Vec::new();
    let mut has_thousands_sep = false;
    let mut after_decimal = false;

    for part in &section.parts {
        match part {
            FormatPart::Digit(d) => {
                if after_decimal {
                    decimal_placeholders.push(*d);
                } else {
                    integer_placeholders.push(*d);
                }
            }
            FormatPart::DecimalPoint => {
                after_decimal = true;
            }
            FormatPart::ThousandsSeparator => {
                if !after_decimal {
                    has_thousands_sep = true;
                }
            }
            _ => {}
        }
    }

    FormatAnalysis {
        decimal_places: decimal_placeholders.len(),
        integer_placeholders,
        decimal_placeholders,
        has_thousands_sep,
    }
}

fn format_with_placeholders(value: f64, analysis: &FormatAnalysis, opts: &FormatOptions) -> String {
    let integer_part = value.trunc() as i64;
    let decimal_part = value.fract();

    // Format integer part
    let mut int_str = format_integer(
        integer_part,
        &analysis.integer_placeholders,
        analysis.has_thousands_sep,
        opts,
    );

    // Format decimal part
    if analysis.decimal_places > 0 {
        let dec_str = format_decimal(
            decimal_part,
            &analysis.decimal_placeholders,
            opts,
        );
        int_str.push(opts.locale.decimal_separator);
        int_str.push_str(&dec_str);
    }

    int_str
}

fn format_integer(
    value: i64,
    placeholders: &[DigitPlaceholder],
    thousands_sep: bool,
    opts: &FormatOptions,
) -> String {
    let digits: Vec<char> = value.abs().to_string().chars().collect();
    let min_digits = placeholders.iter().filter(|p| p.is_required()).count();

    let mut result_digits: Vec<char> = Vec::new();

    // Pad with zeros if needed
    let padding = min_digits.saturating_sub(digits.len());
    for _ in 0..padding {
        result_digits.push('0');
    }
    result_digits.extend(digits);

    // Apply thousands separators
    if thousands_sep && result_digits.len() > 3 {
        let mut with_sep = String::new();
        for (i, ch) in result_digits.iter().rev().enumerate() {
            if i > 0 && i % 3 == 0 {
                with_sep.push(opts.locale.thousands_separator);
            }
            with_sep.push(*ch);
        }
        with_sep.chars().rev().collect()
    } else {
        result_digits.into_iter().collect()
    }
}

fn format_decimal(
    value: f64,
    placeholders: &[DigitPlaceholder],
    _opts: &FormatOptions,
) -> String {
    let num_places = placeholders.len();
    let multiplier = 10_f64.powi(num_places as i32);
    let decimal_int = (value * multiplier).round() as i64;

    let mut digits: Vec<char> = format!("{:0>width$}", decimal_int, width = num_places)
        .chars()
        .collect();

    // Trim trailing based on placeholders
    for (i, placeholder) in placeholders.iter().enumerate().rev() {
        if i < digits.len() {
            match placeholder {
                DigitPlaceholder::Hash => {
                    if digits[i] == '0' {
                        digits.pop();
                    } else {
                        break;
                    }
                }
                DigitPlaceholder::Question => {
                    if digits[i] == '0' {
                        digits[i] = ' ';
                    }
                    break;
                }
                DigitPlaceholder::Zero => break,
            }
        }
    }

    digits.into_iter().collect()
}

fn build_result(section: &Section, formatted_number: &str, _opts: &FormatOptions) -> String {
    let mut result = String::new();
    let mut number_placed = false;

    for part in &section.parts {
        match part {
            FormatPart::Digit(_) | FormatPart::DecimalPoint | FormatPart::ThousandsSeparator => {
                if !number_placed {
                    result.push_str(formatted_number);
                    number_placed = true;
                }
            }
            FormatPart::Literal(s) => result.push_str(s),
            FormatPart::Percent => result.push('%'),
            FormatPart::Fill(_) => {} // TODO: implement fill
            FormatPart::Skip(ch) => {
                // Skip adds a space with the width of the character
                result.push(' ');
            }
            _ => {}
        }
    }

    if !number_placed {
        result.push_str(formatted_number);
    }

    result
}
```

Create stub files for other formatter modules:

`src/formatter/date.rs`:
```rust
//! Date formatting implementation.

use crate::ast::Section;
use crate::error::FormatError;
use crate::options::FormatOptions;

pub fn format_date(
    _value: f64,
    _section: &Section,
    _opts: &FormatOptions,
) -> Result<String, FormatError> {
    // TODO: implement
    Ok("TODO".to_string())
}
```

`src/formatter/text.rs`:
```rust
//! Text formatting implementation.
```

`src/formatter/fraction.rs`:
```rust
//! Fraction formatting implementation.
```

**Step 4: Run test to verify it passes**

Run: `cargo test --test format_numbers`
Expected: PASS

**Step 5: Commit**

```bash
git add src/formatter/ tests/format_numbers.rs
git commit -m "feat: implement basic number formatting"
```

---

## Phase 9: Date Formatting

### Task 9.1: Implement Date Formatting

**Files:**
- Modify: `src/formatter/date.rs`
- Create: `tests/format_dates.rs`

**Step 1: Write the failing test**

Create `tests/format_dates.rs`:

```rust
use ssfmt::{NumberFormat, FormatOptions};

#[test]
fn test_format_date_ymd() {
    let fmt = NumberFormat::parse("yyyy-mm-dd").unwrap();
    let opts = FormatOptions::default();

    // January 9, 2026 = serial 46031
    assert_eq!(fmt.format(46031.0, &opts), "2026-01-09");
}

#[test]
fn test_format_date_mdy() {
    let fmt = NumberFormat::parse("m/d/yyyy").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(46031.0, &opts), "1/9/2026");
}

#[test]
fn test_format_time() {
    let fmt = NumberFormat::parse("h:mm:ss").unwrap();
    let opts = FormatOptions::default();

    // 0.5 = 12:00:00 (noon)
    assert_eq!(fmt.format(0.5, &opts), "12:00:00");
}

#[test]
fn test_format_time_ampm() {
    let fmt = NumberFormat::parse("h:mm AM/PM").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(0.5, &opts), "12:00 PM");
    assert_eq!(fmt.format(0.25, &opts), "6:00 AM");
}

#[test]
fn test_format_datetime() {
    let fmt = NumberFormat::parse("yyyy-mm-dd h:mm").unwrap();
    let opts = FormatOptions::default();

    // January 9, 2026 at noon
    assert_eq!(fmt.format(46031.5, &opts), "2026-01-09 12:00");
}

#[test]
fn test_format_month_name() {
    let fmt = NumberFormat::parse("mmmm d, yyyy").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(46031.0, &opts), "January 9, 2026");
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test format_dates`
Expected: FAIL - date formatting returns "TODO"

**Step 3: Implement date formatting**

Update `src/formatter/date.rs`:

```rust
//! Date formatting implementation.

use crate::ast::{Section, FormatPart, DatePart, AmPmStyle};
use crate::date_serial::{serial_to_date, serial_to_time, serial_to_weekday};
use crate::error::FormatError;
use crate::options::FormatOptions;

pub fn format_date(
    value: f64,
    section: &Section,
    opts: &FormatOptions,
) -> Result<String, FormatError> {
    let (year, month, day) = serial_to_date(value, opts.date_system)
        .ok_or(FormatError::DateOutOfRange { serial: value })?;

    let (hour, minute, second) = serial_to_time(value);
    let weekday = serial_to_weekday(value, opts.date_system);

    // Check for AM/PM to determine 12-hour mode
    let has_ampm = section.parts.iter().any(|p| matches!(p, FormatPart::AmPm(_)));
    let (display_hour, is_pm) = if has_ampm {
        let h = if hour == 0 {
            12
        } else if hour > 12 {
            hour - 12
        } else {
            hour
        };
        (h, hour >= 12)
    } else {
        (hour, false)
    };

    let mut result = String::new();

    for part in &section.parts {
        match part {
            FormatPart::DatePart(dp) => {
                let formatted = format_date_part(
                    *dp,
                    year,
                    month,
                    day,
                    display_hour,
                    minute,
                    second,
                    weekday,
                    opts,
                );
                result.push_str(&formatted);
            }
            FormatPart::AmPm(style) => {
                let s = format_ampm(*style, is_pm, opts);
                result.push_str(&s);
            }
            FormatPart::Literal(s) => result.push_str(s),
            _ => {}
        }
    }

    Ok(result)
}

fn format_date_part(
    part: DatePart,
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    weekday: u32,
    opts: &FormatOptions,
) -> String {
    match part {
        DatePart::Year2 => format!("{:02}", year % 100),
        DatePart::Year4 => format!("{:04}", year),

        DatePart::Month => format!("{}", month),
        DatePart::Month2 => format!("{:02}", month),
        DatePart::MonthAbbr => opts.locale.month_names_short[month as usize - 1].to_string(),
        DatePart::MonthFull => opts.locale.month_names_full[month as usize - 1].to_string(),
        DatePart::MonthLetter => opts.locale.month_names_full[month as usize - 1]
            .chars()
            .next()
            .unwrap_or('?')
            .to_string(),

        DatePart::Day => format!("{}", day),
        DatePart::Day2 => format!("{:02}", day),
        DatePart::DayAbbr => opts.locale.day_names_short[weekday as usize].to_string(),
        DatePart::DayFull => opts.locale.day_names_full[weekday as usize].to_string(),

        DatePart::Hour => format!("{}", hour),
        DatePart::Hour2 => format!("{:02}", hour),

        DatePart::Minute => format!("{}", minute),
        DatePart::Minute2 => format!("{:02}", minute),

        DatePart::Second => format!("{}", second),
        DatePart::Second2 => format!("{:02}", second),

        DatePart::SubSecond(digits) => {
            // TODO: implement subseconds
            "0".repeat(digits as usize)
        }
    }
}

fn format_ampm(style: AmPmStyle, is_pm: bool, opts: &FormatOptions) -> String {
    match style {
        AmPmStyle::Upper => {
            if is_pm {
                opts.locale.pm_string.to_string()
            } else {
                opts.locale.am_string.to_string()
            }
        }
        AmPmStyle::Lower => {
            if is_pm {
                opts.locale.pm_string.to_lowercase()
            } else {
                opts.locale.am_string.to_lowercase()
            }
        }
        AmPmStyle::ShortUpper => {
            if is_pm { "P" } else { "A" }.to_string()
        }
        AmPmStyle::ShortLower => {
            if is_pm { "p" } else { "a" }.to_string()
        }
    }
}
```

**Step 4: Run test to verify it passes**

Run: `cargo test --test format_dates`
Expected: PASS

**Step 5: Commit**

```bash
git add src/formatter/date.rs tests/format_dates.rs
git commit -m "feat: implement date/time formatting"
```

---

## Phase 10: Convenience Functions and Caching

### Task 10.1: Implement Convenience Functions

**Files:**
- Modify: `src/lib.rs`
- Modify: `src/cache.rs`
- Create: `tests/convenience_tests.rs`

**Step 1: Write the failing test**

Create `tests/convenience_tests.rs`:

```rust
use ssfmt::{format, format_default};

#[test]
fn test_format_convenience() {
    let opts = ssfmt::FormatOptions::default();
    let result = format(1234.5, "#,##0.00", &opts).unwrap();
    assert_eq!(result, "1,234.50");
}

#[test]
fn test_format_default_convenience() {
    let result = format_default(0.42, "0%").unwrap();
    assert_eq!(result, "42%");
}

#[test]
fn test_format_invalid_code() {
    let opts = ssfmt::FormatOptions::default();
    // Empty format should error
    let result = format(42.0, "", &opts);
    assert!(result.is_err());
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test --test convenience_tests`
Expected: FAIL - functions not found

**Step 3: Implement cache and convenience functions**

In `src/cache.rs`:

```rust
//! Format code caching.

use lru::LruCache;
use std::num::NonZeroUsize;
use std::sync::Mutex;

use crate::ast::NumberFormat;
use crate::error::ParseError;

/// Global cache for parsed format codes.
static CACHE: Mutex<Option<LruCache<String, NumberFormat>>> = Mutex::new(None);

const CACHE_SIZE: usize = 100;

/// Get or parse a format code, using the cache.
pub fn get_or_parse(format_code: &str) -> Result<NumberFormat, ParseError> {
    let mut cache_guard = CACHE.lock().unwrap();

    let cache = cache_guard.get_or_insert_with(|| {
        LruCache::new(NonZeroUsize::new(CACHE_SIZE).unwrap())
    });

    if let Some(fmt) = cache.get(format_code) {
        return Ok(fmt.clone());
    }

    let fmt = NumberFormat::parse(format_code)?;
    cache.put(format_code.to_string(), fmt.clone());
    Ok(fmt)
}
```

In `src/lib.rs`, add at the end:

```rust
// Convenience functions

/// Parse and format a value in one call.
///
/// This function caches recently used format codes for efficiency.
pub fn format<'a>(
    value: impl Into<Value<'a>>,
    format_code: &str,
    opts: &FormatOptions,
) -> Result<String, ParseError> {
    let fmt = cache::get_or_parse(format_code)?;
    Ok(fmt.format(value, opts))
}

/// Format a value with default options (1900 date system, en-US locale).
///
/// This function caches recently used format codes for efficiency.
pub fn format_default<'a>(
    value: impl Into<Value<'a>>,
    format_code: &str,
) -> Result<String, ParseError> {
    let opts = FormatOptions::default();
    format(value, format_code, &opts)
}
```

**Step 4: Run test to verify it passes**

Run: `cargo test --test convenience_tests`
Expected: PASS

**Step 5: Commit**

```bash
git add src/cache.rs src/lib.rs tests/convenience_tests.rs
git commit -m "feat: implement convenience functions with caching"
```

---

## Phase 11: Integration Tests

### Task 11.1: Create Integration Test Suite

**Files:**
- Create: `tests/integration.rs`

**Step 1: Create comprehensive integration tests**

```rust
//! Integration tests covering realistic Excel format codes.

use ssfmt::{NumberFormat, FormatOptions, DateSystem};

// === Number Formats ===

#[test]
fn test_general_number() {
    let fmt = NumberFormat::parse("General").unwrap_or_else(|_| {
        NumberFormat::parse("0").unwrap()
    });
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(123.0, &opts), "123");
}

#[test]
fn test_accounting_format() {
    let fmt = NumberFormat::parse("_($* #,##0.00_)").unwrap();
    let opts = FormatOptions::default();

    let result = fmt.format(1234.56, &opts);
    assert!(result.contains("1,234.56"));
}

#[test]
fn test_negative_in_parens() {
    let fmt = NumberFormat::parse("#,##0;(#,##0)").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(1234.0, &opts), "1,234");
    assert_eq!(fmt.format(-1234.0, &opts), "(1,234)");
}

#[test]
fn test_zero_section() {
    let fmt = NumberFormat::parse("0;-0;\"zero\"").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(0.0, &opts), "zero");
}

// === Date Formats ===

#[test]
fn test_iso_date() {
    let fmt = NumberFormat::parse("yyyy-mm-dd").unwrap();
    let opts = FormatOptions::default();

    // 2026-01-09
    assert_eq!(fmt.format(46031.0, &opts), "2026-01-09");
}

#[test]
fn test_us_date() {
    let fmt = NumberFormat::parse("m/d/yy").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(46031.0, &opts), "1/9/26");
}

#[test]
fn test_long_date() {
    let fmt = NumberFormat::parse("dddd, mmmm d, yyyy").unwrap();
    let opts = FormatOptions::default();

    let result = fmt.format(46031.0, &opts);
    assert!(result.contains("January"));
    assert!(result.contains("2026"));
}

// === Time Formats ===

#[test]
fn test_24h_time() {
    let fmt = NumberFormat::parse("hh:mm:ss").unwrap();
    let opts = FormatOptions::default();

    // 0.75 = 18:00:00
    assert_eq!(fmt.format(0.75, &opts), "18:00:00");
}

#[test]
fn test_12h_time() {
    let fmt = NumberFormat::parse("h:mm AM/PM").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(0.75, &opts), "6:00 PM");
}

// === Date System ===

#[test]
fn test_1904_date_system() {
    let fmt = NumberFormat::parse("yyyy-mm-dd").unwrap();
    let opts = FormatOptions {
        date_system: DateSystem::Date1904,
        ..Default::default()
    };

    // In 1904 system, day 0 = Jan 1, 1904
    let result = fmt.format(0.0, &opts);
    assert!(result.contains("1904"));
}

// === Text Formats ===

#[test]
fn test_text_format() {
    let fmt = NumberFormat::parse("@").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format("Hello", &opts), "Hello");
}

#[test]
fn test_text_with_prefix() {
    let fmt = NumberFormat::parse("\"ID: \"@").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format("12345", &opts), "ID: 12345");
}

// === Colors (parsing only, no ANSI output) ===

#[test]
fn test_color_parsing() {
    let fmt = NumberFormat::parse("[Red]0").unwrap();
    assert!(fmt.has_color());
}

// === Conditions ===

#[test]
fn test_conditional_format() {
    let fmt = NumberFormat::parse("[>=100]\"high\";[<100]\"low\"").unwrap();
    assert!(fmt.has_condition());
}
```

**Step 2: Run tests**

Run: `cargo test --test integration`
Expected: Most should PASS

**Step 3: Commit**

```bash
git add tests/integration.rs
git commit -m "feat: add integration test suite"
```

---

## Phase 12: Documentation and Polish

### Task 12.1: Add Documentation

**Files:**
- Modify: `src/lib.rs`
- Modify: `README.md` (create if needed)

**Step 1: Add crate-level documentation**

Update the doc comment at the top of `src/lib.rs`:

```rust
//! # ssfmt
//!
//! Excel-compatible ECMA-376 number format codes for Rust.
//!
//! This crate provides parsing and formatting of spreadsheet number format codes,
//! matching Excel's actual behavior including undocumented quirks.
//!
//! ## Quick Start
//!
//! ```rust
//! use ssfmt::{format_default, NumberFormat, FormatOptions};
//!
//! // One-off formatting
//! let result = format_default(1234.56, "#,##0.00").unwrap();
//! assert_eq!(result, "1,234.56");
//!
//! // Compile once, format many
//! let fmt = NumberFormat::parse("#,##0.00").unwrap();
//! let opts = FormatOptions::default();
//! assert_eq!(fmt.format(1234.56, &opts), "1,234.56");
//! assert_eq!(fmt.format(9876.54, &opts), "9,876.54");
//! ```
//!
//! ## Format Code Syntax
//!
//! Format codes can have up to 4 sections separated by semicolons:
//! 1. Positive numbers
//! 2. Negative numbers
//! 3. Zero
//! 4. Text
//!
//! ### Number Placeholders
//! - `0` - Display digit or zero
//! - `#` - Display digit or nothing
//! - `?` - Display digit or space
//!
//! ### Date/Time Codes
//! - `yyyy` - Four-digit year
//! - `mm` - Two-digit month
//! - `dd` - Two-digit day
//! - `hh` - Two-digit hour
//! - `mm` - Two-digit minute (after hour)
//! - `ss` - Two-digit second
//!
//! ## Feature Flags
//!
//! - `chrono` (default) - Enable chrono type support
```

**Step 2: Create README.md**

```markdown
# ssfmt

Excel-compatible ECMA-376 number format codes for Rust.

## Features

- Parse and format Excel/OOXML number format codes
- Match Excel's actual behavior, including quirks
- Support for dates, times, percentages, fractions
- Multiple format sections (positive/negative/zero/text)
- Color and conditional format detection
- Both 1900 and 1904 date systems
- Efficient compile-once, format-many pattern

## Usage

```rust
use ssfmt::{format_default, NumberFormat, FormatOptions};

// Simple one-off formatting
let result = format_default(1234.56, "#,##0.00").unwrap();
assert_eq!(result, "1,234.56");

// Compile once, format many values
let fmt = NumberFormat::parse("yyyy-mm-dd").unwrap();
let opts = FormatOptions::default();
assert_eq!(fmt.format(46031.0, &opts), "2026-01-09");
```

## License

MIT OR Apache-2.0
```

**Step 3: Commit**

```bash
git add src/lib.rs README.md
git commit -m "docs: add crate documentation and README"
```

---

### Task 12.2: Final Cleanup and Verify

**Step 1: Run all tests**

```bash
cargo test
```

**Step 2: Run clippy**

```bash
cargo clippy -- -D warnings
```

**Step 3: Check formatting**

```bash
cargo fmt -- --check
```

**Step 4: Fix any issues found**

**Step 5: Final commit**

```bash
git add -A
git commit -m "chore: final cleanup and formatting"
```

---

## Summary

This plan implements ssfmt in 12 phases:

1. **Project Foundation** - Cargo.toml, module structure
2. **Error Types** - ParseError, FormatError
3. **AST Types** - Color, Condition, FormatPart, Section, NumberFormat
4. **Value Types** - Value enum, FormatOptions
5. **Date Serial** - Excel date conversion with leap year bug
6. **Lexer** - Tokenizer for format codes
7. **Parser** - Parse format codes to AST
8. **Number Formatting** - Basic numeric formatting
9. **Date Formatting** - Date/time formatting
10. **Convenience Functions** - Caching and one-liner APIs
11. **Integration Tests** - Comprehensive test suite
12. **Documentation** - Docs and README

Each phase builds on the previous, following TDD with bite-sized commits.
