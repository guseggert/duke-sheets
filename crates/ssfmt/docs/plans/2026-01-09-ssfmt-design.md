# ssfmt Design Document

A Rust crate for Excel-compatible ECMA-376 number format codes.

## Overview

`ssfmt` provides parsing and formatting of spreadsheet number format codes, matching Excel's actual behavior including undocumented quirks. It's designed for use with spreadsheet libraries like umya-spreadsheet and calamine.

## Architecture

```
Format Code String → Parser → AST → Formatter → Output String
                      ↓
               (reusable for many values)
```

### Components

1. **Parser** - Tokenizes and parses ECMA-376 format codes into an AST. Handles Excel quirks like implicit sections and legacy codes.

2. **AST** - Public data structures representing parsed format codes.

3. **Formatter** - Walks the AST and produces output strings. Uses ICU4X (when enabled) for locale-aware formatting.

4. **Value Types** - An enum accepting multiple input types (f64, chrono types, strings).

## AST Structure

```rust
pub struct NumberFormat {
    sections: Vec<Section>,  // 1-4 sections: positive, negative, zero, text
}

pub struct Section {
    condition: Option<Condition>,  // e.g., [>100], [=0]
    color: Option<Color>,          // e.g., [Red], [Color3]
    parts: Vec<FormatPart>,
}

pub enum Condition {
    GreaterThan(f64),
    LessThan(f64),
    Equal(f64),
    GreaterOrEqual(f64),
    LessOrEqual(f64),
    NotEqual(f64),
}

pub enum Color {
    Named(NamedColor),  // Red, Blue, Green, etc.
    Indexed(u8),        // [Color1] through [Color56]
}

pub enum FormatPart {
    Literal(String),           // Literal text, escaped chars
    Digit(DigitPlaceholder),   // 0, #, ?
    DecimalPoint,
    ThousandsSeparator,
    Percent,
    Scientific { digits: u8 }, // E+, E-, e+, e-
    Fraction { denom: FractionDenom },
    DatePart(DatePart),        // y, m, d, h, s, etc.
    AmPm(AmPmStyle),
    Elapsed(ElapsedPart),      // [h], [m], [s]
    TextPlaceholder,           // @ for text
    Fill(char),                // *x repeat fill
    Skip(u8),                  // _x skip width
    Locale(LocaleCode),        // [$-409], [$€-407]
}
```

### Inspection Methods

```rust
impl NumberFormat {
    pub fn sections(&self) -> &[Section];
    pub fn is_date_format(&self) -> bool;
    pub fn is_text_format(&self) -> bool;
    pub fn is_percentage(&self) -> bool;
    pub fn has_color(&self) -> bool;
    pub fn has_condition(&self) -> bool;
}
```

## Value Types

```rust
pub enum Value<'a> {
    Number(f64),
    DateTime(chrono::NaiveDateTime),
    Date(chrono::NaiveDate),
    Time(chrono::NaiveTime),
    Text(&'a str),
    Bool(bool),
    Empty,
}
```

With `From` implementations for convenient conversions.

## Format Options

```rust
pub struct FormatOptions {
    pub date_system: DateSystem,
    pub locale: Locale,
}

pub enum DateSystem {
    Date1900,      // Windows Excel default (includes leap year bug)
    Date1904,      // Mac Excel legacy
}
```

### Locale

Without `icu` feature - minimal built-in struct:

```rust
pub struct Locale {
    pub decimal_separator: char,
    pub thousands_separator: char,
    pub currency_symbol: &'static str,
    pub am_string: &'static str,
    pub pm_string: &'static str,
    // ... date/month names
}
```

With `icu` feature - integrates with ICU4X locale data.

## Public API

### Core API

```rust
impl NumberFormat {
    /// Parse a format code (can fail)
    pub fn parse(format_code: &str) -> Result<Self, ParseError>;

    /// Format a value (infallible, falls back gracefully)
    pub fn format(&self, value: impl Into<Value<'_>>, opts: &FormatOptions) -> String;

    /// Format a value (fallible, returns error on type mismatch)
    pub fn try_format(&self, value: impl Into<Value<'_>>, opts: &FormatOptions)
        -> Result<String, FormatError>;
}
```

### Convenience Functions

```rust
/// Parse and format in one call
pub fn format(value: impl Into<Value<'_>>, format_code: &str, opts: &FormatOptions)
    -> Result<String, ParseError>;

/// Format with default options (1900 date system, en-US locale)
pub fn format_default(value: impl Into<Value<'_>>, format_code: &str)
    -> Result<String, ParseError>;
```

## Error Types

```rust
pub enum ParseError {
    UnexpectedToken { position: usize, found: char },
    UnterminatedBracket { position: usize },
    InvalidCondition { position: usize, reason: String },
    InvalidLocaleCode { position: usize },
    TooManySections,
}

pub enum FormatError {
    TypeMismatch { expected: &'static str, got: &'static str },
    DateOutOfRange { serial: f64 },
    InvalidSerialNumber { value: f64 },
}
```

## Module Structure

```
src/
├── lib.rs              # Public API re-exports
├── ast.rs              # Public AST types
├── parser/
│   ├── mod.rs          # Parser entry point
│   ├── lexer.rs        # Tokenizer
│   └── tokens.rs       # Token types
├── formatter/
│   ├── mod.rs          # Formatter entry point
│   ├── number.rs       # Numeric formatting
│   ├── date.rs         # Date/time formatting
│   ├── text.rs         # Text placeholder handling
│   └── fraction.rs     # Fraction formatting
├── locale/
│   ├── mod.rs          # Locale trait and types
│   ├── builtin.rs      # Built-in minimal locales
│   └── icu.rs          # ICU4X integration
├── value.rs            # Value enum and conversions
├── options.rs          # FormatOptions, DateSystem
├── error.rs            # ParseError, FormatError
├── date_serial.rs      # Excel serial ↔ date conversion
└── cache.rs            # LRU cache for convenience functions
```

## Dependencies

```toml
[dependencies]
chrono = { version = "0.4", optional = true, default-features = false }
icu = { version = "1.4", optional = true }
lru = "0.12"

[features]
default = ["chrono"]
icu = ["dep:icu"]
chrono = ["dep:chrono"]
```

## Testing Strategy

### Test Sources

1. **numfmt.js test suite** - Port test cases as golden tests in `tests/fixtures/numfmt_cases.json`

2. **Excel verification workbook** - `tests/fixtures/excel_reference.xlsx` with tricky format codes; extract expected outputs to `tests/fixtures/excel_expected.json`

3. **ECMA-376 spec examples** - Cover documented behaviors

### Test Structure

```
tests/
├── fixtures/
│   ├── numfmt_cases.json
│   ├── excel_reference.xlsx
│   └── excel_expected.json
├── parser_tests.rs
├── format_numbers.rs
├── format_dates.rs
├── format_text.rs
├── locale_tests.rs
└── edge_cases.rs
```

### Key Edge Cases

- Feb 29, 1900 (leap year bug)
- Serial number boundaries: 0, 1, 60, 61
- Elapsed time > 24 hours (`[h]:mm:ss`)
- Fractions with varying denominators
- Locale codes: `[$-409]`, `[$€-407]`, `[$-F800]`
- Conditional sections with colors
- Scientific notation

## Scope

`ssfmt` handles **number format codes only**. The following are out of scope (handled by spreadsheet libraries):

- Borders
- Cell fills/backgrounds
- Fonts
- Cell alignment
- Overall cell styling composition

## Future Considerations

- `no_std` support (code structured to make this feasible)
- Additional locale data bundles
- Format code validation/linting utilities
