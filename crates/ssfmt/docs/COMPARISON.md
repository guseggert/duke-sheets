# Comparison: ssfmt vs numfmt-rs

This document compares `ssfmt` with `numfmt-rs`, another Rust implementation of Excel number format codes. Both crates were created in response to [umya-spreadsheet issue #277](https://github.com/MathNya/umya-spreadsheet/issues/277).

## Overview

| Crate | Author | Approach |
|-------|--------|----------|
| ssfmt | ketbra | Compile format to typed AST, format many values |
| numfmt-rs | hongjr03 | Parse to flat struct, cache internally |

## Excel Compliance

Both crates target SheetJS SSF compatibility. ssfmt has been validated against the full SSF test suite:

| Test Suite | Tests | ssfmt Pass Rate |
|------------|-------|-----------------|
| SSF Dates | 3,846,024 | 100.0% |
| SSF Times | 15,728,625 | 83.6%* |
| SSF General | 493 | 100.0% |
| SSF Fractions | 106 | 100.0% |
| SSF Valid | 442 | 100.0% |
| SSF Oddities | 275 | 89.5% |
| SSF Implied | 672 | 78.7% |

\* Time failures are edge cases with subsecond rounding at precision boundaries.

When run against the same test fixtures, both libraries produce identical results for core formatting operations.

## Performance

Benchmarks were run with 100,000 iterations per test on a release build.

### Scenario 1: One-Shot Formatting (Parse + Format)

This measures the time to parse a format string and format a single value. This is relevant when formats are used only once.

| Format Type | ssfmt (ns/op) | numfmt-rs (ns/op) | Winner |
|-------------|---------------|-------------------|--------|
| General (integer) | 101 | 242 | **ssfmt 2.40x** |
| General (decimal) | 256 | 683 | **ssfmt 2.66x** |
| Number with decimals | 575 | 799 | **ssfmt 1.39x** |
| Percentage | 523 | 775 | **ssfmt 1.48x** |
| Scientific notation | 385 | 830 | **ssfmt 2.16x** |
| Fraction | 355 | 550 | **ssfmt 1.55x** |
| Date format | 394 | 385 | ~tied |
| Time format | 464 | 385 | numfmt-rs 1.21x |
| Complex date/time | 675 | 524 | numfmt-rs 1.29x |
| Elapsed time | 407 | 382 | numfmt-rs 1.06x |

### Scenario 2: Pre-Compiled AST (ssfmt) vs Internal Cache (numfmt-rs)

This measures the time when the format has already been parsed. For ssfmt, this uses `NumberFormat::parse()` once and then `format()` repeatedly. For numfmt-rs, this relies on its internal `HashMap` cache.

This is the typical spreadsheet use case where the same format is applied to thousands of cells.

| Format Type | ssfmt (ns/op) | numfmt-rs (ns/op) | Winner |
|-------------|---------------|-------------------|--------|
| General (integer) | 33 | 246 | **ssfmt 7.5x** |
| Number with decimals | 407 | 798 | **ssfmt 1.96x** |
| Date format | 243 | 387 | **ssfmt 1.59x** |
| Time format | 314 | 387 | **ssfmt 1.23x** |

### Analysis

- **ssfmt is faster for number formatting** in both scenarios (1.4-2.7x one-shot, up to 7.5x pre-compiled)
- **For date/time formatting**: numfmt-rs slightly faster in one-shot scenario (1.1-1.3x due to winnow parser), but ssfmt faster when pre-compiled (1.2-1.6x)
- **Pre-compilation matters**: ssfmt's explicit AST approach shows its advantage in the cached scenario
- **General integer fast path**: ssfmt uses an integer fast path that avoids floating-point operations for whole numbers
- **Parser optimization**: ssfmt avoids allocations during lexing with case-insensitive comparisons

### Performance Note

ssfmt uses O(1) algorithms for date serial number conversion (Fliegel & Van Flandern, 1968), avoiding the O(years) iteration that would be required for modern dates. This provides consistent performance regardless of the date being formatted.

### Caching Strategies

| Aspect | ssfmt | numfmt-rs |
|--------|-------|-----------|
| Cache type | LRU (bounded, 100 entries) | HashMap (unbounded) |
| Pre-compile API | `NumberFormat::parse()` | Not exposed |
| Memory safety | Bounded, won't grow | Can grow indefinitely |
| Zero-copy reuse | Yes, via `&NumberFormat` | Cache lookup each call |

## Code Idiomaticity

### ssfmt

**Strongly-typed AST:**
```rust
pub enum FormatPart {
    Literal(String),
    Digit(DigitPlaceholder),
    DecimalPoint,
    ThousandsSeparator,
    DatePart(DatePart),
    Scientific { upper: bool, show_plus: bool },
    Fraction {
        integer_digits: Vec<DigitPlaceholder>,
        numerator_digits: Vec<DigitPlaceholder>,
        denominator: FractionDenom,
        ...
    },
    ...
}
```

**Pre-computed metadata:**
```rust
pub struct SectionMetadata {
    pub has_ampm: bool,
    pub is_hijri: bool,
    pub max_subsecond_precision: Option<u8>,
    pub smallest_time_unit: TimeUnit,
    pub format_type: FormatType,
}
```

**Structured errors with thiserror:**
```rust
#[derive(Debug, Clone, PartialEq, Error)]
pub enum ParseError {
    #[error("unexpected token at position {position}: found '{found}'")]
    UnexpectedToken { position: usize, found: char },
    #[error("unterminated bracket at position {position}")]
    UnterminatedBracket { position: usize },
    ...
}
```

### numfmt-rs

**Flat struct with many fields:**
```rust
pub struct Section {
    pub scale: f64,
    pub percent: bool,
    pub text: bool,
    pub date: DateUnits,
    pub int_pattern: Vec<String>,  // String-based patterns
    pub frac_pattern: Vec<String>,
    pub man_pattern: Vec<String>,
    pub den_pattern: Vec<String>,
    pub num_pattern: Vec<String>,
    pub grouping: bool,
    pub fractions: bool,
    // ... 20+ more fields
}
```

**String-based errors:**
```rust
pub struct ParseError {
    message: String,  // Less structured
}
```

**Builder pattern for options:**
```rust
FormatterOptions::default()
    .with_locale("de")
    .with_nbsp(true)
```

## Feature Comparison

| Feature | ssfmt | numfmt-rs |
|---------|-------|-----------|
| Parse once, format many | Yes (explicit API) | Yes (internal cache) |
| Bounded cache | Yes (LRU, 100) | No (HashMap) |
| Typed AST | Yes | No (flat struct) |
| Structured errors | Yes | No |
| Locale support | Basic (en-US) | Extensive (many locales) |
| BigInt support | No | Yes |
| WASM ready | Possible | Built-in |
| Chrono integration | Yes (optional) | No (uses serial numbers) |
| Dependencies | Lighter | Heavier |

## Dependencies

### ssfmt
- `lru` - LRU cache
- `thiserror` - Error derive macro
- `chrono` (optional) - Date/time types

### numfmt-rs
- `winnow` - Parser combinator library
- `bitflags` - Bit flag macros
- `num-bigint` - Arbitrary precision integers
- `num-traits` - Numeric traits
- `wasm-bindgen` - WASM bindings
- `serde` + `serde_json` - Serialization

## Recommendations

### For umya-spreadsheet
**ssfmt is recommended** because:
1. Faster number formatting (most common operation in spreadsheets)
2. Pre-compiled AST matches spreadsheet use case (same format, many cells)
3. Bounded LRU cache prevents memory issues
4. Lighter dependency footprint
5. More maintainable typed code

### When to choose numfmt-rs
- Need BigInt support for very large numbers
- Need extensive multi-locale formatting
- Building a WASM application
- Need the builder pattern for options

## Conclusion

Both crates are solid implementations of Excel number formatting. ssfmt follows a more traditional compiler design (lexer → parser → AST → formatter) with idiomatic Rust patterns, while numfmt-rs is a more direct port of the JavaScript numfmt library with Rust-specific additions.

For typical spreadsheet processing, ssfmt's pre-compiled AST approach provides better performance and memory characteristics, while numfmt-rs offers more features for specialized use cases.
