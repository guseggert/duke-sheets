# ssfmt Architecture

This document describes the architecture of ssfmt and key design decisions based on learnings from the SheetJS SSF reference implementation.

## Overview

ssfmt is a Rust implementation of Excel/ECMA-376 number formatting that achieves 99.9999% compatibility with SheetJS SSF across 19.5+ million test cases.

## Core Design Principles

### 1. Parse Once, Format Many

The library follows a two-phase design:
- **Parse phase**: Format string → AST with pre-computed metadata
- **Format phase**: Use AST + metadata to format values efficiently

This matches SSF's approach and enables:
- Efficient reuse of compiled formats
- No repeated scanning of format parts
- Performance optimization through metadata

### 2. Separate Code Paths by Type

Following SSF's design (bits/66_numint.js vs bits/63_numflt.js vs bits/35_datecode.js):
- **Integer path**: i64 arithmetic for exact integers (< 2^53)
- **Float path**: f64 arithmetic for decimals and large numbers
- **Date/time path**: Serial number conversion and formatting
- **Fraction path**: Continued fractions algorithm

Each path is optimized for its specific use case.

### 3. Metadata-Driven Formatting

Format characteristics are computed once during parsing and stored in `SectionMetadata`:

```rust
pub struct SectionMetadata {
    pub has_ampm: bool,
    pub is_hijri: bool,
    pub max_subsecond_precision: Option<u8>,
    pub has_elapsed_time: bool,
    pub smallest_time_unit: TimeUnit,
    pub format_type: FormatType,
}
```

This eliminates O(n) scans during formatting hot paths.

## Key Architecture Components

### AST (Abstract Syntax Tree)

**Location**: `src/ast.rs`

The AST represents the parsed structure of a format string:

```
NumberFormat
  └─ Section(s) [1-4 sections: positive, negative, zero, text]
       ├─ condition: Option<Condition>
       ├─ color: Option<Color>
       ├─ parts: Vec<FormatPart>
       └─ metadata: SectionMetadata
```

Key types:
- `FormatPart`: Enum of all possible format elements (digit, literal, date part, etc.)
- `SectionMetadata`: Pre-computed characteristics for fast formatting
- `TimeUnit`: Hierarchy for time pre-rounding (None < Hours < Minutes < Seconds < Subseconds)
- `FormatType`: Classification (General, DateTime, Number, Fraction, Text)

### Parser

**Location**: `src/parser/`

Two-stage parsing:

1. **Lexer** (`lexer.rs`): Format string → Token stream
   - Handles escaped characters, quoted literals, color codes
   - Recognizes date/time codes, digit placeholders, special characters

2. **Parser** (`mod.rs`): Token stream → AST with metadata
   - Splits sections by semicolons
   - Parses conditions and colors
   - **Computes metadata during parsing** (critical for performance)
   - Validates structure

### Formatters

**Location**: `src/formatter/`

#### Number Formatter (`number.rs`)

Handles integers, decimals, percentages, scientific notation:

```rust
pub fn format_number(value: f64, section: &Section, opts: &FormatOptions)
```

**Integer Fast Path**:
- Detects exact integers within safe range (< 2^53)
- Only used when NO decimal placeholders present
- Uses i64 arithmetic to avoid precision loss
- Based on SSF's bits/66_numint.js

**Float Path**:
- Handles all decimal formatting
- Supports optional vs required placeholders (# vs 0)
- Handles thousands separators, percent scaling, trailing commas

**Key optimizations**:
- O(n) string building (Vec + reverse instead of insert(0))
- Exact capacity pre-allocation
- Unified placeholder formatting helper

#### Date/Time Formatter (`date.rs`)

Handles date and time formatting:

```rust
pub fn format_date(value: f64, section: &Section, opts: &FormatOptions)
```

**Pre-rounding algorithm** (from SSF bits/82_eval.js):
- Rounds time components based on smallest displayed unit
- Ensures consistent rounding (e.g., 23:59:59.999 → 24:00 if minutes displayed)
- Applied uniformly to all time formats (regular and elapsed)

**Special handling**:
- Day 0 (Dec 31, 1899) = Saturday (off-by-one from typical week calculation)
- Hijri calendar: Subtract 581 years from Gregorian
- Three-digit year format: Minimum 3 digits
- Date overflow: Empty string for values < 0 or > 2958465

#### Fraction Formatter (`fraction.rs`)

Implements continued fractions algorithm for finding best approximation:

```rust
pub fn format_fraction(value: f64, section: &Section, opts: &FormatOptions)
```

**SSF-compliant padding** (from bits/63_numflt.js and bits/30_frac.js):
- Mixed fractions: Space after integer part detection
- Padding width (`ri`): min(max(numerator_len, denominator_len), 7)
- Numerator: Left-pad with spaces
- Denominator: Right-pad with spaces
- Maximum denominator: 10^7 - 1

### Date Serial Conversion

**Location**: `src/date_serial.rs`

Handles Excel's date system (days since 1900-01-01):

**Design choice**: Uses simple O(n) year-by-year loop instead of complex O(1) algorithm
- Reasoning: 200 iterations ≈ 0.00005ms (negligible)
- Benefit: Simple, correct, maintainable code
- Trade-off analysis: Complexity vs performance gain not justified

Supports both 1900 and 1904 date systems.

## Performance Optimizations

### 1. Metadata Pre-computation

**Impact**: Eliminates O(n) scans in formatting hot path

Before:
```rust
let has_ampm = section.parts.iter().any(|p| matches!(p, FormatPart::AmPm(_)));
let is_hijri = section.parts.iter().any(...);
let subsecond_places: Vec<u8> = section.parts.iter().filter_map(...).collect();
```

After:
```rust
// Computed once during parsing, stored in section.metadata
if section.metadata.has_ampm { ... }
if section.metadata.is_hijri { ... }
```

### 2. Integer Fast Path

**Impact**: Avoids precision loss for large integers

Detects safe integers and uses i64 arithmetic:
```rust
const MAX_SAFE_INTEGER: f64 = 9007199254740992.0; // 2^53
if value.fract() == 0.0
    && value.abs() < MAX_SAFE_INTEGER
    && analysis.decimal_placeholders.is_empty()
{
    return format_number_as_integer(value as i64, section, opts);
}
```

**Important**: Only used when no decimal placeholders, as the float path handles optional placeholder logic.

### 3. O(n) String Building

**Impact**: Fixed O(n²) performance bugs

Before (O(n²)):
```rust
for ch in chars {
    result.insert(0, ch);  // Shifts all existing chars
}
```

After (O(n)):
```rust
let mut chars = Vec::with_capacity(estimated_size);
for ch in chars_to_add {
    chars.push(ch);
}
chars.reverse();  // Single O(n) operation
let result: String = chars.into_iter().collect();
```

### 4. Exact Capacity Pre-allocation

**Impact**: Zero reallocations, zero memory waste

```rust
fn build_result(analysis: &FormatAnalysis, number: &str, opts: &FormatOptions) -> String {
    let capacity = count_part_chars(&analysis.prefix_parts)
        + number.len()
        + count_part_chars(&analysis.suffix_parts);
    let mut result = String::with_capacity(capacity);
    // ... build result with no reallocations
}
```

## SSF Reference Implementation

Our implementation closely follows SSF's battle-tested algorithms:

### Key SSF Source Files Referenced

- `bits/82_eval.js`: Pre-rounding algorithm (`bt` variable)
- `bits/66_numint.js`: Integer formatting path
- `bits/63_numflt.js`: Float formatting path
- `bits/59_numhelp.js`: Placeholder formatting helper (`write_num`)
- `bits/35_datecode.js`: Date/time formatting
- `bits/45_hijri.js`: Hijri calendar conversion
- `bits/30_frac.js`: Fraction algorithm and padding

### Divergences from SSF

1. **Date serial conversion**: O(n) loop instead of O(1) algorithm (simpler, correct, fast enough)
2. **Type system**: Rust's strong typing vs JavaScript's dynamic typing
3. **Error handling**: Result types instead of exceptions
4. **Memory management**: Rust ownership vs JavaScript GC

## Testing Strategy

### Test Coverage

- **19,577,033 total test cases** from SheetJS SSF canonical repository
- **99.9999% pass rate** (19,577,017 passing)
- Test fixtures verified against latest SSF source

### Test Categories

1. Built-in format IDs (672 tests)
2. General number formatting (493 tests)
3. Scientific notation (180 tests)
4. Comma/thousands formatting (105 tests)
5. Format string validation (442 tests)
6. Date formatting (3.8M tests)
7. Time formatting (15.7M tests)
8. Fraction formatting (106 tests)
9. Edge cases/oddities (275 tests)

### Known Limitations

16 remaining test failures, all due to:
- IEEE 754 precision limits (5 tests)
- Test data errors in SSF test expectations (2 tests)
- Edge cases in implied format tests (6 tests)
- Out-of-range date tests (84 skipped)

See `docs/TESTING.md` for detailed test coverage documentation.

## Future Considerations

### Performance

Current optimizations are sufficient for typical use cases. Further optimization opportunities:
- SIMD for digit processing (likely unnecessary)
- Parallel formatting of multiple values (user can parallelize at call site)
- Custom allocator for string building (premature)

### Features

Potential future additions:
- Custom format code validation with detailed error messages
- Format code simplification/normalization
- Locale-specific formatting beyond basic decimal/thousands separators
- Excel 365 new format codes (if/when standardized)

### Maintenance

- Keep test fixtures synchronized with SSF canonical source
- Monitor SSF for algorithm updates or bug fixes
- Maintain compatibility with Excel's actual behavior (SSF is the reference)

## Contributing Guidelines

When modifying the codebase:

1. **Preserve SSF compatibility**: Changes should maintain or improve test pass rate
2. **Benchmark performance**: Use criterion for performance-sensitive changes
3. **Document SSF references**: Link to specific SSF source files/line numbers
4. **Test comprehensively**: Run full test suite, not just unit tests
5. **Follow existing patterns**: Metadata-driven, separate code paths by type
6. **Prefer simplicity**: Only optimize when measurably necessary
7. **Maintain documentation**: Update this file and TESTING.md

## References

- [SheetJS SSF Repository](https://github.com/SheetJS/ssf) (legacy)
- [SheetJS SSF Canonical Source](https://git.sheetjs.com/sheetjs/sheetjs/src/branch/master/packages/ssf)
- [ECMA-376 Standard](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [Excel Number Format Specification](https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat)
