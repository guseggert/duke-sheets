# Duke Sheets - Architecture Plan

A Rust library for spreadsheet manipulation, similar to Aspose Cells.

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Requirements](#2-requirements)
3. [Project Structure](#3-project-structure)
4. [Core Data Structures](#4-core-data-structures)
5. [Formula Engine](#5-formula-engine)
6. [File I/O](#6-file-io)
7. [Chart Support](#7-chart-support)
8. [C FFI Layer](#8-c-ffi-layer)
9. [Testing Strategy](#9-testing-strategy)
10. [Dependencies](#10-dependencies)
11. [API Design](#11-api-design)
12. [Implementation Phases](#12-implementation-phases)
13. [Success Criteria](#13-success-criteria)

---

## 1. Executive Summary

**duke-sheets** is a Rust library for spreadsheet manipulation with an API similar to Aspose Cells.

### Key Features

| Feature | Description |
|---------|-------------|
| **File Formats** | XLSX, XLS (legacy BIFF8), CSV |
| **Formula Engine** | Full evaluation of ~450 Excel functions |
| **Charts** | Create, read, modify all chart types |
| **Large Files** | Streaming APIs for >1M cells |
| **Dual API** | Idiomatic Rust + C FFI (handle-based) |

### Design Principles

1. **End-to-end testing** as the primary quality assurance strategy
2. **Excel compatibility** as the ground truth for correctness
3. **Memory efficiency** for large file handling
4. **Safe C FFI** via handle-based design

---

## 2. Requirements

### Functional Requirements

- Read and write XLSX files (Office Open XML)
- Read and write XLS files (BIFF8 legacy format)
- Read and write CSV files
- Full formula evaluation with dependency tracking
- Support all Excel chart types
- Cell styling (fonts, fills, borders, number formats)
- Named ranges
- Conditional formatting
- Data validation
- Comments and annotations

### Non-Functional Requirements

- Handle files with >1,000,000 cells via streaming
- Memory usage <500MB for 1M cell files in streaming mode
- Read 1M cells in <10 seconds
- Write 1M cells in <15 seconds
- No crashes on malformed input (fuzz tested)
- Thread-safe read operations
- C FFI for language bindings

### API Requirements

- **Rust API**: Idiomatic with `Result`, iterators, `&mut` borrowing
- **C API**: Handle-based, error codes, no memory management burden on caller

---

## 3. Project Structure

```
duke-sheets/
├── Cargo.toml                    # Workspace root
├── ARCHITECTURE.md               # This document
├── README.md
├── LICENSE
│
├── crates/
│   ├── duke-sheets-core/         # Core data structures
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── workbook.rs       # Workbook, WorkbookSettings
│   │       ├── worksheet.rs      # Worksheet
│   │       ├── cell/
│   │       │   ├── mod.rs
│   │       │   ├── value.rs      # CellValue enum
│   │       │   ├── storage.rs    # Sparse cell storage
│   │       │   └── address.rs    # CellAddress, CellRange, parsing
│   │       ├── style/
│   │       │   ├── mod.rs
│   │       │   ├── font.rs
│   │       │   ├── border.rs
│   │       │   ├── fill.rs
│   │       │   ├── alignment.rs
│   │       │   ├── number_format.rs
│   │       │   └── pool.rs       # Style deduplication
│   │       ├── row.rs
│   │       ├── column.rs
│   │       ├── range.rs
│   │       └── error.rs
│   │
│   ├── duke-sheets-formula/      # Formula engine
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── ast.rs            # Formula AST types
│   │       ├── parser.rs         # Formula text → AST
│   │       ├── evaluator.rs      # AST → value
│   │       ├── dependency.rs     # Dependency graph
│   │       ├── error.rs          # #VALUE!, #REF!, etc.
│   │       └── functions/
│   │           ├── mod.rs        # Function registry
│   │           ├── math.rs       # SUM, AVERAGE, ROUND, etc.
│   │           ├── text.rs       # CONCATENATE, LEFT, RIGHT, etc.
│   │           ├── logical.rs    # IF, AND, OR, NOT, etc.
│   │           ├── lookup.rs     # VLOOKUP, INDEX, MATCH, XLOOKUP
│   │           ├── date.rs       # DATE, NOW, TODAY, etc.
│   │           ├── statistical.rs # STDEV, VAR, PERCENTILE, etc.
│   │           ├── financial.rs  # NPV, IRR, PMT, etc.
│   │           ├── info.rs       # ISERROR, ISBLANK, TYPE, etc.
│   │           ├── engineering.rs
│   │           └── database.rs   # DSUM, DAVERAGE, etc.
│   │
│   ├── duke-sheets-xlsx/         # XLSX format support
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── reader/
│   │       │   ├── mod.rs
│   │       │   ├── workbook.rs
│   │       │   ├── worksheet.rs
│   │       │   ├── shared_strings.rs
│   │       │   ├── styles.rs
│   │       │   ├── relationships.rs
│   │       │   └── streaming.rs  # SAX-style for large files
│   │       └── writer/
│   │           ├── mod.rs
│   │           ├── workbook.rs
│   │           ├── worksheet.rs
│   │           ├── shared_strings.rs
│   │           ├── styles.rs
│   │           └── streaming.rs  # Streaming writer
│   │
│   ├── duke-sheets-xls/          # XLS (BIFF8) format support
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── biff/
│   │       │   ├── mod.rs
│   │       │   ├── records.rs    # BIFF8 record types
│   │       │   ├── formulas.rs   # BIFF formula tokens
│   │       │   └── encryption.rs # XLS encryption
│   │       ├── cfb.rs            # Compound File Binary handling
│   │       ├── reader.rs
│   │       └── writer.rs
│   │
│   ├── duke-sheets-csv/          # CSV format support
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── reader.rs
│   │       ├── writer.rs
│   │       └── options.rs        # Delimiter, encoding, etc.
│   │
│   ├── duke-sheets-chart/        # Chart support
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── chart.rs          # Chart struct
│   │       ├── series.rs         # Data series
│   │       ├── axis.rs           # Axis configuration
│   │       ├── legend.rs
│   │       ├── title.rs
│   │       ├── data_labels.rs
│   │       └── types/
│   │           ├── mod.rs
│   │           ├── bar.rs
│   │           ├── column.rs
│   │           ├── line.rs
│   │           ├── pie.rs
│   │           ├── scatter.rs
│   │           ├── area.rs
│   │           ├── radar.rs
│   │           ├── stock.rs
│   │           ├── surface.rs
│   │           └── combo.rs
│   │
│   ├── duke-sheets/              # Main public API crate
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs            # Re-exports all public API
│   │       └── prelude.rs        # Common imports
│   │
│   └── duke-sheets-ffi/          # C FFI bindings
│       ├── Cargo.toml
│       ├── src/
│       │   ├── lib.rs
│       │   ├── handles.rs        # Handle management system
│       │   ├── error.rs          # Error codes
│       │   ├── workbook.rs       # Workbook FFI functions
│       │   ├── worksheet.rs      # Worksheet FFI functions
│       │   ├── cell.rs           # Cell FFI functions
│       │   ├── style.rs          # Style FFI functions
│       │   └── chart.rs          # Chart FFI functions
│       ├── include/
│       │   └── duke_sheets.h     # C header file
│       └── cbindgen.toml         # Header generation config
│
├── tests/                        # End-to-end tests (PRIMARY)
│   ├── e2e/
│   │   ├── mod.rs
│   │   ├── roundtrip.rs          # Create → Save → Read → Verify
│   │   ├── excel_compat.rs       # Test against Excel-created files
│   │   ├── formulas.rs           # Formula calculation E2E
│   │   ├── large_files.rs        # Streaming tests
│   │   ├── charts.rs             # Chart operations
│   │   ├── styles.rs             # Style preservation
│   │   ├── xls_format.rs         # Legacy format tests
│   │   ├── csv_format.rs         # CSV tests
│   │   └── ffi.rs                # C API integration tests
│   │
│   └── fixtures/                 # Test files
│       ├── excel_created/        # Files created in actual Excel
│       │   ├── simple.xlsx
│       │   ├── formulas.xlsx
│       │   ├── styles.xlsx
│       │   ├── charts.xlsx
│       │   └── large.xlsx
│       ├── libreoffice_created/  # Files from LibreOffice
│       ├── legacy/               # XLS format files
│       │   └── *.xls
│       ├── csv/
│       │   └── *.csv
│       ├── edge_cases/           # Unusual but valid files
│       └── expected_outputs/     # Expected results for verification
│
├── benches/                      # Performance benchmarks
│   ├── Cargo.toml
│   └── src/
│       ├── read_xlsx.rs
│       ├── write_xlsx.rs
│       ├── formula_calc.rs
│       └── large_files.rs
│
├── fuzz/                         # Fuzzing targets
│   ├── Cargo.toml
│   └── fuzz_targets/
│       ├── xlsx_reader.rs
│       ├── xls_reader.rs
│       ├── csv_reader.rs
│       └── formula_parser.rs
│
├── examples/
│   ├── rust/
│   │   ├── basic_usage.rs
│   │   ├── create_report.rs
│   │   ├── read_data.rs
│   │   ├── formulas.rs
│   │   ├── charts.rs
│   │   ├── large_file.rs
│   │   └── streaming.rs
│   └── c/
│       ├── Makefile
│       ├── basic_usage.c
│       └── read_data.c
│
└── xtask/                        # Build automation
    ├── Cargo.toml
    └── src/
        └── main.rs               # Test fixture generation, etc.
```

---

## 4. Core Data Structures

### 4.1 Cell Value

```rust
/// Represents a cell's value
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell
    Empty,
    
    /// Boolean value (TRUE/FALSE)
    Boolean(bool),
    
    /// Numeric value (all numbers stored as f64)
    Number(f64),
    
    /// String value (interned for memory efficiency)
    String(SharedString),
    
    /// Error value (#VALUE!, #REF!, etc.)
    Error(CellError),
    
    /// Formula with cached result
    Formula {
        /// Original formula text (e.g., "=SUM(A1:A10)")
        text: String,
        /// Parsed AST (lazy, computed on first calculation)
        ast: Option<Box<FormulaExpr>>,
        /// Last calculated value
        cached_value: Box<CellValue>,
        /// Whether recalculation is needed
        needs_recalc: bool,
    },
}

/// Excel error values
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CellError {
    Null,        // #NULL!  - Incorrect range operator
    Div0,        // #DIV/0! - Division by zero
    Value,       // #VALUE! - Wrong type of argument
    Ref,         // #REF!   - Invalid reference
    Name,        // #NAME?  - Unrecognized formula name
    Num,         // #NUM!   - Invalid numeric value
    Na,          // #N/A    - Value not available
    GettingData, // #GETTING_DATA - External data loading
    Spill,       // #SPILL! - Dynamic array can't spill
    Calc,        // #CALC!  - Calculation error
}
```

### 4.2 Cell Storage

```rust
use std::collections::BTreeMap;

/// Sparse row-based cell storage
/// 
/// Design decisions:
/// - BTreeMap for ordered iteration (required for streaming writes)
/// - Row-major layout matches Excel's internal structure
/// - Only non-empty cells are stored
/// - Can handle millions of cells with reasonable memory
pub struct CellStorage {
    /// Row index → Row data
    rows: BTreeMap<u32, RowData>,
    
    /// Shared string pool for deduplication
    string_pool: StringPool,
    
    /// Shared style pool for deduplication
    style_pool: StylePool,
    
    /// Default row height in points
    default_row_height: f64,
    
    /// Default column width in characters
    default_column_width: f64,
    
    /// Column definitions (custom widths, styles, hidden state)
    columns: BTreeMap<u16, ColumnData>,
    
    /// Merged cell regions
    merged_regions: Vec<CellRange>,
    
    /// Storage mode (in-memory vs streaming)
    mode: StorageMode,
}

/// Data for a single row
pub struct RowData {
    /// Column index → Cell data
    cells: BTreeMap<u16, CellData>,
    
    /// Custom row height (None = default)
    height: Option<f64>,
    
    /// Row-level style index
    style_index: Option<u32>,
    
    /// Whether row is hidden
    hidden: bool,
    
    /// Outline/grouping level (0-7)
    outline_level: u8,
}

/// Data for a single cell
#[derive(Clone)]
pub struct CellData {
    /// Cell value
    pub value: CellValue,
    
    /// Index into style pool (0 = default style)
    pub style_index: u32,
}

/// Column metadata
pub struct ColumnData {
    /// Custom width (None = default)
    width: Option<f64>,
    
    /// Column-level style index
    style_index: Option<u32>,
    
    /// Whether column is hidden
    hidden: bool,
    
    /// Outline/grouping level
    outline_level: u8,
}

/// Storage modes for different use cases
pub enum StorageMode {
    /// Standard in-memory storage
    InMemory,
    
    /// Streaming mode for very large files
    /// Writes rows to temp file, only keeps recent rows in memory
    Streaming {
        temp_file: PathBuf,
        buffer_rows: usize,
        current_row: u32,
    },
}
```

### 4.3 String Pool

```rust
use std::collections::HashMap;
use std::sync::Arc;

/// Interned string for memory efficiency
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct SharedString(Arc<str>);

/// String pool for deduplication
/// 
/// Strings are often repeated across cells (e.g., "Yes", "No", dates).
/// Interning reduces memory usage significantly for large files.
pub struct StringPool {
    /// String value → Shared reference
    strings: HashMap<Arc<str>, SharedString>,
}

impl StringPool {
    /// Get or create a shared string
    pub fn intern(&mut self, s: &str) -> SharedString {
        if let Some(shared) = self.strings.get(s) {
            shared.clone()
        } else {
            let arc: Arc<str> = s.into();
            let shared = SharedString(arc.clone());
            self.strings.insert(arc, shared.clone());
            shared
        }
    }
}
```

### 4.4 Style System

```rust
/// Complete cell style
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Style {
    /// Font settings
    pub font: FontStyle,
    
    /// Fill/background settings
    pub fill: FillStyle,
    
    /// Border settings
    pub border: BorderStyle,
    
    /// Text alignment
    pub alignment: Alignment,
    
    /// Number format (e.g., "#,##0.00", "yyyy-mm-dd")
    pub number_format: NumberFormat,
    
    /// Cell protection settings
    pub protection: Protection,
}

/// Font settings
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct FontStyle {
    pub name: String,
    pub size: u16,          // In half-points (e.g., 22 = 11pt)
    pub bold: bool,
    pub italic: bool,
    pub underline: Underline,
    pub strikethrough: bool,
    pub color: Color,
}

/// Fill/background settings
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum FillStyle {
    None,
    Solid(Color),
    Pattern {
        pattern: PatternType,
        foreground: Color,
        background: Color,
    },
    Gradient {
        gradient_type: GradientType,
        angle: f64,
        stops: Vec<GradientStop>,
    },
}

/// Border settings
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct BorderStyle {
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
    pub diagonal: Option<BorderEdge>,
    pub diagonal_direction: DiagonalDirection,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct BorderEdge {
    pub style: BorderLineStyle,
    pub color: Color,
}

/// Text alignment
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Alignment {
    pub horizontal: HorizontalAlignment,
    pub vertical: VerticalAlignment,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
    pub indent: u8,
    pub rotation: i16,  // -90 to 90 degrees
    pub reading_order: ReadingOrder,
}

/// Color representation
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Color {
    /// RGB color
    Rgb { r: u8, g: u8, b: u8 },
    
    /// ARGB color with alpha
    Argb { a: u8, r: u8, g: u8, b: u8 },
    
    /// Theme color with optional tint
    Theme { theme_index: u8, tint: i8 },
    
    /// Indexed color (legacy)
    Indexed(u8),
    
    /// Automatic/default color
    Auto,
}

/// Style pool for deduplication
/// 
/// Excel files typically have many cells sharing the same style.
/// Deduplication reduces memory and file size.
pub struct StylePool {
    /// All unique styles
    styles: Vec<Style>,
    
    /// Fast lookup for deduplication
    index_map: HashMap<Style, u32>,
}

impl StylePool {
    /// Get or create a style, returning its index
    pub fn get_or_insert(&mut self, style: Style) -> u32 {
        if let Some(&idx) = self.index_map.get(&style) {
            idx
        } else {
            let idx = self.styles.len() as u32;
            self.index_map.insert(style.clone(), idx);
            self.styles.push(style);
            idx
        }
    }
    
    /// Get style by index
    pub fn get(&self, index: u32) -> Option<&Style> {
        self.styles.get(index as usize)
    }
}
```

### 4.5 Cell Address

```rust
/// A cell address (e.g., "A1", "$B$2")
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellAddress {
    pub row: u32,
    pub col: u16,
    pub row_absolute: bool,
    pub col_absolute: bool,
}

impl CellAddress {
    /// Parse from A1-style notation
    pub fn parse(s: &str) -> Result<Self, AddressError> {
        // Implementation handles: A1, $A1, A$1, $A$1
    }
    
    /// Format to A1-style notation
    pub fn to_string(&self) -> String {
        let col_str = Self::column_to_letters(self.col);
        format!(
            "{}{}{}{}",
            if self.col_absolute { "$" } else { "" },
            col_str,
            if self.row_absolute { "$" } else { "" },
            self.row + 1  // Excel is 1-indexed
        )
    }
    
    /// Convert column index to letters (0 = A, 25 = Z, 26 = AA)
    pub fn column_to_letters(col: u16) -> String {
        // Implementation
    }
    
    /// Convert letters to column index
    pub fn letters_to_column(letters: &str) -> Result<u16, AddressError> {
        // Implementation
    }
}

/// A range of cells (e.g., "A1:B10")
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRange {
    pub start: CellAddress,
    pub end: CellAddress,
}

impl CellRange {
    /// Parse from A1:B10-style notation
    pub fn parse(s: &str) -> Result<Self, AddressError> {
        // Implementation
    }
    
    /// Check if a cell is within this range
    pub fn contains(&self, addr: &CellAddress) -> bool {
        addr.row >= self.start.row && addr.row <= self.end.row &&
        addr.col >= self.start.col && addr.col <= self.end.col
    }
    
    /// Iterate over all cells in range
    pub fn cells(&self) -> impl Iterator<Item = CellAddress> {
        // Implementation
    }
    
    /// Number of rows in range
    pub fn row_count(&self) -> u32 {
        self.end.row - self.start.row + 1
    }
    
    /// Number of columns in range
    pub fn col_count(&self) -> u16 {
        self.end.col - self.start.col + 1
    }
}
```

---

## 5. Formula Engine

### 5.1 Abstract Syntax Tree

```rust
/// Formula expression AST
#[derive(Debug, Clone)]
pub enum FormulaExpr {
    // === Literals ===
    
    /// Numeric literal (e.g., 42, 3.14)
    Number(f64),
    
    /// String literal (e.g., "hello")
    String(String),
    
    /// Boolean literal (TRUE, FALSE)
    Boolean(bool),
    
    /// Error literal (#N/A, #VALUE!, etc.)
    Error(CellError),
    
    // === References ===
    
    /// Single cell reference (e.g., A1, $B$2)
    CellRef(CellReference),
    
    /// Range reference (e.g., A1:B10)
    RangeRef(RangeReference),
    
    /// Named range or defined name
    NameRef(String),
    
    /// External workbook reference (e.g., [Book1.xlsx]Sheet1!A1)
    ExternalRef {
        workbook: String,
        sheet: String,
        reference: Box<FormulaExpr>,
    },
    
    // === Operators ===
    
    /// Binary operation
    BinaryOp {
        op: BinaryOperator,
        left: Box<FormulaExpr>,
        right: Box<FormulaExpr>,
    },
    
    /// Unary operation
    UnaryOp {
        op: UnaryOperator,
        operand: Box<FormulaExpr>,
    },
    
    // === Function call ===
    
    Function {
        name: String,
        args: Vec<FormulaExpr>,
    },
    
    // === Array ===
    
    /// Array constant (e.g., {1,2,3;4,5,6})
    Array(Vec<Vec<FormulaExpr>>),
}

/// Cell reference with optional sheet
#[derive(Debug, Clone)]
pub struct CellReference {
    pub sheet: Option<String>,
    pub address: CellAddress,
}

/// Range reference with optional sheet
#[derive(Debug, Clone)]
pub struct RangeReference {
    pub sheet: Option<String>,
    pub range: CellRange,
}

/// Binary operators
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOperator {
    // Arithmetic
    Add,        // +
    Subtract,   // -
    Multiply,   // *
    Divide,     // /
    Power,      // ^
    
    // Comparison
    Equal,      // =
    NotEqual,   // <>
    LessThan,   // <
    LessEqual,  // <=
    GreaterThan,    // >
    GreaterEqual,   // >=
    
    // Text
    Concat,     // &
    
    // Range
    Range,      // : (creates range)
    Union,      // , (union of ranges)
    Intersect,  // (space) (intersection)
}

/// Unary operators
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOperator {
    Negate,     // - (negative)
    Percent,    // % (divide by 100)
    // Note: + is allowed but is a no-op
}
```

### 5.2 Formula Parser

```rust
/// Formula parser
/// 
/// Parses Excel formula syntax into AST.
/// Handles:
/// - Cell references (A1, $A$1, Sheet1!A1)
/// - Ranges (A1:B10)
/// - Operators with correct precedence
/// - Function calls with arguments
/// - Array constants
/// - Structured references (Table[Column])
pub struct FormulaParser<'a> {
    input: &'a str,
    pos: usize,
}

impl<'a> FormulaParser<'a> {
    pub fn parse(formula: &str) -> Result<FormulaExpr, ParseError> {
        // Formula must start with =
        let formula = formula.strip_prefix('=')
            .ok_or(ParseError::MissingEquals)?;
        
        let mut parser = FormulaParser { input: formula, pos: 0 };
        parser.parse_expression()
    }
    
    /// Parse expression with operator precedence
    fn parse_expression(&mut self) -> Result<FormulaExpr, ParseError> {
        self.parse_comparison()
    }
    
    /// Operator precedence (lowest to highest):
    /// 1. Comparison (=, <>, <, <=, >, >=)
    /// 2. Concatenation (&)
    /// 3. Addition/Subtraction (+, -)
    /// 4. Multiplication/Division (*, /)
    /// 5. Exponentiation (^)
    /// 6. Percent (%)
    /// 7. Negation (-)
    /// 8. Range (:), Union (,), Intersect ( )
    
    fn parse_comparison(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
    fn parse_concat(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
    fn parse_additive(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
    fn parse_multiplicative(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
    fn parse_power(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
    fn parse_unary(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
    fn parse_primary(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
    fn parse_function(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
    fn parse_reference(&mut self) -> Result<FormulaExpr, ParseError> { /* ... */ }
}

/// Parse errors
#[derive(Debug, thiserror::Error)]
pub enum ParseError {
    #[error("Formula must start with '='")]
    MissingEquals,
    
    #[error("Unexpected character '{0}' at position {1}")]
    UnexpectedChar(char, usize),
    
    #[error("Unexpected end of formula")]
    UnexpectedEnd,
    
    #[error("Unmatched parenthesis")]
    UnmatchedParen,
    
    #[error("Invalid cell reference: {0}")]
    InvalidReference(String),
    
    #[error("Invalid number: {0}")]
    InvalidNumber(String),
    
    #[error("Formula too complex (nesting level exceeded)")]
    TooNested,
}
```

### 5.3 Evaluator

```rust
/// Formula evaluation context
pub struct EvaluationContext<'a> {
    /// Reference to workbook
    workbook: &'a Workbook,
    
    /// Currently evaluating cell (for relative references)
    current_cell: CellAddress,
    
    /// Current worksheet index
    current_sheet: usize,
    
    /// Cells currently being calculated (for circular reference detection)
    calculating: HashSet<(usize, CellAddress)>,
    
    /// Calculation options
    options: &'a CalculationOptions,
}

/// Calculation options
pub struct CalculationOptions {
    /// Allow circular references with iterative calculation
    pub iterative: bool,
    
    /// Maximum iterations for circular references
    pub max_iterations: u32,
    
    /// Maximum change for convergence
    pub max_change: f64,
}

/// Formula evaluation result
pub type EvalResult = Result<FormulaValue, EvalError>;

/// Value types during formula evaluation
#[derive(Debug, Clone)]
pub enum FormulaValue {
    Number(f64),
    String(String),
    Boolean(bool),
    Error(CellError),
    Array(Vec<Vec<FormulaValue>>),
    Range(RangeReference),
    Empty,
}

/// Evaluation errors (different from cell errors)
#[derive(Debug, thiserror::Error)]
pub enum EvalError {
    #[error("Circular reference detected")]
    CircularReference,
    
    #[error("Reference to deleted cell")]
    DeletedReference,
    
    #[error("Unknown function: {0}")]
    UnknownFunction(String),
    
    #[error("Invalid argument count for {0}: expected {1}, got {2}")]
    ArgumentCount(String, String, usize),
}

/// Main evaluator
pub struct FormulaEvaluator;

impl FormulaEvaluator {
    /// Evaluate a formula expression
    pub fn evaluate(
        expr: &FormulaExpr,
        ctx: &mut EvaluationContext,
    ) -> EvalResult {
        match expr {
            FormulaExpr::Number(n) => Ok(FormulaValue::Number(*n)),
            FormulaExpr::String(s) => Ok(FormulaValue::String(s.clone())),
            FormulaExpr::Boolean(b) => Ok(FormulaValue::Boolean(*b)),
            FormulaExpr::Error(e) => Ok(FormulaValue::Error(*e)),
            
            FormulaExpr::CellRef(r) => Self::eval_cell_ref(r, ctx),
            FormulaExpr::RangeRef(r) => Self::eval_range_ref(r, ctx),
            FormulaExpr::NameRef(name) => Self::eval_name_ref(name, ctx),
            
            FormulaExpr::BinaryOp { op, left, right } => {
                Self::eval_binary(*op, left, right, ctx)
            }
            FormulaExpr::UnaryOp { op, operand } => {
                Self::eval_unary(*op, operand, ctx)
            }
            
            FormulaExpr::Function { name, args } => {
                Self::eval_function(name, args, ctx)
            }
            
            FormulaExpr::Array(rows) => Self::eval_array(rows, ctx),
            
            _ => Ok(FormulaValue::Error(CellError::Value)),
        }
    }
    
    // ... implementation methods
}
```

### 5.4 Dependency Graph

```rust
use std::collections::{HashMap, HashSet};

/// Tracks dependencies between cells for efficient recalculation
pub struct DependencyGraph {
    /// Cell → Cells that depend on it (dependents)
    /// When a cell changes, these cells need recalculation
    dependents: HashMap<CellKey, HashSet<CellKey>>,
    
    /// Cell → Cells it depends on (precedents)
    /// Used for tracing and validation
    precedents: HashMap<CellKey, HashSet<CellKey>>,
}

/// Unique key for a cell (sheet index + address)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellKey {
    pub sheet: usize,
    pub row: u32,
    pub col: u16,
}

impl DependencyGraph {
    /// Register that `dependent` depends on `precedent`
    pub fn add_dependency(&mut self, precedent: CellKey, dependent: CellKey) {
        self.dependents
            .entry(precedent)
            .or_default()
            .insert(dependent);
        self.precedents
            .entry(dependent)
            .or_default()
            .insert(precedent);
    }
    
    /// Remove all dependencies for a cell (when formula changes)
    pub fn clear_dependencies(&mut self, cell: CellKey) {
        if let Some(precedents) = self.precedents.remove(&cell) {
            for precedent in precedents {
                if let Some(deps) = self.dependents.get_mut(&precedent) {
                    deps.remove(&cell);
                }
            }
        }
    }
    
    /// Get cells that need recalculation when `cell` changes
    pub fn get_dependents(&self, cell: CellKey) -> impl Iterator<Item = CellKey> + '_ {
        self.dependents
            .get(&cell)
            .into_iter()
            .flat_map(|set| set.iter().copied())
    }
    
    /// Get calculation order (topological sort)
    /// Returns None if circular reference detected
    pub fn calculation_order(&self, dirty_cells: &[CellKey]) -> Option<Vec<CellKey>> {
        // Topological sort of dirty cells and their dependents
        // Returns None if cycle detected
        // Implementation uses Kahn's algorithm or DFS
    }
    
    /// Detect circular references
    pub fn find_circular_references(&self) -> Vec<Vec<CellKey>> {
        // Find all strongly connected components with more than one cell
        // or self-referencing cells
    }
}
```

### 5.5 Function Registry

```rust
/// Function implementation signature
pub type FunctionImpl = fn(&[FormulaValue], &mut EvaluationContext) -> EvalResult;

/// Function definition
pub struct FunctionDef {
    /// Function name (uppercase)
    pub name: &'static str,
    
    /// Minimum number of arguments
    pub min_args: usize,
    
    /// Maximum number of arguments (None = unlimited)
    pub max_args: Option<usize>,
    
    /// Implementation
    pub implementation: FunctionImpl,
    
    /// Whether function is volatile (always recalculates)
    pub volatile: bool,
    
    /// Category for documentation
    pub category: FunctionCategory,
}

/// Function categories
#[derive(Debug, Clone, Copy)]
pub enum FunctionCategory {
    Math,
    Statistical,
    Text,
    Logical,
    Lookup,
    DateTime,
    Financial,
    Information,
    Engineering,
    Database,
    Cube,
    Web,
    Compatibility,
}

/// Function registry
pub struct FunctionRegistry {
    functions: HashMap<String, FunctionDef>,
}

impl FunctionRegistry {
    /// Create registry with all built-in functions
    pub fn new() -> Self {
        let mut registry = Self { functions: HashMap::new() };
        
        // Register all function categories
        registry.register_math_functions();
        registry.register_statistical_functions();
        registry.register_text_functions();
        registry.register_logical_functions();
        registry.register_lookup_functions();
        registry.register_datetime_functions();
        registry.register_financial_functions();
        registry.register_info_functions();
        // ... etc
        
        registry
    }
    
    /// Look up a function by name
    pub fn get(&self, name: &str) -> Option<&FunctionDef> {
        self.functions.get(&name.to_uppercase())
    }
    
    fn register_math_functions(&mut self) {
        self.register(FunctionDef {
            name: "SUM",
            min_args: 1,
            max_args: None,
            implementation: functions::math::fn_sum,
            volatile: false,
            category: FunctionCategory::Math,
        });
        
        self.register(FunctionDef {
            name: "AVERAGE",
            min_args: 1,
            max_args: None,
            implementation: functions::math::fn_average,
            volatile: false,
            category: FunctionCategory::Math,
        });
        
        // ... many more
    }
}

// Example function implementations
pub mod functions {
    pub mod math {
        use super::*;
        
        pub fn fn_sum(args: &[FormulaValue], ctx: &mut EvaluationContext) -> EvalResult {
            let mut sum = 0.0;
            
            for arg in args {
                match arg {
                    FormulaValue::Number(n) => sum += n,
                    FormulaValue::Range(range) => {
                        for cell in ctx.iterate_range(range)? {
                            if let Some(n) = cell.as_number() {
                                sum += n;
                            }
                            // Ignore non-numeric, propagate errors
                            if let FormulaValue::Error(e) = cell {
                                return Ok(FormulaValue::Error(e));
                            }
                        }
                    }
                    FormulaValue::Error(e) => return Ok(FormulaValue::Error(*e)),
                    _ => {} // Ignore non-numeric
                }
            }
            
            Ok(FormulaValue::Number(sum))
        }
        
        pub fn fn_average(args: &[FormulaValue], ctx: &mut EvaluationContext) -> EvalResult {
            let mut sum = 0.0;
            let mut count = 0;
            
            for arg in args {
                // Similar to SUM but also counts
            }
            
            if count == 0 {
                Ok(FormulaValue::Error(CellError::Div0))
            } else {
                Ok(FormulaValue::Number(sum / count as f64))
            }
        }
    }
    
    pub mod logical {
        pub fn fn_if(args: &[FormulaValue], ctx: &mut EvaluationContext) -> EvalResult {
            // IF(condition, value_if_true, [value_if_false])
            let condition = args.get(0).ok_or(/* error */)?;
            let if_true = args.get(1).ok_or(/* error */)?;
            let if_false = args.get(2);
            
            let condition_bool = condition.as_bool()?;
            
            if condition_bool {
                Ok(if_true.clone())
            } else {
                Ok(if_false.cloned().unwrap_or(FormulaValue::Boolean(false)))
            }
        }
    }
    
    pub mod lookup {
        pub fn fn_vlookup(args: &[FormulaValue], ctx: &mut EvaluationContext) -> EvalResult {
            // VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
            // Implementation
        }
        
        pub fn fn_index(args: &[FormulaValue], ctx: &mut EvaluationContext) -> EvalResult {
            // INDEX(array, row_num, [col_num])
            // Implementation
        }
        
        pub fn fn_match(args: &[FormulaValue], ctx: &mut EvaluationContext) -> EvalResult {
            // MATCH(lookup_value, lookup_array, [match_type])
            // Implementation
        }
    }
}
```

---

## 6. File I/O

### 6.1 XLSX Format (Office Open XML)

```rust
/// XLSX reader
pub struct XlsxReader<R: Read + Seek> {
    archive: ZipArchive<R>,
    shared_strings: Vec<String>,
    styles: StyleSheet,
    workbook_rels: Relationships,
}

impl<R: Read + Seek> XlsxReader<R> {
    /// Open XLSX file
    pub fn new(reader: R) -> Result<Self, ReadError> {
        let mut archive = ZipArchive::new(reader)?;
        
        // Read required parts
        let shared_strings = Self::read_shared_strings(&mut archive)?;
        let styles = Self::read_styles(&mut archive)?;
        let workbook_rels = Self::read_relationships(&mut archive, "xl/_rels/workbook.xml.rels")?;
        
        Ok(Self { archive, shared_strings, styles, workbook_rels })
    }
    
    /// Read workbook structure
    pub fn read_workbook(&mut self) -> Result<Workbook, ReadError> {
        // Read xl/workbook.xml for sheet list
        // Read each worksheet
        // Build Workbook structure
    }
    
    /// Read single worksheet
    fn read_worksheet(&mut self, path: &str) -> Result<Worksheet, ReadError> {
        // Parse xl/worksheets/sheet1.xml
        // Build cells, styles, merges, etc.
    }
}

/// Streaming XLSX reader for large files
/// 
/// Uses SAX-style parsing to avoid loading entire file in memory.
pub struct StreamingXlsxReader<R: Read + Seek> {
    archive: ZipArchive<R>,
    shared_strings: SharedStringsReader,  // Streaming string reader
    styles: StyleSheet,
}

/// Callback interface for streaming reads
pub trait SheetHandler {
    /// Called when starting a new row
    fn start_row(&mut self, row_index: u32);
    
    /// Called for each cell
    fn cell(&mut self, col_index: u16, value: CellValue, style_index: u32);
    
    /// Called when row is complete
    fn end_row(&mut self);
}

impl<R: Read + Seek> StreamingXlsxReader<R> {
    /// Stream-read a worksheet
    pub fn read_sheet_streaming<H: SheetHandler>(
        &mut self,
        sheet_index: usize,
        handler: &mut H,
    ) -> Result<(), ReadError> {
        // SAX-parse worksheet XML, calling handler for each cell
        // Never loads full sheet into memory
    }
}
```

```rust
/// XLSX writer
pub struct XlsxWriter {
    workbook: Workbook,
    shared_strings: SharedStringTable,
    styles: StyleSheet,
}

impl XlsxWriter {
    /// Write workbook to XLSX
    pub fn write<W: Write + Seek>(workbook: &Workbook, writer: W) -> Result<(), WriteError> {
        let mut zip = ZipWriter::new(writer);
        
        // Write [Content_Types].xml
        Self::write_content_types(&mut zip)?;
        
        // Write _rels/.rels
        Self::write_root_rels(&mut zip)?;
        
        // Write xl/workbook.xml
        Self::write_workbook_xml(&mut zip, workbook)?;
        
        // Write xl/sharedStrings.xml
        Self::write_shared_strings(&mut zip, workbook)?;
        
        // Write xl/styles.xml
        Self::write_styles(&mut zip, workbook)?;
        
        // Write each worksheet
        for (i, sheet) in workbook.worksheets().enumerate() {
            Self::write_worksheet(&mut zip, sheet, i)?;
        }
        
        zip.finish()?;
        Ok(())
    }
}

/// Streaming XLSX writer for large files
/// 
/// Writes rows directly to disk, suitable for millions of rows.
pub struct StreamingXlsxWriter {
    temp_dir: TempDir,
    shared_strings: StreamingStringTable,
    styles: StyleSheet,
    sheets: Vec<StreamingSheetWriter>,
}

pub struct StreamingSheetWriter {
    path: PathBuf,
    writer: BufWriter<File>,
    current_row: u32,
}

impl StreamingXlsxWriter {
    pub fn new() -> Result<Self, WriteError> {
        let temp_dir = TempDir::new()?;
        Ok(Self {
            temp_dir,
            shared_strings: StreamingStringTable::new(),
            styles: StyleSheet::new(),
            sheets: Vec::new(),
        })
    }
    
    /// Add a new sheet
    pub fn add_sheet(&mut self, name: &str) -> Result<usize, WriteError> {
        let idx = self.sheets.len();
        let path = self.temp_dir.path().join(format!("sheet{}.xml", idx));
        let writer = BufWriter::new(File::create(&path)?);
        self.sheets.push(StreamingSheetWriter {
            path,
            writer,
            current_row: 0,
        });
        Ok(idx)
    }
    
    /// Write a row (must be in order)
    pub fn write_row(&mut self, sheet: usize, row: u32, cells: &[CellData]) -> Result<(), WriteError> {
        let sheet_writer = &mut self.sheets[sheet];
        if row < sheet_writer.current_row {
            return Err(WriteError::RowsOutOfOrder);
        }
        // Write row XML directly to file
        sheet_writer.current_row = row + 1;
        Ok(())
    }
    
    /// Finalize and create XLSX file
    pub fn finish<W: Write + Seek>(self, writer: W) -> Result<(), WriteError> {
        // Combine temp files into final ZIP
    }
}
```

### 6.2 XLS Format (BIFF8)

```rust
/// XLS (BIFF8) reader
pub struct XlsReader<R: Read + Seek> {
    cfb: CompoundFile<R>,
}

impl<R: Read + Seek> XlsReader<R> {
    pub fn new(reader: R) -> Result<Self, ReadError> {
        let cfb = CompoundFile::open(reader)?;
        Ok(Self { cfb })
    }
    
    pub fn read_workbook(&mut self) -> Result<Workbook, ReadError> {
        // Open "Workbook" or "Book" stream
        let stream = self.cfb.open_stream("Workbook")
            .or_else(|_| self.cfb.open_stream("Book"))?;
        
        // Parse BIFF8 records
        let mut reader = BiffReader::new(stream);
        let mut workbook = Workbook::new();
        
        while let Some(record) = reader.next_record()? {
            match record.record_type {
                BIFF_BOF => { /* Beginning of file */ }
                BIFF_SHEET => { /* Sheet definition */ }
                BIFF_SST => { /* Shared string table */ }
                BIFF_XF => { /* Cell format */ }
                BIFF_LABELSST => { /* String cell */ }
                BIFF_NUMBER => { /* Number cell */ }
                BIFF_FORMULA => { /* Formula cell */ }
                // ... many more record types
                _ => { /* Unknown record, skip */ }
            }
        }
        
        Ok(workbook)
    }
}

/// BIFF record reader
struct BiffReader<R: Read> {
    reader: R,
}

struct BiffRecord {
    record_type: u16,
    data: Vec<u8>,
}

// BIFF8 record type constants
const BIFF_BOF: u16 = 0x0809;
const BIFF_EOF: u16 = 0x000A;
const BIFF_SHEET: u16 = 0x0085;
const BIFF_SST: u16 = 0x00FC;
const BIFF_XF: u16 = 0x00E0;
const BIFF_LABELSST: u16 = 0x00FD;
const BIFF_NUMBER: u16 = 0x0203;
const BIFF_FORMULA: u16 = 0x0006;
// ... many more
```

### 6.3 CSV Format

```rust
/// CSV reader options
pub struct CsvReadOptions {
    /// Field delimiter (default: comma)
    pub delimiter: u8,
    
    /// Quote character (default: double quote)
    pub quote: u8,
    
    /// Whether first row is header
    pub has_header: bool,
    
    /// Text encoding
    pub encoding: Encoding,
    
    /// How to detect types (all strings, auto-detect, etc.)
    pub type_detection: TypeDetection,
}

/// CSV reader
pub struct CsvReader<R: Read> {
    reader: csv::Reader<R>,
    options: CsvReadOptions,
}

impl<R: Read> CsvReader<R> {
    pub fn read_to_worksheet(&mut self, sheet: &mut Worksheet) -> Result<(), ReadError> {
        let mut row_idx = 0u32;
        
        for result in self.reader.records() {
            let record = result?;
            
            for (col_idx, field) in record.iter().enumerate() {
                let value = self.detect_type(field);
                sheet.set_cell_value_at(row_idx, col_idx as u16, value)?;
            }
            
            row_idx += 1;
        }
        
        Ok(())
    }
    
    fn detect_type(&self, field: &str) -> CellValue {
        match self.options.type_detection {
            TypeDetection::AllStrings => CellValue::String(field.into()),
            TypeDetection::Auto => {
                // Try to parse as number, date, boolean, etc.
                if let Ok(n) = field.parse::<f64>() {
                    CellValue::Number(n)
                } else if field.eq_ignore_ascii_case("true") {
                    CellValue::Boolean(true)
                } else if field.eq_ignore_ascii_case("false") {
                    CellValue::Boolean(false)
                } else {
                    CellValue::String(field.into())
                }
            }
        }
    }
}

/// CSV writer
pub struct CsvWriter<W: Write> {
    writer: csv::Writer<W>,
    options: CsvWriteOptions,
}

pub struct CsvWriteOptions {
    pub delimiter: u8,
    pub quote_style: QuoteStyle,
    pub line_terminator: LineTerminator,
}
```

---

## 7. Chart Support

### 7.1 Chart Model

```rust
/// Chart types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartType {
    // Bar/Column
    ColumnClustered,
    ColumnStacked,
    ColumnPercentStacked,
    BarClustered,
    BarStacked,
    BarPercentStacked,
    
    // Line
    Line,
    LineStacked,
    LineMarkers,
    LineMarkersStacked,
    
    // Pie
    Pie,
    PieExploded,
    Doughnut,
    
    // Area
    Area,
    AreaStacked,
    AreaPercentStacked,
    
    // Scatter/XY
    ScatterMarkers,
    ScatterSmooth,
    ScatterSmoothMarkers,
    ScatterLines,
    ScatterLinesMarkers,
    
    // Other
    Bubble,
    Radar,
    Stock,
    Surface,
    Combo,
}

/// Chart definition
pub struct Chart {
    /// Chart type
    pub chart_type: ChartType,
    
    /// Position and size
    pub anchor: ChartAnchor,
    
    /// Chart title
    pub title: Option<ChartTitle>,
    
    /// Data series
    pub series: Vec<DataSeries>,
    
    /// Category axis (X)
    pub category_axis: Option<Axis>,
    
    /// Value axis (Y)
    pub value_axis: Option<Axis>,
    
    /// Secondary axes for combo charts
    pub secondary_category_axis: Option<Axis>,
    pub secondary_value_axis: Option<Axis>,
    
    /// Legend
    pub legend: Option<Legend>,
    
    /// Plot area formatting
    pub plot_area: PlotArea,
    
    /// Chart area formatting
    pub chart_area: ChartArea,
    
    /// 3D settings
    pub view_3d: Option<View3D>,
}

/// How chart is anchored to cells
pub enum ChartAnchor {
    /// Anchored to cell range
    TwoCellAnchor {
        from: AnchorPoint,
        to: AnchorPoint,
    },
    /// Absolute position
    AbsoluteAnchor {
        x: i64,  // EMUs
        y: i64,
        width: i64,
        height: i64,
    },
}

pub struct AnchorPoint {
    pub col: u16,
    pub col_offset: i64,  // EMUs from left of cell
    pub row: u32,
    pub row_offset: i64,  // EMUs from top of cell
}

/// Data series
pub struct DataSeries {
    /// Series name/title
    pub name: Option<SeriesName>,
    
    /// Values (Y axis data)
    pub values: DataReference,
    
    /// Categories (X axis data)
    pub categories: Option<DataReference>,
    
    /// For bubble charts: bubble sizes
    pub bubble_sizes: Option<DataReference>,
    
    /// Series formatting
    pub format: SeriesFormat,
    
    /// Data labels
    pub data_labels: Option<DataLabels>,
    
    /// Trendline
    pub trendline: Option<Trendline>,
    
    /// Error bars
    pub error_bars: Option<ErrorBars>,
}

/// Reference to data (can be formula or literal values)
pub enum DataReference {
    /// Formula reference (e.g., "Sheet1!$A$1:$A$10")
    Formula(String),
    
    /// Literal numeric values
    NumberLiteral(Vec<f64>),
    
    /// Literal string values
    StringLiteral(Vec<String>),
}

/// Axis configuration
pub struct Axis {
    /// Axis title
    pub title: Option<AxisTitle>,
    
    /// Number format
    pub number_format: Option<String>,
    
    /// Minimum value (None = auto)
    pub minimum: Option<f64>,
    
    /// Maximum value (None = auto)
    pub maximum: Option<f64>,
    
    /// Major unit
    pub major_unit: Option<f64>,
    
    /// Minor unit
    pub minor_unit: Option<f64>,
    
    /// Axis position
    pub position: AxisPosition,
    
    /// Major gridlines
    pub major_gridlines: Option<GridLines>,
    
    /// Minor gridlines
    pub minor_gridlines: Option<GridLines>,
    
    /// Axis line formatting
    pub line_format: LineFormat,
    
    /// Tick mark style
    pub major_tick_mark: TickMark,
    pub minor_tick_mark: TickMark,
    
    /// Tick label position
    pub tick_label_position: TickLabelPosition,
}

/// Legend configuration
pub struct Legend {
    pub position: LegendPosition,
    pub overlay: bool,
    pub format: ShapeFormat,
}

#[derive(Debug, Clone, Copy)]
pub enum LegendPosition {
    Bottom,
    Top,
    Left,
    Right,
    TopRight,
}
```

### 7.2 Chart Reading/Writing

```rust
/// Chart reader for XLSX
impl XlsxReader {
    fn read_chart(&mut self, chart_path: &str) -> Result<Chart, ReadError> {
        // Parse xl/charts/chart1.xml
        // Build Chart structure
    }
    
    fn read_drawing(&mut self, sheet_idx: usize) -> Result<Vec<Drawing>, ReadError> {
        // Parse xl/drawings/drawing1.xml
        // Find embedded charts
    }
}

/// Chart writer for XLSX
impl XlsxWriter {
    fn write_chart(&self, zip: &mut ZipWriter, chart: &Chart, idx: usize) -> Result<(), WriteError> {
        // Generate xl/charts/chart{idx}.xml
    }
    
    fn write_drawing(&self, zip: &mut ZipWriter, drawings: &[Drawing], sheet_idx: usize) -> Result<(), WriteError> {
        // Generate xl/drawings/drawing{idx}.xml
    }
}
```

---

## 8. C FFI Layer

### 8.1 Handle System

```rust
// duke-sheets-ffi/src/handles.rs

use std::collections::HashMap;
use std::sync::Mutex;

/// Opaque handle for C API
pub type Handle = u64;

/// Null/invalid handle constant
pub const HANDLE_NULL: Handle = 0;

/// Global context for managing all FFI objects
pub struct FfiContext {
    workbooks: HashMap<Handle, Workbook>,
    next_handle: Handle,
}

lazy_static::lazy_static! {
    static ref CONTEXT: Mutex<FfiContext> = Mutex::new(FfiContext {
        workbooks: HashMap::new(),
        next_handle: 1,  // Start at 1, 0 is null
    });
}

impl FfiContext {
    pub fn create_workbook(&mut self, wb: Workbook) -> Handle {
        let handle = self.next_handle;
        self.next_handle += 1;
        self.workbooks.insert(handle, wb);
        handle
    }
    
    pub fn get_workbook(&self, handle: Handle) -> Option<&Workbook> {
        self.workbooks.get(&handle)
    }
    
    pub fn get_workbook_mut(&mut self, handle: Handle) -> Option<&mut Workbook> {
        self.workbooks.get_mut(&handle)
    }
    
    pub fn destroy_workbook(&mut self, handle: Handle) -> bool {
        self.workbooks.remove(&handle).is_some()
    }
}

/// Helper macro for FFI functions
macro_rules! with_context {
    ($body:expr) => {
        match CONTEXT.lock() {
            Ok(ctx) => $body(ctx),
            Err(_) => CELLS_ERR_INTERNAL,
        }
    };
}
```

### 8.2 Error Codes

```rust
// duke-sheets-ffi/src/error.rs

use std::os::raw::c_int;

// Success
pub const CELLS_OK: c_int = 0;

// General errors
pub const CELLS_ERR_NULL_PTR: c_int = -1;
pub const CELLS_ERR_INVALID_HANDLE: c_int = -2;
pub const CELLS_ERR_INTERNAL: c_int = -3;

// I/O errors
pub const CELLS_ERR_FILE_NOT_FOUND: c_int = -10;
pub const CELLS_ERR_PERMISSION_DENIED: c_int = -11;
pub const CELLS_ERR_IO: c_int = -12;

// Format errors
pub const CELLS_ERR_INVALID_FORMAT: c_int = -20;
pub const CELLS_ERR_CORRUPT_FILE: c_int = -21;
pub const CELLS_ERR_UNSUPPORTED_VERSION: c_int = -22;

// Data errors
pub const CELLS_ERR_OUT_OF_BOUNDS: c_int = -30;
pub const CELLS_ERR_INVALID_ARGUMENT: c_int = -31;
pub const CELLS_ERR_BUFFER_TOO_SMALL: c_int = -32;

// Formula errors
pub const CELLS_ERR_FORMULA_PARSE: c_int = -40;
pub const CELLS_ERR_CIRCULAR_REF: c_int = -41;

/// Get human-readable error message
#[no_mangle]
pub extern "C" fn cells_error_message(code: c_int) -> *const c_char {
    let msg = match code {
        CELLS_OK => "Success",
        CELLS_ERR_NULL_PTR => "Null pointer argument",
        CELLS_ERR_INVALID_HANDLE => "Invalid handle",
        // ... etc
        _ => "Unknown error",
    };
    
    // Return static string pointer
    msg.as_ptr() as *const c_char
}
```

### 8.3 Workbook Functions

```rust
// duke-sheets-ffi/src/workbook.rs

use std::ffi::{CStr, CString};
use std::os::raw::{c_char, c_int};
use super::*;

/// Create a new empty workbook
#[no_mangle]
pub extern "C" fn cells_workbook_new(out_handle: *mut Handle) -> c_int {
    if out_handle.is_null() {
        return CELLS_ERR_NULL_PTR;
    }
    
    with_context!(|mut ctx| {
        let wb = Workbook::new();
        let handle = ctx.create_workbook(wb);
        unsafe { *out_handle = handle; }
        CELLS_OK
    })
}

/// Open workbook from file
#[no_mangle]
pub extern "C" fn cells_workbook_open(
    path: *const c_char,
    out_handle: *mut Handle,
) -> c_int {
    if path.is_null() || out_handle.is_null() {
        return CELLS_ERR_NULL_PTR;
    }
    
    let path_str = unsafe {
        match CStr::from_ptr(path).to_str() {
            Ok(s) => s,
            Err(_) => return CELLS_ERR_INVALID_ARGUMENT,
        }
    };
    
    with_context!(|mut ctx| {
        match Workbook::open(path_str) {
            Ok(wb) => {
                let handle = ctx.create_workbook(wb);
                unsafe { *out_handle = handle; }
                CELLS_OK
            }
            Err(e) => error_to_code(&e),
        }
    })
}

/// Save workbook to file
#[no_mangle]
pub extern "C" fn cells_workbook_save(
    handle: Handle,
    path: *const c_char,
) -> c_int {
    if path.is_null() {
        return CELLS_ERR_NULL_PTR;
    }
    
    let path_str = unsafe {
        match CStr::from_ptr(path).to_str() {
            Ok(s) => s,
            Err(_) => return CELLS_ERR_INVALID_ARGUMENT,
        }
    };
    
    with_context!(|ctx| {
        match ctx.get_workbook(handle) {
            Some(wb) => {
                match wb.save(path_str) {
                    Ok(()) => CELLS_OK,
                    Err(e) => error_to_code(&e),
                }
            }
            None => CELLS_ERR_INVALID_HANDLE,
        }
    })
}

/// Free workbook
#[no_mangle]
pub extern "C" fn cells_workbook_free(handle: Handle) -> c_int {
    with_context!(|mut ctx| {
        if ctx.destroy_workbook(handle) {
            CELLS_OK
        } else {
            CELLS_ERR_INVALID_HANDLE
        }
    })
}

/// Calculate all formulas
#[no_mangle]
pub extern "C" fn cells_workbook_calculate(handle: Handle) -> c_int {
    with_context!(|mut ctx| {
        match ctx.get_workbook_mut(handle) {
            Some(wb) => {
                match wb.calculate() {
                    Ok(()) => CELLS_OK,
                    Err(e) => error_to_code(&e),
                }
            }
            None => CELLS_ERR_INVALID_HANDLE,
        }
    })
}

/// Get number of worksheets
#[no_mangle]
pub extern "C" fn cells_workbook_sheet_count(
    handle: Handle,
    out_count: *mut c_int,
) -> c_int {
    if out_count.is_null() {
        return CELLS_ERR_NULL_PTR;
    }
    
    with_context!(|ctx| {
        match ctx.get_workbook(handle) {
            Some(wb) => {
                unsafe { *out_count = wb.sheet_count() as c_int; }
                CELLS_OK
            }
            None => CELLS_ERR_INVALID_HANDLE,
        }
    })
}
```

### 8.4 Cell Functions

```rust
// duke-sheets-ffi/src/cell.rs

/// Get cell value as string
#[no_mangle]
pub extern "C" fn cells_get_string(
    handle: Handle,
    sheet: c_int,
    row: c_int,
    col: c_int,
    buffer: *mut c_char,
    buffer_size: c_int,
    out_len: *mut c_int,
) -> c_int {
    if buffer.is_null() || out_len.is_null() {
        return CELLS_ERR_NULL_PTR;
    }
    
    with_context!(|ctx| {
        let wb = match ctx.get_workbook(handle) {
            Some(wb) => wb,
            None => return CELLS_ERR_INVALID_HANDLE,
        };
        
        let sheet = match wb.worksheet(sheet as usize) {
            Some(s) => s,
            None => return CELLS_ERR_OUT_OF_BOUNDS,
        };
        
        let cell = match sheet.cell_at(row as u32, col as u16) {
            Some(c) => c,
            None => {
                // Empty cell
                unsafe { *out_len = 0; }
                if buffer_size > 0 {
                    unsafe { *buffer = 0; }
                }
                return CELLS_OK;
            }
        };
        
        let value_str = cell.to_string();
        let bytes = value_str.as_bytes();
        
        unsafe { *out_len = bytes.len() as c_int; }
        
        if buffer_size < (bytes.len() + 1) as c_int {
            return CELLS_ERR_BUFFER_TOO_SMALL;
        }
        
        unsafe {
            std::ptr::copy_nonoverlapping(
                bytes.as_ptr(),
                buffer as *mut u8,
                bytes.len(),
            );
            *buffer.add(bytes.len()) = 0;  // Null terminator
        }
        
        CELLS_OK
    })
}

/// Set cell string value
#[no_mangle]
pub extern "C" fn cells_set_string(
    handle: Handle,
    sheet: c_int,
    row: c_int,
    col: c_int,
    value: *const c_char,
) -> c_int {
    let value_str = if value.is_null() {
        ""
    } else {
        unsafe {
            match CStr::from_ptr(value).to_str() {
                Ok(s) => s,
                Err(_) => return CELLS_ERR_INVALID_ARGUMENT,
            }
        }
    };
    
    with_context!(|mut ctx| {
        let wb = match ctx.get_workbook_mut(handle) {
            Some(wb) => wb,
            None => return CELLS_ERR_INVALID_HANDLE,
        };
        
        let sheet = match wb.worksheet_mut(sheet as usize) {
            Some(s) => s,
            None => return CELLS_ERR_OUT_OF_BOUNDS,
        };
        
        match sheet.set_cell_value_at(row as u32, col as u16, value_str) {
            Ok(()) => CELLS_OK,
            Err(e) => error_to_code(&e),
        }
    })
}

/// Set cell numeric value
#[no_mangle]
pub extern "C" fn cells_set_number(
    handle: Handle,
    sheet: c_int,
    row: c_int,
    col: c_int,
    value: c_double,
) -> c_int {
    with_context!(|mut ctx| {
        // Similar implementation
        CELLS_OK
    })
}

/// Set cell formula
#[no_mangle]
pub extern "C" fn cells_set_formula(
    handle: Handle,
    sheet: c_int,
    row: c_int,
    col: c_int,
    formula: *const c_char,
) -> c_int {
    // Implementation
    CELLS_OK
}
```

### 8.5 C Header File

```c
// duke-sheets-ffi/include/duke_sheets.h

#ifndef DUKE_SHEETS_H
#define DUKE_SHEETS_H

#include <stdint.h>
#include <stdbool.h>

#ifdef __cplusplus
extern "C" {
#endif

/*
 * Duke Sheets - C API
 * 
 * All functions return an error code (0 = success, negative = error).
 * Handle values are opaque 64-bit integers.
 */

/* Handle type */
typedef uint64_t cells_handle_t;
#define CELLS_HANDLE_NULL 0

/* Error codes */
#define CELLS_OK                     0
#define CELLS_ERR_NULL_PTR          -1
#define CELLS_ERR_INVALID_HANDLE    -2
#define CELLS_ERR_INTERNAL          -3
#define CELLS_ERR_FILE_NOT_FOUND    -10
#define CELLS_ERR_PERMISSION_DENIED -11
#define CELLS_ERR_IO                -12
#define CELLS_ERR_INVALID_FORMAT    -20
#define CELLS_ERR_CORRUPT_FILE      -21
#define CELLS_ERR_UNSUPPORTED       -22
#define CELLS_ERR_OUT_OF_BOUNDS     -30
#define CELLS_ERR_INVALID_ARGUMENT  -31
#define CELLS_ERR_BUFFER_TOO_SMALL  -32
#define CELLS_ERR_FORMULA_PARSE     -40
#define CELLS_ERR_CIRCULAR_REF      -41

/* Error handling */
const char* cells_error_message(int code);

/* ============ Workbook Functions ============ */

/**
 * Create a new empty workbook.
 * 
 * @param out_handle Output: Handle to the new workbook
 * @return Error code
 */
int cells_workbook_new(cells_handle_t* out_handle);

/**
 * Open a workbook from file.
 * 
 * Supports: .xlsx, .xls, .csv
 * 
 * @param path Path to the file
 * @param out_handle Output: Handle to the workbook
 * @return Error code
 */
int cells_workbook_open(const char* path, cells_handle_t* out_handle);

/**
 * Save workbook to file.
 * 
 * Format is determined by file extension.
 * 
 * @param handle Workbook handle
 * @param path Output file path
 * @return Error code
 */
int cells_workbook_save(cells_handle_t handle, const char* path);

/**
 * Free a workbook and all associated resources.
 * 
 * @param handle Workbook handle
 * @return Error code
 */
int cells_workbook_free(cells_handle_t handle);

/**
 * Calculate all formulas in the workbook.
 * 
 * @param handle Workbook handle
 * @return Error code
 */
int cells_workbook_calculate(cells_handle_t handle);

/**
 * Get the number of worksheets.
 * 
 * @param handle Workbook handle
 * @param out_count Output: Number of worksheets
 * @return Error code
 */
int cells_workbook_sheet_count(cells_handle_t handle, int* out_count);

/* ============ Worksheet Functions ============ */

/**
 * Get worksheet name.
 * 
 * @param handle Workbook handle
 * @param sheet_index Worksheet index (0-based)
 * @param buffer Output buffer for name
 * @param buffer_size Size of buffer
 * @param out_len Output: Actual length of name
 * @return Error code
 */
int cells_worksheet_get_name(
    cells_handle_t handle,
    int sheet_index,
    char* buffer,
    int buffer_size,
    int* out_len
);

/**
 * Set worksheet name.
 * 
 * @param handle Workbook handle
 * @param sheet_index Worksheet index
 * @param name New name
 * @return Error code
 */
int cells_worksheet_set_name(
    cells_handle_t handle,
    int sheet_index,
    const char* name
);

/**
 * Add a new worksheet.
 * 
 * @param handle Workbook handle
 * @param name Sheet name (or NULL for default)
 * @param out_index Output: Index of new sheet
 * @return Error code
 */
int cells_worksheet_add(
    cells_handle_t handle,
    const char* name,
    int* out_index
);

/* ============ Cell Functions ============ */

/**
 * Get cell value as string.
 * 
 * Numbers and dates are formatted according to cell format.
 * 
 * @param handle Workbook handle
 * @param sheet Worksheet index
 * @param row Row index (0-based)
 * @param col Column index (0-based)
 * @param buffer Output buffer
 * @param buffer_size Size of buffer
 * @param out_len Output: Actual length
 * @return Error code
 */
int cells_get_string(
    cells_handle_t handle,
    int sheet,
    int row,
    int col,
    char* buffer,
    int buffer_size,
    int* out_len
);

/**
 * Get cell numeric value.
 * 
 * @param handle Workbook handle
 * @param sheet Worksheet index
 * @param row Row index
 * @param col Column index
 * @param out_value Output: Numeric value
 * @return Error code (error if cell is not numeric)
 */
int cells_get_number(
    cells_handle_t handle,
    int sheet,
    int row,
    int col,
    double* out_value
);

/**
 * Set cell string value.
 */
int cells_set_string(
    cells_handle_t handle,
    int sheet,
    int row,
    int col,
    const char* value
);

/**
 * Set cell numeric value.
 */
int cells_set_number(
    cells_handle_t handle,
    int sheet,
    int row,
    int col,
    double value
);

/**
 * Set cell formula.
 * 
 * @param formula Formula string including '=' (e.g., "=SUM(A1:A10)")
 */
int cells_set_formula(
    cells_handle_t handle,
    int sheet,
    int row,
    int col,
    const char* formula
);

/**
 * Get cell formula (if any).
 */
int cells_get_formula(
    cells_handle_t handle,
    int sheet,
    int row,
    int col,
    char* buffer,
    int buffer_size,
    int* out_len
);

/* ============ Style Functions ============ */

/**
 * Set cell font bold.
 */
int cells_set_bold(
    cells_handle_t handle,
    int sheet,
    int row,
    int col,
    bool bold
);

/**
 * Set cell number format.
 * 
 * @param format Excel number format string (e.g., "#,##0.00", "yyyy-mm-dd")
 */
int cells_set_number_format(
    cells_handle_t handle,
    int sheet,
    int row,
    int col,
    const char* format
);

/* ============ Range Functions ============ */

/**
 * Get used range dimensions.
 * 
 * @param out_start_row Output: First row with data
 * @param out_start_col Output: First column with data
 * @param out_end_row Output: Last row with data
 * @param out_end_col Output: Last column with data
 */
int cells_get_used_range(
    cells_handle_t handle,
    int sheet,
    int* out_start_row,
    int* out_start_col,
    int* out_end_row,
    int* out_end_col
);

#ifdef __cplusplus
}
#endif

#endif /* DUKE_SHEETS_H */
```

---

## 9. Testing Strategy

### 9.1 Philosophy: End-to-End First

Unit tests are only used for things that are hard to test end-to-end. The primary testing approach is:

1. **Create** a spreadsheet programmatically
2. **Save** to file (XLSX/XLS/CSV)
3. **Read** back from file
4. **Verify** all data matches

This ensures the full stack works together and catches integration issues.

### 9.2 Test Categories

| Category | Purpose | Example |
|----------|---------|---------|
| **Roundtrip** | Write then read, verify data | Create cells, save XLSX, read back, compare |
| **Excel Compatibility** | Read files from real Excel | Open file created in Excel, verify values |
| **Formula Calculation** | Test formula engine | Set formula, calculate, verify result |
| **Large Files** | Test streaming APIs | Read/write 1M+ rows without OOM |
| **Format Conversion** | Cross-format tests | XLS → XLSX, XLSX → CSV |
| **Edge Cases** | Unusual but valid files | Unicode, max values, special characters |

### 9.3 End-to-End Test Examples

```rust
// tests/e2e/roundtrip.rs

#[test]
fn test_basic_values_roundtrip() {
    // Create workbook with various value types
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    
    // Different value types
    sheet.set_cell_value("A1", "Hello").unwrap();
    sheet.set_cell_value("A2", 42.5).unwrap();
    sheet.set_cell_value("A3", true).unwrap();
    sheet.set_cell_value("A4", "").unwrap();  // Empty string
    
    // Save and reload
    let temp = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    wb.save(temp.path()).unwrap();
    
    let wb2 = Workbook::open(temp.path()).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    
    // Verify
    assert_eq!(sheet2.cell("A1").unwrap().as_string(), Some("Hello"));
    assert_eq!(sheet2.cell("A2").unwrap().as_number(), Some(42.5));
    assert_eq!(sheet2.cell("A3").unwrap().as_bool(), Some(true));
    assert!(sheet2.cell("A4").unwrap().is_empty() || 
            sheet2.cell("A4").unwrap().as_string() == Some(""));
}

#[test]
fn test_formulas_calculate_correctly() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    
    // Set up data
    for i in 1..=10 {
        sheet.set_cell_value(&format!("A{}", i), i as f64).unwrap();
    }
    
    // Set formula
    sheet.set_cell_formula("B1", "=SUM(A1:A10)").unwrap();
    sheet.set_cell_formula("B2", "=AVERAGE(A1:A10)").unwrap();
    
    // Save, reload, calculate
    let temp = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    wb.save(temp.path()).unwrap();
    
    let mut wb2 = Workbook::open(temp.path()).unwrap();
    wb2.calculate().unwrap();
    
    let sheet2 = wb2.worksheet(0).unwrap();
    assert_eq!(sheet2.cell("B1").unwrap().as_number(), Some(55.0));
    assert_eq!(sheet2.cell("B2").unwrap().as_number(), Some(5.5));
}

#[test]
fn test_styles_preserved() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    
    sheet.set_cell_value("A1", "Bold").unwrap();
    sheet.set_cell_style("A1", Style::new().bold(true)).unwrap();
    
    sheet.set_cell_value("A2", 1234.5).unwrap();
    sheet.set_cell_style("A2", Style::new().number_format("#,##0.00")).unwrap();
    
    // Roundtrip
    let temp = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    wb.save(temp.path()).unwrap();
    
    let wb2 = Workbook::open(temp.path()).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    
    assert!(sheet2.cell("A1").unwrap().style().font().bold());
    assert_eq!(
        sheet2.cell("A2").unwrap().style().number_format(),
        "#,##0.00"
    );
}
```

### 9.4 Excel Compatibility Tests

```rust
// tests/e2e/excel_compat.rs

/// Test files created in actual Excel
/// These are ground truth for correctness
#[test]
fn test_read_excel_formulas() {
    // This file was created in Excel 365, saved with calculated values
    let wb = Workbook::open("tests/fixtures/excel_created/formulas.xlsx").unwrap();
    let sheet = wb.worksheet(0).unwrap();
    
    // Excel's calculated values (our ground truth)
    let expected: Vec<(&str, f64)> = vec![
        ("D2", 150.0),    // =SUM(A1:A5)
        ("D3", 30.0),     // =AVERAGE(A1:A5)
        ("D4", 10.0),     // =MIN(A1:A5)
        ("D5", 50.0),     // =MAX(A1:A5)
        ("D6", 5.0),      // =COUNT(A1:A5)
        ("D7", 100.0),    // =IF(A1>5, 100, 0)
        ("D8", 30.0),     // =VLOOKUP(3, A1:B5, 2, FALSE)
    ];
    
    // First verify we read Excel's cached values correctly
    for (addr, value) in &expected {
        let cell = sheet.cell(addr).unwrap();
        assert_eq!(cell.as_number(), Some(*value),
            "Excel cached value mismatch at {}", addr);
    }
    
    // Then verify our calculation matches Excel
    let mut wb = wb;
    wb.calculate().unwrap();
    
    for (addr, value) in &expected {
        let cell = wb.worksheet(0).unwrap().cell(addr).unwrap();
        assert!((cell.as_number().unwrap() - value).abs() < 1e-10,
            "Calculated value mismatch at {}: expected {}, got {:?}",
            addr, value, cell.as_number());
    }
}

#[test]
fn test_read_excel_styles() {
    let wb = Workbook::open("tests/fixtures/excel_created/styles.xlsx").unwrap();
    let sheet = wb.worksheet(0).unwrap();
    
    // Bold cell
    assert!(sheet.cell("A1").unwrap().style().font().bold());
    
    // Italic cell
    assert!(sheet.cell("A2").unwrap().style().font().italic());
    
    // Red background
    let fill = sheet.cell("A3").unwrap().style().fill();
    assert_eq!(fill, FillStyle::Solid(Color::Rgb { r: 255, g: 0, b: 0 }));
    
    // Currency format
    let fmt = sheet.cell("A4").unwrap().style().number_format();
    assert!(fmt.contains("$") || fmt.contains("Currency"));
}
```

### 9.5 Large File Tests

```rust
// tests/e2e/large_files.rs

#[test]
fn test_read_large_file_streaming() {
    // 100K rows test file
    let wb = Workbook::open_streaming("tests/fixtures/large_100k.xlsx").unwrap();
    
    let mut row_count = 0;
    wb.stream_sheet(0, |row_idx, cells| {
        row_count += 1;
        // Process cells
        Ok(())
    }).unwrap();
    
    assert!(row_count >= 100_000);
}

#[test]
fn test_write_large_file_streaming() {
    let mut writer = StreamingXlsxWriter::new().unwrap();
    let sheet = writer.add_sheet("Data").unwrap();
    
    // Write 1M rows
    for row in 0..1_000_000 {
        writer.write_row(sheet, row, &[
            CellData::number(row as f64),
            CellData::string(&format!("Row {}", row)),
        ]).unwrap();
    }
    
    let temp = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    writer.finish(temp.path()).unwrap();
    
    // Verify file is readable
    let wb = Workbook::open(temp.path()).unwrap();
    assert_eq!(wb.worksheet(0).unwrap().used_range().row_count(), 1_000_000);
}
```

### 9.6 Fuzzing

```rust
// fuzz/fuzz_targets/xlsx_reader.rs
#![no_main]
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    // Should never panic on arbitrary input
    let _ = duke_sheets::Workbook::from_bytes(data);
});
```

```rust
// fuzz/fuzz_targets/formula_parser.rs
#![no_main]
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &str| {
    // Should never panic on arbitrary formula strings
    let _ = duke_sheets_formula::parse_formula(data);
});
```

### 9.7 Minimal Unit Tests

Unit tests only for pure functions that are hard to trigger via E2E:

```rust
// In duke-sheets-core/src/cell/address.rs

#[cfg(test)]
mod tests {
    use super::*;
    
    // These edge cases are hard to trigger via file I/O
    #[test]
    fn test_column_letters_max() {
        // Maximum column in Excel is XFD (16383)
        assert_eq!(CellAddress::column_to_letters(16383), "XFD");
        assert_eq!(CellAddress::letters_to_column("XFD").unwrap(), 16383);
    }
    
    #[test]
    fn test_address_parse_edge_cases() {
        assert!(CellAddress::parse("").is_err());
        assert!(CellAddress::parse("123").is_err());
        assert!(CellAddress::parse("A").is_err());
        assert!(CellAddress::parse("A0").is_err());  // Row 0 invalid in A1 notation
    }
}
```

---

## 10. Dependencies

```toml
[workspace.package]
version = "0.1.0"
edition = "2021"
license = "MIT OR Apache-2.0"
repository = "https://github.com/yourorg/duke-sheets"

[workspace.dependencies]
# Core
thiserror = "1.0"
log = "0.4"

# XML parsing (for XLSX)
quick-xml = "0.31"

# ZIP handling (for XLSX)
zip = "0.6"

# Compound File Binary (for XLS)
cfb = "0.9"

# CSV handling
csv = "1.3"

# Date/time
chrono = { version = "0.4", default-features = false, features = ["std"] }

# Decimal precision for financial calculations
rust_decimal = "1.33"

# Regex (for formulas, number formats)
regex = "1.10"

# Lazy static (for FFI global context)
lazy_static = "1.4"

# Hash functions (for style deduplication)
ahash = "0.8"

# Optional: async support
tokio = { version = "1", features = ["fs", "io-util"], optional = true }

[workspace.dev-dependencies]
# Testing
tempfile = "3"
criterion = { version = "0.5", features = ["html_reports"] }
proptest = "1.4"

# FFI header generation
cbindgen = "0.26"
```

---

## 11. API Design

### 11.1 Rust API

```rust
use duke_sheets::prelude::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create new workbook
    let mut wb = Workbook::new();
    
    // Access worksheet
    let sheet = wb.worksheet_mut(0)?;
    sheet.set_name("Sales Data")?;
    
    // Set cell values
    sheet.set_cell_value("A1", "Product")?;
    sheet.set_cell_value("B1", "Revenue")?;
    sheet.set_cell_value("A2", "Widget")?;
    sheet.set_cell_value("B2", 1500.0)?;
    sheet.set_cell_value("A3", "Gadget")?;
    sheet.set_cell_value("B3", 2500.0)?;
    
    // Set formula
    sheet.set_cell_formula("B4", "=SUM(B2:B3)")?;
    
    // Style header row
    let header_style = Style::new()
        .bold(true)
        .fill(FillStyle::Solid(Color::rgb(200, 200, 200)));
    sheet.set_row_style(0, &header_style)?;
    
    // Calculate formulas
    wb.calculate()?;
    
    // Save
    wb.save("sales.xlsx")?;
    
    // Read existing file
    let wb = Workbook::open("report.xlsx")?;
    
    // Iterate cells
    for sheet in wb.worksheets() {
        println!("Sheet: {}", sheet.name());
        for row in sheet.used_range().rows() {
            for cell in row.cells() {
                print!("{:?}\t", cell.value());
            }
            println!();
        }
    }
    
    Ok(())
}
```

### 11.2 C API Usage

```c
#include "duke_sheets.h"
#include <stdio.h>

int main() {
    cells_handle_t wb;
    int err;
    
    // Create workbook
    err = cells_workbook_new(&wb);
    if (err != CELLS_OK) {
        printf("Error: %s\n", cells_error_message(err));
        return 1;
    }
    
    // Set values
    cells_set_string(wb, 0, 0, 0, "Product");
    cells_set_string(wb, 0, 0, 1, "Revenue");
    cells_set_string(wb, 0, 1, 0, "Widget");
    cells_set_number(wb, 0, 1, 1, 1500.0);
    cells_set_formula(wb, 0, 2, 1, "=SUM(B2:B2)");
    
    // Calculate and save
    cells_workbook_calculate(wb);
    cells_workbook_save(wb, "output.xlsx");
    
    // Read a cell
    char buffer[256];
    int len;
    cells_get_string(wb, 0, 2, 1, buffer, sizeof(buffer), &len);
    printf("B3 = %s\n", buffer);
    
    // Cleanup
    cells_workbook_free(wb);
    return 0;
}
```

---

## 12. Implementation Phases

### Phase 1: Core + Basic XLSX (6-8 weeks)

**Deliverables:**
- Core data structures (Workbook, Worksheet, Cell, CellValue)
- Sparse cell storage with BTreeMap
- Style system with deduplication
- Cell address parsing (A1, $A$1, ranges)
- XLSX reader (DOM-style for small files)
- XLSX writer
- CSV reader/writer
- Basic Rust API

**E2E Tests:**
- Roundtrip: create → save → read → verify
- Read Excel-created files
- Style preservation

### Phase 2: Formula Engine (8-10 weeks)

**Deliverables:**
- Formula parser (text → AST)
- Expression evaluator
- ~450 built-in functions
- Dependency graph
- Calculation chain
- Circular reference handling
- Array formulas

**E2E Tests:**
- Calculate formulas, compare to Excel's cached values
- Complex formula chains
- Circular reference detection

### Phase 3: Large File Support (4-6 weeks)

**Deliverables:**
- Streaming XLSX reader (SAX-style)
- Streaming XLSX writer
- Memory-optimized cell storage mode
- Progress callbacks

**E2E Tests:**
- Read/write 1M+ row files
- Memory usage verification
- Performance benchmarks

### Phase 4: XLS Support (4-6 weeks)

**Deliverables:**
- Compound File Binary (CFB) reader
- BIFF8 record parsing
- XLS reader
- XLS writer
- Format conversion (XLS ↔ XLSX)

**E2E Tests:**
- Read legacy XLS files
- Roundtrip XLS files
- Cross-format conversion

### Phase 5: Charts (6-8 weeks)

**Deliverables:**
- Chart data model
- Read charts from XLSX
- Write charts to XLSX
- Create charts via API
- All major chart types (bar, line, pie, scatter, etc.)

**E2E Tests:**
- Preserve charts on read/write
- Create charts, verify in Excel
- Modify chart data

### Phase 6: C FFI (3-4 weeks)

**Deliverables:**
- Handle-based FFI design
- Core FFI functions
- Error handling
- C header generation
- Documentation
- Example C program

**E2E Tests:**
- C program integration tests
- Memory safety (valgrind/ASan)
- Thread safety

### Phase 7: Polish & Advanced Features (Ongoing)

**Deliverables:**
- Conditional formatting
- Data validation
- Named ranges
- Comments
- Images/Pictures
- Pivot tables (read-only)
- Print settings
- Documentation
- Performance optimization

---

## 13. Success Criteria

### Correctness
- [ ] All formula calculations match Excel's results
- [ ] Files round-trip without data loss
- [ ] Read files from Excel, LibreOffice, Google Sheets

### Compatibility
- [ ] Output files open correctly in Excel
- [ ] Output files open correctly in LibreOffice
- [ ] Output files open correctly in Google Sheets

### Performance
- [ ] Read 1M cells from XLSX in <10 seconds
- [ ] Write 1M cells to XLSX in <15 seconds
- [ ] Memory <500MB for 1M cell file in streaming mode

### Reliability
- [ ] No panics on malformed input (fuzz tested)
- [ ] Clear error messages for all failure modes
- [ ] Thread-safe read operations

### Usability
- [ ] Clean, documented Rust API
- [ ] Working C FFI with header file
- [ ] Examples for common use cases
- [ ] API documentation

---

## Appendix A: Excel Function Categories

### Math & Trig (~60 functions)
SUM, AVERAGE, COUNT, MIN, MAX, ABS, ROUND, FLOOR, CEILING, MOD, POWER, SQRT, LOG, LN, EXP, SIN, COS, TAN, PI, RAND, RANDBETWEEN, SUMIF, SUMIFS, SUMPRODUCT, ...

### Statistical (~80 functions)
STDEV, VAR, MEDIAN, MODE, PERCENTILE, QUARTILE, CORREL, COVAR, FORECAST, TREND, GROWTH, FREQUENCY, RANK, LARGE, SMALL, ...

### Text (~40 functions)
CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, FIND, SEARCH, REPLACE, SUBSTITUTE, LOWER, UPPER, PROPER, TRIM, TEXT, VALUE, ...

### Logical (~10 functions)
IF, AND, OR, NOT, XOR, TRUE, FALSE, IFERROR, IFNA, IFS, SWITCH, ...

### Lookup & Reference (~20 functions)
VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, OFFSET, INDIRECT, ROW, COLUMN, ROWS, COLUMNS, CHOOSE, LOOKUP, ...

### Date & Time (~25 functions)
DATE, TIME, NOW, TODAY, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, WEEKNUM, DATEVALUE, TIMEVALUE, DATEDIF, EDATE, EOMONTH, NETWORKDAYS, WORKDAY, ...

### Financial (~50 functions)
NPV, IRR, PMT, IPMT, PPMT, FV, PV, RATE, NPER, SLN, DDB, VDB, XNPV, XIRR, ...

### Information (~20 functions)
ISERROR, ISNA, ISBLANK, ISNUMBER, ISTEXT, ISLOGICAL, ISREF, TYPE, ERROR.TYPE, CELL, INFO, ...

### Engineering (~40 functions)
CONVERT, BIN2DEC, DEC2BIN, HEX2DEC, COMPLEX, IMAGINARY, IMREAL, IMSUM, IMPRODUCT, ...

### Database (~12 functions)
DSUM, DAVERAGE, DCOUNT, DMAX, DMIN, DGET, DPRODUCT, DSTDEV, DVAR, ...

---

## Appendix B: Chart Types

| Category | Types |
|----------|-------|
| **Column** | Clustered, Stacked, 100% Stacked, 3D variants |
| **Bar** | Clustered, Stacked, 100% Stacked, 3D variants |
| **Line** | Line, Stacked, 100% Stacked, With Markers, 3D |
| **Pie** | Pie, Exploded, 3D, Doughnut |
| **Area** | Area, Stacked, 100% Stacked, 3D variants |
| **Scatter** | Markers Only, Lines, Smooth Lines |
| **Stock** | High-Low-Close, Open-High-Low-Close, Volume variants |
| **Surface** | 3D Surface, Wireframe, Contour |
| **Radar** | Radar, Filled |
| **Combo** | Combined chart types |
| **Other** | Bubble, Funnel, Waterfall, Treemap, Sunburst |

---

## Appendix C: Reference Documents

- [ECMA-376: Office Open XML File Formats](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [Microsoft XLS Binary File Format (BIFF8)](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/)
- [Compound File Binary Format](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/)
- [Excel Function Reference](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188)
