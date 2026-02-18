# Duke Sheets - Development Status & Roadmap

## Completed Features

### Core Infrastructure
- [x] Cell storage with sparse representation
- [x] Workbook/Worksheet data model
- [x] Cell addressing (A1 notation, R1C1)
- [x] Cell ranges and iteration
- [x] Style system (fonts, colors, borders, fills, alignment)
- [x] Number formatting
- [x] Merged cells support

### XLSX Support
- [x] Read XLSX files
- [x] Write XLSX files
- [x] Shared strings table
- [x] Style preservation on roundtrip
- [x] Excel `_xHHHH_` escape sequence decoding
- [x] Formula cached value preservation (error, boolean, string, number)
- [x] Data validation reading (list, whole, decimal, date, time, textLength, custom)
- [x] Conditional formatting reading (cellIs, expression, colorScale, dataBar, iconSet, etc.)
- [x] DXF (differential format) style reading for conditional formatting
- [x] Cell comment reading (text, author, rich text flattening)
- [x] Merged cells read/write
- [x] Font vertical align (superscript/subscript) read/write
- [x] Row heights / column widths / hidden rows & columns read/write
- [x] Gradient fills read/write (`<gradientFill>` with stops, linear & path types)

### LibreOffice URP Bridge (`duke-sheets-libreoffice`)
- [x] UNO Remote Protocol client over TCP
- [x] Binary protocol negotiation, type/OID/TID caches
- [x] Workbook create/open/save/close
- [x] Cell values, formulas, styles, comments
- [x] Merged cells, row height, column width
- [x] Number formats, conditional formatting, data validation
- [x] On-demand E2E test fixtures via global LO connection singleton

### CSV Support
- [x] Read CSV files
- [x] Write CSV files
- [x] Configurable delimiters

### Formula Engine
- [x] Formula parser (text → AST)
- [x] Expression evaluator
- [x] Dependency graph
- [x] **Calculation chain** (`workbook.calculate()`)
- [x] **Circular reference detection**
- [x] **Iterative calculation** for circular refs
- [x] **Volatile function support** (NOW, TODAY, RAND, RANDBETWEEN)
- [x] Cell reference resolution (single cells, ranges)
- [x] Cross-sheet references (`Sheet2!A1`)

### Implemented Functions (35 total)

| Category | Functions |
|----------|-----------|
| Math | SUM, AVERAGE, MIN, MAX, COUNT, RAND, RANDBETWEEN |
| Logical | IF, AND, OR, NOT |
| Text | LEN, LEFT, RIGHT, MID, LOWER, UPPER, TRIM, CONCAT, CONCATENATE |
| Date | DATE, YEAR, MONTH, DAY, NOW, TODAY |
| Lookup | INDEX, MATCH, VLOOKUP |
| Info | ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISNA, NA |

### CLI Tool (`duke`)
- [x] `duke to-csv` - Convert spreadsheet to CSV
- [x] `duke info` - Show file information
- [x] `duke sheets` - List sheets in workbook
- [x] Formula calculation flag (`-c`)
- [x] Custom delimiter support

---

## In Progress / Partial

### Formula Parser
- [ ] Some complex formulas fail to parse (seen "Unexpected token: Eof" errors)
- [ ] Structured references (`Table1[Column]`)
- [ ] External workbook references (`[Book1.xlsx]Sheet1!A1`)

### Array Formulas
- [x] Array literals (`{1,2,3}`)
- [ ] Array formula entry (`Ctrl+Shift+Enter` style)
- [ ] Dynamic array spilling

---

## Not Started

### High Priority

#### More Excel Functions (~415 remaining)
Common functions needed:
- [ ] **Math**: ROUND, ROUNDUP, ROUNDDOWN, ABS, SQRT, POWER, MOD, INT, CEILING, FLOOR, SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS
- [ ] **Logical**: IFERROR, IFNA, IFS, SWITCH, XOR
- [ ] **Text**: FIND, SEARCH, SUBSTITUTE, REPLACE, REPT, TEXT, VALUE, EXACT, CLEAN, CHAR, CODE, T, N
- [ ] **Lookup**: HLOOKUP, XLOOKUP, LOOKUP, CHOOSE, OFFSET, INDIRECT, ROW, COLUMN, ROWS, COLUMNS
- [ ] **Date/Time**: TIME, HOUR, MINUTE, SECOND, WEEKDAY, WEEKNUM, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY
- [ ] **Statistical**: STDEV, STDEVP, VAR, VARP, MEDIAN, MODE, LARGE, SMALL, RANK, PERCENTILE, QUARTILE
- [ ] **Financial**: PMT, FV, PV, NPV, IRR, RATE, NPER, SLN, DB, DDB

#### Formula Parser Fixes
- [ ] Investigate parse failures on real-world files
- [ ] Add support for implicit intersection (`@`)
- [ ] Add support for spill operator (`#`)

### Medium Priority

#### XLS Support (Legacy Excel)
- [ ] Compound File Binary (CFB) reader
- [ ] BIFF8 record parsing  
- [ ] XLS reader implementation
- [ ] XLS writer implementation

#### Large File Support
- [ ] Streaming XLSX reader (SAX-style, low memory)
- [ ] Streaming XLSX writer
- [ ] Progress callbacks
- [ ] Memory-optimized cell storage mode

### Low Priority

#### Charts
- [ ] Chart data model
- [ ] Read charts from XLSX
- [ ] Write charts to XLSX
- [ ] Basic chart types (bar, line, pie, scatter)

#### XLSX Reader Gaps
- [x] ~~**Read merged cells**~~ — done
- [x] ~~**Read row heights / column widths**~~ — done
- [x] ~~**Gradient fills**~~ — done (linear & path types with stops)
- [x] ~~**Font vertical align**~~ — done (superscript/subscript)
- [ ] **Theme/indexed colors in CF** — conditional format color elements only handle `rgb`, not `theme`/`indexed`/`tint`
- [ ] **Comment visibility** — model has `visible` field, reader doesn't parse VML drawings (large effort)

#### Advanced Features
- [ ] Pivot tables (read-only)
- [ ] Hyperlinks
- [ ] Images
- [ ] Print settings

#### C FFI
- [ ] Complete FFI bindings
- [ ] Python bindings via FFI
- [ ] Documentation

---

## Testing Status

| Test Suite | Count | Status |
|------------|-------|--------|
| Core (cell, workbook, worksheet) | 36 | ✅ |
| Formula parser | 17 | ✅ |
| Formula evaluator | 24 | ✅ |
| Calculation engine | 8 | ✅ |
| XLSX roundtrip | 12 | ✅ |
| XLSX style roundtrip | 10 | ✅ |
| XLSX escape decoding | 9 | ✅ |
| Formula E2E | 10 | ✅ |
| E2E via LibreOffice URP | 56 | ✅ |
| Other (unit, doc, integration) | 155 | ✅ |
| **Total** | **337** | ✅ |

---

## Known Issues

1. **Formula parsing failures** - Some complex real-world formulas fail with "Unexpected token: Eof"
2. **XLS not supported** - `.xls` files cannot be read (stub only)
3. **Limited function coverage** - Only 35 of ~450 Excel functions implemented

---

## Architecture Notes

### Crate Structure
```
duke-sheets/
├── duke-sheets-core      # Data model, cell storage
├── duke-sheets-formula   # Parser, evaluator, functions
├── duke-sheets-xlsx      # XLSX read/write
├── duke-sheets-xls       # XLS read/write (stub)
├── duke-sheets-csv       # CSV read/write
├── duke-sheets-chart     # Chart support (stub)
├── duke-sheets-ffi       # C FFI bindings
├── duke-sheets-cli       # CLI tool
└── duke-sheets           # Main crate, re-exports
```

### Key Types
- `Workbook` - Container for worksheets
- `Worksheet` - Grid of cells with metadata
- `CellValue` - Number, String, Boolean, Error, Formula
- `FormulaExpr` - AST for parsed formulas
- `DependencyGraph` - Tracks cell dependencies
- `CalculationEngine` - Evaluates formulas in order

---

## Quick Reference

### Build & Test
```bash
cargo build                    # Build all
cargo test                     # Run all tests
cargo build -p duke-sheets-cli # Build CLI only
```

### CLI Usage
```bash
duke to-csv input.xlsx              # Convert to CSV (stdout)
duke to-csv -c input.xlsx           # Calculate formulas first
duke to-csv -o out.csv input.xlsx   # Output to file
duke info input.xlsx                # Show file info
duke sheets input.xlsx              # List sheets
```

### Library Usage
```rust
use duke_sheets::prelude::*;
use duke_sheets::WorkbookCalculationExt;

let mut wb = Workbook::open("input.xlsx")?;
wb.calculate()?;  // Evaluate all formulas
wb.save("output.xlsx")?;
```
