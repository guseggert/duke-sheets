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
- [x] Formula cached value preservation in reader (error, boolean, string, number)
- [x] Formula cached value preservation in writer (`<v>` element + `t` attribute)
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

### Excel COM Bridge (`duke-sheets-excel-com` + C# bridge server)
- [x] Generic COM proxy protocol (Get/Set/Invoke over NDJSON-over-TCP)
- [x] C# bridge server for Windows VM (`tools/excel-bridge-server/`)
- [x] Rust TCP client with typed Excel convenience methods
- [x] QEMU/KVM VM management scripts (`tools/vm/`)
- [x] Shared file access via QEMU SMB
- [x] Excel E2E smoke test (round-trip: numbers, strings, booleans, formulas)
- [ ] Excel parity E2E test suite (mirror remaining LibreOffice E2E tests)
- [ ] Style operations (font, fill, border, alignment via generic COM proxy)
- [ ] CI integration (self-hosted runner with KVM)

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
- [x] Quoted sheet references (`'Sheet Name'!A1`)
- [x] Structured references (`Table1[Column]`)
- [x] External workbook references (`[Book1.xlsx]Sheet1!A1`)
- [x] Implicit intersection operator (`@`)
- [x] Spill range operator (`#`)
- [x] Better error messages for unknown characters (no more "Unexpected token: Eof")
- [ ] Some complex formulas may still fail to parse (edge cases)

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
- [x] Investigate parse failures on real-world files (quoted sheet refs were the #1 cause)
- [x] Add support for implicit intersection (`@`)
- [x] Add support for spill operator (`#`)

### Medium Priority

#### XLS Reader (Legacy Excel) — Remaining Items
- [x] Compound File Binary (CFB) reader (via `cfb` crate)
- [x] BIFF8 record parsing with CONTINUE record merging and boundary tracking
- [x] Cell data: LABELSST, LABEL, NUMBER, RK, MULRK, BLANK, MULBLANK, BOOLERR, FORMULA
- [x] Shared String Table with CONTINUE encoding-change handling (Latin-1 ↔ UTF-16LE)
- [x] Styles: FONT, FORMAT, XF, PALETTE → full Style resolution (font, fill, border, alignment, number format, protection)
- [x] Structure: MERGECELLS, ROW, COLINFO, BOUNDSHEET, DATEMODE
- [x] Integrated into `WorkbookExt::open()` via `xls` feature gate
- [x] Formula token parsing Phase 1 — decompiles RPN tokens to infix text (~90% coverage: operators, constants, refs, functions)
- [x] Formula token parsing Phase 2 — 3D references (tRef3d/tArea3d), defined names (tName/tNameX), EXTERNSHEET/SUPBOOK/NAME record parsing (~98% coverage)
- [ ] Formula token parsing Phase 3 — shared formulas (tRefN/tAreaN), array constants (tArray), memory tokens
- [x] Style E2E tests: strikethrough, underline (single + double), text rotation, shrink-to-fit
- [x] Style E2E test: indent level
- [x] Sheet-level properties: hidden sheets (BOUNDSHEET visibility), active sheet (WINDOW1), sheet protection (PROTECT/PASSWORD)
- [x] Style E2E tests + LO bridge extension: diagonal borders, cell protection, reading order
- [ ] Pattern fills (non-solid) E2E — LO Calc cells don't support Excel-style pattern fills; reader + unit test coverage only

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

#### Performance Benchmarks
- [ ] Criterion benchmarks for XLSX read (small, medium, large files)
- [ ] Criterion benchmarks for XLSX write
- [ ] Criterion benchmarks for XLS read
- [ ] Formula parser benchmarks (throughput, complex expressions)
- [ ] Calculation engine benchmarks (large dependency graphs)
- [ ] Memory usage profiling / tracking for large workbooks
- [ ] Comparative benchmarks vs other libraries:
  - [ ] **calamine** (Rust, read-only) — XLSX/XLS read speed, memory usage
  - [ ] **umya-spreadsheet** (Rust, read/write) — XLSX read/write speed, style handling
  - [ ] **rust_xlsxwriter** (Rust, write-only) — XLSX write speed
  - [ ] **excelize** (Go) — XLSX read/write via CLI or FFI wrapper
  - [ ] **openpyxl** (Python) — XLSX read/write as baseline reference
  - [ ] Generate comparison tables/charts for README

#### Fuzz Testing
- [ ] Fuzz XLSX reader (`cargo-fuzz` / `libFuzzer`) — malformed ZIP, corrupt XML, truncated streams
- [ ] Fuzz XLS reader — malformed BIFF8 records, bad CONTINUE boundaries, corrupt SST
- [ ] Fuzz formula parser — arbitrary expression strings
- [ ] Fuzz CSV reader — malformed delimiters, encoding edge cases

#### C FFI
- [ ] Complete FFI bindings
- [ ] Python bindings via FFI
- [ ] Documentation

---

## Testing Status

| Test Suite | Count | Status |
|------------|-------|--------|
| Core (cell, workbook, worksheet) | 36 | ✅ |
| Formula parser | 37 | ✅ |
| Formula evaluator + functions | 74 | ✅ |
| Calculation engine | 8 | ✅ |
| XLSX roundtrip | 17 | ✅ |
| XLSX style roundtrip | 10 | ✅ |
| XLSX escape decoding | 9 | ✅ |
| Formula E2E | 10 | ✅ |
| XLS unit (BIFF parser, strings, styles, formula decompiler) | 87 | ✅ |
| XLS E2E (data types, styles, merged cells, dimensions, sheet props, formulas) | 56 | ✅ |
| XLS real-file integration | 2 | ✅ |
| E2E via LibreOffice URP (XLSX) | 56 | ✅ |
| E2E via Excel COM (XLSX) | 1 | ✅ |
| Other (unit, doc, integration) | 279 | ✅ |
| **Total** | **480** | ✅ |

---

## Known Issues

1. ~~**Formula parsing failures**~~ — Fixed. Quoted sheet refs, structured refs, external refs, @/# operators now parsed. Unknown chars give clear error messages.
2. **XLS formula text partial** - Phases 1+2 decompile ~98% of formulas (operators, constants, refs, functions, 3D refs, defined names); Phase 3 (shared formulas, arrays) pending
3. **Limited function coverage** - Only 35 of ~450 Excel functions implemented
4. **Structured refs / external refs not evaluated** — Parser handles them, but evaluator returns #NAME? / #REF! (tables and external workbooks not implemented)

---

## Architecture Notes

### Crate Structure
```
duke-sheets/
├── duke-sheets-core        # Data model, cell storage
├── duke-sheets-formula     # Parser, evaluator, functions
├── duke-sheets-xlsx        # XLSX read/write
├── duke-sheets-xls         # XLS reader (BIFF8, read-only)
├── duke-sheets-csv         # CSV read/write
├── duke-sheets-chart       # Chart support (stub)
├── duke-sheets-ffi         # C FFI bindings
├── duke-sheets-cli         # CLI tool
├── duke-sheets-libreoffice # LibreOffice URP bridge (E2E testing)
├── libreoffice-urp         # URP protocol implementation
├── excel-com-protocol      # Generic COM proxy protocol types
├── duke-sheets-excel-com   # Excel COM client (TCP to Windows VM)
├── duke-sheets             # Main crate, re-exports
└── tools/
    ├── excel-bridge-server # C# bridge server (runs in Windows VM)
    └── vm/                 # QEMU/KVM VM management scripts
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
