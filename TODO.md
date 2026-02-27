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

### XLSX Reader
- [x] Read XLSX files (cell data, styles, shared strings)
- [x] Shared strings table (read)
- [x] Style preservation on roundtrip
- [x] Excel `_xHHHH_` escape sequence decoding
- [x] Formula cached value preservation (error, boolean, string, number)
- [x] Data validation reading (list, whole, decimal, date, time, textLength, custom)
- [x] Conditional formatting reading (cellIs, expression, colorScale, dataBar, iconSet, etc.)
- [x] DXF (differential format) style reading for conditional formatting
- [x] Cell comment reading (text, author, rich text flattening)
- [x] Merged cells
- [x] Font vertical align (superscript/subscript)
- [x] Row heights / column widths / hidden rows & columns
- [x] Gradient fills (`<gradientFill>` with stops, linear & path types)
- [x] `workbookPr` date1904 reading
- [x] Named ranges (`definedNames` from `workbook.xml`)
- [x] XML namespace handling (`local_name()` across 58 call sites)
- [x] Graceful error recovery (bad SST ref → `#REF!`, bad style → default, bad cell ref → skip)

### XLSX Writer
- [x] Write XLSX files (cell data, styles)
- [x] Formula cached value preservation (`<v>` element + `t` attribute)
- [x] Merged cells
- [x] Font vertical align (superscript/subscript)
- [x] Row heights / column widths / hidden rows & columns
- [x] Gradient fills
- [x] Shared string table (SST) — deduplicated across all sheets
- [x] `workbookPr` (date_1904 preservation on roundtrip)
- [x] `calcPr` (tells Excel to recalculate formulas on open)
- [x] `bookViews` (active sheet index)
- [x] Sheet visibility (`state="hidden"`)
- [x] Tab color (`sheetPr/tabColor`)
- [x] Freeze panes (`sheetViews/pane`)
- [x] Sheet protection (`sheetProtection` with password hash and permissions)
- [x] Page setup (`pageSetup` + `pageMargins`)
- [x] Named ranges (`definedNames` read + write, workbook and sheet scope)
- [x] Conditional formatting writing (all rule types with DXF styles)
- [x] Data validation writing (all types)
- [x] Refactored from string munging to `quick-xml` Writer API (writer/mod.rs + styles.rs)

### XLS Reader (Legacy BIFF8)
- [x] Compound File Binary (CFB) reader (via `cfb` crate)
- [x] BIFF8 record parsing with CONTINUE record merging and boundary tracking
- [x] Cell data: LABELSST, LABEL, NUMBER, RK, MULRK, BLANK, MULBLANK, BOOLERR, FORMULA, STRING
- [x] Shared String Table with CONTINUE encoding-change handling (Latin-1 ↔ UTF-16LE)
- [x] Styles: FONT, FORMAT, XF, PALETTE → full Style resolution
- [x] Structure: MERGECELLS, ROW, COLINFO, BOUNDSHEET, DATEMODE
- [x] Integrated into `WorkbookExt::open()` via `xls` feature gate
- [x] Formula token parsing Phase 1 — RPN tokens to infix (~90% coverage)
- [x] Formula token parsing Phase 2 — 3D refs, defined names, EXTERNSHEET/SUPBOOK/NAME (~98%)
- [x] Formula token parsing Phase 3 — shared formulas (tRefN/tAreaN/SHAREDFMLA), array formulas (tArray/ARRAY), memory tokens (tMemFunc/tMemArea/tMemErr/tMemNoMem), tExp resolution
- [x] Sheet-level properties: hidden sheets, active sheet (WINDOW1), sheet protection (PROTECT/PASSWORD)
- [x] Style E2E tests: strikethrough, underline, text rotation, shrink-to-fit, indent, diagonal borders, cell protection, reading order

### Cell Display Formatting
- [x] Number format rendering engine (`CellView::formatted()` via ssfmt)
- [x] Date serial number formatting (1900 and 1904 date systems)
- [x] `format_cell_value()` + `Worksheet::formatted_value_at()` convenience methods
- [x] CLI `--formatted` flag (`duke to-csv -f`)
- [x] Locale support (`Locale` type on Worksheet, built-in: en_us, de_de, fr_fr, en_gb, ja_jp)

### LibreOffice URP Bridge
- [x] UNO Remote Protocol client over TCP
- [x] Binary protocol negotiation, type/OID/TID caches
- [x] Workbook create/open/save/close
- [x] Cell values, formulas, styles, comments
- [x] Merged cells, row height, column width
- [x] Number formats, conditional formatting, data validation
- [x] On-demand E2E test fixtures via global LO connection singleton

### Excel COM Bridge
- [x] Generic COM proxy protocol (Get/Set/Invoke over NDJSON-over-TCP)
- [x] C# bridge server for Windows VM (`tools/excel-bridge-server/`)
- [x] Rust TCP client with typed Excel convenience methods
- [x] QEMU/KVM VM management scripts (`tools/vm/`)
- [x] Shared file access via QEMU SMB
- [x] Excel parity E2E test suite — Phase 1-5 complete (59 reader tests)
- [x] Writer E2E tests — 31 double-roundtrip tests (fonts, fills, borders, alignment, number formats, dimensions, merged cells, conditional formatting, data validation)

### CSV Support
- [x] Read CSV files
- [x] Write CSV files
- [x] Configurable delimiters

### Formula Engine
- [x] Formula parser (text → AST)
- [x] Expression evaluator
- [x] Dependency graph
- [x] Calculation chain (`workbook.calculate()`)
- [x] Circular reference detection
- [x] Iterative calculation for circular refs
- [x] Volatile function support (NOW, TODAY, RAND, RANDBETWEEN)
- [x] Cell reference resolution (single cells, ranges)
- [x] Cross-sheet references (`Sheet2!A1`)

### Implemented Functions (108 of ~506)

| Category | Count | Functions (highlights) |
|----------|-------|----------------------|
| Math & Trig | 35 | SUM, SUMIF, SUMIFS, AVERAGE, MIN, MAX, COUNT, COUNTIF, COUNTIFS, ROUND, ABS, MOD, INT, CEILING, FLOOR, POWER, SQRT, RAND, LOG, LN, PI, ... |
| Text | 29 | LEN, LEFT, RIGHT, MID, LOWER, UPPER, TRIM, CONCAT, CONCATENATE, FIND, SEARCH, SUBSTITUTE, REPLACE, REPT, EXACT, CLEAN, CHAR, CODE, T, N, VALUE, ... |
| Statistical | 13 | AVERAGEIF, AVERAGEIFS, COUNTBLANK, LARGE, SMALL, ... |
| Logical | 11 | IF, AND, OR, NOT, IFERROR, IFNA, IFS, SWITCH, XOR, TRUE, FALSE |
| Lookup | 8 | INDEX, MATCH, VLOOKUP, CHOOSE, ROW, COLUMN, ROWS, COLUMNS |
| Date | 6 | DATE, YEAR, MONTH, DAY, NOW, TODAY |
| Information | 6 | ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISNA, NA |

### CLI Tool (`duke`)
- [x] `duke to-csv` — convert spreadsheet to CSV (with `-f` for formatted output)
- [x] `duke info` — show file information
- [x] `duke sheets` — list sheets in workbook
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
- [x] Better error messages for unknown characters
- [ ] Some complex formulas may still fail to parse (edge cases)

### Array Formulas
- [x] Array literals (`{1,2,3}`)
- [ ] Array formula entry (`Ctrl+Shift+Enter` style) in XLSX reader
- [ ] Dynamic array spilling in evaluator

---

## Not Started

### High Priority

#### XLSX Writer Gaps (remaining)
- [ ] **Comment VML drawings** — comments written to `comments{N}.xml` but without VML positioning; Excel may not display them
- [x] **`headerFooter`** — `PageSetup` now includes odd header/footer strings; XLSX reader/writer parse and emit `<headerFooter>`

#### XLSX Reader Gaps
- [x] **Theme colors** — reader now parses `xl/theme/theme1.xml` (`clrScheme`) and resolves theme+tint colors in styles/CF
- [ ] **Shared/array/dataTable formulas** — shared + array anchor parsed; `dataTable` now mapped to `=TABLE(r1,r2)` placeholder using OOXML attrs, but full behavior remains incomplete
- [x] **Theme/indexed colors in CF** — `parse_color_element()` now handles `rgb`/`theme`/`indexed`/`tint`/`auto` and resolves with workbook theme palette
- [ ] **`cellStyleXfs` / named cell styles** — reader now parses `cellStyleXfs` inheritance (`xfId` + apply flags); writer still hardcodes one entry + "Normal"
- [x] **Font scheme/family/charset** — modeled in `FontStyle`, parsed from XLSX fonts, and emitted by writer when present
- [x] **Outline/grouping levels** — row/column `outlineLevel` + `collapsed` now read from XLSX into worksheet metadata
- [x] **Sheet views** — tab selection, zoom scale, active selection, frozen panes, and non-frozen split panes now roundtrip
- [ ] **Comment visibility** — model has `visible` field, reader doesn't parse VML drawings
- [ ] **Rich text in shared strings** — reader flattens `<rPr>` formatting runs to plain text

#### More Excel Functions (~398 remaining)
See `FUNCTIONS.md` for the complete tracking list. High-priority gaps:

| Category | Implemented | Total | Key missing functions |
|----------|------------|-------|----------------------|
| Financial | 0 | 55 | PMT, FV, PV, NPV, IRR, RATE, NPER, SLN |
| Engineering | 0 | 54 | BIN2DEC, DEC2BIN, HEX2DEC, CONVERT, COMPLEX |
| Database | 0 | 12 | DSUM, DCOUNT, DAVERAGE, DGET |
| Compatibility | 0 | 40 | STDEV, VAR, MODE, PERCENTILE, RANK, CEILING, FLOOR |
| Statistical | 13 / 110 | 12% | STDEV.S, STDEV.P, VAR.S, VAR.P, MEDIAN, MODE.SNGL, MAXIFS, MINIFS, RANK.EQ, PERCENTILE.INC |
| Date & Time | 6 / 25 | 24% | TIME, HOUR, MINUTE, SECOND, WEEKDAY, EDATE, EOMONTH, DATEDIF, NETWORKDAYS |
| Lookup | 8 / 34 | 24% | HLOOKUP, XLOOKUP, XMATCH, INDIRECT, OFFSET, FILTER, SORT, UNIQUE |
| Text | 29 / 42 | 69% | TEXT, TEXTJOIN, FIXED, DOLLAR |
| Logical | 11 / 19 | 58% | LET, LAMBDA |

#### Reader Robustness
- [x] Fix XML namespace handling (58 call sites)
- [x] Graceful error recovery (8 hard-fail sites converted)
- [ ] **Real-world file corpus + differential testing** — generate XLSX from multiple tools (openpyxl, xlsxwriter, Apache POI, LibreOffice), export from Google Sheets and Apple Numbers; compare against calamine

#### Writer Correctness
- [x] Roundtrip fidelity tests via Excel COM bridge (31 writer E2E tests)
- [x] Cross-app write compatibility (no Excel repair warnings)
- [x] Preserve theme/indexed color attributes in sheet XML (`tabColor`, CF `colorScale`, `dataBar`)
- [x] Emit `xl/theme/theme1.xml` + workbook theme relationship on write (default Office theme)
- [x] Emit row/column `outlineLevel` + `collapsed` attributes from worksheet metadata
- [ ] **OOXML spec validation** — run Open XML SDK validator on duke-sheets output

#### General Quality
- [ ] **Property-based testing** (proptest) — CellAddress roundtrip, Style write/read, formula parse/print
- [ ] **Broader locale coverage** — more built-in `Locale` constructors; CLI `--locale` flag; consider system locale auto-detection

#### CI
- [ ] `cargo test` on every push (GitHub Actions)
- [ ] Excel COM E2E on self-hosted runner with KVM
- [ ] Nightly job for slow tasks: fuzz corpus, benchmarks, real-world file corpus
- [ ] Clippy + `cargo fmt --check` gate

### Medium Priority

#### Hyperlinks
- [ ] **Data model** — `Hyperlink` struct (URL, display text, tooltip, location)
- [ ] **XLSX reader** — read `<hyperlinks>` + relationship targets
- [ ] **XLSX writer** — write hyperlinks
- [ ] **XLS reader** — parse `HLINK` record (0x01B8, constant already defined)

#### Rich Text Runs
- [ ] **Data model** — `RichTextRun` (text segment + font override) in `CellValue::String`
- [ ] **XLSX reader** — preserve `<rPr>` runs in shared strings and inline strings
- [ ] **XLSX writer** — write `<r>/<rPr>/<t>` runs for rich-text cells

#### Tables / ListObjects
- [ ] **Data model** — `Table` struct (name, range, columns, totals row, style)
- [ ] **XLSX reader** — read `<tableParts>` + `xl/tables/table{N}.xml`
- [ ] **XLSX writer** — write table definitions
- [ ] **Structured reference evaluation** — resolve `Table1[Column]` refs in formula evaluator

#### Auto-Filters
- [ ] **Data model** — `AutoFilter` struct (range, column filters, sort state)
- [ ] **XLSX reader** — read `<autoFilter>` from sheet XML
- [ ] **XLSX writer** — write `<autoFilter>`
- [ ] **XLS reader** — parse `AUTOFILTER` record (0x009E)

#### XLS Reader — Remaining Items
- [x] Formula Phase 3 (shared formulas, array formulas, memory tokens) — **done**
- [ ] **tTbl** (data table formula indicator) — parsed but emits `Unknown`
- [ ] **Future function mapping** — index >= 0x8000 not mapped to `_xlfn.NAME` format
- [ ] **Cell comments** — `NOTE` record (0x001C) not parsed
- [ ] **Conditional formatting** — `CONDFMT`/`CF` records not parsed (XLSX reader supports this)
- [ ] **Data validation** — `DVAL`/`DV` records not parsed (XLSX reader supports this)
- [ ] **Hyperlinks** — `HLINK` record defined but not handled
- [ ] **Freeze panes** — `WINDOW2`/`PANE` records defined but not parsed
- [ ] **Default row/column dimensions** — `DEFCOLWIDTH`/`DEFAULTROWHEIGHT` not parsed
- [ ] **Outline/grouping** — outline levels from ROW/COLINFO not extracted
- [ ] **Sheet tab colors** — `SHEETLAYOUT` record not parsed
- [ ] **Pattern fills (non-solid) E2E** — LO Calc doesn't support pattern fills; unit test only

#### Named Ranges I/O
- [x] **XLSX reader** — read `<definedNames>` from `workbook.xml`
- [x] **XLSX writer** — write `<definedNames>`
- [ ] **XLS reader** — parse `NAME` record formula bodies (names parsed but definitions not stored)
- [ ] **Print areas / titles** — read `_xlnm.Print_Area` and `_xlnm.Print_Titles` defined names

#### Large File Support
- [ ] Streaming XLSX reader (SAX-style, low memory)
- [ ] Streaming XLSX writer
- [ ] Progress callbacks
- [ ] Memory-optimized cell storage mode

### Low Priority

#### Theme Support
- [x] **Read `xl/theme/theme1.xml`** — parse `clrScheme` theme colors (slots used by style/CF color refs)
- [x] **Theme color resolution** — resolve `theme` + `tint` color references in styles/CF to RGB using workbook theme palette
- [ ] **Write theme** — preserve or generate theme on roundtrip

#### Print Settings
- [ ] **Data model** — expand `PageSetup` (paper size, orientation, margins, header/footer, print area, page breaks)
- [ ] **XLSX reader** — read `<pageSetup>`, `<pageMargins>`, `<headerFooter>`, `<rowBreaks>`, `<colBreaks>`
- [ ] **XLSX writer** — write print settings
- [ ] **XLS reader** — parse `SETUP`, `HEADER`, `FOOTER`, margin records, page break records

#### Sheet Views
- [ ] **Data model** — zoom level, selected cell, pane state, gridline visibility
- [ ] **XLSX reader** — read `<sheetViews>/<sheetView>`
- [ ] **XLSX writer** — write `<sheetViews>`
- [ ] **XLS reader** — parse `WINDOW2`, `PANE`, `SELECTION`

#### Charts
- [ ] Chart data model
- [ ] Read charts from XLSX
- [ ] Write charts to XLSX
- [ ] Basic chart types (bar, line, pie, scatter)

#### Images / Drawings
- [ ] Data model for embedded images
- [ ] Read `<drawing>` relationships + `xl/drawings/`
- [ ] Write images
- [ ] Two-cell anchor positioning

#### Sparklines
- [ ] Data model
- [ ] XLSX reader (SparklineGroups extension)

#### Pivot Tables
- [ ] Read-only data model
- [ ] XLSX reader

#### Advanced Style Features
- [ ] **Strikethrough type** — model only has boolean; Excel has single/double
- [ ] **Font condense/extend** — not modeled
- [ ] **Gradient fill path attributes** — `left`/`right`/`top`/`bottom` for path gradients not read

#### Performance Benchmarks
- [ ] Criterion benchmarks for XLSX read (small, medium, large files)
- [ ] Criterion benchmarks for XLSX write
- [ ] Criterion benchmarks for XLS read
- [ ] Formula parser benchmarks (throughput, complex expressions)
- [ ] Calculation engine benchmarks (large dependency graphs)
- [ ] Memory usage profiling / tracking for large workbooks
- [ ] Comparative benchmarks vs calamine, umya-spreadsheet, rust_xlsxwriter, excelize, openpyxl

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
| Cell display formatting (CellView) | 51 | ✅ |
| Formula parser | 37 | ✅ |
| Formula evaluator + functions | 74 | ✅ |
| Calculation engine | 8 | ✅ |
| XLSX roundtrip | 18 | ✅ |
| XLSX style roundtrip | 10 | ✅ |
| XLSX escape decoding | 9 | ✅ |
| Formula E2E | 10 | ✅ |
| XLS unit (BIFF parser, strings, styles, formula decompiler) | 87 | ✅ |
| XLS E2E (data types, styles, merged cells, dimensions, sheet props, formulas) | 56 | ✅ |
| XLS real-file integration | 2 | ✅ |
| E2E XLSX reader integration (LO + handcrafted OOXML) | 59 | ✅ |
| E2E via Excel COM — reader (XLSX) | 59 | ✅ |
| E2E via Excel COM — writer (XLSX) | 31 | ✅ |
| XLSX formatting roundtrip | 16 | ✅ |
| Other (unit, doc, integration) | 289 | ✅ |
| **Total** | **711** | ✅ |

---

## Known Issues

1. ~~**Formula parsing failures**~~ — Fixed. Quoted sheet refs, structured refs, external refs, @/# operators now parsed.
2. **Structured refs / external refs not evaluated** — Parser handles them, but evaluator returns #NAME? / #REF! (tables and external workbooks not implemented)
3. ~~**XLSX writer uses inline strings**~~ — Fixed. Now uses shared string table (SST).
4. ~~**Theme colors not resolved**~~ — Fixed. XLSX reader parses `xl/theme/theme1.xml` and resolves theme+tint colors in styles and conditional formatting.
5. **XLS reader drops comments, hyperlinks, CF, DV** — these features are supported by the XLSX reader but silently skipped in XLS
6. **Comment VML not written** — comments XML is written but VML positioning shapes are not; some Excel builds may not display them

---

## Architecture Notes

### Crate Structure
```
duke-sheets/
├── duke-sheets-core        # Data model, cell storage, locale
├── duke-sheets-formula     # Parser, evaluator, 108 functions
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
- `Worksheet` - Grid of cells with metadata, locale, date system
- `CellValue` - Number, String, Boolean, Error, Formula, SpillTarget, Empty
- `CellView` - Lightweight borrow wrapper with `formatted()` display
- `Locale` - Formatting locale (decimal separators, month names, currency)
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
duke to-csv input.xlsx              # Convert to CSV (stdout, raw values)
duke to-csv -f input.xlsx           # Apply Excel number formats
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
