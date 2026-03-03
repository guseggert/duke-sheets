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
- [x] Writer E2E tests — 33 double-roundtrip tests (fonts, fills, borders, alignment, number formats, dimensions, merged cells, conditional formatting, data validation, freeze panes + multi-selection)

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
- [x] Structured reference evaluation (`Table1[Column]`, `Table1[@Col]`, `[#Headers]`, `[#Totals]`, `[#All]`, `[#Data]`)
### Implemented Functions (506 of 506)

| Category | Count | Functions (highlights) |
|----------|-------|----------------------|
| Math & Trig | 78 | SUM, SUMIF, SUMIFS, AVERAGE, MIN, MAX, COUNT, COUNTIF, COUNTIFS, ROUND, ABS, MOD, INT, CEILING, FLOOR, POWER, SQRT, RAND, LOG, LN, PI, ACOSH, ASINH, ATANH, ACOTH, COSH, SINH, TANH, COT, COTH, CSC, CSCH, SEC, SECH, COMBIN, COMBINA, FACT, FACTDOUBLE, GCD, LCM, PRODUCT, QUOTIENT, MROUND, SUMSQ, SQRTPI, BASE, DECIMAL, ROMAN, ARABIC, MDETERM, MINVERSE, MMULT, MUNIT, AGGREGATE, SUBTOTAL, SERIESSUM, RANDARRAY, ... |
| Text | 43 | LEN, LEFT, RIGHT, MID, LOWER, UPPER, TRIM, CONCAT, CONCATENATE, FIND, SEARCH, SUBSTITUTE, TEXT, TEXTJOIN, FIXED, DOLLAR, NUMBERVALUE, REPLACE, REPLACEB, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, UNICHAR, UNICODE, VALUETOTEXT, ARRAYTOTEXT, ASC, BAHTTEXT, DBCS, JIS, PHONETIC, ... |
| Statistical | 110 | AVERAGEIF, AVERAGEIFS, COUNTBLANK, LARGE, SMALL, STDEV.S, STDEV.P, VAR.S, VAR.P, MAXIFS, MINIFS, RANK.EQ, RANK.AVG, PERCENTILE.INC, QUARTILE.INC, MODE.SNGL, BETA.DIST, BINOM.DIST, CHISQ.DIST, CORREL, COVARIANCE.P, COVARIANCE.S, EXPON.DIST, F.DIST, FISHER, FORECAST, GAMMA, GAMMA.DIST, HYPGEOM.DIST, NORM.DIST, NORM.S.DIST, PEARSON, POISSON.DIST, T.DIST, WEIBULL.DIST, Z.TEST, LOGNORM.DIST, LOGNORM.INV, LINEST, LOGEST, GROWTH, TREND, FORECAST.ETS, ... |
| Logical | 19 | IF, AND, OR, NOT, IFERROR, IFNA, IFS, SWITCH, XOR, TRUE, FALSE, LET, LAMBDA, MAP, REDUCE, SCAN, BYCOL, BYROW, MAKEARRAY |
| Lookup & Reference | 37 | INDEX, MATCH, VLOOKUP, HLOOKUP, XLOOKUP, XMATCH, CHOOSE, ROW, COLUMN, ROWS, COLUMNS, INDIRECT, OFFSET, ADDRESS, AREAS, FILTER, SORT, SORTBY, UNIQUE, TRANSPOSE, LOOKUP, HSTACK, VSTACK, TAKE, DROP, EXPAND, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW, WRAPCOLS, WRAPROWS, HYPERLINK, ... |
| Date | 25 | DATE, YEAR, MONTH, DAY, NOW, TODAY, TIME, HOUR, MINUTE, SECOND, WEEKDAY, WEEKNUM, ISOWEEKNUM, EDATE, EOMONTH, DAYS, DAYS360, DATEDIF, YEARFRAC, DATEVALUE, TIMEVALUE, NETWORKDAYS, WORKDAY, NETWORKDAYS.INTL, WORKDAY.INTL |
| Information | 21 | ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISNA, NA, ISERR, ISEVEN, ISODD, ISLOGICAL, ISNONTEXT, ISREF, ERROR.TYPE, TYPE, CELL, INFO, SHEET, SHEETS, ISFORMULA, ISOMITTED, STOCKHISTORY |
| Compatibility | 40 | BETADIST, BETAINV, BINOMDIST, CEILING, CHIDIST, CHIINV, CHITEST, CONFIDENCE, COVAR, CRITBINOM, EXPONDIST, FDIST, FINV, FLOOR, FTEST, GAMMADIST, GAMMAINV, HYPGEOMDIST, LOGINV, LOGNORMDIST, NEGBINOMDIST, NORM.INV, NORMDIST, NORMSDIST, NORMSINV, POISSON, TDIST, TINV, TTEST, WEIBULL, ZTEST, MODE, PERCENTILE, PERCENTRANK, QUARTILE, RANK, STDEV, STDEVP, VAR, VARP |
| Financial | 56 | PMT, FV, PV, NPER, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC, NPV, IRR, MIRR, XNPV, XIRR, SLN, SYD, DB, DDB, VDB, EFFECT, NOMINAL, PDURATION, ACCRINT, ACCRINTM, AMORDEGRC, AMORLINC, COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD, DISC, DOLLARDE, DOLLARFR, DURATION, EUROCONVERT, FVSCHEDULE, INTRATE, MDURATION, ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD, PRICE, PRICEDISC, PRICEMAT, RECEIVED, RRI, TBILLEQ, TBILLPRICE, TBILLYIELD, YIELD, YIELDDISC, YIELDMAT, ISPMT |
| Database | 12 | DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR, DVARP |
| Engineering | 54 | BESSELI, BESSELJ, BESSELK, BESSELY, BIN2DEC, BIN2HEX, BIN2OCT, BITAND, BITOR, BITXOR, COMPLEX, CONVERT, DEC2BIN, DEC2HEX, DEC2OCT, DELTA, ERF, ERFC, GESTEP, HEX2BIN, HEX2DEC, HEX2OCT, IMABS, IMAGINARY, IMCOS, IMDIV, IMEXP, IMLN, IMPOWER, IMPRODUCT, IMREAL, IMSIN, IMSQRT, IMSUB, IMSUM, OCT2BIN, OCT2DEC, OCT2HEX, ... |

### CLI Tool (`duke`)
- [x] `duke to-csv` — convert spreadsheet to CSV (with `-f` for formatted output)
- [x] `duke info` — show file information
- [x] `duke sheets` — list sheets in workbook
- [x] Formula calculation flag (`-c`)
- [x] Custom delimiter support

---

## In Progress / Partial

### Database Functions
- [x] Implemented all 12 D-functions in `crates/duke-sheets-formula/src/functions/database.rs` with shared criteria filtering and unit tests
- [x] Wire `database.rs` into `functions/mod.rs` registry (module + function registrations)

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
- [x] Dynamic array spilling in evaluator — SpillTarget value resolution, `#` (spill range) operator, `@` (implicit intersection) operator, two-pass calculation for spill-dependent formulas, #SPILL! error on blocked ranges
- [x] XLSX dynamic array metadata (`cm` attribute, `xl/metadata.xml`) — writer emits `cm="1"` on anchor cells, `cm="2"` on ghost cells, writes `xl/metadata.xml` with XLDAPR structure; reader parses `cm` attribute
- [x] Element-wise array binary operators — `=SEQUENCE(3)>1` produces `{FALSE,TRUE,TRUE}`, all arithmetic/comparison/concatenation operators lift to arrays with broadcasting

### Financial Functions
- [x] Core financial function implementations added in `crates/duke-sheets-formula/src/functions/financial.rs` (PMT, FV, PV, NPER, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC, NPV, IRR, MIRR, XNPV, SLN, SYD, DB, DDB, EFFECT, NOMINAL, PDURATION)
- [x] Wire `financial.rs` into `functions/mod.rs` registry entries (module + function registration)

---

## Not Started

### High Priority

#### XLSX Writer Gaps (remaining)
- [x] **Comment VML drawings** — writer now emits `vmlDrawing{N}.vml` note shapes + `legacyDrawing`/rels/content-types wiring so Excel displays comments
- [x] **`headerFooter` (odd only)** — `PageSetup` includes odd header/footer strings; XLSX reader/writer parse and emit `<headerFooter>`
- [x] **`headerFooter` (even/first + flags)** — even/first header/footer strings and flags (`differentOddEven`, `differentFirst`, `alignWithMargins`, `scaleWithDoc`) fully supported in data model, XLSX reader, and writer

#### XLSX Reader Gaps
- [x] **Theme colors** — reader now parses `xl/theme/theme1.xml` (`clrScheme`) and resolves theme+tint colors in styles/CF
- [ ] **Shared/array/dataTable formulas** — shared + array anchor parsed; `dataTable` now mapped to `=TABLE(r1,r2)` placeholder using OOXML attrs, but full behavior remains incomplete
- [x] **Theme/indexed colors in CF** — `parse_color_element()` now handles `rgb`/`theme`/`indexed`/`tint`/`auto` and resolves with workbook theme palette
- [x] **`cellStyleXfs` / named cell styles** — reader parses `cellStyleXfs` inheritance and `cellStyles`; writer preserves parsed `cellStyleXfs` + named styles on roundtrip (via roundtrip data registry)
- [x] **Font scheme/family/charset** — modeled in `FontStyle`, parsed from XLSX fonts, and emitted by writer when present
- [x] **Outline/grouping levels** — row/column `outlineLevel` + `collapsed` now read from XLSX into worksheet metadata
- [x] **Sheet views** — tab selection, zoom scale, active selection, frozen panes, and non-frozen split panes now roundtrip (single selection only)
- [x] **Sheet views (multi-selection)** — preserve multiple `<selection>` entries and multi-range `sqref` values; `Selection` struct with pane/active_cell/sqref, `Vec<Selection>` on Worksheet, backward-compatible convenience API
- [x] **Sheet views (pane edge cases)** — handle non-empty `<pane>` start/end tags via `parse_pane_attrs()` helper
- [x] **Comment visibility** — reader parses VML note shapes and sets `CellComment.visible` (currently uses style `visibility:visible`)
- [x] **Comments/VML relationships** — resolve `comments*.xml` and `vmlDrawing*.vml` via `sheet*.xml.rels` with fallback to index-based naming
- [x] **Comment visibility (robust)** — parse VML `<x:Visible/>` element and tolerate whitespace in style `visibility:visible` check
- [x] **Rich text in shared strings** — reader preserves `<rPr>` formatting runs as `CellValue::RichText`

#### Excel Function Coverage — Complete (506/506)

All 506 Excel functions are registered. 11 are stubs returning #N/A because they
require external runtime resources (OS DLL calls, OLAP servers, web data feeds,
or cell-level metadata not available in standalone evaluation):

| Category | Stub functions |
|----------|---------------|
| Add-in | CALL, REGISTER.ID (OS DLL calls) |
| Cube | CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE (OLAP server) |
| Information | STOCKHISTORY (web service) |
| Text | PHONETIC (cell `<rPh>` metadata) |

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
- [x] **Worksheet XML element ordering** — emit children in spec order (dimension, sheetFormatPr, printOptions before pageMargins); verified via positional assertions on raw XML
- [ ] **OOXML spec validation** — run Open XML SDK validator on duke-sheets output

#### General Quality
- [ ] **Refactor XLSX reader/writer modules** — break `crates/duke-sheets-xlsx/src/reader/mod.rs` and `crates/duke-sheets-xlsx/src/writer/mod.rs` into section parsers (sheet views, page setup, comments, CF/DV, etc.) + relationship-based part resolver
- [x] **XLSX reader modular split (phase 1)** — extracted `reader/theme.rs`, `reader/formulas.rs`, `reader/data_validation.rs`, `reader/conditional_format.rs`, `reader/comments.rs`; `reader/mod.rs` now keeps orchestration/workbook/sheet loop/cell processing
- [x] **XLSX reader/writer modular split (phase 2)** — extracted `writer/conditional_format.rs`, `writer/data_validation.rs`, `writer/comments.rs`, `writer/tables.rs`, and `reader/workbook.rs`; converted call sites to module-level free functions
- [ ] **Property-based testing** (proptest) — CellAddress roundtrip, Style write/read, formula parse/print
- [ ] **Broader locale coverage** — more built-in `Locale` constructors; CLI `--locale` flag; consider system locale auto-detection

#### CI
- [ ] `cargo test` on every push (GitHub Actions)
- [ ] Excel COM E2E on self-hosted runner with KVM
- [ ] Nightly job for slow tasks: fuzz corpus, benchmarks, real-world file corpus
- [ ] Clippy + `cargo fmt --check` gate

### Medium Priority

#### Hyperlinks
- [x] **Data model** — `Hyperlink` struct (URL, display text, tooltip, location)
- [x] **XLSX reader** — read `<hyperlinks>` + relationship targets
- [x] **XLSX writer** — write hyperlinks
- [x] **XLS reader** — parse `HLINK`/`HLINKTOOLTIP` records (URL, file, and internal monikers with tooltip support)

#### Rich Text Runs
- [x] **Data model** — `RichTextRun` (text + `RunFont` override), `CellValue::RichText` variant with 25+ match-site updates
- [x] **XLSX reader (SST)** — preserve `<rPr>` runs in shared strings via `SharedStringEntry::Rich`
- [x] **XLSX reader (inline)** — parse `<r>/<rPr>/<t>` inside `<is>` elements with whitespace preservation
- [x] **XLSX writer** — write `<is><r><rPr>...</rPr><t>text</t></r></is>` for rich text cells (inline strings)
- [x] **Roundtrip test** — mixed formatting (bold, italic, color), plain runs, multi-property runs
- [x] **Excel COM E2E tests** — writer roundtrip (bold/italic/color survives Excel re-save) + reader test (Characters API formatting parsed from SST)

#### Tables / ListObjects
- [x] **Data model** — `Table`, `TableColumn`, `TotalsRowFunction`, `TableStyleInfo` structs with header/totals row counts, calculated column formulas, totals row functions/labels/formulas, style info
- [x] **XLSX reader** — parse `xl/tables/tableN.xml` via sheet rels; handles all column attributes and child formula elements
- [x] **XLSX writer** — emit table parts (`xl/tables/tableN.xml`), `<tableParts>` in sheet XML, content types, and relationships; autoFilter ref correctly excludes totals row
- [x] **Structured reference evaluation** — resolve `Table1[Column]`, `Table1[@Column]`, `[#Headers]`, `[#Totals]`, `[#All]`, `[#Data]` specifiers in formula evaluator; dependency graph extraction for calculation engine

#### Auto-Filters
- [x] **Data model** — `AutoFilter` struct (range, column filters, sort state)
- [x] **XLSX reader** — read `<autoFilter>` from sheet XML
- [x] **XLSX writer** — write `<autoFilter>`
- [x] **XLS reader** — parse `AUTOFILTER`/`AUTOFILTERINFO` records and `_FilterDatabase` range

#### XLS Reader — Remaining Items
- [x] Formula Phase 3 (shared formulas, array formulas, memory tokens) — **done**
- [x] **tTbl** (data table formula indicator) — parser emits `ParsedToken::Table`, decompiler outputs `TABLE(<cell>)` placeholder
- [x] **Future function mapping** — unknown function indices with future-function bit now emit `_xlfn.FUNC<n>(...)` placeholders
- [x] **Cell comments** — `NOTE`/`OBJ`/`TXO` records parsed with object ID correlation and Latin-1/UTF-16LE text support
- [x] **Conditional formatting** — `CONDFMT`/`CF` records parsed (CellIs with all 8 operators, Expression rule type)
- [x] **Data validation** — `DVAL`/`DV` records parsed (all 8 validation types, all operators, input/error messages)
- [x] **Hyperlinks** — `HLINK`/`HLINKTOOLTIP` records parsed (URL moniker, file moniker, internal links, display text, tooltips)
- [x] **Freeze panes** — `WINDOW2`/`PANE` + `SELECTION` parsed into worksheet freeze pane and sheet-view selection state
- [x] **Default row/column dimensions** — `DEFCOLWIDTH`/`DEFAULTROWHEIGHT` parsed and applied to worksheet defaults
- [x] **Outline/grouping** — row/column outline level + collapsed flags extracted from ROW/COLINFO options
- [x] **Sheet tab colors** — `SHEETLAYOUT` record parsed for tab RGB color
- [ ] **Pattern fills (non-solid) E2E** — LO Calc doesn't support pattern fills; unit test only

#### Named Ranges I/O
- [x] **XLSX reader** — read `<definedNames>` from `workbook.xml`
- [x] **XLSX writer** — write `<definedNames>`
- [x] **XLS reader** — parse `NAME` record formula bodies (store raw token bytes)
- [x] **Print areas / titles** — read `_xlnm.Print_Area` and `_xlnm.Print_Titles` defined names

#### Large File Support
- [ ] Streaming XLSX reader (SAX-style, low memory)
- [ ] Streaming XLSX writer
- [ ] Progress callbacks
- [ ] Memory-optimized cell storage mode

### Low Priority

#### Theme Support
- [x] **Read `xl/theme/theme1.xml`** — parse `clrScheme` theme colors (slots used by style/CF color refs)
- [x] **Theme color resolution** — resolve `theme` + `tint` color references in styles/CF to RGB using workbook theme palette
- [x] **Write theme** — roundtrip preserves original `xl/theme/theme1.xml` bytes; new workbooks get default Office theme

#### Print Settings
- [x] **Data model** — `PageSetup` supports paper/orientation/scale/fit, margins, print options, print area, repeat rows/cols titles, odd/even/first header/footer strings, and `headerFooter` flags
- [x] **XLSX reader** — reads `<pageSetup>`, `<pageMargins>`, `<printOptions>`, `<headerFooter>` (odd + even + first + flags), and `<rowBreaks>/<colBreaks>`
- [x] **XLSX writer** — writes the same, including `<rowBreaks>/<colBreaks>` in spec order
- [ ] **XLS reader** — parse `SETUP`, `HEADER`, `FOOTER`, margin records, page break records

#### Sheet Views
- [x] **Data model** — zoom level, selection (single + multi via `Vec<Selection>`), freeze/split panes; still missing gridline visibility
- [x] **XLSX reader** — reads `<sheetViews>/<sheetView>` for zoom/selection/panes; multi-selection + multi-range sqref preserved
- [x] **XLSX writer** — writes `<sheetViews>` for zoom/selection/panes; emits all selections with pane inference
- [x] **XLS reader** — parse `WINDOW2`, `PANE`, `SELECTION`

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
| Core (cell, workbook, worksheet) | 41 | ✅ |
| Cell display formatting (CellView) | 51 | ✅ |
| Formula parser | 43 | ✅ |
| Formula evaluator + functions | 519 | ✅ |
| Calculation engine | 46 | ✅ |
| XLSX roundtrip | 58 | ✅ |
| XLSX style roundtrip | 10 | ✅ |
| XLSX escape decoding | 9 | ✅ |
| Formula E2E | 10 | ✅ |
| XLS unit (BIFF parser, strings, styles, formula decompiler, reader) | 115 | ✅ |
| XLS E2E (data types, styles, merged cells, dimensions, sheet props, formulas, comments, CF, DV) | 68 | ✅ |
| XLS real-file integration | 2 | ✅ |
| E2E XLSX reader integration (LO + handcrafted OOXML) | 63 | ✅ |
| E2E via Excel COM — reader (XLSX) | 62 | ✅ |
| E2E via Excel COM — writer (XLSX) | 35 | ✅ |
| XLSX formatting roundtrip | 17 | ✅ |
| Shared string reader | 9 | ✅ |
| Rich text unit tests | 4 | ✅ |
| Other (unit, doc, integration) | 261 | ✅ |
| **Total** | **1285** | ✅ |

---

## Known Issues

1. ~~**Formula parsing failures**~~ — Fixed. Quoted sheet refs, structured refs, external refs, @/# operators now parsed.
2. ~~**Structured refs / external refs not evaluated**~~ — Structured refs now fully evaluated (Table1[Col], @ThisRow, #Headers, #Totals, #All, #Data). External refs still return #REF! (external workbooks not loaded).
3. ~~**XLSX writer uses inline strings**~~ — Fixed. Now uses shared string table (SST).
4. ~~**Theme colors not resolved**~~ — Fixed. XLSX reader parses `xl/theme/theme1.xml` and resolves theme+tint colors in styles and conditional formatting.
5. ~~**XLS reader drops comments, hyperlinks, CF, DV**~~ — Fixed. XLS reader now parses NOTE/OBJ/TXO (comments), HLINK/HLINKTOOLTIP (hyperlinks), CONDFMT/CF (conditional formatting), and DVAL/DV (data validation)
6. ~~**Comment VML not written**~~ — Fixed. Writer now emits VML note shapes, worksheet `legacyDrawing`, and VML/comment relationships.

---

## Architecture Notes

### Crate Structure
```
duke-sheets/
├── duke-sheets-core        # Data model, cell storage, locale
├── duke-sheets-formula     # Parser, evaluator, 506 functions
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
- `CellValue` - Number, String, Boolean, Error, Formula, RichText, SpillTarget, Empty
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
