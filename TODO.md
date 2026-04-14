# TODO

## In Progress

- [x] Reject XLSB files with a clear error instead of misreading them as XLSX.
  - [x] Add `zip` as a regular dependency in `crates/duke-sheets/Cargo.toml`.
  - [x] Add `FileFormat::Xlsb` variant to `detect_format`.
  - [x] Peek inside ZIP for `xl/workbook.bin` to distinguish XLSB from XLSX.
  - [x] Return clear error from `from_bytes` and `open` for XLSB files.
  - [x] Add 4 unit tests (detect xlsb, detect xlsx not misidentified, from_bytes rejects, open rejects).
  - [x] `cargo test -p duke-sheets --lib` passed (62 tests).
  - [x] `cargo check --manifest-path bindings/nodejs/Cargo.toml` passed.

- [x] Extract shared formula decompiler into duke-sheets-formula.
  - [x] Create `duke-sheets-formula/src/decompile/` with 4 modules: mod.rs, parsed_token.rs, decompiler.rs, function_table.rs.
  - [x] Move FormulaContext, SupBook, ExternSheetEntry, NameRecord, BUILTIN_NAMES, ParsedToken, decompiler stack machine, 485-entry function table.
  - [x] Widen row fields from u16 to u32 (BIFF12 supports 1M rows).
  - [x] Update duke-sheets-xls to import shared types, remove duplicate code.
  - [x] XLS token_parser.rs promotes u16 rows to u32 at ParsedToken construction sites.
  - [x] `cargo test -p duke-sheets-formula` passed (1123 tests).
  - [x] `cargo test -p duke-sheets-xls` passed (68 tests).
  - [x] `cargo test -p duke-sheets --features full --lib` passed (62 tests).
  - [x] `lsp_diagnostics` clean on all changed files.

- [x] XLSB read + write support (duke-sheets-xlsb crate).
  See `.sisyphus/plans/xlsb-support.md` for full plan.
  - [x] Phase 1: Create duke-sheets-xlsb crate with BIFF12 parsing infrastructure.
    - [x] Crate scaffold (Cargo.toml, workspace integration, lib.rs).
    - [x] Error types (XlsbError, XlsbResult).
    - [x] BIFF12 record type constants (records.rs).
    - [x] Binary parsing helpers: read_u16/u32/i32/f64, decode_rk, wide_str, cell_style_ref.
    - [x] RecordIter: variable-length type/size decoding, next_record, skip_to with block skipping.
    - [x] 29 unit tests covering all parsing primitives.
    - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (29 tests).
    - [x] `lsp_diagnostics` clean on all source files.
  - [x] Phase 3: Read styles from xl/styles.bin and apply to cells via iStyleRef.
    - [x] Add style record constants (BRT_FONT, BRT_FILL, BRT_BORDER, BRT_STYLE, Begin/End blocks).
    - [x] Create reader/styles.rs: parse BrtFmt, BrtFont, BrtFill, BrtBorder, BrtXF records.
    - [x] Resolve cellXfs via component tables (fonts, fills, borders, numfmts).
    - [x] Merge cellXf with cellStyleXf base using apply flags.
    - [x] Apply styles to cells via 24-bit iStyleRef in all cell record arms.
    - [x] 4 new style tests (number format, built-in format, bold font, solid fill).
    - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (44 tests).
    - [x] `cargo test -p duke-sheets --features full --lib` passed (62 tests).
    - [x] `lsp_diagnostics` clean on all changed files.
  - [x] Phase 5: Worksheet features — merged cells, hyperlinks, row/column dimensions, freeze panes, page setup, autofilter.
    - [x] Add record constants: BRT_COL_INFO, BRT_PANE, BRT_SEL, BRT_BEGIN_A_FILTER, BRT_END_A_FILTER, BRT_MARGINS, BRT_PAGE_SETUP, BRT_PRINT_OPTIONS, BRT_HEADER_FOOTER, BRT_BEGIN_SHEET_VIEW, BRT_END_SHEET_VIEW.
    - [x] Restructure worksheet reader into 3 phases: pre-data (views, col info), sheet data (cells), post-data (merge cells, hyperlinks, page setup, autofilter).
    - [x] Parse BrtColInfo for column widths and hidden columns.
    - [x] Parse BrtRowHdr for custom row heights and hidden rows.
    - [x] Parse BrtMergeCell for merged cell regions.
    - [x] Parse BrtHLink + per-sheet .rels for external/internal hyperlinks.
    - [x] Parse BrtPane for freeze panes (frozen/frozenSplit states).
    - [x] Parse BrtBeginAFilter for autofilter range.
    - [x] Parse BrtMargins, BrtPageSetup, BrtPrintOptions, BrtHeaderFooter for page setup.
    - [x] Add read_sheet_rels() in reader/mod.rs for per-sheet relationship files.
    - [x] 12 new tests: merged cells, custom row heights, hidden rows, column widths, freeze panes, autofilter, page margins, page setup landscape, header/footer, external hyperlinks, internal hyperlinks, combined features.
    - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (82 tests).
    - [x] `cargo test -p duke-sheets --features full --lib` passed (62 tests).
    - [x] `lsp_diagnostics` clean on all changed files.
  - [x] Phase 6: Advanced read features — comments, data validation, tables, conditional formatting.
    - [x] Expand read_sheet_rels to capture all relationship types (comments, tables, drawings).
    - [x] Add resolve_rel_path for resolving relative paths from .rels entries.
    - [x] Add BIFF12 record constants for comments, data validation, conditional formatting, tables.
    - [x] Implement comments reader (reader/comments.rs): binary BrtComment* records + XML fallback.
    - [x] Implement table reader (reader/table.rs): BrtBeginList, BrtBeginListCol, BrtTableStyleClient.
    - [x] Implement data validation parsing (BrtDVal): all 8 validation types, operators, messages, sqref ranges.
    - [x] Implement conditional formatting parsing (BrtBeginCFRule): cellIs, expression, top10, aboveAverage, text rules, duplicates, blanks, errors.
    - [x] Wire up comments, tables via sheet .rels in mod.rs read loop.
    - [x] 6 new tests: comments binary, comments XML fallback, data validation list, conditional format cellIs, table from binary, no-comments graceful.
    - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (88 tests).
    - [x] `cargo test -p duke-sheets --features full --lib` passed (62 tests).
    - [x] `lsp_diagnostics` clean on all changed files.
  - [x] Phase 7: Core writer — produce valid XLSB ZIP files from a Workbook.
    - [x] Add RecordWriter utility and encode_wide_str to biff12/mod.rs (made encode_type/encode_len/build_record non-test-only).
    - [x] Create writer/mod.rs: XlsbWriter orchestrator with ZIP entry creation, content types, .rels files.
    - [x] Create writer/shared_strings.rs: SstMap builder + BrtBeginSst/BrtSSTItem writer.
    - [x] Create writer/workbook.rs: BrtWbProp (date1904) + BrtBundleSh (per-sheet) writer.
    - [x] Create writer/styles.rs: minimal default styles (1 font, 2 fills, 1 border, 1 cellStyleXf, 1 cellXf, 1 style).
    - [x] Create writer/worksheet.rs: BrtRowHdr + cell records (BrtCellReal, BrtCellIsst, BrtCellSt, BrtCellBool, BrtCellError).
    - [x] Wire XlsbWriter into duke-sheets main crate save() dispatch for .xlsb extension.
    - [x] Export XlsbWriter from duke-sheets-xlsb and duke-sheets crates.
    - [x] 8 round-trip tests: empty workbook, strings, numbers, booleans, errors, multiple sheets, mixed types, unicode.
    - [x] 1 file-based round-trip test (write_file + read_file).
    - [x] 1 save dispatch test in main crate (save_xlsb_roundtrip via WorkbookExt::save).
    - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (97 tests).
    - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).
    - [x] `lsp_diagnostics` clean on all changed files.
  - [x] Phase 8: Full style writing and formula writing in the XLSB writer.
    - [x] Rewrite writer/styles.rs: collect styles from all worksheets, build component tables (fonts, fills, borders, numfmts), write BrtFmt/BrtFont/BrtFill/BrtBorder/BrtXF records, return StyleMapping.
    - [x] Update writer/worksheet.rs: use StyleMapping for correct 24-bit iStyleRef in cell prefixes, write formula cells as BrtFmla* records with cached values and empty token stream.
    - [x] Update writer/mod.rs: pass Workbook to write_styles, thread StyleMapping to write_worksheet.
    - [x] 12 new round-trip tests: bold+font size+color, custom number format, builtin number format, fill color, border style, formula cached number/string/bool/error, multiple styles across cells, style deduplication across sheets, formula with style.
    - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (109 tests).
    - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).
    - [x] `lsp_diagnostics` clean on all changed files.
  - [x] Phase 9: Write all worksheet features that the reader supports.
    - [x] Add BrtWsDim, BrtPane, BrtColInfo writing before sheet data section.
    - [x] Update BrtRowHdr to write custom row heights (miyRw) and hidden flags.
    - [x] Write rows with height/hidden but no data.
    - [x] Add BrtBeginSheetView/BrtEndSheetView wrapper with BrtSel for freeze panes.
    - [x] Write BrtMergeCell records for merged regions.
    - [x] Write BrtHLink records for hyperlinks with per-sheet .rels for external URLs.
    - [x] Write BrtBeginAFilter/BrtEndAFilter for autofilter range.
    - [x] Write BrtMargins, BrtPageSetup, BrtPrintOptions, BrtHeaderFooter for page setup.
    - [x] Write BrtDVal records for data validation (structural roundtrip: ranges, messages, type/operator).
    - [x] Write BrtBeginCFRule records for conditional formatting (structural roundtrip: rule type, ranges, priority, flags).
    - [x] Create writer/comments.rs: write comments as XML parts with per-sheet .rels.
    - [x] Update content types for comment parts.
    - [x] 15 new round-trip tests: merged cells, row height+hidden, column width+hidden, freeze pane, page margins, page setup landscape, autofilter, header/footer, comments, external hyperlink, internal hyperlink, hidden row no data, data validation, conditional format, combined features.
    - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (124 tests).
    - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).
    - [x] `lsp_diagnostics` clean on all changed files.
  - [x] Phase 10: Bindings integration and polish.
    - [x] Node.js binding: `open("file.xlsb")`, `fromBytes(xlsbData)`, `save("file.xlsb")` work via WorkbookExt dispatch (no code changes needed).
    - [x] Python binding: `Workbook.open()`, `Workbook.from_bytes()`, `Workbook.save()` work via WorkbookExt dispatch (no code changes needed).
    - [x] WASM binding: `fromBytes()` works automatically; added `saveXlsbBytes()` method alongside `saveXlsxBytes()`.
    - [x] All bindings use `duke-sheets` with `features = ["full"]` which includes `xlsb`.
    - [x] `cargo check` clean for all three binding Cargo.toml files.
    - [x] `cargo test --workspace --features full` passed (all crates).
    - [x] `cargo test -p duke-sheets-xlsb` passed (124 tests).

- [x] Reject XLSB files with a clear error instead of misreading them as XLSX.
  - [x] Add `zip` as a regular dependency in `crates/duke-sheets/Cargo.toml`.
  - [x] Add `FileFormat::Xlsb` variant to `detect_format` (peeks inside ZIP for `xl/workbook.bin`).
  - [x] Return clear error from `from_bytes` and `open` for XLSB files.
  - [x] Add `detect_format_identifies_xlsb` and `from_bytes_rejects_xlsb` tests.
  - [x] `lsp_diagnostics` clean, `cargo test -p duke-sheets --lib` passed (60 tests).

  - [x] Model: 35 types in chart_ex.rs (ChartExLayout, ChartEx, ChartExSeries, etc.).
  - [x] Reader: chart_ex.rs parser (~1700 lines), drawing reader detects chartEx via graphicData URI.
  - [x] Writer: chart_ex.rs emitter (~1080 lines), drawing writer emits mc:AlternateContent.
  - [x] Drawing: mc:Fallback preservation for chartEx anchors; fixed mc:Fallback corruption on non-chartEx files.
  - [x] Plumbing: content types, relationship types, style/color part numbering.
  - [x] Corpus: 11 chartEx charts read from 5 files, all layout types recognized.
  - [x] Roundtrip: all 5 chartEx corpus files roundtrip identically.
  - [x] Excel validation: waterfall chartEx roundtrip opens in Excel without repair.
  - [x] Bindings: Node.js, Python, WASM all expose ChartEx types.
  - Bugs fixed: attribution emitted as child element (spec: attribute), valueColors double-wrapped.

- [x] 2026-04-03: Chart validation against real-world files and Excel.
  - [x] Corpus smoke test: scan 128K xlsx files for charts, read with XlsxReader.
    - 558 chart files found (0.4%), 552/558 read OK (98.9%), 0 panics.
    - 6,260 total charts (6,234 worksheet + 26 chartsheets).
    - 6 failures: all corrupted files (missing workbook.xml.rels).
    - 11 Unsupported(unknown) charts: empty plotArea, no chart type element.
  - [x] Chart parity spreadsheet via Excel COM (one fat file, one roundtrip).
    - Generated `data/chart-parity.xlsx` via Excel COM (8 chart types).
    - Offline verifier: 75 deep field-level assertions, all passing.
    - `mise run generate:chart-parity` / `mise run test:chart-parity`.
  - [x] Chart roundtrip through Excel (write → Excel open → re-save → read back).
    - 22 charts + 1 chartsheet survive full Excel roundtrip with deep field-level assertions.
    - Bisected all 22 chart types individually: all pass.
  - [x] Fix bugs found during Excel validation:
    - Writer: ScatterMarkers emitted invalid `scatterStyle val="markersOnly"` (spec: `"marker"`).
    - Writer: Surface charts missing 3rd axId + serAx element (Excel requires all 3 axes).
    - Reader: PieExploded not detected (now checks series explosion attribute).
  - [x] Extend xlsx_roundtrip fuzz target to generate charts via Arbitrary.
    - 21 chart types, 1-5 series, formula/literal data refs, markers, trendlines, data labels.
    - 24K iterations with 0 crashes.
  - [x] Corpus roundtrip test: read → write → read, compare every chart field.
    - 547 chart files tested, 539 pass identical roundtrip (98.5%).
    - Fixed: writer omitted `smooth val="0"` (affected 165 files).
    - Remaining 8 mismatches: 6 axis property gaps (tickLblPos, axis spPr, txPr),
      2 regionMap chartEx with complex valueColorPositions.

- [x] 2026-04-03: Add chart lines (drop lines, high-low lines, series lines, leader lines) and up-down bars.
  - [x] Add ChartLines and UpDownBars model types to duke-sheets-chart.
  - [x] Add leader_lines field to DataLabels.
  - [x] Parse dropLines, hiLowLines, serLines, upDownBars, leaderLines in XLSX reader.
  - [x] Emit chart lines and up-down bars in XLSX writer (correct element ordering).
  - [x] Add 5 roundtrip tests (drop lines, high-low lines, series lines, up-down bars, leader lines).
  - [x] Update CHART_SUPPORT.md checklist (7 items checked off).
  - [x] `lsp_diagnostics` clean for changed files.
  - [x] `cargo check` and `cargo test` pass.

- [x] 2026-03-24: Add sparse row skip-value filtering across all bindings.
  - [x] Filter raw empty vs blank cell values in Node.js, WASM, and Python row iteration APIs.
  - [x] Omit rows whose cells are fully filtered out.
  - [x] Add core, Node.js, WASM, and Python tests for skip flag semantics.
  - [x] `lsp_diagnostics` clean for changed Rust files.
  - [x] Run targeted binding/core tests and `mise run test:report`.

- [x] 2026-03-19: Build two-phase Excel formula parity coverage.
  - [x] Add Excel COM generator for a real-Excel `formula-parity.xlsx` fixture.
  - [x] Add ignored duke-sheets assertion test that recalculates and compares
    against Excel cached values.
  - [x] Register the new Excel COM e2e module.
  - [x] `lsp_diagnostics` clean for changed formula parity test files.
  - [x] `cargo check -p duke-sheets-excel-com --tests` passed.
  - [x] `cargo check -p duke-sheets --tests` passed.
  - [x] `mise run test:report` ran; Rust/lib/nodejs/python suites passed, wasm
    node tests failed in existing `bindings/wasm` coverage, and the report then
    hit the existing `tools/test-report.sh: line 95: BASH_REMATCH[1]: unbound
    variable` failure.

- [x] 2026-03-16: Align WASM `getRowsBatch()` metadata flags with Node.js.
  - [x] Add WASM sparse row metadata fields/options in Rust and TS declarations.
  - [x] Match Node.js coordinate expansion so metadata-only cells are returned.
  - [x] `lsp_diagnostics` clean for changed WASM Rust/TS files.
  - [x] `cargo check --manifest-path bindings/wasm/Cargo.toml --target wasm32-unknown-unknown` passed.

- [x] 2026-03-16: Match Python row metadata flags with Node.js sparse row API.
  - [x] Extend `RowCell`/row iteration metadata for styles, merges, hyperlinks,
    comments, formulas, and images.
  - [x] Include metadata-only cells in sparse row batches when requested.
  - [x] `lsp_diagnostics` clean for changed Python binding Rust files.
  - [x] `cargo check --manifest-path bindings/python/Cargo.toml` passed.

- [x] 2026-03-16: Add sparse row iteration API for Python bindings.
  - [x] Add PyO3 sparse row batch types and `Worksheet.get_rows_batch()`.
  - [x] Add `RowIterator` plus `Worksheet.iterate_rows()`.
  - [x] `lsp_diagnostics` clean for changed Python binding Rust files.
  - [x] `cargo check --manifest-path bindings/python/Cargo.toml` passed.

- [x] 2026-03-16: Add sparse row iteration API for WASM bindings.
  - [x] Add WASM sparse row types and `Worksheet.getRowsBatch()` binding.
  - [x] Add JS `RowIterator` wrapper plus `Worksheet.iterateRows()`.
  - [x] Update TypeScript declarations without exposing `getRowsBatch()`.
  - [x] `lsp_diagnostics` clean for changed Rust/JS/TS files.
  - [x] `cargo check --manifest-path bindings/wasm/Cargo.toml --target wasm32-unknown-unknown` passed.
  - [x] `wasm-pack build --target web --dev` and `wasm-pack build --target web --release --out-dir pkg-web` passed.

- [x] 2026-03-16: Add sparse row iteration API for Node.js bindings.
  - [x] Add `CellStorage::cells_map()` and `Worksheet::populated_cells_in_range()`.
  - [x] Add Node.js `Worksheet.getRowsBatch()` sparse batch API.
  - [x] Add JS `RowIterator` wrapper plus `Worksheet.rows()`.
  - [x] Extend `index.d.ts` for sparse row iteration types.
  - [x] `lsp_diagnostics` clean for changed Rust/JS/TS files.
  - [x] `cargo check -p duke-sheets-core` passed.
  - [x] `cargo check --manifest-path bindings/nodejs/Cargo.toml` passed.
  - [x] `mise run test:report` reported passing core/formula/nodejs/python/wasm suites,
    then hit the existing `tools/test-report.sh: line 95: BASH_REMATCH[1]:
    unbound variable` failure.

- [x] Implement FILTERXML, WEBSERVICE, RTD, and IMAGE across the formula crate,
  calculation engine, and all bindings.
  - [x] Add quick-xml to `duke-sheets-formula` and wire callbacks/IMAGE metadata
    through calculation.
  - [x] Implement FILTERXML, WEBSERVICE, RTD, and IMAGE behavior with tests.
  - [x] Update Node.js, Python, and WASM calculation options for WEBSERVICE/RTD
    caches and expose IMAGE metadata from calculation results.
  - [x] Update function tracking docs.
  - [x] Run diagnostics, targeted tests, `cargo build`, and `mise run test:report`.
  - [x] Replace unsafe JsCallbackHandle with ThreadsafeFunction-based async
    callbacks in Node.js bindings. Callbacks now require `calculateWithOptionsAsync`
    and must return Promises. No `max_threads=1` constraint.

- [x] 2026-03-13: Phase 3 — Implement remaining stub/unimplemented functions.
  - [x] ACOT, ACOTH — math arccotangent functions.
  - [x] TRIMRANGE — trim blank rows/cols from array edges.
  - [x] PERCENTOF — Excel 365 subset/total aggregation.
  - [x] CELL — partial implementation (12 info_types).
  - [x] FORMULATEXT — evaluator special-case for raw ref access.
  - [x] ODDFPRICE, ODDFYIELD — odd first period bond pricing/yield.
  - [x] ODDLPRICE, ODDLYIELD — odd last period bond pricing/yield.
  - [x] INDIRECT — docs-based tests added (already implemented).
  - [x] OFFSET — evaluator special-case + full implementation.
  - Formula crate: 1048 tests passed (up from 1034).

- [x] 2026-03-13: Async callbacks in Node.js bindings via ThreadsafeFunction.
  - [x] Added `tokio_rt` feature to napi in Cargo.toml.
  - [x] Removed unsafe `JsCallbackHandle` struct (~140 lines of raw napi FFI).
  - [x] Simplified sync `calculateWithOptions` — no callback support (deadlock-safe).
  - [x] Refactored `calculateWithOptionsAsync` to accept callbacks via
    ThreadsafeFunction with `block_on` + async double-await pattern.
  - [x] Removed `max_threads=1` constraint — callbacks now work with parallel calc.
  - [x] Updated all callback tests to use async path.
  - [x] All 142 Node.js tests pass.

## Completed

- [x] 2026-03-20: Fix Excel formula parity mismatches in the formula engine.
  - [x] Reproduce the 14 reported parity failures with targeted unit coverage.
  - [x] Fix Excel behavior mismatches for POWER, INDEX, OR/AND, ISFORMULA,
    COUNTIFS criteria, SUMPRODUCT unary coercion, and DURATION.
  - [x] Improve numerical parity for NORM.S.INV, GAUSS, XIRR, ERF/ERFC,
    and BESSELI/BESSELJ, with a targeted XIRR epsilon override only.
  - [x] Reconcile the stale cached POWER/INDEX parity cells whose workbook value
    no longer matches the explicit expected type in `formula_parity.rs`.
  - [x] `cargo test --release -p duke-sheets-formula` passed (`1061` tests).
  - [x] `cargo test -p duke-sheets --test formula_evaluation` passed.
  - [x] `cargo test -p duke-sheets --test formula_parity -- --nocapture --ignored` passed.
  - [x] `lsp_diagnostics` clean for changed formula/evaluator/test files.
  - [x] `mise run test:report` passed.

- [x] 2026-03-19: Make cached plans reusable across ordinary value edits with targeted invalidation.
  - [x] Restore planner-time `MATCH`/`XMATCH` narrowing for `INDEX`, but record the workbook ranges consulted during that narrowing.
  - [x] Add `Worksheet::topology_generation()` separate from `mutation_count()`.
  - [x] Add `dirty_value_ranges` tracking on worksheets and clear them after successful calculation.
  - [x] Change `CalcCache::is_valid()` to use topology generations plus overlap checks against `value_sensitive_ranges`.
  - [x] Make `CalcCache` scope-aware so scoped calculations can also reuse cached plans.
  - [x] `cargo check -p duke-sheets` passed.
  - [x] `cargo test -p duke-sheets-formula` passed.
  - [x] `cargo test -p duke-sheets --lib calc_cache_survives` passed.
  - [x] `cargo test -p duke-sheets --lib calc_cache_invalidates` passed.
  - [x] `cargo test -p duke-sheets --lib sensitive_match_range` passed.
  - [x] Cold real-workbook sheet 4 trace stayed fast:
    - `plan_ms`: `15028.24`
    - `calc time`: `20.54s`
  - [x] Full-workbook real value-edit harness (editing unrelated `sheet 9!A1`) showed:
    - first run: `cache_hit=false`, `20.31s`
    - second run: `cache_hit=true`, `2.74s`
  - [x] Scoped sheet-4 real value-edit harness (editing unrelated `sheet 9!A1`) showed:
    - first run: `cache_hit=false`, `19.75s`
    - second run: `cache_hit=true`, `2.29s`
  - [x] Sensitive-range value edits still invalidate the cache as intended.

- [x] 2026-03-19: Add dirty-subgraph recalculation on cache hits.
  - [x] Extend the cached plan with reverse dependents and per-formula non-formula input ranges.
  - [x] On cache hits, seed affected formulas from dirty value ranges and walk cached reverse dependents.
  - [x] Skip evaluation entirely when a value edit affects no formulas in the cached scope.
  - [x] `cargo check -p duke-sheets` passed.
  - [x] Targeted cache validity tests still passed.
  - [x] Scoped sheet-4 unrelated value edit (editing `sheet 9!A1`) now showed:
    - first run: `cache_hit=false`, `19.84s`
    - second run: `cache_hit=true`, `150ms`
  - [x] Scoped sheet-4 same-sheet input edit (editing `K2`) showed:
    - first run: `cache_hit=false`, `19.60s`
    - second run: `cache_hit=true`, `112ms`
    - `cells_calculated=6`

- [x] 2026-03-19: Replace serial on-the-fly DFS dependency extraction with a precomputed dense precedent graph.
  - [x] Add a dense adjacency build (`DensePrecedents`) before DFS planning.
  - [x] Materialize precedents once per formula cell, in parallel when enabled.
  - [x] Refactor `build_eval_plan()` DFS to consume precomputed dense dep indices instead of calling `extract_references_recursive()` inside the stack walk.
  - [x] `cargo check -p duke-sheets` passed.
  - [x] `cargo test -p duke-sheets` passed.
  - [x] Real workbook sheet 4 engine trace after the planner redesign:
    - `parse_ms`: `2226.60`
    - `plan_ms`: `16319.10`
    - `eval_ms`: `2210.44`
    - `spill_fixup_ms`: `983.04`
    - `calc time`: `21.60s`
  - [x] Planner result: `plan_ms` dropped from `122135.93` to `16319.10`
    (~`-86.64%`), and total cold calc dropped from `127.87s` to `21.60s`
    (~`-83.11%`).

- [x] 2026-03-19: Reduce `build_eval_plan()` hash lookups by storing dense dependency indices on the DFS stack.
  - [x] Change DFS stack dep vectors from `CellKey` to dense `u32` indices.
  - [x] Remove repeated `cell_to_idx.get()` lookups during pop-time depth computation.
  - [x] `cargo check -p duke-sheets` passed.
  - [x] `cargo test -p duke-sheets` passed.
  - [x] Real workbook sheet 4 trace after the change:
    - `plan_ms`: `124923.72` → `122135.93` (~`-2.23%`)
    - `calc time`: `130.57s` → `127.87s` (~`-2.1%`)
  - [x] Conclusion: this is a real but small win; plan construction is still overwhelmingly dominant.

- [x] 2026-03-19: Flatten `FormulaCellIndex` to a per-sheet sorted `(row, col)` vector.
  - [x] Replace `BTreeMap<u32, Vec<u16>>` with `Vec<(u32, u16)>` + `partition_point` scans.
  - [x] `cargo check -p duke-sheets` passed.
  - [x] `cargo test -p duke-sheets` passed.
  - [x] Real workbook sheet 4 trace showed no meaningful improvement:
    - `plan_ms`: `124490.91` → `124923.72`
  - [x] Conclusion: row-range index layout was not the dominant plan-build bottleneck.

- [x] 2026-03-19: Add engine-side parallel trace and measure the real workbook’s actual level width.
  - [x] Add env-gated `DUKE_SHEETS_PARALLEL_TRACE=1` diagnostics in `CalculationEngine`.
  - [x] Report parse/build/eval/spill-fixup timings from the real execution path.
  - [x] Report dependency-level stats (total levels, widest level, levels >= threads,
    top slow levels with eval vs serial write timing).
  - [x] Synthetic `repeated-lookups` trace showed a single wide level of `16000` formulas.
  - [x] Real workbook sheet 4 trace showed:
    - `parse_ms`: `2284.17`
    - `plan_ms`: `124490.91`
    - `eval_ms`: `2206.74`
    - `spill_fixup_ms`: `959.73`
    - `total_levels`: `32`
    - `widest_level`: depth `1` with `630535` formulas
    - `levels >= threads`: `9`
  - [x] Conclusion: the real workbook has plenty of available width; the dominant
    remaining bottleneck is serial plan building (and some serial write/barrier cost),
    not lack of parallelizable work.

- [x] 2026-03-18: Add profiler-only parallelization report and compare serial vs parallel on the real lookup workbook.
  - [x] Add `--parallel-report` to `profile_calc` to compute a static formula dependency-width summary without changing the public API.
  - [x] Validate the report on the synthetic `repeated-lookups` fixture.
  - [x] Measure real workbook calc-only serial vs parallel runtime:
    - parallel: `125.42s`
    - serial: `134.76s`
    - speedup: ~`7%`
  - [x] Conclusion: parallelization helps, but it is not the dominant remaining bottleneck.

- [x] 2026-03-18: Inline `Worksheet::get_value_at_ref()` hot-path logic.
  - [x] Remove the extra `get_calculated_value_at()`/`Option` wrapper from the
    borrowed scalar accessor and resolve spill targets directly.
  - [x] `cargo check -p duke-sheets` passed.
  - [x] `cargo test -p duke-sheets-formula` passed.
  - [x] `cargo test -p duke-sheets` passed.
  - [x] `mise run perf:calc:compare -- --fixture repeated-lookups --sheet 0 --serial --baseline perf-snapshots/repeated-lookups-baseline.json --tolerance 0.10` passed.
  - [x] `calc_ir_per_formula` is now `394083.9054375` vs baseline
    `2104385.238` (-81.27% total).

- [x] 2026-03-18: Add source-coordinate exact lookup indexes for evaluator fast paths.
  - [x] Add direct worksheet-source lookup index cache in `EvalCache`.
  - [x] Route exact-match `MATCH`, `XMATCH`, `XLOOKUP`, and `VLOOKUP`
    through source-coordinate indexes instead of materialized arrays.
  - [x] Extend `INDEX` fast path to fetch a single source cell directly for
    scalar row/column requests.
  - [x] `cargo check -p duke-sheets` passed.
  - [x] `cargo test -p duke-sheets-formula` passed (1056 tests).
  - [x] `cargo test -p duke-sheets` passed.
  - [x] `mise run perf:calc:compare -- --fixture repeated-lookups --sheet 0 --serial --baseline perf-snapshots/repeated-lookups-baseline.json --tolerance 0.10` passed.
  - [x] `calc_ir_per_formula` improved from `2104385.238` to `398100.756125`
    (-81.08%) on `repeated-lookups`.

- [x] 2026-03-18: Add borrowed worksheet value access for evaluator hot paths.
  - [x] Add `Worksheet::get_value_at_ref()` with spill-resolution semantics.
  - [x] Add `impl From<&CellValue> for FormulaValue`.
  - [x] Switch evaluator single-cell, range materialization, and spill reads to
    the borrowed accessor.
  - [x] `cargo check -p duke-sheets` passed.
  - [x] `cargo test -p duke-sheets-formula` passed (1056 tests).
  - [x] `cargo test -p duke-sheets` passed.
  - [x] `mise run perf:calc:compare -- --fixture repeated-lookups --sheet 0 --serial --baseline perf-snapshots/repeated-lookups-baseline.json --tolerance 0.10` passed.
  - [x] `calc_ir_per_formula` improved from `2104385.238` to `2064792.7548125`
    (-1.88%) on `repeated-lookups`.

- [x] 2026-03-18: Add synthetic perf regression workflow for repeated lookup calculation.
  - [x] Add neutral generated `repeated-lookups` workload shared by the calc profiler
    and Criterion calculation benchmark.
  - [x] Extend `profile_calc` with `--fixture`, `--json`, `--once`,
    `--open-only`, and `--sheet` modes.
  - [x] Add `mise run perf:calc:callgrind`, `perf:calc:snapshot`, and
    `perf:calc:compare` tasks.
  - [x] Add `docs/PERF_REGRESSION.md` and README link for the workflow.
  - [x] `cargo build --release -p duke-sheets --features full --example profile_calc` passed.
  - [x] `cargo bench --features full -p duke-sheets --bench calculation --no-run` passed.
  - [x] `mise run perf:calc:callgrind -- --fixture repeated-lookups --sheet 0 --serial` passed.
  - [x] Snapshot/compare flow passed against `perf-snapshots/repeated-lookups-baseline.json`.

- [x] 2026-03-13: Implement OFFSET evaluator special-casing in `duke-sheets-formula`.
  - [x] Review Microsoft docs examples and evaluator/lookup integration points.
  - [x] Add evaluator-level `OFFSET` interception and reference-aware evaluation.
  - [x] Add docs-based OFFSET tests in `lookup.rs`.
  - [x] Run diagnostics, cargo tests, and `mise run test:report`.
  - `lsp_diagnostics` clean for `evaluator.rs` and `lookup.rs`.
  - `cargo test -p duke-sheets-formula` passed with 1048 tests.
  - `cargo test` passed.
  - `mise run test:report` exercised all suites, then exited non-zero in
    `tools/test-report.sh` after reporting passes.

- [x] 2026-03-13: Ran `mise run test:report` baseline before changes.
  - Formula crate: 1034 tests passed.
  - Overall repo test report passed.

- [x] Phase D1: Theme color resolution for XLSB reader/writer.
  - [x] Create reader/theme.rs: ThemePalette, parse_theme_palette, resolve_style_theme_colors, resolve_color_theme.
  - [x] Update reader/workbook.rs: WorkbookRelationships struct captures theme path from Type attribute.
  - [x] Update reader/mod.rs: load theme palette after styles, resolve theme colors on all cell styles.
  - [x] Update writer/mod.rs: emit xl/theme/theme1.xml with default Office theme, add content type override and relationship.
  - [x] 11 new tests: palette parsing, color resolution (no tint, positive tint, negative tint, passthrough), style resolution, theme XML emission, content types/rels, RGB color round-trip.
  - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (133 tests).
  - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).
  - [x] `lsp_diagnostics` clean on all changed files.

- [x] Add XLSB parity test infrastructure.
  - [x] Add `temp_fixture_xlsb()` to e2e common.rs.
  - [x] Add `roundtrip_through_excel_xlsb()` helper to e2e common.rs.
  - [x] Create XLSB parity generator test (converts formula-parity.xlsx → .xlsb via Excel COM).
  - [x] Register xlsb_parity module in e2e/main.rs.
  - [x] Create XLSB parity test in duke-sheets/tests/ (type checks + roundtrip, `#[ignore]`).
  - [x] Add `mise run generate:xlsb-parity` and `mise run test:xlsb-parity` tasks.
  - [x] Add duke-sheets-xlsb dev dependency to excel-com crate.
  - [x] `lsp_diagnostics` clean on all new files.
  - [x] `cargo check -p duke-sheets --features full --test xlsb_parity` passed.
  - [x] `cargo test -p duke-sheets-xlsb` passed (124 tests).
  - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).

- [x] Phase D2: Rich text support in XLSB shared strings.
  - [x] Update reader/styles.rs: return StylesData with both styles and font table.
  - [x] Update reader/shared_strings.rs: add SharedStringEntry enum, parse BrtSSTItem rich text runs (fRichStr flag, cRuns, ich/ifnt pairs), convert font indices to RunFont via font table.
  - [x] Update reader/mod.rs: reorder to parse styles before SST, pass font table to read_shared_strings.
  - [x] Update reader/worksheet.rs: handle SharedStringEntry::Rich via CellValue::rich_text.
  - [x] Update writer/shared_strings.rs: SstEntry enum (Plain/Rich), build_sst collects RichText cells with font dedup, write_sst emits fRichStr BrtSSTItem records with ich/ifnt run data.
  - [x] Update writer/styles.rs: accept extra fonts from SST rich text, intern and return font indices.
  - [x] Update writer/mod.rs: wire extra fonts through styles→SST.
  - [x] Update writer/worksheet.rs: use get_rich/get_plain for SST lookups.
  - [x] 7 new tests: reader plain/rich/single-run unit tests, writer round-trip rich text, mixed plain+rich, deduplication.
  - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (140 tests).
  - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).
  - [x] `lsp_diagnostics` clean on all changed files.

- [x] Phase D3: Formula token compilation — compile formula text → BIFF12 Ptg bytes.
  - [x] Add `function_index()` reverse lookup to `duke-sheets-formula` function table (OnceLock HashMap, case-insensitive).
  - [x] Create `biff12/compiler.rs`: AST postorder walk emitting BIFF12 Ptg bytes (tInt, tNum, tStr, tBool, tErr, tMissArg, tRef/tArea V-class, tRef3d/tArea3d, binary/unary ops, tFunc/tFuncVar, tAttrSum optimization).
  - [x] Update `writer/workbook.rs`: emit BrtExternSheet with per-sheet self-ref entries when formulas exist.
  - [x] Update `writer/worksheet.rs`: compile formula text via `compiler::compile_formula()`, emit real grbit+cce+token_bytes in BrtFmla* records.
  - [x] Update `writer/mod.rs`: build CompileContext with sheet names, detect has_formulas, wire through to worksheet writer.
  - [x] 3 function_index tests, 18 compiler unit tests, 8 formula text round-trip tests.
  - [x] `cargo test -p duke-sheets-formula` passed (1129 tests).
  - [x] `cargo test -p duke-sheets-xlsb` passed (166 tests).
  - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).
  - [x] `lsp_diagnostics` clean on all changed files.

- [x] Phase D4: Chart/drawing passthrough for XLSB.
  - [x] Add BRT_DRAWING record constant (0x0235) to records.rs.
  - [x] Create drawing_bundle.rs: binary bundle format for packing multiple ZIP entries (drawing XML, chart XML, .rels, media) into a single Vec<u8> for storage in raw_drawing_objects.
  - [x] Create reader/drawing.rs: read_drawing_bundle() scans sheet .rels for drawing relationships, reads drawing XML + drawing .rels + chart/style/color parts + media as raw bytes, packs into DrawingBundle.
  - [x] Update reader/mod.rs: after worksheet read, call read_drawing_bundle and store encoded bundle in ws.raw_drawing_objects.
  - [x] Create writer/drawing.rs: write_drawing_parts() decodes bundle and writes all entries to ZIP, or wraps raw anchor XML fragments in <xdr:wsDr> for XLSX-sourced data.
  - [x] Update writer/worksheet.rs: emit BrtDrawing record with correct rId when sheet has drawings.
  - [x] Update writer/mod.rs: call write_drawing_parts, write drawing relationship in sheet .rels, add content type overrides for drawing/chart parts.
  - [x] 4 new tests: drawing round-trip (bundle → write → read → verify bytes), drawing+chart round-trip, no-drawing baseline, anchor XML fallback.
  - [x] `cargo check -p duke-sheets-xlsb` clean, `cargo test -p duke-sheets-xlsb` passed (173 tests).
  - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).
  - [x] `lsp_diagnostics` clean on all changed files.

- [x] Fix BIFF12 formula decompilation for cce=0 records and tArray BIFF12 format.
  - [x] Read formula tokens from rgcb when cce=0 (Excel stores tokens there for certain formulas).
  - [x] Fix tArray placeholder size from 7 to 14 reserved bytes for BIFF12.
  - [x] Fix tArray extra data parsing: BIFF12 uses u32 1-based cols/rows and different SerAr type codes.
  - [x] Add tArray compilation support: emit 14 reserved bytes, u32 cols/rows, BIFF12 SerAr format.
  - [x] Write full CellParsedFormula with cb+rgcb in writer for tArray extra data.
  - [x] `cargo test -p duke-sheets-xlsb` passed (174 tests).
  - [x] XLSB parity tests: all 4 pass including roundtrip.

- [x] Fix BIFF12 formula token parser bugs causing 265 parity failures.
  - [x] Bug 1: tStr used u32 char count instead of u16 (ShortXLUnicodeString), causing parse loop to bail on string literals.
  - [x] Bug 2: IF/IFERROR/SUMIFS/COUNTIFS lost because tFuncVar tokens came after broken tStr — resolved by Bug 1 fix.
  - [x] Bug 3: `_xlfn.*` functions decompiled as `_nameNNN` — fixed BrtName record offset (chKey is 1 byte not 4), added CE function decompilation for func_idx=0xFF, strip `_xlfn.` prefix.
  - [x] `cargo test -p duke-sheets-xlsb` passed (173 tests).
  - [x] `cargo test -p duke-sheets-formula` passed (1126 tests).
  - [x] `cargo test -p duke-sheets --features full --lib` passed (63 tests).
  - [x] XLSB parity: 265 failures → 22 (remaining are IM* complex number precision).

- [x] Move chart parsing to shared duke-sheets-chart crate.
  - [x] Fix 103 compilation errors in `duke-sheets-chart::parse` (imports, error types, test adaptation).
  - [x] Update XLSX reader to use `duke_sheets_chart::parse::parse_chart_xml` / `parse_chart_ex_xml`.
  - [x] Update XLSB reader to parse charts from drawing rels via shared parsers.
  - [x] All crates compile, all tests pass (chart 49, xlsx 63, xlsb 173, parity 5).

- [x] Fix XLSB writer record payloads to match Excel output.
  - [x] BrtRowHdr: fix layout to [MS-XLSB] spec (3-byte grbitRw with flags in byte[1], ccolspan+spans). Record grows from 17→25 bytes.
  - [x] BrtWbProp: set default flags to 0x20 and byte[2]=0x01.
  - [x] BrtWsProp: set default worksheet property flags and sentinel values.
  - [x] BrtSheetView: set flags to 0x03DC, byte[13]=0x40, move zoom to offset 28.
  - [x] Fix reader parse_row_hdr to match spec (flags at byte[11], bit 4=hidden, bit 5=custom height).
  - [x] `cargo test -p duke-sheets-xlsb` passed (174 tests).
  - [x] `cargo test -p duke-sheets --features full --test xlsb_parity -- --ignored` passed (4 tests).

- [x] Fix XLSB writer styles.bin and sheet1.bin so Excel opens files without repair.
  - [x] styles.bin: encode Color::Auto as Indexed(64) in BrtColor, match Excel wire format.
  - [x] styles.bin: fix default fills to use theme colors (fg=theme 64, bg=theme 65).
  - [x] styles.bin: fix default border edges to use Indexed(0) color instead of all-zero.
  - [x] styles.bin: fix BrtXF byte 13 — encode fLocked/fHidden flags, set fLocked=1 for defaults.
  - [x] sheet1.bin: fix BrtWsProp trailing bytes (positions 19-21: 0x00 not 0xFF).
  - [x] sheet1.bin: fix BrtSheetView icvHdr at byte 14 (not 13), wScale at bytes 18-19 (not 28-29).
  - [x] sheet1.bin: fix BrtSel sqrefCount at offset 16 (not 20).
  - [x] sheet1.bin: add BrtWsFmtInfo (0x01E5) record before BrtBeginSheetData.
  - [x] sheet1.bin: write compact 2-byte BrtPageSetup for default settings, reorder records.
  - [x] sheet1.bin: skip BrtPrintOptions when no print options are set.
  - [x] sheet1.bin: fix formula-only cell collection (cells removed from storage but with formula data).
  - [x] `cargo test -p duke-sheets-xlsb` passed (174 tests).
  - [x] `cargo test -p duke-sheets --features full --test xlsb_parity -- --ignored` passed (4 tests).
  - [x] `cargo test -p duke-sheets-excel-com --test e2e test_xlsb_roundtrip_no_repair` passed.

## Known Issues

- XLSB parity: 22 complex number (IM*) function precision mismatches remain — formula engine produces more digits than Excel's cached values.

## XLSB Feature Matrix

Features marked ✅ work in both reader and writer. Features marked
📖 are read-only. Features marked ❌ are not handled.

### Core data
- ✅ Cell values (number, string, boolean, error, blank with style)
- ✅ Shared string table (plain + rich text with font runs)
- ✅ Inline strings
- ✅ Cell formulas (regular, with cached values)
- ✅ Formula token compilation (_xlfn.* functions, cross-sheet refs, all operators, 485 FTAB functions)
- ✅ Array formula token compilation (tArray with BIFF12 extra data format)
- 📖 Shared formulas (expanded to individual on write — no formula sharing optimization)
- ✅ Rich text in SST (font runs with bold/italic/size/color/name)

### Styles
- ✅ Number formats (built-in 0-49, custom 164+)
- ✅ Fonts (name, size, bold, italic, underline, strikethrough, color, family, charset, scheme, vertAlign)
- ✅ Fills (solid, pattern, gradient with stops)
- ✅ Borders (all 5 edges with u16 style + color, diagonal direction)
- ✅ Cell alignment (horizontal, vertical, wrap, shrink, indent, rotation, reading order)
- ✅ Cell protection (locked, hidden)
- ✅ Theme colors (resolved via xl/theme/theme1.xml)
- ✅ DXF styles (differential formats for CF — font and fill)

### Worksheet features
- ✅ Merged cells
- ✅ Hyperlinks (external URLs + internal cell refs via sheet .rels)
- ✅ Row heights (custom + default)
- ✅ Row hidden
- ✅ Column widths (custom + default)
- ✅ Column hidden
- ✅ Freeze panes
- ✅ Autofilter range
- ✅ Page margins
- ✅ Page setup (paper size, orientation, scale, fit-to-page)
- ✅ Print options (gridlines, headings)
- ✅ Header/footer (odd/even/first)
- ✅ Data validation (all 8 types, all operators, messages, formulas)
- ✅ Cell comments (read + write BIFF12 binary, VML shapes for display)

### Worksheet features — partial or missing
- ✅ Tables (read from xl/tables/*.bin, written to xl/tables/table{N}.bin with .rels + content types)
- ✅ Sheet visibility (visible, hidden, veryHidden)
- ✅ Conditional formatting: cellIs, expression, colorScale, dataBar, iconSet, top10, aboveAverage, text rules, duplicates, blanks, errors
- ✅ Row/column outline levels
- ✅ Active cell / selection state
- ✅ Zoom scale
- ✅ Row/column page breaks
- ✅ Conditional formatting: timePeriod rules
- ❌ Autofilter sort state (no sort_state in core AutoFilter model)
- ✅ Split panes
- ✅ Tab selected state

### Charts and drawings
- ✅ Standard charts (via shared parser in duke-sheets-chart)
- ✅ ChartEx (via shared parser)
- ✅ Chart style/color XML (passthrough)
- ✅ Drawing shapes (non-chart) — passthrough blob round-trip
- ❌ Images (not read or written independently — only via drawing passthrough)
- ❌ Programmatic shape creation/modification

### Workbook features
- ✅ Multiple sheets
- ✅ Date 1904 system
- ✅ BrtExternSheet (formula cross-sheet references)
- ✅ BrtName (_xlfn.* function names)
- ✅ User-defined named ranges (read + write with formula token compilation)
- ✅ Sheet visibility (visible, hidden, veryHidden)
- ❌ Workbook protection (no WorkbookProtection type in core model)
- ✅ Sheet protection
- ✅ Tab colors (RGB, theme, indexed)
- ✅ Active sheet / book views (activeTab in BrtBookView)
- ✅ Print area / print titles (builtin defined names)
- ❌ External links
- ❌ Custom views
- ❌ Calculation properties (needs core CalcMode enum + WorkbookSettings field, plus XLSX reader/writer parity — currently both formats hardcode defaults)

### Not implemented at all
- ❌ Pivot tables
- ❌ Sparklines
- ❌ Slicers
- ❌ Data connections / external data
- ❌ VBA macros (macro-enabled workbooks lose macros on round-trip)
- ❌ Form controls
- ❌ OLE objects
- ❌ Threaded comments (only legacy cell notes)

### Ancillary parts (lost on round-trip)
- ❌ Printer settings (xl/printerSettings/*.bin)
- ❌ Custom XML parts (customXml/)
- ❌ Document metadata (docMetadata/LabelInfo.xml)
- ❌ Custom properties (docProps/custom.xml)
- ❌ Calculation chain (xl/calcChain.bin)
- ❌ Binary indices (xl/worksheets/binaryIndex*.bin)

### Known issues
None.
