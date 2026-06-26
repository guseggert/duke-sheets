# Duke Sheets Feature Matrix

`✔` supported · `●` partial (see Notes) · `✖` not yet · `-` N/A.

## Cell values

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Number values | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_numbers`, `cell_values_round_trip::number_value_round_trips` | §18.3.1.4 |
| String values (SST) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_strings`, `sst_round_trip::ascii_string_round_trips` | §18.4.8 |
| Inline strings | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_e2e::data_types::string_values` | §18.3.1.53 |
| Boolean values | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_booleans`, `cell_values_round_trip::boolean_value_round_trips` | §18.18.11 |
| Error values (#REF!, #VALUE!, etc.) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::data_types::error_values`, `cell_values_round_trip::error_value_round_trips_for_standard_codes` | §18.18.11 |
| Formula cells with cached values | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_formula_cached_number`, `formula_round_trip::arithmetic_formula_round_trips_with_text_intact` | §18.3.1.40 |
| Formula cells without cached values | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_formula_no_cached_value` | §18.3.1.40 |
| Empty cells (formatting only) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_style_only_cells`, `cell_values_round_trip::empty_cell_with_format_only_round_trips` | §18.3.1.4 |
| Date values (stored as numbers) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::number_formats::date_format`, `number_format_round_trip::builtin_date_format_round_trips` | §18.17.4 |
| Rich text values | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_rich_text`, `rich_text_round_trip::three_run_rich_text_round_trips_text_and_run_count`, `excel_com_e2e::writing_xls::excel_can_read_rich_text_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_rich_text_we_emit` | §18.4.1 |
| Cell metadata index (`cm`) | R✔ W✔ | R✖ W✖ | R- W- | R- W- | `xlsx_e2e::formula_metadata::reader_parses_cm_attribute` | §18.3.1.4 |
| Value metadata index (`vm`) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.3.1.4 |
| Phonetic hint (`ph`) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.3.1.4 |
| Large row/column indices | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_large_indices`, `cell_values_round_trip::large_row_and_column_indices_round_trip` | - |
| Sparse data (non-contiguous cells) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_sparse_data`, `cell_values_round_trip::cells_round_trip_when_set_in_scrambled_order` | - |

## Formulas

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Arithmetic operators (+ - * / ^) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `formula_evaluation::test_evaluate_simple_formulas`, `formula_round_trip::arithmetic_formula_round_trips_with_text_intact` | §18.17.2 |
| Comparison operators (= <> > < >= <=) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `formula_evaluation::test_boolean_functions`, `formula_round_trip::comparison_formula_round_trips` | §18.17.2 |
| String concatenation (&) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `formula_evaluation::test_string_operations`, `formula_round_trip::concat_operator_round_trips_via_formula_path` | §18.17.2 |
| Percent operator (%) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `formula_evaluation::test_evaluate_simple_formulas`, `formula_round_trip::percent_operator_round_trips` | §18.17.2 |
| Unary plus and redundant parentheses preserved | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_formula_unary_plus_and_parens`, `duke-sheets-xlsb::writer::tests::formula_uplus_paren_roundtrip`, `xls_formula_round_trip::unary_plus_emits_ptg_uplus`, `xls_formula_round_trip::parentheses_emit_ptg_paren`, `excel_com_e2e::writing_xls::excel_byte_parity_for_uplus_paren_we_emit` | §18.17.2.2 | Parser keeps `+x` and `(x)` as explicit nodes instead of folding them away; XLS and XLSB compilers emit PtgUplus (0x12) / PtgParen (0x15) like Excel does, so `=+A1` and `=(A1+1)` round-trip verbatim. |
| Cell references (A1) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `formula_evaluation::test_evaluate_with_cell_references`, `formula_round_trip::range_reference_in_formula_round_trips` | §18.17.2.3 |
| Range references (A1:B2) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `formula_evaluation::test_evaluate_with_range_references`, `formula_round_trip::sum_function_round_trips` | §18.17.2.3 |
| Full row/column refs (A:A, 1:1) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `formula_evaluation::test_evaluate_sum`, `formula_round_trip::full_column_ref_round_trips_clamped_to_biff8_row_limit`, `formula_round_trip::full_row_ref_round_trips_with_biff8_col_extent` | §18.17.2.3 |
| Cross-sheet references | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::formulas::cross_sheet_ref`, `cross_sheet_formula_round_trip::cross_sheet_cell_ref_round_trips`, `excel_com_e2e::writing_xlsb::excel_can_evaluate_cross_sheet_formulas_we_emit` | §18.17.2.3 |
| Quoted sheet names | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::formulas::cross_sheet_quoted_name`, `cross_sheet_formula_round_trip::quoted_sheet_name_round_trips` | §18.17.2.3 |
| 3D references (Sheet1:Sheet3!A1) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.17.2.3 | Parses but evaluation limited |
| Named range refs in formulas | R✔ W✔ | R✔ W✔ | R✔ W● | R✖ W✖ | `xls_e2e::formulas::named_range_in_expression`, `named_range_round_trip::user_named_range_in_cell_formula_round_trips`, `excel_com_e2e::writing_xlsb::excel_can_evaluate_named_range_formulas_we_emit` | §18.2.5 | XLS round-trips formula text via tName ptg + NAME record; `Workbook::named_ranges` is not yet repopulated by the reader; XLSB writer emits PtgName (R-class) referencing BrtName index |
| Structured references (tables) | R✔ W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_with_calculated_column` | §18.17.2.3 | XLSB compiler drops structured refs |
| Shared formulas | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xls_e2e::formulas::shared_formula` | §18.3.1.40 |
| Array formulas (CSE) | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xls_e2e::formulas::cse_array_formula` | §18.3.1.40 |
| Dynamic array formulas (spill) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::roundtrip_dynamic_array_sequence` | [MS-XLSX] §2.6.3 |
| Data table formulas | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xls_e2e::formulas::data_table_formula` | §18.3.1.72 |
| External workbook refs `[book]!A1` | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.17.2.3 | XLSB reader captures, XLSX parser doesn't |
| Defined name refs in formulas | R✔ W✔ | R✔ W✔ | R✔ W● | R✖ W✖ | `xls_e2e::formulas::named_range`, `named_range_round_trip::workbook_scoped_constant_name_round_trips_in_formula` | §18.2.5 | XLS round-trips formula text via tName ptg + NAME record; `Workbook::named_ranges` is not yet repopulated by the reader |
| R1C1 reference mode (storage) | R✖ W✖ | R- W- | R- W- | R- W- | - | §18.2.29 |
| Array constants in formulas | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::formulas::array_constant`, `xls_formula_round_trip::array_constant_in_sum_emits_ptg_array_and_rgcb`, `xls_formula_round_trip::array_constant_round_trips_text`, `duke-sheets-xlsb::writer::tests::formula_array_constant_roundtrip`, `excel_com_e2e::writing_xls::excel_byte_parity_for_array_constants_we_emit` | §18.17.2.5 | XLS emits PtgArray + trailing rgcb element block; XLSB emits the BIFF12 PtgArray form. Both byte-parity verified against Excel-authored output. |
| Intersection operator (space) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `parser::tests::test_parse_intersection`, `parser::tests::test_parse_intersection_in_function`, `xls_formula_round_trip::intersection_formula_text_survives_xls_roundtrip`, `excel_com_e2e::writing::excel_can_evaluate_intersection_we_emit`, `excel_com_e2e::writing_xls::excel_can_evaluate_intersection_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_evaluate_intersection_we_emit` | §18.17.2.2 | Tokeniser emits Token::Space when whitespace separates two value-producing tokens; parser folds it into BinaryOp(Intersect). All three writers emit PtgIsect (BIFF8 / BIFF12 opcode 0x0F) with R-class PtgArea operands so Excel can intersect the two ranges; V-class operands cause Excel to collapse the range to its last cell. Formula text round-trips through Excel re-save for all three formats. Our own evaluator returns an error for the operator (Excel does the math). |
| Range union operator (comma) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `parser::tests::test_parse_union_in_parens`, `xls_formula_round_trip::union_formula_text_survives_xls_roundtrip`, `xls_formula_round_trip::union_parens_survive_outside_sum`, `duke-sheets-xlsb::writer::tests::union_parens_survive_outside_sum`, `excel_com_e2e::writing::excel_can_evaluate_union_we_emit`, `excel_com_e2e::writing_xls::excel_can_evaluate_union_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_evaluate_union_we_emit` | §18.17.2.2 | Parser distinguishes function-call commas from union commas: bare parens (`(A,B)`) fold the comma into BinaryOp(Union) wrapped in a Paren node — the parens are semantic (they make the union a single argument), so they survive in any context (`COUNT((A1,B1))`, multi-area `INDEX`, top-level `=(A1,B1)`), not just SUM. All three writers emit PtgUnion (0x10) with R-class PtgArea operands; the XLS SUM path keeps Excel's exact MemFunc + PtgParen + AttrSum shape. Formula text round-trips through Excel for all three formats. |
| Analysis ToolPak add-in functions (EDATE, NETWORKDAYS, GCD, ...) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_atp_addin_functions`, `xls_formula_round_trip::atp_function_emits_namex_funcvar_udf`, `excel_com_e2e::writing_xls::excel_byte_parity_for_all_xls_atp_functions_we_emit`, `excel_com_e2e::writing_xlsb::excel_byte_parity_for_all_xlsb_atp_functions_we_emit` | [MS-XLS] §2.5.198.91 | XLS serializes the legacy add-in form (PtgNameX → EXTERNNAME in an AddIn SUPBOOK + PtgFuncVar iftab=255), with R-class arguments; XLSB emits native BIFF12 PtgFunc/PtgFuncVar carrying the real Ftab index (also R-class args); XLSX stores plain formula text. All 93 functions in the range (Ftab 384..=476) are byte-parity verified against Excel-authored output for XLS and XLSB. |
| External add-in UDF calls (`[N]!Name(args)`) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `parser::tests::test_parse_external_function_call`, `calculation::tests::test_external_function_callback_via_calculation_options`, `xlsx_roundtrip::test_roundtrip_external_udf_formula_text`, `xls_formula_round_trip::external_udf_round_trips_text_via_externname`, `duke-sheets-xlsb::writer::tests::formula_text_roundtrip_external_udf`, `excel_com_e2e::writing_xls::excel_preserves_external_udf_xls_we_emit`, `excel_com_e2e::writing_xlsb::excel_preserves_external_udf_xlsb_we_emit` | [MS-XLSX] §2.2.2; [MS-XLS] §2.5.198.17; [MS-XLSB] §2.4.759, §2.4.810 | Evaluation delegates to `CalculationOptions::external_fn`; if no callback resolves the call, calculation preserves the existing cached cell value. XLS/XLSB write UDF calls through add-in supporting-link/name records and `PtgNameX + PtgFuncVar(0x00FF)`. |
| Formula evaluation (on read) | ✔ | ✔ | - | - | `formula_parity::formula_parity_matches_excel_cached_values` | - |
| Workbook recalculation | ✔ | ✔ | - | - | `calculation::test_workbook_calculate_*` | - |
| Circular references | ✔ | ✔ | - | - | `calculation::circular_ref*` | - |
| Iterative calculation | ✖ | ✖ | - | - | - | §18.2.2 | Settings parsed, limited evaluator support |
| Volatile function tracking (NOW, RAND, TODAY) | ✔ | ✔ | - | - | `formula_evaluation::test_evaluate_simple_formulas` | - |
| Dependency tracking | ✔ | ✔ | - | - | `calculation::dependency*` | - |
| Spill blocking detection | ✔ | ✔ | - | - | `calculation::spill*` | - |

## Formula functions

One row per category, backed by the formula engine test suite. Individual function coverage lives in the evaluator unit tests (~1,100 tests).

| Category | Runtime | Tests | Notes |
|----------|---------|-------|-------|
| Math (SUM, ROUND, ABS, POWER, TRIG, ...) | ✔ | `duke-sheets-formula::functions::math` (~96 fns) | |
| Statistical (AVERAGE, STDEV, NORM.DIST, ...) | ✔ | `duke-sheets-formula::functions::statistical` (~200 tests) | |
| Logical (IF, AND, OR, IFS, IFERROR, ...) | ✔ | `formula_evaluation::test_evaluate_if` | |
| Text (LEFT, MID, CONCAT, TEXTJOIN, ...) | ✔ | `duke-sheets-formula::functions::text` (~74 tests) | |
| Date/time (YEAR, DATE, WEEKDAY, DATEDIF, NETWORKDAYS, ...) | ✔ | `duke-sheets-formula::functions::date` (~42 tests) | |
| Financial (PMT, NPV, IRR, FV, ...) | ✔ | `duke-sheets-formula::functions::financial` (~152 tests) | |
| Lookup/reference (VLOOKUP, INDEX, XLOOKUP, OFFSET, ...) | ✔ | `duke-sheets-formula::functions::lookup` (~52 tests) | |
| Engineering (CONVERT, BIN2DEC, IMABS, ...) | ✔ | `duke-sheets-formula::functions::engineering` (~112 tests) | |
| Information (ISBLANK, ISERROR, CELL, ...) | ✔ | `duke-sheets-formula::functions::info` (~7 tests) | |
| Database (DSUM, DAVERAGE, ...) | ✔ | `duke-sheets-formula::functions::database` (~25 tests) | |
| Compatibility (legacy aliases) | ✔ | `duke-sheets-formula::functions::compatibility` (~108 tests) | |
| Dynamic array (SEQUENCE, FILTER, SORT, UNIQUE, TRANSPOSE, RANDARRAY) | ✔ | `calculation::spill*` | |
| Statistical advanced (LINEST, LOGEST, FORECAST.ETS, ...) | ✔ | `duke-sheets-formula::functions::statistical_extra` (~16 tests) | |
| LAMBDA (define) | ✖ | - | Parses as function token; evaluation passes through |
| LET | ✖ | - | Parses; passthrough only; no named binding |
| MAP, REDUCE, SCAN, BYCOL, BYROW, MAKEARRAY | ✖ | - | Stub returns #N/A; needs evaluator lazy-eval |
| Cube functions (CUBEVALUE, CUBEMEMBER, ...) | ✖ | - | Requires OLAP server; returns #N/A |
| STOCKHISTORY | ✖ | - | Requires Microsoft service; returns #N/A |
| CALL, REGISTER.ID (DLL interop) | ✖ | - | Not supported by design; returns #N/A |
| WEBSERVICE, ENCODEURL, FILTERXML | ✖ | - | Not implemented |

## Rich text

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Rich text in shared strings | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_rich_text`, `rich_text_round_trip::three_run_rich_text_round_trips_text_and_run_count` | §18.4.2 |
| Rich text in inline strings | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_rich_text` | §18.4.4 |
| Rich text in comments | R✔ W✔ | R✔ W✔ | R● W✖ | R✖ W✖ | `com_e2e::rich_text::read_rich_text_from_excel` | §18.7.5 |
| Run properties: bold/italic | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_rich_text`, `rich_text_round_trip::rich_text_with_bold_run_round_trips` | §18.4.7 |
| Run properties: font size/name/color | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_rich_text`, `rich_text_round_trip::rich_text_with_size_and_color_round_trips` | §18.4.7 |
| Run properties: underline styles | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_font_styles`, `rich_text_round_trip::rich_text_with_underline_run_round_trips` | §18.4.7 |
| Run properties: strikethrough | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_font_styles`, `rich_text_round_trip::rich_text_with_strikethrough_run_round_trips` | §18.4.7 |
| Run properties: sub/superscript | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_font_vertical_align`, `rich_text_round_trip::rich_text_with_superscript_run_round_trips` | §18.4.7 |
| Run properties: shadow/outline/emboss/engrave | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.4.7 |
| Line breaks within a cell | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_strings`, `sst_round_trip::line_break_within_string_round_trips` | - |
| Phonetic guide (Japanese furigana) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.4.3 |

## Cell styles

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Font: bold | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_font_styles`, `font_round_trip::bold_round_trips` | §18.8.22 |
| Font: italic | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_font_styles`, `font_round_trip::italic_round_trips` | §18.8.22 |
| Font: size | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::font_styles::font_size`, `font_round_trip::font_size_round_trips` | §18.8.29 |
| Font: name/family | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::font_styles::font_name`, `font_round_trip::font_name_round_trips` | §18.8.22 |
| Font: color (RGB) | R✔ W✔ | R✔ W✔ | R✔ W● | R✖ W✖ | `xlsx_e2e::font_styles::font_color`, `font_round_trip::rgb_color_falls_back_to_auto` | §18.8.19 | XLS writer falls back to Auto/Indexed for arbitrary RGB; needs PALETTE-emission slice for full RGB support |
| Font: color (theme) | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_chart_style_color_passthrough` | §18.8.19 |
| Font: color (indexed) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::font_styles::font_color`, `font_round_trip::indexed_color_round_trips` | §18.8.19 |
| Font: underline (single) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::font_styles::underline_single`, `font_round_trip::underline_round_trips` | §18.4.13 |
| Font: underline (double) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::styles::font_underline_double`, `font_round_trip::double_underline_round_trips` | §18.4.13 |
| Font: underline (accounting single/double) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_font_styles`, `font_round_trip::accounting_underline_round_trips` | §18.4.13 |
| Font: strikethrough | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::font_styles::strikethrough`, `font_round_trip::strikethrough_round_trips` | §18.8.37 |
| Font: superscript | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::font_styles::superscript`, `font_round_trip::superscript_round_trips` | §18.18.85 |
| Font: subscript | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::font_styles::subscript`, `font_round_trip::subscript_round_trips` | §18.18.85 |
| Fill: solid color | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::fill_styles::solid_fill_red`, `cell_format_round_trip::solid_fill_with_indexed_color_round_trips` | §18.8.20 |
| Fill: pattern fill | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::styles::fill_pattern_fills`, `cell_format_round_trip::pattern_fill_round_trips` | §18.8.22 |
| Fill: gradient fill | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_gradient_fill` | §18.8.24 |
| Border: all sides (top/bottom/left/right) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::border_styles::thin_border_all_sides`, `cell_format_round_trip::border_thin_all_sides_round_trips` | §18.8.4 |
| Border: diagonal | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::styles::diagonal_border_down`, `cell_format_round_trip::diagonal_border_round_trips` | §18.8.4 |
| Border: styles (thin/medium/thick/dashed/dotted/double) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_border_styles`, `cell_format_round_trip::border_individual_sides_round_trip` | §18.18.3 |
| Border: color | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::border_styles::border_color`, `cell_format_round_trip::border_color_indexed_non_black_round_trips` | §18.8.4 |
| Alignment: horizontal (left/center/right/justify/fill) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::alignment::horizontal_center`, `cell_format_round_trip::horizontal_alignment_round_trips` | §18.18.40 |
| Alignment: vertical (top/center/bottom) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::alignment::vertical_bottom`, `cell_format_round_trip::vertical_alignment_round_trips` | §18.18.88 |
| Alignment: wrap text | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::alignment::wrap_text`, `cell_format_round_trip::wrap_text_round_trips` | §18.8.1 |
| Alignment: shrink to fit | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::alignment::shrink_to_fit`, `cell_format_round_trip::shrink_to_fit_round_trips` | §18.8.1 |
| Alignment: rotation | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::alignment::rotation`, `cell_format_round_trip::rotation_round_trips` | §18.8.1 |
| Alignment: indent | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::alignment::indent`, `cell_format_round_trip::indent_round_trips` | §18.8.1 |
| Alignment: reading order (LTR/RTL) | R● W✔ | R● W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::styles::reading_order_rtl`, `cell_format_round_trip::reading_order_round_trips` | §18.8.1 |
| Alignment: justifyLastLine | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | §18.8.1 |
| Number format: builtin IDs (0-49, 164+) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_style_roundtrip::test_roundtrip_number_format_styles`, `number_format_round_trip::builtin_percent_format_round_trips` | §18.8.31 |
| Number format: custom format strings | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::number_formats::custom_decimal_format`, `number_format_round_trip::custom_currency_format_round_trips` | §18.8.31 |
| Cell protection: locked | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::styles::cell_protection_locked`, `cell_protection_round_trip::explicitly_locked_state_round_trips` | §18.8.33 |
| Cell protection: formula hidden | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::styles::cell_protection_formula_hidden`, `cell_protection_round_trip::formula_hidden_cell_round_trips` | §18.8.33 |
| Named cell styles (cellStyleXf) | R✔ W✔ | R✔ W✔ | R● W✖ | R✖ W✖ | `xlsx_e2e::named_cell_styles_roundtrip::roundtrip_preserves_cell_style_xfs_and_named_styles` | §18.8.8 |
| Differential formats (DXF) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_formatting_roundtrip::test_roundtrip_dxf_styles` | §18.8.14 |
| Table cell style (tableStyleInfo) | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_basic` | §18.8.42 |

## Themes and theme colors

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Theme color scheme (12 slots) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_style_color_passthrough` | §20.1.6.2 (DrawingML) |
| Theme font scheme | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §20.1.4.1.18 (DrawingML) |
| Theme format scheme (fills/lines/effects) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §20.1.4.1.8 (DrawingML) |
| Theme override per-sheet | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.2.4 |
| Custom theme (replace default) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §20.1.6.9 (DrawingML) |

## Row/column dimensions

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Custom row heights | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_row_heights_column_widths`, `row_column_dimensions_round_trip::custom_row_height_round_trips` | §18.3.1.73 |
| Custom column widths | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_row_heights_column_widths`, `row_column_dimensions_round_trip::custom_column_width_round_trips` | §18.3.1.13 |
| Hidden rows | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_hidden_rows_columns`, `row_column_dimensions_round_trip::hidden_rows_round_trip` | §18.3.1.73 |
| Hidden columns | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_hidden_rows_columns`, `row_column_dimensions_round_trip::hidden_columns_round_trip` | §18.3.1.13 |
| Outline levels / grouping | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_outline_metadata`, `row_column_dimensions_round_trip::row_outline_level_round_trips`, `row_column_dimensions_round_trip::column_outline_level_round_trips` | §18.3.1.73 |
| Collapsed outline state | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::formula_metadata::outline_and_sheet_view_metadata`, `row_column_dimensions_round_trip::row_collapsed_state_round_trips`, `row_column_dimensions_round_trip::column_collapsed_state_round_trips` | §18.3.1.73 |
| Default row height | R✖ W✖ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `excel_com_e2e::writing_xlsb::excel_can_read_dimensions_we_emit` | §18.3.1.30 | XLSB BrtWsFmtInfo carries miyDefRwHeight in twips; the fUnsynced flag (bit 0) tells Excel the value is user-set, otherwise Excel resets it to 15pt on save. |
| Default column width | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.30 | Excel's BrtWsFmtInfo cchDefColWidth is rewritten to the base font width (8) on save regardless of input; Excel persists per-column custom widths via BrtColInfo records instead, not via this default. Treat as best-effort hint only. |
| Outline summary position | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | §18.3.1.35 |

## Merged cells

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Basic merged cell ranges | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::merged_cells::merged_cells_horizontal`, `merged_cells_round_trip::horizontal_merge_round_trips` | §18.3.1.55 |
| Merge spanning many rows/cols | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::merged_cells::merged_cells_block`, `merged_cells_round_trip::block_merge_round_trips` | §18.3.1.55 |
| Multiple merged regions per sheet | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::merged_cells::multiple_merged_regions`, `merged_cells_round_trip::multiple_disjoint_merges_round_trip` | §18.3.1.55 |

## Hyperlinks

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| External URL hyperlinks | R✔ W● | R✔ W✔ | R✔ W✔ | R✖ W✖ | `duke-sheets-xlsx::reader::mod::hyperlink*`, `hyperlinks_round_trip::external_url_hyperlink_round_trips`, `hyperlinks_round_trip::lo_can_read_hyperlinks_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_hyperlinks_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_hyperlinks_we_emit` | §18.3.1.36 | Reader preserves; XLSX writer round-trip incomplete; XLSB writer + Excel parity verified |
| Internal hyperlinks (Sheet!A1) | R✔ W● | R✔ W✔ | R✔ W✔ | R✖ W✖ | `duke-sheets-xlsx::reader::mod::hyperlink*`, `hyperlinks_round_trip::internal_hash_target_round_trips`, `hyperlinks_round_trip::lo_can_read_hyperlinks_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_hyperlinks_we_emit` | §18.3.1.36 | |
| Mailto hyperlinks | R✖ W✖ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `excel_com_e2e::writing_xlsb::excel_can_read_mailto_hyperlink_we_emit` | §18.3.1.36 | `mailto:` URLs ride the existing external-hyperlink rel-id path; Excel preserves the prefix and query string |
| Hyperlink tooltips | R✖ W✖ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `excel_com_e2e::writing_xlsb::excel_can_read_hyperlinks_we_emit` | §18.3.1.36 | XLSB BrtHLink trailing XLWideString carries the tooltip; survives Excel round-trip |
| Hyperlink display text | R✖ W✖ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `hyperlinks_round_trip::external_url_hyperlink_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_hyperlinks_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_hyperlinks_we_emit` | §18.3.1.36 | XLS HLINK record carries display name in `cch + UTF-16LE` block; XLSB BrtHLink trailing XLWideString does the same; both round-trip survive Excel parity |
| Hyperlinks in rich text runs | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.4.7 |

## Comments

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Plain-text comments with author | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `duke-sheets-xls::comments_round_trip::single_comment_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_comment_we_emit` | §18.7.5 |
| Comments with Unicode text | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `duke-sheets-xls::comments_round_trip::unicode_comment_text_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_unicode_comment_we_emit` | §18.7.5 |
| Rich text in comments | R✔ W✔ | R✔ W✔ | R● W✖ | R✖ W✖ | `com_e2e::rich_text::read_rich_text_from_excel` | §18.7.5 |
| Multiple comments per sheet | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `duke-sheets-xls::comments_round_trip::multiple_comments_on_same_sheet_round_trip`, `excel_com_e2e::writing_xls::excel_can_read_multiple_comments_we_emit` | §18.7.5 |
| VML legacy drawing for comments | R✔ W✔ | R- W- | R- W- | R- W- | `xlsx_formatting_roundtrip::test_comments_emit_vml_and_legacy_drawing` | §14.1 (VML) |
| Comment positioning (anchor) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §14.1 (VML) |
| Threaded comments (modern) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | [MS-XLSX] §2.3.19 |
| Threaded comments: mentions | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | [MS-XLSX] §2.3.19 |

## Named ranges

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Workbook-scoped names | R✔ W✔ | R✔ W✔ | R● W● | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_named_ranges`, `named_range_round_trip::workbook_scoped_constant_name_round_trips_in_formula` | §18.2.5 | XLS reader doesn't yet repopulate `workbook.named_ranges`; only formula-text references survive round-trip |
| Sheet-scoped names | R✔ W✔ | R✔ W✔ | R● W● | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_named_ranges`, `named_range_round_trip::sheet_scoped_name_round_trips_in_formula` | §18.2.5 | XLS reader doesn't yet repopulate `workbook.named_ranges`; only formula-text references survive round-trip |
| Names with formula bodies | R✔ W✔ | R✔ W✔ | R● W● | R✖ W✖ | `xls_e2e::formulas::named_range_in_expression`, `named_range_round_trip::workbook_scoped_constant_name_round_trips_in_formula` | §18.2.5 | XLS reader doesn't yet repopulate `workbook.named_ranges` |
| Names referencing ranges | R✔ W✔ | R✔ W✔ | R● W● | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_named_ranges`, `named_range_round_trip::user_named_range_in_cell_formula_round_trips` | §18.2.5 | XLS reader doesn't yet repopulate `workbook.named_ranges` |
| Hidden / built-in names (_xlnm.*) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_print_area`, `print_names_round_trip::print_area_round_trips` | §18.2.5 |
| Names with comments | R✖ W✖ | R✔ W✔ | R- W- | R✖ W✖ | `duke-sheets-xlsb::writer::tests::named_range_comment_roundtrip`, `excel_com_e2e::writing_xlsb::excel_can_read_named_range_comment_we_emit` | §18.2.5 | XLSB BrtName trailing strings carry comment + customMenu + description + help + statusBar (5 XLNullableWideStrings); writer emits `comment` field, reader parses first slot back into `NamedRange.comment`. Excel preserves the description through round-trip. |

## Tables

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Basic tables with headers | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_basic` | §18.5.1 |
| Totals row | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_with_totals` | §18.5.1 |
| Totals row functions (SUM/AVG/COUNT/MIN/MAX/STDEV/VAR/CUSTOM) | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_with_totals` | §18.5.1.1 |
| Calculated columns (formula) | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_with_calculated_column` | §18.5.1.1 |
| Header row visibility | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_basic` | §18.5.1 |
| Table styles (built-in names) | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_basic` | §18.8.42 |
| Table styles: showFirstColumn/showLastColumn | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_basic` | §18.8.42 |
| Table styles: showRowStripes/showColumnStripes | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_table_basic` | §18.8.42 |
| Custom table styles (tableStyle + dxf) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.8.40 |
| Multiple tables per sheet | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_multiple_tables` | - |
| Table AutoFilter integration | R✔ W✔ | R✔ W✔ | R- W- | R✖ W✖ | `xlsx_roundtrip::test_auto_filter_range_only` | §18.3.2 |

## Conditional formatting

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| cellIs rules (comparison operators) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_conditional_format_cell_is`, `conditional_format_round_trip::cellis_greater_than_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_conditional_formats_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_conditional_formats_we_emit` | §18.3.1.10 |
| Formula-based (expression) rules | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_conditional_format_expression`, `conditional_format_round_trip::expression_rule_round_trips` | §18.3.1.43 |
| beginsWith / endsWith | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.10 |
| containsText / notContainsText | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.10 |
| containsBlanks / notContainsBlanks | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.10 |
| containsErrors / notContainsErrors | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.10 |
| timePeriod rules (today/yesterday/...) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.10 |
| aboveAverage / belowAverage | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.10 |
| top10 / bottom10 | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.10 |
| duplicateValues / uniqueValues | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.10 |
| Color scale (2-color) | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_color_scale` | §18.3.1.16 |
| Color scale (3-color) | R● W● | R● W● | R✖ W✖ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_color_scale` | §18.3.1.16 | Midpoint config limited |
| Data bars (solid) | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_data_bar` | §18.3.1.28 |
| Data bars (gradient) | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | §18.3.1.28 |
| Data bars (negative bar color/config) | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | [MS-XLSX] §2.6.7 |
| Data bars (axis position) | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | [MS-XLSX] §2.6.7 |
| Icon set: 3 arrows / 3 arrows gray | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_icon_set` | §18.3.1.49 |
| Icon set: 3 traffic lights | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_icon_set` | §18.3.1.49 |
| Icon set: 3 signs / symbols / flags | R● W● | R● W● | R✖ W✖ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_icon_set` | §18.3.1.49 | Not all variants tagged |
| Icon set: 4 variants | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.49 |
| Icon set: 5 variants | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.49 |
| Icon set: custom icons (MS-XLSX ext) | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | [MS-XLSX] §2.6.8 |
| Multiple CF rules per range | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_multiple_rules`, `conditional_format_round_trip::multiple_rules_per_sheet_round_trip` | §18.3.1.18 |
| Rule priority / stopIfTrue | R✔ W✔ | R✔ W✔ | R● W✖ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_multiple_rules` | §18.3.1.10 |
| Extension-list CF rules (x14) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | [MS-XLSX] §2.6.4 |

## Data validation

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| List validation (inline values) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_data_validation_list`, `data_validation_round_trip::list_inline_values_round_trip`, `excel_com_e2e::writing_xls::excel_can_read_data_validations_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_data_validations_we_emit` | §18.3.1.32 |
| List validation (cell-range source) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::data_validation::list_validation`, `data_validation_round_trip::list_cell_range_source_round_trips` | §18.3.1.32 |
| List validation (named range source) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.32 |
| Whole number validation | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_formatting_roundtrip::test_roundtrip_data_validation_number`, `data_validation_round_trip::whole_number_between_round_trips` | §18.3.1.32 |
| Decimal validation | R● W● | R● W● | R● W● | R✖ W✖ | `xlsx_e2e::data_validation::whole_number_validation`, `data_validation_round_trip::decimal_validation_round_trips` | §18.3.1.32 | XLS round-trips numeric value through a parse-and-recompile cycle; lossy for non-trivial precision |
| Date validation | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.32 |
| Time validation | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.32 |
| Text length validation | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::data_validation::text_length_validation`, `data_validation_round_trip::text_length_round_trips` | §18.3.1.32 |
| Custom formula validation | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.32 |
| Input messages (prompt) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::data_validation::validation_with_messages`, `data_validation_round_trip::input_message_round_trips` | §18.3.1.32 |
| Error alerts (stop/warning/info styles) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_e2e::data_validation::validation_with_messages`, `data_validation_round_trip::error_alert_round_trips` | §18.3.1.32 |
| Drop-down arrow visibility | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.32 |
| Extension-list validation (x14 custom) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | [MS-XLSX] §2.6.5 |

## AutoFilter

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Filter range | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_auto_filter_range_only`, `autofilter_round_trip::filter_range_only_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_autofilter_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_autofilter_we_emit` | §18.3.2 |
| Value filter (includes) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_auto_filter_value_filter`, `autofilter_round_trip::value_filter_round_trips_as_equal_or`, `excel_com_e2e::writing_xlsb::excel_can_read_discrete_value_filter_we_emit` | §18.3.2.8 |
| Custom filter (operator-based) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_auto_filter_custom_filter`, `autofilter_round_trip::custom_dual_condition_and_round_trips` | §18.3.2.1 |
| Top 10 filter | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_auto_filter_top10`, `autofilter_round_trip::top10_filter_round_trips`, `excel_com_e2e::writing_xlsb::excel_can_read_autofilter_we_emit` | §18.3.2.10 |
| Dynamic filter (above/below avg, etc.) | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xlsx_roundtrip::test_auto_filter_dynamic`, `duke-sheets-xlsb::writer::tests::dynamic_filter_in_process_roundtrip`, `excel_com_e2e::writing_xlsb::excel_can_read_dynamic_filter_we_emit` | §18.3.2.5 | BrtDynamicFilter id 0x00AB confirmed by dumping bytes of an Excel-emitted XLSB with the AboveAverage filter applied. |
| Date group filter (year/month/day) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.2.6 |
| Color filter | R✖ W✖ | R✔ W✔ | R- W- | R✖ W✖ | `excel_com_e2e::writing_xlsb::excel_can_read_color_filter_we_emit`, `duke-sheets-xlsb::writer::tests::color_filter_in_process_roundtrip` | §18.3.2.2 | XLSB BrtColorFilter (id 0x00A9, 8-byte payload: dxfid u32 + fCellColor u32) survives Excel round-trip. |
| Icon filter | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | §18.3.2.7 |
| Multiple column filters | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_auto_filter_multiple_columns`, `autofilter_round_trip::multiple_columns_round_trip` | §18.3.2 |
| Sort state on autofilter | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.68 |

## Sheet views

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Freeze panes (rows) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_multi_selection_with_freeze_panes`, `freeze_panes_round_trip::freeze_first_two_rows_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_visual_state_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_visual_state_we_emit` | §18.3.1.66 |
| Freeze panes (columns) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_multi_selection_with_freeze_panes`, `freeze_panes_round_trip::freeze_first_column_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_visual_state_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_visual_state_we_emit` | §18.3.1.66 |
| Freeze + split combination | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_split_panes_and_selection` | §18.3.1.66 |
| Split panes | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_split_panes_and_selection` | §18.3.1.66 |
| Active cell / active range | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_multi_range_sqref`, `sheet_view_round_trip::active_cell_round_trips` | §18.3.1.78 |
| Multi-range selection (sqref) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_multi_range_sqref`, `sheet_view_round_trip::multi_range_selection_round_trips` | §18.3.1.78 |
| Zoom level (normal/page-layout/page-break) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `com_e2e::xls_reader::xls_zoom`, `sheet_view_round_trip::zoom_at_75_percent_round_trips` | §18.3.1.87 |
| View mode (normal/pageBreakPreview/pageLayout) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.87 |
| Gridlines visibility | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.87 |
| Row/column header visibility | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.87 |
| Right-to-left view | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.87 |
| Tab color | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.82 |
| Sheet visibility (visible/hidden/veryHidden) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::sheet_properties::hidden_sheet`, `sheet_visibility_round_trip::very_hidden_sheet_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_visual_state_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_visual_state_we_emit` | §18.2.19 |
| Named sheet views (x16 ext) | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | [MS-XLSX] §2.3.17 |

## Page setup and print

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Page orientation (portrait/landscape) | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_page_setup_and_header_footer`, `page_setup_round_trip::landscape_orientation_round_trips` | §18.3.1.62 |
| Paper size | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_page_setup_and_header_footer`, `page_setup_round_trip::paper_size_a4_round_trips` | §18.3.1.62 |
| Scale percentage | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_page_setup_and_header_footer`, `page_setup_round_trip::scale_percentage_round_trips` | §18.3.1.62 |
| Fit to width/height | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.62 |
| Margins | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_page_setup_and_header_footer`, `page_setup_round_trip::page_margins_round_trip` | §18.3.1.59 |
| Center horizontally / vertically on page | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.61 |
| Print area | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_print_area`, `print_names_round_trip::print_area_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_print_names_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_print_names_we_emit` | §18.2.5 |
| Print titles: repeat rows | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_repeat_rows`, `print_names_round_trip::repeat_rows_only_round_trips` | §18.2.5 |
| Print titles: repeat columns | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_repeat_cols`, `print_names_round_trip::repeat_cols_only_round_trips` | §18.2.5 |
| Print gridlines | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `com_e2e::xls_reader::xls_print_gridlines`, `page_setup_round_trip::print_gridlines_round_trips` | §18.3.1.70 |
| Print row/column headings | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `com_e2e::xls_reader::xls_print_headings`, `page_setup_round_trip::print_headings_round_trips` | §18.3.1.70 |
| Black and white printing | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.62 |
| Draft quality | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.62 |
| Comments printing option | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.62 |
| Cell errors display (blank/dash/#N/A) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.62 |
| Page order (downThenOver/overThenDown) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.62 |
| Header/footer: odd pages | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_header_footer_odd_only_defaults`, `page_setup_round_trip::odd_header_text_round_trips` | §18.3.1.46 |
| Header/footer: even pages | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_header_footer_even_first_and_flags` | §18.3.1.46 |
| Header/footer: first page different | R✔ W✔ | R✔ W✔ | R✔ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_header_footer_even_first_and_flags` | §18.3.1.46 |
| Header/footer: scaleWithDoc / alignWithMargins | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | §18.3.1.46 |
| Header/footer: formatting codes (bold, color, font) | R✔ W✔ | R✔ W✔ | R● W● | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_page_setup_and_header_footer`, `page_setup_round_trip::header_formatting_codes_round_trip` | §18.3.1.46 | XLS round-trips formatting codes verbatim through HEADER / FOOTER strings; semantic interpretation is reader-side only |
| Row page breaks | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::roundtrip_row_breaks`, `page_setup_round_trip::row_page_breaks_round_trip` | §18.3.1.74 |
| Column page breaks | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::roundtrip_col_breaks`, `page_setup_round_trip::col_page_breaks_round_trip` | §18.3.1.16 |

## Workbook and sheet protection

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Sheet protection flag | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::sheet_properties::sheet_protection`, `sheet_protection_round_trip::protected_sheet_round_trips_flag`, `excel_com_e2e::writing::test_write_sheet_protection`, `excel_com_e2e::writing_xls::excel_can_read_protection_state_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_protection_state_we_emit` | §18.3.1.79 | XLSX writer emits `<sheetProtection sheet="1" .../>`; reader parses it back and Excel preserves it through round-trip; XLSB writer emits BrtSheetProtection and Excel parity verified |
| Sheet protection: password hash | R✖ W✖ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `sheet_protection_round_trip::protected_sheet_with_password_hash_round_trips`, `excel_com_e2e::writing_xls::excel_can_read_protection_state_we_emit`, `excel_com_e2e::writing_xlsb::excel_can_read_protection_state_we_emit` | §18.3.1.79 | XLS PASSWORD record (0x0013) carries the 16-bit hash; survives Excel round-trip; XLSB BrtSheetProtection u16 password persists likewise |
| Sheet protection: specific permissions (objects/scenarios/formatCells/...) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.79 |
| Protected ranges (protectedRange) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.64 |
| Workbook structure protection | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.30 |
| Workbook window protection | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.30 |
| File-level encryption: OOXML Agile (AES + HMAC) | R✔ W✔ | R✖ W✖ | - | - | `encrypted_agile::excel_agile_xlsx_decrypts_with_correct_password`, `encrypted_agile_write::write_then_read_agile_round_trips_workbook_contents` | [MS-OFFCRYPTO] §2.3.4.10 |
| File-level encryption: OOXML Standard (AES-ECB) | R✔ W✔ | R✖ W✖ | - | - | `encrypted_agile::agile_xlsx_decrypts_with_correct_password`, `encrypted_standard_write::write_then_read_standard_round_trips_workbook_contents` | [MS-OFFCRYPTO] §2.3.4.5 |
| File-level encryption: OOXML Binary RC4 | R✔ W✖ | R✖ W✖ | - | - | `duke_sheets_crypto::ooxml::binary_rc4::tests::verify_password_round_trips` | [MS-OFFCRYPTO] §2.3.5 |
| File-level encryption: XLS RC4 CryptoAPI (SHA-1) | - | - | R✔ W✔ | - | `encrypted_rc4_cryptoapi::xls_rc4_cryptoapi_excel_decrypts_with_correct_password`, `encrypted_rc4_write::round_trip_via_xls_reader_rc4_cryptoapi_128` | [MS-OFFCRYPTO] §2.3.6.4 |
| File-level encryption: XLS RC4 legacy (MD5) | - | - | R✔ W✔ | - | `encrypted_rc4_cryptoapi::xls_rc4_legacy_decrypts_with_correct_password`, `encrypted_rc4_write::round_trip_via_xls_reader_rc4_legacy` | [MS-OFFCRYPTO] §2.3.6.2 |
| File-level encryption: XLS XOR Obfuscation | - | - | R✔ W● | - | `duke_sheets_crypto::xls::xor_obfuscation::tests::round_trip_one_record`, `encrypted_rc4_write::round_trip_via_xls_reader_xor_obfuscation` | [MS-OFFCRYPTO] §2.3.7 | Round-trips via own reader; not certified to interop with modern Excel (no reference XOR fixture) |
| VelvetSweatshop sentinel auto-decrypt | R✔ - | R✖ - | R✔ - | - | `encrypted_agile::velvet_sweatshop_auto_decrypts_when_enabled` | (well-known Excel password) |
| Write-reservation password | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.5.3 |

## Charts - types

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Bar: clustered | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_chart_bar` | §21.2.2.27 |
| Bar: stacked | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_bar_stacked` | §21.2.2.27 |
| Bar: 100% stacked | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_bar_percent_stacked` | §21.2.2.27 |
| Column: clustered/stacked/100% | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_column_stacked` | §21.2.2.27 |
| Line: basic | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_chart_line` | §21.2.2.48 |
| Line: stacked / 100% stacked | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_line_stacked` | §21.2.2.48 |
| Line: smooth | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_chart_series_smooth` | §21.2.2.48 |
| Pie | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_chart_pie` | §21.2.2.57 |
| Pie: exploded | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_pie_exploded` | §21.2.2.57 |
| Pie: 3D | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_3d_chart` | §21.2.2.58 |
| Doughnut | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §21.2.2.37 |
| Area: basic / stacked / 100% stacked | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_area_stacked` | §21.2.2.5 |
| Scatter: markers | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_chart_scatter` | §21.2.2.64 |
| Scatter: lines | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_scatter_lines` | §21.2.2.64 |
| Scatter: smooth lines | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_scatter_smooth` | §21.2.2.64 |
| Bubble | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_bubble` | §21.2.2.29 |
| Radar | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_radar` | §21.2.2.62 |
| Stock | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_stock` | §21.2.2.77 |
| Surface | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_surface` | §21.2.2.80 |
| 3D chart variants | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_3d_chart` | §21.2.2 |
| Combo (mixed bar + line) | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_combo_bar_line` | - |
| Combo: secondary axis | R✔ W✔ | R✔ W✔ | R✖ W✖ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_combo_secondary_axis_position` | - |
| ChartEx: Funnel | R● W● | R● W● | R- W- | R- W- | `chart_corpus::chart_corpus_chartex_read` | [MS-ODRAWXML] §5.22 | Reader only, write partial |
| ChartEx: Treemap | R● W● | R● W● | R- W- | R- W- | `chart_corpus::chart_corpus_chartex_read` | [MS-ODRAWXML] §5.22 |
| ChartEx: Sunburst | R● W● | R● W● | R- W- | R- W- | `chart_corpus::chart_corpus_chartex_read` | [MS-ODRAWXML] §5.22 |
| ChartEx: Waterfall | R● W● | R● W● | R- W- | R- W- | `chart_corpus::chart_corpus_chartex_read` | [MS-ODRAWXML] §5.22 |
| ChartEx: Box and Whisker | R● W● | R● W● | R- W- | R- W- | `chart_corpus::chart_corpus_chartex_read` | [MS-ODRAWXML] §5.22 |
| ChartEx: Histogram / Pareto | R● W● | R● W● | R- W- | R- W- | `chart_corpus::chart_corpus_chartex_read` | [MS-ODRAWXML] §5.22 |
| ChartEx: Region Map | R● W● | R● W● | R- W- | R- W- | `chart_corpus::chart_corpus_chartex_read` | [MS-ODRAWXML] §5.22 |
| ChartEx: Clustered Column | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | [MS-ODRAWXML] §5.22 |

## Charts - elements and options

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Chart title | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_bar` | §21.2.2.83 |
| Axis titles | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §21.2.2.26 |
| Legend (position, overlay) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_legend_positions` | §21.2.2.47 |
| Data labels (chart-wide) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_data_labels` | §21.2.2.34 |
| Data labels (per series) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_series_data_labels` | §21.2.2.34 |
| Data labels (per point) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_data_points` | §21.2.2.34 |
| Data label number format | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_data_label_number_format` | §21.2.2.34 |
| Data label leader lines | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_leader_lines` | §21.2.2.50 |
| Data label positions (inEnd/outEnd/ctr/etc.) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_data_label_positions` | §21.2.2.34 |
| Axis: min / max / scaling | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_axis_enhancements` | §21.2.2.65 |
| Axis: major / minor tick marks | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_axis_tick_marks_cross_none` | §21.2.2.52 |
| Axis: label positions | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_axis_label_positions` | §21.2.2.85 |
| Axis: date type | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_axis_date_type` | §21.2.2.32 |
| Axis: crossing at max / min | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_axis_crosses_min_max` | §21.2.2.30 |
| Axis: deleted | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_axis_delete` | §21.2.2.36 |
| Trendline: linear | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_trendline` | §21.2.2.93 |
| Trendline: polynomial / log / power / moving avg | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_trendline_polynomial` | §21.2.2.93 |
| Error bars (std error, percentage, stddev, custom) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_error_bars` | §21.2.2.42 |
| Markers (symbol types, sizes) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_marker_symbols` | §21.2.2.51 |
| Drop lines | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_drop_lines` | §21.2.2.38 |
| High-low lines | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_high_low_lines` | §21.2.2.43 |
| Series lines | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_series_lines` | §21.2.2.71 |
| Up/down bars | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_up_down_bars` | §21.2.2.96 |
| 3D view (rotX/rotY/perspective/depthPercent) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_view_3d` | §21.2.2.99 |
| Gap width / overlap | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_vary_colors_gap_overlap` | §21.2.2.41 |
| First slice angle / hole size | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_first_slice_angle_hole_size` | §21.2.2.41 |
| Chart layout (manual positioning) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_layout` | §21.2.2.55 |
| Chart styles (built-in style ID) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_style_color_passthrough` | §21.2.2.84 |
| Shape properties (fill, line) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_shape_properties`, `xlsx_roundtrip::test_roundtrip_chart_data_points`, `xlsx_roundtrip::test_roundtrip_chart_axis_enhancements`, `excel_com_e2e::chart_parity::chart_roundtrip_through_excel` | §21.2.2.72 |
| Invert-if-negative | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_invert_if_negative` | §21.2.2.45 |
| Data table displayed below chart | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chart_data_table` | §21.2.2.33 |
| Display blanks as (gap/zero/span) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_display_blanks_as_span` | §21.2.2.39 |
| Chartsheet (chart-only sheet) | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_chartsheet` | §18.3.1.12 |
| Multiple chartsheets | R✔ W✔ | R✔ W✔ | R- W- | R- W- | `xlsx_roundtrip::test_roundtrip_multiple_chartsheets` | §18.3.1.12 |

## Images and drawings

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Image parsing (PNG/JPEG) | R✔ W✔ | R✖ W✖ | R● W● | R✖ W✖ | `xlsx_roundtrip::xlsx_png_image_round_trips`, `duke-sheets-xls::pictures_round_trip::single_picture_round_trips`, `duke-sheets-xls::pictures_round_trip::single_jpeg_picture_round_trips`, `duke-sheets-xls::pictures_round_trip::single_bmp_picture_round_trips`, `duke-sheets-xls::pictures_round_trip::emf_blip_wrapper_round_trips_opaque_bytes_in_process`, `duke-sheets-xls::pictures_round_trip::wmf_blip_wrapper_round_trips_opaque_bytes_in_process`, `duke-sheets-xls::pictures_round_trip::tiff_blip_wrapper_round_trips_in_process`, `duke-sheets-xls::pictures_round_trip::gif_input_is_routed_through_png_blip_and_format_tag_flips`, `duke-sheets-xls::pictures_round_trip::picture_anchor_within_cell_offsets_round_trip`, `duke-sheets-xls::pictures_round_trip::xls_top_level_width_emu_does_not_survive_user_value`, `duke-sheets-xls::pictures_round_trip::rotation_and_flip_flags_round_trip`, `excel_com_e2e::writing::excel_can_read_xlsx_png_image_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_xls_png_image_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_xls_jpeg_image_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_xls_bmp_image_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_xls_picture_rotation_and_flip_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_xls_picture_and_comment_on_same_sheet_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_xls_pictures_across_multiple_sheets_we_emit` | §20.4 (DrawingML), MS-ODRAW | **PNG / JPEG / BMP** have Excel parity coverage. **EMF / WMF / TIFF** have in-process round-trip only (blip-wrapper framing verified; payload treated as opaque bytes; Excel converts metafiles + TIFF to PNG on SaveAs so no Excel parity is meaningful). **GIF** has no native Office binary blip variant — writer routes GIF input through the PNG blip path; bytes survive in-process but the format tag flips from Gif to Png on read. **EmbeddedImage.width_emu/height_emu** top-level fields are NOT preserved through the XLS writer; the reader synthesises them from the anchor cell range × default cell sizes (609,600 EMU/col, 190,500 EMU/row) — pinned by xls_top_level_width_emu_does_not_survive_user_value. PNG round-trip preserves within-cell EMU offsets at the 1024ths/256ths anchor quantisation granularity. Rotation (FOPT 0x0004) and flip H/V (FSP grfPersistence bits) round-trip exactly. |
| Image positioning (two-cell anchor) | R✔ W✔ | R✖ W✖ | R✔ W✔ | R✖ W✖ | `duke-sheets-xls::pictures_round_trip::single_picture_round_trips`, `xlsx_roundtrip::xlsx_png_image_round_trips` | §20.4.2.10 |
| Image positioning (one-cell anchor) | R✔ W✔ | R✖ W✖ | R● W● | R✖ W✖ | `xlsx_roundtrip::xlsx_onecell_anchor_round_trips`, `duke-sheets-xls::pictures_round_trip::onecell_anchor_round_trips_with_visual_area_preserved`, `excel_com_e2e::writing::excel_can_read_xlsx_onecell_picture_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_xls_onecell_image_we_emit` | §20.4.2.1 | **XLSX**: full support via `<xdr:oneCellAnchor>` with `<xdr:ext cx cy/>`; reader and writer agree on from cell + within-cell offsets + width/height. **XLS**: ClientAnchor has one byte layout shared across all OOXML anchor variants; OneCell input collapses to TwoCell with editAs=OneCell on read. Visual area (cells covered) is preserved via flag=2 (move only) and width/height encoded into the to-cell offset. |
| Image positioning (absolute anchor) | R✔ W✔ | R✖ W✖ | R● W● | R✖ W✖ | `xlsx_roundtrip::xlsx_absolute_anchor_round_trips`, `duke-sheets-xls::pictures_round_trip::absolute_anchor_round_trips_with_visual_area_preserved`, `excel_com_e2e::writing::excel_can_read_xlsx_absolute_picture_we_emit`, `excel_com_e2e::writing_xls::excel_can_read_xls_absolute_image_we_emit` | §20.4.2.3 | **XLSX**: full support via `<xdr:absoluteAnchor>` with `<xdr:pos x y/>` + `<xdr:ext cx cy/>`. **XLS**: collapses to TwoCell with editAs=Absolute on read; the x_emu/y_emu position lands the picture at the equivalent cell at default cell sizes. flag=3 (no move, no resize) signals the absolute semantics to readers. |
| Image editAs (move/size with cells) | R✔ W✔ | R✖ W✖ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::xlsx_twocell_anchor_editas_round_trips`, `duke-sheets-xls::pictures_round_trip::twocell_anchor_edit_as_round_trips`, `duke-sheets-xls::pictures_round_trip::onecell_anchor_round_trips_with_visual_area_preserved`, `duke-sheets-xls::pictures_round_trip::absolute_anchor_round_trips_with_visual_area_preserved` | §20.4.2.8 | **XLSX**: `editAs="..."` attribute on twoCellAnchor preserved. **XLS**: editAs preserved via the ClientAnchor flag bit: TwoCell=0, OneCell=2, Absolute=3. |
| SVG image support | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.3.18 |
| Shapes (rectangles, arrows, ...) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §20.1.2.2 |
| Text boxes | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §20.1.2.2 |
| SmartArt diagrams | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | §21.4 |
| WordArt | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §20.1.2.2 |
| Form controls (buttons, checkboxes, ...) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.26 |
| ActiveX controls | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.26 |
| OLE objects (embedded files) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.3.1.56 |

## Pivot tables

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Pivot cache (source data) | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::tests::test_writer_emits_pivot_table_and_cache_parts`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_table_source_pivot`, `duke-sheets-xlsx::writer::tests::test_writer_separates_pivot_caches_for_refresh_policy`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_external_pivot_source_definition`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_external_pivot_database_connection`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_olap_pivot_source_definition`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_scenario_pivot_source_definition`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_consolidation_pivot_source_ranges`, `duke-sheets-pivot::tests::refreshes_consolidation_source_ranges`, `duke-sheets-pivot::tests::shared_consolidation_sources_hit_the_internal_snapshot_cache`, `excel_com_e2e::writing::test_write_basic_pivot_survives_excel_roundtrip` | §18.10.1 CT_PivotCacheDefinition/CT_CacheSource/CT_Consolidation | Worksheet/table source cache definitions, generated cache records, cache-affecting refresh policy partitioning, non-refreshable external/scenario/OLAP cache-source metadata, and external database command metadata through workbook connections are writable; consolidation ranges, range names, and page labels round-trip, and worksheet/table/consolidation sources refresh locally with private snapshot caching. Scenario names, connection parameters/credentials, and full OLAP/data-model cache writing are not complete yet |
| Pivot table definition | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::tests::test_writer_emits_pivot_table_and_cache_parts`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_table_source_pivot`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_layout_flags`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_refresh_policy`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_advanced_filters`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_table_extensions`, `duke-sheets-xlsx::reader::pivot::tests::unsupported_pivot_filter_types_are_preserved`, `duke-sheets-pivot::tests::refreshes_compact_layout_hierarchy_without_column_fields`, `duke-sheets-pivot::tests::refreshes_compact_layout_hierarchy_with_column_fields`, `duke-sheets-pivot::tests::refreshes_tabular_layout_without_repeated_item_labels`, `duke-sheets-pivot::tests::refreshes_tabular_layout_with_repeated_item_labels`, `duke-sheets-pivot::tests::external_sources_are_marked_external_without_failing_refresh`, `duke-sheets-pivot::tests::external_sources_do_not_block_local_pivot_refresh`, `excel_com_e2e::writing::test_write_basic_pivot_survives_excel_roundtrip` | §18.10.1 CT_PivotTableDefinition/CT_PivotCacheDefinition | Basic semantic definitions round-trip: source, location, layout/grand-total/header/expand-collapse/print-title/page-area/merge-label/caption/error-missing-display flags, interaction/drop-zone/data-tip/empty-row-column display flags, refresh policy, fields, measures, style, item/label/value/top-N filters, unsupported advanced-filter diagnostics, and table-level extension payload preservation; compact row hierarchy and tabular/outline repeat-item-label refresh are supported; external-source pivots are marked externally managed during local refresh; advanced extension semantics and style/outline UI details are not complete yet |
| Row / column / value fields | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::tests::test_writer_emits_pivot_table_and_cache_parts`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_table_source_pivot`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_field_sort`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_axis_field_options`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_subtotal_functions`, `duke-sheets-pivot::tests::refreshes_row_field_subtotals`, `duke-sheets-pivot::tests::refreshes_column_field_subtotals`, `duke-sheets-pivot::tests::refreshes_custom_row_subtotal_function`, `duke-sheets-pivot::tests::refreshes_custom_column_subtotal_function`, `duke-sheets-pivot::tests::refreshes_show_empty_items_on_row_fields`, `duke-sheets-pivot::tests::refreshes_show_empty_items_on_column_fields`, `excel_com_e2e::writing::test_write_basic_pivot_survives_excel_roundtrip` | §18.10.1 CT_PivotField@sortType/showAll/subtotal flags | Basic row, column, and data fields with manual/ascending/descending item sort, show-empty-items, dropdown visibility, subtotal placement, blank-row/page-break, include-new-items, and item-page-count options round-trip; show-empty-items and row-/column-axis subtotal refresh with custom single subtotal functions are supported; hierarchies and OLAP fields are not supported yet |
| Filter (page) fields | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_page_fields`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_page_field_multi_select`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_layout_flags` | §18.10.1 CT_PageField/CT_PivotField | Basic report/page axis plus single-item and multi-item page filters and page-area wrapping round-trip; advanced page-field UI details are not complete yet |
| Aggregate functions (Sum/Count/Avg/...) | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::tests::test_writer_emits_pivot_table_and_cache_parts`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_table_source_pivot`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_show_as_percentages`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_show_as_base_field`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_rank_show_as`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_index_show_as`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_measure_number_format`, `duke-sheets-pivot::tests::refresh_applies_measure_number_format`, `excel_com_e2e::writing::test_write_basic_pivot_survives_excel_roundtrip` | §18.10.1 dataField@showDataAs/dataField@numFmtId, [MS-XLSX] x14 data field extension | Data-field subtotal functions, percent-of-row/column/grand-total/index, ECMA base-field difference/percent-difference/running-total, x14 rank show-as calculations, and measure number-format round-trip plus semantic refresh |
| Pivot table styles | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::tests::test_writer_emits_pivot_table_and_cache_parts`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_table_source_pivot`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_style_flags` | §18.10.1 CT_PivotTableStyle | Basic `pivotTableStyleInfo` name/header/stripe/last-column flags only |
| Calculated fields | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_calculated_fields`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_calculated_items`, `duke-sheets-pivot::tests::refreshes_calculated_field_measure` | §18.10.1 CT_CacheField, CT_CalculatedItem | Formula-backed cache fields round-trip and refresh as semantic source fields; calculated items round-trip as semantic cache metadata; calculated-item local refresh/evaluation and workbook/range references inside pivot formulas are not complete yet |
| Grouping (dates, numbers) | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::tests::test_writer_round_trips_pivot_grouping`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_multi_unit_date_grouping`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_manual_pivot_grouping`, `duke-sheets-pivot::tests::refreshes_single_unit_date_grouping`, `duke-sheets-pivot::tests::refreshes_multi_unit_date_grouping_hierarchy`, `duke-sheets-pivot::tests::refreshes_manual_item_grouping`, `duke-sheets-xlsx::pivot_compat::lo_can_open_manual_grouped_pivot`, `excel_com_e2e::writing::test_write_manual_pivot_grouping_survives_excel_roundtrip` | §18.10.1 CT_FieldGroup/CT_RangePr/CT_DiscretePr/CT_GroupItems | Basic numeric `rangePr` bins, single-unit date grouping, multi-unit date grouped-field hierarchies, and manual item `groupItems` round-trip with Excel SaveAs parity for manual groups; semantic refresh supports multi-unit date hierarchies and manual item groups; advanced grouping UI hierarchy details are not complete yet |
| Slicers | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | [MS-XLSX] §2.3.16 |
| Timelines | R✖ W✖ | R✖ W✖ | R- W- | R✖ W✖ | - | [MS-XLSX] §2.3.20 |
| PivotChart | R● W● | R✖ W✖ | R✖ W✖ | R✖ W✖ | `duke-sheets-xlsx::writer::chart::tests::test_pivot_source_roundtrip` | §21.3 CT_PivotSource | Basic chart-level `c:pivotSource` name/fmtId round-trips; generated PivotChart series, pivot formats, slicer/timeline chart state, and Excel UI parity are not complete yet |

## External workbook links

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| External workbook reference metadata | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.14 |
| Cached values from external books | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.14.4 |
| External-book formula parsing | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.17.2.3 |
| External named ranges | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.14.4 |
| OLE/DDE links | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.14.2 |
| Data connections (basic) | R● W● | R✖ W✖ | R- W- | R- W- | `duke-sheets-xlsx::writer::tests::test_writer_round_trips_external_pivot_database_connection`, `duke-sheets-xlsx::writer::tests::test_writer_round_trips_basic_non_database_connections` | §18.13 CT_Connections/CT_Connection/CT_DbPr/CT_OlapPr/CT_WebPr/CT_TextPr | Basic workbook `connections.xml` database, web, text, and OLAP connections with id/name, refresh flags, and core payload attributes round-trip; credentials, parameters, external ODC files, OLE DB/model extension details, and live refresh are not complete yet |
| Web queries | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.13 |
| Query tables | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.15 |

## Workbook calculation settings

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Calculation mode (auto/manual) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.2 |
| Iterate calculation enable | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.2 |
| Iteration count and delta | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.2 |
| Full precision flag | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.2 |
| A1 / R1C1 reference mode | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.2 |
| calcCompleted flag | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.2.2 |
| fullCalcOnLoad flag | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.2.2 |
| concurrentCalc flag | R✖ W✖ | R✖ W✖ | R- W- | R- W- | - | §18.2.2 |

## Workbook structure

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Multiple worksheets | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_multiple_sheets`, `skeleton_round_trip::multi_sheet_round_trips_with_all_names` | §18.2.20 |
| Sheet order preservation | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_interleaved_tab_order`, `skeleton_round_trip::multi_sheet_round_trips_with_all_names` | §18.2.20 |
| Special / XML-unsafe sheet names | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xlsx_roundtrip::test_roundtrip_xml_special_chars_in_sheet_names`, `skeleton_round_trip::special_and_unicode_sheet_names_round_trip` | §18.2.19 |
| Active sheet | R✔ W✔ | R✔ W✔ | R✔ W✔ | R✖ W✖ | `xls_e2e::sheet_properties::active_sheet`, `active_sheet_round_trip::middle_sheet_active_round_trips`, `active_sheet_round_trip::lo_can_read_active_sheet_we_emit` | §18.2.27 |
| Book view: window position / size | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.27 |
| Book view: tabRatio | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.27 |
| Book view: minimized / maximized | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | §18.2.27 |
| 1904 date system | R✖ W✖ | R✖ W✖ | R✖ W✖ | R- W- | - | §18.2.28 |

## Document properties

| Feature | XLSX | XLSB | XLS | ODS | Test | Spec |
|---------|------|------|-----|-----|------|------|
| Core properties (title, author, subject, keywords) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | [OPC] §11.1 |
| Extended properties (application, version, ...) | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | [OPC] §11.2 |
| Custom properties | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | [OPC] §11.3 |
| Last modified by / created | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | [OPC] §11.1 |
| Revision count | R✖ W✖ | R✖ W✖ | R✖ W✖ | R✖ W✖ | - | [OPC] §11.1 |

## MS-XLSX / OOXML extensions

Each row represents one [MS-XLSX] Appendix A extension namespace. These are the modern ("Excel 2014+", "Excel 2018+") extensions stored in `<extLst>` blocks.

| Feature | XLSX | XLSB | Test | Spec |
|---------|------|------|------|------|
| Dynamic array properties (2017) | R✔ W✔ | R✔ W✔ | `xlsx_roundtrip::roundtrip_dynamic_array_sequence` | [MS-XLSX] §2.6.3 |
| LAMBDA calc features (2018) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.9 |
| Threaded comments (2018) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.3.19 |
| Rich data / data types (2017) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.12 |
| Rich data web images (2020) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.14 |
| Rich data refresh (2020) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.15 |
| Pivot 2014/2017/2020/2022/2023 | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.3.* |
| Named sheet views (2019) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.3.17 |
| External link props (2019, 2021) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.3.14 |
| Python in Excel (2023) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.19 |
| Feature property bag (2022) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.18 |
| MSForms (2023) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.20 |
| External code service (2023, 2025) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.21 |
| Slicers (2010) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.3.16 |
| Timeline slicers (2012) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.3.20 |
| Sparklines | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.3.11 |
| Conditional formatting ext (x14) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.4 |
| Data validation ext (x14) | R✖ W✖ | R✖ W✖ | - | [MS-XLSX] §2.6.5 |

## File formats (top-level support)

| Feature | Read | Write | Notes |
|---------|------|-------|-------|
| XLSX (OOXML SpreadsheetML) | ✔ | ✔ | Primary format |
| XLSB (binary OOXML) | ✔ | ✔ | |
| XLS (BIFF8) | ✔ | ✖ | Read-only; writer is a 3-line stub |
| XLS (BIFF5, BIFF7) | ✖ | ✖ | Older binary variants |
| ODS (OpenDocument) | ✖ | ✖ | No implementation |
| CSV import/export | ✖ | ✖ | No implementation in duke-sheets (use std/csv crates) |
| Encrypted files (password-protected) | ✖ | ✖ | |
