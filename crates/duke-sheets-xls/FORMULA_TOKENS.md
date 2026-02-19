# BIFF8 Formula Token Parser — Implementation Checklist

This document is a comprehensive reference and checklist for implementing the
BIFF8 formula token (Ptg) parser. The goal is to decompile the raw RPN byte
stream stored in FORMULA records into human-readable formula text (e.g.,
`=SUM(A1:A10)`).

**Scope**: Read-only decompilation. We do NOT need to compile text→tokens or
evaluate tokens — we already have cached formula results from the FORMULA
record.

---

## Module Structure

```
crates/duke-sheets-xls/src/biff/formula/
  mod.rs            — public API: decompile(data, ctx) -> String
  ptg.rs            — token type enum (~39 base types + tAttr sub-types)
  function_table.rs — BIFF8 function index -> name (485 entries)
  token_parser.rs   — byte stream -> Vec<Token>
  decompiler.rs     — RPN token stack -> infix formula string
```

### Context struct needed by `decompile()`:

```rust
pub struct FormulaContext<'a> {
    /// Sheet names from BOUNDSHEET records (for 3D refs)
    pub sheet_names: &'a [String],
    /// EXTERNSHEET table: (sup_book_idx, first_sheet, last_sheet)
    pub extern_sheet: &'a [(u16, u16, u16)],
    /// Defined NAME records: (name, sheet_scope, formula_tokens)
    pub names: &'a [NameRecord],
}
```

---

## Phase 1 — MVP (~1,200 lines)

Covers: literals, all operators, cell/area refs, basic functions.
This phase handles ~90% of real-world formulas.

### Token Types

#### Unary/Binary Operators (unclassified, base 0x00–0x13)

| Byte | Name        | Description               | Size | Status |
|------|-------------|---------------------------|------|--------|
| 0x03 | tAdd        | Addition (+)              | 1    | [x]    |
| 0x04 | tSub        | Subtraction (-)           | 1    | [x]    |
| 0x05 | tMul        | Multiplication (*)        | 1    | [x]    |
| 0x06 | tDiv        | Division (/)              | 1    | [x]    |
| 0x07 | tPower      | Exponentiation (^)        | 1    | [x]    |
| 0x08 | tConcat     | Concatenation (&)         | 1    | [x]    |
| 0x09 | tLT         | Less than (<)             | 1    | [x]    |
| 0x0A | tLE         | Less than or equal (<=)   | 1    | [x]    |
| 0x0B | tEQ         | Equal (=)                 | 1    | [x]    |
| 0x0C | tGE         | Greater or equal (>=)     | 1    | [x]    |
| 0x0D | tGT         | Greater than (>)          | 1    | [x]    |
| 0x0E | tNE         | Not equal (<>)            | 1    | [x]    |
| 0x0F | tIsect      | Intersection (space)      | 1    | [x]    |
| 0x10 | tList       | Union (comma in refs)     | 1    | [x]    |
| 0x11 | tRange      | Range operator (:)        | 1    | [x]    |
| 0x12 | tUplus      | Unary plus (+)            | 1    | [x]    |
| 0x13 | tUminus     | Unary minus (-)           | 1    | [x]    |
| 0x14 | tPercent    | Percent (%)               | 1    | [x]    |
| 0x15 | tParen      | Parentheses (display)     | 1    | [x]    |

#### Constant Operands (unclassified, 0x16–0x1E)

| Byte | Name        | Description               | Size | Status |
|------|-------------|---------------------------|------|--------|
| 0x16 | tMissArg    | Missing argument          | 1    | [x]    |
| 0x17 | tStr        | String literal            | var  | [x]    |
| 0x19 | tAttr       | Attribute (sub-types)     | var  | [x]    |
| 0x1C | tErr        | Error constant            | 2    | [x]    |
| 0x1D | tBool       | Boolean constant          | 2    | [x]    |
| 0x1E | tInt        | 16-bit integer constant   | 3    | [x]    |
| 0x1F | tNum        | 64-bit float constant     | 9    | [x]    |

#### tAttr Sub-Types (byte 0x19)

The tAttr token has a flags byte at offset +1 that determines the sub-type:

| Flags | Name           | Description                          | Extra bytes | Status |
|-------|----------------|--------------------------------------|-------------|--------|
| 0x01  | tAttrVolatile  | Mark formula as volatile             | 2           | [x]    |
| 0x02  | tAttrIf        | IF function optimization             | 2           | [x]    |
| 0x04  | tAttrChoose    | CHOOSE function optimization         | var         | [x]    |
| 0x08  | tAttrSkip      | Skip bytes (after IF/CHOOSE branch)  | 2           | [x]    |
| 0x10  | tAttrSum       | Optimized SUM (SUM of single arg)    | 2           | [x]    |
| 0x20  | tAttrAssign    | Assign to name (macro sheets)        | 2           | [x]    |
| 0x40  | tAttrSpace     | Whitespace/formatting preservation   | 2           | [x]    |

#### Classified Operand Tokens (base 0x20–0x3F, with R/V/A variants)

These tokens have three class variants, determined by adding 0x00 (R),
0x20 (V), or 0x40 (A) to the base byte:

| Base | Name       | Description                  | Size (excl. byte) | Status |
|------|------------|------------------------------|--------------------|--------|
| 0x20 | tArray     | Array constant               | 7                  | [ ] P3 |
| 0x21 | tFunc      | Fixed-arg function call      | 2                  | [x]    |
| 0x22 | tFuncVar   | Variable-arg function call   | 3                  | [x]    |
| 0x23 | tName      | Defined name reference       | 4                  | [x]    |
| 0x24 | tRef       | Single cell reference        | 4                  | [x]    |
| 0x25 | tArea      | Cell range reference         | 8                  | [x]    |
| 0x26 | tMemArea   | Memory area (cached range)   | 6                  | [ ] P3 |
| 0x27 | tMemErr    | Erroneous memory area        | 6                  | [ ] P3 |
| 0x28 | tMemNoMem  | Non-cached memory area       | 6                  | [ ] P3 |
| 0x29 | tMemFunc   | Memory function              | 2                  | [ ] P3 |
| 0x2A | tRefErr    | Deleted cell reference       | 4                  | [x]    |
| 0x2B | tAreaErr   | Deleted area reference       | 8                  | [x]    |
| 0x2C | tRefN      | Relative ref (shared fmla)   | 4                  | [ ] P3 |
| 0x2D | tAreaN     | Relative area (shared fmla)  | 8                  | [ ] P3 |
| 0x39 | tNameX     | External name reference      | 6                  | [x]    |
| 0x3A | tRef3d     | 3D single cell reference     | 6                  | [x]    |
| 0x3B | tArea3d    | 3D cell range reference      | 10                 | [x]    |
| 0x3C | tRefErr3d  | Deleted 3D cell reference    | 6                  | [x]    |
| 0x3D | tAreaErr3d | Deleted 3D area reference    | 10                 | [x]    |

> **Class variants**: R (reference) = base+0x00, V (value) = base+0x20,
> A (array) = base+0x40. For decompilation purposes the class does not affect
> the output string — strip with `base = byte & 0x1F` when byte >= 0x20.

### Decompiler Stack Machine

The decompiler processes the RPN stream left-to-right, pushing/popping from
a string stack:

1. **Operand** (tRef, tNum, tStr, etc.) → push formatted string
2. **Unary operator** (tUminus, tPercent, etc.) → pop one, push result
3. **Binary operator** (tAdd, tMul, etc.) → pop two, push result with parens if needed
4. **tFunc** → pop N args (from function table), push `NAME(arg1,arg2,...)`
5. **tFuncVar** → pop N args (from token data), push `NAME(arg1,arg2,...)`
6. **tParen** → wrap top of stack in parens
7. **tAttr** sub-types → mostly no-ops for display (tAttrSpace adds whitespace)

### Cell Reference Encoding (BIFF8)

```
tRef (4 bytes after ptg byte):
  row:     u16  (bytes 0-1) — 0-based row index (0–65535)
  col_rw:  u16  (bytes 2-3) — bits 0-7: column (0-255)
                               bit 14: row is relative
                               bit 15: column is relative

tArea (8 bytes after ptg byte):
  first_row: u16, last_row: u16, first_col_rw: u16, last_col_rw: u16
```

For decompilation in non-shared formulas, relative flags are ignored — always
emit absolute-style A1 notation (e.g., `A1`, `B2:C10`).

### Phase 1 Implementation Steps

- [x] Create `crates/duke-sheets-xls/src/biff/formula/` directory
- [x] `mod.rs` — `pub fn decompile(data: &[u8], sheet_names: &[String]) -> String`
- [x] `ptg.rs` — token byte constants, `base_ptg()` class stripper, `token_data_size()`
- [x] `function_table.rs` — `fn function_name(idx: u16) -> &'static str` (all 485 entries)
- [x] `token_parser.rs` — `fn parse_tokens(data: &[u8]) -> Vec<ParsedToken>`
  - ParsedToken enum with all Phase 1 token types + Phase 2/3 stubs
  - Handles R/V/A class variants transparently
- [x] `decompiler.rs` — `fn decompile(tokens: &[ParsedToken], sheet_names: &[String]) -> String`
  - RPN stack machine → infix string with operator precedence
- [x] Hook into `reader.rs`: extract `cce` + token bytes from FORMULA record, call `decompile()`
- [x] Operator precedence table for minimal parenthesization
- [x] Fix `parse_formula_string` to preserve decompiled formula text

### Phase 1 Test Cases

- [x] Simple arithmetic: `=1+2`, `=A1*B1`, `=A1^2`
- [x] String literal: `="hello"&A1`
- [x] Boolean/error constants: `=TRUE`, `=IF(ISERROR(A1),#N/A,A1)`
- [x] Fixed-arg function: `=LEN(A1)`, `=MID(A1,2,3)`
- [x] Variable-arg function: `=SUM(A1:A10)`, `=IF(A1>0,1,0)`
- [x] Nested functions: `=SUM(LEN(A1),LEN(B1))`
- [x] Missing arg: `=MATCH(1,A1:A10,)`
- [x] Unary operators: `=-A1`, `=+A1`, `=50%`
- [x] Comparison operators: `=A1>=B1`, `=A1<>B1`
- [x] Parentheses: `=(A1+B1)*C1`
- [x] Integer and float constants: `=100`, `=3.14159`

---

## Phase 2 — 3D References & Names (~700 lines)

Covers: cross-sheet refs, defined names, external refs. Requires parsing
additional global records (EXTERNSHEET, NAME).

### New Global Records to Parse

- [x] **EXTERNSHEET** (0x0017) — Maps sheet reference indices to actual sheets.
  Format: `num_entries: u16`, then for each: `sup_book_idx: u16`,
  `first_sheet: u16`, `last_sheet: u16`.
- [x] **SUPBOOK** (0x01AE) — Supporting workbook record (self-ref, add-in, external).
- [x] **NAME** (0x0018/0x0218) — Defined name records.
  Format: flags, keyboard_shortcut, name_length, formula_length,
  sheet_index, name_string, formula_tokens.

### Token Types (Phase 2)

| Base | Name       | Parsing notes                                     | Status |
|------|------------|----------------------------------------------------|--------|
| 0x23 | tName      | 4 bytes: name_idx (u16, 1-based) + 2 reserved      | [x]    |
| 0x39 | tNameX     | 6 bytes: extern_sheet_idx (u16), name_idx (u16), 2 reserved | [x] |
| 0x3A | tRef3d     | 6 bytes: extern_sheet_idx (u16) + tRef data (4)     | [x]    |
| 0x3B | tArea3d    | 10 bytes: extern_sheet_idx (u16) + tArea data (8)   | [x]    |
| 0x3C | tRefErr3d  | 6 bytes: extern_sheet_idx (u16) + 4 ignored         | [x]    |
| 0x3D | tAreaErr3d | 10 bytes: extern_sheet_idx (u16) + 8 ignored        | [x]    |

### 3D Reference Formatting

```
extern_sheet_idx → (sup_book, first_sheet, last_sheet)
  - sup_book == self-reference (internal) → use sheet_names[]
  - first_sheet == last_sheet → "Sheet1!A1"
  - first_sheet != last_sheet → "Sheet1:Sheet3!A1:B10"
  - Sheet name with spaces/special chars → 'Sheet Name'!A1
```

### Populate Full Function Table (485 entries)

- [x] Copy full FTAB from calamine reference (MIT-licensed)
- [x] Include argc for variable-arg validation
- [ ] Map future functions (index >= 0x8000) → `_xlfn.NAME`

### Phase 2 Test Cases

- [x] Cross-sheet ref: `=Sheet2!A1` (E2E: `test_xls_formula_cross_sheet_ref`)
- [ ] Multi-sheet range: `=SUM(Sheet1:Sheet3!A1)` (unit test only — LO bridge can't create multi-sheet formulas)
- [x] Sheet name with spaces: `='My Sheet'!A1` (E2E: `test_xls_formula_cross_sheet_quoted_name`)
- [x] Defined name: `=MyRange*2` (E2E: `test_xls_formula_named_range`, `test_xls_formula_named_range_in_expression`)
- [ ] External name: `=[Book1.xlsx]Sheet1!A1` (unit test only — requires external workbook)

---

## Phase 3 — Edge Cases (~650 lines)

Covers: shared formulas, array constants, memory tokens, range operators.

### Token Types (Phase 3)

| Base | Name      | Notes                                               | Status |
|------|-----------|------------------------------------------------------|--------|
| 0x20 | tArray    | 7 bytes header (cols, rows, reserved). Actual array data appended AFTER the main token stream in "extra data" section. | [ ] |
| 0x26 | tMemArea  | 6 bytes: reserved (4) + subexpression_len (2). Skip subexpression_len bytes of sub-tokens. | [ ] |
| 0x27 | tMemErr   | 6 bytes: error_code (4) + subexpression_len (2). Emit #REF! | [ ] |
| 0x28 | tMemNoMem | 6 bytes: reserved (4) + subexpression_len (2). Same as tMemArea. | [ ] |
| 0x29 | tMemFunc  | 2 bytes: subexpression_len (2). The sub-tokens compute a reference. | [ ] |
| 0x2C | tRefN     | 4 bytes: like tRef but offsets are signed (relative to formula cell). For shared formulas, must adjust by cell position. | [ ] |
| 0x2D | tAreaN    | 8 bytes: like tArea but with signed offsets. | [ ] |

### Shared Formula Handling

BIFF8 SHAREDFMLA records store a single formula for a range of cells. The
formula uses tRefN/tAreaN with signed offsets relative to the top-left cell.
When decompiling for a specific cell at (row, col):

```
actual_row = formula_origin_row + offset_row
actual_col = formula_origin_col + offset_col
```

### Array Constant Encoding

tArray token body (in extra data section after main tokens):
```
cols:  u8   (0 = 256 columns)
rows:  u16  (0-based count, so 0 = 1 row)
Then for each element (col-major order):
  type_byte:
    0x00 → empty
    0x01 → f64 (8 bytes)
    0x02 → string (BIFF8 unicode string)
    0x04 → bool (u8)
    0x10 → error (u8 error code)
```

Output: `{1,2,3;4,5,6}` (semicolons separate rows, commas separate columns)

### Phase 3 Test Cases

- [ ] Array constant: `=SUM({1,2,3})`
- [ ] Shared formula with relative refs
- [ ] Range intersection: `=SUM(A1:C1 B1:D1)`
- [ ] Range union: `=SUM((A1:A5,C1:C5))`
- [ ] Memory tokens (passthrough)

---

## Function Table Reference

485 BIFF8 built-in functions. Index is the `iftab` field in tFunc/tFuncVar.

**Argument count encoding:**
- 0–253: exact fixed argument count
- 254: variable args, minimum from argc, actual count in tFuncVar token
- 255: variable args (0 or more)
- 128: paired variable args (COUNTIFS-style: 128 + min_pairs)
- 129: paired variable args with extra leading arg (SUMIFS-style)

### Full Function Table (index → name)

```
  0: COUNT           1: IF              2: ISNA            3: ISERROR
  4: SUM             5: AVERAGE         6: MIN             7: MAX
  8: ROW             9: COLUMN         10: NA             11: NPV
 12: STDEV          13: DOLLAR         14: FIXED          15: SIN
 16: COS            17: TAN            18: ATAN           19: PI
 20: SQRT           21: EXP            22: LN             23: LOG10
 24: ABS            25: INT            26: SIGN           27: ROUND
 28: LOOKUP         29: INDEX          30: REPT           31: MID
 32: LEN            33: VALUE          34: TRUE           35: FALSE
 36: AND            37: OR             38: NOT            39: MOD
 40: DCOUNT         41: DSUM           42: DAVERAGE       43: DMIN
 44: DMAX           45: DSTDEV         46: VAR            47: DVAR
 48: TEXT           49: LINEST         50: TREND          51: LOGEST
 52: GROWTH         53: GOTO           54: HALT           55: RETURN
 56: PV             57: FV             58: NPER           59: PMT
 60: RATE           61: MIRR           62: IRR            63: RAND
 64: MATCH          65: DATE           66: TIME           67: DAY
 68: MONTH          69: YEAR           70: WEEKDAY        71: HOUR
 72: MINUTE         73: SECOND         74: NOW            75: AREAS
 76: ROWS           77: COLUMNS        78: OFFSET         79: ABSREF
 80: RELREF         81: ARGUMENT       82: SEARCH         83: TRANSPOSE
 84: ERROR          85: STEP           86: TYPE           87: ECHO
 88: SET.NAME       89: CALLER         90: DEREF          91: WINDOWS
 92: SERIES         93: DOCUMENTS      94: ACTIVE.CELL    95: SELECTION
 96: RESULT         97: ATAN2          98: ASIN           99: ACOS
100: CHOOSE        101: HLOOKUP       102: VLOOKUP       103: LINKS
104: INPUT         105: ISREF         106: GET.FORMULA   107: GET.NAME
108: SET.VALUE     109: LOG           110: EXEC          111: CHAR
112: LOWER         113: UPPER         114: PROPER        115: LEFT
116: RIGHT         117: EXACT         118: TRIM          119: REPLACE
120: SUBSTITUTE    121: CODE          122: NAMES         123: DIRECTORY
124: FIND          125: CELL          126: ISERR         127: ISTEXT
128: ISNUMBER      129: ISBLANK       130: T             131: N
132: FOPEN         133: FCLOSE        134: FSIZE         135: FREADLN
136: FREAD         137: FWRITELN      138: FWRITE        139: FPOS
140: DATEVALUE     141: TIMEVALUE     142: SLN           143: SYD
144: DDB           145: GET.DEF       146: REFTEXT       147: TEXTREF
148: INDIRECT      149: REGISTER      150: CALL          151: ADD.BAR
152: ADD.MENU      153: ADD.COMMAND   154: ENABLE.COMMAND 155: CHECK.COMMAND
156: RENAME.COMMAND 157: SHOW.BAR     158: DELETE.MENU   159: DELETE.COMMAND
160: GET.CHART.ITEM 161: DIALOG.BOX   162: CLEAN         163: MDETERM
164: MINVERSE      165: MMULT         166: FILES         167: IPMT
168: PPMT          169: COUNTA        170: CANCEL.KEY    171: FOR
172: WHILE         173: BREAK         174: NEXT          175: INITIATE
176: REQUEST       177: POKE          178: EXECUTE       179: TERMINATE
180: RESTART       181: HELP          182: GET.BAR       183: PRODUCT
184: FACT          185: GET.CELL      186: GET.WORKSPACE 187: GET.WINDOW
188: GET.DOCUMENT  189: DPRODUCT      190: ISNONTEXT     191: GET.NOTE
192: NOTE          193: STDEVP        194: VARP          195: DSTDEVP
196: DVARP         197: TRUNC         198: ISLOGICAL     199: DCOUNTA
200: DELETE.BAR    201: UNREGISTER    202: (reserved)    203: (reserved)
204: USDOLLAR      205: FINDB         206: SEARCHB       207: REPLACEB
208: LEFTB         209: RIGHTB        210: MIDB          211: LENB
212: ROUNDUP       213: ROUNDDOWN     214: ASC           215: DBCS
216: RANK          217: (reserved)    218: (reserved)    219: ADDRESS
220: DAYS360       221: TODAY         222: VDB           223: ELSE
224: ELSE.IF       225: END.IF        226: FOR.CELL      227: MEDIAN
228: SUMPRODUCT    229: SINH          230: COSH          231: TANH
232: ASINH         233: ACOSH         234: ATANH         235: DGET
236: CREATE.OBJECT 237: VOLATILE      238: LAST.ERROR    239: CUSTOM.UNDO
240: CUSTOM.REPEAT 241: FORMULA.CONVERT 242: GET.LINK.INFO 243: TEXT.BOX
244: INFO          245: GROUP         246: GET.OBJECT    247: DB
248: PAUSE         249: (reserved)    250: (reserved)    251: RESUME
252: FREQUENCY     253: ADD.TOOLBAR   254: DELETE.TOOLBAR 255: User
256: RESET.TOOLBAR 257: EVALUATE      258: GET.TOOLBAR   259: GET.TOOL
260: SPELLING.CHECK 261: ERROR.TYPE   262: APP.TITLE     263: WINDOW.TITLE
264: SAVE.TOOLBAR  265: ENABLE.TOOL   266: PRESS.TOOL    267: REGISTER.ID
268: GET.WORKBOOK  269: AVEDEV        270: BETADIST      271: GAMMALN
272: BETAINV       273: BINOMDIST     274: CHIDIST       275: CHIINV
276: COMBIN        277: CONFIDENCE    278: CRITBINOM     279: EVEN
280: EXPONDIST     281: FDIST         282: FINV          283: FISHER
284: FISHERINV     285: FLOOR         286: GAMMADIST     287: GAMMAINV
288: CEILING       289: HYPGEOMDIST   290: LOGNORMDIST   291: LOGINV
292: NEGBINOMDIST  293: NORMDIST      294: NORMSDIST     295: NORMINV
296: NORMSINV      297: STANDARDIZE   298: ODD           299: PERMUT
300: POISSON       301: TDIST         302: WEIBULL       303: SUMXMY2
304: SUMX2MY2      305: SUMX2PY2      306: CHITEST       307: CORREL
308: COVAR         309: FORECAST      310: FTEST         311: INTERCEPT
312: PEARSON       313: RSQ           314: STEYX         315: SLOPE
316: TTEST         317: PROB          318: DEVSQ         319: GEOMEAN
320: HARMEAN       321: SUMSQ         322: KURT          323: SKEW
324: ZTEST         325: LARGE         326: SMALL         327: QUARTILE
328: PERCENTILE    329: PERCENTRANK   330: MODE          331: TRIMMEAN
332: TINV          333: (reserved)    334: MOVIE.COMMAND 335: GET.MOVIE
336: CONCATENATE   337: POWER         338: PIVOT.ADD.DATA 339: GET.PIVOT.TABLE
340: GET.PIVOT.FIELD 341: GET.PIVOT.ITEM 342: RADIANS     343: DEGREES
344: SUBTOTAL      345: SUMIF         346: COUNTIF       347: COUNTBLANK
348: SCENARIO.GET  349: OPTIONS.LISTS.GET 350: ISPMT      351: DATEDIF
352: DATESTRING    353: NUMBERSTRING  354: ROMAN         355: OPEN.DIALOG
356: SAVE.DIALOG   357: VIEW.GET      358: GETPIVOTDATA  359: HYPERLINK
360: PHONETIC      361: AVERAGEA      362: MAXA          363: MINA
364: STDEVPA       365: VARPA         366: STDEVA        367: VARA
368: BAHTTEXT      369: THAIDAYOFWEEK 370: THAIDIGIT     371: THAIMONTHOFYEAR
372: THAINUMSOUND  373: THAINUMSTRING 374: THAISTRINGLENGTH 375: ISTHAIDIGIT
376: ROUNDBAHTDOWN 377: ROUNDBAHTUP   378: THAIYEAR      379: RTD
380: CUBEVALUE     381: CUBEMEMBER    382: CUBEMEMBERPROPERTY 383: CUBERANKEDMEMBER
384: HEX2BIN       385: HEX2DEC       386: HEX2OCT       387: DEC2BIN
388: DEC2HEX       389: DEC2OCT       390: OCT2BIN       391: OCT2HEX
392: OCT2DEC       393: BIN2DEC       394: BIN2OCT       395: BIN2HEX
396: IMSUB         397: IMDIV         398: IMPOWER       399: IMABS
400: IMSQRT        401: IMLN          402: IMLOG2        403: IMLOG10
404: IMSIN         405: IMCOS         406: IMEXP         407: IMARGUMENT
408: IMCONJUGATE   409: IMAGINARY     410: IMREAL        411: COMPLEX
412: IMSUM         413: IMPRODUCT     414: SERIESSUM     415: FACTDOUBLE
416: SQRTPI        417: QUOTIENT      418: DELTA         419: GESTEP
420: ISEVEN        421: ISODD         422: MROUND        423: ERF
424: ERFC          425: BESSELJ       426: BESSELK       427: BESSELY
428: BESSELI       429: XIRR          430: XNPV          431: PRICEMAT
432: YIELDMAT      433: INTRATE       434: RECEIVED      435: DISC
436: PRICEDISC     437: YIELDDISC     438: TBILLEQ       439: TBILLPRICE
440: TBILLYIELD    441: PRICE         442: YIELD         443: DOLLARDE
444: DOLLARFR      445: NOMINAL       446: EFFECT        447: CUMPRINC
448: CUMIPMT       449: EDATE         450: EOMONTH       451: YEARFRAC
452: COUPDAYBS     453: COUPDAYS      454: COUPDAYSNC    455: COUPNCD
456: COUPNUM       457: COUPPCD       458: DURATION      459: MDURATION
460: ODDLPRICE     461: ODDLYIELD     462: ODDFPRICE     463: ODDFYIELD
464: RANDBETWEEN   465: WEEKNUM       466: AMORDEGRC     467: AMORLINC
468: CONVERT       469: ACCRINT       470: ACCRINTM      471: WORKDAY
472: NETWORKDAYS   473: GCD           474: MULTINOMIAL   475: LCM
476: FVSCHEDULE    477: CUBEKPIMEMBER 478: CUBESET       479: CUBESETCOUNT
480: IFERROR       481: COUNTIFS      482: SUMIFS        483: AVERAGEIF
484: AVERAGEIFS
```

### Argument Count Table

Values: 0-253 = fixed; 254 = variable (min from table); 255 = variable (0+);
128 = paired variable; 129 = paired variable + leading arg.

```
  0:255   1:  3   2:  1   3:  1   4:255   5:255   6:255   7:255
  8:  1   9:  1  10:  0  11:254  12:255  13:  2  14:  3  15:  1
 16:  1  17:  1  18:  1  19:  0  20:  1  21:  1  22:  1  23:  1
 24:  1  25:  1  26:  1  27:  2  28:  3  29:  4  30:  2  31:  3
 32:  1  33:  1  34:  0  35:  0  36:255  37:255  38:  1  39:  2
 40:  3  41:  3  42:  3  43:  3  44:  3  45:  3  46:255  47:  3
 48:  2  49:  4  50:  4  51:  4  52:  4  53:  1  54:  1  55:  1
 56:  5  57:  5  58:  5  59:  5  60:  6  61:  3  62:  2  63:  0
 64:  3  65:  3  66:  3  67:  1  68:  1  69:  1  70:  2  71:  1
 72:  1  73:  1  74:  0  75:  1  76:  1  77:  1  78:  5  79:  2
 80:  2  81:  3  82:  3  83:  1  84:  2  85:  0  86:  1  87:  1
 88:  2  89:  0  90:  1  91:  2  92:  2  93:  2  94:  0  95:  0
 96:  1  97:  2  98:  1  99:  1 100:255 101:  4 102:  4 103:  2
104:  7 105:  1 106:  1 107:  2 108:  2 109:  2 110:  4 111:  1
112:  1 113:  1 114:  1 115:  2 116:  2 117:  2 118:  1 119:  4
120:  4 121:  1 122:  3 123:  1 124:  3 125:  2 126:  1 127:  1
128:  1 129:  1 130:  1 131:  1 132:  2 133:  1 134:  1 135:  1
136:  2 137:  2 138:  2 139:  2 140:  1 141:  1 142:  3 143:  4
144:  5 145:  3 146:  2 147:  2 148:  2 149:255 150:255 151:  1
152:  4 153:  5 154:  5 155:  5 156:  5 157:  1 158:  3 159:  4
160:  3 161:  1 162:  1 163:  1 164:  1 165:  1 166:  2 167:  6
168:  6 169:255 170:  2 171:  4 172:  1 173:  0 174:  0 175:  2
176:  2 177:  3 178:  2 179:  1 180:  1 181:  1 182:  4 183:255
184:  1 185:  2 186:  1 187:  2 188:  2 189:  3 190:  1 191:  3
192:  4 193:255 194:255 195:  3 196:  3 197:  2 198:  1 199:  3
200:  1 201:  1 202:  0 203:  0 204:  2 205:  3 206:  3 207:  4
208:  2 209:  2 210:  3 211:  3 212:  2 213:  2 214:  1 215:  1
216:  3 217:  0 218:  0 219:  5 220:  3 221:  0 222:  7 223:  0
224:  1 225:  0 226:  3 227:255 228:255 229:  1 230:  1 231:  1
232:  1 233:  1 234:  1 235:  3 236: 11 237:  1 238:  0 239:  2
240:  3 241:  5 242:  4 243:  4 244:  1 245:  0 246:  5 247:  5
248:  1 249:  0 250:  0 251:  1 252:  2 253:  2 254:  1 255:255
256:  1 257:  1 258:  2 259:  3 260:  3 261:  1 262:  1 263:  1
264:  2 265:  3 266:  3 267:  3 268:  2 269:255 270:  5 271:  1
272:  5 273:  4 274:  2 275:  2 276:  2 277:  3 278:  3 279:  1
280:  3 281:  3 282:  3 283:  1 284:  1 285:  2 286:  4 287:  3
288:  2 289:  4 290:  3 291:  3 292:  3 293:  4 294:  1 295:  3
296:  1 297:  3 298:  1 299:  2 300:  3 301:  3 302:  4 303:  2
304:  2 305:  2 306:  2 307:  2 308:  2 309:  3 310:  2 311:  2
312:  2 313:  2 314:  2 315:  2 316:  4 317:  4 318:255 319:255
320:255 321:255 322:255 323:255 324:  3 325:  2 326:  2 327:  2
328:  2 329:  3 330:255 331:  2 332:  2 333:  4 334:  4 335:  3
336:255 337:  2 338:  9 339:  2 340:  3 341:  4 342:  1 343:  1
344:255 345:  3 346:  2 347:  1 348:  2 349:  1 350:  4 351:  3
352:  1 353:  2 354:  2 355:  4 356:  5 357:  2 358:128 359:  2
360:  1 361:255 362:255 363:255 364:255 365:255 366:255 367:255
368:  1 369:  1 370:  1 371:  1 372:  1 373:  1 374:  1 375:  1
376:  1 377:  1 378:  1 379:255 380:255 381:  3 382:  3 383:  4
384:  2 385:  1 386:  2 387:  2 388:  2 389:  2 390:  2 391:  2
392:  1 393:  1 394:  2 395:  2 396:  2 397:  2 398:  2 399:  1
400:  1 401:  1 402:  1 403:  1 404:  1 405:  1 406:  1 407:  1
408:  1 409:  1 410:  1 411:  3 412:255 413:255 414:  4 415:  1
416:  1 417:  2 418:  2 419:  2 420:  1 421:  1 422:  2 423:  2
424:  1 425:  2 426:  2 427:  2 428:  2 429:  3 430:  3 431:  6
432:  6 433:  5 434:  5 435:  5 436:  5 437:  5 438:  3 439:  3
440:  3 441:  7 442:  7 443:  2 444:  2 445:  2 446:  2 447:  6
448:  6 449:  2 450:  2 451:  3 452:  4 453:  4 454:  4 455:  4
456:  4 457:  4 458:  6 459:  6 460:  8 461:  8 462:  8 463:  8
464:  2 465:  2 466:  7 467:  7 468:  8 469:  8 470:  5 471:  3
472:  3 473:255 474:255 475:255 476:  2 477:  4 478:  5 479:  1
480:  2 481:128 482:129 483:  3 484:129
```

---

## Integration Checklist

- [x] Parse EXTERNSHEET record in globals section of reader.rs
- [x] Parse SUPBOOK records in globals section of reader.rs
- [x] Parse NAME records in globals section of reader.rs
- [x] Build `FormulaContext` from parsed globals
- [x] In FORMULA record handler: call `decompile()` with `FormulaContext`
- [x] Handle decompile failures gracefully — fall back to empty string, log warning
- [ ] Remove "formula text partial" known issue from TODO.md when Phase 3 complete
- [x] Add `formula` module to `biff/mod.rs` exports

---

## Effort Estimates

| Phase | Lines (est.) | Tokens covered | Real-world coverage |
|-------|-------------|----------------|---------------------|
| 1     | ~1,200      | ~30            | ~90%                |
| 2     | ~700        | ~6             | ~98%                |
| 3     | ~650        | ~7             | ~100%               |
| **Total** | **~2,550** | **~43**     | —                   |

---

## References

- [MS-XLS] §2.5.198 — Rgce (formula token stream)
- [MS-XLS] §2.5.198.1–198.93 — Individual Ptg definitions
- [MS-XLS] §2.4.168 — NAME record
- [MS-XLS] §2.4.106 — EXTERNSHEET record
- calamine `src/utils.rs` — FTAB/FTAB_ARGC arrays (MIT license)
- calamine `src/xls/mod.rs` — reference parser implementation
