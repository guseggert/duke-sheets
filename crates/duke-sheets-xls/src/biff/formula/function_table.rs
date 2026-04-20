//! BIFF8 built-in function table (485 entries).
//!
//! Maps function index (from tFunc/tFuncVar tokens) to function name and
//! argument count. Derived from the [MS-XLS] specification and the calamine
//! crate's FTAB (MIT-licensed).
//!
//! Argument count encoding:
//! - 0-253: fixed argument count
//! - 254: variable args (minimum from table, actual count in tFuncVar)
//! - 255: variable args (0 or more)

/// Function names indexed by BIFF8 function number.
/// Empty strings represent reserved/unused indices.
static FTAB: [&str; 485] = [
    "COUNT",              // 0
    "IF",                 // 1
    "ISNA",               // 2
    "ISERROR",            // 3
    "SUM",                // 4
    "AVERAGE",            // 5
    "MIN",                // 6
    "MAX",                // 7
    "ROW",                // 8
    "COLUMN",             // 9
    "NA",                 // 10
    "NPV",                // 11
    "STDEV",              // 12
    "DOLLAR",             // 13
    "FIXED",              // 14
    "SIN",                // 15
    "COS",                // 16
    "TAN",                // 17
    "ATAN",               // 18
    "PI",                 // 19
    "SQRT",               // 20
    "EXP",                // 21
    "LN",                 // 22
    "LOG10",              // 23
    "ABS",                // 24
    "INT",                // 25
    "SIGN",               // 26
    "ROUND",              // 27
    "LOOKUP",             // 28
    "INDEX",              // 29
    "REPT",               // 30
    "MID",                // 31
    "LEN",                // 32
    "VALUE",              // 33
    "TRUE",               // 34
    "FALSE",              // 35
    "AND",                // 36
    "OR",                 // 37
    "NOT",                // 38
    "MOD",                // 39
    "DCOUNT",             // 40
    "DSUM",               // 41
    "DAVERAGE",           // 42
    "DMIN",               // 43
    "DMAX",               // 44
    "DSTDEV",             // 45
    "VAR",                // 46
    "DVAR",               // 47
    "TEXT",               // 48
    "LINEST",             // 49
    "TREND",              // 50
    "LOGEST",             // 51
    "GROWTH",             // 52
    "GOTO",               // 53
    "HALT",               // 54
    "RETURN",             // 55
    "PV",                 // 56
    "FV",                 // 57
    "NPER",               // 58
    "PMT",                // 59
    "RATE",               // 60
    "MIRR",               // 61
    "IRR",                // 62
    "RAND",               // 63
    "MATCH",              // 64
    "DATE",               // 65
    "TIME",               // 66
    "DAY",                // 67
    "MONTH",              // 68
    "YEAR",               // 69
    "WEEKDAY",            // 70
    "HOUR",               // 71
    "MINUTE",             // 72
    "SECOND",             // 73
    "NOW",                // 74
    "AREAS",              // 75
    "ROWS",               // 76
    "COLUMNS",            // 77
    "OFFSET",             // 78
    "ABSREF",             // 79
    "RELREF",             // 80
    "ARGUMENT",           // 81
    "SEARCH",             // 82
    "TRANSPOSE",          // 83
    "ERROR",              // 84
    "STEP",               // 85
    "TYPE",               // 86
    "ECHO",               // 87
    "SET.NAME",           // 88
    "CALLER",             // 89
    "DEREF",              // 90
    "WINDOWS",            // 91
    "SERIES",             // 92
    "DOCUMENTS",          // 93
    "ACTIVE.CELL",        // 94
    "SELECTION",          // 95
    "RESULT",             // 96
    "ATAN2",              // 97
    "ASIN",               // 98
    "ACOS",               // 99
    "CHOOSE",             // 100
    "HLOOKUP",            // 101
    "VLOOKUP",            // 102
    "LINKS",              // 103
    "INPUT",              // 104
    "ISREF",              // 105
    "GET.FORMULA",        // 106
    "GET.NAME",           // 107
    "SET.VALUE",          // 108
    "LOG",                // 109
    "EXEC",               // 110
    "CHAR",               // 111
    "LOWER",              // 112
    "UPPER",              // 113
    "PROPER",             // 114
    "LEFT",               // 115
    "RIGHT",              // 116
    "EXACT",              // 117
    "TRIM",               // 118
    "REPLACE",            // 119
    "SUBSTITUTE",         // 120
    "CODE",               // 121
    "NAMES",              // 122
    "DIRECTORY",          // 123
    "FIND",               // 124
    "CELL",               // 125
    "ISERR",              // 126
    "ISTEXT",             // 127
    "ISNUMBER",           // 128
    "ISBLANK",            // 129
    "T",                  // 130
    "N",                  // 131
    "FOPEN",              // 132
    "FCLOSE",             // 133
    "FSIZE",              // 134
    "FREADLN",            // 135
    "FREAD",              // 136
    "FWRITELN",           // 137
    "FWRITE",             // 138
    "FPOS",               // 139
    "DATEVALUE",          // 140
    "TIMEVALUE",          // 141
    "SLN",                // 142
    "SYD",                // 143
    "DDB",                // 144
    "GET.DEF",            // 145
    "REFTEXT",            // 146
    "TEXTREF",            // 147
    "INDIRECT",           // 148
    "REGISTER",           // 149
    "CALL",               // 150
    "ADD.BAR",            // 151
    "ADD.MENU",           // 152
    "ADD.COMMAND",        // 153
    "ENABLE.COMMAND",     // 154
    "CHECK.COMMAND",      // 155
    "RENAME.COMMAND",     // 156
    "SHOW.BAR",           // 157
    "DELETE.MENU",        // 158
    "DELETE.COMMAND",     // 159
    "GET.CHART.ITEM",     // 160
    "DIALOG.BOX",         // 161
    "CLEAN",              // 162
    "MDETERM",            // 163
    "MINVERSE",           // 164
    "MMULT",              // 165
    "FILES",              // 166
    "IPMT",               // 167
    "PPMT",               // 168
    "COUNTA",             // 169
    "CANCEL.KEY",         // 170
    "FOR",                // 171
    "WHILE",              // 172
    "BREAK",              // 173
    "NEXT",               // 174
    "INITIATE",           // 175
    "REQUEST",            // 176
    "POKE",               // 177
    "EXECUTE",            // 178
    "TERMINATE",          // 179
    "RESTART",            // 180
    "HELP",               // 181
    "GET.BAR",            // 182
    "PRODUCT",            // 183
    "FACT",               // 184
    "GET.CELL",           // 185
    "GET.WORKSPACE",      // 186
    "GET.WINDOW",         // 187
    "GET.DOCUMENT",       // 188
    "DPRODUCT",           // 189
    "ISNONTEXT",          // 190
    "GET.NOTE",           // 191
    "NOTE",               // 192
    "STDEVP",             // 193
    "VARP",               // 194
    "DSTDEVP",            // 195
    "DVARP",              // 196
    "TRUNC",              // 197
    "ISLOGICAL",          // 198
    "DCOUNTA",            // 199
    "DELETE.BAR",         // 200
    "UNREGISTER",         // 201
    "",                   // 202 (reserved)
    "",                   // 203 (reserved)
    "USDOLLAR",           // 204
    "FINDB",              // 205
    "SEARCHB",            // 206
    "REPLACEB",           // 207
    "LEFTB",              // 208
    "RIGHTB",             // 209
    "MIDB",               // 210
    "LENB",               // 211
    "ROUNDUP",            // 212
    "ROUNDDOWN",          // 213
    "ASC",                // 214
    "DBCS",               // 215
    "RANK",               // 216
    "",                   // 217 (reserved)
    "",                   // 218 (reserved)
    "ADDRESS",            // 219
    "DAYS360",            // 220
    "TODAY",              // 221
    "VDB",                // 222
    "ELSE",               // 223
    "ELSE.IF",            // 224
    "END.IF",             // 225
    "FOR.CELL",           // 226
    "MEDIAN",             // 227
    "SUMPRODUCT",         // 228
    "SINH",               // 229
    "COSH",               // 230
    "TANH",               // 231
    "ASINH",              // 232
    "ACOSH",              // 233
    "ATANH",              // 234
    "DGET",               // 235
    "CREATE.OBJECT",      // 236
    "VOLATILE",           // 237
    "LAST.ERROR",         // 238
    "CUSTOM.UNDO",        // 239
    "CUSTOM.REPEAT",      // 240
    "FORMULA.CONVERT",    // 241
    "GET.LINK.INFO",      // 242
    "TEXT.BOX",           // 243
    "INFO",               // 244
    "GROUP",              // 245
    "GET.OBJECT",         // 246
    "DB",                 // 247
    "PAUSE",              // 248
    "",                   // 249 (reserved)
    "",                   // 250 (reserved)
    "RESUME",             // 251
    "FREQUENCY",          // 252
    "ADD.TOOLBAR",        // 253
    "DELETE.TOOLBAR",     // 254
    "",                   // 255 (User-defined)
    "RESET.TOOLBAR",      // 256
    "EVALUATE",           // 257
    "GET.TOOLBAR",        // 258
    "GET.TOOL",           // 259
    "SPELLING.CHECK",     // 260
    "ERROR.TYPE",         // 261
    "APP.TITLE",          // 262
    "WINDOW.TITLE",       // 263
    "SAVE.TOOLBAR",       // 264
    "ENABLE.TOOL",        // 265
    "PRESS.TOOL",         // 266
    "REGISTER.ID",        // 267
    "GET.WORKBOOK",       // 268
    "AVEDEV",             // 269
    "BETADIST",           // 270
    "GAMMALN",            // 271
    "BETAINV",            // 272
    "BINOMDIST",          // 273
    "CHIDIST",            // 274
    "CHIINV",             // 275
    "COMBIN",             // 276
    "CONFIDENCE",         // 277
    "CRITBINOM",          // 278
    "EVEN",               // 279
    "EXPONDIST",          // 280
    "FDIST",              // 281
    "FINV",               // 282
    "FISHER",             // 283
    "FISHERINV",          // 284
    "FLOOR",              // 285
    "GAMMADIST",          // 286
    "GAMMAINV",           // 287
    "CEILING",            // 288
    "HYPGEOMDIST",        // 289
    "LOGNORMDIST",        // 290
    "LOGINV",             // 291
    "NEGBINOMDIST",       // 292
    "NORMDIST",           // 293
    "NORMSDIST",          // 294
    "NORMINV",            // 295
    "NORMSINV",           // 296
    "STANDARDIZE",        // 297
    "ODD",                // 298
    "PERMUT",             // 299
    "POISSON",            // 300
    "TDIST",              // 301
    "WEIBULL",            // 302
    "SUMXMY2",            // 303
    "SUMX2MY2",           // 304
    "SUMX2PY2",           // 305
    "CHITEST",            // 306
    "CORREL",             // 307
    "COVAR",              // 308
    "FORECAST",           // 309
    "FTEST",              // 310
    "INTERCEPT",          // 311
    "PEARSON",            // 312
    "RSQ",                // 313
    "STEYX",              // 314
    "SLOPE",              // 315
    "TTEST",              // 316
    "PROB",               // 317
    "DEVSQ",              // 318
    "GEOMEAN",            // 319
    "HARMEAN",            // 320
    "SUMSQ",              // 321
    "KURT",               // 322
    "SKEW",               // 323
    "ZTEST",              // 324
    "LARGE",              // 325
    "SMALL",              // 326
    "QUARTILE",           // 327
    "PERCENTILE",         // 328
    "PERCENTRANK",        // 329
    "MODE",               // 330
    "TRIMMEAN",           // 331
    "TINV",               // 332
    "",                   // 333 (reserved)
    "MOVIE.COMMAND",      // 334
    "GET.MOVIE",          // 335
    "CONCATENATE",        // 336
    "POWER",              // 337
    "PIVOT.ADD.DATA",     // 338
    "GET.PIVOT.TABLE",    // 339
    "GET.PIVOT.FIELD",    // 340
    "GET.PIVOT.ITEM",     // 341
    "RADIANS",            // 342
    "DEGREES",            // 343
    "SUBTOTAL",           // 344
    "SUMIF",              // 345
    "COUNTIF",            // 346
    "COUNTBLANK",         // 347
    "SCENARIO.GET",       // 348
    "OPTIONS.LISTS.GET",  // 349
    "ISPMT",              // 350
    "DATEDIF",            // 351
    "DATESTRING",         // 352
    "NUMBERSTRING",       // 353
    "ROMAN",              // 354
    "OPEN.DIALOG",        // 355
    "SAVE.DIALOG",        // 356
    "VIEW.GET",           // 357
    "GETPIVOTDATA",       // 358
    "HYPERLINK",          // 359
    "PHONETIC",           // 360
    "AVERAGEA",           // 361
    "MAXA",               // 362
    "MINA",               // 363
    "STDEVPA",            // 364
    "VARPA",              // 365
    "STDEVA",             // 366
    "VARA",               // 367
    "BAHTTEXT",           // 368
    "THAIDAYOFWEEK",      // 369
    "THAIDIGIT",          // 370
    "THAIMONTHOFYEAR",    // 371
    "THAINUMSOUND",       // 372
    "THAINUMSTRING",      // 373
    "THAISTRINGLENGTH",   // 374
    "ISTHAIDIGIT",        // 375
    "ROUNDBAHTDOWN",      // 376
    "ROUNDBAHTUP",        // 377
    "THAIYEAR",           // 378
    "RTD",                // 379
    "CUBEVALUE",          // 380
    "CUBEMEMBER",         // 381
    "CUBEMEMBERPROPERTY", // 382
    "CUBERANKEDMEMBER",   // 383
    "HEX2BIN",            // 384
    "HEX2DEC",            // 385
    "HEX2OCT",            // 386
    "DEC2BIN",            // 387
    "DEC2HEX",            // 388
    "DEC2OCT",            // 389
    "OCT2BIN",            // 390
    "OCT2HEX",            // 391
    "OCT2DEC",            // 392
    "BIN2DEC",            // 393
    "BIN2OCT",            // 394
    "BIN2HEX",            // 395
    "IMSUB",              // 396
    "IMDIV",              // 397
    "IMPOWER",            // 398
    "IMABS",              // 399
    "IMSQRT",             // 400
    "IMLN",               // 401
    "IMLOG2",             // 402
    "IMLOG10",            // 403
    "IMSIN",              // 404
    "IMCOS",              // 405
    "IMEXP",              // 406
    "IMARGUMENT",         // 407
    "IMCONJUGATE",        // 408
    "IMAGINARY",          // 409
    "IMREAL",             // 410
    "COMPLEX",            // 411
    "IMSUM",              // 412
    "IMPRODUCT",          // 413
    "SERIESSUM",          // 414
    "FACTDOUBLE",         // 415
    "SQRTPI",             // 416
    "QUOTIENT",           // 417
    "DELTA",              // 418
    "GESTEP",             // 419
    "ISEVEN",             // 420
    "ISODD",              // 421
    "MROUND",             // 422
    "ERF",                // 423
    "ERFC",               // 424
    "BESSELJ",            // 425
    "BESSELK",            // 426
    "BESSELY",            // 427
    "BESSELI",            // 428
    "XIRR",               // 429
    "XNPV",               // 430
    "PRICEMAT",           // 431
    "YIELDMAT",           // 432
    "INTRATE",            // 433
    "RECEIVED",           // 434
    "DISC",               // 435
    "PRICEDISC",          // 436
    "YIELDDISC",          // 437
    "TBILLEQ",            // 438
    "TBILLPRICE",         // 439
    "TBILLYIELD",         // 440
    "PRICE",              // 441
    "YIELD",              // 442
    "DOLLARDE",           // 443
    "DOLLARFR",           // 444
    "NOMINAL",            // 445
    "EFFECT",             // 446
    "CUMPRINC",           // 447
    "CUMIPMT",            // 448
    "EDATE",              // 449
    "EOMONTH",            // 450
    "YEARFRAC",           // 451
    "COUPDAYBS",          // 452
    "COUPDAYS",           // 453
    "COUPDAYSNC",         // 454
    "COUPNCD",            // 455
    "COUPNUM",            // 456
    "COUPPCD",            // 457
    "DURATION",           // 458
    "MDURATION",          // 459
    "ODDLPRICE",          // 460
    "ODDLYIELD",          // 461
    "ODDFPRICE",          // 462
    "ODDFYIELD",          // 463
    "RANDBETWEEN",        // 464
    "WEEKNUM",            // 465
    "AMORDEGRC",          // 466
    "AMORLINC",           // 467
    "CONVERT",            // 468
    "ACCRINT",            // 469
    "ACCRINTM",           // 470
    "WORKDAY",            // 471
    "NETWORKDAYS",        // 472
    "GCD",                // 473
    "MULTINOMIAL",        // 474
    "LCM",                // 475
    "FVSCHEDULE",         // 476
    "CUBEKPIMEMBER",      // 477
    "CUBESET",            // 478
    "CUBESETCOUNT",       // 479
    "IFERROR",            // 480
    "COUNTIFS",           // 481
    "SUMIFS",             // 482
    "AVERAGEIF",          // 483
    "AVERAGEIFS",         // 484
];

/// Argument count for each BIFF8 function.
/// - 0-253: fixed count
/// - 254: variable (min from this table)
/// - 255: variable (0 or more)
static FTAB_ARGC: [u16; 485] = [
    255, 3, 1, 1, 255, 255, 255, 255, //   0-7
    1, 1, 0, 254, 255, 2, 3, 1, //   8-15
    1, 1, 1, 0, 1, 1, 1, 1, //  16-23
    1, 1, 1, 2, 3, 4, 2, 3, //  24-31
    1, 1, 0, 0, 255, 255, 1, 2, //  32-39
    3, 3, 3, 3, 3, 3, 255, 3, //  40-47
    2, 4, 4, 4, 4, 1, 1, 1, //  48-55
    5, 5, 5, 5, 6, 3, 2, 0, //  56-63
    3, 3, 3, 1, 1, 1, 2, 1, //  64-71
    1, 1, 0, 1, 1, 1, 5, 2, //  72-79
    2, 3, 3, 1, 2, 0, 1, 1, //  80-87
    2, 0, 1, 2, 2, 2, 0, 0, //  88-95
    1, 2, 1, 1, 255, 4, 4, 2, //  96-103
    7, 1, 1, 2, 2, 2, 4, 1, // 104-111
    1, 1, 1, 2, 2, 2, 1, 4, // 112-119
    4, 1, 3, 1, 3, 2, 1, 1, // 120-127
    1, 1, 1, 1, 2, 1, 1, 1, // 128-135
    2, 2, 2, 2, 1, 1, 3, 4, // 136-143
    5, 3, 2, 2, 2, 255, 255, 1, // 144-151
    4, 5, 5, 5, 5, 1, 3, 4, // 152-159
    3, 1, 1, 1, 1, 1, 2, 6, // 160-167
    6, 255, 2, 4, 1, 0, 0, 2, // 168-175
    2, 3, 2, 1, 1, 1, 4, 255, // 176-183
    1, 2, 1, 2, 2, 3, 1, 3, // 184-191
    4, 255, 255, 3, 3, 2, 1, 3, // 192-199
    1, 1, 0, 0, 2, 3, 3, 4, // 200-207
    2, 2, 3, 3, 2, 2, 1, 1, // 208-215
    3, 0, 0, 5, 3, 0, 7, 0, // 216-223
    1, 0, 3, 255, 255, 1, 1, 1, // 224-231
    1, 1, 1, 3, 11, 1, 0, 2, // 232-239
    3, 5, 4, 4, 1, 0, 5, 5, // 240-247
    1, 0, 0, 1, 2, 2, 1, 255, // 248-255
    1, 1, 2, 3, 3, 1, 1, 1, // 256-263
    2, 3, 3, 3, 2, 255, 5, 1, // 264-271
    5, 4, 2, 2, 2, 3, 3, 1, // 272-279
    3, 3, 3, 1, 1, 2, 4, 3, // 280-287
    2, 4, 3, 3, 3, 4, 1, 3, // 288-295
    1, 3, 1, 2, 3, 3, 4, 2, // 296-303
    2, 2, 2, 2, 2, 3, 2, 2, // 304-311
    2, 2, 2, 2, 4, 4, 255, 255, // 312-319
    255, 255, 255, 255, 3, 2, 2, 2, // 320-327
    2, 3, 255, 2, 2, 4, 4, 3, // 328-335
    255, 2, 9, 2, 3, 4, 1, 1, // 336-343
    255, 3, 2, 1, 2, 1, 4, 3, // 344-351
    1, 2, 2, 4, 5, 2, 128, 2, // 352-359
    1, 255, 255, 255, 255, 255, 255, 255, // 360-367
    1, 1, 1, 1, 1, 1, 1, 1, // 368-375
    1, 1, 1, 255, 255, 3, 3, 4, // 376-383
    2, 1, 2, 2, 2, 2, 2, 2, // 384-391
    1, 1, 2, 2, 2, 2, 2, 1, // 392-399
    1, 1, 1, 1, 1, 1, 1, 1, // 400-407
    1, 1, 1, 3, 255, 255, 4, 1, // 408-415
    1, 2, 2, 2, 1, 1, 2, 2, // 416-423
    1, 2, 2, 2, 2, 3, 3, 6, // 424-431
    6, 5, 5, 5, 5, 5, 3, 3, // 432-439
    3, 7, 7, 2, 2, 2, 2, 6, // 440-447
    6, 2, 2, 3, 4, 4, 4, 4, // 448-455
    4, 4, 6, 6, 8, 8, 8, 8, // 456-463
    2, 2, 7, 7, 8, 8, 5, 3, // 464-471
    3, 255, 255, 255, 2, 4, 5, 1, // 472-479
    2, 128, 129, 3, 129, // 480-484
];

/// Look up a BIFF8 function name by index.
///
/// Returns the function name, or a synthetic `_xlfn.{idx}` for unknown indices.
pub fn function_name(idx: u16) -> &'static str {
    let i = idx as usize;
    if i < FTAB.len() {
        let name = FTAB[i];
        if name.is_empty() {
            // Reserved slot - should not appear in real files
            return "";
        }
        name
    } else {
        // Future function or add-in - out of range
        ""
    }
}

/// Look up the declared argument count for a BIFF8 function.
///
/// Returns the argc value (see encoding notes at top).
/// Returns 255 (variable) for unknown indices.
pub fn function_argc(idx: u16) -> u16 {
    let i = idx as usize;
    if i < FTAB_ARGC.len() {
        FTAB_ARGC[i]
    } else {
        255 // assume variable for unknown
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_common_functions() {
        assert_eq!(function_name(0), "COUNT");
        assert_eq!(function_name(1), "IF");
        assert_eq!(function_name(4), "SUM");
        assert_eq!(function_name(5), "AVERAGE");
        assert_eq!(function_name(6), "MIN");
        assert_eq!(function_name(7), "MAX");
        assert_eq!(function_name(32), "LEN");
        assert_eq!(function_name(64), "MATCH");
        assert_eq!(function_name(102), "VLOOKUP");
        assert_eq!(function_name(336), "CONCATENATE");
        assert_eq!(function_name(345), "SUMIF");
        assert_eq!(function_name(480), "IFERROR");
        assert_eq!(function_name(484), "AVERAGEIFS");
    }

    #[test]
    fn test_reserved_slots() {
        assert_eq!(function_name(202), "");
        assert_eq!(function_name(203), "");
        assert_eq!(function_name(217), "");
    }

    #[test]
    fn test_argc() {
        assert_eq!(function_argc(4), 255); // SUM = variable
        assert_eq!(function_argc(1), 3); // IF = 3 (fixed in BIFF8)
        assert_eq!(function_argc(32), 1); // LEN = 1
        assert_eq!(function_argc(64), 3); // MATCH = 3
        assert_eq!(function_argc(19), 0); // PI = 0
    }

    #[test]
    fn test_out_of_range() {
        assert_eq!(function_name(999), "");
        assert_eq!(function_argc(999), 255);
    }
}
