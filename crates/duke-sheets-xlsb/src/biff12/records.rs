pub const BRT_ROW_HDR: u16 = 0x0000;
pub const BRT_CELL_BLANK: u16 = 0x0001;
pub const BRT_CELL_RK: u16 = 0x0002;
pub const BRT_CELL_ERROR: u16 = 0x0003;
pub const BRT_CELL_BOOL: u16 = 0x0004;
pub const BRT_CELL_REAL: u16 = 0x0005;
pub const BRT_CELL_ST: u16 = 0x0006;
pub const BRT_CELL_ISST: u16 = 0x0007;
pub const BRT_FMLA_STRING: u16 = 0x0008;
pub const BRT_FMLA_NUM: u16 = 0x0009;
pub const BRT_FMLA_BOOL: u16 = 0x000A;
pub const BRT_FMLA_ERROR: u16 = 0x000B;

pub const BRT_SS_ITEM: u16 = 0x0013;

pub const BRT_NAME: u16 = 0x0027;
pub const BRT_FONT: u16 = 0x002B;
pub const BRT_FMT: u16 = 0x002C;
pub const BRT_FILL: u16 = 0x002D;
pub const BRT_BORDER: u16 = 0x002E;
pub const BRT_XF: u16 = 0x002F;
pub const BRT_STYLE: u16 = 0x0030;

pub const BRT_FILE_VERSION: u16 = 0x0080;
pub const BRT_BEGIN_SHEET: u16 = 0x0081;
pub const BRT_END_SHEET_DATA: u16 = 0x0092;
pub const BRT_BEGIN_SHEET_DATA: u16 = 0x0091;
pub const BRT_END_BUNDLE_SHS: u16 = 0x0090;
pub const BRT_WS_PROP: u16 = 0x0093;
pub const BRT_WS_DIM: u16 = 0x0094;

pub const BRT_BEGIN_BOOK_VIEWS: u16 = 0x0087;
pub const BRT_END_BOOK_VIEWS: u16 = 0x0088;
pub const BRT_BOOK_VIEW: u16 = 0x008B;
pub const BRT_WB_PROP: u16 = 0x0099;
pub const BRT_BUNDLE_SH: u16 = 0x009C;
pub const BRT_BEGIN_SST: u16 = 0x009F;

pub const BRT_EXTERN_SHEET: u16 = 0x016A;
pub const BRT_PLACEHOLDER_NAME: u16 = 0x0169;
pub const BRT_BEGIN_SUP_BOOK: u16 = 0x0168;
pub const BRT_END_SUP_BOOK: u16 = 0x024C;
pub const BRT_SUP_SELF: u16 = 0x0165;
pub const BRT_SUP_BOOK_SRC: u16 = 0x016E;
pub const BRT_SUP_ADDIN: u16 = 0x029B;
pub const BRT_BEGIN_FILLS: u16 = 0x025B;
pub const BRT_END_FILLS: u16 = 0x025C;
pub const BRT_BEGIN_FONTS: u16 = 0x0263;
pub const BRT_END_FONTS: u16 = 0x0264;
pub const BRT_BEGIN_BORDERS: u16 = 0x0265;
pub const BRT_END_BORDERS: u16 = 0x0266;
pub const BRT_BEGIN_FMTS: u16 = 0x0267;
pub const BRT_END_FMTS: u16 = 0x0268;
pub const BRT_BEGIN_CELL_XFS: u16 = 0x0269;
pub const BRT_END_CELL_XFS: u16 = 0x026A;
pub const BRT_BEGIN_STYLES: u16 = 0x026B;
pub const BRT_END_STYLES: u16 = 0x026C;
pub const BRT_BEGIN_CELL_STYLE_XFS: u16 = 0x0272;
pub const BRT_END_CELL_STYLE_XFS: u16 = 0x0273;

pub const BRT_BEGIN_VIEWS: u16 = 0x0085;
pub const BRT_END_VIEWS: u16 = 0x0086;
pub const BRT_BEGIN_AC_BLOCKS: u16 = 0x0025;
pub const BRT_END_AC_BLOCKS: u16 = 0x0026;
pub const BRT_WS_FMT_INFO: u16 = 0x01E5;
pub const BRT_BEGIN_COL_INFOS: u16 = 0x0186;
pub const BRT_END_COL_INFOS: u16 = 0x0187;

pub const BRT_BEGIN_FRT: u16 = 0x0023;
pub const BRT_END_FRT: u16 = 0x0024;

pub const BRT_COL_INFO: u16 = 0x003C;

pub const BRT_BEGIN_SHEET_VIEW: u16 = 0x0089;
pub const BRT_END_SHEET_VIEW: u16 = 0x008A;
pub const BRT_PANE: u16 = 0x0097;
pub const BRT_SEL: u16 = 0x0098;

pub const BRT_BEGIN_A_FILTER: u16 = 0x00A1;
pub const BRT_END_A_FILTER: u16 = 0x00A2;
pub const BRT_BEGIN_FILTER_COLUMN: u16 = 0x00A3;
pub const BRT_END_FILTER_COLUMN: u16 = 0x00A4;
pub const BRT_BEGIN_FILTERS: u16 = 0x00A5;
pub const BRT_END_FILTERS: u16 = 0x00A6;
pub const BRT_FILTER: u16 = 0x00A7;
// BrtDynamicFilter = 0x00AB confirmed empirically by dumping an XLSB
// file Excel emits when an AboveAverage filter is applied via COM.
// BrtColorFilter = 0x00A8 per the [MS-XLSB] §2.3 record enumeration
// (168 = BrtColorFilter, 169 = BrtIconFilter).
pub const BRT_DYNAMIC_FILTER: u16 = 0x00AB;
pub const BRT_COLOR_FILTER: u16 = 0x00A8;
pub const BRT_TOP10_FILTER: u16 = 0x00AA;
pub const BRT_BEGIN_CUSTOM_FILTERS: u16 = 0x00AC;
pub const BRT_END_CUSTOM_FILTERS: u16 = 0x00AD;
pub const BRT_CUSTOM_FILTER: u16 = 0x00AE;

pub const BRT_MERGE_CELL: u16 = 0x00B0;
pub const BRT_BEGIN_MERGE_CELLS: u16 = 0x00B1;
pub const BRT_END_MERGE_CELLS: u16 = 0x00B2;

pub const BRT_MARGINS: u16 = 0x01DC;
pub const BRT_PRINT_OPTIONS: u16 = 0x01DD;
pub const BRT_PAGE_SETUP: u16 = 0x01DE;
// BrtBeginHeaderFooter is followed by HeaderFooterString fields and
// then BrtEndHeaderFooter even though all the data is in the Begin
// record itself.
pub const BRT_HEADER_FOOTER: u16 = 0x01DF;
pub const BRT_END_HEADER_FOOTER: u16 = 0x01E0;

pub const BRT_H_LINK: u16 = 0x01EE;

pub const BRT_BEGIN_COND_FMT: u16 = 0x01CD; // 461 BrtBeginConditionalFormatting
pub const BRT_END_COND_FMT: u16 = 0x01CE; // 462 BrtEndConditionalFormatting
pub const BRT_BEGIN_CF_RULE: u16 = 0x01CF; // 463 BrtBeginCFRule
pub const BRT_END_CF_RULE: u16 = 0x01D0; // 464 BrtEndCFRule
pub const BRT_BEGIN_ICON_SET: u16 = 0x01D1; // 465
pub const BRT_END_ICON_SET: u16 = 0x01D2; // 466
pub const BRT_BEGIN_DATA_BAR: u16 = 0x01D3; // 467
pub const BRT_END_DATA_BAR: u16 = 0x01D4; // 468
pub const BRT_BEGIN_COLOR_SCALE: u16 = 0x01D5; // 469
pub const BRT_END_COLOR_SCALE: u16 = 0x01D6; // 470
pub const BRT_CFVO: u16 = 0x01D7; // 471
pub const BRT_CF_COLOR: u16 = 0x0234; // 564

// Per LibreOffice's battle-tested BIFF12 constants and Excel emission:
// BrtBeginDVals = 0x023D, BrtEndDVals = 0x023E, BrtDVal = 0x0040.
// The MS-XLSB section numbers (2.4.55, 2.4.356) refer to TOC entries,
// not record IDs; the actual record IDs are different.
pub const BRT_BEGIN_DVAL: u16 = 0x023D;
pub const BRT_END_DVAL: u16 = 0x023E;
pub const BRT_DVAL: u16 = 0x0040;

pub const BRT_BEGIN_EC_DB_PROPS: u16 = 0x00CB;
pub const BRT_END_EC_DB_PROPS: u16 = 0x00CC;
pub const BRT_BEGIN_EC_OLAP_PROPS: u16 = 0x00CD;
pub const BRT_END_EC_OLAP_PROPS: u16 = 0x00CE;
pub const BRT_BEGIN_EC_WEB_PROPS: u16 = 0x0105;
pub const BRT_BEGIN_EC_PARAMS: u16 = 0x0109;
pub const BRT_BEGIN_EC_PARAM: u16 = 0x010B;
pub const BRT_END_EC_PARAM: u16 = 0x010C;
pub const BRT_END_EC_PARAMS: u16 = 0x010A;
pub const BRT_END_EC_WEB_PROPS: u16 = 0x0106;
pub const BRT_BEGIN_EXT_CONNECTION: u16 = 0x00C9;
pub const BRT_END_EXT_CONNECTION: u16 = 0x00CA;
pub const BRT_BEGIN_EXT_CONNECTIONS: u16 = 0x01AD;
pub const BRT_END_EXT_CONNECTIONS: u16 = 0x01AE;
pub const BRT_BEGIN_EC_TXT_WIZ: u16 = 0x021A;
pub const BRT_END_EC_TXT_WIZ: u16 = 0x021B;
pub const BRT_BEGIN_EC_TW_FLD_INFO_LST: u16 = 0x021C;
pub const BRT_END_EC_TW_FLD_INFO_LST: u16 = 0x021D;
pub const BRT_BEGIN_EC_TW_FLD_INFO: u16 = 0x021E;

pub const BRT_BEGIN_PIVOT_CACHE_DEF: u16 = 0x00B3;
pub const BRT_END_PIVOT_CACHE_DEF: u16 = 0x00B4;
pub const BRT_BEGIN_PCD_FIELDS: u16 = 0x00B5;
pub const BRT_END_PCD_FIELDS: u16 = 0x00B6;
pub const BRT_BEGIN_PCD_FIELD: u16 = 0x00B7;
pub const BRT_END_PCD_FIELD: u16 = 0x00B8;
pub const BRT_BEGIN_PCD_SOURCE: u16 = 0x00B9;
pub const BRT_END_PCD_SOURCE: u16 = 0x00BA;
pub const BRT_BEGIN_PCDS_SHEET: u16 = 0x00BB;
pub const BRT_END_PCDS_SHEET: u16 = 0x00BC;
pub const BRT_BEGIN_PCDS_CONSOL: u16 = 0x00CF;
pub const BRT_END_PCDS_CONSOL: u16 = 0x00D0;
pub const BRT_BEGIN_PCDSC_PAGES: u16 = 0x00D1;
pub const BRT_END_PCDSC_PAGES: u16 = 0x00D2;
pub const BRT_BEGIN_PCDSC_PAGE: u16 = 0x00D3;
pub const BRT_END_PCDSC_PAGE: u16 = 0x00D4;
pub const BRT_BEGIN_PCDSC_PITEM: u16 = 0x00D5;
pub const BRT_END_PCDSC_PITEM: u16 = 0x00D6;
pub const BRT_BEGIN_PCDSC_SETS: u16 = 0x00D7;
pub const BRT_END_PCDSC_SETS: u16 = 0x00D8;
pub const BRT_BEGIN_PCDSC_SET: u16 = 0x00D9;
pub const BRT_END_PCDSC_SET: u16 = 0x00DA;
pub const BRT_BEGIN_PCD_SHARED_ITEMS: u16 = 0x00BD;
pub const BRT_END_PCD_SHARED_ITEMS: u16 = 0x00BE;
pub const BRT_BEGIN_PCD_RECORDS: u16 = 0x00C1;
pub const BRT_END_PCD_RECORDS: u16 = 0x00C2;
pub const BRT_BEGIN_PCD_HIERARCHIES: u16 = 0x00C3;
pub const BRT_END_PCD_HIERARCHIES: u16 = 0x00C4;
pub const BRT_BEGIN_PCD_HIERARCHY: u16 = 0x00C5;
pub const BRT_END_PCD_HIERARCHY: u16 = 0x00C6;
pub const BRT_BEGIN_PCDH_FIELDS_USAGE: u16 = 0x00C7;
pub const BRT_END_PCDH_FIELDS_USAGE: u16 = 0x00C8;
pub const BRT_BEGIN_PCD_CALC_ITEMS: u16 = 0x00F3;
pub const BRT_END_PCD_CALC_ITEMS: u16 = 0x00F4;
pub const BRT_BEGIN_PCD_CALC_ITEM: u16 = 0x00F5;
pub const BRT_END_PCD_CALC_ITEM: u16 = 0x00F6;
pub const BRT_BEGIN_P_RULE: u16 = 0x00F7;
pub const BRT_END_P_RULE: u16 = 0x00F8;
pub const BRT_BEGIN_PR_FILTERS: u16 = 0x00F9;
pub const BRT_END_PR_FILTERS: u16 = 0x00FA;
pub const BRT_BEGIN_PR_FILTER: u16 = 0x00FB;
pub const BRT_END_PR_FILTER: u16 = 0x00FC;
pub const BRT_BEGIN_PCDF_GROUP: u16 = 0x00DB;
pub const BRT_END_PCDF_GROUP: u16 = 0x00DC;
pub const BRT_BEGIN_PCDFG_ITEMS: u16 = 0x00DD;
pub const BRT_END_PCDFG_ITEMS: u16 = 0x00DE;
pub const BRT_BEGIN_PCDFG_RANGE: u16 = 0x00DF;
pub const BRT_END_PCDFG_RANGE: u16 = 0x00E0;
pub const BRT_BEGIN_PCDFG_DISCRETE: u16 = 0x00E1;
pub const BRT_END_PCDFG_DISCRETE: u16 = 0x00E2;
pub const BRT_BEGIN_PNAMES: u16 = 0x00FD;
pub const BRT_END_PNAMES: u16 = 0x00FE;
pub const BRT_BEGIN_PNAME: u16 = 0x00FF;
pub const BRT_END_PNAME: u16 = 0x0100;
pub const BRT_BEGIN_PNPAIRS: u16 = 0x0101;
pub const BRT_END_PNPAIRS: u16 = 0x0102;
pub const BRT_BEGIN_PNPAIR: u16 = 0x0103;
pub const BRT_END_PNPAIR: u16 = 0x0104;
pub const BRT_BEGIN_PRF_ITEM: u16 = 0x017E;
pub const BRT_END_PRF_ITEM: u16 = 0x017F;

pub const BRT_PCDI_STRING: u16 = 0x0018;
pub const BRT_PCDI_NUMBER: u16 = 0x0015;
pub const BRT_PCDI_BOOLEAN: u16 = 0x0016;
pub const BRT_PCDI_ERROR: u16 = 0x0017;
pub const BRT_PCDI_MISSING: u16 = 0x0014;
pub const BRT_PCDI_DATETIME: u16 = 0x0019;
pub const BRT_PCDI_INDEX: u16 = 0x001A;
pub const BRT_PCDIA_MISSING: u16 = 0x001B;
pub const BRT_PCDIA_NUMBER: u16 = 0x001C;
pub const BRT_PCDIA_BOOLEAN: u16 = 0x001D;
pub const BRT_PCDIA_ERROR: u16 = 0x001E;
pub const BRT_PCDIA_STRING: u16 = 0x001F;
pub const BRT_PCDIA_DATETIME: u16 = 0x0020;
pub const BRT_PCD_RECORD: u16 = 0x0021;

pub const BRT_BEGIN_PIVOT_CACHE_IDS: u16 = 0x0180;
pub const BRT_END_PIVOT_CACHE_IDS: u16 = 0x0181;
pub const BRT_PIVOT_CACHE_ID: u16 = 0x0182;
pub const BRT_END_PIVOT_CACHE_ID: u16 = 0x0183;

pub const BRT_BEGIN_SXVIEW: u16 = 0x0118;
pub const BRT_END_SXVIEW: u16 = 0x013B;
pub const BRT_BEGIN_SXVI: u16 = 0x011A;
pub const BRT_END_SXVI: u16 = 0x0119;
pub const BRT_BEGIN_SXVIS: u16 = 0x011B;
pub const BRT_END_SXVIS: u16 = 0x011C;
pub const BRT_BEGIN_SXVD: u16 = 0x011D;
pub const BRT_END_SXVD: u16 = 0x011E;
pub const BRT_BEGIN_SXVD14: u16 = 0x01CB;
pub const BRT_END_SXVD14: u16 = 0x01CC;
pub const BRT_BEGIN_SXVDS: u16 = 0x011F;
pub const BRT_END_SXVDS: u16 = 0x0120;
pub const BRT_BEGIN_SXPI: u16 = 0x0121;
pub const BRT_END_SXPI: u16 = 0x0122;
pub const BRT_BEGIN_SXPIS: u16 = 0x0123;
pub const BRT_END_SXPIS: u16 = 0x0124;
pub const BRT_BEGIN_SXDI: u16 = 0x0125;
pub const BRT_END_SXDI: u16 = 0x0126;
pub const BRT_SXDI14: u16 = 0x0414;
pub const BRT_BEGIN_SXDIS: u16 = 0x0127;
pub const BRT_END_SXDIS: u16 = 0x0128;
pub const BRT_BEGIN_PIVOT_AREA: u16 = 0x00F7;
pub const BRT_END_PIVOT_AREA: u16 = 0x00F8;
pub const BRT_BEGIN_PIVOT_AREA_REFS: u16 = 0x00F9;
pub const BRT_END_PIVOT_AREA_REFS: u16 = 0x00FA;
pub const BRT_BEGIN_PIVOT_AREA_REF: u16 = 0x00FB;
pub const BRT_END_PIVOT_AREA_REF: u16 = 0x00FC;
pub const BRT_PIVOT_AREA_REF_ITEM: u16 = 0x017E;
pub const BRT_END_PIVOT_AREA_REF_ITEM: u16 = 0x017F;
pub const BRT_BEGIN_SXLI: u16 = 0x0129;
pub const BRT_END_SXLI: u16 = 0x012A;
pub const BRT_BEGIN_SX_ROW_ITEMS: u16 = 0x012B;
pub const BRT_END_SX_ROW_ITEMS: u16 = 0x012C;
pub const BRT_BEGIN_SX_COL_ITEMS: u16 = 0x012D;
pub const BRT_END_SX_COL_ITEMS: u16 = 0x012E;
pub const BRT_BEGIN_ISXVD_RWS: u16 = 0x0135;
pub const BRT_END_ISXVD_RWS: u16 = 0x0136;
pub const BRT_BEGIN_ISXVD_COLS: u16 = 0x0137;
pub const BRT_END_ISXVD_COLS: u16 = 0x0138;
pub const BRT_END_SX_LOCATION: u16 = 0x0139;
pub const BRT_SX_LOCATION: u16 = 0x013A;
pub const BRT_BEGIN_SXTHS: u16 = 0x013C;
pub const BRT_END_SXTHS: u16 = 0x013D;
pub const BRT_BEGIN_SXTH: u16 = 0x013E;
pub const BRT_END_SXTH: u16 = 0x013F;
pub const BRT_BEGIN_ISXTH_RWS: u16 = 0x0140;
pub const BRT_END_ISXTH_RWS: u16 = 0x0141;
pub const BRT_BEGIN_ISXTH_COLS: u16 = 0x0142;
pub const BRT_END_ISXTH_COLS: u16 = 0x0143;
pub const BRT_SXLI_ITEM: u16 = 0x0184;
pub const BRT_END_SXLI_ITEM: u16 = 0x0185;
pub const BRT_SX_VIEW_STYLE: u16 = 0x0201;
pub const BRT_BEGIN_SX_FILTERS: u16 = 0x0257;
pub const BRT_END_SX_FILTERS: u16 = 0x0258;
pub const BRT_BEGIN_SX_FILTER: u16 = 0x0259;
pub const BRT_END_SX_FILTER: u16 = 0x025A;

// Comments part records per MS-XLSB 2.4.33 (pinned against
// Excel-authored xl/comments1.bin).
pub const BRT_BEGIN_COMMENTS: u16 = 0x0274;
pub const BRT_END_COMMENTS: u16 = 0x0275;
pub const BRT_BEGIN_COMMENT_AUTHORS: u16 = 0x0276;
pub const BRT_END_COMMENT_AUTHORS: u16 = 0x0277;
pub const BRT_COMMENT_AUTHOR: u16 = 0x0278;
pub const BRT_BEGIN_COMMENT_LIST: u16 = 0x0279;
pub const BRT_END_COMMENT_LIST: u16 = 0x027A;
pub const BRT_BEGIN_COMMENT: u16 = 0x027B;
pub const BRT_END_COMMENT: u16 = 0x027C;
pub const BRT_COMMENT_TEXT: u16 = 0x027D;

// Off-spec 0x0278-based ids our writer used to emit (Excel refuses
// them); the reader still accepts them for back-compat. The legacy
// BrtBeginComment body was 12 bytes (iauthor + row + col) instead of
// the spec's 36.
pub const BRT_LEGACY_BEGIN_COMMENT_AUTHORS: u16 = 0x0278;
pub const BRT_LEGACY_COMMENT_AUTHOR: u16 = 0x027A;
pub const BRT_LEGACY_END_COMMENT_LIST: u16 = 0x027C;
pub const BRT_LEGACY_BEGIN_COMMENT: u16 = 0x027D;
pub const BRT_LEGACY_END_COMMENT: u16 = 0x027E;
pub const BRT_LEGACY_COMMENT_TEXT: u16 = 0x027F;

// Future-record wrapper (skipped transparently around comments).
pub const BRT_AC_BEGIN: u16 = 0x0025;
pub const BRT_AC_END: u16 = 0x0026;

// MS-XLSB §2.2.1: BrtDrawing = 550, BrtLegacyDrawing = 551
// (552 is BrtLegacyDrawingHF, the header/footer variant).
pub const BRT_DRAWING: u16 = 0x0226;
pub const BRT_LEGACY_DRAWING: u16 = 0x0227;

pub const BRT_BEGIN_DIMS: u16 = 0x0111;
pub const BRT_END_DIMS: u16 = 0x0112;
pub const BRT_BEGIN_DIM: u16 = 0x0113;
pub const BRT_END_DIM: u16 = 0x0114;

pub const BRT_BEGIN_LIST: u16 = 0x0157; // 343
pub const BRT_END_LIST: u16 = 0x0158; // 344
pub const BRT_BEGIN_LIST_COLS: u16 = 0x0159; // 345
pub const BRT_END_LIST_COLS: u16 = 0x015A; // 346
pub const BRT_BEGIN_LIST_COL: u16 = 0x015B; // 347
pub const BRT_END_LIST_COL: u16 = 0x015C; // 348
pub const BRT_TABLE_STYLE_CLIENT: u16 = 0x0201; // 513

pub const BRT_BEGIN_LIST_PARTS: u16 = 0x0294; // 660
pub const BRT_LIST_PART: u16 = 0x0295; // 661
pub const BRT_END_LIST_PARTS: u16 = 0x0296; // 662

// Pinned from Excel-authored XLSB files. These protection records are
// adjacent in the BIFF12 stream and using the wrong neighbor ID makes Excel
// reject the workbook before repair.
pub const BRT_BOOK_PROTECTION: u16 = 0x0216;
pub const BRT_SHEET_PROTECTION: u16 = 0x0217;
pub const BRT_RANGE_PROTECTION: u16 = 0x0218;

pub const BRT_BEGIN_RW_BRK: u16 = 0x0188; // 392
pub const BRT_END_RW_BRK: u16 = 0x0189; // 393
pub const BRT_BEGIN_COL_BRK: u16 = 0x018A; // 394
pub const BRT_END_COL_BRK: u16 = 0x018B; // 395
pub const BRT_BRK: u16 = 0x018C; // 396
