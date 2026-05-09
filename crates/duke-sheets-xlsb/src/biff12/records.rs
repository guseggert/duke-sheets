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
pub const BRT_BEGIN_SUP_BOOK: u16 = 0x01AE;
pub const BRT_END_SUP_BOOK: u16 = 0x01AF;
pub const BRT_SUP_SELF: u16 = 0x016F;
pub const BRT_SUP_BOOK_SRC: u16 = 0x016E;
pub const BRT_SUP_ADDIN: u16 = 0x01B5;
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
// Both record IDs confirmed empirically by dumping XLSB files Excel
// emits when a dynamic / color filter is applied via the COM API:
// BrtDynamicFilter = 0x00AB, BrtColorFilter = 0x00A9.
pub const BRT_DYNAMIC_FILTER: u16 = 0x00AB;
pub const BRT_COLOR_FILTER: u16 = 0x00A9;
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

pub const BRT_BEGIN_COMMENT_AUTHORS: u16 = 0x0278;
pub const BRT_END_COMMENT_AUTHORS: u16 = 0x0279;
pub const BRT_COMMENT_AUTHOR: u16 = 0x027A;
pub const BRT_BEGIN_COMMENT_LIST: u16 = 0x027B;
pub const BRT_END_COMMENT_LIST: u16 = 0x027C;
pub const BRT_BEGIN_COMMENT: u16 = 0x027D;
pub const BRT_END_COMMENT: u16 = 0x027E;
pub const BRT_COMMENT_TEXT: u16 = 0x027F;

pub const BRT_LEGACY_DRAWING: u16 = 0x0228;
pub const BRT_DRAWING: u16 = 0x0235;

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

pub const BRT_SHEET_PROTECTION: u16 = 0x0217; // 535

pub const BRT_BEGIN_RW_BRK: u16 = 0x0188; // 392
pub const BRT_END_RW_BRK: u16 = 0x0189; // 393
pub const BRT_BEGIN_COL_BRK: u16 = 0x018A; // 394
pub const BRT_END_COL_BRK: u16 = 0x018B; // 395
pub const BRT_BRK: u16 = 0x018C; // 396
