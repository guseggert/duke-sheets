//! BIFF8 record type constants.
//!
//! Reference: [MS-XLS] §2.3 — Record Enumeration

// ── Stream structure ────────────────────────────────────────────────────
pub const BOF: u16 = 0x0809;
pub const EOF: u16 = 0x000A;
pub const CONTINUE: u16 = 0x003C;

// ── Workbook globals ────────────────────────────────────────────────────
pub const BOUNDSHEET: u16 = 0x0085; // Sheet name, type, visibility, stream offset
pub const SST: u16 = 0x00FC; // Shared String Table
pub const EXTSST: u16 = 0x00FF; // Extended SST (hash table — we skip it)
pub const DATEMODE: u16 = 0x0022; // 1900 vs 1904 date system (a.k.a. DATE1904)
pub const CODEPAGE: u16 = 0x0042; // Code page (should be 1200 = UTF-16 for BIFF8)
pub const PALETTE: u16 = 0x0092; // Custom color palette (overrides default 56)
pub const FONT: u16 = 0x0031; // Font definition
pub const FORMAT: u16 = 0x041E; // Number format string
pub const XF: u16 = 0x00E0; // Extended Format (cell format record)
pub const STYLE: u16 = 0x0293; // Named cell style

// ── Cell records ────────────────────────────────────────────────────────
pub const DIMENSION: u16 = 0x0200; // Used range (first/last row/col)
pub const LABELSST: u16 = 0x00FD; // Cell containing SST string index
pub const LABEL: u16 = 0x0204; // Cell with inline string (rare in BIFF8)
pub const NUMBER: u16 = 0x0203; // Cell with IEEE 754 double
pub const RK: u16 = 0x027E; // Cell with compressed number (RK encoding)
pub const MULRK: u16 = 0x00BD; // Multiple RK values in one row
pub const BLANK: u16 = 0x0201; // Empty cell with formatting
pub const MULBLANK: u16 = 0x00BE; // Multiple blanks with formatting
pub const BOOLERR: u16 = 0x0205; // Boolean or error cell
pub const FORMULA: u16 = 0x0006; // Formula cell with cached result
pub const STRING: u16 = 0x0207; // Cached string result for preceding FORMULA
pub const RSTRING: u16 = 0x00D6; // Rich-text inline string (rare)
pub const ARRAY: u16 = 0x0221; // Array formula

// ── Sheet structure ─────────────────────────────────────────────────────
pub const ROW: u16 = 0x0208; // Row height, visibility, default format
pub const COLINFO: u16 = 0x007D; // Column width, visibility, default format
pub const DEFCOLWIDTH: u16 = 0x0055; // Default column width
pub const DEFAULTROWHEIGHT: u16 = 0x0225; // Default row height
pub const MERGECELLS: u16 = 0x00E5; // Merged cell ranges
pub const WINDOW2: u16 = 0x023E; // Sheet view settings (freeze panes, etc.)
pub const PANE: u16 = 0x0041; // Pane split position
pub const SELECTION: u16 = 0x001D; // Selected cell range
pub const HLINK: u16 = 0x01B8; // Hyperlink

// ── BOF subtypes (the `dt` field) ───────────────────────────────────────
pub const BOF_WORKBOOK_GLOBALS: u16 = 0x0005;
pub const BOF_WORKSHEET: u16 = 0x0010;
pub const BOF_CHART: u16 = 0x0020;
pub const BOF_MACRO: u16 = 0x0040;

/// BIFF version we support.
pub const BIFF8_VERSION: u16 = 0x0600;
