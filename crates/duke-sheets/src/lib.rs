//! # duke-sheets
//!
//! A Rust library for reading, writing, and manipulating spreadsheets.
//!
//! Duke-sheets provides an API similar to Aspose Cells for working with Excel files
//! (XLSX, XLS) and CSV files.
//!
//! ## Features
//!
//! - Read and write XLSX files (Office Open XML)
//! - Read and write XLS files (legacy BIFF8 format) - optional
//! - Read and write CSV files
//! - Full formula evaluation
//! - Cell styling (fonts, colors, borders, etc.)
//! - Charts support
//! - Large file support via streaming APIs
//!
//! ## Example
//!
//! ```rust
//! use duke_sheets::prelude::*;
//!
//! // Create a new workbook
//! let mut workbook = Workbook::new();
//!
//! // Get the first worksheet
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! // Set cell values
//! sheet.set_cell_value("A1", "Hello").unwrap();
//! sheet.set_cell_value("B1", 42.0).unwrap();
//! sheet.set_cell_value("C1", true).unwrap();
//!
//! // Set a formula
//! sheet.set_cell_formula("D1", "=B1*2").unwrap();
//!
//! // Save to file
//! // workbook.save("output.xlsx").unwrap();
//! ```

pub mod calculation;
pub mod prelude;

// Re-export calculation types
pub use calculation::{CalculationOptions, CalculationStats, WorkbookCalculationExt};

// Re-export core types
pub use duke_sheets_core::auto_filter::{ColorFilter, DynamicFilter, DynamicFilterType};
pub use duke_sheets_core::{
    default_comment_anchor,
    hash_legacy_protection_password,

    radio_groups,
    rich_text_to_plain,
    validate_anchor,
    validate_group_child,
    Alignment,
    AutoFilter,
    BorderEdge,
    BorderLineStyle,
    BorderStyle,
    CellAddress,
    // Comments
    CellComment,
    CellData,

    CellError,
    CellRange,
    // Cell types
    CellValue,
    CellView,
    // Conditional formatting types
    CfColorValue,
    CfOperator,
    CfRuleType,
    CfValue,
    CfValueType,
    // Chart sheet types
    ChartSheet,
    // Form control types
    CheckState,
    // Drawing types
    ChildTransform,
    Color,
    ColumnFilter,
    CommentRef,
    ConditionalFormatRule,
    ControlText,
    // Data validation types
    CustomFilterCondition,
    CustomFilters,
    DataValidation,
    DrawingKind,
    DrawingMeta,
    DrawingNodeMut,
    DrawingNodeRef,
    DrawingObject,
    DrawingPath,
    DrawingText,
    Placed,
    // Error types
    Error,
    FillStyle,
    FilterColumn,
    FilterOperator,
    FontStyle,
    FormControl,
    FormControlInteractionResult,
    FormControlKind,
    // Sheet-level types
    FreezePanes,
    Group,
    GroupChild,
    GroupTransform,
    HorizontalAlignment,
    Hyperlink,
    IconSetStyle,
    ListSelection,
    // Locale for cell formatting
    Locale,

    NumberFormat,

    PageBreak,
    PageOrientation,
    PageSetup,
    PlacedControl,
    ProtectedRange,
    RawDrawing,
    RawRel,
    RectEmu,
    Result,

    // Rich text types
    RichTextRun,
    RunFont,
    Shape,
    ShapeFill,
    ShapeGeometry,
    ShapeLine,
    SheetProtection,
    SheetSlot,
    // Style types
    Style,
    StylePool,
    // Table types
    Table,
    TableColumn,
    TableStyleInfo,
    ThemePalette,
    TimePeriod,
    Top10Filter,
    TotalsRowFunction,

    ValidationErrorStyle,
    ValidationOperator,
    ValidationType,
    ValueFilter,
    VerticalAlignment,
    // Main types
    Workbook,
    WorkbookProtection,
    WorkbookSettings,
    Worksheet,
    MAX_COLS,
    // Constants
    MAX_ROWS,
    MAX_SHEET_NAME_LEN,
};

// Re-export named range module (contains NamedRange, NameScope, NamedRangeCollection)
pub use duke_sheets_core::named_range;

// Re-export formula types
pub use duke_sheets_formula::{
    evaluate, parse_formula, EvaluationContext, FormulaError, FormulaExpr, FormulaResult,
    FormulaValue, ImageInfo, ImageSizing,
};

// Re-export chart types
pub use duke_sheets_chart::{
    column_width_to_emu, row_height_to_emu, Axis, AxisPosition, CellMarker, Chart, ChartEx,
    ChartType, DataReference, DataSeries, DrawingAnchor, DrawingMetrics, EditAs, EmbeddedImage,
    ImageFormat, Legend,
};

// Re-export I/O types
pub use duke_sheets_csv::{CsvError, CsvReadOptions, CsvReader, CsvWriteOptions, CsvWriter};
#[cfg(feature = "xls")]
pub use duke_sheets_xls::{XlsError, XlsReader, XlsWriter};
#[cfg(feature = "xlsb")]
pub use duke_sheets_xlsb::{XlsbError, XlsbReader, XlsbWriter};
pub use duke_sheets_xlsx::{XlsxError, XlsxReader, XlsxWriter};

use std::io::{Cursor, Read, Seek, SeekFrom};
use std::path::Path;

/// Detected file format from magic bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileFormat {
    /// XLSX / Office Open XML (ZIP container starting with PK\x03\x04)
    Xlsx,
    /// XLS / BIFF8 (CFB container starting with \xD0\xCF\x11\xE0)
    Xls,
    /// XLSB / Excel Binary Workbook (ZIP container with BIFF12 internals)
    Xlsb,
    /// Unknown format
    Unknown,
}

/// True if the buffer starts with the CFB magic. Encrypted XLSX files
/// are CFB envelopes; plain XLSX files are ZIPs.
fn is_cfb_magic(bytes: &[u8]) -> bool {
    bytes.len() >= 8 && bytes[0..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
}

/// Sniff the first few bytes of a buffer to determine its file format.
pub fn detect_format(bytes: &[u8]) -> FileFormat {
    if bytes.len() >= 4 && bytes[0..4] == [0x50, 0x4B, 0x03, 0x04] {
        if let Ok(archive) = zip::ZipArchive::new(Cursor::new(bytes)) {
            if archive.index_for_name("xl/workbook.bin").is_some() {
                return FileFormat::Xlsb;
            }
        }
        FileFormat::Xlsx
    } else if bytes.len() >= 8 && bytes[0..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        FileFormat::Xls
    } else {
        FileFormat::Unknown
    }
}

fn detect_format_from_path(path: &Path) -> Result<FileFormat> {
    let mut file = std::fs::File::open(path).map_err(|e| Error::other(e.to_string()))?;
    let mut magic = [0u8; 8];
    let n = file
        .read(&mut magic)
        .map_err(|e| Error::other(e.to_string()))?;

    if n >= 4 && magic[0..4] == [0x50, 0x4B, 0x03, 0x04] {
        file.seek(SeekFrom::Start(0))
            .map_err(|e| Error::other(e.to_string()))?;
        if let Ok(archive) = zip::ZipArchive::new(file) {
            if archive.index_for_name("xl/workbook.bin").is_some() {
                return Ok(FileFormat::Xlsb);
            }
        }
        Ok(FileFormat::Xlsx)
    } else if is_cfb_magic(&magic[..n]) {
        Ok(FileFormat::Xls)
    } else {
        Ok(FileFormat::Unknown)
    }
}

fn path_has_extension(path: &Path, ext: &str) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case(ext))
}

fn path_has_any_extension(path: &Path, extensions: &[&str]) -> bool {
    extensions.iter().any(|ext| path_has_extension(path, ext))
}

fn path_has_xlsx_family_extension(path: &Path) -> bool {
    path_has_any_extension(path, &["xlsx", "xlsm", "xltx", "xltm"])
}

fn read_csv_workbook(path: &Path) -> Result<Workbook> {
    let worksheet = CsvReader::read_file(path, &CsvReadOptions::default())
        .map_err(|e| Error::other(e.to_string()))?;

    let mut workbook = Workbook::empty();
    workbook.add_existing_worksheet(worksheet)?;
    Ok(workbook)
}

fn unsupported_file_format(path: &Path) -> Error {
    Error::other(format!("Unsupported file format: {}", path.display()))
}

fn is_ooxml_envelope_probe_miss(msg: &str) -> bool {
    msg.starts_with("CFB envelope open failed:") || msg.starts_with("not an OOXML envelope")
}

/// Options controlling how a workbook is opened.
///
/// The struct is non-exhaustive so fields can be added without a major
/// version bump.
#[derive(Clone)]
#[non_exhaustive]
pub struct WorkbookOpenOptions {
    /// Password for encrypted workbooks. `None` falls through to the
    /// `VelvetSweatshop` sentinel retry (when enabled), then to an
    /// `Encrypted` error.
    pub password: Option<String>,

    /// When `true` (the default, matching Excel's behavior), encrypted
    /// files with no supplied password are auto-decrypted with the
    /// well-known password `"VelvetSweatshop"` before reporting them
    /// as encrypted. Files protected this way (mostly Excel-2007-era
    /// templates with tamper-evidence rather than real protection)
    /// open transparently. Set to `false` for strict semantics where
    /// no password supplied always errors.
    pub try_velvet_sweatshop: bool,

    /// Skip the data-integrity HMAC check on Agile-encrypted files.
    /// Default `false` (matches Office). Turn on for forensic /
    /// recovery scenarios where you want to read damaged or tampered
    /// files even though their integrity check fails.
    pub skip_integrity_check: bool,
}

impl Default for WorkbookOpenOptions {
    fn default() -> Self {
        Self {
            password: None,
            try_velvet_sweatshop: true,
            skip_integrity_check: false,
        }
    }
}

impl WorkbookOpenOptions {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn password(mut self, pw: impl Into<String>) -> Self {
        self.password = Some(pw.into());
        self
    }

    /// Disable the `VelvetSweatshop` sentinel retry. When this is set,
    /// an encrypted file with no explicit password always returns an
    /// `Encrypted` error.
    pub fn strict_password(mut self) -> Self {
        self.try_velvet_sweatshop = false;
        self
    }

    /// Skip the integrity-check HMAC after decryption. Use only for
    /// forensic / recovery work; default-off matches Office.
    pub fn skip_integrity_check(mut self) -> Self {
        self.skip_integrity_check = true;
        self
    }
}

// Redact the password in Debug output to avoid accidentally logging it.
impl std::fmt::Debug for WorkbookOpenOptions {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WorkbookOpenOptions")
            .field("password", &self.password.as_ref().map(|_| "<redacted>"))
            .finish()
    }
}

/// Options controlling how a workbook is saved.
///
/// Like [`WorkbookOpenOptions`], this struct is non-exhaustive; new
/// fields will appear as features land. The `encryption` selector is
/// currently ignored because the writer doesn't yet produce encrypted
/// output.
#[derive(Default, Clone)]
#[non_exhaustive]
pub struct WorkbookSaveOptions {
    /// Password to encrypt the output with. `None` writes plaintext.
    pub password: Option<String>,

    /// Which encryption profile to use when `password` is supplied.
    pub encryption: EncryptionProfile,
}

impl WorkbookSaveOptions {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn password(mut self, pw: impl Into<String>) -> Self {
        self.password = Some(pw.into());
        self
    }

    pub fn encryption(mut self, profile: EncryptionProfile) -> Self {
        self.encryption = profile;
        self
    }
}

impl std::fmt::Debug for WorkbookSaveOptions {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WorkbookSaveOptions")
            .field("password", &self.password.as_ref().map(|_| "<redacted>"))
            .field("encryption", &self.encryption)
            .finish()
    }
}

/// Encryption variant selector for [`WorkbookSaveOptions`].
///
/// `Default` picks Agile AES-256 for `.xlsx` and RC4 CryptoAPI 128-bit
/// for `.xls`, matching Excel's current defaults. Callers who need
/// deterministic output (reproducible builds, golden files) can pin
/// an explicit variant.
#[derive(Debug, Default, Clone)]
#[non_exhaustive]
pub enum EncryptionProfile {
    /// Format-appropriate default: Agile-256 for `.xlsx`, RC4
    /// CryptoAPI 128-bit for `.xls`.
    #[default]
    Default,

    /// OOXML ECMA-376 Agile (AES-CBC + HMAC-SHA*). Modern default.
    OoxmlAgile { key_bits: u32, spin_count: u32 },
    /// OOXML ECMA-376 Standard (AES-ECB). Office 2007 compatibility.
    OoxmlStandard { key_bits: u32 },
    /// OOXML Binary Document RC4 CryptoAPI. Rare; for legacy readers.
    OoxmlBinaryRc4 { key_bits: u32 },

    /// XLS RC4 CryptoAPI (SHA-1 KDF, Excel XP+).
    XlsRc4CryptoApi { key_bits: u32 },
    /// XLS Legacy RC4 (MD5 KDF, Excel 97/2000).
    XlsRc4Legacy,
    /// XLS XOR Obfuscation (Excel 95-era, 15-char password cap).
    XlsXor,
}

/// Extension trait for Workbook to add file I/O
pub trait WorkbookExt {
    /// Open a workbook from a file, auto-detecting workbook formats from content.
    /// CSV files use the `.csv` extension as a fallback because they have no magic bytes.
    fn open<P: AsRef<Path>>(path: P) -> Result<Workbook>;

    /// Open a workbook from a file with explicit options
    /// (for password-protected files, etc.).
    fn open_with<P: AsRef<Path>>(path: P, opts: &WorkbookOpenOptions) -> Result<Workbook>;

    /// Open a workbook from bytes, auto-detecting the format (XLSX, XLSB, or XLS)
    fn from_bytes(bytes: &[u8]) -> Result<Workbook>;

    /// Open a workbook from bytes with explicit options.
    fn from_bytes_with(bytes: &[u8], opts: &WorkbookOpenOptions) -> Result<Workbook>;

    /// Save the workbook to a file.
    ///
    /// Form-control state is synchronized into linked cells in the output,
    /// replacing existing values and formulas there; the caller's workbook is
    /// left unchanged.
    fn save<P: AsRef<Path>>(&self, path: P) -> Result<()>;

    /// Save the workbook to a file with explicit options.
    ///
    /// Form-control state is synchronized into linked cells in the output,
    /// replacing existing values and formulas there; the caller's workbook is
    /// left unchanged.
    /// (to write an encrypted file, etc.).
    fn save_with<P: AsRef<Path>>(&self, path: P, opts: &WorkbookSaveOptions) -> Result<()>;
}

impl WorkbookExt for Workbook {
    fn open_with<P: AsRef<Path>>(path: P, opts: &WorkbookOpenOptions) -> Result<Workbook> {
        let path = path.as_ref();
        let pw = opts.password.as_deref();
        let vs = opts.try_velvet_sweatshop;
        let skip_ic = opts.skip_integrity_check;

        let format = detect_format_from_path(path)?;

        if (pw.is_some() || vs) && format == FileFormat::Xls {
            match XlsxReader::read_file_with_options(path, pw, vs, skip_ic) {
                Ok(wb) => return Ok(wb),
                Err(XlsxError::InvalidFormat(msg)) if is_ooxml_envelope_probe_miss(&msg) => {}
                Err(e) => return Err(Error::other(e.to_string())),
            }
        }

        match format {
            FileFormat::Xlsx => XlsxReader::read_file_with_options(path, pw, vs, skip_ic)
                .map_err(|e| Error::other(e.to_string())),
            #[cfg(feature = "xls")]
            FileFormat::Xls => XlsReader::read_file_with_password(path, pw, vs)
                .map_err(|e| Error::other(e.to_string())),
            #[cfg(not(feature = "xls"))]
            FileFormat::Xls => Err(Error::other(
                "XLS format detected but the 'xls' feature is not enabled",
            )),
            #[cfg(feature = "xlsb")]
            FileFormat::Xlsb => {
                XlsbReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xlsb"))]
            FileFormat::Xlsb => Err(Error::other(
                "XLSB format detected but the 'xlsb' feature is not enabled",
            )),
            FileFormat::Unknown if path_has_xlsx_family_extension(path) => {
                XlsxReader::read_file_with_options(path, pw, vs, skip_ic)
                    .map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xls")]
            FileFormat::Unknown if path_has_extension(path, "xls") => {
                XlsReader::read_file_with_password(path, pw, vs)
                    .map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xls"))]
            FileFormat::Unknown if path_has_extension(path, "xls") => Err(Error::other(
                "XLS format detected but the 'xls' feature is not enabled",
            )),
            #[cfg(feature = "xlsb")]
            FileFormat::Unknown if path_has_extension(path, "xlsb") => {
                XlsbReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xlsb"))]
            FileFormat::Unknown if path_has_extension(path, "xlsb") => Err(Error::other(
                "XLSB format detected but the 'xlsb' feature is not enabled",
            )),
            FileFormat::Unknown if path_has_extension(path, "csv") => read_csv_workbook(path),
            FileFormat::Unknown => Err(unsupported_file_format(path)),
        }
    }

    fn from_bytes_with(bytes: &[u8], opts: &WorkbookOpenOptions) -> Result<Workbook> {
        let pw = opts.password.as_deref();
        let vs = opts.try_velvet_sweatshop;
        let skip_ic = opts.skip_integrity_check;

        // Encrypted XLSX files masquerade as XLS to detect_format because
        // both share the CFB magic header. When a password is supplied
        // OR the sentinel retry is allowed, try the XLSX path first (it
        // handles both plain ZIPs and CFB-wrapped encrypted envelopes),
        // then fall back to XLS only if the bytes turn out not to be an
        // OOXML envelope. The fall-through is intentionally narrow:
        // matching only the "this CFB has no /EncryptionInfo" failure
        // mode so that genuine OOXML errors (BadPassword, Encrypted,
        // UnsupportedEncryption, crypto-layer InvalidFormat) propagate
        // back to the caller instead of being silently retried as XLS.
        if (pw.is_some() || vs) && is_cfb_magic(bytes) {
            match XlsxReader::read_bytes_with_options(bytes, pw, vs, skip_ic) {
                Ok(wb) => return Ok(wb),
                Err(XlsxError::InvalidFormat(msg)) if is_ooxml_envelope_probe_miss(&msg) => {}
                Err(e) => return Err(Error::other(e.to_string())),
            }
        }
        match detect_format(bytes) {
            FileFormat::Xlsx => XlsxReader::read_bytes_with_password(bytes, pw, vs)
                .map_err(|e| Error::other(e.to_string())),
            #[cfg(feature = "xls")]
            FileFormat::Xls => {
                let cursor = Cursor::new(bytes);
                XlsReader::read_with_password(cursor, pw, vs)
                    .map_err(|e| Error::other(e.to_string()))
            }
            _ => Self::from_bytes(bytes),
        }
    }

    fn save_with<P: AsRef<Path>>(&self, path: P, opts: &WorkbookSaveOptions) -> Result<()> {
        let Some(password) = opts.password.as_deref() else {
            return self.save(path);
        };

        let path = path.as_ref();
        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_lowercase());

        match extension.as_deref() {
            Some("xlsx") | Some("xlsm") | Some("xltx") | Some("xltm") => {
                let xlsx_profile = match &opts.encryption {
                    EncryptionProfile::Default => {
                        duke_sheets_xlsx::EncryptionProfile::agile_default()
                    }
                    EncryptionProfile::OoxmlAgile {
                        key_bits,
                        spin_count,
                    } => duke_sheets_xlsx::EncryptionProfile::Agile {
                        key_bits: *key_bits,
                        spin_count: *spin_count,
                    },
                    EncryptionProfile::OoxmlStandard { key_bits } => {
                        duke_sheets_xlsx::EncryptionProfile::Standard {
                            key_bits: *key_bits,
                        }
                    }
                    other => {
                        return Err(Error::other(format!(
                            "OOXML write does not yet support encryption profile {other:?}"
                        )));
                    }
                };
                let snapshot = self.synchronized_for_save();
                let workbook = snapshot.as_ref().unwrap_or(self);
                XlsxWriter::write_file_encrypted(workbook, path, password, &xlsx_profile)
                    .map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xls")]
            Some("xls") => {
                let variant = match &opts.encryption {
                    EncryptionProfile::Default => {
                        duke_sheets_crypto::xls::XlsEncryptionVariant::Rc4CryptoApi {
                            key_bits: 128,
                        }
                    }
                    EncryptionProfile::XlsRc4CryptoApi { key_bits } => {
                        duke_sheets_crypto::xls::XlsEncryptionVariant::Rc4CryptoApi {
                            key_bits: *key_bits,
                        }
                    }
                    EncryptionProfile::XlsRc4Legacy => {
                        duke_sheets_crypto::xls::XlsEncryptionVariant::Rc4Legacy
                    }
                    EncryptionProfile::XlsXor => duke_sheets_crypto::xls::XlsEncryptionVariant::Xor,
                    other => {
                        return Err(Error::other(format!(
                            "XLS write does not support encryption profile {other:?}"
                        )));
                    }
                };
                let snapshot = self.synchronized_for_save();
                let workbook = snapshot.as_ref().unwrap_or(self);
                XlsWriter::write_file_encrypted(workbook, path, password, variant)
                    .map_err(|e| Error::other(e.to_string()))
            }
            _ => Err(Error::other(format!(
                "encrypted save is only implemented for .xlsx-family and .xls extensions; got {}",
                path.display()
            ))),
        }
    }

    fn open<P: AsRef<Path>>(path: P) -> Result<Workbook> {
        let path = path.as_ref();
        match detect_format_from_path(path)? {
            FileFormat::Xlsx => {
                XlsxReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xls")]
            FileFormat::Xls => XlsReader::read_file(path).map_err(|e| Error::other(e.to_string())),
            #[cfg(not(feature = "xls"))]
            FileFormat::Xls => Err(Error::other(
                "XLS format detected but the 'xls' feature is not enabled",
            )),
            #[cfg(feature = "xlsb")]
            FileFormat::Xlsb => {
                XlsbReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xlsb"))]
            FileFormat::Xlsb => Err(Error::other(
                "XLSB format detected but the 'xlsb' feature is not enabled",
            )),
            FileFormat::Unknown if path_has_xlsx_family_extension(path) => {
                XlsxReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xls")]
            FileFormat::Unknown if path_has_extension(path, "xls") => {
                XlsReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xls"))]
            FileFormat::Unknown if path_has_extension(path, "xls") => Err(Error::other(
                "XLS format detected but the 'xls' feature is not enabled",
            )),
            #[cfg(feature = "xlsb")]
            FileFormat::Unknown if path_has_extension(path, "xlsb") => {
                XlsbReader::read_file(path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xlsb"))]
            FileFormat::Unknown if path_has_extension(path, "xlsb") => Err(Error::other(
                "XLSB format detected but the 'xlsb' feature is not enabled",
            )),
            FileFormat::Unknown if path_has_extension(path, "csv") => read_csv_workbook(path),
            FileFormat::Unknown => Err(unsupported_file_format(path)),
        }
    }

    fn from_bytes(bytes: &[u8]) -> Result<Workbook> {
        match detect_format(bytes) {
            FileFormat::Xlsx => {
                let cursor = Cursor::new(bytes);
                XlsxReader::read(cursor).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xls")]
            FileFormat::Xls => {
                let cursor = Cursor::new(bytes);
                XlsReader::read(cursor).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xls"))]
            FileFormat::Xls => Err(Error::other(
                "XLS format detected but the 'xls' feature is not enabled",
            )),
            #[cfg(feature = "xlsb")]
            FileFormat::Xlsb => {
                let cursor = Cursor::new(bytes);
                XlsbReader::read(cursor).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xlsb"))]
            FileFormat::Xlsb => Err(Error::other(
                "XLSB format detected but the 'xlsb' feature is not enabled",
            )),
            FileFormat::Unknown => Err(Error::other(
                "Unable to detect file format from bytes (expected XLSX, XLSB, or XLS magic bytes)",
            )),
        }
    }

    fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let path = path.as_ref();
        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_lowercase());

        match extension.as_deref() {
            Some("xlsx") => {
                let snapshot = self.synchronized_for_save();
                let workbook = snapshot.as_ref().unwrap_or(self);
                XlsxWriter::write_file(workbook, path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(feature = "xls")]
            Some("xls") => {
                let snapshot = self.synchronized_for_save();
                let workbook = snapshot.as_ref().unwrap_or(self);
                XlsWriter::write_file(workbook, path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xls"))]
            Some("xls") => Err(Error::other("XLS writing requires the 'xls' feature")),
            #[cfg(feature = "xlsb")]
            Some("xlsb") => {
                let snapshot = self.synchronized_for_save();
                let workbook = snapshot.as_ref().unwrap_or(self);
                XlsbWriter::write_file(workbook, path).map_err(|e| Error::other(e.to_string()))
            }
            #[cfg(not(feature = "xlsb"))]
            Some("xlsb") => Err(Error::other("XLSB writing requires the 'xlsb' feature")),
            Some("csv") => {
                let snapshot = self.synchronized_for_save();
                let workbook = snapshot.as_ref().unwrap_or(self);
                if let Some(sheet) = workbook.worksheet(0) {
                    CsvWriter::write_file(sheet, path, &CsvWriteOptions::default())
                        .map_err(|e| Error::other(e.to_string()))
                } else {
                    Err(Error::other("No worksheets to save"))
                }
            }
            _ => Err(Error::other(format!(
                "Unsupported file format: {}",
                path.display()
            ))),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn make_xlsb_bytes() -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("xl/workbook.bin", options).unwrap();
            zip.write_all(b"fake xlsb content").unwrap();
            zip.finish().unwrap();
        }
        buf
    }

    fn make_sample_workbook() -> Workbook {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "hello").unwrap();
        ws.set_cell_value("B1", 42.0).unwrap();
        ws.set_cell_value("C1", true).unwrap();
        wb
    }

    fn assert_sample_workbook(wb: &Workbook) {
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::string("hello"));
        assert_eq!(ws.get_value_at(0, 1), CellValue::Number(42.0));
        assert_eq!(ws.get_value_at(0, 2), CellValue::Boolean(true));
    }

    #[test]
    fn detect_format_identifies_xlsb() {
        let buf = make_xlsb_bytes();
        assert_eq!(detect_format(&buf), FileFormat::Xlsb);
    }

    #[test]
    fn detect_format_xlsx_not_misidentified_as_xlsb() {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<workbook/>").unwrap();
            zip.finish().unwrap();
        }
        assert_eq!(detect_format(&buf), FileFormat::Xlsx);
    }

    #[test]
    #[cfg(not(feature = "xlsb"))]
    fn from_bytes_rejects_xlsb_without_feature() {
        let buf = make_xlsb_bytes();
        let result = Workbook::from_bytes(&buf);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            err_msg.contains("xlsb"),
            "Error should mention xlsb, got: {err_msg}"
        );
    }

    #[test]
    #[cfg(feature = "xlsb")]
    fn from_bytes_reads_xlsb() {
        let buf = make_xlsb_bytes();
        let result = Workbook::from_bytes(&buf);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            !err_msg.contains("Unsupported"),
            "Should attempt to parse, not reject. Got: {err_msg}"
        );
    }

    #[test]
    #[cfg(not(feature = "xlsb"))]
    fn open_rejects_xlsb_without_feature() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsb");
        std::fs::write(&path, b"anything").unwrap();
        let result = Workbook::open(&path);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            err_msg.contains("xlsb"),
            "Error should mention xlsb, got: {err_msg}"
        );
    }

    #[test]
    #[cfg(feature = "xlsb")]
    fn open_reads_xlsb_extension() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsb");
        std::fs::write(&path, b"anything").unwrap();
        let result = Workbook::open(&path);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            !err_msg.contains("Unsupported"),
            "Should attempt to parse, not reject. Got: {err_msg}"
        );
    }

    #[test]
    fn open_detects_xlsx_content_with_xlsb_extension() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsx");
        let mismatched = dir.path().join("mismatched.xlsb");
        make_sample_workbook().save(&source).unwrap();
        std::fs::copy(&source, &mismatched).unwrap();

        let opened = Workbook::open(&mismatched).expect("open xlsx content from .xlsb path");
        assert_sample_workbook(&opened);
    }

    #[test]
    #[cfg(feature = "xls")]
    fn open_detects_xls_content_with_xlsx_extension() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xls");
        let mismatched = dir.path().join("mismatched.xlsx");
        make_sample_workbook().save(&source).unwrap();
        std::fs::copy(&source, &mismatched).unwrap();

        let opened = Workbook::open(&mismatched).expect("open xls content from .xlsx path");
        assert_sample_workbook(&opened);
    }

    #[test]
    #[cfg(feature = "xlsb")]
    fn open_detects_xlsb_content_with_xlsx_extension() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsb");
        let mismatched = dir.path().join("mismatched.xlsx");
        make_sample_workbook().save(&source).unwrap();
        std::fs::copy(&source, &mismatched).unwrap();

        let opened = Workbook::open(&mismatched).expect("open xlsb content from .xlsx path");
        assert_sample_workbook(&opened);
    }

    #[test]
    #[cfg(feature = "xlsb")]
    fn open_with_detects_xlsb_content_with_xlsx_extension() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsb");
        let mismatched = dir.path().join("mismatched.xlsx");
        make_sample_workbook().save(&source).unwrap();
        std::fs::copy(&source, &mismatched).unwrap();

        let opened = Workbook::open_with(&mismatched, &WorkbookOpenOptions::default())
            .expect("open_with xlsb content from .xlsx path");
        assert_sample_workbook(&opened);
    }

    #[test]
    fn open_keeps_csv_extension_fallback_for_unknown_magic() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("sample.csv");
        std::fs::write(&path, "name,value\nhello,42\n").unwrap();

        let opened = Workbook::open(&path).expect("open csv fallback");
        let ws = opened.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::string("name"));
        assert_eq!(ws.get_value_at(1, 1), CellValue::Number(42.0));
    }

    #[test]
    #[cfg(feature = "xlsb")]
    fn save_xlsb_roundtrip() {
        let wb = make_sample_workbook();

        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("roundtrip.xlsb");
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_sample_workbook(&wb2);
    }

    #[test]
    fn workbook_open_options_debug_redacts_password() {
        let opts = WorkbookOpenOptions::new().password("hunter2");
        let dbg = format!("{opts:?}");
        assert!(
            !dbg.contains("hunter2"),
            "debug output must not leak password: {dbg}"
        );
        assert!(
            dbg.contains("redacted"),
            "should mark password as redacted: {dbg}"
        );
    }

    #[test]
    fn workbook_save_options_debug_redacts_password() {
        let opts = WorkbookSaveOptions::new().password("hunter2");
        let dbg = format!("{opts:?}");
        assert!(
            !dbg.contains("hunter2"),
            "debug output must not leak password: {dbg}"
        );
    }

    #[test]
    fn workbook_open_options_default_has_no_password() {
        let opts = WorkbookOpenOptions::default();
        assert!(opts.password.is_none());
    }

    /// Phase 0 contract: `_with` methods accept an options struct and
    /// delegate to the plain method. Behavior doesn't change until the
    /// crypto phases wire the password through.
    #[test]
    #[cfg(feature = "xlsb")]
    fn from_bytes_with_delegates_to_from_bytes() {
        let buf = make_xlsb_bytes();
        let via_plain = Workbook::from_bytes(&buf);
        let via_with = Workbook::from_bytes_with(&buf, &WorkbookOpenOptions::default());
        // Both should produce the same error (bad xlsb content) or the
        // same success; we compare the error-ness since the test fixture
        // isn't a real xlsb.
        assert_eq!(via_plain.is_err(), via_with.is_err());
    }
}
