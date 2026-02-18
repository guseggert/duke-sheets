//! Workbook handle for manipulating a LibreOffice spreadsheet document.

use libreoffice_urp::connection::UrpConnection;
use libreoffice_urp::interface;
use libreoffice_urp::proxy::{self, UnoProxy};
use libreoffice_urp::types::{Any, Type, UnoValue, type_names};

use crate::error::{BridgeError, Result};
use crate::uno_types::{self, BorderLine2, CellAddress, Locale, StyleSpec};

/// A cell value that can be sent to/from LibreOffice.
#[derive(Debug, Clone)]
pub enum CellValue {
    Empty,
    Number(f64),
    String(String),
    Formula(String),
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::String(s.to_string())
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::String(s)
    }
}

/// A handle to an open spreadsheet document in LibreOffice.
pub struct Workbook<'a> {
    conn: &'a mut UrpConnection,
    doc: UnoProxy,
}

// ============================================================================
// Helper: construct a UNO PropertyValue struct
// ============================================================================

/// Build a `com.sun.star.beans.PropertyValue` as a `UnoValue::Struct`.
///
/// PropertyValue members: Name(string), Handle(long), Value(any), State(enum).
fn make_property_value(name: &str, value: UnoValue, type_desc: Type) -> UnoValue {
    UnoValue::Struct(vec![
        UnoValue::String(name.to_string()),
        UnoValue::Long(0),
        UnoValue::Any(Box::new(Any { type_desc, value })),
        UnoValue::Enum(0), // DIRECT_VALUE
    ])
}

impl<'a> Workbook<'a> {
    pub(crate) fn new(conn: &'a mut UrpConnection, doc: UnoProxy) -> Self {
        Self { conn, doc }
    }

    /// Get direct access to the URP connection (for advanced operations).
    pub fn conn(&mut self) -> &mut UrpConnection {
        self.conn
    }

    /// Get the document proxy.
    pub fn doc(&self) -> &UnoProxy {
        &self.doc
    }

    // ========================================================================
    // Internal helpers: proxy acquisition
    // ========================================================================

    /// Query interface on the document itself.
    async fn doc_qi(&mut self, type_name: &str) -> Result<UnoProxy> {
        let iface_type = Type::interface(type_name);
        self.conn
            .query_interface(&self.doc, iface_type)
            .await?
            .ok_or_else(|| {
                BridgeError::OperationFailed(format!(
                    "document does not support {type_name}"
                ))
            })
    }

    /// Get the XSpreadsheets collection proxy.
    async fn get_sheets_proxy(&mut self) -> Result<UnoProxy> {
        let ssd_proxy = self.doc_qi(type_names::X_SPREADSHEET_DOCUMENT).await?;
        let method = interface::get_sheets();
        let result = self.conn.call(&ssd_proxy, &method, &[]).await?;
        let oid = proxy::extract_oid_from_return(&result)
            .ok_or_else(|| BridgeError::OperationFailed("getSheets returned null".into()))?;
        Ok(UnoProxy::new(oid, Type::interface(type_names::X_SPREADSHEETS)))
    }

    /// Get a sheet proxy by its zero-based index, typed for the given interface.
    async fn get_sheet_proxy_as(
        &mut self,
        sheet_index: i32,
        iface_type_name: &str,
    ) -> Result<UnoProxy> {
        let sheets = self.get_sheets_proxy().await?;
        // XIndexAccess on sheets
        let idx_type = Type::interface(type_names::X_INDEX_ACCESS);
        let sheets_idx = UnoProxy::new(sheets.oid, idx_type.clone());
        let sheets_proxy = self
            .conn
            .query_interface(&sheets_idx, idx_type)
            .await?
            .unwrap_or(sheets_idx);

        let method = interface::get_by_index();
        let result = self
            .conn
            .call(&sheets_proxy, &method, &[UnoValue::Long(sheet_index)])
            .await?;
        let oid = proxy::extract_oid_from_return(&result)
            .ok_or_else(|| BridgeError::OperationFailed("getByIndex returned null".into()))?;

        let target_type = Type::interface(iface_type_name);
        let raw = UnoProxy::new(oid, target_type.clone());
        let proxy = self
            .conn
            .query_interface(&raw, target_type)
            .await?
            .unwrap_or(raw);
        Ok(proxy)
    }

    /// Get a sheet proxy typed as XCellRange (for getCellByPosition / getCellRangeByName).
    async fn get_sheet_cell_range(&mut self, sheet_index: i32) -> Result<UnoProxy> {
        self.get_sheet_proxy_as(sheet_index, type_names::X_CELL_RANGE)
            .await
    }

    /// Extract an interface OID from a return value, or error.
    fn require_oid(value: &UnoValue, context: &str) -> Result<String> {
        proxy::extract_oid_from_return(value)
            .ok_or_else(|| BridgeError::OperationFailed(format!("{context} returned null")))
    }

    /// Query interface on an arbitrary proxy.
    async fn qi(&mut self, proxy: &UnoProxy, type_name: &str) -> Result<UnoProxy> {
        let iface_type = Type::interface(type_name);
        let raw = UnoProxy::new(proxy.oid.clone(), iface_type.clone());
        Ok(self
            .conn
            .query_interface(&raw, iface_type)
            .await?
            .unwrap_or(raw))
    }

    // ========================================================================
    // Cell access
    // ========================================================================

    /// Get a cell proxy by its position (column, row) on the first sheet.
    async fn get_cell(&mut self, col: i32, row: i32) -> Result<UnoProxy> {
        self.get_cell_on_sheet(0, col, row).await
    }

    /// Get a cell proxy by position on a specific sheet index.
    pub async fn get_cell_on_sheet(
        &mut self,
        sheet_index: i32,
        col: i32,
        row: i32,
    ) -> Result<UnoProxy> {
        let sheet = self.get_sheet_cell_range(sheet_index).await?;
        let method = interface::get_cell_by_position();
        let result = self
            .conn
            .call(&sheet, &method, &[UnoValue::Long(col), UnoValue::Long(row)])
            .await?;
        let oid = Self::require_oid(&result, "getCellByPosition")?;
        Ok(UnoProxy::new(oid, Type::interface(type_names::X_CELL)))
    }

    /// Parse a cell reference like "A1" into (col, row) zero-indexed.
    pub fn parse_cell_ref(cell_ref: &str) -> Result<(i32, i32)> {
        let mut col: i32 = 0;
        let mut i = 0;
        let chars: Vec<char> = cell_ref.chars().collect();

        // Parse column letters
        while i < chars.len() && chars[i].is_ascii_alphabetic() {
            col = col * 26 + (chars[i].to_ascii_uppercase() as i32 - 'A' as i32 + 1);
            i += 1;
        }
        col -= 1; // Zero-indexed

        if col < 0 {
            return Err(BridgeError::InvalidCellRef(cell_ref.to_string()));
        }

        // Parse row number
        let row_str = &cell_ref[i..];
        let row: i32 = row_str
            .parse::<i32>()
            .map_err(|_| BridgeError::InvalidCellRef(cell_ref.to_string()))?
            - 1; // Zero-indexed

        if row < 0 {
            return Err(BridgeError::InvalidCellRef(cell_ref.to_string()));
        }

        Ok((col, row))
    }

    // ========================================================================
    // Cell data (existing API, unchanged)
    // ========================================================================

    /// Set a cell's value (number, string, formula, or empty).
    pub async fn set_cell_value(
        &mut self,
        cell_ref: &str,
        value: impl Into<CellValue>,
    ) -> Result<()> {
        let (col, row) = Self::parse_cell_ref(cell_ref)?;
        let cell = self.get_cell(col, row).await?;
        self.set_cell_value_on_proxy(&cell, value.into()).await
    }

    /// Set a cell's value given its proxy.
    pub(crate) async fn set_cell_value_on_proxy(
        &mut self,
        cell: &UnoProxy,
        cv: CellValue,
    ) -> Result<()> {
        match cv {
            CellValue::Number(n) => {
                let method = interface::cell_set_value();
                self.conn.call(cell, &method, &[UnoValue::Double(n)]).await?;
            }
            CellValue::String(s) => {
                let text_proxy = self.qi(cell, type_names::X_TEXT_RANGE).await?;
                let method = interface::text_range_set_string();
                self.conn
                    .call(&text_proxy, &method, &[UnoValue::String(s)])
                    .await?;
            }
            CellValue::Formula(f) => {
                let method = interface::cell_set_formula();
                self.conn
                    .call(cell, &method, &[UnoValue::String(f)])
                    .await?;
            }
            CellValue::Empty => {
                let method = interface::cell_set_formula();
                self.conn
                    .call(cell, &method, &[UnoValue::String(String::new())])
                    .await?;
            }
        }
        Ok(())
    }

    /// Set a cell's formula.
    pub async fn set_cell_formula(&mut self, cell_ref: &str, formula: &str) -> Result<()> {
        let (col, row) = Self::parse_cell_ref(cell_ref)?;
        let cell = self.get_cell(col, row).await?;
        let method = interface::cell_set_formula();
        self.conn
            .call(&cell, &method, &[UnoValue::String(formula.to_string())])
            .await?;
        Ok(())
    }

    /// Get a cell's computed numeric value.
    pub async fn get_cell_value(&mut self, cell_ref: &str) -> Result<f64> {
        let (col, row) = Self::parse_cell_ref(cell_ref)?;
        let cell = self.get_cell(col, row).await?;
        let method = interface::cell_get_value();
        let result = self.conn.call(&cell, &method, &[]).await?;
        match result {
            UnoValue::Double(d) => Ok(d),
            other => Err(BridgeError::OperationFailed(format!(
                "getValue returned unexpected type: {other:?}"
            ))),
        }
    }

    /// Get a cell's formula string.
    pub async fn get_cell_formula(&mut self, cell_ref: &str) -> Result<String> {
        let (col, row) = Self::parse_cell_ref(cell_ref)?;
        let cell = self.get_cell(col, row).await?;
        let method = interface::cell_get_formula();
        let result = self.conn.call(&cell, &method, &[]).await?;
        match result {
            UnoValue::String(s) => Ok(s),
            other => Err(BridgeError::OperationFailed(format!(
                "getFormula returned unexpected type: {other:?}"
            ))),
        }
    }

    /// Get a cell's string value (via XTextRange::getString).
    pub async fn get_cell_string(&mut self, cell_ref: &str) -> Result<String> {
        let (col, row) = Self::parse_cell_ref(cell_ref)?;
        let cell = self.get_cell(col, row).await?;
        let text_proxy = self.qi(&cell, type_names::X_TEXT_RANGE).await?;
        let method = interface::text_range_get_string();
        let result = self.conn.call(&text_proxy, &method, &[]).await?;
        match result {
            UnoValue::String(s) => Ok(s),
            other => Err(BridgeError::OperationFailed(format!(
                "getString returned unexpected type: {other:?}"
            ))),
        }
    }

    // ========================================================================
    // Sheet management
    // ========================================================================

    /// Set the name of a sheet by index.
    pub async fn set_sheet_name(&mut self, sheet_index: i32, name: &str) -> Result<()> {
        let sheet = self
            .get_sheet_proxy_as(sheet_index, type_names::X_NAMED)
            .await?;
        let method = interface::set_name();
        self.conn
            .call(&sheet, &method, &[UnoValue::String(name.to_string())])
            .await?;
        Ok(())
    }

    /// Add a new sheet with the given name at the end.
    pub async fn add_sheet(&mut self, name: &str) -> Result<()> {
        let sheets = self.get_sheets_proxy().await?;
        // Get count first
        let idx_proxy = self.qi(&sheets, type_names::X_INDEX_ACCESS).await?;
        let count_method = interface::get_count();
        let count_result = self.conn.call(&idx_proxy, &count_method, &[]).await?;
        let count = match count_result {
            UnoValue::Long(n) => n,
            _ => 0,
        };

        let method = interface::insert_new_by_name();
        self.conn
            .call(
                &sheets,
                &method,
                &[
                    UnoValue::String(name.to_string()),
                    UnoValue::Short(count as i16),
                ],
            )
            .await?;
        Ok(())
    }

    /// Get the number of sheets.
    pub async fn sheet_count(&mut self) -> Result<i32> {
        let sheets = self.get_sheets_proxy().await?;
        let idx_proxy = self.qi(&sheets, type_names::X_INDEX_ACCESS).await?;
        let method = interface::get_count();
        let result = self.conn.call(&idx_proxy, &method, &[]).await?;
        match result {
            UnoValue::Long(n) => Ok(n),
            _ => Ok(0),
        }
    }

    // ========================================================================
    // Property access (XPropertySet)
    // ========================================================================

    /// Set a property on a proxy object (via XPropertySet::setPropertyValue).
    pub async fn set_property(
        &mut self,
        proxy: &UnoProxy,
        name: &str,
        value: UnoValue,
    ) -> Result<()> {
        let ps_proxy = self.qi(proxy, type_names::X_PROPERTY_SET).await?;
        let method = interface::set_property_value();
        self.conn
            .call(
                &ps_proxy,
                &method,
                &[UnoValue::String(name.to_string()), value],
            )
            .await?;
        Ok(())
    }

    /// Get a property value from a proxy object (via XPropertySet::getPropertyValue).
    pub async fn get_property(
        &mut self,
        proxy: &UnoProxy,
        name: &str,
    ) -> Result<UnoValue> {
        let ps_proxy = self.qi(proxy, type_names::X_PROPERTY_SET).await?;
        let method = interface::get_property_value();
        self.conn
            .call(&ps_proxy, &method, &[UnoValue::String(name.to_string())])
            .await
            .map_err(|e| e.into())
    }

    // ========================================================================
    // Cell range operations
    // ========================================================================

    /// Get a cell range by A1 notation on a sheet (e.g., "A1:B5").
    pub async fn get_cell_range_by_name(
        &mut self,
        sheet_index: i32,
        range_ref: &str,
    ) -> Result<UnoProxy> {
        let sheet = self.get_sheet_cell_range(sheet_index).await?;
        let method = interface::get_cell_range_by_name();
        let result = self
            .conn
            .call(
                &sheet,
                &method,
                &[UnoValue::String(range_ref.to_string())],
            )
            .await?;
        let oid = Self::require_oid(&result, "getCellRangeByName")?;
        Ok(UnoProxy::new(
            oid,
            Type::interface(type_names::X_CELL_RANGE),
        ))
    }

    /// Merge a range of cells on a sheet.
    pub async fn merge_range(&mut self, sheet_index: i32, range_ref: &str) -> Result<()> {
        let range = self.get_cell_range_by_name(sheet_index, range_ref).await?;
        let mergeable = self.qi(&range, type_names::X_MERGEABLE).await?;
        let method = interface::merge();
        self.conn
            .call(&mergeable, &method, &[UnoValue::Bool(true)])
            .await?;
        Ok(())
    }

    // ========================================================================
    // Comments / Annotations
    // ========================================================================

    /// Add a comment (annotation) to a cell.
    ///
    /// Note: `author` is attempted via `setAuthor` but may not be supported
    /// by all LO versions (same as Python code which wraps in try/except).
    pub async fn add_comment(
        &mut self,
        sheet_index: i32,
        cell_ref: &str,
        text: &str,
        _author: Option<&str>,
    ) -> Result<()> {
        let (col, row) = Self::parse_cell_ref(cell_ref)?;

        // Get annotations supplier
        let sheet = self
            .get_sheet_proxy_as(sheet_index, type_names::X_SHEET_ANNOTATIONS_SUPPLIER)
            .await?;
        let method = interface::get_annotations();
        let result = self.conn.call(&sheet, &method, &[]).await?;
        let ann_oid = Self::require_oid(&result, "getAnnotations")?;
        let ann_proxy = UnoProxy::new(
            ann_oid,
            Type::interface(type_names::X_SHEET_ANNOTATIONS),
        );

        // Create CellAddress struct
        let cell_addr = CellAddress::new(sheet_index as i16, col, row);
        let method = interface::annotations_insert_new();
        self.conn
            .call(
                &ann_proxy,
                &method,
                &[
                    cell_addr.to_uno(),
                    UnoValue::String(text.to_string()),
                ],
            )
            .await?;

        // Note: setAuthor is not reliably supported in LO — the Python code
        // wraps it in try/except. We skip it for now.

        Ok(())
    }

    // ========================================================================
    // Row / Column dimensions
    // ========================================================================

    /// Set the height of a row on a sheet (height in points).
    pub async fn set_row_height(
        &mut self,
        sheet_index: i32,
        row: i32,
        height_pt: f64,
    ) -> Result<()> {
        let sheet = self
            .get_sheet_proxy_as(sheet_index, type_names::X_COLUMN_ROW_RANGE)
            .await?;
        let method = interface::get_rows();
        let result = self.conn.call(&sheet, &method, &[]).await?;
        let rows_oid = Self::require_oid(&result, "getRows")?;
        let rows_proxy = UnoProxy::new(rows_oid, Type::interface(type_names::X_TABLE_ROWS));

        // getByIndex on rows
        let rows_idx = self.qi(&rows_proxy, type_names::X_INDEX_ACCESS).await?;
        let get_by_idx = interface::get_by_index();
        let row_result = self
            .conn
            .call(&rows_idx, &get_by_idx, &[UnoValue::Long(row)])
            .await?;
        let row_oid = Self::require_oid(&row_result, "rows.getByIndex")?;
        let row_proxy = UnoProxy::new(
            row_oid,
            Type::interface(type_names::X_PROPERTY_SET),
        );

        // Height is in 1/100 mm; 1pt = 0.3528mm → multiply by 35.28
        let height_100mm = (height_pt * 35.28) as i32;
        self.set_property(
            &row_proxy,
            "Height",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::long(),
                value: UnoValue::Long(height_100mm),
            })),
        )
        .await
    }

    /// Set the width of a column on a sheet (width in character units, approximate).
    pub async fn set_column_width(
        &mut self,
        sheet_index: i32,
        col: i32,
        width_chars: f64,
    ) -> Result<()> {
        let sheet = self
            .get_sheet_proxy_as(sheet_index, type_names::X_COLUMN_ROW_RANGE)
            .await?;
        let method = interface::get_columns();
        let result = self.conn.call(&sheet, &method, &[]).await?;
        let cols_oid = Self::require_oid(&result, "getColumns")?;
        let cols_proxy = UnoProxy::new(
            cols_oid,
            Type::interface(type_names::X_TABLE_COLUMNS),
        );

        let cols_idx = self.qi(&cols_proxy, type_names::X_INDEX_ACCESS).await?;
        let get_by_idx = interface::get_by_index();
        let col_result = self
            .conn
            .call(&cols_idx, &get_by_idx, &[UnoValue::Long(col)])
            .await?;
        let col_oid = Self::require_oid(&col_result, "cols.getByIndex")?;
        let col_proxy = UnoProxy::new(
            col_oid,
            Type::interface(type_names::X_PROPERTY_SET),
        );

        // Width in 1/100 mm; 1 char ≈ 2.5mm → multiply by 250
        let width_100mm = (width_chars * 250.0) as i32;
        self.set_property(
            &col_proxy,
            "Width",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::long(),
                value: UnoValue::Long(width_100mm),
            })),
        )
        .await
    }

    /// Hide or show a row on a sheet.
    pub async fn set_row_hidden(
        &mut self,
        sheet_index: i32,
        row: i32,
        hidden: bool,
    ) -> Result<()> {
        let sheet = self
            .get_sheet_proxy_as(sheet_index, type_names::X_COLUMN_ROW_RANGE)
            .await?;
        let method = interface::get_rows();
        let result = self.conn.call(&sheet, &method, &[]).await?;
        let rows_oid = Self::require_oid(&result, "getRows")?;
        let rows_proxy = UnoProxy::new(rows_oid, Type::interface(type_names::X_TABLE_ROWS));

        let rows_idx = self.qi(&rows_proxy, type_names::X_INDEX_ACCESS).await?;
        let get_by_idx = interface::get_by_index();
        let row_result = self
            .conn
            .call(&rows_idx, &get_by_idx, &[UnoValue::Long(row)])
            .await?;
        let row_oid = Self::require_oid(&row_result, "rows.getByIndex")?;
        let row_proxy = UnoProxy::new(
            row_oid,
            Type::interface(type_names::X_PROPERTY_SET),
        );

        self.set_property(
            &row_proxy,
            "IsVisible",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::boolean(),
                value: UnoValue::Bool(!hidden),
            })),
        )
        .await
    }

    // ========================================================================
    // Number formats
    // ========================================================================

    /// Look up or register a number format string, returning its format ID.
    ///
    /// Tries `queryKey` first; if not found (-1), calls `addNew`.
    pub async fn get_or_create_number_format(&mut self, format_str: &str) -> Result<i32> {
        // Get XNumberFormatsSupplier from the document
        let nfs_proxy = self
            .doc_qi(type_names::X_NUMBER_FORMATS_SUPPLIER)
            .await?;
        let method = interface::get_number_formats();
        let result = self.conn.call(&nfs_proxy, &method, &[]).await?;
        let nf_oid = Self::require_oid(&result, "getNumberFormats")?;
        let nf_proxy = UnoProxy::new(
            nf_oid,
            Type::interface(type_names::X_NUMBER_FORMATS),
        );

        let locale = Locale::empty();

        // queryKey(format_str, locale, false)
        let query_method = interface::number_formats_query_key();
        let key_result = self
            .conn
            .call(
                &nf_proxy,
                &query_method,
                &[
                    UnoValue::String(format_str.to_string()),
                    locale.to_uno(),
                    UnoValue::Bool(false),
                ],
            )
            .await?;

        let key = match key_result {
            UnoValue::Long(n) => n,
            _ => -1,
        };

        if key != -1 {
            return Ok(key);
        }

        // addNew(format_str, locale)
        let add_method = interface::number_formats_add_new();
        let new_result = self
            .conn
            .call(
                &nf_proxy,
                &add_method,
                &[
                    UnoValue::String(format_str.to_string()),
                    locale.to_uno(),
                ],
            )
            .await?;

        match new_result {
            UnoValue::Long(n) => Ok(n),
            other => Err(BridgeError::OperationFailed(format!(
                "addNew returned unexpected type: {other:?}"
            ))),
        }
    }

    // ========================================================================
    // Cell styling
    // ========================================================================

    /// Apply a `StyleSpec` to a cell or range proxy.
    ///
    /// This mirrors the Python `_apply_style_to_cell` method, setting
    /// properties via XPropertySet.
    pub async fn apply_style(
        &mut self,
        target: &UnoProxy,
        spec: &StyleSpec,
    ) -> Result<()> {
        // Font properties
        if spec.bold {
            self.set_property(
                target,
                "CharWeight",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::float(),
                    value: UnoValue::Float(uno_types::font_weight::BOLD),
                })),
            )
            .await?;
        }
        if spec.italic {
            self.set_property(
                target,
                "CharPosture",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::short(),
                    value: UnoValue::Short(uno_types::font_slant::ITALIC),
                })),
            )
            .await?;
        }
        if let Some(ref underline) = spec.underline {
            let ul_val = uno_types::font_underline::from_name(underline);
            self.set_property(
                target,
                "CharUnderline",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::short(),
                    value: UnoValue::Short(ul_val),
                })),
            )
            .await?;
        }
        if spec.strikethrough {
            self.set_property(
                target,
                "CharStrikeout",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::short(),
                    value: UnoValue::Short(uno_types::font_strikeout::SINGLE),
                })),
            )
            .await?;
        }
        if let Some(color) = spec.font_color {
            self.set_property(
                target,
                "CharColor",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::long(),
                    value: UnoValue::Long(color),
                })),
            )
            .await?;
        }
        if let Some(size) = spec.font_size {
            self.set_property(
                target,
                "CharHeight",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::float(),
                    value: UnoValue::Float(size),
                })),
            )
            .await?;
        }
        if let Some(ref name) = spec.font_name {
            self.set_property(
                target,
                "CharFontName",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::string(),
                    value: UnoValue::String(name.clone()),
                })),
            )
            .await?;
        }
        if let Some(ref va) = spec.font_vertical_align {
            // CharEscapement: percentage shift (33=superscript, -33=subscript, 0=baseline)
            // Note: CharEscapementHeight is not supported on cell-level properties in Calc,
            // only on text portions. CharEscapement alone triggers the XLSX export to write
            // <vertAlign val="superscript|subscript"/>.
            let escapement: i16 = match va.as_str() {
                "superscript" => 33,
                "subscript" => -33,
                _ => 0,
            };
            self.set_property(
                target,
                "CharEscapement",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::short(),
                    value: UnoValue::Short(escapement),
                })),
            )
            .await?;
        }

        // Fill
        // Note: fill_gradient is kept in StyleSpec for future use but LO Calc 7.3
        // cells don't support FillStyle/FillGradient drawing properties.
        if let Some(color) = spec.fill_color {
            self.set_property(
                target,
                "CellBackColor",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::long(),
                    value: UnoValue::Long(color),
                })),
            )
            .await?;
        }

        // Alignment
        if let Some(ref h) = spec.horizontal {
            let val = uno_types::hori_justify::from_name(h);
            self.set_property(
                target,
                "HoriJustify",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::long(),
                    value: UnoValue::Long(val),
                })),
            )
            .await?;
        }
        if let Some(ref v) = spec.vertical {
            let val = uno_types::vert_justify::from_name(v);
            self.set_property(
                target,
                "VertJustify",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::long(),
                    value: UnoValue::Long(val),
                })),
            )
            .await?;
        }
        if spec.wrap_text {
            self.set_property(
                target,
                "IsTextWrapped",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::boolean(),
                    value: UnoValue::Bool(true),
                })),
            )
            .await?;
        }
        if spec.shrink_to_fit {
            self.set_property(
                target,
                "ShrinkToFit",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::boolean(),
                    value: UnoValue::Bool(true),
                })),
            )
            .await?;
        }
        if spec.rotation != 0 {
            if spec.rotation == 255 {
                // Stacked text
                self.set_property(
                    target,
                    "Orientation",
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::long(),
                        value: UnoValue::Long(1),
                    })),
                )
                .await?;
            } else {
                // Rotation in 1/100 degree
                self.set_property(
                    target,
                    "RotateAngle",
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::long(),
                        value: UnoValue::Long(spec.rotation * 100),
                    })),
                )
                .await?;
            }
        }
        if spec.indent > 0 {
            self.set_property(
                target,
                "ParaIndent",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::short(),
                    value: UnoValue::Short((spec.indent * 200) as i16),
                })),
            )
            .await?;
        }

        // Borders — all sides
        if let Some(ref style_name) = spec.border_style {
            let color = spec.border_color.unwrap_or(0x000000);
            let (line_style, line_width) =
                uno_types::border_line_style::from_name(style_name);
            let border = BorderLine2::new(color, line_style, line_width);
            for side in &["TopBorder", "BottomBorder", "LeftBorder", "RightBorder"] {
                self.set_property(
                    target,
                    side,
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::r#struct(
                            uno_types::struct_type_names::BORDER_LINE2,
                        ),
                        value: border.to_uno(),
                    })),
                )
                .await?;
            }
        }

        // Individual borders
        for (prop_name, border_opt) in [
            ("LeftBorder", &spec.left_border),
            ("RightBorder", &spec.right_border),
            ("TopBorder", &spec.top_border),
            ("BottomBorder", &spec.bottom_border),
        ] {
            if let Some((ref style_name, color)) = border_opt {
                let (line_style, line_width) =
                    uno_types::border_line_style::from_name(style_name);
                let border = BorderLine2::new(*color, line_style, line_width);
                self.set_property(
                    target,
                    prop_name,
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::r#struct(
                            uno_types::struct_type_names::BORDER_LINE2,
                        ),
                        value: border.to_uno(),
                    })),
                )
                .await?;
            }
        }

        // Number format
        if let Some(ref fmt) = spec.number_format {
            let fmt_id = self.get_or_create_number_format(fmt).await?;
            self.set_property(
                target,
                "NumberFormat",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::long(),
                    value: UnoValue::Long(fmt_id),
                })),
            )
            .await?;
        }

        Ok(())
    }

    /// Apply a `StyleSpec` to a cell identified by reference on a sheet.
    pub async fn set_cell_style(
        &mut self,
        sheet_index: i32,
        cell_ref: &str,
        spec: &StyleSpec,
    ) -> Result<()> {
        let (col, row) = Self::parse_cell_ref(cell_ref)?;
        let cell = self.get_cell_on_sheet(sheet_index, col, row).await?;
        self.apply_style(&cell, spec).await
    }

    /// Apply a `StyleSpec` to a range identified by A1 notation on a sheet.
    pub async fn set_range_style(
        &mut self,
        sheet_index: i32,
        range_ref: &str,
        spec: &StyleSpec,
    ) -> Result<()> {
        let range = self.get_cell_range_by_name(sheet_index, range_ref).await?;
        self.apply_style(&range, spec).await
    }

    // ========================================================================
    // Conditional formatting
    // ========================================================================

    /// Create a named cell style for conditional formatting and return its name.
    async fn create_cf_style(
        &mut self,
        style_name: &str,
        spec: &StyleSpec,
    ) -> Result<String> {
        // Get style families
        let sf_proxy = self
            .doc_qi(type_names::X_STYLE_FAMILIES_SUPPLIER)
            .await?;
        let method = interface::get_style_families();
        let result = self.conn.call(&sf_proxy, &method, &[]).await?;
        let sf_oid = Self::require_oid(&result, "getStyleFamilies")?;
        let sf = UnoProxy::new(sf_oid, Type::interface(type_names::X_NAME_ACCESS));

        // getByName("CellStyles")
        let method = interface::get_by_name();
        let cs_result = self
            .conn
            .call(
                &sf,
                &method,
                &[UnoValue::String("CellStyles".to_string())],
            )
            .await?;
        let cs_oid = Self::require_oid(&cs_result, "getByName(CellStyles)")?;
        let cs_proxy = UnoProxy::new(
            cs_oid,
            Type::interface(type_names::X_NAME_CONTAINER),
        );

        // doc.createInstance("com.sun.star.style.CellStyle")
        let msf_proxy = self.doc_qi(type_names::X_MULTI_SERVICE_FACTORY).await?;
        let method = interface::doc_create_instance();
        let style_result = self
            .conn
            .call(
                &msf_proxy,
                &method,
                &[UnoValue::String(
                    uno_types::struct_type_names::CELL_STYLE.to_string(),
                )],
            )
            .await?;
        let style_oid = Self::require_oid(&style_result, "createInstance(CellStyle)")?;
        let style_proxy = UnoProxy::new(
            style_oid,
            Type::interface(type_names::X_PROPERTY_SET),
        );

        // insertByName(style_name, style)
        let method = interface::insert_by_name();
        self.conn
            .call(
                &cs_proxy,
                &method,
                &[
                    UnoValue::String(style_name.to_string()),
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::interface(type_names::X_INTERFACE),
                        value: UnoValue::Interface(style_proxy.oid.clone()),
                    })),
                ],
            )
            .await?;

        // Apply style properties
        self.apply_style(&style_proxy, spec).await?;

        Ok(style_name.to_string())
    }

    /// Add a conditional format rule to a range on a sheet.
    ///
    /// - `range_ref`: A1 notation like "A1:A10"
    /// - `operator`: Condition operator name (e.g., "greater_than", "less_than")
    /// - `formula`: The formula/value string
    /// - `style_name`: Name for the CF cell style
    /// - `style`: The style to apply when condition matches
    pub async fn add_conditional_format(
        &mut self,
        sheet_index: i32,
        range_ref: &str,
        operator: &str,
        formula: &str,
        style_name: &str,
        style: &StyleSpec,
    ) -> Result<()> {
        // Create the named cell style
        self.create_cf_style(style_name, style).await?;

        // Get the *range's* ConditionalFormat property so we append to existing
        // rules rather than overwriting them.
        let range = self.get_cell_range_by_name(sheet_index, range_ref).await?;
        let cf_value = self.get_property(&range, "ConditionalFormat").await?;

        // Extract the conditional format entries OID
        let cf_oid = proxy::extract_oid_from_return(&cf_value).ok_or_else(|| {
            BridgeError::OperationFailed(
                "getPropertyValue(ConditionalFormat) returned null".into(),
            )
        })?;
        let cf_proxy = UnoProxy::new(
            cf_oid,
            Type::interface(type_names::X_SHEET_CONDITIONAL_ENTRIES),
        );

        // Build the condition properties tuple
        let op_val = uno_types::condition_operator::from_name(operator);
        let props = UnoValue::Sequence(vec![
            make_property_value(
                "Operator",
                UnoValue::Enum(op_val),
                Type::r#enum("com.sun.star.sheet.ConditionOperator"),
            ),
            make_property_value("Formula1", UnoValue::String(formula.to_string()), Type::string()),
            make_property_value(
                "StyleName",
                UnoValue::String(style_name.to_string()),
                Type::string(),
            ),
        ]);

        // addNew(props)
        let method = interface::conditional_entries_add_new();
        self.conn.call(&cf_proxy, &method, &[props]).await?;

        // Apply CF entries back to the cell range (reuse the range proxy from above)
        self.set_property(
            &range,
            "ConditionalFormat",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::interface(type_names::X_SHEET_CONDITIONAL_ENTRIES),
                value: UnoValue::Interface(cf_proxy.oid),
            })),
        )
        .await?;

        Ok(())
    }

    // ========================================================================
    // Data validation
    // ========================================================================

    /// Add data validation to a range on a sheet.
    ///
    /// Parameters mirror the Python `add_data_validation` method.
    #[allow(clippy::too_many_arguments)]
    pub async fn add_data_validation(
        &mut self,
        sheet_index: i32,
        range_ref: &str,
        validation_type: &str,
        operator: &str,
        formula1: &str,
        formula2: &str,
        allow_blank: bool,
        show_dropdown: bool,
        input_title: Option<&str>,
        input_message: Option<&str>,
        error_title: Option<&str>,
        error_message: Option<&str>,
        error_style: &str,
    ) -> Result<()> {
        let range = self.get_cell_range_by_name(sheet_index, range_ref).await?;

        // Get validation sub-object
        let validation_value = self.get_property(&range, "Validation").await?;
        let val_oid = proxy::extract_oid_from_return(&validation_value).ok_or_else(|| {
            BridgeError::OperationFailed(
                "getPropertyValue(Validation) returned null".into(),
            )
        })?;
        let val_proxy = UnoProxy::new(
            val_oid,
            Type::interface(type_names::X_PROPERTY_SET),
        );

        // Set Type
        let vtype = uno_types::validation_type::from_name(validation_type);
        self.set_property(
            &val_proxy,
            "Type",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::r#enum("com.sun.star.sheet.ValidationType"),
                value: UnoValue::Enum(vtype),
            })),
        )
        .await?;

        // Set Operator
        let vop = uno_types::condition_operator::from_name(operator);
        self.set_property(
            &val_proxy,
            "Operator",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::r#enum("com.sun.star.sheet.ConditionOperator"),
                value: UnoValue::Enum(vop),
            })),
        )
        .await?;

        // Set formulas
        if !formula1.is_empty() {
            self.set_property(
                &val_proxy,
                "Formula1",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::string(),
                    value: UnoValue::String(formula1.to_string()),
                })),
            )
            .await?;
        }
        if !formula2.is_empty() {
            self.set_property(
                &val_proxy,
                "Formula2",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::string(),
                    value: UnoValue::String(formula2.to_string()),
                })),
            )
            .await?;
        }

        // Options
        self.set_property(
            &val_proxy,
            "IgnoreBlankCells",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::boolean(),
                value: UnoValue::Bool(allow_blank),
            })),
        )
        .await?;
        self.set_property(
            &val_proxy,
            "ShowList",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::boolean(),
                value: UnoValue::Bool(show_dropdown),
            })),
        )
        .await?;

        // Input message
        if input_title.is_some() || input_message.is_some() {
            self.set_property(
                &val_proxy,
                "ShowInputMessage",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::boolean(),
                    value: UnoValue::Bool(true),
                })),
            )
            .await?;
            if let Some(title) = input_title {
                self.set_property(
                    &val_proxy,
                    "InputTitle",
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::string(),
                        value: UnoValue::String(title.to_string()),
                    })),
                )
                .await?;
            }
            if let Some(msg) = input_message {
                self.set_property(
                    &val_proxy,
                    "InputMessage",
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::string(),
                        value: UnoValue::String(msg.to_string()),
                    })),
                )
                .await?;
            }
        }

        // Error message
        if error_title.is_some() || error_message.is_some() {
            self.set_property(
                &val_proxy,
                "ShowErrorMessage",
                UnoValue::Any(Box::new(Any {
                    type_desc: Type::boolean(),
                    value: UnoValue::Bool(true),
                })),
            )
            .await?;
            if let Some(title) = error_title {
                self.set_property(
                    &val_proxy,
                    "ErrorTitle",
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::string(),
                        value: UnoValue::String(title.to_string()),
                    })),
                )
                .await?;
            }
            if let Some(msg) = error_message {
                self.set_property(
                    &val_proxy,
                    "ErrorMessage",
                    UnoValue::Any(Box::new(Any {
                        type_desc: Type::string(),
                        value: UnoValue::String(msg.to_string()),
                    })),
                )
                .await?;
            }
        }

        // Error alert style
        let alert_val = uno_types::validation_alert_style::from_name(error_style);
        self.set_property(
            &val_proxy,
            "ErrorAlertStyle",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::r#enum("com.sun.star.sheet.ValidationAlertStyle"),
                value: UnoValue::Enum(alert_val),
            })),
        )
        .await?;

        // Apply validation back to the range
        self.set_property(
            &range,
            "Validation",
            UnoValue::Any(Box::new(Any {
                type_desc: Type::interface(type_names::X_PROPERTY_SET),
                value: UnoValue::Interface(val_proxy.oid),
            })),
        )
        .await?;

        Ok(())
    }

    // ========================================================================
    // Save / Close
    // ========================================================================

    /// Save the workbook as XLSX to the given file path.
    pub async fn save(&mut self, path: &str) -> Result<()> {
        // Convert to file:// URL
        let url = if path.starts_with("file://") {
            path.to_string()
        } else {
            let abs = if path.starts_with('/') {
                path.to_string()
            } else {
                std::env::current_dir()
                    .unwrap_or_default()
                    .join(path)
                    .display()
                    .to_string()
            };
            format!("file://{abs}")
        };

        // queryInterface for XStorable
        let storable_proxy = self.doc_qi(type_names::X_STORABLE).await?;

        let filter_pv = make_property_value(
            "FilterName",
            UnoValue::String("Calc MS Excel 2007 XML".to_string()),
            Type::string(),
        );
        let overwrite_pv =
            make_property_value("Overwrite", UnoValue::Bool(true), Type::boolean());
        let props = UnoValue::Sequence(vec![filter_pv, overwrite_pv]);

        let method = interface::store_to_url();
        self.conn
            .call(&storable_proxy, &method, &[UnoValue::String(url), props])
            .await?;

        tracing::info!("Saved workbook to {path}");
        Ok(())
    }

    /// Close the workbook without saving.
    pub async fn close(self) -> Result<()> {
        let closeable_type = Type::interface(type_names::X_CLOSEABLE);
        if let Ok(Some(closeable)) =
            self.conn.query_interface(&self.doc, closeable_type).await
        {
            let method = interface::closeable_close();
            let _ = self
                .conn
                .call(&closeable, &method, &[UnoValue::Bool(true)])
                .await;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(Workbook::parse_cell_ref("A1").unwrap(), (0, 0));
        assert_eq!(Workbook::parse_cell_ref("B2").unwrap(), (1, 1));
        assert_eq!(Workbook::parse_cell_ref("Z1").unwrap(), (25, 0));
        assert_eq!(Workbook::parse_cell_ref("AA1").unwrap(), (26, 0));
        assert_eq!(Workbook::parse_cell_ref("AB3").unwrap(), (27, 2));

        assert!(Workbook::parse_cell_ref("").is_err());
        assert!(Workbook::parse_cell_ref("1").is_err());
        assert!(Workbook::parse_cell_ref("A0").is_err());
        assert!(Workbook::parse_cell_ref("A").is_err());
    }

    #[test]
    fn test_make_property_value() {
        let pv = make_property_value("FilterName", UnoValue::String("test".into()), Type::string());
        match pv {
            UnoValue::Struct(fields) => {
                assert_eq!(fields.len(), 4);
                assert_eq!(fields[0], UnoValue::String("FilterName".to_string()));
                assert_eq!(fields[1], UnoValue::Long(0));
                // fields[2] is Any
                assert_eq!(fields[3], UnoValue::Enum(0));
            }
            _ => panic!("Expected Struct"),
        }
    }
}
