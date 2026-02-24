//! Workbook handle — ergonomic API for working with an Excel workbook via the bridge.

use excel_com_protocol::{CellValue, ChainStep, ResponseData, SheetRef};

use crate::bridge::{BridgeError, ExcelBridge};

/// A handle to an open workbook in the Excel COM bridge.
///
/// Operations on this workbook are forwarded to the bridge server.
/// By default, operations target the first worksheet (index 0).
pub struct Workbook<'a> {
    bridge: &'a ExcelBridge,
    handle: u64,
    /// The active sheet for shorthand methods. Defaults to index 0.
    active_sheet: SheetRef,
}

/// Convert RGB (0x00RRGGBB) to BGR (0x00BBGGRR) for Excel COM.
fn rgb_to_bgr(rgb: u32) -> u32 {
    let r = (rgb >> 16) & 0xFF;
    let g = (rgb >> 8) & 0xFF;
    let b = rgb & 0xFF;
    (b << 16) | (g << 8) | r
}

impl<'a> Workbook<'a> {
    pub(crate) fn new(bridge: &'a ExcelBridge, handle: u64) -> Self {
        Self {
            bridge,
            handle,
            active_sheet: SheetRef::Index(0),
        }
    }

    /// Get the internal handle ID.
    pub fn handle(&self) -> u64 {
        self.handle
    }

    /// Set the active sheet for shorthand methods (by 0-based index).
    pub fn set_active_sheet_index(&mut self, index: u32) {
        self.active_sheet = SheetRef::Index(index);
    }

    /// Set the active sheet for shorthand methods (by name).
    pub fn set_active_sheet_name(&mut self, name: impl Into<String>) {
        self.active_sheet = SheetRef::Name(name.into());
    }

    // -- Shorthand methods that use the active sheet --

    /// Set a cell's value on the active sheet.
    ///
    /// Accepts anything that converts to `CellValue`:
    /// - `&str` / `String` -> String value
    /// - `f64`, `i32`, etc. -> Number value
    /// - `bool` -> Boolean value
    pub fn set_cell_value(
        &self,
        cell: &str,
        value: impl Into<CellValue>,
    ) -> Result<(), BridgeError> {
        self.bridge
            .set_cell_value(self.handle, self.active_sheet.clone(), cell, value.into())
    }

    /// Set a cell's formula on the active sheet (e.g., "=SUM(A1:A10)").
    pub fn set_cell_formula(&self, cell: &str, formula: &str) -> Result<(), BridgeError> {
        self.bridge
            .set_cell_formula(self.handle, self.active_sheet.clone(), cell, formula)
    }

    /// Get a cell's computed value from the active sheet.
    pub fn get_cell_value(&self, cell: &str) -> Result<CellValue, BridgeError> {
        self.bridge
            .get_cell_value(self.handle, self.active_sheet.clone(), cell)
    }

    /// Get a cell's formula from the active sheet (empty string if no formula).
    pub fn get_cell_formula(&self, cell: &str) -> Result<String, BridgeError> {
        self.bridge
            .get_cell_formula(self.handle, self.active_sheet.clone(), cell)
    }

    // -- Sheet-specific methods --

    /// Set a cell value on a specific sheet.
    pub fn set_cell_value_on_sheet(
        &self,
        sheet: SheetRef,
        cell: &str,
        value: impl Into<CellValue>,
    ) -> Result<(), BridgeError> {
        self.bridge
            .set_cell_value(self.handle, sheet, cell, value.into())
    }

    /// Get a cell value from a specific sheet.
    pub fn get_cell_value_on_sheet(
        &self,
        sheet: SheetRef,
        cell: &str,
    ) -> Result<CellValue, BridgeError> {
        self.bridge.get_cell_value(self.handle, sheet, cell)
    }

    // -------------------------------------------------------------------------
    // Font styling
    // -------------------------------------------------------------------------

    /// Set font bold on a cell.
    pub fn set_font_bold(&self, cell: &str, bold: bool) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Bold",
            serde_json::Value::from(bold),
        )
    }

    /// Set font italic on a cell.
    pub fn set_font_italic(&self, cell: &str, italic: bool) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Italic",
            serde_json::Value::from(italic),
        )
    }

    /// Set font size on a cell (in points).
    pub fn set_font_size(&self, cell: &str, size: f64) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Size",
            serde_json::Value::from(size),
        )
    }

    /// Set font name on a cell.
    pub fn set_font_name(&self, cell: &str, name: &str) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Name",
            serde_json::Value::from(name),
        )
    }

    /// Set font color on a cell (RGB as 0xRRGGBB).
    pub fn set_font_color(&self, cell: &str, color: u32) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Color",
            serde_json::Value::from(rgb_to_bgr(color)),
        )
    }

    /// Set font underline on a cell.
    /// `style`: xlUnderlineStyleNone=0, xlUnderlineStyleSingle=2,
    ///          xlUnderlineStyleDouble=-4119
    pub fn set_font_underline(&self, cell: &str, style: i32) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Underline",
            serde_json::Value::from(style),
        )
    }

    /// Set font strikethrough on a cell.
    pub fn set_font_strikethrough(
        &self,
        cell: &str,
        strikethrough: bool,
    ) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Strikethrough",
            serde_json::Value::from(strikethrough),
        )
    }

    /// Set font superscript on a cell.
    pub fn set_font_superscript(&self, cell: &str, superscript: bool) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Superscript",
            serde_json::Value::from(superscript),
        )
    }

    /// Set font subscript on a cell.
    pub fn set_font_subscript(&self, cell: &str, subscript: bool) -> Result<(), BridgeError> {
        self.bridge.set_font_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Subscript",
            serde_json::Value::from(subscript),
        )
    }

    // -------------------------------------------------------------------------
    // Fill styling
    // -------------------------------------------------------------------------

    /// Set fill (background) color on a cell (RGB as 0xRRGGBB).
    pub fn set_fill_color(&self, cell: &str, color: u32) -> Result<(), BridgeError> {
        self.bridge.set_interior_color(
            self.handle,
            self.active_sheet.clone(),
            cell,
            rgb_to_bgr(color),
        )
    }

    // -------------------------------------------------------------------------
    // Border styling
    // -------------------------------------------------------------------------

    /// Set border on all four sides of a cell.
    /// `line_style`: xlContinuous=1, xlDash=-4115, xlDot=-4118, etc.
    /// `weight`: xlThin=2, xlMedium=-4138, xlThick=4
    /// `color`: RGB as 0xRRGGBB
    pub fn set_border_all(
        &self,
        cell: &str,
        line_style: i32,
        weight: i32,
        color: u32,
    ) -> Result<(), BridgeError> {
        // xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10
        for edge in [7, 8, 9, 10] {
            self.bridge.set_border_property(
                self.handle,
                self.active_sheet.clone(),
                cell,
                edge,
                "LineStyle",
                serde_json::Value::from(line_style),
            )?;
            self.bridge.set_border_property(
                self.handle,
                self.active_sheet.clone(),
                cell,
                edge,
                "Weight",
                serde_json::Value::from(weight),
            )?;
            self.bridge.set_border_property(
                self.handle,
                self.active_sheet.clone(),
                cell,
                edge,
                "Color",
                serde_json::Value::from(rgb_to_bgr(color)),
            )?;
        }
        Ok(())
    }

    /// Set border on a specific edge.
    /// `edge`: xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10
    pub fn set_border_edge(
        &self,
        cell: &str,
        edge: i32,
        line_style: i32,
        weight: i32,
        color: u32,
    ) -> Result<(), BridgeError> {
        self.bridge.set_border_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            edge,
            "LineStyle",
            serde_json::Value::from(line_style),
        )?;
        self.bridge.set_border_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            edge,
            "Weight",
            serde_json::Value::from(weight),
        )?;
        self.bridge.set_border_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            edge,
            "Color",
            serde_json::Value::from(rgb_to_bgr(color)),
        )?;
        Ok(())
    }

    // -------------------------------------------------------------------------
    // Alignment styling
    // -------------------------------------------------------------------------

    /// Set horizontal alignment.
    /// `align`: xlLeft=-4131, xlCenter=-4108, xlRight=-4152, xlGeneral=1
    pub fn set_horizontal_alignment(&self, cell: &str, align: i32) -> Result<(), BridgeError> {
        self.bridge.set_range_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "HorizontalAlignment",
            serde_json::Value::from(align),
        )
    }

    /// Set vertical alignment.
    /// `align`: xlTop=-4160, xlCenter=-4108, xlBottom=-4107
    pub fn set_vertical_alignment(&self, cell: &str, align: i32) -> Result<(), BridgeError> {
        self.bridge.set_range_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "VerticalAlignment",
            serde_json::Value::from(align),
        )
    }

    /// Set wrap text on a cell.
    pub fn set_wrap_text(&self, cell: &str, wrap: bool) -> Result<(), BridgeError> {
        self.bridge.set_range_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "WrapText",
            serde_json::Value::from(wrap),
        )
    }

    /// Set shrink-to-fit on a cell.
    pub fn set_shrink_to_fit(&self, cell: &str, shrink: bool) -> Result<(), BridgeError> {
        self.bridge.set_range_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "ShrinkToFit",
            serde_json::Value::from(shrink),
        )
    }

    /// Set text rotation on a cell (degrees, 0-90 or 180 for vertical).
    pub fn set_rotation(&self, cell: &str, degrees: i32) -> Result<(), BridgeError> {
        self.bridge.set_range_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "Orientation",
            serde_json::Value::from(degrees),
        )
    }

    /// Set indent level on a cell.
    pub fn set_indent(&self, cell: &str, level: i32) -> Result<(), BridgeError> {
        self.bridge.set_range_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "IndentLevel",
            serde_json::Value::from(level),
        )
    }

    // -------------------------------------------------------------------------
    // Number format
    // -------------------------------------------------------------------------

    /// Set number format on a cell (e.g., "0.00%", "#,##0.00", "YYYY-MM-DD").
    pub fn set_number_format(&self, cell: &str, format: &str) -> Result<(), BridgeError> {
        self.bridge.set_range_property(
            self.handle,
            self.active_sheet.clone(),
            cell,
            "NumberFormat",
            serde_json::Value::from(format),
        )
    }

    // -------------------------------------------------------------------------
    // Row/Column dimensions
    // -------------------------------------------------------------------------

    /// Set row height (in points). Row is 0-based.
    pub fn set_row_height(&self, row: u32, height: f64) -> Result<(), BridgeError> {
        self.bridge
            .set_row_height(self.handle, self.active_sheet.clone(), row + 1, height)
    }

    /// Set row hidden. Row is 0-based.
    pub fn set_row_hidden(&self, row: u32, hidden: bool) -> Result<(), BridgeError> {
        self.bridge
            .set_row_hidden(self.handle, self.active_sheet.clone(), row + 1, hidden)
    }

    /// Set column width (in character widths). Col is 0-based.
    pub fn set_column_width(&self, col: u32, width: f64) -> Result<(), BridgeError> {
        self.bridge
            .set_column_width(self.handle, self.active_sheet.clone(), col + 1, width)
    }

    // -------------------------------------------------------------------------
    // Merged cells
    // -------------------------------------------------------------------------

    /// Merge a range of cells (e.g., "A1:C3").
    pub fn merge_range(&self, range: &str) -> Result<(), BridgeError> {
        self.bridge
            .merge_range(self.handle, self.active_sheet.clone(), range)
    }

    // -------------------------------------------------------------------------
    // Comments
    // -------------------------------------------------------------------------

    /// Add a comment to a cell.
    pub fn add_comment(&self, cell: &str, text: &str) -> Result<(), BridgeError> {
        self.bridge
            .add_comment(self.handle, self.active_sheet.clone(), cell, text)
    }

    // -------------------------------------------------------------------------
    // Conditional formatting
    // -------------------------------------------------------------------------

    /// Add a conditional format rule and return a handle to style it.
    ///
    /// `cf_type`: xlCellValue=1, xlExpression=2
    /// `operator`: xlBetween=1, xlNotBetween=2, xlEqual=3, xlNotEqual=4,
    ///             xlGreater=5, xlLess=6, xlGreaterEqual=7, xlLessEqual=8
    pub fn add_format_condition(
        &self,
        range: &str,
        cf_type: i32,
        operator: i32,
        formula1: &str,
    ) -> Result<FormatCondition<'_>, BridgeError> {
        let handle = self.bridge.add_format_condition(
            self.handle,
            self.active_sheet.clone(),
            range,
            cf_type,
            operator,
            formula1,
        )?;
        Ok(FormatCondition {
            bridge: self.bridge,
            handle,
        })
    }

    // -------------------------------------------------------------------------
    // Data validation
    // -------------------------------------------------------------------------

    /// Add data validation to a range.
    ///
    /// `vtype`: xlValidateWholeNumber=1, xlValidateDecimal=2, xlValidateList=3,
    ///          xlValidateDate=4, xlValidateTime=5, xlValidateTextLength=6,
    ///          xlValidateCustom=7
    /// `alert_style`: xlValidAlertStop=1, xlValidAlertWarning=2, xlValidAlertInformation=3
    /// `operator`: xlBetween=1, xlGreater=5, xlLess=6, etc. (pass None to skip)
    pub fn add_data_validation(
        &self,
        range: &str,
        vtype: i32,
        alert_style: i32,
        operator: Option<i32>,
        formula1: &str,
        formula2: Option<&str>,
    ) -> Result<(), BridgeError> {
        let mut args: Vec<serde_json::Value> = vec![
            serde_json::Value::from(vtype),
            serde_json::Value::from(alert_style),
        ];
        match operator {
            Some(op) => args.push(serde_json::Value::from(op)),
            None => args.push(serde_json::Value::Null),
        }
        args.push(serde_json::Value::from(formula1));
        if let Some(f2) = formula2 {
            args.push(serde_json::Value::from(f2));
        }
        self.bridge
            .add_validation(self.handle, self.active_sheet.clone(), range, args)
    }

    /// Set the input prompt on a validated range.
    pub fn set_validation_input(
        &self,
        range: &str,
        title: &str,
        message: &str,
    ) -> Result<(), BridgeError> {
        let sheet = self.active_sheet.clone();
        self.bridge.set_validation_property(
            self.handle,
            sheet.clone(),
            range,
            "InputTitle",
            serde_json::Value::from(title),
        )?;
        self.bridge.set_validation_property(
            self.handle,
            sheet,
            range,
            "InputMessage",
            serde_json::Value::from(message),
        )
    }

    /// Set the error alert on a validated range.
    pub fn set_validation_error(
        &self,
        range: &str,
        title: &str,
        message: &str,
    ) -> Result<(), BridgeError> {
        let sheet = self.active_sheet.clone();
        self.bridge.set_validation_property(
            self.handle,
            sheet.clone(),
            range,
            "ErrorTitle",
            serde_json::Value::from(title),
        )?;
        self.bridge.set_validation_property(
            self.handle,
            sheet,
            range,
            "ErrorMessage",
            serde_json::Value::from(message),
        )
    }

    // -------------------------------------------------------------------------
    // File operations
    // -------------------------------------------------------------------------

    /// Save the workbook to a Windows file path.
    ///
    /// The path must be a Windows path visible to the VM. For files shared
    /// via QEMU SMB, use a UNC path like `\\10.0.2.4\qemu\output.xlsx`.
    ///
    /// Format is inferred from extension: `.xlsx` = 51, `.xls` = -4143, `.csv` = 6.
    pub fn save(&self, windows_path: &str) -> Result<(), BridgeError> {
        let format = infer_save_format(windows_path);
        self.bridge.save_workbook(self.handle, windows_path, format)
    }

    /// Save the workbook with an explicit Excel file format constant.
    pub fn save_as(&self, windows_path: &str, format: i32) -> Result<(), BridgeError> {
        self.bridge.save_workbook(self.handle, windows_path, format)
    }

    // -------------------------------------------------------------------------
    // Workbook properties (read)
    // -------------------------------------------------------------------------

    /// Get the workbook's name (e.g., "Book1.xlsx").
    ///
    /// When Excel repairs a file on open, the name may change
    /// (e.g., "file [Repaired].xlsx"), so this can detect repairs.
    pub fn name(&self) -> Result<String, BridgeError> {
        let data = self.bridge.get(self.handle, vec![], "Name")?;
        match data {
            Some(ResponseData::Value { value }) => {
                Ok(value.as_str().unwrap_or_default().to_string())
            }
            _ => Ok(String::new()),
        }
    }

    /// Check if the workbook was opened as read-only.
    ///
    /// Excel sometimes opens repaired files in read-only mode.
    pub fn is_read_only(&self) -> Result<bool, BridgeError> {
        let data = self.bridge.get(self.handle, vec![], "ReadOnly")?;
        match data {
            Some(ResponseData::Value { value }) => Ok(value.as_bool().unwrap_or(false)),
            _ => Ok(false),
        }
    }

    /// Close the workbook without saving.
    pub fn close(self) -> Result<(), BridgeError> {
        self.bridge.close_workbook(self.handle)
    }
}

// =============================================================================
// FormatCondition — handle to a CF rule for styling
// =============================================================================

/// A handle to a FormatCondition COM object, used to style a conditional
/// formatting rule after creation.
pub struct FormatCondition<'a> {
    bridge: &'a ExcelBridge,
    handle: u64,
}

impl FormatCondition<'_> {
    /// Set Font.Bold on this format condition.
    pub fn set_font_bold(&self, bold: bool) -> Result<(), BridgeError> {
        self.bridge.set(
            self.handle,
            vec![ChainStep::Property("Font".to_string())],
            "Bold",
            serde_json::Value::from(bold),
        )
    }

    /// Set Font.Color on this format condition (RGB as 0xRRGGBB).
    pub fn set_font_color(&self, color: u32) -> Result<(), BridgeError> {
        self.bridge.set(
            self.handle,
            vec![ChainStep::Property("Font".to_string())],
            "Color",
            serde_json::Value::from(rgb_to_bgr(color)),
        )
    }

    /// Set Interior.Color (fill) on this format condition (RGB as 0xRRGGBB).
    pub fn set_fill_color(&self, color: u32) -> Result<(), BridgeError> {
        self.bridge.set(
            self.handle,
            vec![ChainStep::Property("Interior".to_string())],
            "Color",
            serde_json::Value::from(rgb_to_bgr(color)),
        )
    }

    /// Set borders on all four edges of this format condition.
    ///
    /// Note: `FormatCondition.Borders` uses `XlBordersIndex` individual border
    /// indices (`xlLeft=1, xlRight=2, xlTop=3, xlBottom=4`), NOT the edge
    /// indices used on `Range.Borders` (`xlEdgeLeft=7..xlEdgeRight=10`).
    pub fn set_border_all(
        &self,
        line_style: i32,
        weight: i32,
        color: u32,
    ) -> Result<(), BridgeError> {
        let bgr = rgb_to_bgr(color);
        // xlLeft=1, xlRight=2, xlTop=3, xlBottom=4
        for edge in [1, 2, 3, 4] {
            let chain = vec![ChainStep::Indexed(
                "Borders".to_string(),
                serde_json::Value::from(edge),
            )];
            self.bridge.set(
                self.handle,
                chain.clone(),
                "LineStyle",
                serde_json::Value::from(line_style),
            )?;
            self.bridge.set(
                self.handle,
                chain.clone(),
                "Weight",
                serde_json::Value::from(weight),
            )?;
            self.bridge
                .set(self.handle, chain, "Color", serde_json::Value::from(bgr))?;
        }
        Ok(())
    }

    // NOTE: FormatCondition COM object does NOT expose HorizontalAlignment
    // or WrapText (DISP_E_UNKNOWNNAME). These DXF properties can only be set
    // through Excel's UI, not COM automation. Tests that need DXF alignment
    // must post-process the xlsx to inject the XML per OOXML spec.

    /// Set NumberFormat on this format condition.
    pub fn set_number_format(&self, format: &str) -> Result<(), BridgeError> {
        self.bridge.set(
            self.handle,
            vec![],
            "NumberFormat",
            serde_json::Value::from(format),
        )
    }
}

/// Infer the Excel file format constant from a file extension.
///
/// - `.xlsx` -> 51 (xlOpenXMLWorkbook)
/// - `.xls`  -> -4143 (xlWorkbookNormal)
/// - `.csv`  -> 6 (xlCSV)
/// - other   -> 51 (default to xlsx)
fn infer_save_format(path: &str) -> i32 {
    let lower = path.to_lowercase();
    if lower.ends_with(".xlsx") {
        51
    } else if lower.ends_with(".xls") {
        -4143
    } else if lower.ends_with(".csv") {
        6
    } else {
        51
    }
}
