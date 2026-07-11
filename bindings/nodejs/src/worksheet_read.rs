//! Read-only methods for the Worksheet class.

use napi::bindgen_prelude::*;
use napi_derive::napi;

use super::{
    catch_panic, to_napi_err, JsAutoFilter, JsChart, JsChartEx, JsColor, JsComment, JsCommentEntry,
    JsConditionalFormatRule, JsDataValidation, JsEmbeddedImage, JsFormControl, JsFormulaCell,
    JsFreezePanes, JsHyperlink, JsHyperlinkEntry, JsImageInfo, JsMergeSpan, JsMergedRegion,
    JsPageBreak, JsPageSetup, JsProtectedRange, JsRow, JsRowCell, JsRowsOptions, JsSelection,
    JsSheetProtection, JsSpillSource, JsSplitPanes, JsStyle, JsTable, Worksheet,
};

#[napi]
impl Worksheet {
    // Sheet Properties

    /// Sheet visibility: "visible", "hidden", or "veryHidden".
    #[napi(getter)]
    pub fn visibility(&self) -> Result<String> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(match ws.visibility() {
                duke_sheets_core::SheetVisibility::Visible => "visible".to_string(),
                duke_sheets_core::SheetVisibility::Hidden => "hidden".to_string(),
                duke_sheets_core::SheetVisibility::VeryHidden => "veryHidden".to_string(),
            })
        })
    }

    /// Whether the worksheet is selected.
    #[napi(getter)]
    pub fn is_selected(&self) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_selected())
        })
    }

    /// Zoom scale percentage, or null if default.
    #[napi(getter)]
    pub fn zoom_scale(&self) -> Result<Option<u32>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.zoom_scale().map(|z| z as u32))
        })
    }

    /// Sheet tab color, or null if default.
    #[napi(getter)]
    pub fn tab_color(&self) -> Result<Option<JsColor>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.tab_color().map(|c| JsColor::from(&c)))
        })
    }

    /// Whether the worksheet has no cells.
    #[napi(getter, js_name = "isEmpty")]
    pub fn ws_is_empty(&self) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_empty())
        })
    }

    /// Total number of cells with data.
    #[napi(getter)]
    pub fn cell_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.cell_count() as u32)
        })
    }

    /// Sheet view selections.
    #[napi(getter)]
    pub fn selections(&self) -> Result<Vec<JsSelection>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.selections().iter().map(JsSelection::from).collect())
        })
    }

    // Cell Styles (read)

    /// Get the resolved style for a cell by address (e.g., "A1").
    #[napi]
    pub fn get_cell_style(&self, address: String) -> Result<Option<JsStyle>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            let style = ws.cell_style(&address).map_err(to_napi_err)?;
            Ok(style.map(JsStyle::from))
        })
    }

    /// Get the resolved style for a cell by row/col (0-based).
    #[napi]
    pub fn get_cell_style_at(&self, row: u32, col: u32) -> Result<Option<JsStyle>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.cell_style_at(row, col as u16).map(JsStyle::from))
        })
    }

    /// Get the display-formatted value of a cell (e.g., "$1,234.56").
    #[napi]
    pub fn get_formatted_value(&self, address: String) -> Result<String> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            ws.formatted_value(&address).map_err(to_napi_err)
        })
    }

    /// Get the display-formatted value by row/col (0-based).
    #[napi]
    pub fn get_formatted_value_at(&self, row: u32, col: u32) -> Result<String> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.formatted_value_at(row, col as u16))
        })
    }

    /// Get a batch of sparse rows starting from `start_row`.
    ///
    /// Returns up to `max_rows` rows that contain data or metadata.
    /// Each row contains only relevant cells (sparse representation).
    /// Returns an empty array when no more rows exist.
    ///
    /// When metadata flags are enabled (includeStyles, includeMergeInfo, etc.),
    /// cells with that metadata are included even if their value is empty.
    #[napi]
    pub fn get_rows_batch(
        &self,
        start_row: u32,
        max_rows: u32,
        options: Option<JsRowsOptions>,
    ) -> Result<Vec<JsRow>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let opts = options.as_ref();
            let use_formatted = opts.and_then(|o| o.use_formatted_values).unwrap_or(false);
            let use_calculated = opts.and_then(|o| o.use_calculated_values).unwrap_or(false);
            let inc_styles = opts.and_then(|o| o.include_styles).unwrap_or(false);
            let inc_merge = opts.and_then(|o| o.include_merge_info).unwrap_or(false);
            let inc_hyperlinks = opts.and_then(|o| o.include_hyperlinks).unwrap_or(false);
            let inc_comments = opts.and_then(|o| o.include_comments).unwrap_or(false);
            let inc_formulas = opts.and_then(|o| o.include_formulas).unwrap_or(false);
            let inc_images = opts.and_then(|o| o.include_images).unwrap_or(false);
            let skip_empty = opts.and_then(|o| o.skip_empty_values).unwrap_or(false);
            let skip_blank = opts.and_then(|o| o.skip_blank_values).unwrap_or(false);

            // Compute the effective max row, extending for metadata sources.
            let mut max_row = ws.used_range().map(|r| r.end.row).unwrap_or(0);
            if inc_merge {
                for region in ws.merged_regions() {
                    max_row = max_row.max(region.end.row);
                }
            }
            if inc_hyperlinks {
                for (addr, _) in ws.hyperlinks() {
                    max_row = max_row.max(addr.row);
                }
            }
            if inc_comments {
                for ((row, _), _) in ws.comments() {
                    max_row = max_row.max(row);
                }
            }
            if inc_formulas {
                for (row, _, _) in ws.formula_cells() {
                    max_row = max_row.max(row);
                }
            }
            let end_row = max_row.min(start_row.saturating_add(max_rows).saturating_sub(1));

            if start_row > end_row {
                return Ok(vec![]);
            }

            // Build the set of (row, col) coordinates to include.
            // Start with non-empty value cells.
            let mut coords: std::collections::BTreeSet<(u32, u16)> = ws
                .populated_cells_in_range(start_row, end_row)
                .into_iter()
                .collect();

            // Add cells with metadata when their flag is enabled.
            if inc_styles {
                for (&(row, col), data) in ws.cells_map_in_range(start_row, end_row) {
                    if data.style_index != 0 {
                        coords.insert((row, col));
                    }
                }
            }
            if inc_merge {
                for region in ws.merged_regions() {
                    for row in region.start.row.max(start_row)..=region.end.row.min(end_row) {
                        for col in region.start.col..=region.end.col {
                            coords.insert((row, col));
                        }
                    }
                }
            }
            if inc_hyperlinks {
                for (addr, _) in ws.hyperlinks() {
                    if addr.row >= start_row && addr.row <= end_row {
                        coords.insert((addr.row, addr.col));
                    }
                }
            }
            if inc_comments {
                for ((row, col), _) in ws.comments() {
                    if row >= start_row && row <= end_row {
                        coords.insert((row, col));
                    }
                }
            }
            if inc_formulas {
                for (row, col, _) in ws.formula_cells() {
                    if row >= start_row && row <= end_row {
                        coords.insert((row, col));
                    }
                }
            }

            // Group by row and build output.
            let mut rows: Vec<JsRow> = Vec::new();
            let mut current_row: Option<u32> = None;
            let mut current_cells: Vec<JsRowCell> = Vec::new();

            for (row, col) in &coords {
                let (row, col) = (*row, *col);
                if current_row != Some(row) {
                    if let Some(prev_row) = current_row {
                        if !current_cells.is_empty() {
                            rows.push(JsRow {
                                index: prev_row,
                                cells: std::mem::take(&mut current_cells),
                            });
                        }
                    }
                    current_row = Some(row);
                }

                let value = if use_formatted {
                    ws.formatted_value_at(row, col)
                } else if use_calculated {
                    ws.get_calculated_value_at(row, col)
                        .map(|v| v.to_string())
                        .unwrap_or_default()
                } else {
                    ws.get_value_at(row, col).to_string()
                };

                if skip_blank || skip_empty {
                    let raw = ws.get_value_at(row, col);
                    if skip_blank && raw.is_blank() {
                        continue;
                    }
                    if skip_empty && raw.is_empty() {
                        continue;
                    }
                }

                let style = if inc_styles {
                    ws.cell_style_at(row, col).map(JsStyle::from)
                } else {
                    None
                };

                let merge_span = if inc_merge {
                    ws.get_merge_span(row, col).map(|(rs, cs)| JsMergeSpan {
                        row_span: rs,
                        col_span: cs as u32,
                    })
                } else {
                    None
                };

                let is_merged_secondary = if inc_merge {
                    let v = ws.is_merged_secondary(row, col);
                    if v {
                        Some(true)
                    } else {
                        None
                    }
                } else {
                    None
                };

                let hyperlink = if inc_hyperlinks {
                    ws.hyperlink_at(row, col).map(JsHyperlink::from)
                } else {
                    None
                };

                let comment = if inc_comments {
                    ws.comment_at(row, col).map(JsComment::from)
                } else {
                    None
                };

                let formula = if inc_formulas {
                    ws.get_formula_at(row, col).map(|s| s.to_string())
                } else {
                    None
                };

                let image = if inc_images {
                    ws.get_image_at(row, col).map(|info| JsImageInfo {
                        source: info.source,
                        alt_text: info.alt_text,
                        sizing: info.sizing as u32,
                        width: info.width,
                        height: info.height,
                    })
                } else {
                    None
                };

                current_cells.push(JsRowCell {
                    col: col as u32,
                    value,
                    style,
                    merge_span,
                    is_merged_secondary,
                    hyperlink,
                    comment,
                    formula,
                    image,
                });
            }

            if let Some(last_row) = current_row {
                if !current_cells.is_empty() {
                    rows.push(JsRow {
                        index: last_row,
                        cells: current_cells,
                    });
                }
            }

            Ok(rows)
        })
    }

    // Row / Column Properties (read)

    /// Default row height in points.
    #[napi(getter)]
    pub fn default_row_height(&self) -> Result<f64> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.default_row_height())
        })
    }

    /// Default column width in character units.
    #[napi(getter)]
    pub fn default_column_width(&self) -> Result<f64> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.default_column_width())
        })
    }

    /// Whether a row is hidden.
    #[napi]
    pub fn is_row_hidden(&self, row: u32) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_row_hidden(row))
        })
    }

    /// Whether a column is hidden.
    #[napi]
    pub fn is_column_hidden(&self, col: u32) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_column_hidden(col as u16))
        })
    }

    /// Get the outline (grouping) level of a row.
    #[napi]
    pub fn get_row_outline_level(&self, row: u32) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.row_outline_level(row) as u32)
        })
    }

    /// Get the outline (grouping) level of a column.
    #[napi]
    pub fn get_column_outline_level(&self, col: u32) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.column_outline_level(col as u16) as u32)
        })
    }

    /// Whether a row group is collapsed.
    #[napi]
    pub fn is_row_collapsed(&self, row: u32) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_row_collapsed(row))
        })
    }

    /// Whether a column group is collapsed.
    #[napi]
    pub fn is_column_collapsed(&self, col: u32) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_column_collapsed(col as u16))
        })
    }

    // Freeze / Split Panes (read)

    /// Freeze pane settings, or null if no freeze panes.
    #[napi(getter)]
    pub fn freeze_panes(&self) -> Result<Option<JsFreezePanes>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.freeze_panes().map(JsFreezePanes::from))
        })
    }

    /// Split pane settings, or null if no split panes.
    #[napi(getter)]
    pub fn split_panes(&self) -> Result<Option<JsSplitPanes>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.split_panes().map(JsSplitPanes::from))
        })
    }

    // Hyperlinks (read)

    /// Get the hyperlink on a cell, or null if none.
    #[napi]
    pub fn get_hyperlink(&self, address: String) -> Result<Option<JsHyperlink>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.hyperlink(&address).map(JsHyperlink::from))
        })
    }

    /// Get the hyperlink on a cell by row/col (0-based), or null if none.
    #[napi]
    pub fn get_hyperlink_at(&self, row: u32, col: u32) -> Result<Option<JsHyperlink>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.hyperlink_at(row, col as u16).map(JsHyperlink::from))
        })
    }

    /// Number of hyperlinks in the worksheet.
    #[napi(getter)]
    pub fn hyperlink_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.hyperlink_count() as u32)
        })
    }

    /// Get all hyperlinks as an array of `{ address, hyperlink }`.
    #[napi(getter)]
    pub fn hyperlinks(&self) -> Result<Vec<JsHyperlinkEntry>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .hyperlinks()
                .iter()
                .map(|(addr, hl)| JsHyperlinkEntry {
                    address: addr.to_string(),
                    hyperlink: JsHyperlink::from(hl),
                })
                .collect())
        })
    }

    // Comments (read)

    /// Get the comment on a cell by address, or null if none.
    #[napi]
    pub fn get_comment(&self, address: String) -> Result<Option<JsComment>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            let comment = ws.comment(&address).map_err(to_napi_err)?;
            Ok(comment.map(JsComment::from))
        })
    }

    /// Get the comment on a cell by row/col (0-based), or null if none.
    #[napi]
    pub fn get_comment_at(&self, row: u32, col: u32) -> Result<Option<JsComment>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.comment_at(row, col as u16).map(JsComment::from))
        })
    }

    /// Whether a cell has a comment (by address).
    #[napi]
    pub fn has_comment(&self, address: String) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            ws.has_comment(&address).map_err(to_napi_err)
        })
    }

    /// Whether a cell has a comment (by row/col).
    #[napi]
    pub fn has_comment_at(&self, row: u32, col: u32) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.has_comment_at(row, col as u16))
        })
    }

    /// Number of comments in the worksheet.
    #[napi(getter)]
    pub fn comment_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.comment_count() as u32)
        })
    }

    /// Get all comments as an array of `{ row, col, comment }`.
    #[napi(getter)]
    pub fn comments(&self) -> Result<Vec<JsCommentEntry>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .comments()
                .map(|((row, col), c)| JsCommentEntry {
                    row,
                    col: col as u32,
                    comment: JsComment::from(c),
                })
                .collect())
        })
    }

    /// List of distinct comment authors.
    #[napi(getter)]
    pub fn comment_authors(&self) -> Result<Vec<String>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.comment_authors().to_vec())
        })
    }

    // Tables (read)

    /// Get all tables in the worksheet.
    #[napi(getter)]
    pub fn tables(&self) -> Result<Vec<JsTable>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.tables().iter().map(JsTable::from).collect())
        })
    }

    /// Get a table by name, or null if not found.
    #[napi]
    pub fn get_table_by_name(&self, name: String) -> Result<Option<JsTable>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.table_by_name(&name).map(JsTable::from))
        })
    }

    /// Number of tables in the worksheet.
    #[napi(getter)]
    pub fn table_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.table_count() as u32)
        })
    }

    // Data Validation (read)

    /// Get all data validation rules.
    #[napi(getter)]
    pub fn data_validations(&self) -> Result<Vec<JsDataValidation>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .data_validations()
                .iter()
                .map(JsDataValidation::from)
                .collect())
        })
    }

    /// Number of data validation rules.
    #[napi(getter)]
    pub fn data_validation_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.data_validation_count() as u32)
        })
    }

    // Conditional Formatting (read)

    /// Get all conditional formatting rules.
    #[napi(getter)]
    pub fn conditional_formats(&self) -> Result<Vec<JsConditionalFormatRule>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .conditional_formats()
                .iter()
                .map(JsConditionalFormatRule::from)
                .collect())
        })
    }

    /// Number of conditional formatting rules.
    #[napi(getter)]
    pub fn conditional_format_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.conditional_format_count() as u32)
        })
    }

    // Auto-Filter (read)

    /// The auto-filter on this worksheet, or null if none.
    #[napi(getter)]
    pub fn auto_filter(&self) -> Result<Option<JsAutoFilter>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.auto_filter().map(JsAutoFilter::from))
        })
    }

    // Protection (read)

    /// Sheet protection settings, or null if unprotected.
    #[napi(getter)]
    pub fn protection(&self) -> Result<Option<JsSheetProtection>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.protection().map(JsSheetProtection::from))
        })
    }

    /// Protected editable ranges on this worksheet.
    #[napi(getter)]
    pub fn protected_ranges(&self) -> Result<Vec<JsProtectedRange>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .protected_ranges()
                .iter()
                .map(JsProtectedRange::from)
                .collect())
        })
    }

    // Page Setup (read)

    /// Page setup / print settings.
    #[napi(getter)]
    pub fn page_setup(&self) -> Result<JsPageSetup> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(JsPageSetup::from(ws.page_setup()))
        })
    }

    /// Print area range string, or null if not set.
    #[napi(getter)]
    pub fn print_area(&self) -> Result<Option<String>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.print_area().map(|r| r.to_string()))
        })
    }

    /// Repeat rows at top of each printed page as `[startRow, endRow]`, or null.
    #[napi(getter)]
    pub fn repeat_rows(&self) -> Result<Option<Vec<u32>>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.repeat_rows().map(|(s, e)| vec![s, e]))
        })
    }

    /// Repeat columns at left of each printed page as `[startCol, endCol]`, or null.
    #[napi(getter)]
    pub fn repeat_cols(&self) -> Result<Option<Vec<u32>>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.repeat_cols().map(|(s, e)| vec![s as u32, e as u32]))
        })
    }

    /// Manual row page breaks.
    #[napi(getter)]
    pub fn row_breaks(&self) -> Result<Vec<JsPageBreak>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.row_breaks().iter().map(JsPageBreak::from).collect())
        })
    }

    /// Manual column page breaks.
    #[napi(getter)]
    pub fn col_breaks(&self) -> Result<Vec<JsPageBreak>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.col_breaks().iter().map(JsPageBreak::from).collect())
        })
    }

    // Formulas (read)

    /// Get the formula text of a cell by row/col (0-based), or null if not a formula cell.
    #[napi]
    pub fn get_formula_at(&self, row: u32, col: u32) -> Result<Option<String>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.get_formula_at(row, col as u16).map(|s| s.to_string()))
        })
    }

    /// Get the number of formula cells in this worksheet.
    #[napi(getter)]
    pub fn formula_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.formula_cells().count() as u32)
        })
    }

    /// Get all formula cells as an array of `{ row, col, formula }`.
    #[napi(getter)]
    pub fn formula_cells(&self) -> Result<Vec<JsFormulaCell>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .formula_cells()
                .map(|(row, col, formula)| JsFormulaCell {
                    row,
                    col: col as u32,
                    formula: formula.to_string(),
                })
                .collect())
        })
    }

    // Spill (read)

    /// Whether a cell is a spill target (receives a spilled value from a dynamic array formula).
    #[napi]
    pub fn is_spill_target(&self, row: u32, col: u32) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_spill_target(row, col as u16))
        })
    }

    /// Whether a cell is the source of a spill (has a dynamic array formula with results).
    #[napi]
    pub fn is_spill_source(&self, row: u32, col: u32) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_spill_source(row, col as u16))
        })
    }

    /// Get the source cell of a spill target, or null if not a spill target.
    #[napi]
    pub fn get_spill_source(&self, row: u32, col: u32) -> Result<Option<JsSpillSource>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .get_spill_source(row, col as u16)
                .map(|(r, c)| JsSpillSource {
                    row: r,
                    col: c as u32,
                }))
        })
    }

    // Locale / Date System (read)

    /// Whether the worksheet uses the 1904 date system.
    #[napi(getter)]
    pub fn date_1904(&self) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.date_1904())
        })
    }

    // Merged Regions (read - supplements existing mergeCells/unmergeCells)

    /// Get all merged regions as structured objects with start/end row/col.
    #[napi(getter)]
    pub fn merged_regions(&self) -> Result<Vec<JsMergedRegion>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .merged_regions()
                .iter()
                .map(|r| JsMergedRegion {
                    start_row: r.start.row,
                    start_col: r.start.col as u32,
                    end_row: r.end.row,
                    end_col: r.end.col as u32,
                    range: r.to_string(),
                })
                .collect())
        })
    }

    /// Get the merge span for a cell if it is the top-left origin of a merged region.
    ///
    /// Returns `{ rowSpan, colSpan }` if the cell is a merge origin, or null otherwise.
    #[napi]
    pub fn get_merge_span(&self, row: u32, col: u32) -> Result<Option<JsMergeSpan>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws
                .get_merge_span(row, col as u16)
                .map(|(rs, cs)| JsMergeSpan {
                    row_span: rs,
                    col_span: cs as u32,
                }))
        })
    }

    /// Whether a cell is a non-origin member of a merged region (should be skipped when rendering).
    #[napi]
    pub fn is_merged_secondary(&self, row: u32, col: u32) -> Result<bool> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.is_merged_secondary(row, col as u16))
        })
    }
    // Charts (read)

    /// Get all charts embedded in the worksheet.
    #[napi(getter)]
    pub fn charts(&self) -> Result<Vec<JsChart>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.charts().iter().map(JsChart::from).collect())
        })
    }

    /// Number of charts in the worksheet.
    #[napi(getter)]
    pub fn chart_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.chart_count() as u32)
        })
    }

    /// Get all ChartEx charts (Office 2016+ extended charts) in the worksheet.
    #[napi(getter)]
    pub fn charts_ex(&self) -> Result<Vec<JsChartEx>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.charts_ex().iter().map(JsChartEx::from).collect())
        })
    }

    /// Number of ChartEx charts in the worksheet.
    #[napi(getter)]
    pub fn chart_ex_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.chart_ex_count() as u32)
        })
    }

    /// Get all embedded images in the worksheet.
    #[napi(getter)]
    pub fn images(&self) -> Result<Vec<JsEmbeddedImage>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.images().iter().map(JsEmbeddedImage::from).collect())
        })
    }

    /// Number of embedded images in the worksheet.
    #[napi(getter)]
    pub fn image_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.image_count() as u32)
        })
    }

    /// Get all form controls in worksheet order.
    #[napi(getter)]
    pub fn form_controls(&self) -> Result<Vec<JsFormControl>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.form_controls().iter().map(JsFormControl::from).collect())
        })
    }

    /// Number of form controls in the worksheet.
    #[napi(getter)]
    pub fn form_control_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            Ok(ws.form_control_count() as u32)
        })
    }
}
