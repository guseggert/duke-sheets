//! XLSX writer

use std::collections::HashMap;
use std::fs::File;
use std::io::{Seek, Write};
use std::path::Path;

use crate::error::{XlsxError, XlsxResult};
use crate::styles::XlsxStyleTable;
use duke_sheets_core::{CellAddress, Workbook};

/// Shared string table — maps string content to SST index.
struct SharedStringTable {
    strings: Vec<String>,
    index: HashMap<String, u32>,
}

impl SharedStringTable {
    /// Build the SST by scanning all string cells in the workbook.
    fn build(workbook: &Workbook) -> Self {
        let mut strings = Vec::new();
        let mut index = HashMap::new();

        for sheet in workbook.worksheets() {
            for (_row, _col, cell) in sheet.iter_cells() {
                let s = match &cell.value {
                    duke_sheets_core::CellValue::String(s) => s.as_str(),
                    _ => continue,
                };
                if !index.contains_key(s) {
                    let idx = strings.len() as u32;
                    index.insert(s.to_owned(), idx);
                    strings.push(s.to_owned());
                }
            }
        }

        Self { strings, index }
    }

    /// Look up the SST index for a string.
    fn get(&self, s: &str) -> Option<u32> {
        self.index.get(s).copied()
    }

    fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }
}

/// XLSX file writer
pub struct XlsxWriter;

impl XlsxWriter {
    /// Write a workbook to a file path
    pub fn write_file<P: AsRef<Path>>(workbook: &Workbook, path: P) -> XlsxResult<()> {
        let file = File::create(path)?;
        Self::write(workbook, file)
    }

    /// Write a workbook to a writer
    pub fn write<W: Write + Seek>(workbook: &Workbook, writer: W) -> XlsxResult<()> {
        let mut zip = zip::ZipWriter::new(writer);

        // Build a workbook-wide style table.
        let style_table = XlsxStyleTable::build(workbook);

        // Build shared string table (deduplicated across all sheets).
        let sst = SharedStringTable::build(workbook);

        // Determine which sheets have comments
        let sheets_with_comments: Vec<usize> = workbook
            .worksheets()
            .enumerate()
            .filter(|(_, sheet)| sheet.comment_count() > 0)
            .map(|(i, _)| i)
            .collect();

        // Write [Content_Types].xml
        Self::write_content_types(&mut zip, workbook, &sheets_with_comments, &sst)?;

        // Write _rels/.rels
        Self::write_root_rels(&mut zip)?;

        // Write xl/workbook.xml
        Self::write_workbook_xml(&mut zip, workbook)?;

        // Write xl/_rels/workbook.xml.rels
        Self::write_workbook_rels(&mut zip, workbook, &sst)?;

        // Write xl/styles.xml
        Self::write_styles_xml(&mut zip, &style_table)?;

        // Write shared string table
        if !sst.is_empty() {
            Self::write_shared_strings(&mut zip, &sst)?;
        }

        // Write worksheets and their relationships
        for (i, sheet) in workbook.worksheets().enumerate() {
            Self::write_worksheet(&mut zip, workbook, i, &style_table, &sst)?;

            // Write worksheet relationships if sheet has comments
            if sheet.comment_count() > 0 {
                Self::write_worksheet_rels(&mut zip, i)?;
                Self::write_comments(&mut zip, workbook, i)?;
            }
        }

        zip.finish()?;
        Ok(())
    }

    fn write_content_types<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sheets_with_comments: &[usize],
        sst: &SharedStringTable,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("[Content_Types].xml", options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
        );

        if !sst.is_empty() {
            content.push_str(
                "\n    <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
            );
        }

        // Add an override for each worksheet
        for i in 0..workbook.sheet_count() {
            content.push_str(&format!(
                r#"
    <Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                i + 1
            ));
        }

        // Add content type for comments files
        for &i in sheets_with_comments {
            content.push_str(&format!(
                r#"
    <Override PartName="/xl/comments{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>"#,
                i + 1
            ));
        }

        content.push_str("\n</Types>");

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_root_rels<W: Write + Seek>(zip: &mut zip::ZipWriter<W>) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("_rels/.rels", options)?;

        let content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_workbook_xml<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/workbook.xml", options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        );

        // Write workbookPr (date system, etc.)
        let settings = workbook.settings();
        if settings.date_1904 {
            content.push_str("\n    <workbookPr date1904=\"1\"/>");
        }

        // Write bookViews (active sheet)
        let active = workbook.active_sheet();
        if active > 0 {
            content.push_str(&format!(
                "\n    <bookViews><workbookView activeTab=\"{}\"/></bookViews>",
                active
            ));
        }

        content.push_str("\n    <sheets>");

        for (i, sheet) in workbook.worksheets().enumerate() {
            let state_attr = if !sheet.is_visible() {
                " state=\"hidden\""
            } else {
                ""
            };
            content.push_str(&format!(
                "\n        <sheet name=\"{}\" sheetId=\"{}\"{}  r:id=\"rId{}\"/>",
                Self::escape_xml(sheet.name()),
                i + 1,
                state_attr,
                i + 1
            ));
        }

        content.push_str("\n    </sheets>");

        // Write calcPr (tells Excel to recalculate formulas on open)
        if settings.calc_on_open {
            content.push_str("\n    <calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/>");
        }

        content.push_str("\n</workbook>");

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_workbook_rels<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sst: &SharedStringTable,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/_rels/workbook.xml.rels", options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );

        for i in 0..workbook.sheet_count() {
            content.push_str(&format!(
                "\n    <Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>",
                i + 1,
                i + 1
            ));
        }

        // Styles relationship
        let mut next_rid = workbook.sheet_count() + 1;
        content.push_str(&format!(
            "\n    <Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
            next_rid
        ));
        next_rid += 1;

        // Shared strings relationship
        if !sst.is_empty() {
            content.push_str(&format!(
                "\n    <Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>",
                next_rid
            ));
        }

        content.push_str("\n</Relationships>");

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_shared_strings<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        sst: &SharedStringTable,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/sharedStrings.xml", options)?;

        let count = sst.strings.len();
        let mut content = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{}\" uniqueCount=\"{}\">",
            count, count
        );

        for s in &sst.strings {
            // Use xml:space="preserve" for strings with leading/trailing whitespace
            let needs_preserve = s.starts_with(|c: char| c.is_ascii_whitespace())
                || s.ends_with(|c: char| c.is_ascii_whitespace());
            if needs_preserve {
                content.push_str(&format!(
                    "\n    <si><t xml:space=\"preserve\">{}</t></si>",
                    Self::escape_xml(s)
                ));
            } else {
                content.push_str(&format!("\n    <si><t>{}</t></si>", Self::escape_xml(s)));
            }
        }

        content.push_str("\n</sst>");
        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_styles_xml<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        style_table: &XlsxStyleTable,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/styles.xml", options)?;
        let xml = style_table.to_styles_xml();
        zip.write_all(xml.as_bytes())?;
        Ok(())
    }

    fn write_worksheet<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        index: usize,
        style_table: &XlsxStyleTable,
        sst: &SharedStringTable,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(&format!("xl/worksheets/sheet{}.xml", index + 1), options)?;

        let sheet = workbook
            .worksheet(index)
            .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        );

        // Write sheetPr (tab color)
        if let Some(color) = sheet.tab_color() {
            content.push_str(&format!(
                "\n    <sheetPr><tabColor rgb=\"{}\"/></sheetPr>",
                color.to_argb_hex()
            ));
        }

        // Write sheetViews (freeze panes, tab selection)
        Self::write_sheet_views(&mut content, sheet);

        // Write column definitions (if any custom widths or hidden columns)
        let col_widths = sheet.custom_column_widths();
        let col_hidden = sheet.hidden_columns();
        if !col_widths.is_empty() || !col_hidden.is_empty() {
            content.push_str("\n    <cols>");
            // Collect all columns that need a <col> element
            let mut cols_to_write: std::collections::BTreeSet<u16> = Default::default();
            for &col in col_widths.keys() {
                cols_to_write.insert(col);
            }
            for &col in col_hidden.keys() {
                cols_to_write.insert(col);
            }
            for col in cols_to_write {
                let col1 = col as u32 + 1; // 0-based to 1-based
                let width = col_widths.get(&col).copied().unwrap_or(8.43);
                let hidden = col_hidden.get(&col).copied().unwrap_or(false);
                let mut attrs = format!(
                    " min=\"{}\" max=\"{}\" width=\"{:.2}\" customWidth=\"1\"",
                    col1, col1, width
                );
                if hidden {
                    attrs.push_str(" hidden=\"1\"");
                }
                content.push_str(&format!("\n        <col{}/>", attrs));
            }
            content.push_str("\n    </cols>");
        }

        content.push_str("\n    <sheetData>");

        // Collect metadata-only rows (custom height / hidden, no cells) so
        // they can be interleaved with data rows in ascending order.  OOXML
        // requires <row> elements to appear in strictly ascending r= order.
        let custom_heights = sheet.custom_row_heights();
        let hidden_rows_map = sheet.hidden_rows();
        let mut meta_only_rows: std::collections::BTreeSet<u32> = Default::default();
        for &r in custom_heights.keys() {
            meta_only_rows.insert(r);
        }
        for &r in hidden_rows_map.keys() {
            meta_only_rows.insert(r);
        }

        // Helper: emit a self-closing row element with metadata only.
        let emit_meta_row = |content: &mut String, row: u32| {
            let mut tag = format!("\n        <row r=\"{}\"", row + 1);
            if let Some(&ht) = custom_heights.get(&row) {
                tag.push_str(&format!(" ht=\"{:.2}\" customHeight=\"1\"", ht));
            }
            if hidden_rows_map.get(&row).copied().unwrap_or(false) {
                tag.push_str(" hidden=\"1\"");
            }
            tag.push_str("/>");
            content.push_str(&tag);
        };

        // Peekable iterator over metadata-only rows.
        let mut meta_iter = meta_only_rows.iter().copied().peekable();

        // Write cell data (sparse, row-major), interleaving metadata-only
        // rows before each data row to maintain ascending order.
        let mut current_row: Option<u32> = None;
        let mut written_rows: std::collections::HashSet<u32> = Default::default();
        for (row, col, cell) in sheet.iter_cells() {
            if current_row != Some(row) {
                // Close previous row
                if current_row.is_some() {
                    content.push_str("\n        </row>");
                }

                // Emit any metadata-only rows that come before this data row
                while let Some(&mr) = meta_iter.peek() {
                    if mr >= row {
                        break;
                    }
                    if !written_rows.contains(&mr) {
                        emit_meta_row(&mut content, mr);
                        written_rows.insert(mr);
                    }
                    meta_iter.next();
                }

                // Open new row with optional dimension attributes
                let mut row_tag = format!("\n        <row r=\"{}\"", row + 1);
                if let Some(&ht) = custom_heights.get(&row) {
                    row_tag.push_str(&format!(" ht=\"{:.2}\" customHeight=\"1\"", ht));
                }
                if sheet.is_row_hidden(row) {
                    row_tag.push_str(" hidden=\"1\"");
                }
                row_tag.push('>');
                content.push_str(&row_tag);
                current_row = Some(row);
                written_rows.insert(row);
            }

            let addr = duke_sheets_core::CellAddress::new(row, col);
            let cell_ref = addr.to_a1_string();

            let xf_id = style_table.xf_id_for(index, cell.style_index);
            let style_attr = if xf_id != 0 {
                format!(" s=\"{}\"", xf_id)
            } else {
                String::new()
            };

            match &cell.value {
                duke_sheets_core::CellValue::Number(n) => {
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{}><v>{}</v></c>",
                        cell_ref, style_attr, n
                    ));
                }
                duke_sheets_core::CellValue::String(s) => {
                    if let Some(sst_idx) = sst.get(s.as_str()) {
                        content.push_str(&format!(
                            "\n            <c r=\"{}\"{} t=\"s\"><v>{}</v></c>",
                            cell_ref, style_attr, sst_idx
                        ));
                    } else {
                        // Fallback to inline string (shouldn't happen)
                        content.push_str(&format!(
                            "\n            <c r=\"{}\"{} t=\"inlineStr\"><is><t>{}</t></is></c>",
                            cell_ref,
                            style_attr,
                            Self::escape_xml(s.as_str())
                        ));
                    }
                }
                duke_sheets_core::CellValue::Boolean(b) => {
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{} t=\"b\"><v>{}</v></c>",
                        cell_ref,
                        style_attr,
                        if *b { 1 } else { 0 }
                    ));
                }
                duke_sheets_core::CellValue::Formula {
                    text, cached_value, ..
                } => {
                    let formula_text = if text.starts_with('=') {
                        &text[1..]
                    } else {
                        text.as_str()
                    };
                    // Determine type attribute and <v> element from cached value
                    let (type_attr, value_elem) = match cached_value.as_deref() {
                        Some(duke_sheets_core::CellValue::Number(n)) => {
                            (String::new(), format!("<v>{}</v>", n))
                        }
                        Some(duke_sheets_core::CellValue::String(s)) => (
                            " t=\"str\"".to_string(),
                            format!("<v>{}</v>", Self::escape_xml(s.as_str())),
                        ),
                        Some(duke_sheets_core::CellValue::Boolean(b)) => (
                            " t=\"b\"".to_string(),
                            format!("<v>{}</v>", if *b { 1 } else { 0 }),
                        ),
                        Some(duke_sheets_core::CellValue::Error(e)) => (
                            " t=\"e\"".to_string(),
                            format!("<v>{}</v>", Self::escape_xml(e.as_str())),
                        ),
                        _ => (String::new(), String::new()),
                    };
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{}{}><f>{}</f>{}</c>",
                        cell_ref,
                        style_attr,
                        type_attr,
                        Self::escape_xml(formula_text),
                        value_elem,
                    ));
                }
                duke_sheets_core::CellValue::Error(e) => {
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{} t=\"e\"><v>{}</v></c>",
                        cell_ref,
                        style_attr,
                        Self::escape_xml(e.as_str())
                    ));
                }
                duke_sheets_core::CellValue::Empty => {
                    // Preserve style-only cells
                    if xf_id != 0 {
                        content.push_str(&format!(
                            "\n            <c r=\"{}\"{} />",
                            cell_ref, style_attr
                        ));
                    }
                }
                duke_sheets_core::CellValue::SpillTarget { .. } => {
                    // SpillTarget cells are not written to the file - they are computed
                    // at runtime from the source formula's array result.
                    // In Excel's file format, dynamic array formulas use a special
                    // mechanism, but for simplicity we skip spill targets during write.
                }
            }
        }

        if current_row.is_some() {
            content.push_str("\n        </row>");
        }

        // Emit remaining metadata-only rows that come after all data rows
        for mr in meta_iter {
            if !written_rows.contains(&mr) {
                emit_meta_row(&mut content, mr);
            }
        }

        content.push_str("\n    </sheetData>");

        // Write sheet protection (before mergeCells per OOXML order)
        Self::write_sheet_protection(&mut content, sheet);

        // Write merged cells (if any)
        let merged_regions = sheet.merged_regions();
        if !merged_regions.is_empty() {
            content.push_str(&format!(
                "\n    <mergeCells count=\"{}\">",
                merged_regions.len()
            ));
            for range in merged_regions {
                content.push_str(&format!("\n        <mergeCell ref=\"{}\"/>", range));
            }
            content.push_str("\n    </mergeCells>");
        }

        // Write conditional formatting (if any)
        Self::write_conditional_formatting(&mut content, sheet, index, style_table);

        // Write data validations (if any)
        Self::write_data_validations(&mut content, sheet);

        // Write page margins and page setup
        Self::write_page_setup(&mut content, sheet);

        content.push_str("\n</worksheet>");

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_sheet_views(content: &mut String, sheet: &duke_sheets_core::Worksheet) {
        let freeze = sheet.freeze_panes();
        let selected = sheet.is_selected();

        // Only emit sheetViews if there's something to write
        if freeze.is_none() && !selected {
            return;
        }

        content.push_str("\n    <sheetViews>\n        <sheetView workbookViewId=\"0\"");
        if selected {
            content.push_str(" tabSelected=\"1\"");
        }
        content.push('>');

        if let Some(fp) = freeze {
            // Determine active pane based on what's frozen
            let active_pane = match (fp.col > 0, fp.row > 0) {
                (true, true) => "bottomRight",
                (false, true) => "bottomLeft",
                (true, false) => "topRight",
                (false, false) => "bottomLeft", // shouldn't happen, but safe default
            };

            let top_left = CellAddress::new(fp.row, fp.col).to_a1_string();

            let mut pane_attrs = String::new();
            if fp.col > 0 {
                pane_attrs.push_str(&format!(" xSplit=\"{}\"", fp.col));
            }
            if fp.row > 0 {
                pane_attrs.push_str(&format!(" ySplit=\"{}\"", fp.row));
            }
            content.push_str(&format!(
                "\n            <pane{} topLeftCell=\"{}\" activePane=\"{}\" state=\"frozen\"/>",
                pane_attrs, top_left, active_pane
            ));
            content.push_str(&format!(
                "\n            <selection pane=\"{}\" activeCell=\"{}\" sqref=\"{}\"/>",
                active_pane, top_left, top_left
            ));
        }

        content.push_str("\n        </sheetView>\n    </sheetViews>");
    }

    fn write_sheet_protection(content: &mut String, sheet: &duke_sheets_core::Worksheet) {
        let prot = match sheet.protection() {
            Some(p) if p.protected => p,
            _ => return,
        };

        content.push_str("\n    <sheetProtection sheet=\"1\"");

        if let Some(hash) = prot.password_hash {
            content.push_str(&format!(" password=\"{:04X}\"", hash));
        }

        // ECMA-376 §18.3.1.85 sheetProtection attributes:
        //   - "sheet" = sheet is protected (already emitted above)
        //   - Other attributes: absent or "true"/"1" means NOT allowed.
        //     We emit "0" when our model says the action IS allowed.
        macro_rules! prot_allow {
            ($field:expr, $attr:literal) => {
                if $field {
                    content.push_str(concat!(" ", $attr, "=\"0\""));
                }
            };
        }

        prot_allow!(prot.format_cells, "formatCells");
        prot_allow!(prot.format_columns, "formatColumns");
        prot_allow!(prot.format_rows, "formatRows");
        prot_allow!(prot.insert_columns, "insertColumns");
        prot_allow!(prot.insert_rows, "insertRows");
        prot_allow!(prot.insert_hyperlinks, "insertHyperlinks");
        prot_allow!(prot.delete_columns, "deleteColumns");
        prot_allow!(prot.delete_rows, "deleteRows");
        prot_allow!(prot.sort, "sort");
        prot_allow!(prot.auto_filter, "autoFilter");
        prot_allow!(prot.pivot_tables, "pivotTables");

        // selectLockedCells/selectUnlockedCells: absent = allowed (inverted)
        // So emit "1" when NOT allowed
        if !prot.select_locked_cells {
            content.push_str(" selectLockedCells=\"1\"");
        }
        if !prot.select_unlocked_cells {
            content.push_str(" selectUnlockedCells=\"1\"");
        }

        content.push_str("/>");
    }

    fn write_page_setup(content: &mut String, sheet: &duke_sheets_core::Worksheet) {
        let ps = sheet.page_setup();
        let def = duke_sheets_core::PageSetup::default();

        // Only emit if something differs from defaults
        let margins_differ = (ps.left_margin - def.left_margin).abs() > 1e-9
            || (ps.right_margin - def.right_margin).abs() > 1e-9
            || (ps.top_margin - def.top_margin).abs() > 1e-9
            || (ps.bottom_margin - def.bottom_margin).abs() > 1e-9
            || (ps.header_margin - def.header_margin).abs() > 1e-9
            || (ps.footer_margin - def.footer_margin).abs() > 1e-9;

        let setup_differs = ps.paper_size != def.paper_size
            || ps.orientation != def.orientation
            || ps.scale != def.scale
            || ps.fit_to_width.is_some()
            || ps.fit_to_height.is_some();

        if margins_differ {
            content.push_str(&format!(
                "\n    <pageMargins left=\"{}\" right=\"{}\" top=\"{}\" bottom=\"{}\" header=\"{}\" footer=\"{}\"/>",
                ps.left_margin, ps.right_margin, ps.top_margin, ps.bottom_margin,
                ps.header_margin, ps.footer_margin
            ));
        }

        if setup_differs {
            let orientation = match ps.orientation {
                duke_sheets_core::PageOrientation::Portrait => "portrait",
                duke_sheets_core::PageOrientation::Landscape => "landscape",
            };
            let mut attrs = format!(
                " paperSize=\"{}\" orientation=\"{}\"",
                ps.paper_size, orientation
            );
            if ps.scale != 100 {
                attrs.push_str(&format!(" scale=\"{}\"", ps.scale));
            }
            if let Some(w) = ps.fit_to_width {
                attrs.push_str(&format!(" fitToWidth=\"{}\"", w));
            }
            if let Some(h) = ps.fit_to_height {
                attrs.push_str(&format!(" fitToHeight=\"{}\"", h));
            }
            content.push_str(&format!("\n    <pageSetup{}/>", attrs));
        }
    }

    fn write_conditional_formatting(
        content: &mut String,
        sheet: &duke_sheets_core::Worksheet,
        sheet_index: usize,
        style_table: &XlsxStyleTable,
    ) {
        use duke_sheets_core::conditional_format::CfRuleType;

        let rules = sheet.conditional_formats();
        if rules.is_empty() {
            return;
        }

        // Group rules by their range sets for the <conditionalFormatting> element
        // For simplicity, we output one <conditionalFormatting> per rule
        for (rule_idx, rule) in rules.iter().enumerate() {
            if rule.ranges.is_empty() {
                continue;
            }

            // Build sqref from ranges
            let sqref: String = rule
                .ranges
                .iter()
                .map(|r| r.to_string())
                .collect::<Vec<_>>()
                .join(" ");

            content.push_str(&format!(
                "\n    <conditionalFormatting sqref=\"{}\">",
                sqref
            ));

            // Build the cfRule element
            let rule_type = rule.rule_type.xlsx_type();
            // Get dxf_id from style table (if rule has format) or from rule itself (if loaded from file)
            let dxf_id = style_table
                .dxf_id_for(sheet_index, rule_idx)
                .or(rule.dxf_id);
            let dxf_attr = dxf_id.map_or(String::new(), |id| format!(" dxfId=\"{}\"", id));
            let priority_val = rule.priority.max(1);
            let stop_if_true = if rule.stop_if_true {
                " stopIfTrue=\"1\""
            } else {
                ""
            };

            match &rule.rule_type {
                CfRuleType::CellIs {
                    operator,
                    formula1,
                    formula2,
                } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" operator=\"{}\" priority=\"{}\"{}{}>\n            <formula>{}</formula>",
                        rule_type,
                        operator.xlsx_operator(),
                        priority_val,
                        dxf_attr,
                        stop_if_true,
                        Self::escape_xml(formula1)
                    ));
                    if let Some(f2) = formula2 {
                        content.push_str(&format!(
                            "\n            <formula>{}</formula>",
                            Self::escape_xml(f2)
                        ));
                    }
                    content.push_str("\n        </cfRule>");
                }

                CfRuleType::Expression { formula } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}{}>\n            <formula>{}</formula>\n        </cfRule>",
                        rule_type,
                        priority_val,
                        dxf_attr,
                        stop_if_true,
                        Self::escape_xml(formula)
                    ));
                }

                CfRuleType::ColorScale { colors } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}>\n            <colorScale>",
                        rule_type, priority_val, stop_if_true
                    ));
                    for cv in colors {
                        let val_attr = cv
                            .value
                            .as_ref()
                            .map_or(String::new(), |v| format!(" val=\"{}\"", v));
                        content.push_str(&format!(
                            "\n                <cfvo type=\"{}\"{} />",
                            cv.value_type.xlsx_type(),
                            val_attr
                        ));
                    }
                    for cv in colors {
                        content.push_str(&format!(
                            "\n                <color rgb=\"{}\" />",
                            cv.color.to_argb_hex()
                        ));
                    }
                    content.push_str("\n            </colorScale>\n        </cfRule>");
                }

                CfRuleType::DataBar {
                    min_value,
                    max_value,
                    color,
                    show_value,
                    ..
                } => {
                    let show_val_attr = if *show_value { "" } else { " showValue=\"0\"" };
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}>\n            <dataBar{}>",
                        rule_type, priority_val, stop_if_true, show_val_attr
                    ));

                    // cfvo for min
                    let min_val_attr = min_value
                        .value
                        .as_ref()
                        .map_or(String::new(), |v| format!(" val=\"{}\"", v));
                    content.push_str(&format!(
                        "\n                <cfvo type=\"{}\"{} />",
                        min_value.value_type.xlsx_type(),
                        min_val_attr
                    ));

                    // cfvo for max
                    let max_val_attr = max_value
                        .value
                        .as_ref()
                        .map_or(String::new(), |v| format!(" val=\"{}\"", v));
                    content.push_str(&format!(
                        "\n                <cfvo type=\"{}\"{} />",
                        max_value.value_type.xlsx_type(),
                        max_val_attr
                    ));

                    content.push_str(&format!(
                        "\n                <color rgb=\"{}\" />",
                        color.to_argb_hex()
                    ));
                    content.push_str("\n            </dataBar>\n        </cfRule>");
                }

                CfRuleType::IconSet {
                    icon_style,
                    values,
                    reverse,
                    show_value,
                } => {
                    let reverse_attr = if *reverse { " reverse=\"1\"" } else { "" };
                    let show_val_attr = if *show_value { "" } else { " showValue=\"0\"" };
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}>\n            <iconSet iconSet=\"{}\"{}{}>\n",
                        rule_type, priority_val, stop_if_true, icon_style.xlsx_name(), reverse_attr, show_val_attr
                    ));
                    for val in values {
                        let val_attr = val
                            .value
                            .as_ref()
                            .map_or(String::new(), |v| format!(" val=\"{}\"", v));
                        content.push_str(&format!(
                            "                <cfvo type=\"{}\"{} />\n",
                            val.value_type.xlsx_type(),
                            val_attr
                        ));
                    }
                    content.push_str("            </iconSet>\n        </cfRule>");
                }

                CfRuleType::Top10 {
                    rank,
                    percent,
                    bottom,
                } => {
                    let percent_attr = if *percent { " percent=\"1\"" } else { "" };
                    let bottom_attr = if *bottom { " bottom=\"1\"" } else { "" };
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" rank=\"{}\"{}{}{}{}/>",
                        rule_type,
                        priority_val,
                        rank,
                        percent_attr,
                        bottom_attr,
                        dxf_attr,
                        stop_if_true
                    ));
                }

                CfRuleType::AboveAverage {
                    above,
                    equal_average,
                    std_dev,
                } => {
                    let above_attr = if !*above { " aboveAverage=\"0\"" } else { "" };
                    let equal_attr = if *equal_average {
                        " equalAverage=\"1\""
                    } else {
                        ""
                    };
                    let std_dev_attr =
                        std_dev.map_or(String::new(), |s| format!(" stdDev=\"{}\"", s));
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}{}{}{}{}/>",
                        rule_type,
                        priority_val,
                        above_attr,
                        equal_attr,
                        std_dev_attr,
                        dxf_attr,
                        stop_if_true
                    ));
                }

                CfRuleType::ContainsText { text } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" text=\"{}\"{}{}>\n            <formula>NOT(ISERROR(SEARCH(\"{}\",{})))</formula>\n        </cfRule>",
                        rule_type, priority_val, Self::escape_xml(text), dxf_attr, stop_if_true,
                        Self::escape_xml(text), sqref.split(' ').next().unwrap_or("A1")
                    ));
                }

                CfRuleType::BeginsWith { text } => {
                    let first_cell = sqref
                        .split(' ')
                        .next()
                        .unwrap_or("A1")
                        .split(':')
                        .next()
                        .unwrap_or("A1");
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" text=\"{}\"{}{}>\n            <formula>LEFT({},{})=\"{}\"</formula>\n        </cfRule>",
                        rule_type, priority_val, Self::escape_xml(text), dxf_attr, stop_if_true,
                        first_cell, text.len(), Self::escape_xml(text)
                    ));
                }

                CfRuleType::EndsWith { text } => {
                    let first_cell = sqref
                        .split(' ')
                        .next()
                        .unwrap_or("A1")
                        .split(':')
                        .next()
                        .unwrap_or("A1");
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" text=\"{}\"{}{}>\n            <formula>RIGHT({},{})=\"{}\"</formula>\n        </cfRule>",
                        rule_type, priority_val, Self::escape_xml(text), dxf_attr, stop_if_true,
                        first_cell, text.len(), Self::escape_xml(text)
                    ));
                }

                CfRuleType::DuplicateValues
                | CfRuleType::UniqueValues
                | CfRuleType::ContainsBlanks
                | CfRuleType::NotContainsBlanks
                | CfRuleType::ContainsErrors
                | CfRuleType::NotContainsErrors => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}{}/>",
                        rule_type, priority_val, dxf_attr, stop_if_true
                    ));
                }

                CfRuleType::TimePeriod { period } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" timePeriod=\"{}\"{}{}/>",
                        rule_type,
                        priority_val,
                        period.xlsx_period(),
                        dxf_attr,
                        stop_if_true
                    ));
                }
            }

            content.push_str("\n    </conditionalFormatting>");
        }
    }

    fn write_data_validations(content: &mut String, sheet: &duke_sheets_core::Worksheet) {
        use duke_sheets_core::validation::ValidationType;

        let validations = sheet.data_validations();
        if validations.is_empty() {
            return;
        }

        content.push_str(&format!(
            "\n    <dataValidations count=\"{}\">",
            validations.len()
        ));

        for validation in validations {
            if validation.ranges.is_empty() {
                continue;
            }

            // Build sqref from ranges
            let sqref: String = validation
                .ranges
                .iter()
                .map(|r| r.to_string())
                .collect::<Vec<_>>()
                .join(" ");

            let type_attr = match &validation.validation_type {
                ValidationType::None => String::new(),
                _ => format!(" type=\"{}\"", validation.validation_type.xlsx_type()),
            };

            let operator_attr = match &validation.validation_type {
                ValidationType::Whole { operator, .. }
                | ValidationType::Decimal { operator, .. }
                | ValidationType::Date { operator, .. }
                | ValidationType::Time { operator, .. }
                | ValidationType::TextLength { operator, .. } => {
                    format!(" operator=\"{}\"", operator.xlsx_operator())
                }
                _ => String::new(),
            };

            let allow_blank = if validation.allow_blank {
                " allowBlank=\"1\""
            } else {
                ""
            };
            let show_dropdown = if !validation.show_dropdown {
                " showDropDown=\"1\""
            } else {
                ""
            };
            let show_input = if validation.show_input_message {
                " showInputMessage=\"1\""
            } else {
                ""
            };
            let show_error = if validation.show_error_alert {
                " showErrorMessage=\"1\""
            } else {
                ""
            };

            let error_style = match validation.error_style {
                duke_sheets_core::ValidationErrorStyle::Stop => "",
                duke_sheets_core::ValidationErrorStyle::Warning => " errorStyle=\"warning\"",
                duke_sheets_core::ValidationErrorStyle::Information => {
                    " errorStyle=\"information\""
                }
            };

            let error_title = validation.error_title.as_ref().map_or(String::new(), |t| {
                format!(" errorTitle=\"{}\"", Self::escape_xml(t))
            });
            let error_msg = validation
                .error_message
                .as_ref()
                .map_or(String::new(), |m| {
                    format!(" error=\"{}\"", Self::escape_xml(m))
                });
            let prompt_title = validation.input_title.as_ref().map_or(String::new(), |t| {
                format!(" promptTitle=\"{}\"", Self::escape_xml(t))
            });
            let prompt_msg = validation
                .input_message
                .as_ref()
                .map_or(String::new(), |m| {
                    format!(" prompt=\"{}\"", Self::escape_xml(m))
                });

            content.push_str(&format!(
                "\n        <dataValidation{}{}{}{}{}{}{}{}{}{}{} sqref=\"{}\">",
                type_attr,
                operator_attr,
                allow_blank,
                show_dropdown,
                show_input,
                show_error,
                error_style,
                error_title,
                error_msg,
                prompt_title,
                prompt_msg,
                sqref
            ));

            // Write formulas based on validation type
            match &validation.validation_type {
                ValidationType::List { source } => {
                    // List source: either a range or comma-separated values
                    let formula = if source.starts_with('=') {
                        source[1..].to_string()
                    } else if source.contains('!')
                        || source
                            .chars()
                            .all(|c| c.is_ascii_alphanumeric() || c == '$' || c == ':')
                    {
                        source.clone()
                    } else {
                        // Inline list - wrap in quotes
                        format!("\"{}\"", source)
                    };
                    content.push_str(&format!(
                        "\n            <formula1>{}</formula1>",
                        Self::escape_xml(&formula)
                    ));
                }
                ValidationType::Whole { value1, value2, .. }
                | ValidationType::Decimal { value1, value2, .. }
                | ValidationType::Date { value1, value2, .. }
                | ValidationType::Time { value1, value2, .. }
                | ValidationType::TextLength { value1, value2, .. } => {
                    content.push_str(&format!(
                        "\n            <formula1>{}</formula1>",
                        Self::escape_xml(value1)
                    ));
                    if let Some(v2) = value2 {
                        content.push_str(&format!(
                            "\n            <formula2>{}</formula2>",
                            Self::escape_xml(v2)
                        ));
                    }
                }
                ValidationType::Custom { formula } => {
                    let formula = if formula.starts_with('=') {
                        &formula[1..]
                    } else {
                        formula
                    };
                    content.push_str(&format!(
                        "\n            <formula1>{}</formula1>",
                        Self::escape_xml(formula)
                    ));
                }
                ValidationType::None => {}
            }

            content.push_str("\n        </dataValidation>");
        }

        content.push_str("\n    </dataValidations>");
    }

    fn escape_xml(s: &str) -> String {
        s.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
    }

    /// Write worksheet relationships file (for comments, drawings, etc.)
    fn write_worksheet_rels<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        sheet_index: usize,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(
            &format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_index + 1),
            options,
        )?;

        let content = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments{}.xml"/>
</Relationships>"#,
            sheet_index + 1
        );

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    /// Write comments file for a worksheet
    fn write_comments<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sheet_index: usize,
    ) -> XlsxResult<()> {
        let sheet = workbook
            .worksheet(sheet_index)
            .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

        if sheet.comment_count() == 0 {
            return Ok(());
        }

        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(&format!("xl/comments{}.xml", sheet_index + 1), options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <authors>"#,
        );

        // Write authors
        let authors = sheet.comment_authors();
        for author in authors {
            content.push_str(&format!(
                "\n        <author>{}</author>",
                Self::escape_xml(author)
            ));
        }
        // Add empty author if no authors defined (for comments without author)
        if authors.is_empty() {
            content.push_str("\n        <author></author>");
        }

        content.push_str(
            r#"
    </authors>
    <commentList>"#,
        );

        // Collect and sort comments by cell position for consistent output
        let mut comments: Vec<_> = sheet.comments().collect();
        comments.sort_by_key(|((row, col), _)| (*row, *col));

        // Build author index map
        let author_index: std::collections::HashMap<&str, usize> = authors
            .iter()
            .enumerate()
            .map(|(i, a)| (a.as_str(), i))
            .collect();

        // Write comments
        for ((row, col), comment) in comments {
            let cell_ref = CellAddress::new(row, col).to_a1_string();
            let author_id = if comment.author.is_empty() {
                if authors.is_empty() {
                    0
                } else {
                    0 // Fallback to first author
                }
            } else {
                author_index
                    .get(comment.author.as_str())
                    .copied()
                    .unwrap_or(0)
            };

            content.push_str(&format!(
                r#"
        <comment ref="{}" authorId="{}">
            <text>
                <r>
                    <t>{}</t>
                </r>
            </text>
        </comment>"#,
                cell_ref,
                author_id,
                Self::escape_xml(&comment.text)
            ));
        }

        content.push_str(
            r#"
    </commentList>
</comments>"#,
        );

        zip.write_all(content.as_bytes())?;
        Ok(())
    }
}
