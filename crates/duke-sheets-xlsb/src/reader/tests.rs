#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use duke_sheets_core::{
        CellRange, CellValue, PivotAggregate, PivotFilter, PivotSource, PivotTable, PivotValue,
        PivotValuesAxis, Workbook,
    };
    use zip::write::SimpleFileOptions;

    use crate::biff12::{build_record, records};
    use crate::reader::XlsbReader;
    use crate::writer::XlsbWriter;

    fn utf16le(s: &str) -> Vec<u8> {
        s.encode_utf16().flat_map(|c| c.to_le_bytes()).collect()
    }

    fn xlwide(s: &str) -> Vec<u8> {
        let encoded = utf16le(s);
        let char_count = (encoded.len() / 2) as u32;
        let mut out = Vec::new();
        out.extend_from_slice(&char_count.to_le_bytes());
        out.extend_from_slice(&encoded);
        out
    }

    fn build_sst(strings: &[&str]) -> Vec<u8> {
        let total = strings.len() as u32;
        let unique = strings.len() as u32;
        let mut sst_payload = Vec::new();
        sst_payload.extend_from_slice(&total.to_le_bytes());
        sst_payload.extend_from_slice(&unique.to_le_bytes());
        let mut data = build_record(records::BRT_BEGIN_SST, &sst_payload);

        for s in strings {
            let mut item_payload = vec![0u8];
            item_payload.extend_from_slice(&xlwide(s));
            data.extend_from_slice(&build_record(records::BRT_SS_ITEM, &item_payload));
        }
        data
    }

    fn build_workbook_bin(sheets: &[(&str, &str)], date_1904: bool) -> Vec<u8> {
        let mut data = Vec::new();

        let mut wb_prop = vec![0u8; 8];
        if date_1904 {
            wb_prop[0] = 0x01;
        }
        data.extend_from_slice(&build_record(records::BRT_WB_PROP, &wb_prop));

        for (name, rel_id) in sheets {
            let mut payload = Vec::new();
            payload.extend_from_slice(&0u32.to_le_bytes());
            payload.extend_from_slice(&1u32.to_le_bytes());
            let rel_encoded = utf16le(rel_id);
            let rel_chars = (rel_encoded.len() / 2) as u32;
            payload.extend_from_slice(&rel_chars.to_le_bytes());
            payload.extend_from_slice(&rel_encoded);
            payload.extend_from_slice(&xlwide(name));
            data.extend_from_slice(&build_record(records::BRT_BUNDLE_SH, &payload));
        }

        data.extend_from_slice(&build_record(records::BRT_END_BUNDLE_SHS, &[]));
        data
    }

    fn build_worksheet_bin(cells: &[(u32, u32, WorksheetCell)]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&build_record(records::BRT_BEGIN_SHEET_DATA, &[]));

        let mut current_row: Option<u32> = None;
        for (row, col, cell) in cells {
            if current_row != Some(*row) {
                let mut row_payload = vec![0u8; 20];
                row_payload[..4].copy_from_slice(&row.to_le_bytes());
                data.extend_from_slice(&build_record(records::BRT_ROW_HDR, &row_payload));
                current_row = Some(*row);
            }

            let mut cell_prefix = vec![0u8; 8];
            cell_prefix[..4].copy_from_slice(&col.to_le_bytes());

            match cell {
                WorksheetCell::Real(v) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&v.to_le_bytes());
                    data.extend_from_slice(&build_record(records::BRT_CELL_REAL, &payload));
                }
                WorksheetCell::Bool(v) => {
                    let mut payload = cell_prefix.clone();
                    payload.push(if *v { 1 } else { 0 });
                    data.extend_from_slice(&build_record(records::BRT_CELL_BOOL, &payload));
                }
                WorksheetCell::InlineStr(s) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&xlwide(s));
                    data.extend_from_slice(&build_record(records::BRT_CELL_ST, &payload));
                }
                WorksheetCell::SharedString(idx) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&(*idx as u32).to_le_bytes());
                    data.extend_from_slice(&build_record(records::BRT_CELL_ISST, &payload));
                }
                WorksheetCell::Error(code) => {
                    let mut payload = cell_prefix.clone();
                    payload.push(*code);
                    data.extend_from_slice(&build_record(records::BRT_CELL_ERROR, &payload));
                }
                WorksheetCell::Rk(rk_bits) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&rk_bits.to_le_bytes());
                    data.extend_from_slice(&build_record(records::BRT_CELL_RK, &payload));
                }
                WorksheetCell::FmlaNum(v, tokens) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&v.to_le_bytes()); // f64 cached value
                    payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
                    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes()); // cce
                    payload.extend_from_slice(tokens);
                    data.extend_from_slice(&build_record(records::BRT_FMLA_NUM, &payload));
                }
                WorksheetCell::FmlaBool(v, tokens) => {
                    let mut payload = cell_prefix.clone();
                    payload.push(if *v { 1 } else { 0 }); // bool cached value
                    payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
                    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes()); // cce
                    payload.extend_from_slice(tokens);
                    data.extend_from_slice(&build_record(records::BRT_FMLA_BOOL, &payload));
                }
                WorksheetCell::FmlaError(code, tokens) => {
                    let mut payload = cell_prefix.clone();
                    payload.push(*code); // error cached value
                    payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
                    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes()); // cce
                    payload.extend_from_slice(tokens);
                    data.extend_from_slice(&build_record(records::BRT_FMLA_ERROR, &payload));
                }
                WorksheetCell::FmlaString(s, tokens) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&xlwide(s)); // XLWideString cached value
                    payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
                    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes()); // cce
                    payload.extend_from_slice(tokens);
                    data.extend_from_slice(&build_record(records::BRT_FMLA_STRING, &payload));
                }
            }
        }

        data.extend_from_slice(&build_record(records::BRT_END_SHEET_DATA, &[]));
        data
    }

    enum WorksheetCell {
        Real(f64),
        Bool(bool),
        InlineStr(&'static str),
        SharedString(usize),
        Error(u8),
        Rk(u32),
        FmlaNum(f64, Vec<u8>),
        FmlaBool(bool, Vec<u8>),
        FmlaError(u8, Vec<u8>),
        FmlaString(&'static str, Vec<u8>),
    }

    fn rels_xml(entries: &[(&str, &str)]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for (id, target) in entries {
            xml.push_str(&format!(
                r#"<Relationship Id="{id}" Target="{target}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>"#,
            ));
        }
        xml.push_str("</Relationships>");
        xml
    }

    fn build_xlsb_zip(
        workbook_bin: &[u8],
        rels_xml: &str,
        sst: Option<&[u8]>,
        worksheets: &[(&str, &[u8])],
    ) -> Vec<u8> {
        build_xlsb_zip_with_sheet_rels(workbook_bin, rels_xml, sst, worksheets, &[])
    }

    fn build_xlsb_zip_with_sheet_rels(
        workbook_bin: &[u8],
        rels_xml: &str,
        sst: Option<&[u8]>,
        worksheets: &[(&str, &[u8])],
        sheet_rels: &[(&str, &str)],
    ) -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = SimpleFileOptions::default();
            zip.start_file("xl/workbook.bin", opts).unwrap();
            zip.write_all(workbook_bin).unwrap();
            zip.start_file("xl/_rels/workbook.bin.rels", opts).unwrap();
            zip.write_all(rels_xml.as_bytes()).unwrap();
            if let Some(sst_data) = sst {
                zip.start_file("xl/sharedStrings.bin", opts).unwrap();
                zip.write_all(sst_data).unwrap();
            }
            for (path, data) in worksheets {
                zip.start_file(*path, opts).unwrap();
                zip.write_all(data).unwrap();
            }
            for (path, data) in sheet_rels {
                zip.start_file(*path, opts).unwrap();
                zip.write_all(data.as_bytes()).unwrap();
            }
            zip.finish().unwrap();
        }
        buf
    }

    fn write_xlsb(workbook: &Workbook) -> Vec<u8> {
        let mut buf = Vec::new();
        XlsbWriter::write(workbook, Cursor::new(&mut buf)).unwrap();
        buf
    }

    #[test]
    fn read_single_sheet_with_numbers() {
        let ws_data = build_worksheet_bin(&[
            (0, 0, WorksheetCell::Real(1.0)),
            (0, 1, WorksheetCell::Real(2.5)),
            (1, 0, WorksheetCell::Real(3.0)),
        ]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.worksheet(0).unwrap().name(), "Sheet1");

        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Number(1.0));
        assert_eq!(ws.get_value_at(0, 1), CellValue::Number(2.5));
        assert_eq!(ws.get_value_at(1, 0), CellValue::Number(3.0));
        assert_eq!(ws.get_value_at(1, 1), CellValue::Empty);
    }

    #[test]
    fn reads_writer_pivot_table_semantics() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Segment").unwrap();
        ws.set_cell_value("B1", "Region").unwrap();
        ws.set_cell_value("C1", "Quarter").unwrap();
        ws.set_cell_value("D1", "Revenue").unwrap();
        ws.set_cell_value("E1", "Units").unwrap();
        ws.set_cell_value("A2", "Online").unwrap();
        ws.set_cell_value("B2", "East").unwrap();
        ws.set_cell_value("C2", "Q1").unwrap();
        ws.set_cell_value("D2", 10.0).unwrap();
        ws.set_cell_value("E2", 2.0).unwrap();
        ws.set_cell_value("A3", "Retail").unwrap();
        ws.set_cell_value("B3", "East").unwrap();
        ws.set_cell_value("C3", "Q2").unwrap();
        ws.set_cell_value("D3", 20.0).unwrap();
        ws.set_cell_value("E3", 4.0).unwrap();
        ws.set_cell_value("A4", "Online").unwrap();
        ws.set_cell_value("B4", "West").unwrap();
        ws.set_cell_value("C4", "Q1").unwrap();
        ws.set_cell_value("D4", 30.0).unwrap();
        ws.set_cell_value("E4", 6.0).unwrap();

        let mut pivot = PivotTable::builder("RevenuePivot")
            .source_range(CellRange::parse("A1:E4").unwrap())
            .target_address("G2")
            .unwrap()
            .page("Segment")
            .row("Region")
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
            .named_measure("Units", PivotAggregate::Average, "Average Units")
            .filter(PivotFilter::field_items("Segment", ["Online"]))
            .build()
            .unwrap();
        pivot.layout.values_axis = PivotValuesAxis::Columns;
        pivot.layout.values_axis_position = Some(1);
        wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();

        let bytes = write_xlsb(&wb);
        let read = XlsbReader::read(Cursor::new(bytes)).unwrap();
        let ws = read.worksheet(0).unwrap();
        assert_eq!(ws.pivot_tables().len(), 1);

        let pivot = &ws.pivot_tables()[0];
        assert_eq!(pivot.name, "RevenuePivot");
        assert_eq!(pivot.target.to_a1_string(), "G2");
        assert_eq!(
            pivot.source,
            PivotSource::range_on_sheet("Sheet1", CellRange::parse("A1:E4").unwrap())
        );
        assert_eq!(pivot.rows[0].field.name, "Region");
        assert_eq!(pivot.columns[0].field.name, "Quarter");
        assert_eq!(pivot.page_fields[0].field.name, "Segment");
        assert_eq!(pivot.style.name.as_deref(), Some("PivotStyleMedium9"));
        assert_eq!(pivot.layout.values_axis, PivotValuesAxis::Columns);
        assert_eq!(pivot.layout.values_axis_position, Some(1));
        assert_eq!(pivot.measures.len(), 2);
        assert_eq!(pivot.measures[0].field.name, "Revenue");
        assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
        assert_eq!(pivot.measures[0].name.as_deref(), Some("Total Revenue"));
        assert_eq!(pivot.measures[1].field.name, "Units");
        assert_eq!(pivot.measures[1].aggregate, PivotAggregate::Average);
        assert_eq!(pivot.measures[1].name.as_deref(), Some("Average Units"));
        assert!(pivot.filters.iter().any(|filter| matches!(
            filter,
            PivotFilter::FieldItems {
                field,
                allowed_items,
            } if field.name == "Segment"
                && allowed_items == &vec![PivotValue::String("Online".to_string())]
        )));
        let cache_info = pivot.cache_info().expect("cache diagnostics");
        assert_eq!(
            cache_info.source_kind,
            duke_sheets_core::PivotCacheSourceKind::Worksheet
        );
        assert_eq!(cache_info.record_count, Some(3));
    }

    #[test]
    fn reads_shared_pivot_cache_once_across_sheets() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value("A1", "Region").unwrap();
        ws.set_cell_value("B1", "Revenue").unwrap();
        ws.set_cell_value("A2", "East").unwrap();
        ws.set_cell_value("B2", 10.0).unwrap();
        ws.set_cell_value("A3", "West").unwrap();
        ws.set_cell_value("B3", 20.0).unwrap();
        let second_sheet = wb.add_worksheet().unwrap();

        let first = PivotTable::builder("SalesPivotA")
            .source_range_on_sheet("Sheet1", CellRange::parse("A1:B3").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap();
        let second = PivotTable::builder("SalesPivotB")
            .source_range_on_sheet("Sheet1", CellRange::parse("A1:B3").unwrap())
            .target_address("A1")
            .unwrap()
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_pivot_table(first).unwrap();
        wb.worksheet_mut(second_sheet)
            .unwrap()
            .add_pivot_table(second)
            .unwrap();

        let bytes = write_xlsb(&wb);
        let (read, definition_parses, records_parses) =
            XlsbReader::read_with_pivot_cache_parse_counts(Cursor::new(bytes)).unwrap();

        let first_pivots = read.worksheet(0).unwrap().pivot_tables();
        let second_pivots = read.worksheet(1).unwrap().pivot_tables();
        assert_eq!(first_pivots.len(), 1);
        assert_eq!(second_pivots.len(), 1);
        assert_eq!(first_pivots[0].name, "SalesPivotA");
        assert_eq!(second_pivots[0].name, "SalesPivotB");
        assert_eq!(definition_parses, 1);
        assert_eq!(records_parses, 1);
    }

    #[test]
    fn read_shared_strings() {
        let sst = build_sst(&["Hello", "World"]);
        let ws_data = build_worksheet_bin(&[
            (0, 0, WorksheetCell::SharedString(0)),
            (0, 1, WorksheetCell::SharedString(1)),
        ]);
        let wb_data = build_workbook_bin(&[("Data", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            Some(&sst),
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.name(), "Data");
        assert_eq!(ws.get_value_at(0, 0), CellValue::String("Hello".into()));
        assert_eq!(ws.get_value_at(0, 1), CellValue::String("World".into()));
    }

    #[test]
    fn read_inline_strings() {
        let ws_data = build_worksheet_bin(&[
            (0, 0, WorksheetCell::InlineStr("inline")),
            (1, 0, WorksheetCell::InlineStr("text")),
        ]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::String("inline".into()));
        assert_eq!(ws.get_value_at(1, 0), CellValue::String("text".into()));
    }

    #[test]
    fn read_booleans() {
        let ws_data = build_worksheet_bin(&[
            (0, 0, WorksheetCell::Bool(true)),
            (0, 1, WorksheetCell::Bool(false)),
        ]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Boolean(true));
        assert_eq!(ws.get_value_at(0, 1), CellValue::Boolean(false));
    }

    #[test]
    fn read_errors() {
        let ws_data = build_worksheet_bin(&[
            (0, 0, WorksheetCell::Error(0x07)),
            (0, 1, WorksheetCell::Error(0x2A)),
        ]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(
            ws.get_value_at(0, 0),
            CellValue::Error(duke_sheets_core::CellError::Div0)
        );
        assert_eq!(
            ws.get_value_at(0, 1),
            CellValue::Error(duke_sheets_core::CellError::Na)
        );
    }

    #[test]
    fn read_rk_encoded_values() {
        let rk_42 = ((42i32 << 2) as u32) | 0x02;
        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::Rk(rk_42))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Number(42.0));
    }

    #[test]
    fn read_multiple_sheets() {
        let ws1 = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(10.0))]);
        let ws2 = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(20.0))]);
        let wb_data = build_workbook_bin(&[("Alpha", "rId1"), ("Beta", "rId2")], false);
        let rels = rels_xml(&[
            ("rId1", "worksheets/sheet1.bin"),
            ("rId2", "worksheets/sheet2.bin"),
        ]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[
                ("xl/worksheets/sheet1.bin", &ws1),
                ("xl/worksheets/sheet2.bin", &ws2),
            ],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        assert_eq!(wb.sheet_count(), 2);
        assert_eq!(wb.worksheet(0).unwrap().name(), "Alpha");
        assert_eq!(wb.worksheet(1).unwrap().name(), "Beta");
        assert_eq!(
            wb.worksheet(0).unwrap().get_value_at(0, 0),
            CellValue::Number(10.0)
        );
        assert_eq!(
            wb.worksheet(1).unwrap().get_value_at(0, 0),
            CellValue::Number(20.0)
        );
    }

    #[test]
    fn read_date_1904() {
        let ws_data = build_worksheet_bin(&[]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], true);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        assert!(wb.settings().date_1904);
    }

    #[test]
    fn read_mixed_cell_types() {
        let sst = build_sst(&["shared"]);
        let ws_data = build_worksheet_bin(&[
            (0, 0, WorksheetCell::Real(3.14)),
            (0, 1, WorksheetCell::Bool(true)),
            (0, 2, WorksheetCell::InlineStr("hello")),
            (0, 3, WorksheetCell::SharedString(0)),
            (0, 4, WorksheetCell::Error(0x0F)),
        ]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            Some(&sst),
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Number(3.14));
        assert_eq!(ws.get_value_at(0, 1), CellValue::Boolean(true));
        assert_eq!(ws.get_value_at(0, 2), CellValue::String("hello".into()));
        assert_eq!(ws.get_value_at(0, 3), CellValue::String("shared".into()));
        assert_eq!(
            ws.get_value_at(0, 4),
            CellValue::Error(duke_sheets_core::CellError::Value)
        );
    }

    #[test]
    fn read_no_shared_strings_file() {
        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(99.0))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(
            wb.worksheet(0).unwrap().get_value_at(0, 0),
            CellValue::Number(99.0)
        );
    }

    fn make_biff12_ref(row: u32, col: u16, row_rel: bool, col_rel: bool) -> Vec<u8> {
        let mut out = row.to_le_bytes().to_vec();
        let mut col_word: u16 = col;
        if row_rel {
            col_word |= 0x4000;
        }
        if col_rel {
            col_word |= 0x8000;
        }
        out.extend_from_slice(&col_word.to_le_bytes());
        out
    }

    fn biff12_tstr_token(s: &str) -> Vec<u8> {
        let encoded = utf16le(s);
        let char_count = (encoded.len() / 2) as u16;
        let mut tokens = vec![0x17u8];
        tokens.extend_from_slice(&char_count.to_le_bytes());
        tokens.extend_from_slice(&encoded);
        tokens
    }

    fn biff12_tokens_add_1_2() -> Vec<u8> {
        // tInt(1) tInt(2) tAdd => "1+2"
        vec![0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03]
    }

    fn biff12_tokens_sum_a1_a10() -> Vec<u8> {
        let mut t = vec![0x45]; // tAreaV = 0x25 + 0x20
        t.extend_from_slice(&0u32.to_le_bytes()); // first_row=0
        t.extend_from_slice(&9u32.to_le_bytes()); // last_row=9
        let fc: u16 = 0x0000 | 0x4000 | 0x8000; // col=0, both relative
        t.extend_from_slice(&fc.to_le_bytes());
        t.extend_from_slice(&fc.to_le_bytes());
        t.push(0x19);
        t.push(0x10);
        t.extend_from_slice(&0u16.to_le_bytes());
        t
    }

    fn biff12_tokens_ref_b1(row_rel: bool, col_rel: bool) -> Vec<u8> {
        // tRefV(B1) => "B1" or "$B$1"
        let mut t = vec![0x44]; // tRefV = 0x24 + 0x20
        t.extend_from_slice(&make_biff12_ref(0, 1, row_rel, col_rel));
        t
    }

    #[test]
    fn read_formula_num() {
        let tokens = biff12_tokens_add_1_2();
        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::FmlaNum(3.0, tokens))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Number(3.0));
        assert_eq!(ws.get_formula_at(0, 0), Some("=1+2"));
    }

    #[test]
    fn read_formula_bool() {
        let tokens = biff12_tokens_ref_b1(true, true);
        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::FmlaBool(true, tokens))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Boolean(true));
        assert_eq!(ws.get_formula_at(0, 0), Some("=B1"));
    }

    #[test]
    fn read_formula_error() {
        let tokens = biff12_tokens_ref_b1(false, false);
        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::FmlaError(0x07, tokens))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(
            ws.get_value_at(0, 0),
            CellValue::Error(duke_sheets_core::CellError::Div0)
        );
        assert_eq!(ws.get_formula_at(0, 0), Some("=$B$1"));
    }

    #[test]
    fn read_formula_string() {
        let tokens = biff12_tokens_ref_b1(true, true);
        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::FmlaString("hello", tokens))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::String("hello".into()));
        assert_eq!(ws.get_formula_at(0, 0), Some("=B1"));
    }

    #[test]
    fn read_formula_sum() {
        let tokens = biff12_tokens_sum_a1_a10();
        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::FmlaNum(55.0, tokens))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value_at(0, 0), CellValue::Number(55.0));
        assert_eq!(ws.get_formula_at(0, 0), Some("=SUM(A1:A10)"));
    }

    fn build_merge_cell(first_row: u32, last_row: u32, first_col: u32, last_col: u32) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&first_row.to_le_bytes());
        payload.extend_from_slice(&last_row.to_le_bytes());
        payload.extend_from_slice(&first_col.to_le_bytes());
        payload.extend_from_slice(&last_col.to_le_bytes());
        build_record(records::BRT_MERGE_CELL, &payload)
    }

    fn build_col_info(
        first_col: u32,
        last_col: u32,
        width_256ths: u32,
        hidden: bool,
        custom_width: bool,
    ) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&first_col.to_le_bytes());
        payload.extend_from_slice(&last_col.to_le_bytes());
        payload.extend_from_slice(&width_256ths.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        let mut flags: u16 = 0;
        if hidden {
            flags |= 0x01;
        }
        if custom_width {
            flags |= 0x02;
        }
        payload.extend_from_slice(&flags.to_le_bytes());
        build_record(records::BRT_COL_INFO, &payload)
    }

    fn build_row_hdr(row: u32, height_twips: u16, custom_height: bool, hidden: bool) -> Vec<u8> {
        let mut payload = vec![0u8; 17];
        payload[..4].copy_from_slice(&row.to_le_bytes());
        payload[8..10].copy_from_slice(&height_twips.to_le_bytes());
        let mut flags: u8 = 0;
        if hidden {
            flags |= 0x10;
        }
        if custom_height {
            flags |= 0x20;
        }
        payload[11] = flags;
        build_record(records::BRT_ROW_HDR, &payload)
    }

    fn build_pane(x_split: f64, y_split: f64, state: u8) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&x_split.to_le_bytes());
        payload.extend_from_slice(&y_split.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&3u32.to_le_bytes());
        payload.push(state);
        build_record(records::BRT_PANE, &payload)
    }

    fn build_auto_filter(first_row: u32, last_row: u32, first_col: u32, last_col: u32) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&first_row.to_le_bytes());
        payload.extend_from_slice(&last_row.to_le_bytes());
        payload.extend_from_slice(&first_col.to_le_bytes());
        payload.extend_from_slice(&last_col.to_le_bytes());
        build_record(records::BRT_BEGIN_A_FILTER, &payload)
    }

    fn build_margins(
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&left.to_le_bytes());
        payload.extend_from_slice(&right.to_le_bytes());
        payload.extend_from_slice(&top.to_le_bytes());
        payload.extend_from_slice(&bottom.to_le_bytes());
        payload.extend_from_slice(&header.to_le_bytes());
        payload.extend_from_slice(&footer.to_le_bytes());
        build_record(records::BRT_MARGINS, &payload)
    }

    fn build_page_setup(paper_size: u32, scale: u32, landscape: bool) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&paper_size.to_le_bytes());
        payload.extend_from_slice(&scale.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        let grbit: u16 = if landscape { 0x02 } else { 0x00 };
        payload.extend_from_slice(&grbit.to_le_bytes());
        build_record(records::BRT_PAGE_SETUP, &payload)
    }

    fn build_header_footer(odd_header: &str, odd_footer: &str) -> Vec<u8> {
        let mut payload = Vec::new();
        let flags: u16 = 0x0C;
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&xlwide(odd_header));
        payload.extend_from_slice(&xlwide(odd_footer));
        payload.extend_from_slice(&xlwide(""));
        payload.extend_from_slice(&xlwide(""));
        payload.extend_from_slice(&xlwide(""));
        payload.extend_from_slice(&xlwide(""));
        build_record(records::BRT_HEADER_FOOTER, &payload)
    }

    fn build_hlink(
        row: u32,
        col: u32,
        rel_id: &str,
        location: &str,
        tooltip: &str,
        display: &str,
    ) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&col.to_le_bytes());
        payload.extend_from_slice(&col.to_le_bytes());
        payload.extend_from_slice(&xlwide(rel_id));
        payload.extend_from_slice(&xlwide(location));
        payload.extend_from_slice(&xlwide(tooltip));
        payload.extend_from_slice(&xlwide(display));
        build_record(records::BRT_H_LINK, &payload)
    }

    fn build_worksheet_bin_full(
        pre_data: &[u8],
        cells: &[(u32, u32, WorksheetCell)],
        post_data: &[u8],
    ) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(pre_data);
        data.extend_from_slice(&build_record(records::BRT_BEGIN_SHEET_DATA, &[]));

        let mut current_row: Option<u32> = None;
        for (row, col, cell) in cells {
            if current_row != Some(*row) {
                let mut row_payload = vec![0u8; 20];
                row_payload[..4].copy_from_slice(&row.to_le_bytes());
                data.extend_from_slice(&build_record(records::BRT_ROW_HDR, &row_payload));
                current_row = Some(*row);
            }

            let mut cell_prefix = vec![0u8; 8];
            cell_prefix[..4].copy_from_slice(&col.to_le_bytes());

            match cell {
                WorksheetCell::Real(v) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&v.to_le_bytes());
                    data.extend_from_slice(&build_record(records::BRT_CELL_REAL, &payload));
                }
                WorksheetCell::Bool(v) => {
                    let mut payload = cell_prefix.clone();
                    payload.push(if *v { 1 } else { 0 });
                    data.extend_from_slice(&build_record(records::BRT_CELL_BOOL, &payload));
                }
                WorksheetCell::InlineStr(s) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&xlwide(s));
                    data.extend_from_slice(&build_record(records::BRT_CELL_ST, &payload));
                }
                WorksheetCell::SharedString(idx) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&(*idx as u32).to_le_bytes());
                    data.extend_from_slice(&build_record(records::BRT_CELL_ISST, &payload));
                }
                WorksheetCell::Error(code) => {
                    let mut payload = cell_prefix.clone();
                    payload.push(*code);
                    data.extend_from_slice(&build_record(records::BRT_CELL_ERROR, &payload));
                }
                WorksheetCell::Rk(rk_bits) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&rk_bits.to_le_bytes());
                    data.extend_from_slice(&build_record(records::BRT_CELL_RK, &payload));
                }
                WorksheetCell::FmlaNum(v, tokens) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&v.to_le_bytes());
                    payload.extend_from_slice(&0u16.to_le_bytes());
                    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes());
                    payload.extend_from_slice(tokens);
                    data.extend_from_slice(&build_record(records::BRT_FMLA_NUM, &payload));
                }
                WorksheetCell::FmlaBool(v, tokens) => {
                    let mut payload = cell_prefix.clone();
                    payload.push(if *v { 1 } else { 0 });
                    payload.extend_from_slice(&0u16.to_le_bytes());
                    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes());
                    payload.extend_from_slice(tokens);
                    data.extend_from_slice(&build_record(records::BRT_FMLA_BOOL, &payload));
                }
                WorksheetCell::FmlaError(code, tokens) => {
                    let mut payload = cell_prefix.clone();
                    payload.push(*code);
                    payload.extend_from_slice(&0u16.to_le_bytes());
                    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes());
                    payload.extend_from_slice(tokens);
                    data.extend_from_slice(&build_record(records::BRT_FMLA_ERROR, &payload));
                }
                WorksheetCell::FmlaString(s, tokens) => {
                    let mut payload = cell_prefix.clone();
                    payload.extend_from_slice(&xlwide(s));
                    payload.extend_from_slice(&0u16.to_le_bytes());
                    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes());
                    payload.extend_from_slice(tokens);
                    data.extend_from_slice(&build_record(records::BRT_FMLA_STRING, &payload));
                }
            }
        }

        data.extend_from_slice(&build_record(records::BRT_END_SHEET_DATA, &[]));
        data.extend_from_slice(post_data);
        data
    }

    fn build_worksheet_with_row_hdrs(rows: &[(u32, u16, bool)]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&build_record(records::BRT_BEGIN_SHEET_DATA, &[]));
        for &(row, height_twips, hidden) in rows {
            data.extend_from_slice(&build_row_hdr(row, height_twips, height_twips > 0, hidden));
        }
        data.extend_from_slice(&build_record(records::BRT_END_SHEET_DATA, &[]));
        data
    }

    fn sheet_rels_xml(entries: &[(&str, &str, &str)]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for (id, target, rel_type) in entries {
            xml.push_str(&format!(
                r#"<Relationship Id="{id}" Target="{target}" Type="{rel_type}" TargetMode="External"/>"#,
            ));
        }
        xml.push_str("</Relationships>");
        xml
    }

    #[test]
    fn read_merged_cells() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_record(records::BRT_BEGIN_MERGE_CELLS, &[]));
        post_data.extend_from_slice(&build_merge_cell(0, 2, 0, 3));
        post_data.extend_from_slice(&build_merge_cell(5, 5, 1, 4));
        post_data.extend_from_slice(&build_record(records::BRT_END_MERGE_CELLS, &[]));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let merged = ws.merged_regions();
        assert_eq!(merged.len(), 2);
        assert_eq!(merged[0].start.row, 0);
        assert_eq!(merged[0].start.col, 0);
        assert_eq!(merged[0].end.row, 2);
        assert_eq!(merged[0].end.col, 3);
        assert_eq!(merged[1].start.row, 5);
        assert_eq!(merged[1].start.col, 1);
        assert_eq!(merged[1].end.row, 5);
        assert_eq!(merged[1].end.col, 4);
    }

    #[test]
    fn read_custom_row_heights() {
        let ws_data =
            build_worksheet_with_row_hdrs(&[(0, 300, false), (1, 480, false), (2, 600, false)]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.row_height(0), 15.0);
        assert_eq!(ws.row_height(1), 24.0);
        assert_eq!(ws.row_height(2), 30.0);
    }

    #[test]
    fn read_hidden_rows() {
        let ws_data = build_worksheet_with_row_hdrs(&[(0, 0, false), (1, 0, true), (2, 0, false)]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert!(!ws.is_row_hidden(0));
        assert!(ws.is_row_hidden(1));
        assert!(!ws.is_row_hidden(2));
    }

    #[test]
    fn read_column_widths() {
        let mut pre_data = Vec::new();
        pre_data.extend_from_slice(&build_col_info(0, 0, 2560, false, true));
        pre_data.extend_from_slice(&build_col_info(1, 3, 5120, false, true));
        pre_data.extend_from_slice(&build_col_info(4, 4, 2560, true, true));

        let ws_data = build_worksheet_bin_full(&pre_data, &[], &[]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.column_width(0), 10.0);
        assert_eq!(ws.column_width(1), 20.0);
        assert_eq!(ws.column_width(2), 20.0);
        assert_eq!(ws.column_width(3), 20.0);
        assert!(ws.is_column_hidden(4));
    }

    #[test]
    fn read_freeze_panes() {
        let pre_data = build_pane(2.0, 3.0, 1);

        let ws_data = build_worksheet_bin_full(&pre_data, &[], &[]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let freeze = ws.freeze_panes().unwrap();
        assert_eq!(freeze.row, 3);
        assert_eq!(freeze.col, 2);
    }

    #[test]
    fn read_auto_filter() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_auto_filter(0, 10, 0, 3));
        post_data.extend_from_slice(&build_record(records::BRT_END_A_FILTER, &[]));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let af = ws.auto_filter().unwrap();
        assert_eq!(af.range.start.row, 0);
        assert_eq!(af.range.start.col, 0);
        assert_eq!(af.range.end.row, 10);
        assert_eq!(af.range.end.col, 3);
    }

    #[test]
    fn read_page_margins() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_margins(1.0, 1.0, 0.5, 0.5, 0.25, 0.25));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let ps = ws.page_setup();
        assert_eq!(ps.left_margin, 1.0);
        assert_eq!(ps.right_margin, 1.0);
        assert_eq!(ps.top_margin, 0.5);
        assert_eq!(ps.bottom_margin, 0.5);
        assert_eq!(ps.header_margin, 0.25);
        assert_eq!(ps.footer_margin, 0.25);
    }

    #[test]
    fn read_page_setup_landscape() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_page_setup(9, 75, true));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let ps = ws.page_setup();
        assert_eq!(ps.paper_size, 9);
        assert_eq!(ps.scale, 75);
        assert_eq!(
            ps.orientation,
            duke_sheets_core::worksheet::PageOrientation::Landscape
        );
    }

    #[test]
    fn read_header_footer() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_header_footer("&CPage &P", "&CFooter"));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let ps = ws.page_setup();
        assert_eq!(ps.odd_header.as_deref(), Some("&CPage &P"));
        assert_eq!(ps.odd_footer.as_deref(), Some("&CFooter"));
        assert!(ps.scale_with_doc);
        assert!(ps.align_with_margins);
    }

    #[test]
    fn read_hyperlink_external() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_hlink(0, 0, "rId1", "", "Click me", "Example"));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let wb_rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);

        let sheet_rels_content = sheet_rels_xml(&[(
            "rId1",
            "https://example.com",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        )]);

        let zip = build_xlsb_zip_with_sheet_rels(
            &wb_data,
            &wb_rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
            &[("xl/worksheets/_rels/sheet1.bin.rels", &sheet_rels_content)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let hl = ws.hyperlink_at(0, 0).unwrap();
        assert_eq!(hl.target, "https://example.com");
        assert_eq!(hl.display.as_deref(), Some("Example"));
        assert_eq!(hl.tooltip.as_deref(), Some("Click me"));
    }

    #[test]
    fn read_hyperlink_internal() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_hlink(0, 0, "", "Sheet2!A1", "", ""));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let wb_rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &wb_rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let hl = ws.hyperlink_at(0, 0).unwrap();
        assert_eq!(hl.target, "#Sheet2!A1");
        assert_eq!(hl.location.as_deref(), Some("Sheet2!A1"));
    }

    fn build_dval(
        val_type: u32,
        operator: u32,
        error_style: u32,
        allow_blank: bool,
        show_input: bool,
        show_error: bool,
        formula1: &str,
        formula2: &str,
        error_title: &str,
        error_msg: &str,
        input_title: &str,
        input_msg: &str,
        ranges: &[(u32, u32, u32, u32)],
    ) -> Vec<u8> {
        let mut payload = Vec::new();

        // Bit-packed header per [MS-XLSB] §2.4.356:
        // bits 0-3 valType, 4-6 errStyle, 7 unused, 8 fAllowBlank,
        // 9 fSuppressCombo, 10-17 mdImeMode, 18 fShowInputMsg,
        // 19 fShowErrorMsg, 20-23 typOperator, 24-31 reserved.
        let mut header: u32 = 0;
        header |= val_type & 0xF;
        header |= (error_style & 0x7) << 4;
        if allow_blank {
            header |= 1 << 8;
        }
        if show_input {
            header |= 1 << 18;
        }
        if show_error {
            header |= 1 << 19;
        }
        header |= (operator & 0xF) << 20;
        payload.extend_from_slice(&header.to_le_bytes());

        // sqrfx (UncheckedSqRfX): cFx + cFx × UncheckedRfX
        payload.extend_from_slice(&(ranges.len() as u32).to_le_bytes());
        for &(r1, r2, c1, c2) in ranges {
            payload.extend_from_slice(&r1.to_le_bytes());
            payload.extend_from_slice(&r2.to_le_bytes());
            payload.extend_from_slice(&c1.to_le_bytes());
            payload.extend_from_slice(&c2.to_le_bytes());
        }

        // DValStrings: 4 XLNullableWideStrings
        payload.extend_from_slice(&xlnullwide(error_title));
        payload.extend_from_slice(&xlnullwide(error_msg));
        payload.extend_from_slice(&xlnullwide(input_title));
        payload.extend_from_slice(&xlnullwide(input_msg));

        // formula1, formula2 (DVParsedFormula): cce + rgce + cb + rgcb
        for f in [formula1, formula2] {
            if f.is_empty() {
                payload.extend_from_slice(&0u32.to_le_bytes()); // cce
                payload.extend_from_slice(&0u32.to_le_bytes()); // cb
            } else {
                let tokens = biff12_tstr_token(f);
                payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes()); // cce
                payload.extend_from_slice(&tokens);
                payload.extend_from_slice(&0u32.to_le_bytes()); // cb
            }
        }

        build_record(records::BRT_DVAL, &payload)
    }

    fn xlnullwide(s: &str) -> Vec<u8> {
        if s.is_empty() {
            return 0xFFFFFFFFu32.to_le_bytes().to_vec();
        }
        xlwide(s)
    }

    fn build_cond_fmt_payload(ccf: u32, ranges: &[(u32, u32, u32, u32)]) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&ccf.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes()); // fPivot
        payload.extend_from_slice(&(ranges.len() as u32).to_le_bytes());
        for &(r1, r2, c1, c2) in ranges {
            payload.extend_from_slice(&r1.to_le_bytes());
            payload.extend_from_slice(&r2.to_le_bytes());
            payload.extend_from_slice(&c1.to_le_bytes());
            payload.extend_from_slice(&c2.to_le_bytes());
        }
        payload
    }

    fn build_cf_rule_cell_is(
        operator: u32,
        priority: i32,
        dxf_id: u32,
        _ranges: &[(u32, u32, u32, u32)],
        formula1: &str,
    ) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u32.to_le_bytes()); // iType = cellIs
        payload.extend_from_slice(&0u32.to_le_bytes()); // iTemplate = CF_TEMPLATE_EXPR
        payload.extend_from_slice(&dxf_id.to_le_bytes()); // dxfId
        payload.extend_from_slice(&(priority as u32).to_le_bytes()); // iPri
        payload.extend_from_slice(&operator.to_le_bytes()); // iParam
        payload.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        payload.extend_from_slice(&0u32.to_le_bytes()); // reserved2
        payload.extend_from_slice(&0u16.to_le_bytes()); // flags

        let has_fmla1 = !formula1.is_empty();
        payload.extend_from_slice(&(if has_fmla1 { 1u32 } else { 0u32 }).to_le_bytes()); // cbFmla1
        payload.extend_from_slice(&0u32.to_le_bytes()); // cbFmla2
        payload.extend_from_slice(&0u32.to_le_bytes()); // cbFmla3

        payload.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes()); // strParam null

        if has_fmla1 {
            let tokens = biff12_tstr_token(formula1);
            payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes()); // cce
            payload.extend_from_slice(&tokens); // rgce
            payload.extend_from_slice(&0u32.to_le_bytes()); // cb (no rgcb)
        }

        build_record(records::BRT_BEGIN_CF_RULE, &payload)
    }

    /// Comments part with the MS-XLSB 2.4.33 record sequence
    /// (0x0274-based ids, 36-byte BrtBeginComment).
    fn build_comments_bin(comments: &[(&str, u32, u16, &str)]) -> Vec<u8> {
        let mut data = Vec::new();

        // Collect unique authors
        let mut authors: Vec<String> = Vec::new();
        for &(author, _, _, _) in comments {
            if !authors.contains(&author.to_string()) {
                authors.push(author.to_string());
            }
        }

        data.extend_from_slice(&build_record(records::BRT_BEGIN_COMMENTS, &[]));
        data.extend_from_slice(&build_record(records::BRT_BEGIN_COMMENT_AUTHORS, &[]));
        for author in &authors {
            data.extend_from_slice(&build_record(records::BRT_COMMENT_AUTHOR, &xlwide(author)));
        }
        data.extend_from_slice(&build_record(records::BRT_END_COMMENT_AUTHORS, &[]));

        data.extend_from_slice(&build_record(records::BRT_BEGIN_COMMENT_LIST, &[]));
        for &(author, row, col, text) in comments {
            let author_id = authors.iter().position(|a| a == author).unwrap_or(0) as u32;
            // iauthor + UncheckedRfX + GUID
            let mut comment_payload = Vec::new();
            comment_payload.extend_from_slice(&author_id.to_le_bytes());
            comment_payload.extend_from_slice(&row.to_le_bytes());
            comment_payload.extend_from_slice(&row.to_le_bytes());
            comment_payload.extend_from_slice(&(col as u32).to_le_bytes());
            comment_payload.extend_from_slice(&(col as u32).to_le_bytes());
            comment_payload.extend_from_slice(&[0u8; 16]);
            data.extend_from_slice(&build_record(records::BRT_BEGIN_COMMENT, &comment_payload));

            // BrtCommentText: RichStr = flags(1) + XLWideString
            let mut text_payload = vec![0u8]; // flags
            text_payload.extend_from_slice(&xlwide(text));
            data.extend_from_slice(&build_record(records::BRT_COMMENT_TEXT, &text_payload));

            data.extend_from_slice(&build_record(records::BRT_END_COMMENT, &[]));
        }
        data.extend_from_slice(&build_record(records::BRT_END_COMMENT_LIST, &[]));
        data.extend_from_slice(&build_record(records::BRT_END_COMMENTS, &[]));

        data
    }

    /// Comments part with the off-spec 0x0278-based ids and 12-byte
    /// BrtBeginComment body our old writer emitted.
    fn build_legacy_comments_bin(comments: &[(&str, u32, u16, &str)]) -> Vec<u8> {
        let mut data = Vec::new();

        let mut authors: Vec<String> = Vec::new();
        for &(author, _, _, _) in comments {
            if !authors.contains(&author.to_string()) {
                authors.push(author.to_string());
            }
        }

        data.extend_from_slice(&build_record(
            records::BRT_LEGACY_BEGIN_COMMENT_AUTHORS,
            &[],
        ));
        for author in &authors {
            data.extend_from_slice(&build_record(
                records::BRT_LEGACY_COMMENT_AUTHOR,
                &xlwide(author),
            ));
        }
        // Legacy end-authors (0x0279)
        data.extend_from_slice(&build_record(0x0279, &[]));

        // Legacy begin-list (0x027B)
        data.extend_from_slice(&build_record(0x027B, &[]));
        for &(author, row, col, text) in comments {
            let author_id = authors.iter().position(|a| a == author).unwrap_or(0) as u32;
            let mut comment_payload = Vec::new();
            comment_payload.extend_from_slice(&author_id.to_le_bytes());
            comment_payload.extend_from_slice(&row.to_le_bytes());
            comment_payload.extend_from_slice(&(col as u32).to_le_bytes());
            data.extend_from_slice(&build_record(
                records::BRT_LEGACY_BEGIN_COMMENT,
                &comment_payload,
            ));

            let mut text_payload = vec![0u8];
            text_payload.extend_from_slice(&xlwide(text));
            data.extend_from_slice(&build_record(
                records::BRT_LEGACY_COMMENT_TEXT,
                &text_payload,
            ));

            data.extend_from_slice(&build_record(records::BRT_LEGACY_END_COMMENT, &[]));
        }
        data.extend_from_slice(&build_record(records::BRT_LEGACY_END_COMMENT_LIST, &[]));

        data
    }

    fn build_table_bin(
        id: u32,
        name: &str,
        display_name: &str,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
        header_row_count: u32,
        columns: &[&str],
    ) -> Vec<u8> {
        let mut data = Vec::new();

        let mut payload = Vec::new();
        payload.extend_from_slice(&first_row.to_le_bytes());
        payload.extend_from_slice(&last_row.to_le_bytes());
        payload.extend_from_slice(&(first_col as u32).to_le_bytes());
        payload.extend_from_slice(&(last_col as u32).to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes()); // lt
        payload.extend_from_slice(&id.to_le_bytes());
        payload.extend_from_slice(&header_row_count.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes()); // crwTotals
        payload.extend_from_slice(&1u32.to_le_bytes()); // flags (fShownTotalRow=1)
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfHeader
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfData
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfAgg
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfBorder
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfHeaderBorder
        payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfAggBorder
        payload.extend_from_slice(&0u32.to_le_bytes()); // dwConnID
        payload.extend_from_slice(&xlwide(name));
        payload.extend_from_slice(&xlwide(display_name));
        payload.extend_from_slice(&xlwide("")); // stComment
        payload.extend_from_slice(&xlwide("")); // stStyleHeader
        payload.extend_from_slice(&xlwide("")); // stStyleData
        payload.extend_from_slice(&xlwide("")); // stStyleAgg
        data.extend_from_slice(&build_record(records::BRT_BEGIN_LIST, &payload));

        data.extend_from_slice(&build_record(records::BRT_BEGIN_LIST_COLS, &[]));
        for (i, col_name) in columns.iter().enumerate() {
            let mut col_payload = Vec::new();
            col_payload.extend_from_slice(&((i + 1) as u32).to_le_bytes());
            col_payload.extend_from_slice(&0u32.to_le_bytes()); // ilta
            col_payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfHdr
            col_payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfInsertRow
            col_payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfAgg
            col_payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // idqsif
            col_payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfData
            col_payload.extend_from_slice(&xlwide(col_name)); // stName
            col_payload.extend_from_slice(&xlwide(col_name)); // stCaption
            col_payload.extend_from_slice(&xlwide("")); // stTotal
            col_payload.extend_from_slice(&xlwide("")); // stStyleHeader
            col_payload.extend_from_slice(&xlwide("")); // stStyleInsertRow
            col_payload.extend_from_slice(&xlwide("")); // stStyleAgg
            data.extend_from_slice(&build_record(records::BRT_BEGIN_LIST_COL, &col_payload));
            data.extend_from_slice(&build_record(records::BRT_END_LIST_COL, &[]));
        }
        data.extend_from_slice(&build_record(records::BRT_END_LIST_COLS, &[]));

        data.extend_from_slice(&build_record(records::BRT_END_LIST, &[]));
        data
    }

    fn build_xlsb_zip_with_extras(
        workbook_bin: &[u8],
        rels_xml: &str,
        sst: Option<&[u8]>,
        worksheets: &[(&str, &[u8])],
        sheet_rels: &[(&str, &str)],
        extra_bins: &[(&str, &[u8])],
    ) -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = SimpleFileOptions::default();
            zip.start_file("xl/workbook.bin", opts).unwrap();
            zip.write_all(workbook_bin).unwrap();
            zip.start_file("xl/_rels/workbook.bin.rels", opts).unwrap();
            zip.write_all(rels_xml.as_bytes()).unwrap();
            if let Some(sst_data) = sst {
                zip.start_file("xl/sharedStrings.bin", opts).unwrap();
                zip.write_all(sst_data).unwrap();
            }
            for (path, data) in worksheets {
                zip.start_file(*path, opts).unwrap();
                zip.write_all(data).unwrap();
            }
            for (path, data) in sheet_rels {
                zip.start_file(*path, opts).unwrap();
                zip.write_all(data.as_bytes()).unwrap();
            }
            for (path, data) in extra_bins {
                zip.start_file(*path, opts).unwrap();
                zip.write_all(data).unwrap();
            }
            zip.finish().unwrap();
        }
        buf
    }

    fn sheet_rels_xml_full(entries: &[(&str, &str, &str)]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for (id, target, rel_type) in entries {
            xml.push_str(&format!(
                r#"<Relationship Id="{id}" Target="{target}" Type="{rel_type}"/>"#,
            ));
        }
        xml.push_str("</Relationships>");
        xml
    }

    #[test]
    fn read_comments_bin() {
        let comments_data = build_comments_bin(&[
            ("Alice", 0, 0, "First comment"),
            ("Bob", 1, 2, "Second comment"),
        ]);

        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(1.0))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let wb_rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);

        let sr = sheet_rels_xml_full(&[(
            "rId10",
            "../comments1.bin",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        )]);

        let zip = build_xlsb_zip_with_extras(
            &wb_data,
            &wb_rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
            &[("xl/worksheets/_rels/sheet1.bin.rels", &sr)],
            &[("xl/comments1.bin", &comments_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();

        let c1 = ws.comment_at(0, 0).expect("comment at A1");
        assert_eq!(c1.author, "Alice");
        assert_eq!(c1.plain_text(), "First comment");

        let c2 = ws.comment_at(1, 2).expect("comment at C2");
        assert_eq!(c2.author, "Bob");
        assert_eq!(c2.plain_text(), "Second comment");

        assert!(ws.comment_at(0, 1).is_none());
    }

    /// Files our old writer produced used off-spec 0x0278-based
    /// comment record ids with a 12-byte BrtBeginComment; they must
    /// keep reading.
    #[test]
    fn read_comments_bin_legacy_ids() {
        let comments_data = build_legacy_comments_bin(&[
            ("Alice", 0, 0, "First comment"),
            ("Bob", 1, 2, "Second comment"),
        ]);

        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(1.0))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let wb_rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);

        let sr = sheet_rels_xml_full(&[(
            "rId10",
            "../comments1.bin",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        )]);

        let zip = build_xlsb_zip_with_extras(
            &wb_data,
            &wb_rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
            &[("xl/worksheets/_rels/sheet1.bin.rels", &sr)],
            &[("xl/comments1.bin", &comments_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();

        let c1 = ws.comment_at(0, 0).expect("comment at A1");
        assert_eq!(c1.author, "Alice");
        assert_eq!(c1.plain_text(), "First comment");

        let c2 = ws.comment_at(1, 2).expect("comment at C2");
        assert_eq!(c2.author, "Bob");
        assert_eq!(c2.plain_text(), "Second comment");
    }

    #[test]
    fn read_comments_xml_fallback() {
        let comments_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>TestUser</author></authors>
  <commentList>
    <comment ref="B3" authorId="0">
      <text><r><t>XML comment</t></r></text>
    </comment>
  </commentList>
</comments>"#;

        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(1.0))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let wb_rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);

        let sr = sheet_rels_xml_full(&[(
            "rId10",
            "../comments1.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        )]);

        let zip = build_xlsb_zip_with_extras(
            &wb_data,
            &wb_rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
            &[
                ("xl/worksheets/_rels/sheet1.bin.rels", &sr),
                ("xl/comments1.xml", comments_xml),
            ],
            &[],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();

        let c = ws.comment_at(2, 1).expect("comment at B3");
        assert_eq!(c.author, "TestUser");
        assert_eq!(c.plain_text(), "XML comment");
    }

    #[test]
    fn read_data_validation_list() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_record(records::BRT_BEGIN_DVAL, &[]));
        post_data.extend_from_slice(&build_dval(
            3,    // list
            0,    // operator (unused for list)
            0,    // error_style
            true, // allow_blank
            true, // show_input
            true, // show_error
            "Yes,No,Maybe",
            "",
            "Error",
            "Pick from list",
            "Choose",
            "Select a value",
            &[(0, 9, 0, 0)], // A1:A10
        ));
        post_data.extend_from_slice(&build_record(records::BRT_END_DVAL, &[]));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let validations = ws.data_validations();
        assert_eq!(validations.len(), 1);

        let dv = &validations[0];
        assert!(dv.allow_blank);
        assert!(dv.show_input_message);
        assert!(dv.show_error_alert);
        assert_eq!(dv.error_title.as_deref(), Some("Error"));
        assert_eq!(dv.error_message.as_deref(), Some("Pick from list"));
        assert_eq!(dv.input_title.as_deref(), Some("Choose"));
        assert_eq!(dv.input_message.as_deref(), Some("Select a value"));
        assert_eq!(dv.ranges.len(), 1);
        assert_eq!(dv.ranges[0].start.row, 0);
        assert_eq!(dv.ranges[0].end.row, 9);

        if let duke_sheets_core::validation::ValidationType::List { source } = &dv.validation_type {
            assert_eq!(source, "Yes,No,Maybe");
        } else {
            panic!("Expected List validation type");
        }
    }

    #[test]
    fn read_conditional_format_cell_is() {
        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_record(
            records::BRT_BEGIN_COND_FMT,
            &build_cond_fmt_payload(1, &[(0, 9, 0, 0)]),
        ));
        post_data.extend_from_slice(&build_cf_rule_cell_is(
            5,               // greaterThan
            1,               // priority
            0,               // dxf_id
            &[(0, 9, 0, 0)], // A1:A10
            "100",
        ));
        post_data.extend_from_slice(&build_record(records::BRT_END_CF_RULE, &[]));
        post_data.extend_from_slice(&build_record(records::BRT_END_COND_FMT, &[]));

        let ws_data = build_worksheet_bin_full(&[], &[], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let cf_rules = ws.conditional_formats();
        assert_eq!(cf_rules.len(), 1);

        let rule = &cf_rules[0];
        assert_eq!(rule.priority, 1);
        assert_eq!(rule.dxf_id, Some(0));
        assert_eq!(rule.ranges.len(), 1);
        assert_eq!(rule.ranges[0].start.row, 0);
        assert_eq!(rule.ranges[0].end.row, 9);

        if let duke_sheets_core::conditional_format::CfRuleType::CellIs {
            operator, formula1, ..
        } = &rule.rule_type
        {
            assert_eq!(
                *operator,
                duke_sheets_core::conditional_format::CfOperator::GreaterThan
            );
            assert_eq!(formula1, "\"100\"");
        } else {
            panic!("Expected CellIs rule type, got {:?}", rule.rule_type);
        }
    }

    #[test]
    fn read_table_from_bin() {
        let table_data = build_table_bin(
            1,
            "Sales",
            "Sales",
            0,
            5,
            0,
            2,
            1,
            &["Product", "Region", "Revenue"],
        );

        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(1.0))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let wb_rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);

        let sr = sheet_rels_xml_full(&[(
            "rId20",
            "../tables/table1.bin",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        )]);

        let zip = build_xlsb_zip_with_extras(
            &wb_data,
            &wb_rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
            &[("xl/worksheets/_rels/sheet1.bin.rels", &sr)],
            &[("xl/tables/table1.bin", &table_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        let tables = ws.tables();
        assert_eq!(tables.len(), 1);

        let t = &tables[0];
        assert_eq!(t.id, 1);
        assert_eq!(t.name, "Sales");
        assert_eq!(t.display_name, "Sales");
        assert_eq!(t.reference.start.row, 0);
        assert_eq!(t.reference.end.row, 5);
        assert_eq!(t.reference.start.col, 0);
        assert_eq!(t.reference.end.col, 2);
        assert_eq!(t.header_row_count, 1);
        assert!(t.has_header_row());
        assert!(!t.has_totals_row());

        assert_eq!(t.columns.len(), 3);
        assert_eq!(t.columns[0].name, "Product");
        assert_eq!(t.columns[1].name, "Region");
        assert_eq!(t.columns[2].name, "Revenue");
    }

    #[test]
    fn read_no_comments_graceful() {
        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(1.0))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.comment_count(), 0);
    }

    #[test]
    fn read_all_worksheet_features() {
        let mut pre_data = Vec::new();
        pre_data.extend_from_slice(&build_pane(1.0, 2.0, 1));
        pre_data.extend_from_slice(&build_col_info(0, 0, 3840, false, true));

        let mut post_data = Vec::new();
        post_data.extend_from_slice(&build_record(records::BRT_BEGIN_MERGE_CELLS, &[]));
        post_data.extend_from_slice(&build_merge_cell(0, 1, 0, 2));
        post_data.extend_from_slice(&build_record(records::BRT_END_MERGE_CELLS, &[]));
        post_data.extend_from_slice(&build_auto_filter(0, 5, 0, 2));
        post_data.extend_from_slice(&build_record(records::BRT_END_A_FILTER, &[]));
        post_data.extend_from_slice(&build_margins(0.8, 0.8, 0.6, 0.6, 0.2, 0.2));
        post_data.extend_from_slice(&build_page_setup(1, 100, false));

        let ws_data =
            build_worksheet_bin_full(&pre_data, &[(0, 0, WorksheetCell::Real(42.0))], &post_data);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let zip = build_xlsb_zip(
            &wb_data,
            &rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();

        assert_eq!(ws.get_value_at(0, 0), CellValue::Number(42.0));
        assert_eq!(ws.merged_regions().len(), 1);
        assert_eq!(ws.column_width(0), 15.0);
        let freeze = ws.freeze_panes().unwrap();
        assert_eq!(freeze.row, 2);
        assert_eq!(freeze.col, 1);
        let af = ws.auto_filter().unwrap();
        assert_eq!(af.range.start.row, 0);
        assert_eq!(af.range.end.row, 5);
        assert_eq!(ws.page_setup().left_margin, 0.8);
        assert_eq!(
            ws.page_setup().orientation,
            duke_sheets_core::worksheet::PageOrientation::Portrait
        );
    }

    #[test]
    fn brt_filter_does_not_decode_stale_buffer_bytes() {
        use duke_sheets_core::auto_filter::ColumnFilter;
        use duke_sheets_core::Worksheet;
        use duke_sheets_formula::decompile::FormulaContext;
        use std::collections::HashMap;

        // The record iterator's reuse buffer only grows, so a BrtFilter
        // whose cch exceeds its own record length must not decode bytes
        // left over from a previous, larger record.
        let mut stream = Vec::new();
        stream.extend_from_slice(&build_record(records::BRT_BEGIN_SHEET_DATA, &[]));
        stream.extend_from_slice(&build_record(records::BRT_END_SHEET_DATA, &[]));
        // Big ignored record fills the reuse buffer with 0x41 bytes.
        stream.extend_from_slice(&build_record(0x0FFF, &[0x41u8; 120]));
        let mut afr = Vec::new();
        afr.extend_from_slice(&0u32.to_le_bytes());
        afr.extend_from_slice(&4u32.to_le_bytes());
        afr.extend_from_slice(&0u32.to_le_bytes());
        afr.extend_from_slice(&0u32.to_le_bytes());
        stream.extend_from_slice(&build_record(records::BRT_BEGIN_A_FILTER, &afr));
        let mut fc = Vec::new();
        fc.extend_from_slice(&0u32.to_le_bytes());
        fc.extend_from_slice(&0u16.to_le_bytes());
        stream.extend_from_slice(&build_record(records::BRT_BEGIN_FILTER_COLUMN, &fc));
        stream.extend_from_slice(&build_record(
            records::BRT_BEGIN_FILTERS,
            &0u32.to_le_bytes(),
        ));
        // Malicious BrtFilter: cch claims 20 chars, carries one.
        let mut bad = Vec::new();
        bad.extend_from_slice(&20u32.to_le_bytes());
        bad.extend_from_slice(&[0x58, 0x00]); // "X"
        stream.extend_from_slice(&build_record(records::BRT_FILTER, &bad));
        stream.extend_from_slice(&build_record(records::BRT_END_FILTERS, &[]));
        stream.extend_from_slice(&build_record(records::BRT_END_FILTER_COLUMN, &[]));
        stream.extend_from_slice(&build_record(records::BRT_END_A_FILTER, &[]));

        let mut ws = Worksheet::new("S1".to_string());
        crate::reader::worksheet::read_worksheet(
            std::io::Cursor::new(stream),
            &mut ws,
            &[],
            &[],
            &FormulaContext::new(Vec::new()),
            &HashMap::new(),
        )
        .expect("worksheet parse");

        let af = ws.auto_filter().expect("autofilter parsed");
        match &af.filter_columns[0].filter {
            ColumnFilter::Values(vf) => {
                assert!(
                    vf.values.is_empty(),
                    "truncated BrtFilter must be skipped, not decoded from stale bytes: {:?}",
                    vf.values
                );
            }
            other => panic!("expected Values filter, got {other:?}"),
        }
    }

    /// A drawing-part control twin (either flavor) is matched to its
    /// VML control by spid and never duplicated as a raw entry; the
    /// twin's cNvPr name rides onto the control.
    #[test]
    fn drawing_part_control_twin_dedupes_against_vml_control() {
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1025" name="Check Box 1"/><xdr:cNvGraphicFramePr><a:graphicFrameLocks/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/compatibility"><com14:compatSp xmlns:com14="http://schemas.microsoft.com/office/drawing/2010/compatibility" spid="_x0000_s1025"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent></xdr:wsDr>"#;
        let vml = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <v:textbox><div style='text-align:left'>tick</div></v:textbox>
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor>
   <x:Checked>1</x:Checked>
  </x:ClientData>
 </v:shape>
</xml>"##;

        let ws_data = build_worksheet_bin(&[(0, 0, WorksheetCell::Real(1.0))]);
        let wb_data = build_workbook_bin(&[("Sheet1", "rId1")], false);
        let wb_rels = rels_xml(&[("rId1", "worksheets/sheet1.bin")]);
        let sr = sheet_rels_xml_full(&[
            (
                "rId1",
                "../drawings/drawing1.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            ),
            (
                "rId2",
                "../drawings/vmlDrawing1.vml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            ),
        ]);

        let zip = build_xlsb_zip_with_extras(
            &wb_data,
            &wb_rels,
            None,
            &[("xl/worksheets/sheet1.bin", &ws_data)],
            &[
                ("xl/worksheets/_rels/sheet1.bin.rels", &sr),
                ("xl/drawings/drawing1.xml", drawing_xml),
                ("xl/drawings/vmlDrawing1.vml", vml),
            ],
            &[],
        );

        let wb = XlsbReader::read(Cursor::new(zip)).unwrap();
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(
            ws.drawings().len(),
            1,
            "twin must not duplicate: {:?}",
            ws.drawings()
        );
        assert_eq!(ws.form_control_count(), 1);
        let control = ws.form_controls().next().unwrap();
        assert_eq!(control.payload.caption_text().as_deref(), Some("tick"));
        assert_eq!(control.object.unwrap().meta.name.as_deref(), Some("Check Box 1"));
    }
}
